"""Fontes de referência por origem (coir.fontes): provedores, precedência e divergência.

A composição sobre o data/ atual devolve o que a tabela única devolvia; o mesmo evento em duas
fontes entra uma vez, com as origens somadas; conteúdo diferente na mesma chave vale a fonte de
maior precedência e fica registrado.
"""

from dataclasses import replace
from pathlib import Path

from coir import fontes
from coir.fontes.composicao import juntar_eventos
from coir.fontes.modelo import Evento, evento_valido
from coir.fontes.provedores import b3, cvm
from coir.recursos import pasta_data
from tools.referencias.gerar import ler_eventos_pacote

BASE = Evento(emissor="FICT", isin="BRFICTACNOR0", tipo="DESDOBRAMENTO", fator="100",
              data_com="2024-05-02", aprovado_em="2024-04-10", isin_emitido="", preco="",
              periodo="", data_subscricao="", origem="cvm_ipe", referencia="IPE 1",
              custo_atribuido="")  # fmt: skip


def _linhas_validas(pasta: Path) -> list[dict[str, str]]:
    """A tabela de eventos numa linha por evento (os arquivos da CVM e da bolsa juntos)."""
    return [r for r in ler_eventos_pacote(pasta) if evento_valido(r)]


def test_provedores_cobrem_a_tabela_sem_perder_linha() -> None:
    """Cada linha válida da tabela sai de pelo menos um provedor, com a origem que tem."""
    pasta = pasta_data()
    chaves = {e.chave for e in cvm.eventos(pasta)} | {e.chave for e in b3.eventos(pasta)}
    linhas = _linhas_validas(pasta)
    assert chaves == {Evento.de_linha(r).chave for r in linhas}
    assert all("b3" in e.origem.split("+") for e in b3.eventos(pasta))
    assert all(e.origem.split("+") != ["b3"] for e in cvm.eventos(pasta))


def test_composicao_igual_a_tabela_na_mesma_ordem() -> None:
    """A junção dos provedores devolve as linhas da tabela, na ordem dela (o pipeline grava pela
    chave), menos os pares que se anulam, e sem divergência (a tabela já vem conciliada)."""
    pasta = pasta_data()
    lista, divergencias = fontes.eventos_e_divergencias(pasta)
    originais = [Evento.de_linha(r) for r in _linhas_validas(pasta)]
    assert divergencias == []
    assert lista == [e for e in originais if e in set(lista)]
    assert len(originais) - len(lista) == 24  # 12 pares grupamento + desdobramento


def test_mesmo_evento_em_duas_fontes_soma_origem() -> None:
    da_bolsa = replace(BASE, origem="b3", referencia="", custo_atribuido="")
    com_custo = replace(BASE, custo_atribuido="1.5")
    [junto], divergencias = juntar_eventos([com_custo], [da_bolsa])
    assert divergencias == []
    assert (junto.origem, junto.referencia, junto.custo_atribuido) == ("b3+cvm_ipe", "IPE 1", "1.5")


def test_conteudo_diferente_vale_a_precedencia_e_registra() -> None:
    da_bolsa = replace(BASE, origem="b3", aprovado_em="2024-04-12")
    [mantido], [divergencia] = juntar_eventos([BASE], [da_bolsa])
    assert mantido == BASE
    assert (divergencia.mantida, divergencia.descartada) == ("cvm_ipe", "b3")
    assert divergencia.chave == BASE.chave
    # custo atribuído diferente nas duas fontes também é divergência, não soma
    _, [outra] = juntar_eventos([replace(BASE, custo_atribuido="1")],
                                [replace(BASE, origem="b3", custo_atribuido="2")])  # fmt: skip
    assert outra.mantida == "cvm_ipe"


def test_ordem_pela_chave_qualquer_que_seja_a_entrada() -> None:
    """Pacote antigo ou editado à mão: a saída sai na ordem da chave, sempre a mesma."""
    eventos = [replace(BASE, data_com=f"2024-05-0{d}", origem=o)
               for d, o in ((3, "b3"), (1, "cvm_fre"), (2, "b3"))]  # fmt: skip
    lista, _ = juntar_eventos(eventos[1:2], [eventos[2], eventos[0]])
    assert [e.data_com for e in lista] == ["2024-05-01", "2024-05-02", "2024-05-03"]
    assert juntar_eventos([], list(reversed(eventos)))[0] == juntar_eventos(eventos, [])[0]


def test_origem_vazia_vale_bolsa_ao_somar() -> None:
    """Pacote antigo, sem a coluna origem: a linha é da bolsa, e a soma não perde o "b3"."""
    [junto], _ = juntar_eventos([replace(BASE, origem="")], [replace(BASE, origem="cvm_fre")])
    assert junto.origem == "b3+cvm_fre"


def test_pares_que_se_anulam_saem_em_qualquer_fonte() -> None:
    grupamento = replace(BASE, tipo="GRUPAMENTO", fator="0.01")
    desdobramento = replace(BASE, fator="9900", origem="b3")
    lista, _ = juntar_eventos([grupamento], [desdobramento])
    assert lista == []


def test_fonte_ausente_devolve_vazio(tmp_path: Path) -> None:
    """Sem os arquivos de uma fonte (pacote antigo ou fonte desligada), o conceito fica vazio."""
    assert fontes.eventos(tmp_path) == []
    assert fontes.proventos(tmp_path) == fontes.proventos_historicos(tmp_path) == []
    assert fontes.instrumentos(tmp_path) == fontes.emissores(tmp_path) == {}
    assert fontes.tickers_por_isin(tmp_path) == {}
    fontes.validar_pacote(tmp_path)


def test_servicos_nao_leem_arquivo_da_fonte() -> None:
    """Os serviços pedem o conceito a coir.fontes: nenhum abre tabela da bolsa pelo nome."""
    servicos = Path(__file__).resolve().parents[2] / "coir" / "servicos"
    proibidos = ("eventos_b3(", "proventos_b3(", "proventos_historico_b3(", "saidas_b3(",
                 "validar_tabelas_b3(")  # fmt: skip
    achados = [f"{arq.name}: {p}" for arq in servicos.glob("*.py")
               for p in proibidos if p in arq.read_text(encoding="utf-8")]  # fmt: skip
    assert achados == []
