"""F6: serviços das telas de importação, notas e operações (banco temporário, notas anonimizadas
das fixtures no lugar dos PDFs)."""

from datetime import datetime
from decimal import Decimal
from pathlib import Path

import pytest
from reportlab.pdfgen.canvas import Canvas

from coir import referencias
from coir.config import VAR_CONFIG, VAR_DADOS, arquivo_config, dados_dir, salvar_dados_dir
from coir.erros import ConfiguracaoError, ExtracaoError, ImportacaoError
from coir.servicos import cadastros, notas, operacoes
from coir.servicos.estado import Estado
from coir.servicos.importacao import copiar_para_pasta, importar_lista, por_arquivo
from tests.apoio import cpf_de_teste
from tools.banco_demo import criar, extrator

AGORA = datetime(2026, 10, 16, 12, 0, 0)
FIXTURES = Path(__file__).parents[1] / "fixtures" / "notas"


@pytest.fixture(scope="module")
def demo(tmp_path_factory: pytest.TempPathFactory) -> Path:
    return criar(tmp_path_factory.mktemp("demo"))


# ------------------------------------------------------------------------------ configuração
def test_pasta_escolhida_na_interface_vale_sem_a_variavel(tmp_path: Path) -> None:
    ambiente = {VAR_CONFIG: str(tmp_path / "config.json")}
    salvar_dados_dir(tmp_path / "dados", ambiente)
    assert dados_dir(ambiente=ambiente) == tmp_path / "dados"
    # a variável de ambiente continua mandando
    assert dados_dir(ambiente={**ambiente, VAR_DADOS: str(tmp_path / "x")}) == tmp_path / "x"


def test_configuracao_fica_no_appdata() -> None:
    assert arquivo_config({"APPDATA": "C:/u"}) == Path("C:/u") / "COIR" / "config.json"
    assert arquivo_config({}) is None


def test_configuracao_ilegivel_falha_alto(tmp_path: Path) -> None:
    arquivo = tmp_path / "config.json"
    arquivo.write_text("{quebrado", encoding="utf-8")
    with pytest.raises(ConfiguracaoError, match="ilegível"):
        dados_dir(ambiente={VAR_CONFIG: str(arquivo)})


def test_trocar_pasta_de_dados_grava_e_zera_a_escolha(tmp_path: Path) -> None:
    gravadas: list[Path] = []
    estado = Estado(tmp_path / "a" / "coir.db")
    estado.titular, estado.ano = 1, 2026
    estado.trocar_dados(tmp_path / "b", salvar=gravadas.append)
    assert estado.banco == tmp_path / "b" / "coir.db"
    assert gravadas == [tmp_path / "b"] and estado.titular is None
    assert estado.pasta_notas == tmp_path / "b" / "_notas-corretagens"


# --------------------------------------------------------------------------------- cadastros
def test_cadastrar_titular_confere_o_cpf(tmp_path: Path) -> None:
    estado = Estado(tmp_path / "coir.db")
    estado.preparar(AGORA)
    with pytest.raises(ImportacaoError, match="CPF inválido"):
        cadastros.adicionar_titular(estado.banco, "111.111.111-11", "Fulano")
    with pytest.raises(ImportacaoError, match="nome"):
        cadastros.adicionar_titular(estado.banco, cpf_de_teste(), "  ")
    cpf = cpf_de_teste()
    id_ = cadastros.adicionar_titular(
        estado.banco, f"{cpf[:3]}.{cpf[3:6]}.{cpf[6:9]}-{cpf[9:]}", " Fulano "
    )
    assert [(t.id, t.nome) for t in estado.titulares()] == [(id_, "Fulano")]
    assert estado.pronto


# ------------------------------------------------------------------------- cópia e importação
def test_copiar_para_pasta_nao_duplica_nem_sobrescreve(tmp_path: Path) -> None:
    pasta = tmp_path / "_notas-corretagens"
    a = copiar_para_pasta([("nota.pdf", b"um")], pasta)
    b = copiar_para_pasta([("nota.pdf", b"um"), ("nota.pdf", b"dois")], pasta)
    assert a == [pasta / "nota.pdf"]
    assert b == [pasta / "nota.pdf", pasta / "nota-2.pdf"]
    assert (pasta / "nota.pdf").read_bytes() == b"um"
    assert (pasta / "nota-2.pdf").read_bytes() == b"dois"


@pytest.mark.parametrize("nome", ["nota.txt", "..\\..\\fora.pdf", "../fora.pdf"])
def test_copiar_para_pasta_nunca_sai_da_pasta(tmp_path: Path, nome: str) -> None:
    pasta = tmp_path / "notas"
    if nome.endswith(".txt"):
        with pytest.raises(ImportacaoError, match="não é PDF"):
            copiar_para_pasta([(nome, b"x")], pasta)
        return
    [destino] = copiar_para_pasta([(nome, b"x")], pasta)
    assert destino.parent == pasta


def test_importar_lista_informa_progresso_e_resultado_por_arquivo(tmp_path: Path) -> None:
    banco = tmp_path / "coir.db"
    estado = Estado(banco)
    estado.preparar(AGORA)
    cpf = cpf_de_teste()
    cadastros.adicionar_titular(banco, cpf, "Titular")
    origens = [FIXTURES / "xp_2026_08_bovespa.txt", FIXTURES / "rico_2015_11_bovespa.txt"]
    copias = copiar_para_pasta(
        [(o.stem + ".pdf", o.read_bytes()) for o in origens], estado.pasta_notas
    )
    passos: list[tuple[int, int, str]] = []
    rel = importar_lista(
        banco,
        copias,
        referencias.carregar(),
        AGORA,
        None,
        extrator(cpf),
        lambda *p: passos.append(p),
    )
    assert passos[0] == (0, 2, "xp_2026_08_bovespa.pdf") and passos[-1] == (2, 2, "")
    linhas = por_arquivo(rel)
    assert [(x.arquivo, x.situacao) for x in linhas] == [
        ("xp_2026_08_bovespa.pdf", "gravado"),
        ("rico_2015_11_bovespa.pdf", "gravado"),
    ]
    assert all(x.notas >= 1 for x in linhas)
    # de novo: arquivos completos são ignorados
    rel2 = importar_lista(banco, copias, referencias.carregar(), AGORA, None, extrator(cpf))
    assert {x.situacao for x in por_arquivo(rel2)} == {"repetido"}


def test_rejeicao_conhecida_e_marcada(tmp_path: Path) -> None:
    from coir.parsing.extrair import Rejeitada
    from coir.servicos.importacao import RelatorioImportacao

    data = datetime(2026, 8, 3).date()
    rel = RelatorioImportacao(
        arquivos_lidos=["a.pdf"],
        rejeitadas=[Rejeitada("a.pdf", (2,), "nota 9: líquido", "9", data)],
    )
    [linha] = por_arquivo(rel, frozenset({("9", data)}))
    assert linha.situacao == "rejeitado" and linha.conhecida
    assert linha.detalhes == ("pág. 2: nota 9: líquido",)


# ------------------------------------------------------------------------- notas e operações
def test_lista_e_detalhe_da_nota_conciliada(demo: Path) -> None:
    lista = notas.listar(demo, 1, 2026)
    assert lista and all(n.data.year == 2026 for n in lista)
    assert lista == sorted(lista, key=lambda n: (n.data, n.id), reverse=True)
    d = notas.detalhe(demo, demo.parent, lista[0].id)
    assert d.conciliada and d.motivo == ""
    assert all(c.confere for c in d.conferencias)
    assert d.pdf is None  # as fixtures são texto: não há PDF na pasta das notas
    assert d.operacoes and d.resumo.numero == lista[0].numero


def test_detalhe_acha_o_pdf_na_pasta_das_notas(demo: Path, tmp_path: Path) -> None:
    nota = notas.listar(demo, 1, None)[0]
    (tmp_path / notas.PASTA_NOTAS).mkdir()
    (tmp_path / notas.PASTA_NOTAS / nota.arquivo).write_bytes(b"%PDF")
    assert notas.detalhe(demo, tmp_path, nota.id).pdf == tmp_path / notas.PASTA_NOTAS / nota.arquivo


def test_pagina_do_pdf_em_png(tmp_path: Path) -> None:
    pdf = tmp_path / "nota.pdf"
    tela = Canvas(str(pdf))
    tela.drawString(72, 720, "NOTA DE CORRETAGEM")
    tela.showPage()
    tela.drawString(72, 720, "pagina 2")
    tela.save()
    png = notas.pagina_png(pdf, 2, largura=300)
    assert png.startswith(b"\x89PNG")
    with pytest.raises(ExtracaoError, match="página 3"):
        notas.pagina_png(pdf, 3)


def test_livro_rateia_os_custos_como_o_motor(demo: Path) -> None:
    linhas = operacoes.livro(demo, 1, 2026)
    assert linhas
    for nota in notas.listar(demo, 1, 2026):
        if nota.segmento != "bovespa":
            continue
        custos = [o.custos for o in linhas if o.nota == nota.numero and o.data == nota.data]
        assert sum((c or Decimal(0) for c in custos), Decimal(0)) == nota.custos


def test_estado_guarda_notas_e_livro_ate_o_banco_mudar(demo: Path) -> None:
    estado = Estado(demo)
    estado.titular, estado.ano = 1, 2026
    assert estado.lista_notas() is estado.lista_notas()
    assert estado.livro() is estado.livro()


def test_nota_de_origem_de_cada_movimento() -> None:
    from coir.servicos.carteira import nota_de

    mapa = {12: 3, 7: 2}
    assert nota_de("op:12,13", mapa) == 3
    assert nota_de("termo:7", mapa) == 2
    assert nota_de("nota:4", mapa) == 4
    assert nota_de("evento:9", mapa) is None
    assert nota_de("vencimento", mapa) is None


def test_posicoes_e_trilha_do_titular(demo: Path) -> None:
    from datetime import date

    from coir.domain.fiscal import Visao
    from coir.servicos.carteira import posicoes, trilha

    estado = Estado(demo)
    ate = date(2026, 10, 16)
    carteira, contas = estado.carteira(ate, Visao.TITULAR)
    linhas = posicoes(carteira, contas, 1)
    assert linhas and all(p.quantidade != 0 and p.corretora == "" for p in linhas)
    passos = trilha(estado.movimentos(ate, Visao.TITULAR), 1, linhas[0], ate)
    assert passos and passos[-1].quantidade_apos == linhas[0].quantidade
    assert passos[-1].custo_apos == linhas[0].custo


# ------------------------------------------------------- registro, preferências e backups
def test_registro_de_atividades_nunca_grava_cpf(tmp_path: Path) -> None:
    import logging

    from coir.config import configurar_registro, ler_registro

    arquivo = configurar_registro(tmp_path)
    try:
        cpf = cpf_de_teste()
        formatado = f"{cpf[:3]}.{cpf[3:6]}.{cpf[6:9]}-{cpf[9:]}"
        logging.getLogger("coir.teste").info("nota do CPF %s importada", formatado)
        logging.getLogger("coir.teste").info("outro %s no texto", cpf)
        for h in logging.getLogger().handlers:
            h.flush()
        linhas = ler_registro(tmp_path)
        assert arquivo == tmp_path / "logs" / "coir.log"
        assert any("nota do CPF *** importada" in x for x in linhas)
        assert not any(cpf[6:9] in x for x in linhas)
    finally:
        raiz = logging.getLogger()
        for h in [h for h in raiz.handlers if getattr(h, "baseFilename", "") == str(arquivo)]:
            raiz.removeHandler(h)
            h.close()


def test_preferencia_de_tema_fica_na_configuracao(tmp_path: Path) -> None:
    from coir.config import ler_preferencia, salvar_preferencia

    ambiente = {VAR_CONFIG: str(tmp_path / "config.json")}
    assert ler_preferencia("tema", ambiente) is None
    salvar_preferencia("tema", "escuro", ambiente)
    salvar_dados_dir(tmp_path / "dados", ambiente)  # não apaga a preferência
    assert ler_preferencia("tema", ambiente) == "escuro"


def test_backup_restaurar_volta_o_banco_e_guarda_o_atual(tmp_path: Path) -> None:
    from coir.erros import BancoError
    from coir.servicos import manutencao

    estado = Estado(tmp_path / "coir.db")
    with pytest.raises(BancoError, match="ainda não há banco"):
        manutencao.fazer_agora(estado.banco, AGORA)
    estado.preparar(AGORA)
    cadastros.adicionar_titular(estado.banco, cpf_de_teste(), "Antes")
    copia = manutencao.fazer_agora(estado.banco, AGORA)
    cadastros.renomear_titular(estado.banco, 1, "Depois")
    depois = datetime(2026, 10, 16, 13, 0, 0)
    atual = manutencao.restaurar(estado.banco, copia, depois)
    assert [t.nome for t in cadastros.titulares(estado.banco)] == ["Antes"]
    assert atual is not None and [b.feito_em for b in manutencao.backups(estado.banco)] == [
        depois, AGORA]  # fmt: skip
    with pytest.raises(BancoError, match="desconhecido"):
        manutencao.restaurar(estado.banco, tmp_path / "outro.db", depois)


def test_cadastros_mostra_contas_e_tickers(demo: Path) -> None:
    [titular] = cadastros.titulares(demo)
    assert titular.cpf == "***.***.247-25" and "XP 0" in titular.contas
    contas = cadastros.contas(demo)
    assert sum(c.notas for c in contas) == len(notas.listar(demo, None, None))
    assert any(t.ticker == "CSMG3" for t in cadastros.tickers(demo))
