"""Bonificação sem custo atribuído (caso UNIP6, 22/04/2022, 60 ações): a proposta da movimentação
da B3 ("Bonificação em ativos") vinha sem valor e a tela gravava assim, sem aviso, e o preço médio
da venda saía errado. Agora a proposta busca o custo no pacote público (mesmo ativo, data ex perto
do crédito) e, sem ele, pede o custo; o que já está gravado sem custo vira aviso na tela Eventos e
no sino. Dados fictícios (banco de demonstração e ISINs de teste)."""

import shutil
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path
from types import SimpleNamespace

import pytest

from coir import referencias
from coir.domain.fiscal import TipoEvento
from coir.recursos import pasta_data
from coir.servicos import eventos
from coir.servicos.b3 import importar_planilhas
from coir.servicos.estado import Estado
from coir.servicos.eventos import DadosEvento, EventoLancado
from tests.unit.test_importador_b3 import CAB_MOV, planilha
from tools.banco_demo import criar

AGORA = datetime(2026, 11, 11, 12, 0, 0)
REFS, FISCAIS = referencias.carregar(), referencias.carregar_fiscais()
AGORA_B3 = "AGORA CORRETORA DE TITULOS E VALORES MOBILIARIOS S/A"
AGORA_CONTA = 1  # GGBR4: 300 ações de 2022 em diante
CAB_EVENTOS = ("emissor,isin,tipo,fator,data_com,aprovado_em,isin_emitido,preco,periodo,"
               "data_subscricao,origem,referencia,custo_atribuido\n")  # fmt: skip
# data com 20/04/2022; 21/04 é feriado (Tiradentes): data ex 22/04/2022, o dia do crédito
UNIP_2022 = "UNIP,BRUNIPTESTE8,BONIFICACAO,10,2022-04-20,2022-04-20,BRUNIPTESTE8,,,,cvm_ipe,IPE 1,"
D = Decimal


@pytest.fixture
def demo(tmp_path: Path) -> Path:
    return criar(tmp_path / "demo")


def _pacote(tmp: Path, *linhas: str) -> Path:
    pasta = tmp / "pacote"
    shutil.copytree(pasta_data(), pasta)
    (pasta / "instrumentos_b3.csv").write_text(
        "ticker,isin,categoria,especificacao,emissor,nome,visto_desde,visto_ate\n"
        "GGBR4,BRGGBRTESTE1,SHARES,PN N1,GGBR,GERDAU,2026-10-20,2026-10-20\n"
        "UNIP6,BRUNIPTESTE8,SHARES,PNB,UNIP,UNIPAR,2026-10-20,2026-10-20\n",
        encoding="utf-8",
    )  # fmt: skip
    (pasta / "eventos_b3.csv").write_text(CAB_EVENTOS + "".join(f"{x}\n" for x in linhas),
                                          encoding="utf-8")  # fmt: skip
    return pasta


def _movimentacao_unip(demo: Path) -> None:
    arquivo = planilha(demo.parent / "movimentacao-2022.xlsx", {"Movimentação": [
        CAB_MOV,
        ["Credito", "22/04/2022", "Bonificação em Ativos", "UNIP6 - UNIPAR CARBOCLORO S.A.",
         AGORA_B3, 60, "-", "-"],
    ]})  # fmt: skip
    importar_planilhas(demo, [arquivo], 1, AGORA)


def _proposta_unip(demo: Path, pasta: Path | None) -> eventos.Proposta:
    [p] = [p for p in eventos.propostas(demo, 1, REFS, FISCAIS if pasta else None, pasta)
           if p.codigo == "UNIP6"]  # fmt: skip
    return p


# --- caminho 1: a importação da movimentação preenche o custo pelo pacote ---------------------


def test_bonificacao_da_movimentacao_sai_com_o_custo_do_pacote(demo: Path) -> None:
    """UNIP6: 60 ações a R$ 24,24 (aviso aos acionistas na CVM) = R$ 1.454,40."""
    _movimentacao_unip(demo)
    p = _proposta_unip(demo, _pacote(demo.parent, UNIP_2022 + "24.24"))
    assert p.movimentacao == "Bonificação em Ativos" and p.dados.tipo is TipoEvento.BONIFICACAO
    assert (p.dados.quantidade, p.dados.valor, p.custo_por_acao) == (60, D("1454.40"), D("24.24"))
    assert p.referencia == "IPE 1"
    assert "R$ 24.24 por ação" in p.nota and "informe o custo" not in p.nota
    assert (p.situacao, p.segura) == ("pendente", False)  # a data ainda é a do crédito: confere


def test_bonificacao_da_movimentacao_sem_custo_no_pacote_fica_pendente(demo: Path) -> None:
    """O pacote tem o evento, mas sem custo: a proposta fica pendente e pede o custo."""
    _movimentacao_unip(demo)
    p = _proposta_unip(demo, _pacote(demo.parent, UNIP_2022))
    assert (p.dados.valor, p.custo_por_acao, p.situacao) == (0, None, "pendente")
    assert p.nota.startswith(eventos.SEM_CUSTO_NO_PACOTE)


def test_bonificacao_do_pacote_longe_do_credito_nao_e_usada(demo: Path) -> None:
    """Outra bonificação do mesmo ativo, anos antes: não é a desta linha da movimentação."""
    _movimentacao_unip(demo)
    antiga = "UNIP,BRUNIPTESTE8,BONIFICACAO,10,2018-12-20,2018-12-20,BRUNIPTESTE8,,,,cvm_ipe,,3.10"
    p = _proposta_unip(demo, _pacote(demo.parent, antiga))
    assert (p.dados.valor, p.custo_por_acao) == (0, None)


def test_dois_custos_diferentes_na_janela_nao_preenchem() -> None:
    dia = date(2022, 4, 22)
    pacote = {"UNIP6": [(dia, D("24.24"), "a"), (dia, D("25"), "b")]}
    assert eventos._custo_da_bonificacao(pacote, "UNIP6", dia) is None
    pacote = {"UNIP6": [(dia, D("24.24"), "a")]}
    assert eventos._custo_da_bonificacao(pacote, "UNIP6", date(2022, 4, 29)) == (D("24.24"), "a")
    assert eventos._custo_da_bonificacao(pacote, "UNIP6", date(2022, 5, 10)) is None


def test_sem_o_pacote_a_proposta_da_movimentacao_pede_o_custo(demo: Path) -> None:
    _movimentacao_unip(demo)
    p = _proposta_unip(demo, None)
    assert (p.dados.valor, p.situacao) == (0, "pendente") and "custo atribuído" in p.nota


# --- caminho 2: aviso das bonificações já gravadas sem custo ----------------------------------


def _bonificacao_ggbr(valor: Decimal) -> DadosEvento:
    return DadosEvento(tipo=TipoEvento.BONIFICACAO, data=date(2023, 6, 1), conta=AGORA_CONTA,
                       ticker="GGBR4", cv="C", quantidade=D(30), valor=valor)  # fmt: skip


def test_bonificacao_gravada_sem_custo_vira_aviso(demo: Path) -> None:
    assert eventos.bonificacoes_sem_custo(demo, 1) == []
    evento_id = eventos.salvar(demo, 1, _bonificacao_ggbr(D(0)), REFS, FISCAIS)
    [sem] = eventos.bonificacoes_sem_custo(demo, 1)
    assert sem.id == evento_id
    aviso = eventos.aviso_sem_custo(sem)
    assert "GGBR4 em 01/06/2023" in aviso and "30 ações" in aviso
    assert "informe o custo atribuído (aviso aos acionistas)" in aviso
    # corrigido com o custo, o aviso some
    eventos.salvar(demo, 1, _bonificacao_ggbr(D("346.20")), REFS, FISCAIS, evento_id=evento_id)
    assert eventos.bonificacoes_sem_custo(demo, 1) == []


def test_ajuste_de_serie_de_opcao_sem_custo_nao_e_avisado() -> None:
    """A B3 ajusta a quantidade da série aberta pela bonificação do objeto: não há custo."""

    def lancado(ticker: str, valor: Decimal) -> EventoLancado:
        dados = DadosEvento(tipo=TipoEvento.BONIFICACAO, data=date(2022, 4, 22), conta=1,
                            ticker=ticker, cv="C", quantidade=D(10), valor=valor)  # fmt: skip
        return EventoLancado(1, "AGORA 0", None, dados)

    lista = [lancado("UNIPD250", D(0)), lancado("UNIP6", D(0)), lancado("GGBR4", D(5))]
    assert [e.dados.ticker for e in eventos.bonificacoes_sem_custo_de(lista)] == ["UNIP6"]


def test_sino_conta_as_bonificacoes_sem_custo(demo: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    estado = Estado(demo)
    estado.preparar(AGORA)
    estado.titular = 1
    monkeypatch.setattr(estado, "propostas", lambda: [])
    monkeypatch.setattr(estado, "ativos_que_sairam", lambda: [])
    assert not [p for p in estado.pendencias() if "custo atribuído" in p.texto]
    eventos.salvar(demo, 1, _bonificacao_ggbr(D(0)), REFS, FISCAIS)
    estado.invalidar()
    [item] = [p for p in estado.pendencias() if "custo atribuído" in p.texto]
    assert (item.quantidade, item.rota) == (1, "/eventos")


def test_aviso_sem_quantidade() -> None:
    dados = SimpleNamespace(ticker="UNIP6", data=date(2022, 4, 22), quantidade=None)
    texto = eventos.aviso_sem_custo(SimpleNamespace(dados=dados, conta_nome="XP 1"))  # type: ignore[arg-type]
    assert texto.startswith("Bonificação de UNIP6 em 22/04/2022 (XP 1) lançada sem custo")
