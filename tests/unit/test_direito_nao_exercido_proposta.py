"""Proposta da movimentação da B3 para o direito não exercido e para a amortização de FI-Infra e
FIP-IE (parecer do especialista de 09/11/2026)."""

from datetime import date
from pathlib import Path

from coir.db.modelos import MovimentacaoB3
from coir.domain.fiscal import TipoAtivo, TipoEvento
from coir.domain.parametros import TiposAtivo
from coir.servicos.eventos import (
    _fim_do_prazo,  # pyright: ignore[reportPrivateUsage]
    _mapa_proposta,  # pyright: ignore[reportPrivateUsage]
)
from tests.fiscal_cases.apoio import REFS

CABECALHO = ("emissor,isin,tipo,fator,data_com,aprovado_em,isin_emitido,preco,periodo,"
             "data_subscricao,origem,referencia,custo_atribuido\n")  # fmt: skip
FISCAIS = REFS.model_copy(
    update={
        "tipos": TiposAtivo(
            tabela={**REFS.tipos.tabela, "XINF11": TipoAtivo.FUNDO_INFRA, "KNRI11": TipoAtivo.FII}
        )
    }
)


def _mov(movimentacao: str, entrada: bool) -> MovimentacaoB3:
    return MovimentacaoB3(movimentacao=movimentacao, entrada=entrada)


def test_baixa_do_direito_nao_exercido_propoe_o_evento() -> None:
    saida = _mov("Direitos de Subscrição - Não Exercido", entrada=False)
    assert _mapa_proposta(saida, "XPTO1", FISCAIS) == (TipoEvento.DIREITO_NAO_EXERCIDO, None)
    assert _mapa_proposta(_mov("Direitos de Subscrição - Não Exercido", True), "XPTO1",
                          FISCAIS) is None  # fmt: skip


def test_amortizacao_so_em_fi_infra_e_fip_ie() -> None:
    amortizacao = _mov("Amortização", entrada=True)
    assert _mapa_proposta(amortizacao, "XINF11", FISCAIS) == (TipoEvento.REDUCAO_CAPITAL, "V")
    assert _mapa_proposta(amortizacao, "KNRI11", FISCAIS) is None
    assert _mapa_proposta(amortizacao, "XINF11", None) is None


def test_data_e_o_fim_do_prazo_de_subscricao_do_pacote(tmp_path: Path) -> None:
    """Período de negociação até 20/03 e subscrição em 27/03: vale a mais tarde (nunca antecipa
    a perda); a baixa na B3 em 31/03."""
    (tmp_path / "eventos_b3.csv").write_text(
        CABECALHO + "XPTO,BRXPTOACNOR0,SUBSCRICAO,10,2026-02-27,2026-02-20,BRXPTOACNOR0,1,"
        "02/03/2026 a 20/03/2026,2026-03-27,b3,,\n", encoding="utf-8")  # fmt: skip
    data, nota = _fim_do_prazo("XPTO1", date(2026, 3, 31), FISCAIS, tmp_path)
    assert data == date(2026, 3, 27)
    assert "fim do prazo" in nota


def test_sem_subscricao_no_pacote_fica_a_data_da_baixa(tmp_path: Path) -> None:
    (tmp_path / "eventos_b3.csv").write_text(CABECALHO, encoding="utf-8")
    data, nota = _fim_do_prazo("XPTO1", date(2026, 3, 31), FISCAIS, tmp_path)
    assert data == date(2026, 3, 31)
    assert "baixa na B3" in nota


def test_parametro_baixa_b3_usa_a_data_da_baixa(tmp_path: Path) -> None:
    from coir.domain.parametros import ParametrosFiscais

    params = ParametrosFiscais(parametros=tuple(
        p.model_copy(update={"valor": "baixa_b3"}) if p.nome == "direito_vencido_data" else p
        for p in REFS.parametros.parametros))  # fmt: skip
    (tmp_path / "eventos_b3.csv").write_text(
        CABECALHO + "XPTO,BRXPTOACNOR0,SUBSCRICAO,10,2026-02-27,2026-02-20,BRXPTOACNOR0,1,"
        "02/03/2026 a 20/03/2026,2026-03-27,b3,,\n", encoding="utf-8")  # fmt: skip
    data, _ = _fim_do_prazo(
        "XPTO1", date(2026, 3, 31), FISCAIS.model_copy(update={"parametros": params}), tmp_path
    )
    assert data == date(2026, 3, 31)
