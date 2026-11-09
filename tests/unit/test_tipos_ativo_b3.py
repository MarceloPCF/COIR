"""Tabela de tipos: fundo imobiliário não pode faltar, e o que tem tratamento à parte não pode
entrar como o papel (parecer do especialista de 09/11/2026). O HGBS11 ficou sem tipo porque o nome
no cadastro ("HEDGE BRASIL SHOPPING FII RESP LIM") não diz "imobiliário", e a apuração parava."""

import csv
from pathlib import Path

import pytest

from coir.domain.fiscal import TipoAtivo
from coir.domain.parametros import TiposAtivo
from coir.erros import ReferenciaError
from tools.referencias import gerar

DATA = Path(__file__).parents[2] / "data"


def _fundo(
    ticker: str, nome: str, especificacao: str = "CI", emissor: str = "XXXX"
) -> dict[str, str]:
    return {"ticker": ticker, "categoria": "FUNDS", "especificacao": especificacao,
            "emissor": emissor, "nome": nome}  # fmt: skip


def _pregao(nome: str, ticker: str, tipo: str = "") -> dict[str, str]:
    return {"nome_pregao": nome, "ticker": ticker, "tipo": tipo}


def _tipos(instrumentos=(), nomes=(), emissores=()) -> dict[str, str]:  # type: ignore[no-untyped-def]
    return {r["ticker"]: r["tipo"] for r in gerar.tipos(list(instrumentos), nomes, emissores)}


@pytest.mark.parametrize(
    ("fundo", "emissores"),
    [
        (_fundo("HGBS11", "HEDGE BRASIL SHOPPING FII RESP LIM"), ()),  # o caso real
        (_fundo("HGLG11", "FDO INV IMOB CSHG LOGISTICA"), ()),
        (_fundo("AAAA11", "NOME QUALQUER LTDA", emissor="AAAA"),
         ({"emissor": "AAAA", "tipo": "FII"},)),
        (_fundo("BRFT11", "BRADESCO FARMTECH FIDC", "CI EA", "BRFT"),
         ({"emissor": "BRFT", "tipo": "FIAGRO"},)),  # Fiagro é FII em qualquer forma
    ],
)  # fmt: skip
def test_fundo_imobiliario_ou_fiagro_e_fii(fundo: dict[str, str], emissores: tuple) -> None:  # type: ignore[type-arg]
    assert _tipos([fundo], emissores=emissores) == {fundo["ticker"]: "FII"}


@pytest.mark.parametrize(
    "fundo",
    [
        _fundo("AATH11", "ATHON ENERGIA ESG I FDO INV PART EM INFRAESTRUTURA"),  # FIP-IE
        _fundo("BDIF11", "BTG PACTUAL DIVIDA INFRA FIC. FDO. INC. IE. RF."),  # FI-Infra
        _fundo("BBVH12", "FIDC BB VOTORANTIM HIGHLAND INFRA", "FIDC"),
    ],
)
def test_o_que_tem_tratamento_a_parte_nao_vira_fii(fundo: dict[str, str]) -> None:
    assert _tipos([fundo]) == {}


def test_recibo_de_cota_no_cadastro_e_direito_nao_fii() -> None:
    """Recibo de cota de FII: 15%, fora do limite (posição B do parecer de 09/11/2026)."""
    recibo = _fundo("XXXX15", "ALGUM FII RESP LIM", "REC", "XXXX")
    assert _tipos([recibo], emissores=[{"emissor": "XXXX", "tipo": "FII"}]) == {"XXXX15": "DIREITO"}


def test_fundo_fora_do_fii_nao_vira_fii_pelo_cotahist() -> None:
    """O COTAHIST põe FIP e FI-Infra junto dos FII (BDI 12): o cadastro decide."""
    assert _tipos([_fundo("AATH11", "FIP ATHON")], [_pregao("FIP ATHON CI", "AATH11", "FII")]) == {}


@pytest.mark.parametrize(
    ("nome", "ticker", "tipo"),
    [
        ("FII ALMIRANT CI", "FAMB11B", "FII"),
        ("FII ANH EDUC CI ERS", "FAED11B", "FII"),
        ("FIAGRO BRFT CI EA", "BRFT11", "FII"),
        ("FII REC RECEBIV CI", "RRRR11", "FII"),  # "REC" no nome do fundo não é recibo
        ("3M DRN", "MMMC11B", "BDR"),
        ("CEEE-D ON *", "CEED3B", "ACAO"),
        ("CONST BETER PNA*", "COBE5B", "ACAO"),
        ("TAESA UNT", "TAEE11", "UNIT"),
    ],
)
def test_codigo_fora_do_cadastro_pelo_nome_de_pregao(nome: str, ticker: str, tipo: str) -> None:
    assert _tipos(nomes=[_pregao(nome, ticker)]) == {ticker: tipo}


@pytest.mark.parametrize(
    ("nome", "ticker"),
    [
        ("FII HEDGEBS DIR", "HGBS12"),
        ("3R PETROLEUM DIR ORD", "RRRP1"),
        ("3R PETROLEUM ON REC", "RRRP9"),
        ("AES TIETE E DIR UNT", "TIET12"),
        ("AZEVEDO BNS", "AZEV11"),
    ],
)
def test_direito_recibo_e_bonus_sao_direito(nome: str, ticker: str) -> None:
    """SC Cosit 145 e 166/2021: 15%, sem isenção e fora do limite, nunca o tipo do papel."""
    assert _tipos(nomes=[_pregao(nome, ticker)]) == {ticker: "DIREITO"}


def test_direito_de_fii_nao_e_fii_mesmo_com_o_cotahist_dizendo_fii() -> None:
    """O COTAHIST põe o direito de cota com os FII (BDI 12); a 20% seria imposto errado."""
    assert _tipos(nomes=[_pregao("FII BC FUND DIR CI", "BRCR12", "FII")]) == {"BRCR12": "DIREITO"}
    nomes = [_pregao("FII BC FUND CI", "BRCR12", "FII"), _pregao("FII BC FUND DIR CI", "BRCR12")]
    assert _tipos(nomes=nomes) == {"BRCR12": "DIREITO"}  # cota de FII é sempre final 11


def test_rec_de_recebiveis_no_nome_do_fii_nao_e_recibo() -> None:
    assert _tipos(nomes=[_pregao("FII BLUE REC CI", "BLUR11", "FII")]) == {"BLUR11": "FII"}


@pytest.mark.parametrize(
    ("nome", "ticker"),
    [
        ("FIP BKO BREI CI", "BKOI11"),
        ("TREND ASIA CI", "ASIA11"),  # ETF antigo: pode ser renda fixa, sem palpite
    ],
)
def test_o_que_nao_se_sabe_fica_de_fora(nome: str, ticker: str) -> None:
    assert _tipos(nomes=[_pregao(nome, ticker)]) == {}


def test_nomes_que_dao_dois_tipos_nao_dao_tipo_nenhum() -> None:
    nomes = [_pregao("TAESA UNT", "ZZZZ11"), _pregao("FII ZZZZ CI", "ZZZZ11")]
    assert _tipos(nomes=nomes) == {}


def test_cadastro_vale_sobre_o_cotahist() -> None:
    fundo = _fundo("HGBS11", "HEDGE BRASIL SHOPPING FII RESP LIM")
    assert _tipos([fundo], [_pregao("HEDGE UNT", "HGBS11")]) == {"HGBS11": "FII"}


def _csv(nome: str) -> list[dict[str, str]]:
    with (DATA / nome).open(encoding="utf-8", newline="") as arquivo:
        return list(csv.DictReader(arquivo))


def test_nenhum_fundo_imobiliario_das_tabelas_fica_sem_tipo() -> None:
    """Tabela real: o que a lista oficial de fundos da B3 ou o COTAHIST dizem que é FII tem tipo."""
    com_tipo = {r["ticker"]: r["tipo"] for r in _csv("tipos_ativo_b3.csv")}
    fundos = {e["emissor"] for e in _csv("emissores_b3.csv") if e["tipo"] in ("FII", "FIAGRO")}
    faltam = [r["ticker"] for r in _csv("instrumentos_b3.csv")
              if r["categoria"] == "FUNDS" and r["emissor"] in fundos
              and r["especificacao"].split()[:1] == ["CI"]
              and r["ticker"] not in com_tipo]  # fmt: skip
    assert faltam == []
    faltam = [n["ticker"] for n in _csv("nomes_pregao_b3.csv")
              if n["nome_pregao"].split()[0] == "FII" and " CI" in n["nome_pregao"]
              and " DIR " not in n["nome_pregao"]  # "FII MAXI REN DIR CI": direito de subscrição
              and n["ticker"] not in com_tipo]  # fmt: skip
    assert faltam == []
    assert com_tipo["HGBS11"] == com_tipo["BRCR11"] == "FII"
    assert com_tipo["BRCR12"] == com_tipo["HGBS12"] == "DIREITO"


def test_tipo_e_estrito_e_a_apuracao_usa_a_versao_que_isola() -> None:
    tipos = TiposAtivo(tabela={"HGLG11": TipoAtivo.FII})
    assert tipos.tipo_ou_sem("HGLG11") is TipoAtivo.FII
    assert tipos.tipo_ou_sem("PETR4") is TipoAtivo.ACAO
    assert tipos.tipo_ou_sem("ZZZZ11") is TipoAtivo.SEM_TIPO
    with pytest.raises(ReferenciaError):
        tipos.tipo("ZZZZ11")
