"""Avisos de FII e Fiagro na declaração (decisões do usuário de 10/11/2026, leitura conservadora):
compensação cruzada sem norma (IN RFB 1.585/2015 art. 37 §2º), rendimento de Fiagro antes da
isenção (Lei 11.033/2004 art. 3º III, red. Lei 14.130/2021, DOU 11/06/2021) e requisitos da
isenção (art. 3º §1º: 50 cotistas até 2023, 100 desde 2024)."""

from datetime import date

from coir.domain.fiscal import TipoAtivo
from coir.domain.proventos import Provento, TipoProvento
from coir.engine.fiagro import avisos_fii_fiagro
from tests.fiscal_cases.apoio import REFS, D, Livro

FII = TipoAtivo.FII


def _rendimento(dia: date, codigo: str, valor: str) -> Provento:
    return Provento(data=dia, titular=1, codigo=codigo, empresa=codigo, instituicao="XP",
                    tipo=TipoProvento.RENDIMENTO_FII, quantidade=D("100"), liquido=D(valor),
                    bruto=D(valor), irrf=D("0"), aliquota=D("0"), ano_declaracao=dia.year,
                    origem="mov:1")  # fmt: skip


def test_ganho_de_fiagro_e_perda_de_fii_no_mesmo_ano_avisa() -> None:
    """Março/2025: V 100 FIIA11 (FII) a 90,00 sobre PM 100,00 (perda 1.000,00) e V 100 AGRO11
    (Fiagro) a 110,00 sobre PM 100,00 (ganho 1.000,00): o motor compensa, o aviso pede conferir."""
    livro = Livro()
    livro.nota(date(2025, 1, 6), livro.op("C", "FIIA11", 100, "100", tipo=FII),
               livro.op("C", "AGRO11", 100, "100", tipo=FII))  # fmt: skip
    livro.nota(date(2025, 3, 10), livro.op("V", "FIIA11", 100, "90", tipo=FII),
               livro.op("V", "AGRO11", 100, "110", tipo=FII))  # fmt: skip
    ap = livro.apurar(date(2025, 12, 31))
    avisos = avisos_fii_fiagro(1, 2025, ap.movimentos, (), {"AGRO11"}, REFS.parametros)
    assert len(avisos) == 1 and "art. 37 §2º" in avisos[0]
    assert not avisos_fii_fiagro(1, 2025, ap.movimentos, (), set(), REFS.parametros)


def test_rendimento_de_fiagro_antes_de_11_06_2021_nao_e_isento() -> None:
    avisos = avisos_fii_fiagro(1, 2021, (), (_rendimento(date(2021, 5, 14), "AGRO11", "80.00"),
                                             _rendimento(date(2021, 7, 14), "AGRO11", "80.00")),
                               {"AGRO11"}, REFS.parametros)  # fmt: skip
    assert avisos[0].startswith("14/05/2021: rendimento do Fiagro AGRO11 (R$ 80.00) antes")
    assert len([a for a in avisos if "antes da isenção" in a]) == 1
    assert "50 cotistas" in avisos[-1]


def test_requisitos_da_isencao_de_2024_em_diante_pedem_100_cotistas() -> None:
    [aviso] = avisos_fii_fiagro(1, 2024, (), (_rendimento(date(2024, 5, 14), "FIIA11", "80.00"),),
                                set(), REFS.parametros)  # fmt: skip
    assert "100 cotistas" in aviso and "30%" in aviso
