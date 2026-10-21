"""F8: crédito do JCP proposto pela DIRPF entregue (Bens e Direitos 99.07), dados fictícios."""

from datetime import date
from decimal import Decimal

from coir.domain.proventos import Provento, TipoProvento
from coir.gabarito.dirpf import BemDeclarado, Dirpf
from coir.servicos.declaracao import propor_creditos

D = Decimal
A, B, C = "11111111000111", "22222222000122", "33333333000133"
CNPJ = {"AAAA4": A, "BBBB4": B, "CCCC3": C}


def _jcp(dia: date, codigo: str, liquido: str, chave: str, obs: str = "") -> Provento:
    return Provento(data=dia, titular=1, codigo=codigo, empresa=codigo, instituicao="XP",
                    tipo=TipoProvento.JCP, quantidade=D(100), liquido=D(liquido),
                    bruto=D(liquido), irrf=D(0), aliquota=D("0.15"), ano_declaracao=dia.year,
                    origem="mov:1", observacao=obs, chave=chave)  # fmt: skip


def _a_receber(cnpj: str, anterior: str, atual: str) -> BemDeclarado:
    return BemDeclarado(grupo="99", codigo="07", ticker=None, cnpj=cnpj, discriminacao="JCP",
                        anterior=D(anterior), atual=D(atual))  # fmt: skip


def _dirpf(*bens: BemDeclarado) -> Dirpf:
    return Dirpf(exercicio=2026, ano=2025, renda_variavel=(), fii=(), bens=bens, totais={},
                 rendimentos=(), sem_fonte=(), ir_fonte_lei_11033=D(0))  # fmt: skip


def test_parcelas_que_fecham_o_valor_viram_proposta() -> None:
    """AAAA: 300,00 a receber em 31/12/2024, pagos em duas parcelas em 2025 (100 + 200; a de 50
    é outro JCP). BBBB: 75,00 com duas combinações possíveis (25 + 50 ou 75): sem proposta. CCCC:
    a receber em 31/12/2025, pago em 2026."""
    proventos = [
        _jcp(date(2025, 1, 10), "AAAA4", "100", "a1"),
        _jcp(date(2025, 6, 10), "AAAA4", "200", "a2"),
        _jcp(date(2025, 8, 10), "AAAA4", "50", "a3"),
        _jcp(date(2025, 1, 10), "BBBB4", "25", "b1"),
        _jcp(date(2025, 2, 10), "BBBB4", "50", "b2"),
        _jcp(date(2025, 3, 10), "BBBB4", "75", "b3"),
        _jcp(date(2026, 2, 10), "CCCC3", "40", "c1"),
    ]
    declaracao = _dirpf(_a_receber(A, "300", "0"), _a_receber(B, "75", "0"),
                        _a_receber(C, "0", "40"))  # fmt: skip
    propostas = propor_creditos(proventos, CNPJ, [("dirpf-2026.pdf", declaracao)])
    assert [(x.chave, x.credito) for x in propostas] == [
        ("a1", date(2024, 12, 31)),
        ("a2", date(2024, 12, 31)),
        ("c1", date(2025, 12, 31)),
    ]
    assert {x.origem for x in propostas} == {"dirpf-2026.pdf"}


def test_credito_ja_informado_ou_no_ano_certo_fica_de_fora() -> None:
    informado = _jcp(date(2025, 1, 10), "AAAA4", "300", "a1").model_copy(
        update={"credito": date(2024, 12, 20), "credito_origem": "informada"}
    )
    assert propor_creditos([informado], CNPJ, [("d.pdf", _dirpf(_a_receber(A, "300", "0")))]) == []
    ja_certo = _jcp(date(2025, 1, 10), "AAAA4", "300", "a1").model_copy(
        update={"ano_declaracao": 2024}
    )
    assert propor_creditos([ja_certo], CNPJ, [("d.pdf", _dirpf(_a_receber(A, "300", "0")))]) == []
