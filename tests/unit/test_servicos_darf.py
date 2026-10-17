"""F6: situação das DARFs, multa até hoje ou até o pagamento e o registro do pagamento."""

from datetime import date, datetime
from decimal import Decimal
from pathlib import Path

import pytest

from coir import referencias
from coir.domain.fiscal import MesApurado
from coir.erros import ImportacaoError
from coir.servicos import darf
from coir.servicos.cadastros import adicionar_titular
from coir.servicos.estado import Estado
from tests.apoio import cpf_de_teste

D = Decimal
Z = D(0)
FISCAIS = referencias.carregar_fiscais()
AGORA = datetime(2026, 4, 2, 10, 0, 0)


def mes(m: int, darf_: str = "0", acumulado: str = "0", vencimento: date | None = None,
        ano: int = 2026) -> MesApurado:  # fmt: skip
    return MesApurado(
        titular=1, ano=ano, mes=m, categorias=(), imposto=D(darf_), irrf_mes=Z,
        irrf_saldo_anterior=Z, irrf_deduzido=Z, irrf_saldo=Z, imposto_devido=D(darf_),
        darf_acumulado_anterior=Z, darf=D(darf_), darf_acumulado=D(acumulado),
        vencimento=vencimento,
    )  # fmt: skip


MESES = [
    mes(12, "0", acumulado="8.20", ano=2025),  # abaixo do mínimo: soma à de janeiro
    mes(1, "120.00", vencimento=date(2026, 2, 27)),  # paga
    mes(2, "91.59", vencimento=date(2026, 3, 31)),  # em atraso em 02/04
    mes(3, "18.20", vencimento=date(2026, 4, 30)),  # ainda no prazo
    mes(4, "12.00"),  # mês corrente (02/04): em apuração, valor parcial
    mes(5),  # sem imposto: fora da lista
]


def test_situacao_e_multa_de_cada_darf() -> None:
    pagos = {(2026, 1): darf.Pagamento(date(2026, 3, 3), D("120.80"), None, "")}
    linhas = darf.darfs(MESES, pagos, AGORA.date(), FISCAIS)
    assert [(x.mes, x.situacao) for x in linhas] == [
        (12, darf.Situacao.ACUMULADA), (1, darf.Situacao.PAGA), (2, darf.Situacao.EM_ATRASO),
        (3, darf.Situacao.A_PAGAR), (4, darf.Situacao.EM_APURACAO)]  # fmt: skip
    acumulada, paga, atrasada, a_pagar, em_apuracao = linhas
    # paga em 03/03 (terça), venceu sexta 27/02: conta de 02/03 (1º dia útil) = 2 dias = 0,66%
    # 120,00 x 0,66% = 0,792 -> 0,79 (truncado); total 120,79; pago 120,80 -> 0,01 a maior
    assert paga.multa is not None and (paga.multa.dias, paga.multa.valor) == (2, D("0.79"))
    assert paga.diferenca == D("0.01") and paga.pago_em_atraso  # a diferença: juros (F7)
    # vence terça 31/03; hoje 02/04: 2 dias; 91,59 x 0,66% = 0,604494 -> 0,60
    assert atrasada.multa is not None and atrasada.multa.valor == D("0.60")
    assert atrasada.total == D("92.19") and atrasada.diferenca is None
    assert acumulada.vencimento is None and acumulada.multa is None
    assert a_pagar.multa is None and a_pagar.total == D("18.20")
    assert em_apuracao.vencimento is None and em_apuracao.total == D("12.00")
    assert {x.codigo for x in linhas} == {"6015"}


@pytest.fixture
def estado(tmp_path: Path) -> Estado:
    e = Estado(tmp_path / "coir.db")
    e.preparar(AGORA)
    adicionar_titular(e.banco, cpf_de_teste(), "Titular")
    return e


def test_registrar_pagamento_com_comprovante_e_desfazer(estado: Estado) -> None:
    darf.registrar_pagamento(estado.banco, estado.dados, 1, 2026, 2, date(2026, 4, 1),
                             D("92.19"), AGORA, ("recibo.pdf", b"%PDF-1.4 x"), " pix ")  # fmt: skip
    [(chave, pago)] = darf.pagamentos(estado.banco, 1).items()
    assert chave == (2026, 2) and pago.valor == D("92.19") and pago.observacao == "pix"
    assert pago.comprovante == "comprovantes/2026/darf-2026-02-titular1.pdf"
    assert (estado.dados / pago.comprovante).read_bytes() == b"%PDF-1.4 x"
    # alterar mantém o comprovante já guardado
    darf.registrar_pagamento(estado.banco, estado.dados, 1, 2026, 2, date(2026, 4, 1),
                             D("92.20"), AGORA)  # fmt: skip
    assert darf.pagamentos(estado.banco, 1)[(2026, 2)].comprovante == pago.comprovante
    darf.excluir_pagamento(estado.banco, 1, 2026, 2)
    assert darf.pagamentos(estado.banco, 1) == {}
    assert (estado.dados / pago.comprovante).is_file()  # o comprovante não é apagado


@pytest.mark.parametrize(
    ("dia", "valor", "comprovante", "erro"),
    [
        (date(2026, 4, 3), "10", None, "futura"),
        (date(2026, 2, 27), "10", None, "anterior ao fim da competência"),
        (date(2026, 4, 1), "0", None, "valor"),
        (date(2026, 4, 1), "10", ("x.pdf", b"nao e pdf"), "não é PDF"),
    ],
)
def test_pagamento_invalido_e_recusado(
    estado: Estado, dia: date, valor: str, comprovante: tuple[str, bytes] | None, erro: str
) -> None:
    with pytest.raises(ImportacaoError, match=erro):
        darf.registrar_pagamento(estado.banco, estado.dados, 1, 2026, 2, dia, D(valor), AGORA,
                                 comprovante)  # fmt: skip


@pytest.mark.parametrize(
    ("texto", "valor"),
    [("1.234,56", "1234.56"), ("1234,56", "1234.56"), ("523.47", "523.47"), ("523.4", "523.4"),
     ("1.234", "1234"), ("1.234.567", "1234567"), ("R$ 92,19", "92.19"), ("92", "92"),
     ("abc", None), ("", None)],
)  # fmt: skip
def test_valor_digitado_no_formato_brasileiro(texto: str, valor: str | None) -> None:
    """Revisão I3: "523.47" (ponto decimal) não pode virar 52.347."""
    from coir.ui.entrada import valor_br

    assert valor_br(texto) == (D(valor) if valor else None)


def test_competencia_e_comprovante_validados(estado: Estado) -> None:
    with pytest.raises(ImportacaoError, match="competência inválida"):
        darf.registrar_pagamento(estado.banco, estado.dados, 1, 2026, 13, date(2026, 4, 1),
                                 D("1"), AGORA)  # fmt: skip
    grande = b"%PDF" + b"0" * darf.LIMITE_COMPROVANTE
    with pytest.raises(ImportacaoError, match="10 MB"):
        darf.registrar_pagamento(estado.banco, estado.dados, 1, 2026, 2, date(2026, 4, 1),
                                 D("1"), AGORA, ("x.pdf", grande))  # fmt: skip
    assert not (estado.dados / darf.PASTA_COMPROVANTES).exists()


def test_pagamento_aparece_nas_darfs_do_estado(estado: Estado) -> None:
    """Revisão I6: o caminho real do cache (sem substituir estado.darfs) vê o pagamento."""
    from coir.domain.fiscal import Apuracao

    estado.agora = lambda: AGORA
    estado.titular, estado.ano = 1, 2026
    apuracao = Apuracao(movimentos=(), meses=tuple(MESES), exclusiva=(), irrf_anual=())
    estado.apuracao = lambda ate: apuracao  # type: ignore[method-assign]
    [antes] = [x for x in estado.darfs() if x.mes == 2]
    assert antes.situacao is darf.Situacao.EM_ATRASO
    darf.registrar_pagamento(estado.banco, estado.dados, 1, 2026, 2, date(2026, 4, 1),
                             D("92.19"), AGORA)  # fmt: skip
    estado.invalidar()
    [depois] = [x for x in estado.darfs() if x.mes == 2]
    assert depois.situacao is darf.Situacao.PAGA and depois.pagamento is not None
    assert {x.ano for x in estado.darfs_todos_os_anos()} == {2025, 2026}


def test_resumo_do_painel() -> None:
    from coir.servicos import painel

    linhas = darf.darfs(MESES, {}, AGORA.date(), FISCAIS)
    r = painel.resumo([m for m in MESES if m.ano == 2026], linhas)
    assert r.imposto_devido == D("241.79")  # 120,00 + 91,59 + 18,20 + 12,00
    assert (r.darfs_pagas, r.darfs_emitidas) == (0, 3)  # jan, fev, mar (fora acumulada e abril)
    vez = painel.da_vez(linhas)
    assert vez is not None and (vez.ano, vez.mes) == (2026, 1)  # a mais antiga em atraso
