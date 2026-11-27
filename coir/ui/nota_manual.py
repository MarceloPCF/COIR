"""Formulário da nota manual (tela Notas): a nota inteira digitada, Bovespa ou BM&F, como a
corretora a imprimiria. Só apresenta: monta, confere e grava pelo ``servicos.nota_manual``."""

from collections.abc import Awaitable, Callable
from dataclasses import dataclass, field
from datetime import date
from decimal import Decimal

from nicegui import run, ui

from coir.domain.formato import reais
from coir.domain.nota import CV, CategoriaCusto, Segmento
from coir.erros import CoirError
from coir.servicos import nota_manual as nm
from coir.servicos.estado import Estado
from coir.ui import mascara, ocupado
from coir.ui.entrada import valor_br

NOVA = 0  # opção "conta nova" no seletor de contas
CUSTOS_BOVESPA = (
    CategoriaCusto.LIQUIDACAO, CategoriaCusto.REGISTRO, CategoriaCusto.EMOLUMENTOS,
    CategoriaCusto.TERMO_OPCOES, CategoriaCusto.ANA, CategoriaCusto.CORRETAGEM,
    CategoriaCusto.IMPOSTOS, CategoriaCusto.CUSTODIA, CategoriaCusto.OUTROS,
)  # fmt: skip
CUSTOS_BMF = (
    CategoriaCusto.TAXAS_BMF, CategoriaCusto.REGISTRO, CategoriaCusto.EMOLUMENTOS,
    CategoriaCusto.CORRETAGEM, CategoriaCusto.IMPOSTOS, CategoriaCusto.OUTROS,
)  # fmt: skip


@dataclass
class _Linha:
    cv: CV = "C"
    mercado: str = "VISTA"  # Bovespa
    ativo: str = ""  # Bovespa: código ou nome de pregão; BM&F: contrato
    prazo: str = ""  # opção: MM/AA; termo: dias; BM&F: vencimento AAAA-MM-DD
    quantidade: str = ""
    preco: str = ""
    day_trade: bool = False  # Bovespa
    tipo: str = "NORMAL"  # BM&F
    ajuste: str = ""  # BM&F, valor da linha
    dc: str = "C"  # BM&F


@dataclass
class _Formulario:
    conta: int = NOVA
    corretora: str = ""
    numero_conta: str = ""
    data: str = ""
    numero: str = ""
    segmento: Segmento = Segmento.BOVESPA
    linhas: list[_Linha] = field(default_factory=list[_Linha])
    custos: dict[CategoriaCusto, str] = field(default_factory=dict[CategoriaCusto, str])
    irrf: str = ""
    irrf_day_trade: str = ""


def _decimal(texto: str, campo: str, vazio_zero: bool = False) -> Decimal:
    if vazio_zero and not texto.strip():
        return Decimal(0)
    valor = valor_br(texto)
    if valor is None:
        raise CoirError(f"{campo}: valor inválido")
    return valor


def _br(valor: Decimal) -> str:
    return format(valor, "f").replace(".", ",")


def _dia(texto: str, campo: str) -> date:
    try:
        return date.fromisoformat(texto)
    except ValueError as erro:
        raise CoirError(f"{campo}: data inválida") from erro


def _de_nota(n: nm.NotaManual) -> _Formulario:
    """A nota gravada de volta nos campos (para corrigir)."""
    linhas = [
        _Linha(cv=x.cv, mercado=x.mercado, ativo=x.ativo, prazo=x.prazo,
               quantidade=_br(x.quantidade), preco=_br(x.preco),
               day_trade=x.day_trade)
        for x in n.bovespa
    ] + [
        _Linha(cv=x.cv, ativo=x.contrato, prazo=x.vencimento.isoformat() if x.vencimento else "",
               quantidade=_br(x.quantidade), preco=_br(x.preco), tipo=x.tipo,
               ajuste=_br(abs(x.ajuste)), dc="D" if x.ajuste < 0 else "C")
        for x in n.bmf
    ]  # fmt: skip
    return _Formulario(
        conta=n.conta_id or NOVA, data=n.data.isoformat(), numero=n.numero,
        segmento=n.segmento, linhas=linhas,
        custos={c: _br(v) for c, v in n.custos.items()},
        irrf=_br(n.irrf) if n.irrf else "",
        irrf_day_trade=_br(n.irrf_day_trade) if n.irrf_day_trade else "",
    )  # fmt: skip


def _nota(f: _Formulario) -> nm.NotaManual:
    """Os campos como ``NotaManual`` (CoirError no primeiro valor que não é número)."""
    if not f.data:
        raise CoirError("informe a data do pregão")
    bovespa = f.segmento is Segmento.BOVESPA
    linhas = [x for x in f.linhas if x.ativo.strip() or x.quantidade.strip()]
    lb: list[nm.LinhaBovespa] = []
    lf: list[nm.LinhaBMF] = []
    for i, x in enumerate(linhas, start=1):
        qtd = _decimal(x.quantidade, f"operação {i}, quantidade")
        preco = _decimal(x.preco, f"operação {i}, preço")
        if bovespa:
            lb.append(
                nm.LinhaBovespa(cv=x.cv, mercado=x.mercado, ativo=x.ativo, quantidade=qtd,
                                preco=preco, prazo=x.prazo, day_trade=x.day_trade)
            )  # fmt: skip
        else:
            ajuste = _decimal(x.ajuste, f"operação {i}, valor")
            vencimento = _dia(x.prazo, f"operação {i}, vencimento") if x.prazo else None
            lf.append(
                nm.LinhaBMF(cv=x.cv, contrato=x.ativo, vencimento=vencimento, quantidade=qtd,
                            preco=preco, ajuste=-ajuste if x.dc == "D" else ajuste, tipo=x.tipo)
            )  # fmt: skip
    categorias = CUSTOS_BOVESPA if bovespa else CUSTOS_BMF
    return nm.NotaManual(
        conta_id=f.conta or None, data=_dia(f.data, "pregão"), numero=f.numero,
        segmento=f.segmento, bovespa=tuple(lb), bmf=tuple(lf),
        custos={c: _decimal(f.custos.get(c, ""), nm.ROTULOS[c], vazio_zero=True)
                for c in categorias},
        irrf=_decimal(f.irrf, "IRRF", vazio_zero=True),
        irrf_day_trade=_decimal(f.irrf_day_trade, "IRRF de day trade", vazio_zero=True),
        corretora=f.corretora, conta=f.numero_conta,
    )  # fmt: skip


async def abrir(
    estado: Estado, ao_gravar: Callable[[int], Awaitable[None] | None], nota_id: int | None = None
) -> None:
    """Abre o formulário (``nota_id``: a nota manual a corrigir) e chama ``ao_gravar`` com o id."""
    titular = estado.titular
    if titular is None:
        ui.notify("Escolha um titular.", type="warning")
        return
    try:
        contas = await run.io_bound(nm.contas_do_titular, estado.banco, titular)
        if nota_id is None:
            numero = await run.io_bound(nm.proximo_numero, estado.banco, titular)
            gravada = None
        else:
            numero = ""
            gravada = await run.io_bound(nm.para_corrigir, estado.banco, titular, nota_id)
    except CoirError as erro:
        ui.notify(str(erro), type="negative")
        return
    if contas is None or numero is None or (nota_id is not None and gravada is None):
        return  # a janela fechou no meio
    if gravada is not None:
        f = _de_nota(gravada)
    else:
        f = _Formulario(conta=next(iter(contas), NOVA), numero=numero, linhas=[_Linha()])
    _dialogo(estado, titular, contas, f, ao_gravar, nota_id)


def _dialogo(
    estado: Estado, titular: int, contas: dict[int, str], f: _Formulario,
    ao_gravar: Callable[[int], Awaitable[None] | None], nota_id: int | None,
) -> None:  # fmt: skip
    titulo = "Corrigir nota manual" if nota_id else "Nova nota manual"
    with (
        ui.dialog().props("persistent") as janela,
        ui.card().classes("w-[1100px] max-w-full gap-3"),
    ):
        ui.label(titulo).classes("coir-titulo text-base")
        ui.label(
            "Digite a nota como a corretora a imprimiu (ou pelos dados do extrato da B3). Ela "
            "entra no cálculo como as notas lidas do PDF. Os totais saem das linhas digitadas: "
            "use Conferir totais e compare com a nota ou o extrato antes de gravar. Se achar o PDF "
            "depois, importe-o e remova esta nota, senão as operações contam duas vezes."
        ).classes("text-sm")
        with ui.row().classes("w-full items-center gap-3"):
            ui.select({**contas, NOVA: "Conta nova..."}, label="Conta").bind_value(
                f, "conta"
            ).props("outlined dense options-dense").classes("w-[220px]").mark("manual-conta")
            with ui.row().classes("gap-3").bind_visibility_from(f, "conta", lambda c: c == NOVA):
                ui.input("Corretora").bind_value(f, "corretora").props("outlined dense").classes(
                    "w-[140px]"
                ).mark("manual-corretora")
                ui.input("Nº da conta").bind_value(f, "numero_conta").props(
                    "outlined dense"
                ).classes("w-[140px]").mark("manual-numero-conta")
            ui.input("Pregão").bind_value(f, "data").props("outlined dense type=date").mark(
                "manual-data"
            )
            ui.input("Nº da nota").bind_value(f, "numero").props("outlined dense").classes(
                "w-[120px]"
            ).mark("manual-numero")
            segmento = ui.radio({Segmento.BOVESPA: "Bovespa", Segmento.BMF: "BM&F"}).bind_value(
                f, "segmento").props("inline dense").mark("manual-segmento")  # fmt: skip

        @ui.refreshable
        def linhas() -> None:
            bovespa = f.segmento is Segmento.BOVESPA
            for i, x in enumerate(f.linhas):
                with ui.row().classes("w-full items-center gap-2 no-wrap"):
                    ui.select({"C": "C", "V": "V"}).bind_value(x, "cv").props(
                        "outlined dense"
                    ).classes("w-[64px]").mark(f"manual-cv-{i}")
                    if bovespa:
                        ui.select(nm.MERCADOS_BOVESPA, label="Mercado").bind_value(
                            x, "mercado"
                        ).props("outlined dense options-dense").classes("w-[190px]")
                        ui.input("Ativo").bind_value(x, "ativo").props("outlined dense").classes(
                            "w-[150px]"
                        ).mark(f"manual-ativo-{i}")
                        ui.input("Venc. (MM/AA) ou prazo").bind_value(x, "prazo").props(
                            "outlined dense"
                        ).classes("w-[150px]")
                    else:
                        ui.input("Contrato").bind_value(x, "ativo").props("outlined dense").classes(
                            "w-[120px]"
                        ).mark(f"manual-ativo-{i}")
                        ui.input("Vencimento").bind_value(x, "prazo").props(
                            "outlined dense type=date"
                        )
                        ui.select({t: t for t in nm.TIPOS_BMF}, label="Tipo").bind_value(
                            x, "tipo"
                        ).props("outlined dense").classes("w-[130px]")
                    ui.input("Quantidade").bind_value(x, "quantidade").props(
                        "outlined dense"
                    ).classes("w-[110px]").mark(f"manual-quantidade-{i}")
                    ui.input("Preço").bind_value(x, "preco").props("outlined dense").classes(
                        "w-[110px]"
                    ).mark(f"manual-preco-{i}")
                    if bovespa:
                        ui.checkbox("DT").bind_value(x, "day_trade").tooltip(
                            "Marcada como day trade (D) na nota"
                        )
                    else:
                        ui.input("Valor").bind_value(x, "ajuste").props("outlined dense").classes(
                            "w-[110px]"
                        )
                        ui.select({"C": "C", "D": "D"}).bind_value(x, "dc").props(
                            "outlined dense"
                        ).classes("w-[64px]")
                    ui.button(icon="delete", on_click=lambda _, i=i: remover_linha(i)).props(
                        "flat round dense"
                    ).tooltip("Tirar a operação")
            ui.button("Operação", icon="add", on_click=nova_linha).props("flat dense").mark(
                "manual-nova-linha"
            )

        def nova_linha() -> None:
            f.linhas.append(_Linha())
            linhas.refresh()

        def remover_linha(i: int) -> None:
            del f.linhas[i]
            linhas.refresh()

        ui.label("Negócios").classes("coir-campo-rotulo")
        with ui.column().classes("w-full gap-1"):
            linhas()

        @ui.refreshable
        def custos() -> None:
            categorias = CUSTOS_BOVESPA if f.segmento is Segmento.BOVESPA else CUSTOS_BMF
            with ui.row().classes("w-full gap-2"):
                for c in categorias:
                    campo = (
                        ui.input(nm.ROTULOS[c], value=f.custos.get(c, ""))
                        .props("outlined dense")
                        .classes("w-[170px]")
                        .mark(f"manual-custo-{c.value}")
                    )
                    campo.on_value_change(lambda e, c=c: f.custos.__setitem__(c, e.value or ""))
                ui.input("IRRF s/ operações").bind_value(f, "irrf").props("outlined dense").classes(
                    "w-[130px]"
                ).mark("manual-irrf")
                ui.input("IRRF day trade").bind_value(f, "irrf_day_trade").props(
                    "outlined dense"
                ).classes("w-[130px]")

        ui.label("Custos e IRRF da nota (em branco = zero)").classes("coir-campo-rotulo")
        custos()
        # só depois de linhas e custos existirem: o valor inicial também dispara a troca
        segmento.on_value_change(lambda: (linhas.refresh(), custos.refresh()))
        previa = ui.label("").classes("coir-subtitulo text-sm").mark("manual-previa")

        def conferir() -> None:
            try:
                n = nm.montar(_nota(f), "", "")
            except CoirError as erro:
                previa.set_text(f"Ainda não fecha: {erro}")
                return
            r = n.resumo_bovespa or n.resumo_bmf
            assert r is not None
            if n.resumo_bovespa is not None:
                total = n.resumo_bovespa.valor_operacoes
            else:
                total = n.resumo_bmf.valor_negocios.assinado if n.resumo_bmf else Decimal(0)
            previa.set_text(
                f"Valor das operações {reais(total)} · custos "
                f"{reais(n.custos_total)} · líquido {reais(r.liquido.assinado)}"
            )

        def confirmar_parecidas(parecidas: list[str]) -> None:
            """Mesma conta e pregão de nota já gravada: pode ser a mesma nota com outro número."""
            with ui.dialog() as aviso, ui.card().classes("w-[720px] max-w-full gap-3"):
                ui.label("Já há nota nesta conta e pregão").classes("coir-titulo text-base")
                ui.label(
                    "Se for a mesma nota com outro número, as operações vão contar duas vezes. "
                    "Confira antes de gravar:"
                ).classes("text-sm")
                for texto in parecidas:
                    ui.label(f"• {texto}").classes("text-sm").mark("manual-parecida")

                async def mesmo_assim() -> None:
                    aviso.close()
                    await gravar(confirmado=True)

                with ui.row().classes("w-full justify-end gap-2"):
                    ui.button("Voltar", on_click=aviso.close).props("flat").mark("parecida-voltar")
                    ui.button("Gravar mesmo assim", icon="save", on_click=mesmo_assim).props(
                        "unelevated"
                    ).mark("parecida-gravar")
            aviso.open()

        async def gravar(confirmado: bool = False) -> None:
            try:
                nota = _nota(f)
            except CoirError as erro:
                ui.notify(str(erro), type="warning")
                return
            try:
                with ocupado.processando("Gravando a nota"):
                    gravada = await run.io_bound(
                        nm.gravar, estado.banco, titular, nota, estado.refs, estado.fiscais,
                        estado.agora(), nota_id, confirmado,
                    )  # fmt: skip
            except nm.NotaParecida as erro:
                confirmar_parecidas(erro.parecidas)
                return
            except CoirError as erro:
                ui.notify(str(erro), type="negative", multi_line=True)
                return
            if gravada is None:
                return
            janela.close()
            estado.invalidar()
            ui.notify(f"Nota {nota.numero.strip()} gravada.", type="positive")
            for aviso in gravada.avisos:
                ui.notify(mascara.texto(aviso), type="warning", multi_line=True)
            resultado = ao_gravar(gravada.nota_id)
            if resultado is not None:
                await resultado

        with ui.row().classes("w-full justify-end gap-2"):
            ui.button("Conferir totais", icon="calculate", on_click=conferir).props("flat").mark(
                "manual-conferir"
            )
            ui.button("Cancelar", on_click=janela.close).props("flat").mark("manual-cancelar")
            ui.button("Gravar", icon="save", on_click=lambda: gravar()).props("unelevated").mark(
                "manual-gravar"
            )
    janela.open()


def confirmar_remocao(
    estado: Estado, nota_id: int, numero: str, ao_remover: Callable[[], Awaitable[None] | None]
) -> None:
    """Pede confirmação e remove a nota manual (a versão fica guardada, com backup antes)."""
    titular = estado.titular
    if titular is None:
        return
    with ui.dialog() as janela, ui.card().classes("w-[560px] max-w-full gap-3"):
        ui.label(f"Remover a nota manual {numero}?").classes("coir-titulo text-base")
        ui.label(
            "A carteira e a apuração serão refeitas sem ela. O COIR faz backup antes e guarda "
            "a nota no histórico. Se a remoção deixar o livro sem fechar (por exemplo, uma venda "
            "que dependia desta compra), nada muda e o motivo aparece na tela."
        ).classes("text-sm")

        async def remover() -> None:
            janela.close()
            try:
                with ocupado.processando("Removendo a nota"):
                    avisos = await run.io_bound(
                        nm.remover, estado.banco, titular, nota_id, estado.refs, estado.fiscais,
                        estado.agora(),
                    )  # fmt: skip
            except CoirError as erro:
                ui.notify(str(erro), type="negative", multi_line=True)
                return
            estado.invalidar()
            ui.notify(f"Nota {numero} removida.", type="positive")
            for aviso in avisos or []:
                ui.notify(mascara.texto(aviso), type="warning", multi_line=True)
            resultado = ao_remover()
            if resultado is not None:
                await resultado

        with ui.row().classes("w-full justify-end gap-2"):
            ui.button("Agora não", on_click=janela.close).props("flat").mark("remover-nao")
            ui.button("Remover", icon="delete", on_click=remover).props(
                "unelevated color=negative"
            ).mark("remover-sim")
    janela.open()
