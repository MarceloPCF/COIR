"""Tela Simulação: ordens hipotéticas (vendas e compras à vista) sobre a carteira apurada, com o
efeito no imposto, no limite de isenção, no prejuízo a compensar e no day trade. Dois cenários
(A e B) com as mesmas ordens e datas próprias: a ordem sem data fica fora do cenário (execução
parcial), e datas em meses diferentes comparam vender de uma vez ou aos poucos. Nada é gravado."""

from collections.abc import Callable
from dataclasses import dataclass, field
from datetime import date
from decimal import Decimal
from functools import partial

from nicegui import run, ui

from coir.domain.fiscal import Categoria
from coir.erros import CoirError
from coir.reports.fiscais import MESES
from coir.reports.modelo import Coluna, TipoColuna, Valor, formatar
from coir.servicos import simulacao
from coir.servicos.estado import Estado
from coir.servicos.simulacao import Contexto, OrdemDigitada, Regras, ResultadoSimulacao
from coir.ui import ocupado, tabela
from coir.ui.entrada import valor_br

T = TipoColuna
ZERO = Decimal(0)
CENARIOS = ("A", "B")
AVISO_GERAL = (
    "Simulação informativa. Mostra o efeito no imposto de renda das ordens que você digitou, com "
    "as regras fiscais vigentes hoje e as notas já importadas. Nada é gravado na apuração. Não é "
    "recomendação de compra, venda ou manutenção de valores mobiliários, nem consultoria de "
    "investimento ou análise de valores mobiliários, e o COIR não sugere operações. O imposto real "
    "depende dos preços, quantidades, custos e retenções das notas de corretagem e das regras em "
    "vigor na data de cada venda. Não substitui um profissional habilitado."
)
HIPOTESES = (
    "Preços e quantidades são os que você digitou. A ordem pode não ser executada, ou ser "
    "executada em parte ou a outro preço.",
    "O preço médio é o da carteira apurada com as notas importadas até {ultima}. Compras, vendas, "
    "transferências e eventos (bonificação, desdobramento, grupamento, incorporação, cisão) até a "
    "data simulada mudam a quantidade e o preço médio.",
    "As regras são as vigentes hoje: {comum} em operações comuns, {dt} em day trade, {fii} em FII "
    "e Fiagro, isenção de ações até {limite} de vendas no mês. Lei nova pode mudar as regras de "
    "meses futuros; vale a lei da data de cada venda.",
    "O limite de {limite} é do valor bruto de todas as vendas de ações do mês (data do "
    "pregão), somando todas as corretoras e contas do titular (Lei 11.033/2004 art. 3º I). Custos "
    "da corretora não reduzem esse valor. As vendas de ações em day trade, por exercício de opção "
    "e na liquidação de termo entram na soma (critério do COIR, leitura literal da lei), mas o "
    "ganho delas nunca é isento.",
    "Units, ETF, BDR, FII, Fiagro, direitos de subscrição, opções e futuros não entram na soma do "
    "limite e não têm isenção (units: critério do COIR).",
    "Prejuízo só compensa ganho da mesma categoria: operações comuns, day trade ou FII e Fiagro. "
    "Ganho isento não consome prejuízo.",
    "Imposto abaixo de {darf} não gera DARF: soma-se ao dos meses seguintes e vence com o mês em "
    "que o total chega a {darf} (Lei 9.430/1996 art. 68). A DARF vence no último dia útil do mês "
    "seguinte ao da venda.",
    "As ordens simuladas não têm IRRF deduzido (veja o valor estimado no quadro).",
    "Vendas em anos diferentes: o saldo de IRRF não passa de dezembro para janeiro (o de 0,005% "
    "vai para a declaração anual; o de day trade, só por restituição). Desde 2026, os ganhos em "
    "bolsa, inclusive os isentos, entram na base do imposto mínimo de quem recebe mais de "
    "R$ 600 mil no ano (Lei 9.250/1995 art. 16-A, incluído pela Lei 15.270/2025).",
    "Opções, termos e futuros em aberto que vencem até o fim da simulação são tratados como "
    "vencidos nos dois lados da comparação, sem os ajustes que ainda não existem: os valores do "
    "mês ficam hipotéticos, mas a diferença das ordens continua certa.",
)
DAY_TRADE_FIXO = (
    "Day trade é a compra e a venda do mesmo ativo no mesmo dia, na mesma corretora, ainda que em "
    "contas diferentes dela. As ações que você já tinha não contam: a venda casa primeiro com as "
    "compras do dia, pela ordem dos negócios. Compra e venda em corretoras diferentes não são day "
    "trade, nem o exercício de opção seguido de venda no mesmo dia. O ganho de day trade paga "
    "{dt}, não tem isenção, sofre retenção de {irrf_dt} na fonte, e o prejuízo só compensa ganho "
    "de day trade "
    "(Lei 9.959/2000 art. 8º; Lei 11.033/2004 art. 2º; IN RFB 1.585/2015 art. 65)."
)
CUSTOS = (
    "Corretagem, emolumentos e taxas reduzem o ganho (IN RFB 1.585/2015 art. 56 §3º), mas não o "
    "valor que conta para o limite de {limite}. Sem custos informados, o imposto simulado sai "
    "um pouco maior que o real."
)
COLUNAS_MES = (
    Coluna("Mês"),
    Coluna("Vendas de ações (soma do limite)", T.DINHEIRO),
    Coluna("Isento", T.DINHEIRO),
    Coluna("Comum", T.DINHEIRO),
    Coluna("Day trade", T.DINHEIRO),
    Coluna("FII", T.DINHEIRO),
    Coluna("Imposto", T.DINHEIRO),
    Coluna("DARF", T.DINHEIRO),
    Coluna("Vencimento", T.DATA),
)
COLUNAS_ORDEM = (
    Coluna("Ordem", T.INTEIRO),
    Coluna("Ativo"),
    Coluna("C/V"),
    Coluna("Qtd.", T.QUANTIDADE),
    Coluna("Valor", T.DINHEIRO),
    Coluna("Resultado comum", T.DINHEIRO),
    Coluna("Qtd. em day trade", T.QUANTIDADE),
    Coluna("Resultado day trade", T.DINHEIRO),
    Coluna("Qtd. após", T.QUANTIDADE),
    Coluna("PM após", T.PM),
)
COLUNAS_POSICOES = (
    Coluna("Ativo"),
    Coluna("Conta"),
    Coluna("Quantidade", T.QUANTIDADE),
    Coluna("PM na conta", T.PM),
)


def _r(v: Decimal) -> str:
    return formatar(v, T.DINHEIRO)


def _q(v: Decimal) -> str:
    return formatar(v, T.QUANTIDADE)


def _pct(v: Decimal) -> str:
    texto = f"{v * 100:f}"
    if "." in texto:
        texto = texto.rstrip("0").rstrip(".")
    return texto.replace(".", ",") + "%"


def valores(r: Regras) -> dict[str, str]:
    """Os parâmetros fiscais nos textos da tela (nada fixo no texto)."""
    return {"comum": _pct(r.aliquota_comum), "dt": _pct(r.aliquota_day_trade),
            "fii": _pct(r.aliquota_fii), "limite": f"R$ {_r(r.limite_isencao)}",
            "darf": f"R$ {_r(r.darf_minimo)}", "irrf_dt": _pct(r.irrf_day_trade)}  # fmt: skip


@dataclass
class _Ordem:
    ticker: str = ""
    conta: int | None = None
    cv: str = "V"
    quantidade: str = ""
    preco: str = ""
    custos: str = ""
    datas: dict[str, str] = field(default_factory=lambda: {c: "" for c in CENARIOS})


def _data(texto: str) -> date | None:
    try:
        return date.fromisoformat(texto) if texto else None
    except ValueError:
        return None


def digitadas(ordens: list[_Ordem]) -> dict[str, list[OrdemDigitada]]:
    """As linhas da tela em ordens de cada cenário; a ordem sem data fica fora do cenário. Erro de
    digitação levanta ValueError com a mensagem para o usuário."""
    saida: dict[str, list[OrdemDigitada]] = {c: [] for c in CENARIOS}
    for n, o in enumerate(ordens, start=1):
        datas = {c: _data(o.datas[c]) for c in CENARIOS}
        if not any(o.datas.values()):
            continue
        if any(o.datas[c] and datas[c] is None for c in CENARIOS):
            raise ValueError(f"ordem {n}: data inválida")
        quantidade, preco = valor_br(o.quantidade or ""), valor_br(o.preco or "")
        custos = valor_br(o.custos or "0")
        if quantidade is None or preco is None or custos is None:
            raise ValueError(f"ordem {n}: quantidade, preço ou custos inválidos")
        if o.conta is None:
            raise ValueError(f"ordem {n}: escolha a conta")
        for c in CENARIOS:
            dia = datas[c]
            if dia is not None:
                saida[c].append(OrdemDigitada(
                    numero=n, ticker=o.ticker, conta=o.conta,
                    cv="C" if o.cv == "C" else "V", quantidade=quantidade, preco=preco,
                    custos=custos, data=dia,
                ))  # fmt: skip
    return saida


def _linhas_mes(r: ResultadoSimulacao) -> list[tuple[Valor, ...]]:
    linhas: list[tuple[Valor, ...]] = []
    for m in r.meses:
        c = {x.categoria: x for x in m.cenario.categorias}
        comum = c[Categoria.COMUM]
        linhas.append((
            f"{MESES[m.mes - 1]}/{m.ano}", comum.vendas_limite, comum.isento, comum.resultado,
            c[Categoria.DAY_TRADE].resultado, c[Categoria.FII].resultado, m.cenario.imposto,
            m.cenario.darf, m.cenario.vencimento,
        ))  # fmt: skip
    return linhas


def _quadro(base: ResultadoSimulacao, resultados: dict[str, ResultadoSimulacao]) -> None:
    """Os números lado a lado (sem as ordens, cenário A, cenário B), sem rótulo de melhor."""
    nomes = ["Sem as ordens", *(f"Cenário {c}" for c in resultados)]
    linhas: list[tuple[str, list[str]]] = [
        ("Imposto no período", [_r(base.total("imposto", False)),
                                *(_r(r.total("imposto")) for r in resultados.values())]),
        ("DARF no período", [_r(base.total("darf", False)),
                             *(_r(r.total("darf")) for r in resultados.values())]),
        ("Ganho isento", [_r(base.isento(False)), *(_r(r.isento()) for r in resultados.values())]),
    ]  # fmt: skip
    for cat, nome in ((Categoria.COMUM, "comum"), (Categoria.DAY_TRADE, "day trade"),
                      (Categoria.FII, "FII")):  # fmt: skip
        linhas.append((f"Prejuízo a compensar no fim ({nome})",
                       [_r(base.prejuizo_final(False)[cat]),
                        *(_r(r.prejuizo_final()[cat]) for r in resultados.values())]))  # fmt: skip
    linhas.append(("IRRF estimado nas ordens (não deduzido)",
                   ["—", *(_r(r.irrf_estimado) for r in resultados.values())]))  # fmt: skip
    por_ticker = [{p.ticker: p for p in r.posicoes} for r in resultados.values()]
    for ticker in sorted({t for d in por_ticker for t in d}):
        de_base = next(d[ticker] for d in por_ticker if ticker in d)
        valores = [f"{_q(de_base.quantidade_base)} / {_r(de_base.custo_base)}"]
        for d in por_ticker:
            p = d.get(ticker)
            valores.append("sem ordem" if p is None else f"{_q(p.quantidade)} / {_r(p.custo)}")
        linhas.append((f"{ticker}: quantidade e custo que ficam", valores))
    colunas = ["", *nomes]
    with ui.grid(columns=f"minmax(240px, 2fr) repeat({len(nomes)}, minmax(140px, 1fr))").classes(
        "coir-folha w-full gap-x-4 gap-y-1 items-center").mark("sim-comparacao"):  # fmt: skip
        for rotulo in colunas:
            ui.label(rotulo).classes("coir-campo-rotulo")
        for rotulo, valores in linhas:
            ui.label(rotulo)
            for v in valores:
                ui.label(v).classes("coir-numero text-right")


def _day_trade(r: ResultadoSimulacao, nome: str, contas: dict[int, str],
               ordens: list[OrdemDigitada], v: dict[str, str]) -> list[str]:  # fmt: skip
    por_numero = {o.numero: o for o in ordens}
    textos: list[str] = []
    for o in r.ordens:
        if not o.quantidade_day_trade:
            continue
        d = por_numero[o.numero]
        resto = o.quantidade_comum
        textos.append(
            f"Cenário {nome}, ordem {o.numero}: em {d.data:%d/%m/%Y}, {_q(o.quantidade_day_trade)} "
            f"{o.ticker} formam day trade com negócio do mesmo dia na {contas.get(d.conta, '')}: "
            f"resultado de {_r(o.resultado_day_trade)} no grupo de day trade, pelo preço do dia e "
            f"não pelo preço médio, tributado a {v['dt']}, sem isenção e com IRRF de "
            f"{v['irrf_dt']} sobre o ganho do dia."
            + (f" As outras {_q(resto)} são operação comum, pelo preço médio." if resto else "")
            + " A simulação supõe a ordem executada depois dos negócios já feitos no dia."
        )  # fmt: skip
    return textos


def montar(estado: Estado, ao_mudar: Callable[[], object]) -> None:
    del ao_mudar  # só leitura: nada muda no banco
    hoje = estado.agora().date()
    v = valores(simulacao.regras(estado.fiscais, hoje))
    ui.label("Simulação").classes("coir-titulo-pagina")
    ui.label(
        "Digite ordens de venda ou de compra à vista e veja o efeito no imposto, no limite de "
        f"isenção de {v['limite']}, no prejuízo a compensar e no day trade. Os cenários A e B "
        "usam as mesmas ordens com datas próprias: deixe a data em branco para a ordem ficar fora "
        "do cenário, ou use meses diferentes para comparar."
    ).classes("coir-subtitulo")
    ui.label(AVISO_GERAL).classes("coir-carimbo text-sm").mark("sim-aviso")
    if not estado.existe or estado.titular is None:
        ui.label("Cadastre o titular e importe as notas primeiro.").classes("coir-vazio w-full")
        return
    titular = estado.titular
    ordens: list[_Ordem] = []
    ctx_box: list[Contexto] = []
    corpo = ui.column().classes("w-full gap-2")
    with corpo:
        ui.label("Lendo a carteira…").classes("coir-vazio")
    with ui.row().classes("w-full items-center gap-2"):
        botao_adicionar = ui.button("Adicionar ordem", icon="add").props("flat").mark(
            "sim-adicionar")  # fmt: skip
        botao_copiar = ui.button("Copiar as datas de A para B", icon="content_copy").props(
            "flat").mark("sim-copiar")  # fmt: skip
        ui.space()
        botao_simular = ui.button("Simular", icon="play_arrow").props("unelevated").mark(
            "sim-simular")  # fmt: skip
    grade_ordens = ui.column().classes("w-full gap-1")
    ui.label(CUSTOS.format(**v)).classes("coir-subtitulo text-xs")
    with ui.expansion("Hipóteses desta simulação").classes("w-full").mark("sim-hipoteses"):
        lista_hipoteses = ui.column().classes("gap-1")
    resultado = ui.column().classes("w-full gap-3")

    def contas() -> dict[int, str]:
        return {c.id: c.rotulo for c in ctx_box[0].contas} if ctx_box else {}

    def hipoteses() -> None:
        lista_hipoteses.clear()
        ultima = (f"{ctx_box[0].ultima_nota:%d/%m/%Y}"
                  if ctx_box and ctx_box[0].ultima_nota else "a última nota")  # fmt: skip
        with lista_hipoteses:
            for h in HIPOTESES:
                ui.label("• " + h.format(ultima=ultima, **v)).classes("text-sm")

    def adicionar(ordem: _Ordem | None = None) -> None:
        nova = ordem or _Ordem()
        if nova.conta is None and len(contas()) == 1:
            nova.conta = next(iter(contas()))
        padrao = max(hoje, ctx_box[0].ultima_nota or hoje) if ctx_box else hoje
        nova.datas = {c: padrao.isoformat() for c in CENARIOS}
        ordens.append(nova)
        desenhar_ordens()

    def remover(k: int) -> None:
        ordens.pop(k)
        desenhar_ordens()

    def copiar_datas() -> None:
        for o in ordens:
            o.datas["B"] = o.datas["A"]
        desenhar_ordens()

    def desenhar_ordens() -> None:
        grade_ordens.clear()
        with grade_ordens:
            if not ordens:
                ui.label("Nenhuma ordem. Clique numa posição ou em Adicionar ordem.").classes(
                    "coir-vazio w-full")  # fmt: skip
                return
            with ui.grid(columns="40px 110px 120px 200px 120px 110px 110px 160px 160px 40px"
                         ).classes("coir-folha w-full gap-x-2 gap-y-1 items-center"):  # fmt: skip
                for rotulo in ("Nº", "C/V", "Ativo", "Conta", "Quantidade", "Preço", "Custos",
                               "Data no cenário A", "Data no cenário B", ""):  # fmt: skip
                    ui.label(rotulo).classes("coir-campo-rotulo")
                for k, o in enumerate(ordens):
                    linha_ordem(k, o)

    def linha_ordem(k: int, o: _Ordem) -> None:
        n = k + 1
        ui.label(str(n))
        ui.select({"V": "Venda", "C": "Compra"}, value=o.cv,
                  on_change=lambda e, o=o: setattr(o, "cv", e.value)).props(
            "outlined dense options-dense").mark(f"sim-cv-{n}")  # fmt: skip
        ui.input(value=o.ticker, on_change=lambda e, o=o: setattr(o, "ticker", e.value or "")
                 ).props("outlined dense").mark(f"sim-ticker-{n}")  # fmt: skip
        ui.select(contas(), value=o.conta,
                  on_change=lambda e, o=o: setattr(o, "conta", e.value)).props(
            "outlined dense options-dense").mark(f"sim-conta-{n}")  # fmt: skip
        for campo in ("quantidade", "preco", "custos"):
            ui.input(value=getattr(o, campo),
                     on_change=lambda e, o=o, campo=campo: setattr(o, campo, e.value or "")
                     ).props("outlined dense input-class=text-right").mark(
                f"sim-{campo}-{n}")  # fmt: skip
        for c in CENARIOS:
            ui.input(value=o.datas[c],
                     on_change=lambda e, o=o, c=c: o.datas.__setitem__(c, e.value or "")
                     ).props("outlined dense type=date clearable").mark(
                f"sim-data-{c.lower()}-{n}")  # fmt: skip
        ui.button(icon="close", on_click=partial(remover, k)).props(
            "flat round dense").tooltip("Tirar a ordem").mark(f"sim-remover-{n}")  # fmt: skip

    async def simular() -> None:
        try:
            cenarios = digitadas(ordens)
        except ValueError as erro:
            ui.notify(str(erro), type="warning")
            return
        if not any(cenarios.values()):
            ui.notify("Dê uma data a pelo menos uma ordem.", type="warning")
            return
        botao_simular.disable()
        try:
            with ocupado.processando("Simulando as ordens"):
                resultados = await run.io_bound(
                    simulacao.simular_banco, estado.banco, estado.refs, estado.fiscais, titular,
                    cenarios, estado.apuracao,
                )  # fmt: skip
        except CoirError as erro:
            if not botao_simular.is_deleted:
                ui.notify(f"Simulação não feita: {erro}", type="negative", multi_line=True)
            return
        finally:
            if not botao_simular.is_deleted:
                botao_simular.enable()
        if resultados is not None and not resultado.is_deleted:
            mostrar(resultados, cenarios)

    def mostrar(resultados: dict[str, ResultadoSimulacao],
                cenarios: dict[str, list[OrdemDigitada]]) -> None:  # fmt: skip
        resultado.clear()
        if not resultados:
            return
        base = next(iter(resultados.values()))
        with resultado:
            ui.label(f"Resultado até {base.ate:%m/%Y}").classes("coir-titulo text-base")
            _quadro(base, resultados)
            for nome, r in resultados.items():
                ui.label(f"Cenário {nome}: mês a mês e ordens").classes("coir-titulo text-base")
                tabela.grade(COLUNAS_MES, _linhas_mes(r)).mark(f"sim-meses-{nome}")
                tabela.grade(COLUNAS_ORDEM, [
                    (o.numero, o.ticker, o.cv, o.quantidade, o.valor, o.resultado,
                     o.quantidade_day_trade, o.resultado_day_trade, o.quantidade_apos, o.pm_apos)
                    for o in r.ordens
                ]).mark(f"sim-ordens-{nome}")  # fmt: skip
            textos = [t for nome, r in resultados.items()
                      for t in _day_trade(r, nome, contas(), cenarios[nome], v)]  # fmt: skip
            with ui.column().classes("w-full gap-1").mark("sim-day-trade"):
                ui.label("Day trade").classes("coir-titulo text-base")
                for t in textos:
                    ui.label(t).classes("coir-carimbo text-sm")
                ui.label(DAY_TRADE_FIXO.format(**v)).classes("coir-subtitulo text-xs")
            avisos = sorted({a for r in resultados.values() for a in r.avisos})
            if avisos:
                with ui.column().classes("w-full gap-1").mark("sim-avisos"):
                    ui.label("Avisos que só aparecem com as ordens").classes(
                        "coir-titulo text-base")  # fmt: skip
                    for a in avisos:
                        ui.label(a).classes("text-sm")

    async def carregar() -> None:
        try:
            ctx = await run.io_bound(simulacao.contexto, estado.banco, estado.refs, estado.fiscais,
                                     titular, hoje)  # fmt: skip
        except CoirError as erro:
            if not corpo.is_deleted:
                corpo.clear()
                with corpo:
                    ui.label(f"Não foi possível ler a carteira: {erro}").classes("coir-vazio")
            return
        if ctx is None or corpo.is_deleted:
            return
        ctx_box.append(ctx)
        hipoteses()
        posicoes = ctx.posicoes
        nomes = contas()

        def de_posicao(chave: object) -> None:
            p = posicoes[int(str(chave))]
            adicionar(_Ordem(ticker=p.ticker, conta=p.conta, quantidade=_q(p.quantidade)))

        corpo.clear()
        with corpo:
            ultima = f"{ctx.ultima_nota:%d/%m/%Y}" if ctx.ultima_nota else "nenhuma"
            ano, mes = ctx.ultimo_mes
            ui.label(
                f"Notas importadas até {ultima}. As ordens vão desta data até {mes:02d}/{ano}, em "
                "dias de pregão. Clique numa posição para criar uma venda dela."
            ).classes("coir-subtitulo text-sm").mark("sim-faixa")
            tabela.grade(
                COLUNAS_POSICOES,
                [(p.ticker, nomes.get(p.conta, ""), p.quantidade, p.pm) for p in posicoes],
                chaves=list(range(len(posicoes))), altura="h-[260px]", ao_clicar=de_posicao,
            ).mark("sim-posicoes")  # fmt: skip
        desenhar_ordens()

    botao_adicionar.on_click(lambda: adicionar())
    botao_copiar.on_click(copiar_datas)
    botao_simular.on_click(simular)
    hipoteses()
    desenhar_ordens()
    ui.timer(0.01, carregar, once=True)
