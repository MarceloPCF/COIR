"""Declaração anual (F8): o que o COIR preenche no programa IRPF de um titular e ano.

`fichas()` monta as fichas (Bens e Direitos, rendimentos isentos e exclusivos, imposto pago, renda
variável mês a mês e FII) a partir de `DadosRelatorio` e do mapeamento do exercício
(`data/irpf/<exercício>.yaml`); `declaracao()` as transforma em `Relatorio`. Nenhum cálculo fiscal
novo: os valores vêm do motor (apuração, carteira fiscal, IRRF por ficha) e dos proventos; aqui só
se agrupa por linha, código e fonte pagadora (regra 12).

Decisões (docs/regras-fiscais.md, "Declaração anual (F8)"): um bem por ativo, custo médio do
titular (P&R IRPF 2026 pergunta 720), discriminação com a custódia por corretora; prejuízo de
janeiro recalculado pelo COIR, com aviso.
"""

from collections import defaultdict
from collections.abc import Iterable
from decimal import ROUND_HALF_UP, Decimal

from coir.domain.cpf import formatar_cnpj
from coir.domain.fiscal import (
    ZERO,
    Categoria,
    Ficha,
    FichaMes,
    MesApurado,
    PosicaoAtivo,
    ResultadoCategoria,
    TipoAtivo,
    TipoEvento,
)
from coir.domain.irpf import (
    BemIrpf,
    DividaIrpf,
    FichasIrpf,
    LinhaFicha,
    MapeamentoIrpf,
    MesFichaIrpf,
    RendimentoIrpf,
    TipoBem,
)
from coir.domain.proventos import TipoProvento
from coir.domain.referencias import especie_opcao
from coir.domain.relatorio import DadosRelatorio, InfoAtivo
from coir.erros import ExportacaoError
from coir.reports.modelo import Cabecalho, Coluna, Relatorio, Secao, TipoColuna, Valor, formatar

T = TipoColuna
MESES = ("Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro",
         "Outubro", "Novembro", "Dezembro")  # fmt: skip

# rótulos impressos pelo programa IRPF depois das linhas de mercado (ficha de comuns/day trade)
RESULTADO = "RESULTADO LÍQUIDO DO MÊS"
NEGATIVO_ANTERIOR = "Resultado negativo até o mês anterior"
BASE = "BASE DE CÁLCULO DO IMPOSTO"
PREJUIZO = "Prejuízo a compensar"
ALIQUOTA = "Alíquota do imposto"
IMPOSTO = "IMPOSTO DEVIDO"
TOTAL_DEVIDO = "Total do imposto devido"
DT_MES = "IR fonte de day-trade do mês"
DT_ANTERIORES = "IR fonte de day-trade dos meses anteriores"
DT_COMPENSAR = "IR fonte de day-trade a compensar"
A_PAGAR = "Imposto a pagar"
PAGO = "Imposto pago"
LEI_MES = "IR fonte (Lei nº 11.033/2004) no mês"
LEI_ANTERIORES = "IR fonte (Lei nº 11.033/2004) nos meses anteriores"
LEI_COMPENSAR = "IR fonte (Lei nº 11.033/2004) a compensar"
# ficha de FII/Fiagro (leiaute DIRPF 2023 reg. 42: sem linha de IR fonte de day trade)
FII_RETIDO_MES = "Imposto retido no mês (Lei nº 11.033/2004)"
FII_RETIDO_ANTERIORES = "Imposto retido nos meses anteriores (Lei nº 11.033/2004)"
FII_RETIDO_COMPENSAR = "Imposto retido a compensar (Lei nº 11.033/2004)"
PERCENTUAIS = frozenset({ALIQUOTA})

TIPO_BEM: dict[TipoAtivo, TipoBem] = {
    TipoAtivo.ACAO: "ACAO", TipoAtivo.UNIT: "UNIT", TipoAtivo.BDR: "BDR",
    TipoAtivo.OPCAO: "OPCAO", TipoAtivo.ETF: "ETF", TipoAtivo.ETF_RF: "ETF_RF",
    TipoAtivo.FII: "FII",
}  # fmt: skip
QUANTIDADE_DE = {"UNIT": "units", "BDR": "BDRs", "ETF": "cotas", "ETF_RF": "cotas",
                 "FII": "cotas", "FIAGRO": "cotas"}  # fmt: skip
ART_58 = "(IN RFB 1.585/2015 art. 58)"


def _centavos(valor: Decimal) -> Decimal:
    """Situação do bem em reais e centavos (o programa só aceita 2 casas)."""
    return valor.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


def _q(valor: Decimal) -> str:
    return formatar(valor, T.QUANTIDADE)


def _r(valor: Decimal) -> str:
    return formatar(valor, T.DINHEIRO)


def _meses(d: DadosRelatorio) -> list[MesApurado]:
    return sorted((m for m in d.apuracao.meses if m.titular == d.titular and m.ano == d.ano),
                  key=lambda m: m.mes)  # fmt: skip


def _cat(m: MesApurado, categoria: Categoria) -> ResultadoCategoria:
    try:
        return next(c for c in m.categorias if c.categoria is categoria)
    except StopIteration as erro:
        mes = f"{m.mes:02d}/{m.ano}"
        raise ExportacaoError(f"{mes}: apuração sem a categoria {categoria}") from erro


# --------------------------------------------------------------------------- bens e direitos
def _tipo_bem(p: PosicaoAtivo, info: InfoAtivo | None) -> TipoBem | None:
    tipo = TIPO_BEM.get(p.tipo)
    if tipo == "FII" and info is not None and info.tipo_b3 == "FIAGRO":
        return "FIAGRO"
    return tipo


def _custodia(posicoes: Iterable[PosicaoAtivo], ativo: int) -> str:
    por: dict[str, Decimal] = defaultdict(lambda: ZERO)
    for p in posicoes:
        if p.ativo == ativo and p.quantidade > 0:
            por[p.corretora or "?"] += p.quantidade
    return "; ".join(f"{c}: {_q(q)}" for c, q in sorted(por.items()))


def _discriminacao(tipo: TipoBem, p: PosicaoAtivo, info: InfoAtivo | None, custodia: str,
                   quando: str) -> str:  # fmt: skip
    if tipo == "OPCAO":
        objeto = f", ativo-objeto {info.objeto}" if info is not None and info.objeto else ""
        venc = f", vencimento {p.vencimento:%d/%m/%Y}" if p.vencimento else ""
        corretora = custodia.split(":")[0] if custodia and ";" not in custodia else custodia
        onde = f", corretora {corretora}" if corretora else ""
        return (f"{_q(p.quantidade)} opções de {especie_opcao(p.ticker)} {p.ticker}{objeto}"
                f"{venc}{onde}{quando}.")  # fmt: skip
    classe = (info.especificacao.split() or [""])[0] if info is not None else ""
    especie = QUANTIDADE_DE.get(tipo, f"ações {classe}".strip() if classe in (
        "ON", "PN", "PNA", "PNB", "PNC") else "ações")  # fmt: skip
    nome = info.razao_social if info is not None and info.razao_social else p.ticker
    onde = f", custodiadas em {custodia}" if custodia else ""
    return (f"{_q(p.quantidade)} {especie} de {nome}, código {p.ticker}{onde}{quando}. "
            f"Custo médio R$ {formatar(p.pm, T.PRECO)} {ART_58}.")  # fmt: skip


def _bens(d: DadosRelatorio, mapa: MapeamentoIrpf, avisos: list[str]) -> list[BemIrpf]:
    def abertas(posicoes: Iterable[PosicaoAtivo]) -> dict[int, PosicaoAtivo]:
        return {p.ativo: p for p in posicoes
                if p.titular == d.titular and p.quantidade > 0 and p.tipo in TIPO_BEM}  # fmt: skip

    atual, anterior = abertas(d.carteira_fiscal), abertas(d.carteira_fiscal_anterior)
    bens: list[BemIrpf] = []
    for ativo in sorted(atual.keys() | anterior.keys(),
                        key=lambda a: (atual.get(a) or anterior[a]).ticker):  # fmt: skip
        p = atual.get(ativo) or anterior[ativo]
        info = d.info_ativos.get(p.ticker)
        tipo = _tipo_bem(p, info)
        if tipo is None:
            continue
        item = mapa.bens.item(tipo)
        if ativo in atual:
            custodia, quando = _custodia(d.custodia, ativo), ""
        else:
            custodia = _custodia(d.custodia_anterior, ativo)
            quando = f" em 31/12/{d.ano - 1} (alienadas no ano)"
        cnpj = info.cnpj if info is not None and item.cnpj != "nenhum" else None
        if item.cnpj != "nenhum" and cnpj is None:
            avisos.append(f"Bens e Direitos: CNPJ de {p.ticker} não encontrado no pacote da B3; "
                          "inclua o emissor em data/emissores.csv")  # fmt: skip
        bens.append(BemIrpf(
            tipo=tipo, grupo=item.grupo, codigo=item.codigo, ticker=p.ticker, cnpj=cnpj,
            discriminacao=_discriminacao(tipo, p, info, custodia, quando),
            anterior=_centavos(anterior[ativo].custo) if ativo in anterior else ZERO,
            atual=_centavos(atual[ativo].custo) if ativo in atual else ZERO,
        ))  # fmt: skip
    bens += _jcp_a_receber(d, mapa)
    return bens


def _dividas(d: DadosRelatorio, mapa: MapeamentoIrpf) -> list[DividaIrpf]:
    """Posições vendidas em 31/12 (opção lançada, venda a descoberto) pelo valor recebido, em
    Dívidas e ônus reais (decisão do usuário de 20/10/2026; sem norma específica)."""
    if mapa.dividas is None:
        return []
    item = next(i for i in mapa.dividas.itens if i.tipo == "POSICAO_VENDIDA")

    def vendidas(posicoes: Iterable[PosicaoAtivo]) -> dict[int, PosicaoAtivo]:
        return {p.ativo: p for p in posicoes if p.titular == d.titular and p.quantidade < 0}

    atual, anterior = vendidas(d.carteira_fiscal), vendidas(d.carteira_fiscal_anterior)
    saida: list[DividaIrpf] = []
    for ativo in sorted(atual.keys() | anterior.keys(),
                        key=lambda a: (atual.get(a) or anterior[a]).ticker):  # fmt: skip
        p = atual.get(ativo) or anterior[ativo]
        posicoes = d.custodia if ativo in atual else d.custodia_anterior
        corretoras = sorted({x.corretora or "?" for x in posicoes
                             if x.ativo == ativo and x.quantidade < 0})  # fmt: skip
        if p.tipo is TipoAtivo.OPCAO:
            especie = especie_opcao(p.ticker)
            venc = f", vencimento {p.vencimento:%d/%m/%Y}" if p.vencimento else ""
            texto = (f"Opção de {especie} {p.ticker} lançada ({_q(-p.quantidade)}){venc}: prêmio "
                     f"recebido, posição em aberto")  # fmt: skip
        else:
            texto = f"Venda a descoberto de {_q(-p.quantidade)} {p.ticker}: valor recebido"
        saida.append(DividaIrpf(
            codigo=item.codigo, ticker=p.ticker,
            discriminacao=f"{texto}, corretora {', '.join(corretoras)}.",
            anterior=_centavos(-anterior[ativo].custo) if ativo in anterior else ZERO,
            atual=_centavos(-atual[ativo].custo) if ativo in atual else ZERO,
        ))  # fmt: skip
    return saida


def _jcp_a_receber(d: DadosRelatorio, mapa: MapeamentoIrpf) -> list[BemIrpf]:
    """JCP creditado num ano e pago no seguinte (decisão do usuário de 19/10/2026): a receber em
    31/12 do ano do crédito."""
    item = mapa.bens.item("JCP_A_RECEBER")
    por: dict[str, list[Decimal]] = defaultdict(lambda: [ZERO, ZERO])
    nomes: dict[str, tuple[str, str | None]] = {}
    for p in d.proventos:
        if p.tipo is not TipoProvento.JCP or not p.a_receber_em_31_12:
            continue
        if p.ano_declaracao == d.ano - 1 and p.data.year == d.ano:
            coluna = 0
        elif p.ano_declaracao == d.ano:
            coluna = 1
        else:
            continue
        info = d.info_ativos.get(p.codigo)
        chave = info.cnpj if info is not None and info.cnpj else p.empresa
        nomes[chave] = (info.razao_social if info is not None else p.empresa,
                        info.cnpj if info is not None else None)  # fmt: skip
        por[chave][coluna] += p.liquido
    texto = "Juros sobre capital próprio creditados por {} e a receber, valor líquido."
    return [
        BemIrpf(tipo="JCP_A_RECEBER", grupo=item.grupo, codigo=item.codigo, ticker=None,
                cnpj=nomes[k][1], anterior=v[0], atual=v[1],
                discriminacao=texto.format(nomes[k][0]))
        for k, v in sorted(por.items())
    ]  # fmt: skip


# -------------------------------------------------------------------------------- rendimentos
def _por_fonte(
    ficha: str, chave: str, codigo: str, itens: Iterable[tuple[str, Decimal]],
    d: DadosRelatorio, avisos: list[str],
) -> list[RendimentoIrpf]:  # fmt: skip
    """(ticker, valor) -> uma linha por CNPJ (ITUB3 e ITUB4 na mesma linha)."""
    soma: dict[str, Decimal] = defaultdict(lambda: ZERO)
    nomes: dict[str, tuple[str, str | None]] = {}
    for ticker, valor in itens:
        info = d.info_ativos.get(ticker)
        cnpj = info.cnpj if info is not None else None
        if cnpj is None:
            avisos.append(f"Rendimento {codigo} de {ticker}: CNPJ da fonte pagadora não "
                          "encontrado; inclua o emissor em data/emissores.csv")  # fmt: skip
        k = cnpj or ticker
        nomes[k] = (info.razao_social if info is not None and info.razao_social else ticker, cnpj)
        soma[k] += valor
    return [RendimentoIrpf(ficha=ficha, chave=chave, codigo=codigo, cnpj=nomes[k][1],  # type: ignore[arg-type]
                           fonte=nomes[k][0], valor=v)
            for k, v in sorted(soma.items(), key=lambda x: nomes[x[0]][0]) if v]  # fmt: skip


def _rendimentos(d: DadosRelatorio, mapa: MapeamentoIrpf, avisos: list[str]
                 ) -> list[RendimentoIrpf]:  # fmt: skip
    do_ano = [p for p in d.proventos if p.ano_declaracao == d.ano]

    def proventos(tipo: TipoProvento) -> list[tuple[str, Decimal]]:
        return [(p.codigo, p.liquido) for p in do_ano if p.tipo is tipo]

    saida: list[RendimentoIrpf] = []
    isentos, exclusivos = mapa.isentos, mapa.exclusivos

    def linha(ficha: str, chave: str, itens: Iterable[tuple[str, Decimal]]) -> None:
        secao = isentos if ficha == "isentos" else exclusivos
        saida.extend(_por_fonte(ficha, chave, secao.linha(chave).codigo, itens, d, avisos))

    linha("isentos", "dividendos", proventos(TipoProvento.DIVIDENDO))
    bonificacoes = [(m.ticker, m.valor) for m in d.apuracao.movimentos
                    if m.titular == d.titular and m.data.year == d.ano and m.valor > 0
                    and m.tipo is not TipoAtivo.OPCAO
                    and m.evento is TipoEvento.BONIFICACAO]  # fmt: skip
    linha("isentos", "bonificacao", bonificacoes)
    isento = sum((_cat(m, Categoria.COMUM).isento for m in _meses(d)), ZERO)
    if isento:
        lucro = isentos.linha("lucro_acoes_ate_limite")
        saida.append(RendimentoIrpf(ficha="isentos", chave=lucro.chave, codigo=lucro.codigo,
                                    cnpj=None, fonte=lucro.descricao, valor=isento))  # fmt: skip
    pme = sum((m.isento_pme for m in _meses(d)), ZERO)
    if pme:
        avisos.append(
            f"Ganho isento de PME (Lei 13.043/2014 art. 16) de R$ {_r(pme)} sem "
            f"linha no mapeamento do exercício {mapa.exercicio}: declare em Outros"
        )
    linha("isentos", "rendimento_fii", proventos(TipoProvento.RENDIMENTO_FII))
    linha("isentos", "reembolso_btc", proventos(TipoProvento.REEMBOLSO_BTC))
    linha("exclusivos", "jcp", proventos(TipoProvento.JCP))
    etf_rf = [(e.ticker, e.resultado) for e in d.apuracao.exclusiva
              if e.titular == d.titular and e.data.year == d.ano and e.resultado > 0]  # fmt: skip
    if etf_rf:
        avisos.append("ETF de renda fixa: o COIR estima o ganho bruto; declare o valor líquido do "
                      "informe de rendimentos do administrador")  # fmt: skip
    linha("exclusivos", "etf_rf", etf_rf)
    linha("exclusivos", "atualizacao_proventos", proventos(TipoProvento.ATUALIZACAO))
    anual = next((a for a in d.apuracao.irrf_anual
                  if a.titular == d.titular and a.ano == d.ano), None)  # fmt: skip
    if anual is not None and anual.saldo_comum:
        lei = mapa.imposto_pago.linha("ir_fonte_lei_11033")
        saida.append(RendimentoIrpf(
            ficha="imposto_pago", chave=lei.chave, codigo=lei.codigo, cnpj=None,
            fonte=lei.descricao, valor=anual.saldo_comum,
        ))  # fmt: skip
    if anual is not None and anual.saldo_day_trade:
        avisos.append(f"Saldo de IRRF de day trade em dezembro de R$ {_r(anual.saldo_day_trade)}: "
                      "não entra na declaração, só por pedido de restituição (PER/DCOMP; IN RFB "
                      "1.585/2015 art. 65 §§9º e 12 II)")  # fmt: skip
    return saida


# ------------------------------------------------------------------------- renda variável
def _mes_comuns(m: MesApurado, mapa: MapeamentoIrpf) -> MesFichaIrpf:
    c, dt = _cat(m, Categoria.COMUM), _cat(m, Categoria.DAY_TRADE)
    f: FichaMes = m.ficha(Ficha.COMUNS)

    def par(rotulo: str, a: Decimal, b: Decimal) -> LinhaFicha:
        return LinhaFicha(rotulo=rotulo, comuns=a, day_trade=b)

    def um(rotulo: str, a: Decimal) -> LinhaFicha:
        return LinhaFicha(rotulo=rotulo, comuns=a)

    linhas = [par(x.rotulo, sum((c.por_grupo.get(g, ZERO) for g in x.grupos), ZERO),
                  sum((dt.por_grupo.get(g, ZERO) for g in x.grupos), ZERO))
              for x in mapa.renda_variavel.linhas]  # fmt: skip
    linhas += [
        par(RESULTADO, c.resultado, dt.resultado),
        par(NEGATIVO_ANTERIOR, -c.prejuizo_anterior, -dt.prejuizo_anterior),
        par(BASE, c.base, dt.base),
        par(PREJUIZO, -c.prejuizo_a_compensar, -dt.prejuizo_a_compensar),
        par(ALIQUOTA, c.aliquota, dt.aliquota),
        par(IMPOSTO, c.imposto, dt.imposto),
        um(TOTAL_DEVIDO, f.imposto), um(DT_MES, f.irrf_day_trade_mes),
        um(DT_ANTERIORES, f.irrf_day_trade_anterior),
        um(DT_COMPENSAR, f.irrf_day_trade_a_compensar),
        um(A_PAGAR, f.a_pagar), um(PAGO, f.pago),
        um(LEI_MES, f.irrf_mes), um(LEI_ANTERIORES, f.irrf_anterior),
        um(LEI_COMPENSAR, f.irrf_a_compensar),
    ]  # fmt: skip
    mercados = {x.rotulo for x in mapa.renda_variavel.linhas}
    return MesFichaIrpf(mes=m.mes, linhas=_digitadas(linhas, m.mes, mercados | DIGITADAS))


def _mes_fii(m: MesApurado) -> MesFichaIrpf:
    r, f = _cat(m, Categoria.FII), m.ficha(Ficha.FII)
    valores = (
        (RESULTADO, r.resultado), (NEGATIVO_ANTERIOR, -r.prejuizo_anterior), (BASE, r.base),
        (PREJUIZO, -r.prejuizo_a_compensar), (ALIQUOTA, r.aliquota), (IMPOSTO, r.imposto),
        (FII_RETIDO_MES, f.irrf_mes), (FII_RETIDO_ANTERIORES, f.irrf_anterior),
        (FII_RETIDO_COMPENSAR, f.irrf_a_compensar), (A_PAGAR, f.a_pagar), (PAGO, f.pago),
    )  # fmt: skip
    linhas = [LinhaFicha(rotulo=r_, comuns=v) for r_, v in valores]
    return MesFichaIrpf(mes=m.mes, linhas=_digitadas(linhas, m.mes, DIGITADAS_FII))


# linhas que o usuário digita no programa (as outras o programa calcula: servem para conferir);
# em janeiro também o resultado negativo do ano anterior
DIGITADAS = frozenset({DT_MES, LEI_MES, PAGO})
DIGITADAS_FII = frozenset({FII_RETIDO_MES, PAGO})


def _digitadas(linhas: list[LinhaFicha], mes: int, rotulos: frozenset[str] | set[str]
               ) -> tuple[LinhaFicha, ...]:  # fmt: skip
    return tuple(x.model_copy(update={"digita": x.rotulo in rotulos or (
        mes == 1 and x.rotulo == NEGATIVO_ANTERIOR)}) for x in linhas)  # fmt: skip


def _fii_ativa(meses: list[MesApurado]) -> bool:
    return any(_cat(m, Categoria.FII).resultado or _cat(m, Categoria.FII).prejuizo_anterior
               or m.ficha(Ficha.FII).irrf_mes for m in meses)  # fmt: skip


def fichas(d: DadosRelatorio) -> FichasIrpf | None:
    """As fichas do ano; None se o exercício não tem mapeamento em data/irpf/."""
    mapa = d.irpf
    if mapa is None:
        return None
    avisos: list[str] = list(d.avisos_irpf)
    if mapa.provisorio:
        avisos.append(f"Mapeamento do exercício {mapa.exercicio} provisório: o programa IRPF "
                      f"{mapa.exercicio} ainda não foi publicado; confira os códigos")  # fmt: skip
    if (d.ate.month, d.ate.day) != (12, 31):
        avisos.append(f"Ano em curso: valores até {d.ate:%d/%m/%Y}")
    meses = _meses(d)
    bens = _bens(d, mapa, avisos)
    rendimentos = _rendimentos(d, mapa, avisos)
    avisos.append("Resultado negativo até o mês anterior (janeiro): o prejuízo recalculado pelo "
                  "COIR; se a declaração do ano anterior trouxe outro valor, confira e, se for o "
                  "caso, retifique (decisão do usuário de 20/10/2026)")  # fmt: skip
    avisos.append("Imposto pago: o a pagar do mês quando a DARF que o cobriu já foi emitida; no "
                  "programa, informe o valor efetivamente pago no DARF 6015")  # fmt: skip
    return FichasIrpf(
        exercicio=mapa.exercicio, ano=d.ano, provisorio=mapa.provisorio, bens=tuple(bens),
        dividas=tuple(_dividas(d, mapa)),
        rendimentos=tuple(rendimentos),
        renda_variavel=tuple(_mes_comuns(m, mapa) for m in meses),
        fii=tuple(_mes_fii(m) for m in meses) if _fii_ativa(meses) else (),
        avisos=tuple(dict.fromkeys(avisos)),
    )  # fmt: skip


# ------------------------------------------------------------------------------- relatório
def aliquota(valor: Decimal) -> str:
    """0,15 -> "15%" (como o programa IRPF mostra a alíquota)."""
    return formatar(valor, T.PERCENTUAL).replace(",00", "")


def _cab(d: DadosRelatorio) -> Cabecalho:
    return Cabecalho(d.nome, d.cpf, f"Exercício {d.ano + 1}, ano-calendário {d.ano}",
                     d.gerado_em)  # fmt: skip


def _cnpj(cnpj: str | None) -> str:
    return formatar_cnpj(cnpj) if cnpj else ""


def _celula(linha: LinhaFicha, coluna: Decimal | None) -> Valor:
    if coluna is None:
        return None
    if linha.rotulo in PERCENTUAIS:
        return aliquota(coluna)
    return coluna


def _secoes_mes(titulo: str, meses: Iterable[MesFichaIrpf], duas: bool) -> list[Secao]:
    colunas = (Coluna("Linha"), Coluna("Operações comuns" if duas else "Valor", T.DINHEIRO),
               *((Coluna("Day trade", T.DINHEIRO),) if duas else ()))  # fmt: skip
    return [
        Secao(f"{titulo} - {MESES[m.mes - 1]}", colunas, tuple(
            (x.rotulo, _celula(x, x.comuns), *((_celula(x, x.day_trade),) if duas else ()))
            for x in m.linhas))
        for m in meses
    ]  # fmt: skip


def declaracao(d: DadosRelatorio) -> Relatorio:
    f = fichas(d)
    if f is None:
        return Relatorio("declaracao-irpf", "Declaração anual (IRPF)", _cab(d), (), pdf=True,
                         notas=(f"Sem mapeamento das fichas do programa IRPF {d.ano + 1}: falta "
                                f"data/irpf/{d.ano + 1}.yaml.",))  # fmt: skip
    colunas = (Coluna("Grupo"), Coluna("Código"), Coluna("Código de negociação"),
               Coluna("CNPJ"), Coluna("Discriminação"),
               Coluna(f"Situação em 31/12/{d.ano - 1}", T.DINHEIRO),
               Coluna(f"Situação em 31/12/{d.ano}", T.DINHEIRO))  # fmt: skip
    linhas: list[tuple[Valor, ...]] = [
        (b.grupo, b.codigo, b.ticker or "", _cnpj(b.cnpj), b.discriminacao, b.anterior, b.atual)
        for b in f.bens
    ]
    total = ("Total", None, None, None, None, sum((b.anterior for b in f.bens), ZERO),
             sum((b.atual for b in f.bens), ZERO))  # fmt: skip
    secoes = [Secao("Bens e Direitos", colunas, tuple(linhas), total if linhas else None,
                    nota="Um item por ativo, custo médio do titular somando as corretoras "
                    "(P&R IRPF 2026 pergunta 720). Posição vendida vai para Dívidas.")]  # fmt: skip
    colunas = (Coluna("Código"), Coluna("Código de negociação"), Coluna("Discriminação"),
               Coluna(f"Situação em 31/12/{d.ano - 1}", T.DINHEIRO),
               Coluna(f"Situação em 31/12/{d.ano}", T.DINHEIRO))  # fmt: skip
    linhas = [(x.codigo, x.ticker, x.discriminacao, x.anterior, x.atual) for x in f.dividas]
    total = ("Total", None, None, sum((x.anterior for x in f.dividas), ZERO),
             sum((x.atual for x in f.dividas), ZERO))  # fmt: skip
    secoes.append(Secao("Dívidas e ônus reais", colunas, tuple(linhas), total if linhas else None,
                        nota="Posição vendida em aberto (opção lançada, venda a descoberto) pelo "
                        "valor recebido; sem norma específica (decisão do usuário de "
                        "20/10/2026)."))  # fmt: skip
    colunas = (Coluna("Código"), Coluna("CNPJ da fonte pagadora"),
               Coluna("Fonte pagadora / descrição"), Coluna("Valor", T.DINHEIRO))  # fmt: skip
    for ficha, titulo in (("isentos", "Rendimentos isentos e não tributáveis"),
                          ("exclusivos", "Rendimentos sujeitos à tributação exclusiva/definitiva"),
                          ("imposto_pago", "Imposto pago/retido")):  # fmt: skip
        rs = [r for r in f.rendimentos if r.ficha == ficha]
        linhas = [(r.codigo, _cnpj(r.cnpj), r.fonte, r.valor) for r in rs]
        total = ("Total", None, None, sum((r.valor for r in rs), ZERO))
        secoes.append(Secao(titulo, colunas, tuple(linhas), total if linhas else None))
    secoes += _secoes_mes("Renda variável - operações comuns/day trade", f.renda_variavel, True)
    secoes += _secoes_mes("Fundos imobiliários e Fiagro", f.fii, False)
    return Relatorio("declaracao-irpf", "Declaração anual (IRPF)", _cab(d), tuple(secoes),
                     pdf=True, notas=f.avisos)  # fmt: skip
