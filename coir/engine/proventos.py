"""Proventos a partir da movimentação da B3 (decisões do usuário de 19/10/2026).

- Dividendo de empresa brasileira (ação, unit): isento (Lei 9.249/1995 art. 10). De BDR, ETF ou
  FII: fora da v3.0, listado com aviso (BDR: Lei 14.754/2023, não é isento).
- JCP: tributação exclusiva; a B3 credita o líquido e o preço unitário é o bruto por ação, o que
  mostra a alíquota retida: 15% até 2025, 17,5% desde 2026, pela data do pagamento ou do crédito,
  o que vier primeiro (LC 224/2025 art. 8º). Pago num ano com a alíquota do ano anterior = creditado
  no ano anterior: declarado no ano do crédito, com o crédito a receber em 31/12 (Bens e Direitos
  99.07). Bruto = líquido / (1 - alíquota), centavos arredondados (estimativa: vale o informe).
- Rendimento de FII/Fiagro: isento (Lei 11.033/2004 art. 3º III). "Rendimento" em ação pago junto
  com um dividendo ou JCP do mesmo código e dia: a atualização do provento pela Selic entre o
  balanço e o pagamento (BB, BB Seguridade, Caixa Seguridade); não é lucro distribuído (Lei 9.249
  art. 10): tributação exclusiva na fonte como renda fixa (Lei 8.981/1995 art. 65 §4º c; Lei
  11.033/2004 art. 1º), como as fontes retêm (22,5% nos dados reais); bruto = quantidade x preço,
  IRRF = bruto - líquido; ficha exclusiva (código 06 ou 12, conforme o informe da empresa).
  "Rendimento" em ação sem esse par: fora, listado com aviso.
- Reembolso de BTC (doador): isento até 2025 (IN RFB 1.585/2015 art. 74); desde 2026 a Lei
  15.265/2025 é omissa para dividendos: isento por equivalência, com aviso (parâmetro).
- Restituição de capital: reduz o custo das ações (IN 1.585 art. 58 §8º): listada e proposta como
  evento de redução de capital; não é rendimento.
- Linhas "- Transferido": o provento acompanhando a transferência de custódia, em pares que se
  anulam; ignoradas, mas cada par tem de fechar (senão, aviso). Débito de provento fora desse
  padrão falha alto.
- Lei 15.270/2025 (desde 2026): dividendos da mesma empresa no mês acima de R$ 50 mil têm IRRF de
  10% sobre o total. O COIR só avisa, somando os valores líquidos por empresa (raiz do código) e
  mês: acima do limite, ou acima de limite x (1 - 10%) como "possível retenção".
"""

from collections import Counter, defaultdict
from collections.abc import Iterable, Mapping, Sequence
from datetime import date
from decimal import ROUND_HALF_UP, Decimal

from coir.domain.b3 import LinhaMovimentacao
from coir.domain.fiscal import ZERO, TipoAtivo
from coir.domain.parametros import ParametrosFiscais, TiposAtivo
from coir.domain.proventos import LinhaB3, Provento, TipoProvento
from coir.erros import MotorError

CENTAVO = Decimal("0.01")
DIVIDENDO, JCP, RENDIMENTO = "Dividendo", "Juros Sobre Capital Próprio", "Rendimento"
REEMBOLSO, RESTITUICAO = "Reembolso", "Restituição de Capital"
TRANSFERIDO = " - Transferido"
ISENTO_DIVIDENDO = (TipoAtivo.ACAO, TipoAtivo.UNIT)


Candidatos = Mapping[tuple[str, date], Sequence[tuple[str, Decimal]]]  # (raiz, pagamento)
ARREDONDAMENTO_PRECO = Decimal("0.0005")  # a movimentação imprime o preço com 3 casas


def _origem_reembolso(x: LinhaMovimentacao, params: ParametrosFiscais,
                      candidatos: Candidatos) -> tuple[str, int]:  # fmt: skip
    """Reembolso de BTC: substitui dividendo ou JCP? O valor é a quantidade emprestada vezes o
    valor por ação do provento do mesmo dia e empresa, menos o IR do JCP (Lei 13.043/2014 art. 7º
    §1º II; Lei 15.265/2025 art. 20). Só uma origem com quantidade inteira identifica; nenhuma ou
    mais de uma: ("", 0), com aviso (parecer do especialista e decisão do usuário de 21/10/2026)."""
    valor = x.valor or ZERO
    taxas = {t for t in (_vigente(params, "aliquota_irrf_jcp", x.data),
                         _vigente(params, "aliquota_irrf_jcp", date(x.data.year - 1, 12, 31)))
             if t is not None}  # fmt: skip
    achados: dict[str, int] = {}
    for tipo, por_acao in candidatos.get((x.codigo[:4], x.data), ()):
        fatores = [Decimal(1)] if tipo == "DIVIDENDO" else [1 - t for t in taxas]
        for f in fatores:
            if not por_acao or not f:
                continue
            n = int((valor / (por_acao * f)).to_integral_value())
            margem = Decimal("0.01") + n * ARREDONDAMENTO_PRECO * f
            if n >= 1 and abs(valor - n * por_acao * f) <= margem:
                achados[tipo] = n
    if len(achados) == 1:
        [(tipo, n)] = achados.items()
        return tipo, n
    return "", 0


def _vigente(params: ParametrosFiscais, nome: str, dia: date) -> Decimal | None:
    achados = [p for p in params.parametros if p.nome == nome and p.vigente(dia)]
    return Decimal(achados[0].valor) if len(achados) == 1 else None


def _empresa(produto: str) -> str:
    partes = produto.split(" - ", 1)
    return partes[1].strip() if len(partes) > 1 else produto


def proventos(
    linhas: Iterable[LinhaB3],
    tipos: TiposAtivo,
    params: ParametrosFiscais,
    avisos: list[str] | None = None,
    publicos: Candidatos | None = None,
) -> list[Provento]:
    """``publicos``: proventos em dinheiro do pacote da B3 por (raiz do código, pagamento) ->
    [(DIVIDENDO ou JCP, valor por ação)], para identificar a origem dos reembolsos de BTC; a
    movimentação do titular entra junto (as linhas de dividendo e JCP do mesmo dia)."""
    avisos = avisos if avisos is not None else []
    saida: list[Provento] = []
    pares: Counter[tuple[int, date, str, str, Decimal]] = Counter()
    linhas = list(linhas)
    pagos = {
        (i.titular, i.linha.codigo, i.linha.data)
        for i in linhas
        if i.linha.movimentacao in (DIVIDENDO, JCP) and i.linha.entrada
    }
    por_dia: dict[tuple[str, date], list[tuple[str, Decimal]]] = defaultdict(list)
    for (raiz, dia), lista in (publicos or {}).items():
        por_dia[(raiz, dia)] += lista
    for i in linhas:
        x = i.linha
        if x.movimentacao in (DIVIDENDO, JCP) and x.entrada and x.preco:
            origem = ("JCP" if x.movimentacao == JCP else "DIVIDENDO", x.preco)
            if origem not in por_dia[(x.codigo[:4], x.data)]:
                por_dia[(x.codigo[:4], x.data)].append(origem)
    for item in linhas:
        x = item.linha
        if x.movimentacao.endswith(TRANSFERIDO):
            base = x.movimentacao.removesuffix(TRANSFERIDO)
            pares[(item.titular, x.data, x.codigo, base, abs(x.valor or ZERO))] += (
                1 if x.entrada else -1
            )
            continue
        if x.movimentacao not in (DIVIDENDO, JCP, RENDIMENTO, REEMBOLSO, RESTITUICAO):
            continue
        if not x.entrada:
            raise MotorError(
                f"{x.data}: {x.movimentacao} de {x.codigo} a débito (mov:{item.id}) "
                "fora de uma transferência"
            )
        if x.valor is None:
            avisos.append(
                f"{x.data}: {x.movimentacao} de {x.codigo} sem valor na movimentação "
                "(confira no informe da empresa)"
            )
            continue
        saida.append(_provento(item, tipos, params, pagos, por_dia))
    for (_, dia, codigo, base, valor), saldo in sorted(pares.items()):
        if saldo:
            avisos.append(
                f"{dia}: {base} de {codigo} ({valor}) transferido sem o par de crédito e "
                "débito; confira a movimentação"
            )
    for p in saida:
        if p.observacao and p.tipo is TipoProvento.FORA_DO_ESCOPO:
            avisos.append(f"{p.data}: {p.codigo}: {p.observacao}")
    avisos.extend(_lei_15270(saida, params))
    return saida


def _criar(
    item: LinhaB3, tipo: TipoProvento, bruto: Decimal, aliquota: Decimal, ano: int, observacao: str,
    reembolso_de: str = "",
) -> Provento:  # fmt: skip
    x = item.linha
    liquido = x.valor or ZERO
    return Provento(
        data=x.data,
        titular=item.titular,
        codigo=x.codigo,
        empresa=_empresa(x.produto),
        instituicao=x.instituicao,
        tipo=tipo,
        quantidade=x.quantidade,
        liquido=liquido,
        bruto=bruto,
        irrf=bruto - liquido,
        aliquota=aliquota,
        ano_declaracao=ano,
        origem=f"mov:{item.id}",
        observacao=observacao,
        credito=item.credito if tipo is TipoProvento.JCP else None,
        chave=item.chave,
        reembolso_de=reembolso_de,
    )


def _provento(
    item: LinhaB3, tipos: TiposAtivo, params: ParametrosFiscais, pagos: set[tuple[int, str, date]],
    candidatos: Candidatos,
) -> Provento:  # fmt: skip
    x = item.linha
    if x.movimentacao == JCP:
        return _jcp(item, params)
    tipo_ativo = tipos.tipo(x.codigo)
    tipo, obs = TipoProvento.FORA_DO_ESCOPO, ""
    if x.movimentacao == DIVIDENDO:
        if tipo_ativo in ISENTO_DIVIDENDO:
            tipo = TipoProvento.DIVIDENDO
        elif tipo_ativo is TipoAtivo.BDR:
            obs = (
                "dividendo de BDR não é isento (Lei 14.754/2023); a tributação fica para uma "
                "versão futura"
            )
        else:
            obs = f"dividendo de {tipo_ativo.value}: tratamento fora da v3.0"
    elif x.movimentacao == RENDIMENTO:
        if tipo_ativo is TipoAtivo.FII:
            tipo = TipoProvento.RENDIMENTO_FII
        elif (item.titular, x.codigo, x.data) in pagos and x.preco and x.quantidade:
            bruto = (x.quantidade * x.preco).quantize(CENTAVO, rounding=ROUND_HALF_UP)
            liquido = x.valor or ZERO
            retida = (1 - liquido / bruto) if bruto else ZERO
            return _criar(
                item,
                TipoProvento.ATUALIZACAO,
                max(bruto, liquido),
                retida.quantize(Decimal("0.001")),
                x.data.year,
                "atualização do provento pela Selic: tributação exclusiva na fonte "
                "(código 06 ou 12, conforme o informe da empresa)",
            )
        else:
            obs = f"rendimento em {tipo_ativo.value} sem dividendo ou JCP no mesmo dia: fora"
    elif x.movimentacao == REEMBOLSO:
        de, acoes = _origem_reembolso(x, params, candidatos)
        if de == "JCP":
            obs = (f"reembolso de JCP de {acoes} ações emprestadas (BTC): isento para o doador "
                   "(Lei 13.043/2014 art. 7º; desde 2026, Lei 15.265/2025 art. 21 I), no ano do "
                   "recebimento e sem JCP a receber (99.07)")  # fmt: skip
        else:
            origem = (f"reembolso de dividendos de {acoes} ações emprestadas (BTC)" if de else
                      "reembolso de BTC sem identificar se de dividendos ou de JCP (isento nos "
                      "dois casos até 2025; confira o informe)")  # fmt: skip
            obs = origem
            if params.booleano("reembolso_btc_norma_omissa", x.data):
                obs = (f"{origem}: a Lei 15.265/2025 é omissa para dividendos; isento por "
                       "equivalência (confira com o contador)")  # fmt: skip
            if not params.booleano("reembolso_btc_dividendos_isento", x.data):
                obs = f"{origem}: tributável pelo parâmetro vigente"
        return _criar(item, TipoProvento.REEMBOLSO_BTC, x.valor or ZERO, ZERO, x.data.year, obs,
                      reembolso_de=de)  # fmt: skip
    else:  # restituição de capital
        tipo = TipoProvento.RESTITUICAO_CAPITAL
        obs = "reduz o custo das ações (evento de redução de capital), não é rendimento"
    return _criar(item, tipo, x.valor or ZERO, ZERO, x.data.year, obs)


def _jcp(item: LinhaB3, params: ParametrosFiscais) -> Provento:
    x = item.linha
    valor = x.valor or ZERO
    if item.credito is not None:
        return _jcp_com_credito(item, item.credito, params)
    no_pagamento = params.decimal("aliquota_irrf_jcp", x.data)
    no_ano_anterior = _vigente(params, "aliquota_irrf_jcp", date(x.data.year - 1, 12, 31))
    aliquota, ano, obs = no_pagamento, x.data.year, ""
    if x.preco and x.quantidade and no_ano_anterior is not None and no_ano_anterior != no_pagamento:
        retida = 1 - valor / (x.quantidade * x.preco)
        if abs(retida - no_ano_anterior) < abs(retida - no_pagamento):
            aliquota, ano = no_ano_anterior, x.data.year - 1
            obs = (
                f"creditado em {ano} (retenção de {no_ano_anterior * 100:.1f}%): declarado em "
                f"{ano}, com o crédito a receber em 31/12/{ano}"
            )
    bruto = (valor / (1 - aliquota)).quantize(CENTAVO, rounding=ROUND_HALF_UP)
    return _criar(item, TipoProvento.JCP, bruto, aliquota, ano, obs)


def _jcp_com_credito(item: LinhaB3, credito: date, params: ParametrosFiscais) -> Provento:
    """JCP com a data do crédito conhecida (informe da empresa ou data de aprovação no pacote da
    B3): declarado no ano do crédito (P&R IRPF 2026 pergunta 320; Lei 9.249/1995 art. 9º §2º; a
    receber em 31/12 se pago depois), com a alíquota da data do crédito ou do pagamento, o que vier
    primeiro (LC 224/2025 art. 8º). Crédito depois do pagamento falha alto."""
    x = item.linha
    if credito > x.data:
        raise MotorError(f"{x.data}: JCP de {x.codigo} (mov:{item.id}) creditado em {credito}, "
                         "depois do pagamento")  # fmt: skip
    aliquota = params.decimal("aliquota_irrf_jcp", credito)
    obs = f"creditado em {credito:%d/%m/%Y} ({item.credito_origem or 'informada'})"
    if credito.year < x.data.year:
        obs += f": declarado em {credito.year}, com o crédito a receber em 31/12/{credito.year}"
    if x.preco and x.quantidade:
        retida = 1 - (x.valor or ZERO) / (x.quantidade * x.preco)
        if abs(retida - aliquota) > Decimal("0.005"):
            obs += (f"; retenção implícita de {retida * 100:.1f}% difere da alíquota de "
                    f"{aliquota * 100:.1f}% (confira o informe)")  # fmt: skip
    bruto = ((x.valor or ZERO) / (1 - aliquota)).quantize(CENTAVO, rounding=ROUND_HALF_UP)
    return _criar(item, TipoProvento.JCP, bruto, aliquota, credito.year, obs)


def _lei_15270(saida: list[Provento], params: ParametrosFiscais) -> list[str]:
    somas: defaultdict[tuple[int, str, int, int], Decimal] = defaultdict(Decimal)
    for p in saida:
        # reembolso de dividendos de ações emprestadas soma no aviso (decisão de 21/10/2026)
        if p.tipo is TipoProvento.DIVIDENDO or p.reembolso_de == "DIVIDENDO":
            somas[(p.titular, p.codigo[:4], p.data.year, p.data.month)] += p.liquido
    avisos: list[str] = []
    for (_, raiz, ano, mes), soma in sorted(somas.items()):
        limite = _vigente(params, "dividendos_limite_mensal_mesma_pj", date(ano, mes, 1))
        aliquota = _vigente(params, "aliquota_irrf_dividendos_acima_limite", date(ano, mes, 1))
        if limite is None or aliquota is None:
            continue
        if soma > limite:
            situacao = "acima do limite"
        elif soma > limite * (1 - aliquota):
            situacao = "possível retenção (o valor da B3 é líquido)"
        else:
            continue
        avisos.append(
            f"{mes:02d}/{ano}: dividendos de {raiz} somam {soma} no mês, {situacao} "
            f"de R$ {limite} (Lei 15.270/2025: IRRF de {aliquota * 100:.0f}% sobre o "
            "total; dividendos de lucros até 2025 aprovados até 31/12/2025 ficam fora)"
        )
    return avisos
