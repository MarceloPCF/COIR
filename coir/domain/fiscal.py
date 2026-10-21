"""Entradas e saídas do motor fiscal (coir/engine). Sem I/O; tudo Decimal.

Entradas: notas (com operações e ajustes de futuros já ligados ao ativo), eventos corporativos e
parâmetros. Saídas: movimentos de posição (explicáveis, regra 12), apuração mensal por categoria,
IRRF, DARF e o que vai para a declaração anual.
"""

from datetime import date
from decimal import Decimal
from enum import StrEnum
from typing import Literal

from pydantic import BaseModel, ConfigDict

CV = Literal["C", "V"]
ZERO = Decimal(0)


class Modelo(BaseModel):
    model_config = ConfigDict(frozen=True, extra="forbid")


class TipoAtivo(StrEnum):
    ACAO = "ACAO"
    UNIT = "UNIT"
    ETF = "ETF"
    ETF_RF = "ETF_RF"  # ETF de renda fixa (Lei 13.043/2014): tributação exclusiva na fonte
    BDR = "BDR"
    FII = "FII"  # inclui Fiagro
    OPCAO = "OPCAO"
    FUTURO = "FUTURO"


class Categoria(StrEnum):
    COMUM = "COMUM"
    DAY_TRADE = "DAY_TRADE"
    FII = "FII"


class Visao(StrEnum):
    """Agrupamento das posições. A fiscal (apuração e IRPF) é sempre TITULAR: custo médio único
    por titular e ativo, somando todas as corretoras e contas (IN RFB 1.585/2015 art. 58).
    CORRETORA e CONTA são visões operacionais (tela e comparação com a planilha)."""

    TITULAR = "titular"
    CORRETORA = "corretora"
    CONTA = "conta"


class Grupo(StrEnum):
    """Subtotal da apuração (como no gabarito e no programa da Receita)."""

    A_VISTA = "a_vista"  # ações, units, ETF, BDR, FII (inclui exercício e termo vendido depois)
    OPCOES = "opcoes"  # ativo-objeto ação ou unit (parâmetro opcoes_linha_acoes_tipos_objeto)
    OPCOES_OUTROS = "opcoes_outros"  # ativo-objeto ETF, índice ou outro
    INDICE = "indice"
    DOLAR = "dolar"
    JUROS = "juros"
    COMMODITY = "commodity"
    TERMO = "termo"  # venda à vista na data da liquidação do termo (IN RFB 1.585 art. 62 I)


# ------------------------------------------------------------------------------------- entradas
class OperacaoMotor(Modelo):
    """Negócio Bovespa já ligado ao ativo (ativo_id da F3)."""

    id: int  # operação no banco (explicabilidade)
    ordem: int  # ordem na nota
    ativo: int
    ticker: str  # para exibir
    tipo: TipoAtivo
    mercado: str  # VISTA, FRACIONARIO, TERMO, CALL, PUT, EXERC CALL, EXERC PUT
    cv: CV
    quantidade: Decimal
    preco: Decimal
    total: Decimal  # valor bruto da operação (quantidade x preço)
    vencimento: date | None = None  # opção: data de vencimento da série
    opcao: int | None = None  # exercício: ativo da opção exercida
    objeto_tipo: TipoAtivo | None = None  # opção: tipo do ativo-objeto (None: não identificado)
    liquidacao_termo: date | None = None  # compra a termo: entra na carteira nesta data
    pme_ate: date | None = None  # compra elegível à isenção de PME: vendas isentas até esta data


class TipoAjuste(StrEnum):
    NORMAL = "NORMAL"  # negócio que abre ou encerra posição (valor = ajuste do dia do negócio)
    AJUPOS = "AJUPOS"  # ajuste diário da posição carregada
    DAY_TRADE = "DAY_TRADE"
    OUTRO = "OUTRO"  # taxa de permanência e afins: resultado no dia


class AjusteFuturo(Modelo):
    """Linha da nota BM&F com valor de ajuste (negócio normal, day trade ou AJUPOS)."""

    ordem: int
    mercadoria: str  # contrato com vencimento (WINM24)
    grupo: Grupo  # INDICE, DOLAR ou COMMODITY
    tipo: TipoAjuste
    valor: Decimal  # crédito positivo, débito negativo
    volume: Decimal  # base do rateio de custos, em R$ (preço x qtd x multiplicador; AJUPOS = 0)
    cv: CV | None = None  # negócio normal: compra ou venda
    quantidade: Decimal = ZERO  # contratos do negócio
    vencimento: date | None = None  # vencimento do contrato (encerra a posição que sobrar)
    entrega_fisica: bool = False  # contrato liquidado por entrega (ex.: café): fora do escopo

    @property
    def day_trade(self) -> bool:
        return self.tipo is TipoAjuste.DAY_TRADE


class NotaMotor(Modelo):
    id: int
    titular: int
    conta: int
    corretora: str
    numero: str
    data: date
    custos: Decimal  # custos da nota sem IRRF (rateados nas operações)
    irrf: Decimal  # 0,005% impresso
    irrf_day_trade: Decimal  # 1% impresso
    operacoes: tuple[OperacaoMotor, ...] = ()
    ajustes: tuple[AjusteFuturo, ...] = ()


class PosicaoFuturo(Modelo):
    """Contrato futuro em aberto: os ajustes e custos acumulados ainda não realizados."""

    titular: int
    mercadoria: str
    grupo: Grupo
    quantidade: Decimal  # comprado > 0, vendido < 0
    acumulado: Decimal
    vencimento: date | None


class OutroCusto(Modelo):
    """Custo dedutível lançado pelo usuário num mês (ex.: taxas de aluguel de ações - BTC): reduz o
    resultado da categoria no mês (como a aba BTC_Outros da planilha v2)."""

    id: int
    titular: int
    ano: int
    mes: int
    categoria: Categoria  # COMUM ou DAY_TRADE
    valor: Decimal  # positivo
    descricao: str = ""


class TipoEvento(StrEnum):
    DESDOBRAMENTO = "DESDOBRAMENTO"
    GRUPAMENTO = "GRUPAMENTO"
    BONIFICACAO = "BONIFICACAO"
    SUBSCRICAO = "SUBSCRICAO"
    REDUCAO_CAPITAL = "REDUCAO_CAPITAL"
    TRANSFERENCIA = "TRANSFERENCIA"  # custódia entre contas do mesmo titular
    INCORPORACAO = "INCORPORACAO"  # troca de A por B com o custo (ver NaturezaEvento)
    CISAO = "CISAO"  # parte do custo de A vai para B, pelo percentual da companhia
    LEILAO_FRACAO = "LEILAO_FRACAO"  # fração vendida em leilão pela companhia (alienação)


class NaturezaEvento(StrEnum):
    """Natureza da INCORPORACAO (a regra fiscal muda com ela)."""

    INCORPORACAO_SOCIEDADE = "INCORPORACAO_SOCIEDADE"  # Lei 6.404 art. 227
    INCORPORACAO_ACOES = "INCORPORACAO_ACOES"  # Lei 6.404 art. 252 (divergência RFB x Judiciário)
    FUSAO = "FUSAO"  # Lei 6.404 art. 228
    CONVERSAO = "CONVERSAO"  # de classe na mesma companhia (ex.: PN -> ON)
    AJUSTE_SERIE = "AJUSTE_SERIE"  # série de opção com código ou quantidade ajustados pela B3


TIPOS_COM_DESTINO = frozenset({TipoEvento.INCORPORACAO, TipoEvento.CISAO})


class Evento(Modelo):
    """Evento corporativo ou de custódia. Ajuste de posição: C soma quantidade e custo, V
    subtrai (como na aba Eventos_B3 do gabarito). Transferência: move a quantidade (None = toda
    a posição) da conta para a conta de destino, com o custo proporcional.

    Incorporação e cisão (sem C/V): ``quantidade`` é a quantidade de origem (a da conta) e
    ``quantidade_destino`` a de B recebida; na cisão, ``percentual_custo`` é a parte do custo de
    A que vai para B. Leilão de fração (sem C/V): ``quantidade`` vendida por ``valor`` líquido."""

    id: int
    data: date
    titular: int
    conta: int
    corretora: str
    ativo: int
    ticker: str
    tipo: TipoEvento
    tipo_ativo: TipoAtivo = TipoAtivo.ACAO
    cv: CV | None = None
    quantidade: Decimal | None = None
    valor: Decimal = ZERO
    conta_destino: int | None = None
    corretora_destino: str | None = None
    ativo_destino: int | None = None
    ticker_destino: str | None = None
    tipo_ativo_destino: TipoAtivo | None = None
    quantidade_destino: Decimal | None = None
    natureza: NaturezaEvento | None = None
    percentual_custo: Decimal | None = None  # cisão: fração do custo de A que vai para B (0..1)
    observacao: str = ""


# --------------------------------------------------------------------------------------- saídas
class Movimento(Modelo):
    """Um passo na posição de um ativo (regra 12: todo valor apurado sai daqui)."""

    data: date
    titular: int
    # conta e corretora de origem (rastreabilidade, em qualquer visão); None quando o passo junta
    # contas (day trade entre corretoras antes de 28/07/2010, pó de série comprada em duas contas)
    conta: int | None
    corretora: str | None
    ativo: int | None  # None em futuros (resultado por ajuste, sem posição)
    ticker: str
    origem: str  # "op:<id>[,<id>...]", "evento:<id>", "vencimento", "termo:<id>", "nota:<id>"
    regra: str  # explicação curta da regra aplicada
    categoria: Categoria
    grupo: Grupo
    tipo: TipoAtivo
    cv: CV | None
    quantidade: Decimal
    valor: Decimal  # custo (compra) ou alienação líquida (venda), com custos rateados
    custos: Decimal  # custos rateados da nota
    custo_baixa: Decimal = ZERO  # custo médio baixado (venda) ou receita baixada (recompra)
    resultado: Decimal = ZERO  # ganho/perda realizados neste passo
    venda_acao_a_vista: Decimal = ZERO  # valor bruto que entra no limite de isenção
    isentavel: bool = False  # resultado de ação à vista (comum): isento se vendas <= limite
    mes_limite: date | None = None  # recompra de venda a descoberto: o limite é o do mês da venda
    isento_pme: bool = False  # venda de ações de PME (Lei 13.043): fora da apuração e do limite
    base_irrf: Decimal = ZERO  # venda comum bruta na base do IRRF 0,005% (repartição por ficha)
    evento: TipoEvento | None = None  # passo de um evento corporativo: o tipo
    quantidade_apos: Decimal
    custo_apos: Decimal  # custo total da posição (negativo se vendida)

    @property
    def pm_apos(self) -> Decimal:
        return ZERO if self.quantidade_apos == 0 else self.custo_apos / self.quantidade_apos


class PosicaoAtivo(Modelo):
    """Posição aberta numa data, na visão pedida (conta/corretora None quando a visão as soma)."""

    titular: int
    corretora: str | None
    conta: int | None
    ativo: int
    ticker: str
    tipo: TipoAtivo
    vencimento: date | None  # opção: vencimento da série
    quantidade: Decimal  # negativa se vendida
    custo: Decimal  # negativo se vendida (receita)

    @property
    def pm(self) -> Decimal:
        return ZERO if self.quantidade == 0 else self.custo / self.quantidade


class ResultadoCategoria(Modelo):
    categoria: Categoria
    por_grupo: dict[Grupo, Decimal]
    resultado: Decimal  # soma dos grupos tributáveis no mês (sem o isento)
    vendas_acoes: Decimal = ZERO  # só COMUM: vendas de ações à vista no mês (limite)
    isento: Decimal = ZERO  # só COMUM: ganho de ações à vista isento
    prejuizo_anterior: Decimal = ZERO  # saldo (<= 0) de meses anteriores
    base: Decimal = ZERO  # resultado + prejuízo anterior, se positivo
    prejuizo_a_compensar: Decimal = ZERO  # saldo (<= 0) que passa ao mês seguinte
    aliquota: Decimal
    imposto: Decimal = ZERO


class Ficha(StrEnum):
    """Fichas de renda variável do programa IRPF (cada uma com imposto e IR fonte próprios)."""

    COMUNS = "comuns"  # operações comuns e day trade
    FII = "fii"  # fundos imobiliários e Fiagro


class FichaMes(Modelo):
    """Uma ficha do programa IRPF num mês. A DARF usa o saldo único de IRRF; aqui cada ficha
    recebe o IRRF que ela consumiu nesse saldo, mais o retido nela e nunca usado (a soma dos
    saldos das fichas é o saldo único, sem usar o mesmo IRRF duas vezes)."""

    ficha: Ficha
    imposto: Decimal  # imposto devido da ficha
    irrf_day_trade_mes: Decimal = ZERO  # 1%
    irrf_day_trade_anterior: Decimal = ZERO
    irrf_day_trade_a_compensar: Decimal = ZERO
    irrf_mes: Decimal = ZERO  # 0,005% (Lei 11.033/2004)
    irrf_anterior: Decimal = ZERO
    irrf_a_compensar: Decimal = ZERO
    a_pagar: Decimal = ZERO
    pago: Decimal = ZERO  # o a pagar do mês quando a DARF que o cobriu já foi emitida


class MesApurado(Modelo):
    titular: int
    ano: int
    mes: int
    categorias: tuple[ResultadoCategoria, ...]
    imposto: Decimal  # soma das categorias
    irrf_mes: Decimal
    irrf_saldo_anterior: Decimal
    irrf_deduzido: Decimal
    irrf_saldo: Decimal  # passa aos meses seguintes do ano; dezembro vai para a DIRPF
    imposto_devido: Decimal  # imposto - IRRF deduzido
    darf_acumulado_anterior: Decimal  # < R$ 10 de meses anteriores
    darf: Decimal  # a pagar neste mês (0 se abaixo do mínimo)
    darf_acumulado: Decimal  # passa ao mês seguinte
    vencimento: date | None
    isento_pme: Decimal = ZERO  # resultado das vendas isentas de PME (perdas não compensam)
    # parte de day trade (1%) do IRRF, como no programa da Receita (o resto é o de 0,005%)
    irrf_mes_day_trade: Decimal = ZERO
    irrf_saldo_day_trade_anterior: Decimal = ZERO
    irrf_saldo_day_trade: Decimal = ZERO
    fichas: tuple[FichaMes, ...] = ()  # comuns/day trade e FII, como no programa IRPF

    def ficha(self, ficha: Ficha) -> FichaMes:
        for f in self.fichas:
            if f.ficha is ficha:
                return f
        raise LookupError(f"{self.mes:02d}/{self.ano}: ficha {ficha} não apurada")


class TributacaoExclusiva(Modelo):
    """Venda de ETF de renda fixa: imposto retido pelo administrador (vai para a DIRPF)."""

    data: date
    titular: int
    ativo: int
    ticker: str
    quantidade: Decimal
    alienacao: Decimal
    custo: Decimal
    resultado: Decimal


class IrrfAnual(Modelo):
    """Saldo de IRRF não deduzido até dezembro. O de 0,005% é compensado na declaração anual
    (IN RFB 1.585/2015 art. 63 §8º III); o de day trade é definitivo e só pode ser pedido em
    restituição (art. 65 §§9º e 12 II; P&R IRPF 2026 perguntas 714 e 715)."""

    titular: int
    ano: int
    saldo: Decimal  # total
    saldo_comum: Decimal  # 0,005%: vai para a declaração anual
    saldo_day_trade: Decimal  # 1%: pedido de restituição
    comum: Decimal  # IRRF 0,005% retido no ano
    day_trade: Decimal  # IRRF 1% retido no ano
    saldo_comum_fichas: dict[Ficha, Decimal] = {}  # saldo_comum por ficha do programa IRPF


class MultaMora(Modelo):
    """Multa de mora de uma DARF paga (ou a pagar) depois do vencimento."""

    pagamento: date  # data considerada (dia não útil vai para o próximo dia útil)
    ajustado: bool  # a data informada não era dia útil
    dias: int  # dias de atraso contados
    percentual: Decimal  # já limitado ao teto
    valor: Decimal


class JurosMora(Modelo):
    """Juros de mora de uma DARF paga (ou a pagar) depois do vencimento (regra 12: guarda os meses
    e as taxas usados)."""

    pagamento: date  # data considerada (dia não útil vai para o próximo dia útil)
    percentual: Decimal  # Selic dos meses entre o vencimento e o pagamento + 1% do mês do pagamento
    valor: Decimal
    taxas: tuple[tuple[str, Decimal], ...] = ()  # ("MM/AAAA", % a.m.) somados
    meses_sem_taxa: tuple[str, ...] = ()  # depois do último mês da tabela: valor estimado


class Apuracao(Modelo):
    movimentos: tuple[Movimento, ...]
    meses: tuple[MesApurado, ...]
    exclusiva: tuple[TributacaoExclusiva, ...]
    irrf_anual: tuple[IrrfAnual, ...]
    avisos: tuple[str, ...] = ()
