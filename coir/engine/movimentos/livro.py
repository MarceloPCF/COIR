# pyright: reportPrivateUsage=false, reportUnusedFunction=false
# (os nomes com _ são internos do pacote, usados entre os arquivos dele)
"""O livro de posições (por titular, corretora ou conta), as peças que ele movimenta e as tabelas
comuns ao motor."""

from collections import defaultdict
from collections.abc import Sequence
from dataclasses import dataclass, field
from datetime import date
from decimal import Decimal

from coir.domain.fiscal import (
    ZERO,
    Categoria,
    Evento,
    Grupo,
    Movimento,
    NotaMotor,
    OperacaoMotor,
    TipoAtivo,
    TipoEvento,
    Visao,
)
from coir.domain.parametros import ParametrosFiscais
from coir.engine.posicao import Baixa, Posicao
from coir.erros import MotorError

MERCADOS_DT = frozenset({"VISTA", "FRACIONARIO", "CALL", "PUT"})


MERCADOS_ACAO_A_VISTA = frozenset({"VISTA", "FRACIONARIO", "EXERC CALL", "EXERC PUT"})


MERCADOS_EXERCICIO = frozenset({"EXERC CALL", "EXERC PUT"})


# a venda não entra na base que reparte o IRRF de 0,005% entre as fichas: o imposto desses é
# exclusivo na fonte (ETF de renda fixa) ou de alíquota zero (FI-Infra e FIP-IE)
SEM_IRRF_DA_VENDA = frozenset({TipoAtivo.ETF_RF, TipoAtivo.FUNDO_INFRA})


# cota de fundo com regra própria na venda em bolsa: o parâmetro diz se o COIR apura
REGRA_DO_FUNDO = {TipoAtivo.FIP: "fip_venda_bolsa", TipoAtivo.FIDC: "fidc_venda_bolsa"}


# titular, grupo da visão (None por titular; corretora; conta), ativo, vencimento (opção: o
# código da série se repete de um ano para outro, ex.: TAEEL382 em dez/2023 e dez/2025)
Chave = tuple[int, int | str | None, int, date | None]


Origem = tuple[int, str]  # conta, corretora


OrigemMovimento = tuple[int | None, str | None]  # None quando o passo junta contas/corretoras


@dataclass(frozen=True)
class _Op:
    nota: NotaMotor
    op: OperacaoMotor
    custos: Decimal

    @property
    def valor(self) -> Decimal:
        """Custo (compra) ou alienação líquida (venda)."""
        return self.op.total + self.custos if self.op.cv == "C" else self.op.total - self.custos


@dataclass(frozen=True)
class _Peca:
    """Parte de uma operação aplicada à posição (operação inteira ou excedente do day trade)."""

    base: _Op
    quantidade: Decimal
    valor: Decimal
    custos: Decimal
    bruto: Decimal
    regra: str


@dataclass
class _Dia:
    eventos: list[Evento] = field(default_factory=lambda: list[Evento]())
    termos: list[_Op] = field(default_factory=lambda: list[_Op]())
    operacoes: list[_Op] = field(default_factory=lambda: list[_Op]())


def _conferir_fundo(livro: "Livro", tipo: TipoAtivo | None, ticker: str, dia: date) -> None:
    """FIP comum e FIDC só entram na apuração com o parâmetro em ``ganho_liquido`` (15% na
    categoria comum, sem isenção e fora do limite de R$ 20 mil). FIP: IN RFB 1.585/2015 art. 32
    §1º I a e §3º; Lei 11.312/2006 art. 2º §1º I. FIDC: sem norma expressa para a venda em bolsa
    por pessoa física (o art. 16 I a da IN só alcança fundo fechado com carteira de ações, art. 18
    §2º): o padrão é isolar o ativo, com este motivo (parecer do especialista de 13/11/2026). Sem
    vigência do parâmetro na data, também isola: nada de palpite."""
    if tipo is None or tipo not in REGRA_DO_FUNDO:
        return
    nome = REGRA_DO_FUNDO[tipo]
    regras = {p.valor for p in livro.parametros.parametros if p.nome == nome and p.vigente(dia)}
    if regras == {"ganho_liquido"}:
        return
    if tipo is TipoAtivo.FIDC:
        raise MotorError(
            f"{ticker} é cota de FIDC: não há norma expressa para a venda de FIDC em bolsa por "
            "pessoa física, e o COIR deixa o fundo fora do cálculo desde o primeiro negócio, "
            "inclusive as compras. Confira com o administrador do "
            "fundo. A prática de mercado recolhe 15% (DARF 6015); com o parâmetro "
            "fidc_venda_bolsa = ganho_liquido o COIR apura assim"
        )
    raise MotorError(
        f"{ticker} é cota de FIP: o COIR não tem regra para a venda em {dia:%d/%m/%Y} (parâmetro "
        f"{nome}); confira com o administrador do fundo"
    )


def _avisar_come_cotas(livro: "Livro", ticker: str, dia: date, params: ParametrosFiscais) -> None:
    """Desde 2024 o FIP que não é entidade de investimento tem tributação periódica (Lei
    14.754/2023 art. 26), e o imposto pago nela aumenta o custo da cota (art. 17 §2º II). O COIR
    não tem esse dado: sem ele o ganho na venda sai maior (imposto a mais, nunca a menos). Só
    avisa, uma vez por fundo (parecer do especialista de 13/11/2026)."""
    regime = {p.valor for p in params.parametros
              if p.nome == "fip_amortizacao_custo" and p.vigente(dia)}  # fmt: skip
    aviso = (
        f"Venda de {ticker} (FIP): se o fundo não é entidade de investimento, teve tributação "
        "periódica (come-cotas) desde 2024, e o rendimento já tributado aumenta o custo das cotas "
        "(Lei 14.754/2023 arts. 26 e 17 §2º II). O COIR não tem esse dado e usou só o custo de "
        "compra: o ganho pode ter saído maior. Confira o informe do administrador"
    )
    if regime == {"proporcional_vp"} and aviso not in livro.avisos:
        livro.avisos.append(aviso)


def _avisar_conjugadas(
    livro: "Livro",
    vendas: Sequence[tuple[NotaMotor, OperacaoMotor]],
    compras: set[tuple[int, str, int, date]],
) -> None:
    """Compra à vista e venda a termo do mesmo ativo, no mesmo pregão e na mesma corretora: pode
    ser a operação conjugada com rendimento predeterminado, que é renda fixa (IN RFB 1.585/2015
    art. 47 I b e §1º I). O COIR apura como venda a termo coberta (art. 62 III) e avisa (parecer
    do especialista de 13/11/2026)."""
    for nota, op in vendas:
        if (nota.titular, nota.corretora, op.ativo, nota.data) not in compras:
            continue
        aviso = (
            f"{nota.data:%d/%m/%Y}: compra à vista e venda a termo de {op.ticker} no mesmo "
            f"pregão, na {nota.corretora}: pode ser uma operação conjugada, que a Receita trata "
            "como renda fixa (IN RFB 1.585/2015 art. 47 I b). O COIR apurou como venda a termo, "
            "com o ganho na liquidação. Se o informe da corretora mostrar essa operação como "
            "renda fixa, com imposto retido na fonte, o ganho não entra na apuração de renda "
            "variável: confira com a corretora"
        )
        if aviso not in livro.avisos:
            livro.avisos.append(aviso)


def _grupo(tipo: TipoAtivo) -> Grupo:
    return Grupo.OPCOES if tipo is TipoAtivo.OPCAO else Grupo.A_VISTA


def _grupo_opcao(op: OperacaoMotor, params: ParametrosFiscais, dia: date) -> Grupo:
    """Linha da ficha do IRPF para a opção: "opções - ações" se o ativo-objeto é de um dos tipos
    do parâmetro (ação e unit); ETF, índice e outros vão para "opções - outros". Objeto não
    identificado fica em ações, como antes da F8."""
    acoes = {t.strip() for t in params.texto("opcoes_linha_acoes_tipos_objeto", dia).split(";")}
    if op.objeto_tipo is None or op.objeto_tipo.value in acoes:
        return Grupo.OPCOES
    return Grupo.OPCOES_OUTROS


def _categoria(tipo: TipoAtivo, day_trade: bool) -> Categoria:
    if tipo is TipoAtivo.FII:
        return Categoria.FII  # FII (inclusive day trade): 20%, compensação só entre FII
    return Categoria.DAY_TRADE if day_trade else Categoria.COMUM


def _proporcao(valor: Decimal, parte: Decimal, todo: Decimal) -> Decimal:
    return valor if parte == todo else valor * parte / todo


class Livro:
    """Posições de todos os titulares numa visão: a fiscal (por titular, todas as corretoras e
    contas juntas) ou uma operacional (por corretora ou por conta). A conta e a corretora de
    origem vão em cada movimento, em qualquer visão."""

    def __init__(self, visao: Visao) -> None:
        self.visao = visao
        self.posicoes: dict[Chave, Posicao] = defaultdict(Posicao)
        # contas que formam cada posição aberta (atribui o pó à conta, se for uma só)
        self.origens: dict[Chave, set[Origem]] = defaultdict(set)
        self.corretoras: dict[int, str] = {}  # conta -> corretora
        self.tickers: dict[int, str] = {}
        self.tipos: dict[int, TipoAtivo] = {}
        self.grupos: dict[int, Grupo] = {}  # opção: linha pelo ativo-objeto
        # lote do termo liquidado no dia em processamento que ainda não foi vendido nesta data
        # (quantidade e custo pelo preço a termo), por posição
        self.termo_hoje: dict[Chave, Posicao] = {}
        self.parametros = ParametrosFiscais(parametros=())  # os do livro, ao percorrer
        self.hoje: date | None = None  # o dia em processamento
        self.fins_pme: set[date] = set()  # dias seguintes ao fim de uma isenção de PME
        self.movimentos: list[Movimento] = []
        self.avisos: list[str] = []
        # titular e ativos do item em processamento (o que isolar se o motor parar nele)
        self.em_curso: tuple[int, frozenset[int]] | None = None
        # (titular, ativo) -> dia a partir do qual o ativo ficou fora do cálculo
        self.cortes: dict[tuple[int, int], date] = {}
        # ativo puxado para fora por outro ligado (exercício, incorporação) -> o que o puxou
        self.puxados: dict[tuple[int, int], int] = {}
        # venda que passou de comprado a vendido: posição, dia, quantidade vendida, posição antes
        self.suspeitas: list[tuple[Chave, date, Decimal, Decimal]] = []

    def fora(self, titular: int, ativo: int, dia: date) -> bool:
        corte = self.cortes.get((titular, ativo))
        return corte is not None and corte <= dia

    def processando(self, titular: int, *ativos: int | None) -> None:
        self.em_curso = (titular, frozenset(a for a in ativos if a is not None))
        for ativo in self.em_curso[1]:
            tipo = self.tipos.get(ativo)
            if tipo in REGRA_DO_FUNDO and self.hoje is not None:
                _conferir_fundo(self, tipo, self.tickers.get(ativo, str(ativo)), self.hoje)
            if tipo is TipoAtivo.SEM_TIPO:
                raise MotorError(
                    f"{self.tickers.get(ativo, ativo)} não tem tipo de ativo nas tabelas do COIR, "
                    "e sem o tipo não se sabe como ele é tributado. Atualize as tabelas em "
                    "Configurações; se continuar, mande a nota pela tela Suporte"
                )

    def _grupo(self, de: Origem) -> int | str | None:
        if self.visao is Visao.CONTA:
            return de[0]
        return de[1] if self.visao is Visao.CORRETORA else None

    def chave(self, titular: int, de: Origem, ativo: int, vencimento: date | None = None) -> Chave:
        self.corretoras[de[0]] = de[1]
        return (titular, self._grupo(de), ativo, vencimento)

    def chave_opcao(self, titular: int, de: Origem, ativo: int, dia: date) -> Chave | None:
        """Posição aberta da série com o vencimento mais próximo a partir de `dia`."""
        grupo = self._grupo(de)
        abertas = [
            k for k, p in self.posicoes.items()
            if k[:3] == (titular, grupo, ativo) and p.quantidade
            and (k[3] is None or k[3] >= dia)
        ]  # fmt: skip
        return min(abertas, key=lambda k: k[3] or date.max, default=None)

    def origem_unica(self, chave: Chave) -> OrigemMovimento:
        contas = self.origens.get(chave, set())
        if len(contas) == 1:
            return next(iter(contas))
        corretoras = {c for _, c in contas}
        return (None, next(iter(corretoras)) if len(corretoras) == 1 else None)

    def registrar(
        self, chave: Chave, dia: date, origem: str, regra: str, *, de: OrigemMovimento,
        categoria: Categoria, cv: str | None, quantidade: Decimal, valor: Decimal,
        custos: Decimal = ZERO, baixa: Baixa | None = None, venda_acao: Decimal = ZERO,
        isentavel: bool = False, mes_limite: date | None = None, isento_pme: bool = False,
        toca_posicao: bool = True, base_irrf: Decimal = ZERO, grupo: Grupo | None = None,
        evento: TipoEvento | None = None,
    ) -> None:  # fmt: skip
        titular, _, ativo, _ = chave
        pos = self.posicoes[chave]
        conta, corretora = de
        if toca_posicao and conta is not None and corretora is not None:
            self.origens[chave].add((conta, corretora))
        if pos.quantidade == 0:
            self.origens.pop(chave, None)
        tipo = self.tipos[ativo]
        b = baixa or Baixa()
        self.movimentos.append(
            Movimento(
                data=dia, titular=titular, conta=conta, corretora=corretora, ativo=ativo,
                ticker=self.tickers[ativo],
                origem=origem, regra=regra, categoria=categoria,
                grupo=grupo or self.grupos.get(ativo) or _grupo(tipo), tipo=tipo,
                cv=cv,  # type: ignore[arg-type]
                quantidade=quantidade, valor=valor, custos=custos, custo_baixa=b.custo_baixa,
                resultado=b.resultado, venda_acao_a_vista=venda_acao, isentavel=isentavel,
                mes_limite=mes_limite, isento_pme=isento_pme, base_irrf=base_irrf,
                quantidade_apos=pos.quantidade, custo_apos=pos.custo, evento=evento,
            )
        )  # fmt: skip


Meses = dict[tuple[int, int], set[tuple[int, int]]]  # (titular, ativo) -> (ano, mês)


OpsDoDia = list[tuple[date, _Op]]


def _de(o: _Op) -> Origem:
    return (o.nota.conta, o.nota.corretora)


def _ordem(p: _Peca) -> tuple[date, str, int, int]:
    """Peças do dia na ordem dos negócios (número da nota e ordem impressa), não na ordem em que
    as notas foram gravadas: corrigir uma nota manual não muda o cálculo."""
    return _ordem_negocio(p.base)


def _ordem_negocio(o: _Op) -> tuple[date, str, int, int]:
    """Ordem dos negócios no dia: número da nota e ordem impressa (a nota não traz horário)."""
    numero = o.nota.numero
    return (o.nota.data, numero.zfill(20) if numero.isdigit() else numero, o.op.ordem, o.op.id)
