"""Construtor de notas sintéticas para os casos fiscais (valores conferidos à mão)."""

from __future__ import annotations

from collections.abc import Sequence
from dataclasses import dataclass, field
from datetime import date
from decimal import Decimal
from itertools import count

from coir import referencias
from coir.domain.fiscal import (
    AjusteFuturo,
    Apuracao,
    Categoria,
    Evento,
    Grupo,
    MesApurado,
    Movimento,
    NotaMotor,
    OperacaoMotor,
    PosicaoAtivo,
    ResultadoCategoria,
    TipoAjuste,
    TipoAtivo,
    TipoEvento,
    Visao,
)
from coir.domain.parametros import Parametro, ParametrosFiscais
from coir.engine.apuracao import apurar
from coir.engine.movimentos import carteira, movimentar

D = Decimal
REFS = referencias.carregar_fiscais()


def parametros(**troca: str) -> ParametrosFiscais:
    """Parâmetros de data/ com valores trocados (para testar as duas leituras de uma regra)."""
    lista = [p for p in REFS.parametros.parametros if p.nome not in troca]
    lista += [Parametro(nome=n, valor=v) for n, v in troca.items()]
    return ParametrosFiscais(parametros=tuple(lista))


@dataclass
class Livro:
    """Notas de um titular. `op(...)` cria a operação; `nota(...)` junta operações numa nota."""

    titular: int = 1
    notas: list[NotaMotor] = field(default_factory=lambda: list[NotaMotor]())
    eventos: list[Evento] = field(default_factory=lambda: list[Evento]())
    ativos: dict[str, int] = field(default_factory=lambda: dict[str, int]())
    _ids: count[int] = field(default_factory=lambda: count(1))

    def ativo(self, ticker: str) -> int:
        return self.ativos.setdefault(ticker, len(self.ativos) + 1)

    def op(
        self, cv: str, ticker: str, quantidade: str | int, preco: str | float, *,
        mercado: str = "VISTA", tipo: TipoAtivo | None = None, vencimento: date | None = None,
        opcao: str | None = None, liquidacao: date | None = None, pme_ate: date | None = None,
    ) -> OperacaoMotor:  # fmt: skip
        q, p = D(str(quantidade)), D(str(preco))
        if tipo is None:
            tipo = TipoAtivo.OPCAO if mercado in ("CALL", "PUT") else REFS.tipos.tipo(ticker)
        return OperacaoMotor(
            id=next(self._ids), ordem=0, ativo=self.ativo(ticker), ticker=ticker, tipo=tipo,
            mercado=mercado, cv=cv, quantidade=q, preco=p, total=q * p,  # type: ignore[arg-type]
            vencimento=vencimento, opcao=None if opcao is None else self.ativo(opcao),
            liquidacao_termo=liquidacao, pme_ate=pme_ate,
        )  # fmt: skip

    def nota(
        self, dia: date, *ops: OperacaoMotor, custos: str = "0", irrf: str = "0",
        irrf_dt: str = "0", corretora: str = "XP", conta: int = 1,
        ajustes: Sequence[AjusteFuturo] = (),
    ) -> NotaMotor:  # fmt: skip
        n = NotaMotor(
            id=next(self._ids), titular=self.titular, conta=conta, corretora=corretora,
            numero=str(len(self.notas) + 1), data=dia, custos=D(custos), irrf=D(irrf),
            irrf_day_trade=D(irrf_dt),
            operacoes=tuple(o.model_copy(update={"ordem": i}) for i, o in enumerate(ops)),
            ajustes=tuple(ajustes),
        )  # fmt: skip
        self.notas.append(n)
        return n

    def evento(
        self, dia: date, ticker: str, tipo: TipoEvento, cv: str | None = None,
        quantidade: str | None = None, valor: str = "0", conta: int = 1,
        destino: int | None = None, tipo_ativo: TipoAtivo = TipoAtivo.ACAO,
        corretora: str = "XP", corretora_destino: str | None = None,
    ) -> None:  # fmt: skip
        self.eventos.append(
            Evento(
                id=next(self._ids), data=dia, titular=self.titular, conta=conta,
                corretora=corretora, ativo=self.ativo(ticker), ticker=ticker, tipo=tipo,
                tipo_ativo=tipo_ativo,
                cv=cv,  # type: ignore[arg-type]
                quantidade=None if quantidade is None else D(quantidade), valor=D(valor),
                conta_destino=destino,
                corretora_destino=None if destino is None else corretora_destino or corretora,
            )
        )  # fmt: skip

    def apurar(self, ate: date, params: ParametrosFiscais | None = None) -> Apuracao:
        return apurar(self.notas, self.eventos, params or REFS.parametros, REFS.calendario, ate)

    def movimentos(self, ate: date, visao: Visao = Visao.TITULAR) -> list[Movimento]:
        return movimentar(self.notas, self.eventos, REFS.parametros, ate, visao)

    def carteira(self, ate: date, visao: Visao = Visao.TITULAR) -> list[PosicaoAtivo]:
        return carteira(self.notas, self.eventos, REFS.parametros, ate, visao)


def mes(ap: Apuracao, ano: int, m: int, titular: int = 1) -> MesApurado:
    return next(x for x in ap.meses if (x.titular, x.ano, x.mes) == (titular, ano, m))


def categoria(ap: Apuracao, ano: int, m: int, cat: Categoria) -> ResultadoCategoria:
    return next(c for c in mes(ap, ano, m).categorias if c.categoria is cat)


def ajuste(
    mercadoria: str, valor: str, *, tipo: TipoAjuste = TipoAjuste.NORMAL, cv: str | None = None,
    quantidade: str = "0", volume: str = "0", grupo: Grupo = Grupo.INDICE,
    vencimento: date | None = None,
) -> AjusteFuturo:  # fmt: skip
    return AjusteFuturo(ordem=0, mercadoria=mercadoria, grupo=grupo, tipo=tipo, valor=D(valor),
                        volume=D(volume), cv=cv, quantidade=D(quantidade),  # type: ignore[arg-type]
                        vencimento=vencimento)  # fmt: skip
