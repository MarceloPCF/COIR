"""Posição de um ativo pelo custo médio ponderado (IN RFB 1.585/2015 art. 58: custo médio de
aquisição; custos da compra integram o custo, os da venda reduzem a alienação).

Posição comprada: quantidade > 0 e custo > 0. Vendida (venda a descoberto, lançamento de opção):
quantidade < 0 e custo < 0 (= receita da venda); o resultado sai na recompra ou no vencimento.
A venda não muda o PM; a posição zerada não guarda resto de custo.
"""

from dataclasses import dataclass
from datetime import date
from decimal import Decimal

from coir.domain.fiscal import ZERO


@dataclass
class Baixa:
    resultado: Decimal = ZERO
    custo_baixa: Decimal = ZERO  # parte do custo (ou da receita, se vendida) que saiu


@dataclass
class Posicao:
    quantidade: Decimal = ZERO
    custo: Decimal = ZERO
    descoberto_desde: date | None = None  # venda a descoberto que abriu a posição vendida
    pme_ate: date | None = None  # lote elegível à isenção de PME (vendas isentas até a data)

    @property
    def pm(self) -> Decimal:
        return ZERO if self.quantidade == 0 else self.custo / self.quantidade

    def _parte(self, quantidade: Decimal) -> Decimal:
        """Custo proporcional de `quantidade` (em módulo) da posição atual."""
        if quantidade == abs(self.quantidade):
            return self.custo
        return self.custo * quantidade / abs(self.quantidade)

    def _sair(self, quantidade: Decimal, custo: Decimal) -> None:
        self.quantidade -= quantidade if self.quantidade > 0 else -quantidade
        self.custo = ZERO if self.quantidade == 0 else self.custo - custo

    def comprar(self, quantidade: Decimal, valor: Decimal) -> Baixa:
        """Compra de `quantidade` por `valor` (bruto + custos). Cobre vendida antes de abrir."""
        baixa = Baixa()
        if self.quantidade < 0:
            cobre = min(quantidade, -self.quantidade)
            receita = self._parte(cobre)  # negativa
            custo_cobertura = valor if cobre == quantidade else valor * cobre / quantidade
            baixa = Baixa(resultado=-receita - custo_cobertura, custo_baixa=-receita)
            self._sair(cobre, receita)
            quantidade, valor = quantidade - cobre, valor - custo_cobertura
        if quantidade:
            self.quantidade += quantidade
            self.custo += valor
        return baixa

    def vender(self, quantidade: Decimal, valor: Decimal) -> Baixa:
        """Venda de `quantidade` por `valor` (bruto - custos). Zera comprada antes de vender a
        descoberto."""
        baixa = Baixa()
        if self.quantidade > 0:
            fecha = min(quantidade, self.quantidade)
            custo = self._parte(fecha)
            alienacao = valor if fecha == quantidade else valor * fecha / quantidade
            baixa = Baixa(resultado=alienacao - custo, custo_baixa=custo)
            self._sair(fecha, custo)
            quantidade, valor = quantidade - fecha, valor - alienacao
        if quantidade:
            self.quantidade -= quantidade
            self.custo -= valor
        return baixa

    def vender_lote(self, quantidade: Decimal, valor: Decimal, custo: Decimal) -> Baixa:
        """Venda de um lote com custo próprio de dentro da posição comprada (ação do termo vendida
        na data da liquidação, pelo preço a termo: IN 1.585 art. 62 I); o resto da posição fica com
        o custo que sobra."""
        self._sair(quantidade, custo)
        return Baixa(resultado=valor - custo, custo_baixa=custo)

    def baixar_sem_resultado(self, quantidade: Decimal) -> Decimal:
        """Tira `quantidade` (em módulo) da posição e devolve o custo correspondente (com sinal):
        opção exercida (o prêmio vai para o ativo-objeto) e transferência entre contas."""
        custo = self._parte(quantidade)
        self._sair(quantidade, custo)
        return custo

    def ajustar(self, quantidade: Decimal, custo: Decimal) -> None:
        """Evento corporativo: soma quantidade e custo (com sinal), sem resultado."""
        self.quantidade += quantidade
        self.custo = ZERO if self.quantidade == 0 else self.custo + custo

    def virar_po(self) -> Baixa:
        """Vencimento sem exercício: titular perde o custo; lançador ganha a receita."""
        baixa = Baixa(resultado=-self.custo, custo_baixa=self.custo)
        self.quantidade, self.custo = ZERO, ZERO
        return baixa
