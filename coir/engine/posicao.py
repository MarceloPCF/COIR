"""Posição de um ativo pelo custo médio ponderado (IN RFB 1.585/2015 art. 58: custo médio de
aquisição; custos da compra integram o custo, os da venda reduzem a alienação).

Posição comprada: quantidade > 0 e custo > 0. Vendida (venda a descoberto, lançamento de opção):
quantidade < 0 e custo < 0 (= receita da venda); o resultado sai na recompra ou no vencimento.
A venda não muda o PM; a posição zerada não guarda resto de custo.

Ações de PME (IN 1.585 art. 69-A §4º): a posição guarda à parte a quantidade e o custo do lote
isento; o lote tributado é a diferença. Cada lote tem o seu PM até o fim da isenção.
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
    pme_ate: date | None = None  # lote isento de PME: vendas isentas até a data
    pme_media: Decimal | None = None  # média dos 30 pregões do fim da isenção (Lei 13.043 art. 17)
    isento_quantidade: Decimal = ZERO  # parte da posição no lote isento de PME
    isento_custo: Decimal = ZERO

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

    # ------------------------------------------------------------------ lote isento de PME
    def isento_em(self, dia: date) -> bool:
        return self.isento_quantidade > 0 and self.pme_ate is not None and dia <= self.pme_ate

    def somar_isento(self, quantidade: Decimal, custo: Decimal, ate: date,
                     media: Decimal | None) -> None:  # fmt: skip
        """Marca como isenta uma parte que já entrou na posição."""
        self.isento_quantidade += quantidade
        self.isento_custo += custo
        self.pme_ate, self.pme_media = ate, self.pme_media or media

    def _limpar_isento(self) -> None:
        if self.isento_quantidade == 0:
            self.isento_custo, self.pme_ate, self.pme_media = ZERO, None, None

    def tirar_isento(self, quantidade: Decimal) -> Decimal:
        """Tira `quantidade` do lote isento e devolve o custo dela (a posição não muda aqui)."""
        if quantidade == self.isento_quantidade:
            custo = self.isento_custo
        else:
            custo = self.isento_custo * quantidade / self.isento_quantidade
        self.isento_quantidade -= quantidade
        self.isento_custo -= custo
        self._limpar_isento()
        return custo

    def vender_isento(self, quantidade: Decimal, valor: Decimal) -> Baixa:
        return self.vender_lote(quantidade, valor, self.tirar_isento(quantidade))

    def vender_tributado(self, quantidade: Decimal, valor: Decimal) -> Baixa:
        """Venda de dentro do lote tributado, pelo PM dele, com o lote isento na mesma posição."""
        tem = self.quantidade - self.isento_quantidade
        total = self.custo - self.isento_custo
        custo = total if quantidade == tem else total * quantidade / tem
        return self.vender_lote(quantidade, valor, custo)

    def repartir_isento(self, quantidade: Decimal, custo: Decimal) -> None:
        """Desdobramento, grupamento e redução de capital atingem os dois lotes pela quantidade.
        Chamar antes de `ajustar`."""
        if self.isento_quantidade and self.quantidade:
            parte = self.isento_quantidade / self.quantidade
            self.isento_quantidade += quantidade * parte
            self.isento_custo += custo * parte

    def fim_da_isencao(self) -> Decimal:
        """Lei 13.043/2014 art. 17 §6º: o lote isento sobe ao maior entre o custo pago e quantidade
        x média dos 30 pregões, e os lotes se fundem num PM só (art. 58). Devolve o aumento."""
        media = self.pme_media
        aumento = ZERO
        if media is not None:
            aumento = max(self.isento_quantidade * media - self.isento_custo, ZERO)
        self.custo += aumento
        self.isento_quantidade = ZERO
        self._limpar_isento()
        return aumento

    def virar_po(self) -> Baixa:
        """Vencimento sem exercício: titular perde o custo; lançador ganha a receita."""
        baixa = Baixa(resultado=-self.custo, custo_baixa=self.custo)
        self.quantidade, self.custo = ZERO, ZERO
        return baixa
