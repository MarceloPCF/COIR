"""Mapeamento das fichas do programa IRPF de um exercício (data/irpf/<exercício>.yaml, regra R17).

Só diz onde cada valor do COIR vai no programa (linha, grupo e código); nenhum valor fiscal.
"""

from decimal import Decimal
from typing import Literal

from pydantic import BaseModel, ConfigDict, model_validator

from coir.domain.fiscal import Grupo


class Modelo(BaseModel):
    model_config = ConfigDict(frozen=True, extra="forbid")


class LinhaRendaVariavel(Modelo):
    rotulo: str  # como o programa imprime ("Mercado à Vista - Ações")
    grupos: tuple[Grupo, ...]


class SecaoRendaVariavel(Modelo):
    fonte: str
    linhas: tuple[LinhaRendaVariavel, ...]

    @model_validator(mode="after")
    def _cada_grupo_uma_linha(self) -> "SecaoRendaVariavel":
        vistos = [g for linha in self.linhas for g in linha.grupos]
        faltam = set(Grupo) - set(vistos)
        if len(vistos) != len(set(vistos)) or faltam:
            raise ValueError(f"cada grupo do COIR em exatamente uma linha (faltam {faltam})")
        return self


TipoBem = Literal["ACAO", "UNIT", "BDR", "OPCAO", "ETF", "ETF_RF", "FII", "FIAGRO",
                  "JCP_A_RECEBER"]  # fmt: skip


class ItemBem(Modelo):
    tipo: TipoBem
    grupo: str
    codigo: str
    descricao: str
    cnpj: Literal["emissor", "fundo", "nenhum"]
    nota: str = ""


class SecaoBens(Modelo):
    fonte: str
    itens: tuple[ItemBem, ...]

    @model_validator(mode="after")
    def _um_por_tipo(self) -> "SecaoBens":
        tipos = [i.tipo for i in self.itens]
        if len(tipos) != len(set(tipos)):
            raise ValueError("tipo de bem repetido")
        return self

    def item(self, tipo: TipoBem) -> ItemBem:
        return next(i for i in self.itens if i.tipo == tipo)


class ItemDivida(Modelo):
    tipo: Literal["POSICAO_VENDIDA"]
    codigo: str
    descricao: str
    nota: str = ""


class SecaoDividas(Modelo):
    fonte: str
    itens: tuple[ItemDivida, ...]


class LinhaRendimento(Modelo):
    chave: str
    codigo: str
    descricao: str
    por_fonte: bool = True
    parcial: bool = False  # o código também recebe rendimentos que o COIR não conhece (99, 06)
    nota: str = ""


class SecaoRendimentos(Modelo):
    fonte: str
    linhas: tuple[LinhaRendimento, ...]

    @model_validator(mode="after")
    def _chaves_unicas(self) -> "SecaoRendimentos":
        chaves = [x.chave for x in self.linhas]
        if len(chaves) != len(set(chaves)):
            raise ValueError("chave repetida")
        return self

    def linha(self, chave: str) -> LinhaRendimento:
        try:
            return next(x for x in self.linhas if x.chave == chave)
        except StopIteration as erro:
            raise KeyError(f"linha {chave!r} sem mapeamento") from erro


class MapeamentoIrpf(Modelo):
    exercicio: int
    ano_calendario: int
    provisorio: bool = False  # programa do exercício ainda não publicado: o relatório avisa
    renda_variavel: SecaoRendaVariavel
    bens: SecaoBens
    dividas: SecaoDividas | None = None  # sem a seção, posição vendida não é declarada
    isentos: SecaoRendimentos
    exclusivos: SecaoRendimentos
    imposto_pago: SecaoRendimentos

    @model_validator(mode="after")
    def _ano(self) -> "MapeamentoIrpf":
        if self.ano_calendario != self.exercicio - 1:
            raise ValueError("ano-calendário deve ser o exercício menos 1")
        return self

    def linha_do_grupo(self, grupo: Grupo) -> str:
        return next(x.rotulo for x in self.renda_variavel.linhas if grupo in x.grupos)


# ------------------------------------------------------------------------- fichas geradas
class BemIrpf(Modelo):
    """Um item de Bens e Direitos (um por ativo; P&R IRPF 2026 pergunta 720)."""

    tipo: TipoBem
    grupo: str
    codigo: str
    ticker: str | None  # código de negociação (bens negociados em bolsa)
    cnpj: str | None
    discriminacao: str
    anterior: Decimal  # situação em 31/12 do ano anterior
    atual: Decimal  # situação em 31/12 do ano-calendário


class DividaIrpf(Modelo):
    """Posição vendida em aberto (opção lançada, venda a descoberto), pelo valor recebido."""

    codigo: str
    ticker: str
    discriminacao: str
    anterior: Decimal  # situação em 31/12 do ano anterior
    atual: Decimal


class RendimentoIrpf(Modelo):
    """Uma linha de rendimento (isento, exclusivo) ou de imposto pago, por fonte pagadora."""

    ficha: Literal["isentos", "exclusivos", "imposto_pago"]
    chave: str  # linha do mapeamento (dividendos, jcp...)
    codigo: str
    cnpj: str | None  # None: linha sem fonte (código 20) ou CNPJ não encontrado
    fonte: str  # nome da fonte pagadora (ou descrição da linha)
    valor: Decimal


class LinhaFicha(Modelo):
    rotulo: str  # como o programa imprime
    comuns: Decimal
    day_trade: Decimal | None = None  # None: linha de uma coluna só


class MesFichaIrpf(Modelo):
    mes: int
    linhas: tuple[LinhaFicha, ...]

    def valor(self, rotulo: str) -> LinhaFicha:
        return next(x for x in self.linhas if x.rotulo == rotulo)


class FichasIrpf(Modelo):
    """O que o COIR preenche na declaração de um titular e ano (relatório, tela 14 e comparação
    com a DIRPF entregue)."""

    exercicio: int
    ano: int
    provisorio: bool
    bens: tuple[BemIrpf, ...]
    dividas: tuple[DividaIrpf, ...] = ()
    rendimentos: tuple[RendimentoIrpf, ...]
    renda_variavel: tuple[MesFichaIrpf, ...]
    fii: tuple[MesFichaIrpf, ...]
    avisos: tuple[str, ...] = ()
