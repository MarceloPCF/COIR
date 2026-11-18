"""Tabelas curadas pelo projeto, cada linha com a fonte conferida à mão."""

from pathlib import Path

from coir.domain.cpf import cnpj_valido
from coir.erros import ConfiguracaoError
from coir.fontes.leitura import opcionais
from coir.fontes.modelo import CURADORIA, Emissor, TrocaCodigo

EMISSORES = "emissores.csv"
SAIDAS = "saidas.csv"
SAIDAS_FORMATO_2 = "saidas_b3.csv"  # o mesmo arquivo, com o nome antigo
TROCAS = "trocas_ticker.csv"
COLUNAS_SAIDAS = ("ticker", "ultimo_negocio", "situacao", "destino", "relacao",
                  "dinheiro_por_acao", "confianca", "fonte", "observacao")  # fmt: skip


def emissores(pasta: Path) -> list[Emissor]:
    """CNPJ e razão social conferidos à mão; CNPJ com dígito errado falha alto."""
    saida: list[Emissor] = []
    for r in opcionais(pasta / EMISSORES):
        if not cnpj_valido(r["cnpj"]):
            raise ConfiguracaoError(f"{EMISSORES}: CNPJ inválido para {r['emissor']}")
        saida.append(Emissor(emissor=r["emissor"], cnpj=r["cnpj"], razao_social=r["razao_social"],
                             tipo="", visto_ate="", origem=CURADORIA))  # fmt: skip
    return saida


def saidas(pasta: Path) -> dict[str, dict[str, str]]:
    """O que houve com o ativo que parou de negociar, por ticker, com a fonte de cada linha."""
    nome = SAIDAS if (pasta / SAIDAS).is_file() else SAIDAS_FORMATO_2
    tabela = opcionais(pasta / nome)
    if tabela and set(COLUNAS_SAIDAS) - set(tabela[0]):
        raise ConfiguracaoError(f"{nome}: colunas ausentes")
    return {r["ticker"]: r for r in tabela}


def trocas(pasta: Path) -> list[TrocaCodigo]:
    """Trocas de código de negociação (o ticker impresso na nota e o de depois)."""
    return [
        TrocaCodigo(ticker_impresso=t["ticker_impresso"], ticker=t["ticker"], ate=t["ate"],
                    isin_antigo=t.get("isin_antigo", ""), isin_novo=t.get("isin_novo", ""),
                    origem=CURADORIA)
        for t in opcionais(pasta / TROCAS)
    ]  # fmt: skip
