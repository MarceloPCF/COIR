"""Dados de referência por conceito (eventos, proventos, cadastro de instrumentos e emissores...).

Os serviços pedem o conceito, sem saber a fonte: cada origem (curadoria do projeto, CVM, Banco
Central, bolsa) tem o seu provedor em ``provedores/``, e ``composicao`` junta os provedores pela
precedência de cada conceito, registrando as divergências. As tabelas fixas do motor
(``Referencias`` e ``ReferenciasFiscais``) continuam em ``coir.referencias``.
"""

from pathlib import Path

from coir.fontes.composicao import (
    correcoes_por_ticker,
    emissores,
    eventos,
    eventos_e_divergencias,
    instrumentos,
    saidas,
    tickers_por_isin,
    ultimos_negocios,
)
from coir.fontes.modelo import (
    Divergencia,
    Evento,
    Provento,
    ProventoHistorico,
    evento_valido,
)
from coir.fontes.provedores import b3


def proventos(pasta: Path) -> list[Provento]:
    """Proventos em dinheiro anunciados (vazio sem a fonte)."""
    return b3.proventos(pasta)


def proventos_historicos(pasta: Path) -> list[ProventoHistorico]:
    """JCP e restituição de capital de anos antigos (vazio sem a fonte)."""
    return b3.proventos_historicos(pasta)


def validar_pacote(pasta: Path) -> None:
    """Pacote novo: as tabelas lidas só sob demanda também têm de estar íntegras."""
    b3.validar(pasta)


__all__ = [
    "Divergencia",
    "Evento",
    "Provento",
    "ProventoHistorico",
    "correcoes_por_ticker",
    "emissores",
    "evento_valido",
    "eventos",
    "eventos_e_divergencias",
    "instrumentos",
    "proventos",
    "proventos_historicos",
    "saidas",
    "tickers_por_isin",
    "ultimos_negocios",
    "validar_pacote",
]
