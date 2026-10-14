"""Perfis de corretora (profiles/*.yaml) e classificação de cada página do PDF.

Um PDF pode ter várias notas e misturar Bovespa e BM&F; por isso corretora e segmento são decididos
por página, pelas âncoras do perfil (razão social, CNPJ) e pelo cabeçalho de colunas dos negócios.
"""

from dataclasses import dataclass
from enum import StrEnum
from functools import cache
from importlib import resources
from typing import Literal

import yaml
from pydantic import BaseModel, ConfigDict

from coir.domain.nota import Segmento
from coir.erros import ExtracaoError
from coir.parsing.numeros import compactar

Gramatica = Literal["agora", "xp", "btg"]


class Perfil(BaseModel):
    model_config = ConfigDict(frozen=True, extra="forbid")

    corretora: str
    prioridade: int
    ancoras: tuple[str, ...]
    gramaticas: dict[Segmento, Gramatica]


class TipoPagina(StrEnum):
    NOTA = "nota"
    OPCAO_FLEXIVEL = "opcao_flexivel"  # fora do escopo da v3.0: avisada, nunca ignorada calada


@dataclass(frozen=True)
class Pagina:
    perfil: Perfil
    segmento: Segmento
    gramatica: Gramatica
    tipo: TipoPagina


@cache
def carregar_perfis() -> tuple[Perfil, ...]:
    pasta = resources.files("coir.parsing").joinpath("profiles")
    perfis = [
        Perfil.model_validate(yaml.safe_load(arquivo.read_text(encoding="utf-8")))
        for arquivo in pasta.iterdir()
        if arquivo.name.endswith(".yaml")
    ]
    return tuple(sorted(perfis, key=lambda p: (p.prioridade, p.corretora)))


def classificar(texto: str, perfis: tuple[Perfil, ...] | None = None) -> Pagina:
    compacto = compactar(texto)
    for perfil in perfis or carregar_perfis():
        if any(compactar(a) in compacto for a in perfil.ancoras):
            break
    else:
        raise ExtracaoError("página sem perfil de corretora conhecido (nenhuma âncora encontrada)")
    if "mercadoria:" in compacto and "fca" in compacto:
        tipo = TipoPagina.OPCAO_FLEXIVEL
    else:
        tipo = TipoPagina.NOTA
    segmento = Segmento.BMF if "c/vmercadoria" in compacto else Segmento.BOVESPA
    gramatica = perfil.gramaticas.get(segmento)
    if gramatica is None:
        raise ExtracaoError(f"{perfil.corretora}: perfil sem gramática para {segmento.value}")
    return Pagina(perfil, segmento, gramatica, tipo)
