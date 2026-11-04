"""Perfis de corretora (profiles/*.yaml) e classificação de cada página do PDF.

Um PDF pode ter várias notas e misturar Bovespa e BM&F; por isso corretora e segmento são decididos
por página, pelas âncoras do perfil (razão social, CNPJ) e pelo cabeçalho de colunas dos negócios.
Sem âncora (a corretora mudou o nome impresso, como a Clear em 2026), vale o código da corretora
na B3, impresso antes da conta do cliente ("3-5 000000" na XP, "85-0 ..." na BTG).
"""

import re
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
# opção flexível (BTG): "Mercadoria:" e depois a modalidade, ex.: "FCA - Com Garantia"
# código da corretora na B3 e dígito antes da conta do cliente, como em bovespa.RE_CONTA
RE_CODIGO_CORRETORA = re.compile(r"^\s*(?P<codigo>\d{1,4})-\d\s+\d+\s*$", re.MULTILINE)
RE_OPCAO_FLEXIVEL = re.compile(r"^\s*F[A-Z]{2}\s*-\s*(?:Com|Sem)\s+Garantia", re.MULTILINE)


class Perfil(BaseModel):
    model_config = ConfigDict(frozen=True, extra="forbid")

    corretora: str
    prioridade: int
    ancoras: tuple[str, ...]
    gramaticas: dict[Segmento, Gramatica]
    codigos: tuple[str, ...] = ()  # código da corretora na B3 (sem o dígito)


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


def _pelo_codigo(texto: str, perfis: tuple[Perfil, ...]) -> Perfil:
    """Segunda tentativa, sem nome nem CNPJ conhecidos: o código da corretora na B3."""
    codigos = {m["codigo"] for m in RE_CODIGO_CORRETORA.finditer(texto)}
    achados = [p for p in perfis if codigos & set(p.codigos)]
    if len({p.corretora for p in achados}) == 1:
        return achados[0]
    impresso = f" (código {', '.join(sorted(codigos))} na B3)" if codigos else ""
    raise ExtracaoError(
        f"corretora não identificada{impresso}: nem o nome, nem o CNPJ, nem o código da B3 "
        "batem com uma corretora conhecida. Mande a nota pela tela Suporte"
    )


def classificar(texto: str, perfis: tuple[Perfil, ...] | None = None) -> Pagina:
    compacto = compactar(texto)
    todos = perfis or carregar_perfis()
    perfil = next((p for p in todos if any(compactar(a) in compacto for a in p.ancoras)), None)
    if perfil is None:
        perfil = _pelo_codigo(texto, todos)
    flexivel = "Mercadoria:" in texto and RE_OPCAO_FLEXIVEL.search(texto)
    tipo = TipoPagina.OPCAO_FLEXIVEL if flexivel else TipoPagina.NOTA
    segmento = Segmento.BMF if "c/vmercadoria" in compacto else Segmento.BOVESPA
    gramatica = perfil.gramaticas.get(segmento)
    if gramatica is None:
        raise ExtracaoError(f"{perfil.corretora}: perfil sem gramática para {segmento.value}")
    return Pagina(perfil, segmento, gramatica, tipo)
