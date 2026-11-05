"""Perfis de corretora (profiles/*.yaml) e classificação de cada página do PDF.

Um PDF pode ter várias notas e misturar Bovespa e BM&F; por isso corretora e segmento são decididos
por página, pelas âncoras do perfil (razão social, CNPJ) e pelo cabeçalho de colunas dos negócios.
Sem âncora (a corretora mudou o nome impresso, como a Clear em 2026), vale o código da corretora
na B3, impresso antes da conta do cliente ("3-5 000000" na XP, "85-0 ..." na BTG). Corretora que
nenhum perfil cobre é tentada pelos layouts padrão (`pagina_generica`), e a conciliação decide.
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
# o código no próprio rótulo, sem a conta ("Número da Corretora 88-4 Cliente", na CM Capital): só
# para dizer na mensagem qual código a nota traz; quem identifica a corretora é RE_CODIGO_CORRETORA
RE_CODIGO_NO_ROTULO = re.compile(
    r"N.?mero\s*da\s*Corretora[ \t]+(?P<codigo>\d{1,4})-\d\b", re.IGNORECASE
)
RE_OPCAO_FLEXIVEL = re.compile(r"^\s*F[A-Z]{2}\s*-\s*(?:Com|Sem)\s+Garantia", re.MULTILINE)
GRAMATICAS_PADRAO: tuple[Gramatica, ...] = ("xp", "btg", "agora")  # layouts de nota que o COIR lê


class CorretoraDesconhecida(ExtracaoError):
    """Nenhum perfil cobre a página: nem o nome, nem o CNPJ, nem o código da B3."""


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


def _codigos(texto: str) -> list[str]:
    """Códigos da B3 impressos na página, ao lado da conta ou no rótulo."""
    return sorted(
        {m["codigo"] for m in RE_CODIGO_CORRETORA.finditer(texto)}
        | {m["codigo"] for m in RE_CODIGO_NO_ROTULO.finditer(texto)}
    )


def _pelo_codigo(texto: str, perfis: tuple[Perfil, ...]) -> Perfil:
    """Segunda tentativa, sem nome nem CNPJ conhecidos: o código da corretora na B3."""
    codigos = {m["codigo"] for m in RE_CODIGO_CORRETORA.finditer(texto)}
    achados = [p for p in perfis if codigos & set(p.codigos)]
    if len({p.corretora for p in achados}) == 1:
        return achados[0]
    impresso = ", ".join(_codigos(texto))
    impresso = f" (código {impresso} na B3)" if impresso else ""
    raise CorretoraDesconhecida(
        f"corretora não identificada{impresso}: nem o nome, nem o CNPJ, nem o código da B3 "
        "batem com uma corretora conhecida. Mande a nota pela tela Suporte"
    )


def classificar(texto: str, perfis: tuple[Perfil, ...] | None = None) -> Pagina:
    compacto = compactar(texto)
    todos = perfis or carregar_perfis()
    perfil = next((p for p in todos if any(compactar(a) in compacto for a in p.ancoras)), None)
    if perfil is None:
        perfil = _pelo_codigo(texto, todos)
    return _pagina(texto, compacto, perfil)


def pagina_generica(texto: str, gramatica: Gramatica) -> Pagina:
    """A página de uma corretora que nenhum perfil cobre, lida por um dos layouts padrão. A
    corretora fica com o nome pelo código da B3 impresso, ou genérico sem ele."""
    codigos = _codigos(texto)
    nome = f"CORRETORA {codigos[0]}" if len(codigos) == 1 else "CORRETORA NAO IMPLEMENTADA"
    segmentos: dict[Segmento, Gramatica] = {Segmento.BOVESPA: gramatica}
    if gramatica != "agora":  # o layout da Ágora é só Bovespa
        segmentos[Segmento.BMF] = gramatica
    perfil = Perfil(corretora=nome, prioridade=0, ancoras=(), gramaticas=segmentos)
    return _pagina(texto, compactar(texto), perfil)


def _pagina(texto: str, compacto: str, perfil: Perfil) -> Pagina:
    flexivel = "Mercadoria:" in texto and RE_OPCAO_FLEXIVEL.search(texto)
    tipo = TipoPagina.OPCAO_FLEXIVEL if flexivel else TipoPagina.NOTA
    segmento = Segmento.BMF if "c/vmercadoria" in compacto else Segmento.BOVESPA
    gramatica = perfil.gramaticas.get(segmento)
    if gramatica is None:
        raise ExtracaoError(f"{perfil.corretora}: perfil sem gramática para {segmento.value}")
    return Pagina(perfil, segmento, gramatica, tipo)
