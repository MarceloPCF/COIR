"""PDF -> notas extraídas e conciliadas.

Cada página é classificada (corretora, segmento); páginas consecutivas com o mesmo número de nota e
data formam uma nota (folhas). Os negócios vêm de todas as folhas; resumo e conta, da última. Nota
que não concilia não entra: vai para `rejeitadas` com o motivo.
"""

import re
from collections.abc import Sequence
from dataclasses import dataclass, field
from pathlib import Path

from coir.domain.nota import NotaExtraida, Segmento
from coir.erros import ExtracaoError
from coir.parsing import bmf, bovespa
from coir.parsing.conciliar import conciliar
from coir.parsing.numeros import compactar
from coir.parsing.paginas import Pagina, TipoPagina, classificar
from coir.parsing.texto_pdf import paginas as ler_paginas

RE_CONTA_BMF_XP = re.compile(r"^\s*(?P<conta>\d+)\s*$")


@dataclass
class Rejeitada:
    arquivo: str
    paginas: tuple[int, ...]
    motivo: str


@dataclass
class ResultadoExtracao:
    notas: list[NotaExtraida] = field(default_factory=list)
    rejeitadas: list[Rejeitada] = field(default_factory=list)
    avisos: list[str] = field(default_factory=list)


@dataclass
class _Grupo:
    pagina: Pagina
    cabecalho: bovespa.Cabecalho
    indices: list[int]


def extrair_pdf(caminho: Path) -> ResultadoExtracao:
    return extrair_textos(caminho.name, ler_paginas(caminho))


def extrair_textos(arquivo: str, textos: Sequence[str]) -> ResultadoExtracao:
    res = ResultadoExtracao()
    grupos: list[_Grupo] = []
    for i, texto in enumerate(textos):
        try:
            pagina = classificar(texto)
            if pagina.tipo is TipoPagina.OPCAO_FLEXIVEL:
                res.avisos.append(
                    f"{arquivo} p. {i + 1}: opção flexível (fora do escopo; lançar à mão)"
                )
                continue
            cab = bovespa.cabecalho(texto, pagina.gramatica)
        except ExtracaoError as erro:
            res.rejeitadas.append(Rejeitada(arquivo, (i + 1,), str(erro)))
            continue
        anterior = grupos[-1] if grupos else None
        if (
            anterior
            and anterior.indices[-1] == i - 1
            and anterior.pagina.segmento is pagina.segmento
            and (anterior.cabecalho.numero, anterior.cabecalho.data_pregao)
            == (cab.numero, cab.data_pregao)
        ):
            anterior.indices.append(i)
        else:
            grupos.append(_Grupo(pagina, cab, [i]))
    for g in grupos:
        paginas_nota = tuple(i + 1 for i in g.indices)
        try:
            nota = _montar(arquivo, g, [textos[i] for i in g.indices])
            res.notas.append(conciliar(nota))
        except ExtracaoError as erro:
            res.rejeitadas.append(
                Rejeitada(arquivo, paginas_nota, f"nota {g.cabecalho.numero}: {erro}")
            )
    return res


def _montar(arquivo: str, g: _Grupo, textos: list[str]) -> NotaExtraida:
    ultima = textos[-1]
    comum = {
        "corretora": g.pagina.perfil.corretora,
        "numero": g.cabecalho.numero,
        "data_pregao": g.cabecalho.data_pregao,
        "segmento": g.pagina.segmento,
        "arquivo": arquivo,
        "paginas": tuple(i + 1 for i in g.indices),
    }
    if g.pagina.segmento is Segmento.BOVESPA:
        variante = g.pagina.gramatica
        negocios = [n for t in textos for n in bovespa.negocios(t, variante)]
        if not negocios:
            raise ExtracaoError("nenhuma linha de negócio Bovespa")
        return NotaExtraida(
            **comum,
            conta=bovespa.conta(ultima, variante),
            negocios_bovespa=tuple(negocios),
            resumo_bovespa=bovespa.resumo(ultima, variante),
        )
    variante_bmf = "btg" if g.pagina.gramatica == "btg" else "xp"
    negocios_bmf = [n for t in textos for n in bmf.negocios(t)]
    return NotaExtraida(
        **comum,
        conta=_conta_bmf(ultima, variante_bmf),
        negocios_bmf=tuple(negocios_bmf),
        resumo_bmf=bmf.resumo(ultima, variante_bmf, negocios_bmf),
    )


def _conta_bmf(texto: str, variante: str) -> str:
    linhas = [linha for linha in texto.splitlines() if linha.strip()]
    if variante == "xp":
        # 2022+: a conta vem logo após "Codigo do Cliente"; antes, na última linha. Sem nenhuma das
        # duas (página que termina no CPF), fica vazia e é completada no lote.
        for i, linha in enumerate(linhas[:-1]):
            if compactar(linha) == "codigodocliente":
                m = RE_CONTA_BMF_XP.match(linhas[i + 1])
                if m:
                    return str(int(m["conta"]))
        m = RE_CONTA_BMF_XP.match(linhas[-1])
        return str(int(m["conta"])) if m else ""
    m = re.search(r"(\d+)\s*$", linhas[-1])
    if m:
        return str(int(m.group(1)))
    raise ExtracaoError("conta do cliente (BM&F BTG) não encontrada")


def extrair_pasta(pasta: Path) -> ResultadoExtracao:
    """Todos os PDFs de nota da pasta (ordem alfabética), com as contas ausentes completadas."""
    total = ResultadoExtracao()
    for caminho in sorted(pasta.glob("*.pdf")):
        r = extrair_pdf(caminho)
        total.notas += r.notas
        total.rejeitadas += r.rejeitadas
        total.avisos += r.avisos
    return completar_contas(total)


def completar_contas(res: ResultadoExtracao) -> ResultadoExtracao:
    """Nota sem conta impressa recebe a conta da corretora/segmento se ela for única no lote."""
    vistas: dict[tuple[str, Segmento], set[str]] = {}
    for n in res.notas:
        if n.conta:
            vistas.setdefault((n.corretora, n.segmento), set()).add(n.conta)
    notas: list[NotaExtraida] = []
    for n in res.notas:
        if n.conta:
            notas.append(n)
            continue
        candidatas = vistas.get((n.corretora, n.segmento), set())
        if len(candidatas) != 1:
            res.rejeitadas.append(
                Rejeitada(n.arquivo, n.paginas, f"nota {n.numero}: conta não impressa e "
                          f"{len(candidatas)} contas {n.corretora} {n.segmento.value} no lote")
            )  # fmt: skip
            continue
        aviso = f"conta não impressa: usada a única conta {n.corretora} {n.segmento.value} do lote"
        notas.append(n.model_copy(update={"conta": next(iter(candidatas)),
                                          "avisos": (*n.avisos, aviso)}))  # fmt: skip
    res.notas = notas
    return res
