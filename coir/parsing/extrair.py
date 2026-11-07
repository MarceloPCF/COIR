"""PDF -> notas extraídas e conciliadas.

Cada página é classificada (corretora, segmento); páginas consecutivas com o mesmo número de nota e
data e folhas em sequência (1, 2, ...) formam uma nota. Os negócios vêm de todas as folhas; resumo e
conta, da última. Nota que não concilia, sem conta impressa ou repetida com conteúdo diferente não
entra: vai para `rejeitadas` com o motivo. Nada é descartado em silêncio.

Corretora sem perfil: as páginas são tentadas por cada layout padrão e vale o que mais notas fechar.
A conciliação é o portão: a nota lida entra com um aviso, e a que não fecha é recusada dizendo que
o COIR tentou e o que não deu certo.
"""

import re
from collections.abc import Callable, Sequence
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path

from coir.domain.nota import Cliente, NotaExtraida, Segmento
from coir.erros import ExtracaoError
from coir.parsing import bmf, bovespa
from coir.parsing.cliente import cliente
from coir.parsing.conciliar import conciliar
from coir.parsing.numeros import compactar
from coir.parsing.paginas import (
    GRAMATICAS_PADRAO,
    CorretoraDesconhecida,
    Gramatica,
    Pagina,
    TipoPagina,
    classificar,
    nome_generico,
    pagina_generica,
)
from coir.parsing.texto_pdf import PdfError
from coir.parsing.texto_pdf import paginas as ler_paginas

RE_CONTA_BMF_XP = re.compile(r"^\s*(?P<conta>\d+)\s*$")
# BTG: número inteiro no fim da linha do endereço (não o "0,00" de um negócio)
RE_CONTA_BMF_BTG = re.compile(r"(?:^|\s)(\d{3,})\s*$")


@dataclass
class Rejeitada:
    arquivo: str
    paginas: tuple[int, ...]
    motivo: str
    nota: str = ""  # número da nota, quando o cabeçalho foi lido
    data: date | None = None
    erro: Exception | None = field(default=None, compare=False, repr=False)  # para a auditoria


@dataclass
class ResultadoExtracao:
    notas: list[NotaExtraida] = field(default_factory=list)
    rejeitadas: list[Rejeitada] = field(default_factory=list)
    avisos: list[str] = field(default_factory=list)
    # cliente impresso em cada nota, pela chave (dado pessoal: fora de NotaExtraida)
    clientes: dict[tuple[object, ...], Cliente] = field(default_factory=dict)

    def juntar(self, outro: "ResultadoExtracao") -> None:
        self.notas += outro.notas
        self.rejeitadas += outro.rejeitadas
        self.avisos += outro.avisos
        self.clientes.update(outro.clientes)


@dataclass
class _Grupo:
    pagina: Pagina
    cabecalho: bovespa.Cabecalho
    indices: list[int]
    folhas: list[int]


def extrair_pdf(caminho: Path, senhas: Sequence[str] = ()) -> ResultadoExtracao:
    """Notas de um PDF (``senhas``: tentadas se o PDF for protegido); ilegível vira rejeição."""
    try:
        textos = ler_paginas(caminho, senhas)
    except PdfError as erro:
        return ResultadoExtracao(rejeitadas=[Rejeitada(caminho.name, (), str(erro), erro=erro)])
    return extrair_textos(caminho.name, textos)


def extrair_textos(arquivo: str, textos: Sequence[str]) -> ResultadoExtracao:
    res = ResultadoExtracao()
    desconhecidas = _ler(arquivo, textos, range(len(textos)), classificar, res)
    if desconhecidas:
        _pelo_layout_padrao(arquivo, textos, desconhecidas, res)
    return res


def _generica(gramatica: Gramatica, nome: str) -> Callable[[str], Pagina]:
    return lambda texto: pagina_generica(texto, gramatica, nome)


def _pelo_layout_padrao(
    arquivo: str,
    textos: Sequence[str],
    desconhecidas: dict[int, CorretoraDesconhecida],
    res: ResultadoExtracao,
) -> None:
    """Corretora sem perfil: tenta cada layout padrão e fica com o que mais notas fechar. Página
    em que nenhum layout achou uma nota mantém a recusa original: pode nem ser nota."""
    nome = nome_generico(textos[i] for i in desconhecidas)
    if nome is None:  # sem identificar a corretora, a conta poderia ser de outra
        res.rejeitadas += [
            Rejeitada(arquivo, (i + 1,), str(erro), erro=erro) for i, erro in desconhecidas.items()
        ]
        return
    tentativas: list[ResultadoExtracao] = []
    for gramatica in GRAMATICAS_PADRAO:
        tentativa = ResultadoExtracao()
        _ler(arquivo, textos, list(desconhecidas), _generica(gramatica, nome), tentativa)
        tentativas.append(tentativa)
    melhor = max(tentativas, key=lambda t: (len(t.notas), sum(bool(r.nota) for r in t.rejeitadas)))
    for r in melhor.rejeitadas:
        original = desconhecidas[r.paginas[0] - 1]
        if r.nota:
            r.motivo = (
                f"{original}. O COIR tentou ler pelo layout padrão e não conseguiu tudo certo "
                f"({r.motivo}): nada desta nota foi gravado"
            )
        else:
            r.motivo, r.erro = str(original), original
    if melhor.notas:
        melhor.avisos.append(
            f"{arquivo}: corretora sem perfil no COIR ({nome}), lida pelo layout padrão. Os "
            f"totais de {len(melhor.notas)} nota(s) conferem, mas confira as operações; se algo "
            "estiver errado, mande a nota pela tela Suporte"
        )
    res.juntar(melhor)


def _ler(
    arquivo: str,
    textos: Sequence[str],
    indices: Sequence[int],
    classifica: Callable[[str], Pagina],
    res: ResultadoExtracao,
) -> dict[int, CorretoraDesconhecida]:
    """Lê as páginas `indices`; devolve as de corretora desconhecida, com o erro."""
    desconhecidas: dict[int, CorretoraDesconhecida] = {}
    grupos: list[_Grupo] = []
    for i in indices:
        texto = textos[i]
        try:
            pagina = classifica(texto)
            if pagina.tipo is TipoPagina.OPCAO_FLEXIVEL:
                res.avisos.append(
                    f"{arquivo} p. {i + 1}: opção flexível (fora do escopo; lançar à mão)"
                )
                continue
            cab = bovespa.cabecalho(texto, pagina.gramatica)
        except CorretoraDesconhecida as erro:
            desconhecidas[i] = erro
            continue
        except ExtracaoError as erro:
            res.rejeitadas.append(Rejeitada(arquivo, (i + 1,), str(erro), erro=erro))
            continue
        anterior = grupos[-1] if grupos else None
        if (
            anterior
            and anterior.indices[-1] == i - 1
            and anterior.pagina.perfil == pagina.perfil
            and anterior.pagina.segmento is pagina.segmento
            and (anterior.cabecalho.numero, anterior.cabecalho.data_pregao)
            == (cab.numero, cab.data_pregao)
            and cab.folha == anterior.folhas[-1] + 1
        ):
            anterior.indices.append(i)
            anterior.folhas.append(cab.folha)
        else:
            grupos.append(_Grupo(pagina, cab, [i], [cab.folha]))
    for g in grupos:
        rejeitada = Rejeitada(arquivo, tuple(i + 1 for i in g.indices), "", g.cabecalho.numero,
                              g.cabecalho.data_pregao)  # fmt: skip
        try:
            if g.folhas[0] != 1:
                raise ExtracaoError(f"folhas fora de sequência: {g.folhas}")
            nota = conciliar(_montar(arquivo, g, [textos[i] for i in g.indices]))
            res.notas.append(nota)
            res.avisos += nota.avisos
            impresso = cliente(textos[g.indices[-1]])  # da última folha, como a conta
            if impresso is not None:
                res.clientes[nota.chave] = impresso
        except ExtracaoError as erro:
            rejeitada.motivo, rejeitada.erro = f"nota {g.cabecalho.numero}: {erro}", erro
            res.rejeitadas.append(rejeitada)
    return desconhecidas


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
        # 2022+: a conta vem logo após "Codigo do Cliente"; antes, na última linha
        for i, linha in enumerate(linhas[:-1]):
            if compactar(linha) == "codigodocliente":
                m = RE_CONTA_BMF_XP.match(linhas[i + 1])
                if m:
                    return str(int(m["conta"]))
        m = RE_CONTA_BMF_XP.match(linhas[-1])
        if m:
            return str(int(m["conta"]))
        raise ExtracaoError("conta do cliente (BM&F XP) não impressa")
    # BTG: no fim da linha seguinte a "... Código cliente" (depois do endereço); sem ela, na
    # última linha. A última linha pode ser o rodapé (telefone da ouvidoria, endereço da BTG)
    for i, linha in enumerate(linhas[:-1]):
        if compactar(linha).endswith("codigocliente"):
            m = RE_CONTA_BMF_BTG.search(linhas[i + 1])
            if m:
                return str(int(m.group(1)))
    m = re.search(r"(\d+)\s*$", linhas[-1])
    if m:
        return str(int(m.group(1)))
    raise ExtracaoError("conta do cliente (BM&F BTG) não encontrada")


def extrair_pasta(pasta: Path, sufixo: str = ".pdf") -> ResultadoExtracao:
    """Todos os PDFs de nota da pasta, em ordem; outros arquivos e subpastas geram aviso.
    ``sufixo=".txt"``: os textos de nota (base anonimizada de desenvolvimento)."""
    total = ResultadoExtracao()
    for caminho in sorted(pasta.iterdir()):
        if not (caminho.is_file() and caminho.suffix.lower() == sufixo):
            total.avisos.append(f"ignorado (não é PDF de nota): {caminho.name}")
            continue
        total.juntar(extrair_pdf(caminho))
    return sem_duplicadas(total)


def sem_duplicadas(res: ResultadoExtracao) -> ResultadoExtracao:
    """Mesma nota (chave) em dois lugares: idêntica fica uma (aviso); diferente, saem as duas."""
    por_chave: dict[tuple[object, ...], list[NotaExtraida]] = {}
    for n in res.notas:
        por_chave.setdefault(n.chave, []).append(n)
    notas: list[NotaExtraida] = []
    for repetidas in por_chave.values():
        primeira = repetidas[0]
        conteudo = {
            (n.negocios_bovespa, n.negocios_bmf, n.resumo_bovespa, n.resumo_bmf) for n in repetidas
        }
        if len(repetidas) == 1:
            notas.append(primeira)
        elif len(conteudo) == 1:
            notas.append(primeira)
            onde = ", ".join(f"{n.arquivo} p. {list(n.paginas)}" for n in repetidas)
            res.avisos.append(f"nota {primeira.numero} repetida (idêntica): {onde}; usada uma")
        else:
            for n in repetidas:
                res.rejeitadas.append(
                    Rejeitada(n.arquivo, n.paginas, f"nota {n.numero}: repetida com conteúdo "
                              "diferente em outro arquivo/página", n.numero, n.data_pregao)
                )  # fmt: skip
    res.notas = notas
    return res
