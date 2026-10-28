"""Notas para o suporte: cópia anonimizada do texto das notas (``coir.parsing.anonimizar``).

Quem tem uma nota que o COIR não lê (ou lê errado) manda o texto anonimizado, não o PDF. A cópia
só é gravada se passar nos portões de fidelidade e de privacidade. O mapa de anonimização (segredo
e contas/CPFs fictícios) fica na pasta de dados, em ``_privado``, e nunca vai junto: com ele, a
mesma conta real vira sempre a mesma fictícia em todos os envios.
"""

import logging
import tempfile
import zipfile
from collections.abc import Sequence
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

from pydantic import ValidationError

from coir.erros import AnonimizacaoError
from coir.parsing.anonimizar import MapaAnonimizacao, conferir
from coir.parsing.texto_pdf import PdfError, gravar_texto, ler_texto, paginas

log = logging.getLogger(__name__)

PASTA_SUPORTE = "suporte"
ARQUIVO_MAPA = Path("_privado") / "mapa-anonimizacao.json"
ARQUIVO_TERMOS = "termos-sensiveis.txt"  # a mesma lista local de tools/verificar_dados_pessoais


def carregar_mapa(caminho: Path) -> MapaAnonimizacao:
    """O mapa gravado ou, na primeira vez, um novo com segredo aleatório."""
    if not caminho.is_file():
        return MapaAnonimizacao.novo()
    try:
        return MapaAnonimizacao.model_validate_json(caminho.read_text(encoding="utf-8"))
    except (OSError, ValidationError) as erro:
        raise AnonimizacaoError(f"mapa de anonimização ilegível: {caminho.name}") from erro


def salvar_mapa(mapa: MapaAnonimizacao, caminho: Path) -> None:
    caminho.parent.mkdir(parents=True, exist_ok=True)
    caminho.write_text(mapa.model_dump_json(indent=2), encoding="utf-8")


def termos_sensiveis(dados: Path) -> list[str]:
    """Termos da lista local (um por linha; ``#`` comenta), sem caixa. Ausente: lista vazia."""
    caminho = dados / ARQUIVO_TERMOS
    if not caminho.is_file():
        return []
    linhas = caminho.read_text(encoding="utf-8").splitlines()
    return [t.strip().casefold() for t in linhas if t.strip() and not t.lstrip().startswith("#")]


@dataclass(frozen=True)
class ResultadoArquivo:
    origem: str
    destino: Path | None  # None: recusado (no pacote do suporte, o nome dentro do .zip)
    recusas: tuple[str, ...] = ()


def anonimizar_arquivos(
    arquivos: Sequence[Path],
    saida: Path,
    mapa_caminho: Path,
    sensiveis: Sequence[str] = (),
    nomes: Sequence[str] | None = None,
) -> list[ResultadoArquivo]:
    """Grava em ``saida`` o texto anonimizado de cada nota (PDF ou ``.txt``) que passar nos
    portões; as recusas vêm com o motivo (nota e campo, nunca o valor). ``nomes``: nome de cada
    cópia (padrão: o do arquivo, com ``.txt``)."""
    mapa = carregar_mapa(mapa_caminho)
    saida.mkdir(parents=True, exist_ok=True)
    resultados: list[ResultadoArquivo] = []
    for i, arquivo in enumerate(arquivos):
        nome = nomes[i] if nomes is not None else f"{arquivo.stem}.txt"
        try:
            textos = paginas(arquivo)
        except PdfError as erro:
            resultados.append(ResultadoArquivo(arquivo.name, None, (str(erro),)))
            continue
        copia = conferir(nome, textos, mapa, sensiveis)
        if not copia.aceita:
            resultados.append(ResultadoArquivo(arquivo.name, None, tuple(copia.recusas)))
            continue
        destino = saida / nome
        gravar_texto(destino, copia.textos)
        if ler_texto(destino) != copia.textos:  # o arquivo tem de dar exatamente o texto conferido
            destino.unlink()
            resultados.append(ResultadoArquivo(arquivo.name, None, ("cópia gravada diferente",)))
            continue
        resultados.append(ResultadoArquivo(arquivo.name, destino))
    salvar_mapa(mapa, mapa_caminho)
    aceitos = sum(r.destino is not None for r in resultados)
    log.info(
        "notas anonimizadas: %d gravada(s), %d recusada(s)", aceitos, len(resultados) - aceitos
    )
    return resultados


@dataclass
class PacoteSuporte:
    arquivo: Path | None  # .zip com as cópias aceitas (None: nenhuma)
    resultados: list[ResultadoArquivo] = field(default_factory=list)


def pacote_suporte(
    dados: Path, recebidos: Sequence[tuple[str, bytes]], agora: datetime
) -> PacoteSuporte:
    """Opção "Enviar notas para suporte" (Configurações): os PDFs escolhidos viram textos
    anonimizados num .zip em ``<dados>/suporte``. As cópias levam nomes neutros (nota-01.txt...):
    o nome do PDF pode ter o nome do cliente. Os PDFs não ficam guardados."""
    destino = dados / PASTA_SUPORTE
    with tempfile.TemporaryDirectory(prefix="coir-suporte-") as tmp:
        pasta = Path(tmp)
        origens: list[Path] = []
        for i, (nome, conteudo) in enumerate(recebidos, start=1):
            if Path(nome).suffix.lower() != ".pdf":
                raise AnonimizacaoError(f"não é PDF: {Path(nome).name}")
            origem = pasta / "pdf" / f"nota-{i:02d}.pdf"
            origem.parent.mkdir(exist_ok=True)
            origem.write_bytes(conteudo)
            origens.append(origem)
        nomes = [f"nota-{i:02d}.txt" for i in range(1, len(origens) + 1)]
        resultados = anonimizar_arquivos(
            origens, pasta / "txt", dados / ARQUIVO_MAPA, termos_sensiveis(dados), nomes
        )
        resultados = [
            ResultadoArquivo(Path(nome).name, r.destino and Path(r.destino.name), r.recusas)
            for (nome, _), r in zip(recebidos, resultados, strict=True)
        ]
        aceitos = [pasta / "txt" / r.destino.name for r in resultados if r.destino is not None]
        if not aceitos:
            return PacoteSuporte(None, resultados)
        destino.mkdir(parents=True, exist_ok=True)
        zip_ = destino / f"notas-anonimizadas-{agora:%Y%m%d-%H%M%S}.zip"
        with zipfile.ZipFile(zip_, "w", zipfile.ZIP_DEFLATED) as z:
            for txt in aceitos:
                z.write(txt, txt.name)
    return PacoteSuporte(zip_, resultados)
