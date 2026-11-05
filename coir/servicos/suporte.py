"""Notas para o suporte: cópia anonimizada do texto das notas (``coir.parsing.anonimizar``) e a
imagem de cada página com tarja sobre o que foi trocado (``coir.parsing.imagens``).

Quem tem uma nota que o COIR não lê (ou lê errado) manda o texto anonimizado e as imagens, não o
PDF. A cópia só é gravada se passar nos portões de fidelidade e de privacidade. Ficam na pasta de
dados, em ``_privado``, e nunca vão junto: o mapa de anonimização (segredo e contas/CPFs fictícios;
com ele, a mesma conta real vira sempre a mesma fictícia) e os dados que cada titular informou
para ocultar (endereço, telefones, e-mail...).

O envio é por e-mail, pelo programa de e-mail do usuário: o COIR prepara a pasta, ele confere as
imagens, o COIR grava o .zip e abre a mensagem já endereçada; o usuário anexa o .zip e envia. Nada
sai do computador sem essa ação dele.
"""

import json
import logging
import re
import tempfile
import zipfile
from collections.abc import Iterable, Sequence
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from pathlib import Path
from urllib.parse import quote

from pydantic import BaseModel, ConfigDict, Field, ValidationError
from sqlalchemy import select
from sqlalchemy.exc import SQLAlchemyError

from coir import auditoria
from coir.auditoria import portao
from coir.config import VERSAO
from coir.db import repositorio as repo
from coir.db.modelos import Conta
from coir.db.sessao import conectar
from coir.domain.cpf import cpf_valido, so_digitos
from coir.erros import AnonimizacaoError, CoirError
from coir.parsing.anonimizar import CAMPOS_LIVRES as CAMPOS_LIVRES  # para a tela
from coir.parsing.anonimizar import DadosPessoais, MapaAnonimizacao, conferir
from coir.parsing.imagens import gravar_png, paginas_tarjadas
from coir.parsing.texto_pdf import (
    SUFIXO_TEXTO,
    PdfError,
    gravar_texto,
    ler_texto,
    paginas,
    senhas_do_cpf,
)

log = logging.getLogger(__name__)

PASTA_SUPORTE = "suporte"
ARQUIVO_MAPA = Path("_privado") / "mapa-anonimizacao.json"
ARQUIVO_DADOS = Path("_privado") / "dados-para-ocultar.json"
ARQUIVO_TERMOS = "termos-sensiveis.txt"  # a mesma lista local de tools/verificar_dados_pessoais
EMAIL_SUPORTE = "suporte.coir@gmail.com"
DIAS_DO_REGISTRO = 30
MAXIMO_DO_REGISTRO = 8_000_000  # em bytes, antes de compactar; ficam as linhas mais recentes
ARQUIVO_DO_REGISTRO = "registro-para-suporte.jsonl"
RE_DATA_DA_LINHA = re.compile(r'\{"t": ?"(\d{4}-\d{2}-\d{2})')


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


# ------------------------------------------------------------------------- dados para ocultar
class _ArquivoDados(BaseModel):
    model_config = ConfigDict(extra="forbid")

    titulares: list[DadosPessoais] = Field(default_factory=list)


def carregar_dados(dados: Path) -> list[DadosPessoais]:
    """Os dados para ocultar que cada titular informou (só os gravados, sem os do banco)."""
    caminho = dados / ARQUIVO_DADOS
    if not caminho.is_file():
        return []
    try:
        return _ArquivoDados.model_validate_json(caminho.read_text(encoding="utf-8")).titulares
    except (OSError, ValidationError) as erro:
        raise AnonimizacaoError(f"dados para ocultar ilegíveis: {caminho.name}") from erro


def salvar_titular(dados: Path, informado: DadosPessoais) -> DadosPessoais:
    """Grava (ou substitui, pelo CPF) os dados para ocultar de um titular."""
    cpf = so_digitos(informado.cpf)
    if not cpf_valido(cpf):
        raise AnonimizacaoError("CPF inválido")
    if not informado.nome.strip():
        raise AnonimizacaoError("informe o nome como ele sai nas notas")
    novo = informado.model_copy(
        update={"cpf": cpf, "contas": [c.strip() for c in informado.contas if c.strip()]}
    )
    lista = [d for d in carregar_dados(dados) if so_digitos(d.cpf) != cpf] + [novo]
    caminho = dados / ARQUIVO_DADOS
    caminho.parent.mkdir(parents=True, exist_ok=True)
    texto = _ArquivoDados(titulares=sorted(lista, key=lambda d: d.nome)).model_dump_json(indent=2)
    caminho.write_text(texto, encoding="utf-8")
    return novo


def do_banco(banco: Path) -> list[DadosPessoais]:
    """Nome, CPF e contas de cada titular do banco (o ponto de partida do que ocultar)."""
    if not banco.is_file():
        return []
    with conectar(banco) as fabrica, fabrica() as s:
        contas: dict[int, list[str]] = {}
        for c in s.scalars(select(Conta).order_by(Conta.numero)):
            contas.setdefault(c.titular_id, []).append(c.numero)
        return [DadosPessoais(cpf=t.cpf, nome=t.nome, contas=contas.get(t.id, []))
                for t in repo.titulares(s)]  # fmt: skip


def dados_para_ocultar(dados: Path, banco: Path) -> list[DadosPessoais]:
    """Os do banco completados com os informados (o informado vale; as contas se somam)."""
    juntos: dict[str, DadosPessoais] = {so_digitos(d.cpf): d for d in do_banco(banco)}
    for d in carregar_dados(dados):
        cpf = so_digitos(d.cpf)
        base = juntos.get(cpf)
        if base is not None:
            contas = list(dict.fromkeys([*base.contas, *d.contas]))
            d = d.model_copy(update={"nome": d.nome or base.nome, "contas": contas})
        juntos[cpf] = d
    return sorted(juntos.values(), key=lambda d: d.nome)


# ------------------------------------------------------------------------------- anonimização
@dataclass(frozen=True)
class ResultadoArquivo:
    origem: str
    destino: Path | None  # None: recusado
    recusas: tuple[str, ...] = ()
    imagens: tuple[Path, ...] = ()
    paginas: int = 0
    ocultados: int = 0  # trechos trocados
    alertas: int = 0  # trechos informados que ficaram à mostra (contorno nas imagens)
    cpfs: frozenset[str] = frozenset()


def anonimizar_arquivos(
    arquivos: Sequence[Path],
    saida: Path,
    mapa_caminho: Path,
    sensiveis: Sequence[str] = (),
    nomes: Sequence[str] | None = None,
    informados: Sequence[DadosPessoais] = (),
    imagens: bool = False,
    senhas: Sequence[str] = (),
) -> list[ResultadoArquivo]:
    """Grava em ``saida`` o texto anonimizado de cada nota (PDF ou ``.txt``) que passar nos
    portões; as recusas vêm com o motivo (nota e campo, nunca o valor). ``nomes``: nome de cada
    cópia (padrão: o do arquivo, com ``.txt``). Com ``imagens``, grava também a imagem tarjada de
    cada página dos PDFs (``<cópia>-pagina-N.png``). PDF protegido: tenta as ``senhas`` e as do
    CPF dos ``informados``."""
    with auditoria.operacao("anonimizacao", arquivos=len(arquivos)) as op:
        resultados = _anonimizar_arquivos(arquivos, saida, mapa_caminho, sensiveis, nomes,
                                          informados, imagens, senhas)  # fmt: skip
        aceitas = sum(r.destino is not None for r in resultados)
        op.dados(aceitas=aceitas, rejeitadas=len(resultados) - aceitas)
        for r in resultados:
            auditoria.evento(
                "suporte.nota", arquivo=r.origem, paginas=[r.paginas], ocultados=r.ocultados,
                alertas=r.alertas, situacao="aceita" if r.destino else "recusada",
                recusas=r.recusas,
            )  # fmt: skip
    return resultados


def _anonimizar_arquivos(
    arquivos: Sequence[Path],
    saida: Path,
    mapa_caminho: Path,
    sensiveis: Sequence[str],
    nomes: Sequence[str] | None,
    informados: Sequence[DadosPessoais],
    imagens: bool,
    senhas: Sequence[str],
) -> list[ResultadoArquivo]:
    tentar = [*senhas, *senhas_do_cpf(d.cpf for d in informados)]
    mapa = carregar_mapa(mapa_caminho)
    saida.mkdir(parents=True, exist_ok=True)
    resultados = [
        _anonimizar_um(a, saida, nomes[i] if nomes is not None else f"{a.stem}.txt",
                       mapa, sensiveis, informados, imagens, tentar)
        for i, a in enumerate(arquivos)
    ]  # fmt: skip
    salvar_mapa(mapa, mapa_caminho)
    aceitos = sum(r.destino is not None for r in resultados)
    log.info(
        "notas anonimizadas: %d gravada(s), %d recusada(s)", aceitos, len(resultados) - aceitos
    )
    return resultados


def _anonimizar_um(
    arquivo: Path,
    saida: Path,
    nome: str,
    mapa: MapaAnonimizacao,
    sensiveis: Sequence[str],
    informados: Sequence[DadosPessoais],
    imagens: bool,
    senhas: Sequence[str],
) -> ResultadoArquivo:
    try:
        textos = paginas(arquivo, senhas)
    except PdfError as erro:
        return ResultadoArquivo(arquivo.name, None, (str(erro),))
    copia = conferir(nome, textos, mapa, sensiveis, informados)
    if not copia.aceita:
        return ResultadoArquivo(arquivo.name, None, tuple(copia.recusas), cpfs=copia.cpfs)
    destino = saida / nome
    gravar_texto(destino, copia.textos)
    if ler_texto(destino) != copia.textos:  # o arquivo tem de dar exatamente o texto conferido
        destino.unlink()
        return ResultadoArquivo(arquivo.name, None, ("cópia gravada diferente",))
    figuras: list[Path] = []
    if imagens and arquivo.suffix.lower() != SUFIXO_TEXTO:
        try:
            tarjadas = paginas_tarjadas(arquivo, textos, copia.trocados, copia.alertas,
                                        senhas=senhas)  # fmt: skip
        except AnonimizacaoError as erro:
            destino.unlink()
            return ResultadoArquivo(arquivo.name, None, (f"imagens: {erro}",))
        for n, imagem in enumerate(tarjadas, start=1):
            figura = saida / f"{destino.stem}-pagina-{n}.png"
            gravar_png(imagem, figura)
            figuras.append(figura)
    return ResultadoArquivo(
        arquivo.name,
        destino,
        imagens=tuple(figuras),
        paginas=len(textos),
        ocultados=sum(_trechos(m) for m in copia.trocados),
        alertas=sum(len(a) for a in copia.alertas),
        cpfs=copia.cpfs,
    )


def _trechos(mascara: bytes) -> int:
    return sum(1 for i, b in enumerate(mascara) if b and (i == 0 or not mascara[i - 1]))


# ---------------------------------------------------------------------------------- envio
@dataclass
class Envio:
    pasta: Path  # <dados>/suporte/<data-hora>: textos e imagens, para o usuário conferir
    notas: list[ResultadoArquivo] = field(default_factory=list)
    titulares: dict[str, str] = field(default_factory=dict)  # CPF -> nome (dos dados informados)
    informados: list[DadosPessoais] = field(default_factory=list, repr=False)

    @property
    def aceitas(self) -> list[ResultadoArquivo]:
        return [n for n in self.notas if n.destino is not None]

    def titular(self, nota: ResultadoArquivo) -> str:
        """De quem é a nota, pelo CPF impresso (sem mostrar o CPF)."""
        nomes = sorted({self.titulares.get(c, "CPF sem dados para ocultar") for c in nota.cpfs})
        return ", ".join(nomes) or "nenhum CPF na nota"


def novo_envio(dados: Path, banco: Path, agora: datetime) -> Envio:
    """Opção "Enviar notas para suporte" (tela Suporte): uma pasta nova em ``<dados>/suporte``,
    que recebe nota a nota (``preparar_nota``) o texto anonimizado e as imagens tarjadas. A tela
    prepara uma nota por vez para mostrar o progresso."""
    pasta = dados / PASTA_SUPORTE / f"{agora:%Y%m%d-%H%M%S}"
    if pasta.exists():
        raise AnonimizacaoError("já existe um envio preparado neste segundo; tente de novo")
    informados = dados_para_ocultar(dados, banco)
    return Envio(pasta, [], {so_digitos(d.cpf): d.nome for d in informados}, informados)


def preparar_nota(
    envio: Envio, dados: Path, nome: str, conteudo: bytes, senhas: Sequence[str] = ()
) -> ResultadoArquivo:
    """Uma nota do envio, com nome neutro (nota-01.txt, nota-01-pagina-1.png...): o nome do PDF
    pode ter o nome do cliente. O PDF não fica guardado."""
    if Path(nome).suffix.lower() != ".pdf":
        raise AnonimizacaoError(f"não é PDF: {Path(nome).name}")
    neutro = f"nota-{len(envio.notas) + 1:02d}"
    with tempfile.TemporaryDirectory(prefix="coir-suporte-") as tmp:
        origem = Path(tmp) / f"{neutro}.pdf"
        origem.write_bytes(conteudo)
        [r] = anonimizar_arquivos(
            [origem], envio.pasta, dados / ARQUIVO_MAPA, termos_sensiveis(dados),
            [f"{neutro}.txt"], envio.informados, imagens=True, senhas=senhas,
        )  # fmt: skip
    nota = ResultadoArquivo(Path(nome).name, r.destino, r.recusas, r.imagens, r.paginas,
                            r.ocultados, r.alertas, r.cpfs)  # fmt: skip
    envio.notas.append(nota)
    return nota


def preparar_envio(
    dados: Path, banco: Path, recebidos: Sequence[tuple[str, bytes]], agora: datetime
) -> Envio:
    """Todas as notas de uma vez (``novo_envio`` e ``preparar_nota``). O .zip só sai depois que o
    usuário confere as imagens (``gerar_zip``)."""
    if fora := [n for n, _ in recebidos if Path(n).suffix.lower() != ".pdf"]:
        raise AnonimizacaoError(f"não é PDF: {Path(fora[0]).name}")
    envio = novo_envio(dados, banco, agora)
    for nome, conteudo in recebidos:
        preparar_nota(envio, dados, nome, conteudo)
    return envio


def gerar_zip(envio: Envio) -> Path:
    """O .zip com o texto e as imagens das notas aceitas, na própria pasta do envio."""
    if not envio.aceitas:
        raise AnonimizacaoError("nenhuma nota foi anonimizada; não há o que enviar")
    zip_ = envio.pasta / f"notas-para-suporte-{envio.pasta.name}.zip"
    with zipfile.ZipFile(zip_, "w", zipfile.ZIP_DEFLATED) as z:
        for nota in envio.aceitas:
            for arquivo in (nota.destino, *nota.imagens):
                if arquivo is not None:
                    z.write(arquivo, arquivo.name)
    return zip_


def endereco_email(zip_: Path, versao: str) -> str:
    """Mensagem para o suporte já endereçada (``mailto:``). Leva só a versão e o nome do .zip:
    nada pessoal (nem a pasta, que pode ter o nome do usuário do Windows) vai no endereço."""
    assunto = f"COIR {versao} - notas para suporte"
    corpo = (
        f"Versão do COIR: {versao}\n"
        "Corretora:\n"
        "O que acontece (a nota não é lida, um valor sai errado...):\n\n\n"
        f"(anexe o arquivo {zip_.name})\n"
    )
    return f"mailto:{EMAIL_SUPORTE}?subject={quote(assunto)}&body={quote(corpo)}"


# ------------------------------------------------------------------------- registro de auditoria
@dataclass(frozen=True)
class ExportacaoRegistro:
    pasta: Path  # <dados>/suporte/registro-<data-hora>: o arquivo legível e o .zip
    arquivo: Path
    zip: Path
    desde: date | None
    ate: date | None
    linhas: int
    retidas: dict[str, int]  # motivo -> quantas linhas ficaram de fora
    cortadas: int  # mais antigas que couberam a menos, pelo limite de tamanho
    termos_do_banco: bool = True  # falso: banco inacessível, sem o cadastro na conferência

    @property
    def total_retidas(self) -> int:
        return sum(self.retidas.values())


def _linhas_do_registro(dados: Path, desde: date) -> Iterable[str]:
    """As linhas do período, da mais antiga à mais recente. Sem a data no começo (linha estranha),
    a linha segue: quem decide é o portão."""
    limite = desde.isoformat()
    for arquivo in auditoria.arquivos_do_registro(dados):
        with arquivo.open(encoding="utf-8", errors="replace") as entrada:
            for linha in entrada:
                linha = linha.strip()
                data = RE_DATA_DA_LINHA.match(linha)
                if linha and (data is None or data[1] >= limite):
                    yield linha


def termos_do_titular(dados: Path, banco: Path) -> tuple[portao.Termos, bool]:
    """Os termos que o portão procura, de todos os titulares. Se o banco não abrir (é justamente
    quando se manda o registro), vale o que foi informado na tela e o portão continua valendo por
    forma e por catálogo; o cabeçalho da exportação registra isso."""
    do_banco_ok = True
    try:
        pessoais = dados_para_ocultar(dados, banco)
    except (CoirError, SQLAlchemyError):
        log.warning("portão do registro sem os dados do banco: banco inacessível")
        pessoais = carregar_dados(dados)
        do_banco_ok = False
    valores = [v for d in pessoais for v in (d.nome, d.cpf, *d.contas, *d.valores()) if v]
    segredo = dados / auditoria.ARQUIVO_SEGREDO
    segredos = [segredo.read_text(encoding="utf-8").strip()] if segredo.is_file() else []
    return portao.Termos.de(valores, termos_sensiveis(dados), segredos), do_banco_ok


def exportar_registro(
    dados: Path, banco: Path, agora: datetime, dias: int = DIAS_DO_REGISTRO
) -> ExportacaoRegistro:
    """Grava em ``<dados>/suporte`` o registro de auditoria dos últimos ``dias``, conferido pelo
    portão de privacidade (``coir.auditoria.portao``): o .jsonl, para o usuário olhar, e o .zip,
    para anexar ao e-mail. Nada sai do computador sem essa ação dele."""
    with auditoria.operacao("exportacao_registro", dias=dias) as op:
        exportado = _exportar_registro(dados, banco, agora, dias)
        op.dados(linhas=exportado.linhas, retidas=exportado.total_retidas,
                 cortadas=exportado.cortadas)  # fmt: skip
    return exportado


def _exportar_registro(dados: Path, banco: Path, agora: datetime, dias: int) -> ExportacaoRegistro:
    desde = agora.date() - timedelta(days=dias)
    termos, termos_do_banco = termos_do_titular(dados, banco)
    conferencia = portao.conferir_linhas(_linhas_do_registro(dados, desde), termos)
    if not conferencia.linhas:
        if conferencia.retidas:
            motivos = ", ".join(f"{m} ({q})" for m, q in sorted(conferencia.retidas.items()))
            raise AnonimizacaoError(
                f"nenhuma linha do registro passou na conferência de privacidade: {motivos}"
            )
        raise AnonimizacaoError("ainda não há registro de auditoria para enviar")
    linhas, cortadas = _dentro_do_limite(conferencia.linhas)
    datas = _datas(linhas)
    chave = auditoria.segredo() or bytes(32)
    quando = agora.astimezone()
    extras = [
        auditoria.linha_de_evento(
            "suporte.exportacao", chave, "00000000", 0, quando, versao=VERSAO, dias=dias,
            desde=min(datas), ate=max(datas),
            linhas=len(linhas), retidas=conferencia.retidas.total(), cortadas=cortadas,
            termos_do_banco=termos_do_banco,
        ),
        *(auditoria.linha_de_evento("portao.retidas", chave, "00000000", 0, quando, motivo=motivo,
                                    retidas=quantas)
          for motivo, quantas in sorted(conferencia.retidas.items())),
    ]  # fmt: skip
    # o cabeçalho passa pelo mesmo portão; se ele não passar, a exportação não sai
    cabecalho = portao.conferir_linhas(extras, termos)
    if cabecalho.retidas:
        raise AnonimizacaoError("o cabeçalho da exportação não passou no portão de privacidade")
    texto = "\n".join([*cabecalho.linhas, *linhas]) + "\n"
    portao.exigir_texto_limpo(texto, termos)
    return _gravar_exportacao(dados, agora, texto, conferencia, cortadas, linhas, termos_do_banco)


def _datas(linhas: Sequence[str]) -> list[date]:
    return [date.fromisoformat(json.loads(x)["t"][:10]) for x in linhas]


def _dentro_do_limite(linhas: list[str]) -> tuple[list[str], int]:
    total, inicio = 0, len(linhas)
    while inicio > 0 and total + len(linhas[inicio - 1].encode()) + 1 <= MAXIMO_DO_REGISTRO:
        inicio -= 1
        total += len(linhas[inicio].encode()) + 1
    return linhas[inicio:], inicio


def _gravar_exportacao(dados: Path, agora: datetime, texto: str, conferencia: portao.Conferencia,
                       cortadas: int, linhas: list[str],
                       termos_do_banco: bool) -> ExportacaoRegistro:  # fmt: skip
    pasta = dados / PASTA_SUPORTE / f"registro-{agora:%Y%m%d-%H%M%S}"
    if pasta.exists():
        raise AnonimizacaoError("já existe uma exportação neste segundo; tente de novo")
    pasta.mkdir(parents=True)
    arquivo = pasta / ARQUIVO_DO_REGISTRO
    arquivo.write_text(texto, encoding="utf-8", newline="\n")
    zip_ = pasta / f"registro-para-suporte-{agora:%Y%m%d-%H%M%S}.zip"
    with zipfile.ZipFile(zip_, "w", zipfile.ZIP_DEFLATED) as z:
        z.write(arquivo, arquivo.name)
    with zipfile.ZipFile(zip_) as z:  # o .zip tem de dar exatamente o texto conferido
        conferido = z.read(arquivo.name).decode("utf-8")
    if conferido != texto or arquivo.read_text(encoding="utf-8") != texto:
        zip_.unlink()
        arquivo.unlink()
        raise AnonimizacaoError("cópia gravada do registro diferente da conferida")
    log.info("registro de auditoria exportado: %d linha(s)", len(linhas))
    return ExportacaoRegistro(
        pasta, arquivo, zip_, min(_datas(linhas)), max(_datas(linhas)), len(linhas),
        dict(conferencia.retidas), cortadas, termos_do_banco,
    )  # fmt: skip


def endereco_email_registro(exportado: ExportacaoRegistro, versao: str) -> str:
    """Como ``endereco_email``: leva só a versão e o nome do .zip."""
    assunto = f"COIR {versao} - registro para suporte"
    corpo = (
        f"Versão do COIR: {versao}\n"
        "O que aconteceu e em que tela:\n"
        "Quando (dia e hora aproximados):\n\n\n"
        f"(anexe o arquivo {exportado.zip.name})\n"
    )
    return f"mailto:{EMAIL_SUPORTE}?subject={quote(assunto)}&body={quote(corpo)}"
