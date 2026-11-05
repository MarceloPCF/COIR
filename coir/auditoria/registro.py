"""Gravação do registro de auditoria sobre o ``logging`` da biblioteca padrão.

- ``evento``: um fato com campos do catálogo. O que não couber é recusado (o nome do campo vai para
  o log comum, para corrigir o catálogo); nunca derruba o programa.
- ``operacao``: início e fim de uma etapa (também como decorador), com o tempo gasto e, se falhar,
  a classe do erro, a pilha (módulo, função e linha) e o texto limpo. Os eventos de dentro levam o
  número da operação, para o suporte ler junto o que aconteceu junto.
- A ponte com o ``logging`` traz para cá o que os módulos já registram (informações do COIR; avisos
  e erros de qualquer biblioteca), com o texto limpo.
"""

import hashlib
import hmac
import itertools
import json
import logging
import os
import platform
import secrets
import sys
import threading
from collections import Counter
from collections.abc import Callable, Iterator
from contextlib import contextmanager, suppress
from contextvars import ContextVar
from datetime import datetime
from logging.handlers import RotatingFileHandler
from pathlib import Path
from time import perf_counter
from types import TracebackType
from typing import Any

from coir.auditoria.catalogo import CAMPOS, EVENTOS, PREFIXOS
from coir.auditoria.limpeza import limpar

log = logging.getLogger(__name__)

ARQUIVO = Path("logs") / "auditoria.jsonl"
ARQUIVO_SEGREDO = Path("_privado") / "segredo-registro.txt"
TAMANHO_MAXIMO = 2_000_000
ARQUIVOS_ANTIGOS = 5
TAMANHO_SEGREDO = 32
ADIAR_ROTACAO = 60.0  # segundos até tentar girar de novo um arquivo preso por outro processo
QUADROS_NA_PILHA = 25
REPETICOES_DA_PONTE = 40  # a mesma mensagem por sessão (biblioteca que avisa a cada página)
NIVEIS = {
    logging.DEBUG: "I", logging.INFO: "I", logging.WARNING: "A", logging.ERROR: "E",
    logging.CRITICAL: "C",
}  # fmt: skip
_op_atual: ContextVar[str | None] = ContextVar("coir_auditoria_op", default=None)
_FALHAS_PREVISTAS = (ValueError, TypeError, KeyError, ArithmeticError, OSError)
_Gancho = Callable[[type[BaseException], BaseException, TracebackType | None], object]
_GanchoThread = Callable[[threading.ExceptHookArgs], object]


def _carregar_segredo(dados: Path) -> bytes:
    """O segredo dos pseudônimos: sem ele não se chega ao CPF ou à conta por força bruta. Fica em
    ``_privado`` e nunca vai no envio. Se não der para gravar, vale só nesta sessão."""
    caminho = dados / ARQUIVO_SEGREDO
    try:
        if caminho.is_file():
            gravado = caminho.read_text(encoding="utf-8").strip()
            with suppress(ValueError):  # ilegível: um novo, como o curto
                if len(bytes.fromhex(gravado)) >= TAMANHO_SEGREDO:
                    return bytes.fromhex(gravado)
            log.warning("segredo do registro de auditoria inválido: um novo foi criado")
        novo = secrets.token_hex(TAMANHO_SEGREDO)
        caminho.parent.mkdir(parents=True, exist_ok=True)
        # de uma vez: uma queda no meio não pode deixar um segredo vazio ou curto
        temporario = caminho.with_suffix(".tmp")
        temporario.write_text(novo, encoding="utf-8")
        temporario.replace(caminho)
        return bytes.fromhex(novo)
    except OSError:
        log.warning("segredo do registro de auditoria indisponível: vale só nesta sessão")
        return secrets.token_bytes(TAMANHO_SEGREDO)


class RotacaoTolerante(RotatingFileHandler):
    """No Windows, a rotação falha se outro processo (a janela e um comando da linha de comando)
    está com o arquivo aberto. A padrão tentaria de novo a cada linha, perdendo a linha e empurrando
    arquivos antigos para fora. Aqui, se o arquivo não pode ser renomeado, nada é mexido, a escrita
    continua e a próxima tentativa fica para depois."""

    def __init__(self, *args: Any, **kwargs: Any) -> None:
        super().__init__(*args, **kwargs)
        self._adiar_ate = 0.0

    def shouldRollover(self, record: logging.LogRecord) -> bool:
        return perf_counter() >= self._adiar_ate and bool(super().shouldRollover(record))

    def doRollover(self) -> None:
        if self.stream:
            self.stream.close()
            self.stream = None  # type: ignore[assignment]
        teste = f"{self.baseFilename}.rotacao"
        try:
            os.replace(self.baseFilename, teste)
            os.replace(teste, self.baseFilename)
        except FileNotFoundError:
            pass  # ainda não há arquivo: nada a girar
        except OSError:
            self._adiar_ate = perf_counter() + ADIAR_ROTACAO
            self.stream = self._open()
            return
        super().doRollover()


class _Alvo:
    """O registro de uma pasta de dados, enquanto estiver ativo."""

    def __init__(self, dados: Path) -> None:
        self.dados = dados
        self.segredo = _carregar_segredo(dados)
        self.sessao = secrets.token_hex(4)
        self.sequencia = itertools.count(1)
        self.vistos: set[object] = set()
        self.repeticoes: Counter[tuple[str, str]] = Counter()
        arquivo = dados / ARQUIVO
        arquivo.parent.mkdir(parents=True, exist_ok=True)
        self.manipulador = RotacaoTolerante(
            arquivo, maxBytes=TAMANHO_MAXIMO, backupCount=ARQUIVOS_ANTIGOS, encoding="utf-8",
            delay=True,
        )  # fmt: skip
        self.manipulador.setFormatter(logging.Formatter("%(message)s"))


class _Ponte(logging.Handler):
    """Traz os registros do ``logging`` para a auditoria: as informações do próprio COIR e os
    avisos e erros de qualquer biblioteca (a tela que quebra chega aqui pelo log do NiceGUI)."""

    def __init__(self) -> None:
        super().__init__(logging.INFO)

    def emit(self, record: logging.LogRecord) -> None:
        alvo = _alvo
        if alvo is None or record.name.startswith("coir.auditoria"):
            return
        if record.levelno < logging.WARNING and not record.name.startswith("coir"):
            return
        chave = (record.name, str(record.msg))
        alvo.repeticoes[chave] += 1
        if alvo.repeticoes[chave] > REPETICOES_DA_PONTE:
            return
        try:
            campos: dict[str, object] = {"modulo": record.name, "funcao": record.funcName,
                                         "linha": record.lineno,
                                         "texto": record.getMessage()}  # fmt: skip
            if record.exc_info and record.exc_info[1] is not None:
                erro = dados_do_erro(record.exc_info[1])
                erro["texto"] = f"{campos['texto']}: {erro['texto']}"
                campos |= erro
            evento("log.registro", NIVEIS.get(record.levelno, "E"), **campos)
        except _FALHAS_PREVISTAS:
            self.handleError(record)


_alvo: _Alvo | None = None
_ponte = _Ponte()
_saida = logging.getLogger("coir.auditoria.saida")
_saida.propagate = False  # o JSON não vai para o console nem para o coir.log
_saida.setLevel(logging.INFO)
_ganchos: tuple[_Gancho, _GanchoThread, _Gancho, _GanchoThread] | None = None
_trava = threading.Lock()


def ativa() -> bool:
    """Para não montar dados à toa quando não há registro (testes, comando sem pasta de dados)."""
    return _alvo is not None


def configurar(dados: Path, versao: str, modo: str = "janela") -> None:
    """Liga o registro em ``<dados>/logs``. A mesma pasta de novo não faz nada; outra (o usuário
    escolheu a pasta no primeiro uso) troca o destino."""
    global _alvo
    with _trava:
        if _alvo is not None and _alvo.dados == dados:
            return
        _desligar()
        alvo = _Alvo(dados)
        _saida.addHandler(alvo.manipulador)
        _alvo = alvo
        logging.getLogger().addHandler(_ponte)
        # só o COIR: abaixar o nível da raiz faria as bibliotecas escreverem mais nos outros logs
        coir = logging.getLogger("coir")
        if coir.getEffectiveLevel() > logging.INFO:
            coir.setLevel(logging.INFO)
        _instalar_ganchos()
    evento(
        "sessao.inicio", versao=versao, python=platform.python_version(),
        sistema=platform.platform(terse=True), executavel=bool(getattr(sys, "frozen", False)),
        modo=modo,
    )  # fmt: skip


def desligar() -> None:
    with _trava:
        _desligar()


def _desligar() -> None:
    global _alvo
    alvo, _alvo = _alvo, None
    if alvo is None:
        return
    logging.getLogger().removeHandler(_ponte)
    _saida.removeHandler(alvo.manipulador)
    alvo.manipulador.close()
    _tirar_ganchos()


def _instalar_ganchos() -> None:
    """Erro que ninguém tratou (o programa ia fechar calado): fica registrado antes."""
    global _ganchos
    principal, de_thread = sys.excepthook, threading.excepthook

    def gancho(tipo: type[BaseException], erro: BaseException, pilha: TracebackType | None) -> None:
        _nao_tratado(erro)
        principal(tipo, erro, pilha)

    def gancho_thread(args: threading.ExceptHookArgs) -> None:
        if args.exc_value is not None:
            _nao_tratado(args.exc_value)
        de_thread(args)

    _ganchos = (principal, de_thread, gancho, gancho_thread)
    sys.excepthook, threading.excepthook = gancho, gancho_thread


def _nao_tratado(erro: BaseException) -> None:
    """Direto no registro: pelo ``logging``, o erro sairia duas vezes no console da linha de
    comando. Fechar o programa ou pedir para parar não é erro."""
    if not isinstance(erro, SystemExit | KeyboardInterrupt):
        evento("log.registro", "C", modulo="coir", funcao="excepthook", **dados_do_erro(erro))


def _tirar_ganchos() -> None:
    """Só desfaz o que ainda é nosso: outra biblioteca pode ter posto um gancho depois."""
    global _ganchos
    if _ganchos is None:
        return
    principal, de_thread, gancho, gancho_thread = _ganchos
    if sys.excepthook is gancho:
        sys.excepthook = principal
    if threading.excepthook is gancho_thread:
        threading.excepthook = de_thread
    _ganchos = None


def pseudonimo(campo: str, valor: object, chave: bytes) -> str:
    """Sempre o mesmo para o mesmo valor (dá para seguir um arquivo no log); não se reverte."""
    resumo = hmac.new(chave, f"{campo}|{valor}".encode(), hashlib.sha256).hexdigest()
    return f"{PREFIXOS[campo]}_{resumo[:10]}"


def segredo() -> bytes | None:
    return None if _alvo is None else _alvo.segredo


def uma_vez(*chave: object) -> bool:
    """Verdadeiro só na primeira vez nesta sessão: as telas recalculam a toda hora, e o que não
    mudou não precisa aparecer de novo no registro."""
    alvo = _alvo
    if alvo is None or chave in alvo.vistos:
        return False
    alvo.vistos.add(chave)
    return True


def evento(nome: str, nivel: str = "I", /, **campos: object) -> None:
    """Registra um evento do catálogo. Sem registro ativo, não faz nada."""
    alvo = _alvo
    if alvo is None:
        return
    try:
        _saida.info(
            _texto(
                datetime.now().astimezone(), alvo.sessao, next(alvo.sequencia), nivel, nome,
                _op_atual.get(), _dados(nome, campos, alvo.segredo),
            )
        )  # fmt: skip
    except _FALHAS_PREVISTAS:  # o registro nunca derruba o programa; o erro fica no log comum
        log.exception("auditoria: evento %s não registrado", nome)


def linha_de_evento(nome: str, chave: bytes, sessao: str, n: int, quando: datetime,
                    **campos: object) -> str:  # fmt: skip
    """A linha de um evento fora do registro (o cabeçalho da exportação)."""
    return _texto(quando, sessao, n, "I", nome, None, _dados(nome, campos, chave))


def _texto(quando: datetime, sessao: str, n: int, nivel: str, nome: str, op: str | None,
           dados: dict[str, object]) -> str:  # fmt: skip
    conteudo: dict[str, object] = {"t": quando.isoformat(timespec="milliseconds"), "s": sessao,
                                   "n": n, "nv": nivel, "ev": nome}  # fmt: skip
    if op:
        conteudo["op"] = op
    conteudo["d"] = dados
    return json.dumps(conteudo, ensure_ascii=False, separators=(",", ":"))


def _dados(nome: str, campos: dict[str, object], chave: bytes) -> dict[str, object]:
    permitidos = EVENTOS.get(nome, frozenset[str]())
    dados: dict[str, object] = {}
    recusados: list[str] = []
    for campo, valor in campos.items():
        if valor is None:
            continue
        if campo not in permitidos:
            recusados.append(campo)
            continue
        try:
            dados[campo] = (pseudonimo(campo, valor, chave) if campo in PREFIXOS
                            else limpar(CAMPOS[campo], valor))  # fmt: skip
        except ValueError:
            recusados.append(campo)
    if recusados:  # só o nome do campo, nunca o valor
        log.warning("auditoria: o evento %s recusou %s", nome, sorted(recusados))
    return dados


class Operacao:
    """O que a etapa quer registrar no fim (``op.dados(notas=3)``)."""

    def __init__(self) -> None:
        self.campos: dict[str, object] = {}

    def dados(self, **campos: object) -> None:
        self.campos |= campos


@contextmanager
def operacao(nome: str, /, **campos: object) -> Iterator[Operacao]:
    """Início e fim de uma etapa, com o tempo gasto; a falha é registrada e continua subindo."""
    op = Operacao()
    token = _op_atual.set(secrets.token_hex(3) if ativa() else None)
    inicio = perf_counter()

    def gasto() -> int:
        return int((perf_counter() - inicio) * 1000)

    try:
        evento(f"{nome}.inicio", **campos)
        yield op
    except Exception as erro:
        nivel = "A" if _erro_de_dominio(erro) else "E"
        evento(
            f"{nome}.fim", nivel, **{**op.campos, **dados_do_erro(erro), "ok": False, "ms": gasto()}
        )
        raise
    else:
        evento(f"{nome}.fim", ok=True, ms=gasto(), **op.campos)
    finally:
        _op_atual.reset(token)


def erro_de_dominio(erro: BaseException) -> None:
    """Um erro do COIR acabou de ser criado: quase todos terminam numa mensagem na tela e em mais
    nenhum lugar. Fica aqui de onde ele veio; a pilha ainda não existe nesse momento."""
    if _alvo is None:
        return
    quadro = sys._getframe(1)  # pyright: ignore[reportPrivateUsage]
    while quadro.f_back is not None and quadro.f_globals.get("__name__") == "coir.erros":
        quadro = quadro.f_back
    evento(
        "erro.dominio", "A", classe=type(erro).__name__, texto=_texto_do_erro(erro),
        modulo=str(quadro.f_globals.get("__name__", "desconhecido")),
        funcao=quadro.f_code.co_qualname, linha=quadro.f_lineno,
    )  # fmt: skip


def _erro_de_dominio(erro: BaseException) -> bool:
    return any(c.__name__ == "CoirError" for c in type(erro).__mro__)


def dados_do_erro(erro: BaseException) -> dict[str, object]:
    """Classe, texto limpo, pilha (sem caminhos: módulo, função e linha) e a cadeia de causas."""
    dados: dict[str, object] = {"classe": type(erro).__name__, "texto": _texto_do_erro(erro),
                                "pilha": _pilha(erro), "causas": _causas(erro)}  # fmt: skip
    data = getattr(erro, "data", None)  # MotorError: o dia do livro em que o motor parou
    if data is not None:
        dados["data"] = data
    return dados


def _texto_do_erro(erro: BaseException) -> str:
    """Um ``__str__`` com defeito não pode esconder o erro de verdade nem impedir a criação dele."""
    try:
        return str(erro)
    except Exception:
        log.exception("auditoria: texto do erro %s indisponível", type(erro).__name__)
        return "texto indisponível"


def _pilha(erro: BaseException) -> list[str]:
    quadros: list[str] = []
    atual = erro.__traceback__
    while atual is not None:
        quadro = atual.tb_frame
        modulo = str(quadro.f_globals.get("__name__", "desconhecido"))
        quadros.append(f"{modulo}:{quadro.f_code.co_qualname}:{atual.tb_lineno}")
        atual = atual.tb_next
    return quadros[-QUADROS_NA_PILHA:]


def _causas(erro: BaseException) -> list[str]:
    classes: list[str] = []
    atual = erro.__cause__ or erro.__context__
    while atual is not None and len(classes) < 10:
        classes.append(type(atual).__name__)
        atual = atual.__cause__ or atual.__context__
    return classes


def arquivos_do_registro(dados: Path) -> list[Path]:
    """Do mais antigo ao mais recente."""
    arquivo = dados / ARQUIVO
    antigos = [arquivo.with_name(f"{arquivo.name}.{n}") for n in range(ARQUIVOS_ANTIGOS, 0, -1)]
    return [a for a in (*antigos, arquivo) if a.is_file()]
