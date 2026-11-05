"""Registro de auditoria: o que o programa fez, para o suporte corrigir um problema só com ele.

``<dados>/logs/auditoria.jsonl``, um JSON por linha: ``t`` hora, ``s`` sessão, ``n`` sequência,
``nv`` nível (I, A, E), ``ev`` evento, ``d`` campos. Só entram os eventos de ``EVENTOS`` e os campos
de ``CAMPOS``, cada um numa forma fechada; **não há campo de texto livre**: de um erro ficam a
classe e o ponto do código, nunca a mensagem. Identificador vira pseudônimo (HMAC com segredo local)
e valor vira ordem de grandeza. Vale já no arquivo local, e a exportação confere tudo de novo
(``conferir_linhas`` e ``achados_no_texto``), sem confiar no arquivo. Para registrar algo novo,
acrescente o campo aqui, na forma mais restrita possível; nunca valor exato, nome ou caminho.

Só usa a biblioteca padrão, para ``coir.erros`` poder usar. O motor não registra nada (regra 6): o
serviço registra a partir do resultado (``coir.servicos.trilha``).
"""

import hashlib
import hmac
import itertools
import json
import logging
import platform
import re
import secrets
import sys
import threading
import unicodedata
from collections.abc import Iterable, Iterator
from contextlib import contextmanager, suppress
from dataclasses import dataclass
from datetime import date, datetime
from decimal import ROUND_HALF_EVEN, Decimal, InvalidOperation
from logging.handlers import RotatingFileHandler
from pathlib import Path
from time import perf_counter

log = logging.getLogger(__name__)

ARQUIVO = Path("logs") / "auditoria.jsonl"
ARQUIVO_SEGREDO = Path("_privado") / "segredo-registro.txt"
ARQUIVOS_ANTIGOS = 5
LIMITE_DIFERENCA = Decimal(100)
LIMITE_INTEIRO = 10**10  # 11 dígitos já teria cara de CPF
LIMITE_LISTA = 50

PREFIXOS = {"arquivo": "arq", "nota": "not", "conta": "cta", "titular": "tit", "assinatura": "res"}
RE_PSEUDONIMO = re.compile(rf"(?:{'|'.join(PREFIXOS.values())})_[0-9a-f]{{10}}")
RE_DATA = re.compile(r"\d{4}-\d{2}-\d{2}")
RE_MES = re.compile(r"\d{4}-\d{2}")
RE_GRANDEZA = re.compile(r"0|[+-]1e\d{1,2}")
RE_DIFERENCA = re.compile(r"[+-]\d{1,3}\.\d{2}|0|[+-]1e\d{1,2}")
# código de negociação: raiz de 4 e 1 ou 2 dígitos (PETR4, B3SA3, A1AP34, WINZ24), ou raiz, letra da
# série e até 3 dígitos (PETRK300E); conta ou CPF com letra na frente não têm essa forma
RE_ATIVO = re.compile(r"[A-Z0-9]{3}[A-Z]\d{1,2}F?|[A-Z0-9]{3}[A-Z][A-X]\d{1,3}(?:E|W\d)?")
# valor de uma lista do programa (situação, tipo, corretora): sem @, barra nem número longo
RE_ROTULO = re.compile(r"(?!.*\d{4})(?!.*\d[.\- ]\d)[^\W\d_][\w .&\-]{0,39}")
RE_CODIGO = re.compile(r"[A-Za-z_<][\w.<>\-]{0,99}")
RE_QUADRO = re.compile(r"[A-Za-z_][\w.\-]*:[\w.<>]+:\d+")  # modulo:funcao:linha
RE_VERSAO = re.compile(r"\d{1,2}(?:\.\d{1,3}){1,2}[a-z0-9.+]{0,10}")
RE_ROTA = re.compile(r"/[\w/\-]{0,60}")


def _campos(forma: object, nomes: str) -> dict[str, object]:
    return dict.fromkeys(nomes.split(), forma)


CAMPOS: dict[str, object] = {
    **_campos(RE_PSEUDONIMO, " ".join(PREFIXOS)),
    **_campos(RE_DATA, "data desde ate"),
    "mes": RE_MES,
    **_campos(
        int,
        "avisos dias divergentes gravadas ignoradas isolados ja_no_banco lancadas lidos linha "
        "linhas meses ms negocios novas operacoes pendentes rejeitadas renormalizadas retidas "
        "tamanho total",
    ),
    **_campos(
        bool,
        "alterado com_juros com_multa com_pendencia executavel isento ok provisorio "
        "termos_do_banco",
    ),
    **_campos(
        RE_GRANDEZA,
        "darf darf_acumulado imposto irrf prejuizo quantidade resultado_comum "
        "resultado_day_trade resultado_fii valor vendas_acoes",
    ),
    "diferenca": RE_DIFERENCA,
    **_campos(RE_ATIVO, "ativo ativo_destino"),
    **_campos(RE_ROTULO, "campo categoria corretora origem segmento sistema situacao tipo"),
    **_campos(RE_CODIGO, "classe funcao modulo nome"),
    **_campos(RE_VERSAO, "python versao"),
    "rota": RE_ROTA,
    "paginas": [int],
    "pilha": [RE_QUADRO],
    "causas": [RE_CODIGO],
    **_campos([RE_ATIVO], "incompletos dependentes"),
}
EVENTOS = frozenset(
    [
        "sessao.inicio",
        "ui.tela",
        "log.registro",
        "erro.dominio",
        "operacao.inicio",
        "operacao.fim",
        "importacao.arquivo",
        "importacao.nota",
        "importacao.rejeicao",
        "importacao.resumo",
        "b3.planilha",
        "apuracao.resumo",
        "apuracao.mes",
        "apuracao.isolado",
        "conciliacao.resumo",
        "conciliacao.divergencia",
        "propostas.resumo",
        "eventos.proposta",
        "eventos.salvo",
        "eventos.excluido",
        "darf.linha",
        "suporte.exportacao",
    ]
)


def _na_forma(forma: object, valor: object) -> bool:
    if isinstance(forma, list):
        return (isinstance(valor, list) and len(valor) <= LIMITE_LISTA
                and all(_na_forma(forma[0], v) for v in valor))  # fmt: skip
    if forma is int:
        return type(valor) is int and abs(valor) < LIMITE_INTEIRO
    if forma is bool:
        return type(valor) is bool
    return isinstance(valor, str) and forma.fullmatch(valor) is not None  # type: ignore[attr-defined]


def grandeza(valor: object) -> str:
    """Sinal e ordem de grandeza, nunca o valor (1.234,56 vira +1e3)."""
    numero = _decimal(valor)
    return f"{'-' if numero < 0 else '+'}1e{max(numero.adjusted(), 0)}" if numero else "0"


def diferenca(valor: object) -> str:
    """Até 100 é o que explica uma nota que não fecha e não diz nada do cliente."""
    numero = _decimal(valor)
    if abs(numero) > LIMITE_DIFERENCA:
        return grandeza(numero)
    return f"{numero.quantize(Decimal('0.01'), rounding=ROUND_HALF_EVEN):+.2f}"


def _decimal(valor: object) -> Decimal:
    if isinstance(valor, bool) or not isinstance(valor, Decimal | int | str):
        raise ValueError("valor sem número")
    try:
        numero = Decimal(str(valor))
    except InvalidOperation as erro:
        raise ValueError("valor sem número") from erro
    if not numero.is_finite():
        raise ValueError("valor sem número")
    return numero


def pseudonimo(campo: str, valor: object, chave: bytes) -> str:
    """Sempre o mesmo para o mesmo valor (dá para seguir um arquivo no log); não se reverte."""
    resumo = hmac.new(chave, f"{campo}|{valor}".encode(), hashlib.sha256).hexdigest()
    return f"{PREFIXOS[campo]}_{resumo[:10]}"


def _convertido(campo: str, valor: object, chave: bytes) -> object:
    forma = CAMPOS[campo]
    if campo in PREFIXOS:
        return pseudonimo(campo, valor, chave)
    if forma is RE_GRANDEZA:
        return grandeza(valor)
    if forma is RE_DIFERENCA:
        return diferenca(valor)
    if isinstance(valor, date):
        return valor.isoformat()[: 7 if forma is RE_MES else 10]
    if isinstance(valor, tuple | list):
        return list(valor)  # pyright: ignore[reportUnknownArgumentType, reportUnknownVariableType]
    return valor


def _dados(campos: dict[str, object], chave: bytes) -> dict[str, object]:
    dados: dict[str, object] = {}
    recusados: list[str] = []
    for campo, valor in campos.items():
        if valor is None:
            continue
        try:
            convertido = _convertido(campo, valor, chave) if campo in CAMPOS else None
        except ValueError:
            convertido = None
        if campo in CAMPOS and _na_forma(CAMPOS[campo], convertido):
            dados[campo] = convertido
        else:
            recusados.append(campo)
    if recusados:  # só o nome do campo, nunca o valor
        log.warning("auditoria: campos recusados: %s", sorted(recusados))
    return dados


def linha(nome: str, nivel: str, n: int, sessao: str, chave: bytes, /, **campos: object) -> str:
    """Uma linha do registro, com hora de agora."""
    return json.dumps(
        {"t": datetime.now().astimezone().isoformat(timespec="milliseconds"), "s": sessao, "n": n,
         "nv": nivel, "ev": nome, "d": _dados(campos, chave)},
        ensure_ascii=False, separators=(",", ":"),
    )  # fmt: skip


# ------------------------------------------------------------------------------------- gravação
def _carregar_segredo(dados: Path) -> bytes:
    """O segredo dos pseudônimos: sem ele não se chega ao CPF ou à conta por força bruta. Fica em
    ``_privado`` e nunca vai no envio. Se não der para gravar, vale só nesta sessão."""
    caminho = dados / ARQUIVO_SEGREDO
    try:
        if caminho.is_file():
            with suppress(ValueError):  # curto ou ilegível: um novo
                if len(gravado := bytes.fromhex(caminho.read_text(encoding="utf-8").strip())) >= 32:
                    return gravado
        novo = secrets.token_hex(32)
        caminho.parent.mkdir(parents=True, exist_ok=True)
        temporario = caminho.with_suffix(".tmp")  # de uma vez: uma queda não deixa um segredo curto
        temporario.write_text(novo, encoding="utf-8")
        temporario.replace(caminho)
        return bytes.fromhex(novo)
    except OSError:
        log.warning("segredo do registro de auditoria indisponível: vale só nesta sessão")
        return secrets.token_bytes(32)


class _Alvo:
    """O registro de uma pasta de dados, enquanto estiver ativo."""

    def __init__(self, dados: Path) -> None:
        self.dados = dados
        self.segredo = _carregar_segredo(dados)
        self.sessao = secrets.token_hex(4)
        self.sequencia = itertools.count(1)
        self.vistos: set[object] = set()
        (dados / ARQUIVO).parent.mkdir(parents=True, exist_ok=True)
        self.manipulador = RotatingFileHandler(
            dados / ARQUIVO, maxBytes=2_000_000, backupCount=ARQUIVOS_ANTIGOS, encoding="utf-8",
            delay=True,
        )  # fmt: skip
        self.manipulador.setFormatter(logging.Formatter("%(message)s"))


class _Ponte(logging.Handler):
    """Avisos e erros de qualquer módulo (a tela que quebra chega pelo log do NiceGUI): de onde
    vieram, sem a mensagem. Uma vez por sessão cada ponto do código."""

    def emit(self, record: logging.LogRecord) -> None:
        if record.name.startswith("coir.auditoria"):
            return
        erro = record.exc_info[1] if record.exc_info else None
        if uma_vez("ponte", record.name, record.funcName, record.lineno, type(erro)):
            nivel = "A" if record.levelno < logging.ERROR else "E"
            evento("log.registro", nivel, modulo=record.name, funcao=record.funcName,
                   linha=record.lineno, **(dados_do_erro(erro) if erro else {}))  # fmt: skip


_alvo: _Alvo | None = None
_ponte = _Ponte(logging.WARNING)
_saida = logging.getLogger("coir.auditoria.saida")
_saida.propagate = False  # o JSON não vai para o console nem para o coir.log
_saida.setLevel(logging.INFO)
_trava = threading.Lock()


def ativa() -> bool:
    """Para não montar dados à toa quando não há registro (linha de comando, testes)."""
    return _alvo is not None


def configurar(dados: Path, versao: str) -> None:
    """Liga o registro em ``<dados>/logs``. Outra pasta (do primeiro uso) troca o destino."""
    global _alvo
    with _trava:
        if _alvo is not None and _alvo.dados == dados:
            return
        _desligar()
        _alvo = _Alvo(dados)
        _saida.addHandler(_alvo.manipulador)
        logging.getLogger().addHandler(_ponte)
    evento("sessao.inicio", versao=versao, python=platform.python_version(),
           sistema=platform.platform(terse=True),
           executavel=bool(getattr(sys, "frozen", False)))  # fmt: skip


def desligar() -> None:
    with _trava:
        _desligar()


def _desligar() -> None:
    global _alvo
    if _alvo is not None:
        logging.getLogger().removeHandler(_ponte)
        _saida.removeHandler(_alvo.manipulador)
        _alvo.manipulador.close()
        _alvo = None


def segredo() -> bytes:
    return _alvo.segredo if _alvo is not None else bytes(32)


def uma_vez(*chave: object) -> bool:
    """Verdadeiro só na primeira vez nesta sessão: as telas recalculam a toda hora, e o que não
    mudou não precisa aparecer de novo."""
    if _alvo is None or chave in _alvo.vistos:
        return False
    _alvo.vistos.add(chave)
    return True


def evento(nome: str, nivel: str = "I", /, **campos: object) -> None:
    """Registra um evento; o que não couber na lista fechada fica de fora. Sem registro, nada."""
    alvo = _alvo
    if alvo is not None:
        _saida.info(linha(nome, nivel, next(alvo.sequencia), alvo.sessao, alvo.segredo, **campos))


@contextmanager
def operacao(nome: str, /) -> Iterator[None]:
    """Início e fim de uma etapa (também como decorador), com o tempo gasto; a falha é registrada
    e continua subindo."""
    inicio = perf_counter()
    evento("operacao.inicio", nome=nome)
    try:
        yield
    except Exception as erro:
        evento("operacao.fim", "E", nome=nome, ok=False, ms=int((perf_counter() - inicio) * 1000),
               **dados_do_erro(erro))  # fmt: skip
        raise
    evento("operacao.fim", nome=nome, ok=True, ms=int((perf_counter() - inicio) * 1000))


def erro_de_dominio(erro: BaseException) -> None:
    """Um erro do COIR acabou de ser criado: quase todos terminam numa mensagem na tela e em mais
    nenhum lugar. Fica de onde ele veio."""
    if _alvo is None:
        return
    quadro = sys._getframe(1)  # pyright: ignore[reportPrivateUsage]
    while quadro.f_back is not None and quadro.f_globals.get("__name__") == "coir.erros":
        quadro = quadro.f_back
    evento("erro.dominio", "A", classe=type(erro).__name__,
           modulo=str(quadro.f_globals.get("__name__", "desconhecido")),
           funcao=quadro.f_code.co_qualname, linha=quadro.f_lineno)  # fmt: skip


def dados_do_erro(erro: BaseException) -> dict[str, object]:
    """Classe, pilha (módulo, função e linha, sem caminhos) e a cadeia de causas; nunca o texto."""
    pilha: list[str] = []
    atual = erro.__traceback__
    while atual is not None:
        quadro = atual.tb_frame
        modulo = str(quadro.f_globals.get("__name__", "desconhecido"))
        pilha.append(f"{modulo}:{quadro.f_code.co_qualname}:{atual.tb_lineno}")
        atual = atual.tb_next
    causas: list[str] = []
    causa = erro.__cause__ or erro.__context__
    while causa is not None and len(causas) < 10:
        causas.append(type(causa).__name__)
        causa = causa.__cause__ or causa.__context__
    dados: dict[str, object] = {
        "classe": type(erro).__name__,
        "pilha": pilha[-25:],
        "causas": causas,
    }
    if isinstance(data := getattr(erro, "data", None), date):  # MotorError: onde o motor parou
        dados["data"] = data
    return dados


def arquivos_do_registro(dados: Path) -> list[Path]:
    """Do mais antigo ao mais recente."""
    arquivo = dados / ARQUIVO
    antigos = [arquivo.with_name(f"{arquivo.name}.{n}") for n in range(ARQUIVOS_ANTIGOS, 0, -1)]
    return [a for a in (*antigos, arquivo) if a.is_file()]


# ------------------------------------------------------------------------- portão de privacidade
RE_HORA = re.compile(r"\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}[+-]\d{2}:\d{2}")
RE_SESSAO = re.compile(r"[0-9a-f]{8}")
RE_PALAVRA = re.compile(r"[^\W\d_]+")
RE_NUMERO = re.compile(r"\d[\d.\-/ ]*\d")
# a rede da segunda passagem: nada disso cabe nas formas acima
RE_PESSOAL = re.compile(
    r"(?<![\d.])\d{3}[.\s]?\d{3}[.\s]?\d{3}[-.\s]?\d{2}(?!\d)"  # CPF
    r"|[\w.+-]+@[\w-]+\.[\w-]+"  # e-mail
    r"|(?<![\d.,/-])\d{4,5}-\d{4}(?![\d-])"  # telefone
    r"|(?<![\d.,/-])\d{5}-\d{3}(?![\d-])"  # CEP
    r"|[A-Za-z]:\\|\\Users\\|/home/|/Users/"  # pasta
)
_TEXTUAIS = (RE_ROTULO, RE_CODIGO, RE_ATIVO, RE_QUADRO, RE_ROTA, RE_VERSAO)


def _normal(texto: str) -> str:
    sem = "".join(c for c in unicodedata.normalize("NFKD", texto) if not unicodedata.combining(c))
    return sem.casefold()


@dataclass(frozen=True)
class Termos:
    """O que identifica o titular, na forma de busca. Em memória; nunca vai ao arquivo."""

    palavras: frozenset[str] = frozenset()
    numeros: tuple[re.Pattern[str], ...] = ()
    sensiveis: tuple[str, ...] = ()
    segredo: str = ""

    def __repr__(self) -> str:
        return f"Termos({len(self.palavras)} palavra(s), {len(self.numeros)} número(s))"

    @classmethod
    def de(cls, valores: Iterable[str], sensiveis: Iterable[str] = (), segredo: str = "",
           genericas: Iterable[str] = ()) -> "Termos":  # fmt: skip
        """``valores``: nome, CPF, contas, endereço...; ``genericas``: palavras que não identificam
        ninguém (rua, banco...). Palavras com 4 ou mais letras e números com 4 ou mais dígitos."""
        ignorar = set(genericas)
        palavras = {n for v in valores for p in RE_PALAVRA.findall(v)
                    if len(n := _normal(p)) >= 4 and n not in ignorar}  # fmt: skip
        numeros = {d for v in valores for m in RE_NUMERO.findall(v)
                   if len(d := re.sub(r"\D", "", m).lstrip("0")) >= 4}  # fmt: skip
        padroes = tuple(re.compile(r"(?<!\d)" + r"[.\- ]?".join(n) + r"(?!\d)")
                        for n in sorted(numeros))  # fmt: skip
        return cls(
            frozenset(palavras), padroes, tuple(s.casefold() for s in sensiveis if s), segredo
        )

    def achou(self, texto: str) -> bool:
        return bool(
            set(RE_PALAVRA.findall(_normal(texto))) & self.palavras
            or any(p.search(texto) for p in self.numeros)
            or any(s in texto.casefold() for s in self.sensiveis)
        )


def _textos(campo: str, valor: object) -> list[str]:
    forma = CAMPOS.get(campo, RE_ROTULO)
    if isinstance(forma, list):
        forma = forma[0]
    if forma not in _TEXTUAIS:
        return []
    valores = valor if isinstance(valor, list) else [valor]
    # o número da linha da pilha não é um número do titular
    return [str(v).rsplit(":", 1)[0] if forma is RE_QUADRO else str(v) for v in valores]


def conferir_linha(texto: str, termos: Termos) -> dict[str, object] | None:
    """O evento, se estiver todo na lista fechada e sem nenhum termo do titular. Não confia no
    arquivo: pode ter sido editado, ou escrito por um programa com defeito."""
    try:
        o = json.loads(texto)
        datetime.fromisoformat(o["t"])
    except (ValueError, TypeError, KeyError):
        return None
    if not isinstance(o, dict) or set(o) != {"t", "s", "n", "nv", "ev", "d"}:
        return None
    if not (_na_forma(RE_HORA, o["t"]) and _na_forma(RE_SESSAO, o["s"]) and _na_forma(int, o["n"])
            and o["nv"] in ("I", "A", "E") and isinstance(o["ev"], str) and o["ev"] in EVENTOS
            and isinstance(o["d"], dict)):  # fmt: skip
        return None
    for campo, valor in o["d"].items():
        if campo not in CAMPOS or not _na_forma(CAMPOS[campo], valor):
            return None
        if any(termos.achou(t) for t in _textos(campo, valor)):
            return None
    return o


def conferir_linhas(linhas: Iterable[str], termos: Termos) -> tuple[list[str], int]:
    """Primeira passagem: as linhas que passam, refeitas do que foi conferido, e quantas ficaram
    retidas. Da linha retida não fica nada."""
    boas: list[str] = []
    retidas = 0
    for texto in linhas:
        if (o := conferir_linha(texto, termos)) is None:
            retidas += 1
        else:
            boas.append(json.dumps(o, ensure_ascii=False, separators=(",", ":")))
    return boas, retidas


def tem_dado_pessoal(texto: str, termos: Termos) -> bool:
    """Segunda passagem, no texto final inteiro: a rede para um defeito na primeira."""
    if RE_PESSOAL.search(texto) or (termos.segredo and termos.segredo in texto):
        return True
    return any(termos.achou(t) for x in texto.splitlines()
               for c, v in json.loads(x)["d"].items() for t in _textos(c, v))  # fmt: skip
