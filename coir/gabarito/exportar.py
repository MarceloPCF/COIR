"""Grava e lê o gabarito em CSV (UTF-8, Decimal como texto) + manifesto.

Os CSV ficam na pasta de dados do usuário; o mapa de pseudônimos fica em ``_privado/``.
Nada é gravado dentro do repositório.
"""

import csv
import hashlib
import json
import re
from collections import Counter
from datetime import date
from decimal import Decimal
from enum import Enum
from pathlib import Path

from pydantic import BaseModel

import coir
from coir.erros import GabaritoError
from coir.gabarito.modelo import (
    ApuracaoMensalGabarito,
    CustodiaInicialGabarito,
    DarfGabarito,
    EventoGabarito,
    Gabarito,
    IsentoMensal,
    OperacaoGabarito,
    PosicaoGabarito,
)
from coir.gabarito.pseudonimos import MapaPseudonimos

PASTA_PRIVADA = "_privado"
ARQUIVO_MAPA = "mapa-pseudonimos.json"
ARQUIVO_MANIFESTO = "manifesto.json"
ARQUIVOS: dict[str, tuple[str, type[BaseModel]]] = {
    "operacoes": ("operacoes.csv", OperacaoGabarito),
    "eventos": ("eventos.csv", EventoGabarito),
    "eventos_cadastro": ("eventos_cadastro.csv", EventoGabarito),
    "posicoes": ("posicoes.csv", PosicaoGabarito),
    "apuracao": ("apuracao_mensal.csv", ApuracaoMensalGabarito),
    "darfs": ("darf.csv", DarfGabarito),
    "isentos": ("isentos.csv", IsentoMensal),
    "custodia_inicial": ("custodia_inicial.csv", CustodiaInicialGabarito),
}
RAIZ_REPOSITORIO = Path(coir.__file__).resolve().parents[1]


def _texto(valor: object) -> str:
    if valor is None:
        return ""
    if isinstance(valor, Decimal):
        return format(valor, "f")
    if isinstance(valor, date):
        return valor.isoformat()
    if isinstance(valor, Enum):
        return str(valor.value)
    return str(valor)


def sha256(caminho: Path) -> str:
    return hashlib.sha256(caminho.read_bytes()).hexdigest()


def exigir_fora_do_repositorio(destino: Path) -> None:
    alvo = destino.resolve()
    if alvo == RAIZ_REPOSITORIO or RAIZ_REPOSITORIO in alvo.parents:
        raise GabaritoError(
            f"Destino dentro do repositório ({alvo}); o gabarito tem dados pessoais e deve "
            "ficar na pasta COIR_DADOS."
        )


def gravar_csv(caminho: Path, registros: list[BaseModel], modelo: type[BaseModel]) -> None:
    campos = list(modelo.model_fields)
    with caminho.open("w", encoding="utf-8", newline="") as arquivo:
        escritor = csv.writer(arquivo, lineterminator="\n")
        escritor.writerow(campos)
        for registro in registros:
            escritor.writerow([_texto(getattr(registro, campo)) for campo in campos])


def ler_csv[M: BaseModel](caminho: Path, modelo: type[M]) -> list[M]:
    if not caminho.is_file():
        raise GabaritoError(f"Arquivo do gabarito ausente: {caminho.name}")
    with caminho.open(encoding="utf-8", newline="") as arquivo:
        leitor = csv.DictReader(arquivo)
        if leitor.fieldnames != list(modelo.model_fields):
            raise GabaritoError(f"{caminho.name}: colunas diferentes do esperado")
        return [
            modelo.model_validate({k: (v if v != "" else None) for k, v in linha.items()})
            for linha in leitor
        ]


def verificar_pseudonimizacao(destino: Path, mapa: MapaPseudonimos) -> None:
    """Falha se algum CPF ou número de conta real aparecer nos arquivos públicos do gabarito."""
    reais = [v for v in mapa.valores_reais() if v]
    padrao_cpf = re.compile(r"\d{3}\.?\d{3}\.?\d{3}-?\d{2}")
    for nome, _modelo in ARQUIVOS.values():
        caminho = destino / nome
        with caminho.open(encoding="utf-8", newline="") as arquivo:
            leitor = csv.DictReader(arquivo)
            for n, linha in enumerate(leitor, start=2):
                for coluna in ("titular", "conta"):
                    if coluna in linha and linha[coluna] in reais:
                        raise GabaritoError(f"{nome} linha {n}: valor real em '{coluna}'")
                for valor in linha.values():
                    for achado in padrao_cpf.findall(valor or ""):
                        if re.sub(r"\D", "", achado) in mapa.titulares:
                            raise GabaritoError(f"{nome} linha {n}: CPF real em arquivo público")


def contagens(gabarito: Gabarito) -> dict[str, object]:
    return {
        "operacoes": dict(sorted(Counter(o.categoria.value for o in gabarito.operacoes).items())),
        "notas": len({(o.corretora, o.conta, o.nota, o.data) for o in gabarito.operacoes}),
        "eventos": len(gabarito.eventos),
        "eventos_cadastro": len(gabarito.eventos_cadastro),
        "posicoes": dict(sorted(Counter(p.categoria.value for p in gabarito.posicoes).items())),
        "apuracao_meses": len(gabarito.apuracao),
        "darfs": len(gabarito.darfs),
        "custodia_inicial": len(gabarito.custodia_inicial),
        "anos": [
            min(o.data.year for o in gabarito.operacoes),
            max(o.data.year for o in gabarito.operacoes),
        ],
    }


def exportar(
    gabarito: Gabarito, mapa: MapaPseudonimos, origem: Path, destino: Path
) -> dict[str, object]:
    """Grava os CSV, o mapa privado e o manifesto. Devolve o manifesto."""
    exigir_fora_do_repositorio(destino)
    destino.mkdir(parents=True, exist_ok=True)
    mapa.salvar(destino / PASTA_PRIVADA / ARQUIVO_MAPA)
    arquivos: dict[str, str] = {}
    for campo, (nome, modelo) in ARQUIVOS.items():
        registros: list[BaseModel] = getattr(gabarito, campo)
        gravar_csv(destino / nome, registros, modelo)
        arquivos[nome] = sha256(destino / nome)
    verificar_pseudonimizacao(destino, mapa)
    manifesto: dict[str, object] = {
        "versao_coir": coir.__version__,
        "origem": {"arquivo": origem.name, "sha256": sha256(origem)},
        "contagens": contagens(gabarito),
        "arquivos": arquivos,
        "avisos": gabarito.avisos,
    }
    (destino / ARQUIVO_MANIFESTO).write_text(
        json.dumps(manifesto, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    return manifesto


def carregar(pasta: Path) -> Gabarito:
    """Lê um gabarito exportado (ou um candidato no mesmo formato)."""
    dados: dict[str, object] = {
        campo: ler_csv(pasta / nome, modelo) for campo, (nome, modelo) in ARQUIVOS.items()
    }
    manifesto = pasta / ARQUIVO_MANIFESTO
    avisos = (
        json.loads(manifesto.read_text(encoding="utf-8")).get("avisos", [])
        if manifesto.is_file()
        else []
    )
    return Gabarito.model_validate({**dados, "avisos": avisos})
