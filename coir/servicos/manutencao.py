"""Backups do banco e registro de atividades para a tela de configurações."""

import re
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

from coir.config import ler_registro
from coir.db.sessao import fazer_backup, listar_backups, restaurar_backup
from coir.erros import BancoError


@dataclass(frozen=True)
class Backup:
    caminho: Path
    feito_em: datetime | None  # pelo nome do arquivo (coir-AAAAMMDD-HHMMSS.db)
    tamanho: int


def _quando(arquivo: Path) -> datetime | None:
    """coir-AAAAMMDD-HHMMSS[-NN].db -> data e hora do backup."""
    try:
        return datetime.strptime("-".join(arquivo.stem.split("-")[1:3]), "%Y%m%d-%H%M%S")
    except ValueError:
        return None


def backups(banco: Path) -> list[Backup]:
    return [Backup(b, _quando(b), b.stat().st_size) for b in listar_backups(banco)]


def fazer_agora(banco: Path, agora: datetime) -> Path:
    destino = fazer_backup(banco, agora)
    if destino is None:
        raise BancoError("ainda não há banco para copiar")
    return destino


def restaurar(banco: Path, origem: Path, agora: datetime) -> Path | None:
    return restaurar_backup(banco, origem, agora)


# ------------------------------------------------------------------- registro de atividades
@dataclass(frozen=True)
class EntradaRegistro:
    quando: datetime
    tipo: str  # Informação, Aviso ou Erro
    area: str
    mensagem: str
    detalhe: str = ""  # detalhe técnico (traceback), quando houver

    @property
    def problema(self) -> bool:
        return self.tipo != "Informação"


TIPOS = {"DEBUG": "Informação", "INFO": "Informação", "WARNING": "Aviso", "ERROR": "Erro",
         "CRITICAL": "Erro"}  # fmt: skip
AREAS = (  # prefixo do módulo que registrou -> área (o mais específico primeiro)
    ("coir.servicos.importacao", "Importação de notas"),
    ("coir.servicos.b3", "Planilhas da B3"),
    ("coir.importers", "Planilhas da B3"),
    ("coir.servicos.atualizacao", "Tabelas de referência"),
    ("coir.referencias", "Tabelas de referência"),
    ("coir.servicos.darf", "DARF"),
    ("coir.servicos.proventos", "Proventos"),
    ("coir.servicos.declaracao", "Declaração anual"),
    ("coir.servicos", "Cálculo"),
    ("coir.engine", "Cálculo"),
    ("coir.db", "Banco e backups"),
    ("alembic", "Banco e backups"),
    ("coir.config", "Configuração"),
    ("coir.ui.janela", "Programa"),
    ("coir.ui", "Interface"),
    ("nicegui", "Interface"),
)
RE_LINHA = re.compile(
    r"(?P<quando>\d{4}-\d\d-\d\d \d\d:\d\d:\d\d),\d+ (?P<nivel>[A-Z]+) (?P<origem>\S+): "
    r"(?P<mensagem>.*)"
)
RE_TELA = re.compile(r"coir[\\/]ui[\\/]telas[\\/](\w+)\.py")
RE_ULTIMA = re.compile(r"^\w+(Error|Exception|Warning)\b")


NOMES_TELA = {
    "assistente": "Primeiro uso", "painel": "Painel", "importar": "Importar",
    "importar_b3": "Importar (B3)", "notas": "Notas", "operacoes": "Operações",
    "carteira": "Carteira e preço médio", "eventos": "Eventos corporativos", "opcoes": "Opções",
    "futuros": "Futuros", "proventos": "Proventos", "outros_custos": "Outros custos",
    "apuracao": "Apuração mensal", "darf": "DARF", "declaracao": "Declaração anual",
    "conciliacao": "Conciliação B3", "cadastros": "Cadastros", "relatorios": "Relatórios",
    "configuracoes": "Configurações",
}  # fmt: skip


def _area(origem: str, detalhe: str) -> str:
    if telas := RE_TELA.findall(detalhe):  # erro numa tela: o nome da tela ajuda mais
        return f"Tela {NOMES_TELA.get(telas[-1], telas[-1])}"
    return next((nome for prefixo, nome in AREAS if origem.startswith(prefixo)), "Sistema")


def _mensagem(mensagem: str, detalhe: str) -> str:
    """Erro com traceback: a última linha ("KeyError: ...") diz mais que a mensagem do NiceGUI."""
    ultimas = [x.strip() for x in detalhe.splitlines() if RE_ULTIMA.match(x.strip())]
    if ultimas and (not mensagem.strip() or mensagem.strip() in ultimas[-1]):
        return f"Erro inesperado: {ultimas[-1]}"
    return mensagem


def registro(dados: Path, linhas: int = 3000) -> list[EntradaRegistro]:
    """O registro de atividades em entradas legíveis, a mais recente primeiro. As linhas de
    continuação (tracebacks) viram o detalhe técnico da entrada anterior."""
    brutas: list[tuple[re.Match[str], list[str]]] = []
    for linha in ler_registro(dados, linhas):
        if m := RE_LINHA.fullmatch(linha):
            brutas.append((m, []))
        elif brutas:
            brutas[-1][1].append(linha)
    entradas: list[EntradaRegistro] = []
    for m, resto in brutas:
        detalhe = "\n".join(resto).strip()
        entradas.append(EntradaRegistro(
            quando=datetime.strptime(m["quando"], "%Y-%m-%d %H:%M:%S"),
            tipo=TIPOS.get(m["nivel"], "Informação"),
            area=_area(m["origem"], detalhe),
            mensagem=_mensagem(m["mensagem"], detalhe),
            detalhe=detalhe,
        ))  # fmt: skip
    return entradas[::-1]
