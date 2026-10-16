"""Eventos do gabarito (eventos.csv, pseudonimizado) -> CSV de eventos do COIR (com as contas
reais, pelo mapa privado de pseudônimos). O arquivo gerado fica em COIR_DADOS: tem número de conta
e nunca vai para o repositório. O usuário revisa e acrescenta o que a planilha não tinha.
"""

import csv
from pathlib import Path

from coir.erros import GabaritoError
from coir.gabarito.pseudonimos import MapaPseudonimos

TIPOS = {
    "bonificação": "BONIFICACAO",
    "desdobramento": "DESDOBRAMENTO",
    "grupamento": "GRUPAMENTO",
    "subscrição": "SUBSCRICAO",
    "redução de capital": "REDUCAO_CAPITAL",
}
COLUNAS: tuple[str, ...] = (
    "data", "corretora", "conta", "ticker", "tipo", "cv", "quantidade", "valor",
    "corretora_destino", "conta_destino", "observacao", "origem",
)  # fmt: skip


def exportar_eventos(pasta_gabarito: Path, destino: Path) -> int:
    """Escreve `destino` (não sobrescreve: o usuário pode ter editado) e devolve as linhas."""
    repositorio = Path(__file__).resolve().parents[2]
    if destino.resolve().is_relative_to(repositorio):
        raise GabaritoError("eventos.csv tem números de conta: grave fora do repositório")
    if destino.exists():
        raise GabaritoError(f"{destino.name} já existe (revise-o ou apague para gerar de novo)")
    mapa = MapaPseudonimos.carregar(pasta_gabarito / "_privado" / "mapa-pseudonimos.json")
    contas = {pseudo: real.split("|", 1) for real, pseudo in mapa.contas.items()}
    origem = pasta_gabarito / "eventos.csv"
    if not origem.is_file():
        raise GabaritoError("eventos.csv do gabarito ausente (rode coir gabarito gerar)")
    linhas: list[dict[str, str]] = []
    with origem.open(encoding="utf-8", newline="") as arquivo:
        for n, e in enumerate(csv.DictReader(arquivo), start=2):
            tipo = TIPOS.get(e["tipo"].strip().lower())
            if tipo is None:
                raise GabaritoError(f"eventos.csv linha {n}: tipo {e['tipo']!r} sem tradução")
            if e["conta"] not in contas:
                raise GabaritoError(f"eventos.csv linha {n}: conta {e['conta']} fora do mapa")
            corretora, conta = contas[e["conta"]]
            linhas.append({
                "data": e["data"], "corretora": corretora, "conta": conta, "ticker": e["papel"],
                "tipo": tipo, "cv": e["cv"], "quantidade": _numero(e["quantidade"]),
                "valor": _numero(e["total"]), "corretora_destino": "", "conta_destino": "",
                "observacao": f"gabarito: {e['tipo']} (linha {e['linha']} de {e['categoria']})",
                "origem": "gabarito",
            })  # fmt: skip
    destino.parent.mkdir(parents=True, exist_ok=True)
    with destino.open("w", encoding="utf-8", newline="") as arquivo:
        escritor = csv.DictWriter(arquivo, fieldnames=COLUNAS, lineterminator="\n")
        escritor.writeheader()
        escritor.writerows(linhas)
    return len(linhas)


def _numero(texto: str) -> str:
    """'300.00000000' -> '300'; '890.61500000' -> '890.615'; vazio fica vazio."""
    texto = texto.strip()
    if not texto:
        return ""
    if "." in texto:
        texto = texto.rstrip("0").rstrip(".")
    return texto or "0"
