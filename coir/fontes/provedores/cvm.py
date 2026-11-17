"""Dados abertos da CVM: eventos do formulário de referência (FRE) e dos avisos aos acionistas
(IPE), com o protocolo do documento na referência."""

from pathlib import Path

from coir.fontes.leitura import opcionais
from coir.fontes.modelo import ORIGENS_CVM, Evento, evento_valido, partes_da_origem

# a tabela de eventos ainda junta as origens numa linha (coluna origem, "b3+cvm_fre"): cada
# provedor pega as linhas que a sua fonte traz, e a composição junta as repetidas
EVENTOS = "eventos_b3.csv"


def eventos(pasta: Path) -> list[Evento]:
    """Eventos utilizáveis com alguma origem na CVM."""
    return [Evento.de_linha(r) for r in opcionais(pasta / EVENTOS)
            if partes_da_origem(r.get("origem", "")) & ORIGENS_CVM
            and evento_valido(r)]  # fmt: skip
