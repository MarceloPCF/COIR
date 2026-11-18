"""Dados abertos da CVM: eventos do formulário de referência (FRE) e dos avisos aos acionistas
(IPE), com o protocolo do documento na referência."""

from pathlib import Path

from coir.fontes.leitura import opcionais
from coir.fontes.modelo import ORIGENS_CVM, Evento, evento_valido, partes_da_origem
from coir.fontes.provedores import b3

EVENTOS = "eventos_cvm.csv"  # formato 3: só as linhas da CVM, com conferido_cotahist
EVENTOS_JUNTOS = "eventos_b3.csv"  # formato 2: as origens juntas numa linha (coluna origem)


def eventos(pasta: Path) -> list[Evento]:
    """Eventos utilizáveis com alguma origem na CVM: o arquivo da CVM e, com a fonte da bolsa
    ligada, as linhas da CVM que estiverem na tabela junta (pacote do formato 2). Desligada a
    fonte da bolsa, a tabela junta não é aberta."""
    linhas = opcionais(pasta / EVENTOS)
    if b3.ligada():
        linhas += opcionais(pasta / EVENTOS_JUNTOS)
    return [Evento.de_linha(r) for r in linhas
            if partes_da_origem(r.get("origem", "")) & ORIGENS_CVM
            and evento_valido(r)]  # fmt: skip
