"""Cadastros pela interface: titulares (as contas e corretoras nascem na importação)."""

from pathlib import Path

from coir.db import repositorio as repo
from coir.db.sessao import conectar
from coir.domain.cpf import cpf_valido, so_digitos
from coir.erros import ImportacaoError


def adicionar_titular(banco: Path, cpf: str, nome: str) -> int:
    """Cadastra o titular (CPF com dígitos verificadores conferidos) e devolve o id."""
    if not nome.strip():
        raise ImportacaoError("informe o nome do titular")
    if not cpf_valido(cpf):
        raise ImportacaoError("CPF inválido (confira os dígitos)")
    with conectar(banco) as fabrica, fabrica.begin() as s:
        return repo.adicionar_titular(s, so_digitos(cpf), nome).id


def renomear_titular(banco: Path, id_titular: int, nome: str) -> None:
    with conectar(banco) as fabrica, fabrica.begin() as s:
        repo.renomear_titular(s, id_titular, nome)
