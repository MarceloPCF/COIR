"""Tema da janela (sistema, claro ou escuro), gravado na configuração e aplicado em cada página."""

from weakref import WeakKeyDictionary

from nicegui import Client, ui

from coir.config import ler_preferencia, salvar_preferencia
from coir.erros import ConfiguracaoError

TEMAS = {"sistema": None, "claro": False, "escuro": True}
_ESCUROS: WeakKeyDictionary[Client, ui.dark_mode] = WeakKeyDictionary()


def tema() -> str:
    try:
        valor = ler_preferencia("tema")
    except ConfiguracaoError:
        return "sistema"
    return valor if valor in TEMAS else "sistema"


def aplicar_tema() -> ui.dark_mode:
    """Cria o controle de tema da página com a preferência gravada."""
    escuro = ui.dark_mode(value=TEMAS[tema()])
    _ESCUROS[ui.context.client] = escuro
    return escuro


def escolher_tema(nome: str) -> None:
    escuro = _ESCUROS.get(ui.context.client)
    if escuro is not None:
        escuro.set_value(TEMAS[nome])
    try:
        salvar_preferencia("tema", nome)
    except ConfiguracaoError as erro:
        ui.notify(f"O tema vale só nesta janela: {erro}", type="warning")


def verificar_tabelas() -> bool:
    """Opção "verificar ao abrir o app" das tabelas de referência (desligada por padrão)."""
    try:
        return ler_preferencia("verificar_tabelas") == "sim"
    except ConfiguracaoError:
        return False


def escolher_verificar_tabelas(sim: bool) -> None:
    try:
        salvar_preferencia("verificar_tabelas", "sim" if sim else "nao")
    except ConfiguracaoError as erro:
        ui.notify(f"A opção vale só nesta janela: {erro}", type="warning")


def alternar_tema() -> None:
    """Botão do topo: claro <-> escuro (sai do "sistema")."""
    escuro = _ESCUROS.get(ui.context.client)
    escolher_tema("claro" if escuro is not None and escuro.value else "escuro")
