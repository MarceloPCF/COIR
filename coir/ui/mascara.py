"""Olho do topo: esconde os valores em R$ da tela, como nos apps de banco.

Só apresentação: o cálculo, o banco, os arquivos exportados e a DARF não mudam. As telas pedem o
texto do valor por `dinheiro` (as grades de `tabela` e os gráficos por `grafico` já fazem isso) e,
com os valores ocultos, recebem a máscara. Quantidade, preço médio, preço e percentual continuam à
mostra (escolha do usuário em 27/11/2026), e os campos de digitação ficam com o valor real.
A escolha fica gravada na configuração, como o tema, e vale a partir da próxima tela desenhada.
"""

import re
from decimal import Decimal
from typing import Any
from weakref import WeakKeyDictionary

from nicegui import Client, ui

from coir.config import ler_preferencia, salvar_preferencia
from coir.erros import ConfiguracaoError
from coir.reports.modelo import TipoColuna, Valor, formatar

MASCARA = "•••••"
CHAVE = "valores_ocultos"
# "R$ 1.234,56", "R$ 1234,56", "R$ -0,50", "R$ 12,3456", "R$ 35 mil" (o ponto final da frase fica)
RE_REAIS = re.compile(r"R\$\s?-?\d+(?:\.\d{3})*(?:,\d+)?")
_OCULTOS: WeakKeyDictionary[Client, bool] = WeakKeyDictionary()


def _gravado() -> bool:
    try:
        return ler_preferencia(CHAVE) == "sim"
    except ConfiguracaoError:
        return False


def ocultos() -> bool:
    """Os valores em R$ estão ocultos nesta janela? (lido da configuração uma vez por janela)"""
    try:
        cliente = ui.context.client
    except RuntimeError:  # fora de uma janela (tarefa sem tela): vale o que está gravado
        return _gravado()
    if cliente not in _OCULTOS:
        _OCULTOS[cliente] = _gravado()
    return _OCULTOS[cliente]


def alternar() -> bool:
    """Botão do olho: troca e grava a escolha; devolve se os valores ficaram ocultos."""
    novo = not ocultos()
    _OCULTOS[ui.context.client] = novo
    try:
        salvar_preferencia(CHAVE, "sim" if novo else "nao")
    except ConfiguracaoError as erro:
        ui.notify(f"A escolha vale só nesta janela: {erro}", type="warning")
    return novo


def dinheiro(valor: Decimal, milhar: bool = True) -> str:
    """Valor em R$ para mostrar na tela (nunca para preencher um campo)."""
    return MASCARA if ocultos() else formatar(valor, TipoColuna.DINHEIRO, milhar)


def campo(entrada: ui.input) -> ui.input:
    """Campo que já vem preenchido com um valor em R$: ocultos, os algarismos aparecem como
    pontos até o usuário clicar nele para digitar (o valor do campo não muda; tema.CSS)."""
    if ocultos():
        entrada.classes("coir-campo-oculto")
    return entrada


def texto(frase: str) -> str:
    """Aviso ou explicação com valores em R$ no meio do texto ("excedente de R$ 1.234,56"):
    ocultos, só os valores viram a máscara."""
    return RE_REAIS.sub(f"R$ {MASCARA}", frase) if ocultos() else frase


def celula(valor: Valor, tipo: TipoColuna) -> str:
    """Texto de uma célula de grade: as colunas de dinheiro e os R$ dentro dos textos levam a
    máscara."""
    if not ocultos():
        return formatar(valor, tipo)
    if tipo is TipoColuna.DINHEIRO and isinstance(valor, Decimal | int):
        return MASCARA
    return RE_REAIS.sub(f"R$ {MASCARA}", valor) if isinstance(valor, str) else formatar(valor, tipo)


def grafico(opcoes: dict[str, Any]) -> dict[str, Any]:
    """Opções de um gráfico de valores em R$: ocultos, o eixo de valores e a dica mostram a
    máscara (as barras continuam, sem os números)."""
    if not ocultos():
        return opcoes
    mascara = f"() => '{MASCARA}'"
    saida = dict(opcoes)
    for nome in ("xAxis", "yAxis"):
        eixo = saida.get(nome)
        if isinstance(eixo, dict) and eixo.get("type") == "value":
            rotulo = {**eixo.get("axisLabel", {}), ":formatter": mascara}
            saida[nome] = {**eixo, "axisLabel": rotulo}
    saida["tooltip"] = {**saida.get("tooltip", {}), ":valueFormatter": mascara}
    return saida
