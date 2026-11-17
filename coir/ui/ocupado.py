"""Aviso "Processando dados, aguarde" por cima da janela inteira.

Enquanto o COIR grava (importação, eventos automáticos, tabelas novas, restauração), trocar de
tela ou clicar em outra coisa deixava a operação terminar numa tela que já não existia. A
cortina cobre o menu e a tela, mostra o que está sendo feito e some quando termina.

Uso: ``with ocupado.processando("Importando as notas"): await run.io_bound(...)``. O andamento
(``detalhar``) pode ser escrito da thread de trabalho: a janela lê o texto por um temporizador.
"""

from collections.abc import Iterator
from contextlib import contextmanager
from dataclasses import dataclass, field

from nicegui import ui

INTERVALO = 0.2  # segundos entre as leituras da situação pela janela


@dataclass
class _Situacao:
    textos: list[str] = field(default_factory=list[str])  # um por operação em andamento
    detalhe: str = ""

    @property
    def operacoes(self) -> int:
        return len(self.textos)

    @property
    def ativo(self) -> bool:
        return bool(self.textos)

    @property
    def texto(self) -> str:
        """O da operação mais recente; quando ela termina, volta o da que a envolve."""
        return self.textos[-1] if self.textos else ""


SITUACAO = _Situacao()  # um programa, um usuário: a situação é a mesma em todas as janelas


@contextmanager
def processando(texto: str) -> Iterator[None]:
    """Mostra a cortina com ``texto`` até o bloco terminar, com ou sem erro."""
    SITUACAO.textos.append(texto)
    SITUACAO.detalhe = ""
    try:
        yield
    finally:
        SITUACAO.textos.pop()
        SITUACAO.detalhe = ""


def detalhar(detalhe: str) -> None:
    """Segunda linha da cortina (ex.: "3 de 12 arquivos"); pode ser chamada de outra thread."""
    SITUACAO.detalhe = detalhe


def montar() -> None:
    """A cortina da página: fica escondida e aparece enquanto houver operação em andamento."""
    cortina = ui.element("div").classes("coir-cortina").mark("processando")
    with cortina, ui.column().classes("coir-cortina-cartao items-center gap-2"):
        ui.spinner(size="42px")
        texto = ui.label("").classes("coir-cortina-texto").mark("processando-texto")
        detalhe = ui.label("").classes("coir-subtitulo text-sm").mark("processando-detalhe")
        ui.label("Aguarde: não feche o programa.").classes("coir-subtitulo text-sm")
    cortina.set_visibility(False)

    def atualizar() -> None:
        cortina.set_visibility(SITUACAO.ativo)
        if SITUACAO.ativo:
            texto.set_text(SITUACAO.texto)
            detalhe.set_text(SITUACAO.detalhe)

    ui.timer(INTERVALO, atualizar)
