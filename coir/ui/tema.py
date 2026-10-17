"""Tema do COIR: o único lugar com cores, fontes e medidas da interface.

Direção: o vocabulário de formulário fiscal (campos numerados do DARF, tinta de caneta azul sobre
papel acinzentado), sem copiar plataformas de terceiros. Fontes do próprio Windows (o app é
offline): Bahnschrift, de desenho DIN, em títulos, números de destaque e rótulos de campo; Segoe UI
no texto e nas tabelas, sempre com algarismos tabulares.
"""

from dataclasses import dataclass

from nicegui import ui


@dataclass(frozen=True)
class Paleta:
    fundo: str  # papel
    superficie: str  # folhas, painéis e tabelas
    linha: str  # bordas finas e divisões
    tinta: str  # texto principal
    tinta_suave: str  # texto secundário
    caneta: str  # cor de ação e de seleção (primária)
    caneta_fundo: str  # seleção suave
    carimbo: str  # atenção: vencimento próximo, acumulado
    positivo: str
    negativo: str
    lateral: str  # barra lateral
    lateral_tinta: str


CLARO = Paleta(
    fundo="#E9EDEA",
    superficie="#FFFFFF",
    linha="#C9D1CC",
    tinta="#1C2A33",
    tinta_suave="#5A6872",
    caneta="#24489A",
    caneta_fundo="#E3E9F7",
    carimbo="#A65F00",
    positivo="#256B45",
    negativo="#B02A20",
    lateral="#1E2B36",
    lateral_tinta="#D5DEE5",
)
ESCURO = Paleta(
    fundo="#161D24",
    superficie="#1E2730",
    linha="#33414D",
    tinta="#E2E8ED",
    tinta_suave="#9AA8B3",
    caneta="#8DAEF5",
    caneta_fundo="#26375A",
    carimbo="#E2A54A",
    positivo="#6FC497",
    negativo="#F08A80",
    lateral="#11171D",
    lateral_tinta="#B9C5CF",
)

FONTE_TITULO = "Bahnschrift, 'DIN Alternate', 'Segoe UI Variable Display', 'Segoe UI', sans-serif"
FONTE_TEXTO = "'Segoe UI Variable Text', 'Segoe UI', system-ui, sans-serif"


def _variaveis(p: Paleta) -> str:
    return "".join(f"--coir-{nome.replace('_', '-')}: {valor};" for nome, valor in vars(p).items())


CSS = f"""
:root {{ {_variaveis(CLARO)} }}
body.body--dark {{ {_variaveis(ESCURO)} }}
body {{
  background: var(--coir-fundo); color: var(--coir-tinta);
  font-family: {FONTE_TEXTO}; font-size: 14px; font-variant-numeric: tabular-nums;
}}
h1, h2, h3, .coir-titulo {{ font-family: {FONTE_TITULO}; font-weight: 600; letter-spacing: 0; }}
.coir-titulo-pagina {{ font-family: {FONTE_TITULO}; font-size: 26px; font-weight: 600;
  line-height: 1.15; margin: 0; }}
.coir-subtitulo {{ color: var(--coir-tinta-suave); max-width: 72ch; }}
.coir-topo {{ background: var(--coir-superficie); color: var(--coir-tinta);
  border-bottom: 1px solid var(--coir-linha); box-shadow: none; }}
.coir-marca {{ font-family: {FONTE_TITULO}; font-size: 20px; font-weight: 700;
  letter-spacing: .02em; color: var(--coir-caneta); }}
.coir-lateral {{ background: var(--coir-lateral); color: var(--coir-lateral-tinta); }}
.coir-lateral .q-item {{ color: var(--coir-lateral-tinta); border-radius: 4px; min-height: 36px;
  margin: 1px 8px; }}
.coir-lateral .q-item--active, .coir-lateral .coir-ativo {{ background: rgba(141,174,245,.16);
  color: #FFFFFF; box-shadow: inset 3px 0 0 var(--coir-caneta); }}
.coir-grupo {{ font-family: {FONTE_TITULO}; font-size: 12px; color: var(--coir-tinta-suave);
  opacity: .8; padding: 14px 20px 4px; }}
.coir-lateral .coir-grupo {{ color: var(--coir-lateral-tinta); }}
.coir-folha {{ background: var(--coir-superficie); border: 1px solid var(--coir-linha);
  border-radius: 6px; padding: 16px 20px; }}
.coir-campo-rotulo {{ font-family: {FONTE_TITULO}; font-size: 12px;
  color: var(--coir-tinta-suave); }}
.coir-numero {{ font-family: {FONTE_TITULO}; font-variant-numeric: tabular-nums; }}
.coir-positivo {{ color: var(--coir-positivo); }}
.coir-negativo {{ color: var(--coir-negativo); }}
.coir-carimbo {{ color: var(--coir-carimbo); }}
.coir-vazio {{ border: 1px dashed var(--coir-linha); border-radius: 6px; padding: 32px;
  color: var(--coir-tinta-suave); text-align: left; }}
.coir-grade .ag-root-wrapper {{ --ag-font-family: {FONTE_TEXTO}; --ag-font-size: 13px;
  --ag-background-color: var(--coir-superficie); --ag-header-background-color: var(--coir-fundo);
  --ag-border-color: var(--coir-linha); --ag-foreground-color: var(--coir-tinta);
  --ag-text-color: var(--coir-tinta); --ag-header-text-color: var(--coir-tinta-suave);
  --ag-row-border-color: var(--coir-linha); --ag-odd-row-background-color: var(--coir-superficie);
  --ag-row-hover-color: var(--coir-caneta-fundo); --ag-selected-row-background-color:
  var(--coir-caneta-fundo); --ag-pinned-row-border: 2px solid var(--coir-linha); }}
.coir-grade .ag-row-pinned {{ font-weight: 600; background: var(--coir-fundo); }}
.ag-right-aligned-cell {{ font-variant-numeric: tabular-nums; }}
.coir-item-lista {{ border-bottom: 1px solid var(--coir-linha); }}
.coir-item-lista:last-child {{ border-bottom: none; }}
.coir-envio {{ box-shadow: none; border: 1px dashed var(--coir-caneta);
  background: var(--coir-caneta-fundo); }}
.coir-envio .q-uploader__header {{ background: transparent; color: var(--coir-caneta); }}
.coir-envio .q-uploader__subtitle {{ display: none; }}
.coir-envio .q-uploader__list {{ min-height: 0; padding: 0 8px; }}
:focus-visible {{ outline: 2px solid var(--coir-caneta); outline-offset: 2px; }}
@media (prefers-reduced-motion: reduce) {{ * {{ transition: none !important;
  animation: none !important; }} }}
"""


def aplicar() -> None:
    """Cores do Quasar e CSS do tema na página atual."""
    ui.colors(primary=CLARO.caneta, secondary=CLARO.tinta_suave, accent=CLARO.carimbo,
              positive=CLARO.positivo, negative=CLARO.negativo, warning=CLARO.carimbo,
              dark=ESCURO.fundo, dark_page=ESCURO.fundo)  # fmt: skip
    ui.add_css(CSS)
    ui.button.default_props("no-caps unelevated")
