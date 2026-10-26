"""Tema do COIR: o único lugar com cores, fontes e medidas da interface.

Direção: a identidade do site e da logomarca do COIR — marinho com o azul do símbolo e o verde do
botão de download — sobre papel claro. Fontes do próprio Windows (o app é offline): Bahnschrift,
de desenho DIN, em títulos, números de destaque e rótulos de campo; Segoe UI no texto e nas
tabelas, sempre com algarismos tabulares.
"""

from dataclasses import dataclass
from functools import cache
from pathlib import Path

from nicegui import ui

ESTATICOS = Path(__file__).parent / "estaticos"
SIMBOLO = ESTATICOS / "coir-simbolo.svg"  # logomarca: o "C" com o arco azul
ICONE = ESTATICOS / "coir-icone-256.png"


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
    topo: str  # faixa do topo (marca, titular e ano)
    topo_tinta: str
    destaque: str  # o verde do site: item ativo e ação principal de cada tela
    eixo: str  # rótulos e eixos dos gráficos


CLARO = Paleta(
    fundo="#F4F7FA",
    superficie="#FFFFFF",
    linha="#D9E2EC",
    tinta="#243B53",
    tinta_suave="#52606D",
    caneta="#1B5DA8",
    caneta_fundo="#E3EEFA",
    carimbo="#A86B00",
    positivo="#1A7F37",
    negativo="#B42318",
    lateral="#102A43",
    lateral_tinta="#C9D8E6",
    topo="#102A43",
    topo_tinta="#FFFFFF",
    destaque="#2EA043",
    eixo="#627D98",
)
ESCURO = Paleta(
    fundo="#0B1A2B",
    superficie="#102A43",
    linha="#243B53",
    tinta="#E6EDF4",
    tinta_suave="#9FB3C8",
    caneta="#8FBBEA",
    caneta_fundo="#1B3D5C",
    carimbo="#F0B429",
    positivo="#4AC26B",
    negativo="#FF8A80",
    lateral="#081525",
    lateral_tinta="#BCCCDC",
    topo="#081525",
    topo_tinta="#FFFFFF",
    destaque="#2EA043",
    eixo="#829AB1",
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
body.body--dark {{ --q-primary: #2E6FB8; }}
.coir-topo {{ background: var(--coir-topo); color: var(--coir-topo-tinta);
  border-bottom: 1px solid rgba(255,255,255,.08); box-shadow: none; }}
.coir-topo::before {{ content: ""; position: absolute; left: 0; top: 0; bottom: 0; width: 6px;
  background: var(--coir-destaque); }}
.coir-topo .q-field__label, .coir-topo .q-field__native, .coir-topo .q-field__append,
.coir-topo .q-field__control {{ color: var(--coir-topo-tinta) !important; }}
.coir-topo .q-field--outlined .q-field__control:before {{ border-color: rgba(255,255,255,.28); }}
.coir-simbolo svg {{ width: 34px; height: 34px; display: block; }}
.coir-marca {{ font-family: {FONTE_TITULO}; font-size: 28px; font-weight: 700; line-height: 1;
  letter-spacing: .03em; color: var(--coir-topo-tinta); }}
.coir-marca-nome {{ font-size: 12px; line-height: 1.2; color: #A8D4FF; max-width: 150px; }}
.coir-lateral {{ background: var(--coir-lateral); color: var(--coir-lateral-tinta); }}
.coir-lateral, .coir-lateral * {{ scrollbar-color: rgba(255,255,255,.22) transparent; }}
.coir-lateral .q-item {{ color: var(--coir-lateral-tinta); border-radius: 4px; min-height: 36px;
  margin: 1px 8px; }}
.coir-lateral .q-item--active, .coir-lateral .coir-ativo {{ background: rgba(168,212,255,.14);
  color: #FFFFFF; box-shadow: inset 4px 0 0 var(--coir-destaque); }}
.coir-grupo {{ font-family: {FONTE_TITULO}; font-size: 14px; font-weight: 600;
  letter-spacing: .04em; text-transform: uppercase; color: var(--coir-tinta-suave); }}
.coir-lateral .coir-grupo {{ color: #A8D4FF; min-height: 40px; margin: 8px 8px 0;
  padding: 0 12px; border-top: 1px solid rgba(255,255,255,.08); }}
.coir-lateral .q-expansion-item__toggle-icon {{ color: #A8D4FF; }}
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
.coir-item-ativo {{ background: var(--coir-caneta-fundo); color: var(--coir-caneta);
  font-weight: 600; }}
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


@cache
def simbolo_svg() -> str:
    """O símbolo da logomarca, para pôr no topo (inline: não depende de rota de arquivos)."""
    return SIMBOLO.read_text(encoding="utf-8")
