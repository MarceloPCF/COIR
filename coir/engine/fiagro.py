"""Avisos de FII e Fiagro na declaração anual (decisões do usuário de 10/11/2026, leitura
conservadora). O motor apura FII e Fiagro na mesma categoria, como a ficha única do programa
IRPF; o tipo do fundo só é conhecido pelo cadastro da B3 (tipo FIAGRO), então o que segue são
avisos, sem mudar o cálculo.

- Compensação cruzada: a IN RFB 1.585/2015 art. 37 §2º só autoriza compensar perda de FII com
  ganho de FII; nenhuma norma autoriza compensar Fiagro com FII (a Lei 8.668 art. 20-D só remete
  às regras gerais). Ano com resultado de sinais opostos entre os dois: aviso.
- Rendimento de Fiagro antes de 11/06/2021: sem isenção (parâmetro ``fiagro_rendimento_isento``;
  Lei 11.033/2004 art. 3º III, red. Lei 14.130/2021, veto derrubado no DOU de 11/06/2021).
- Requisitos da isenção de rendimento de FII e Fiagro (Lei 11.033 art. 3º §1º; parâmetro
  ``fii_isencao_minimo_cotistas``): o COIR não sabe quantos cotistas o fundo tem nem a fatia do
  titular; o aviso diz o que conferir no informe do fundo.
"""

from collections.abc import Iterable, Mapping
from datetime import date
from decimal import Decimal

from coir.domain.fiscal import ZERO, Categoria, Movimento
from coir.domain.formato import reais
from coir.domain.parametros import ParametrosFiscais
from coir.domain.proventos import Provento, TipoProvento


def _vigente(params: ParametrosFiscais, nome: str, dia: date) -> str | None:
    achados = [p for p in params.parametros if p.nome == nome and p.vigente(dia)]
    return achados[0].valor if len(achados) == 1 else None


def avisos_fii_fiagro(
    titular: int,
    ano: int,
    movimentos: Iterable[Movimento],
    proventos: Iterable[Provento],
    fiagro: set[str],
    params: ParametrosFiscais,
) -> list[str]:
    """``fiagro``: os tickers que o cadastro da B3 dá como Fiagro."""
    avisos: list[str] = []
    resultado: Mapping[bool, list[Decimal]] = {True: [], False: []}
    for m in movimentos:
        if (m.titular == titular and m.data.year == ano and m.categoria is Categoria.FII
                and m.resultado):  # fmt: skip
            resultado[m.ticker in fiagro].append(m.resultado)
    ganho = {k: any(r > 0 for r in v) for k, v in resultado.items()}
    perda = {k: any(r < 0 for r in v) for k, v in resultado.items()}
    if (ganho[True] and perda[False]) or (ganho[False] and perda[True]):
        avisos.append(
            f"{ano}: há ganho de um tipo de fundo e perda do outro (FII e Fiagro). O COIR os apura "
            "juntos, como a ficha do programa IRPF, mas só a compensação de FII com FII tem "
            "norma (IN RFB 1.585/2015 art. 37 §2º): confira se a perda de um reduziu o imposto "
            "do outro e, na dúvida, recolha sem essa compensação"
        )
    rendimentos = [p for p in proventos if p.titular == titular and p.ano_declaracao == ano
                   and p.tipo is TipoProvento.RENDIMENTO_FII]  # fmt: skip
    for p in rendimentos:
        if p.codigo in fiagro and _vigente(params, "fiagro_rendimento_isento", p.data) == "false":
            avisos.append(
                f"{p.data:%d/%m/%Y}: rendimento do Fiagro {p.codigo} ({reais(p.liquido)}) antes da "
                "isenção, que vale desde 11/06/2021 (Lei 11.033/2004 art. 3º III, red. Lei "
                "14.130/2021): não é isento; declare como rendimento tributável"
            )
    if any(p.liquido > ZERO for p in rendimentos):
        minimo = _vigente(params, "fii_isencao_minimo_cotistas", date(ano, 12, 31))
        if minimo is not None:
            avisos.append(
                f"Rendimentos de FII e Fiagro de {ano} declarados como isentos: a isenção exige "
                f"cotas negociadas só em bolsa ou balcão organizado, fundo com {minimo} cotistas "
                "ou mais e titular com menos de 10% das cotas e dos rendimentos"
                + (" (e, desde 2024, o grupo de pessoas ligadas com menos de 30%)"
                   if minimo != "50" else "")
                + " (Lei 11.033/2004 art. 3º). Confira no informe do fundo; sem os requisitos, o "
                "rendimento é tributado na fonte a 20% e vai para a tributação exclusiva"
            )  # fmt: skip
    return avisos
