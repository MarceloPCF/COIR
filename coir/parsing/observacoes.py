"""Coluna "Obs. (*)" da linha de negócio Bovespa.

A coluna traz zero ou mais marcas independentes, coladas e em qualquer ordem ("D#@", "@D", "H#2"). O
significado vem da legenda impressa no rodapé da própria nota: as letras de lá valem para aquela
nota, além das conhecidas. Só "D" (day trade) muda o cálculo; o resto fica registrado.

A linha é lida como texto, então a marca é reconhecida pela forma do último pedaço da especificação.
Quem decide o que é especificação são os pedaços que a B3 usa nos nomes de pregão (COTAHIST, 2008 em
diante): "CI" e "DB" nunca são marca, embora tenham só letras da legenda. Se o nome não resolver,
``coir.parsing.normalizar`` ainda tenta o nome com o pedaço de volta e sem o pedaço desconhecido.
"""

import re

MARCAS_DA_LEGENDA = frozenset("#28ABCDFHILPTXY")  # legenda SINACOR do rodapé das notas
MARCAS_SEM_LEGENDA = frozenset("@4")  # impressas nas notas, ausentes da legenda
MARCA_DAY_TRADE = "D"
# as que não formam fim de nome de pregão: marca em qualquer vizinhança, mesmo depois de uma
# especificação colada, como o layout antigo da XP imprime ("ONNM HD")
MARCAS_FORTES = frozenset("HD248")
# marcas de operações que o COIR não trata à parte: entram como negócio comum, com aviso
MARCAS_COM_AVISO = {
    "A": "posição futuro",
    "B": "debêntures",
    "F": "cobertura",
    "L": "precatório",
    "X": "box",
    "Y": "desmanche de box",
}
# pedaços do fim do nome de pregão só com caracteres de marca (conferido no COTAHIST 2008-2026)
TOKENS_ESPECIFICACAO = frozenset({"CI", "CIA", "CPA", "DB", "F4", "FID", "FLI", "IBX", "PAB", "PAD",
                                  "PAF", "PAH", "PAL"})  # fmt: skip
# marca ou especificação, conforme o vizinho: é marca depois de classe ou sufixo ("ON NM C")
AMBIGUOS = frozenset("ABCP")
VIZINHOS_DE_MARCA = frozenset({"ON", "OR", "PN", "PNA", "PNB", "PNC", "PND", "PNE", "PNF", "PNG",
                               "UNT", "CI", "NM", "N1", "N2", "MA", "MB", "M2", "ED", "EJ", "ER",
                               "EX", "ES", "EB", "ATZ", "EDJ", "EDB", "EDR", "EJS", "DR1", "DR2",
                               "DR3", "DRE", "DRN", "F11"})  # fmt: skip
# pedaços com "D" que são da especificação, não marca (COTAHIST 2008-2026)
COM_D_NA_ESPECIFICACAO = frozenset({"BDR", "D71", "DB", "DEB", "DIR", "DR1", "DR2", "DR3", "DRE",
                                    "DRN", "ED", "EDB", "EDC", "EDG", "EDJ", "EDR", "EDS", "EJD",
                                    "FID"})  # fmt: skip
TAMANHO_SUSPEITO = 3  # pedaço curto, com "D", que não é especificação: pode ser marca nova
TAMANHO_MAXIMO = 6
RE_CODIGO = re.compile(r"^[A-Z0-9]{4}[A-X]?\d")  # ticker ou série de opção no começo da linha
RE_INICIO_LEGENDA = re.compile(r"\(\*\)\s*-?\s*Observa")
RE_LEGENDA = re.compile(r"(?:^|\s)([A-Z0-9#@])\s*-\s*[A-ZÀ-Ú]")


def legenda(texto: str) -> frozenset[str]:
    """Marcas explicadas no rodapé desta página ("D - Day Trade", "# - Negócio direto")."""
    achado = RE_INICIO_LEGENDA.search(texto)
    return frozenset(RE_LEGENDA.findall(texto[achado.end() :])) if achado else frozenset()


def _e_marca(token: str, anterior: str, marcas: frozenset[str]) -> bool:
    if token in TOKENS_ESPECIFICACAO or len(token) > TAMANHO_MAXIMO or token == "*":
        return False
    if token in AMBIGUOS:
        return anterior in VIZINHOS_DE_MARCA
    if "#" in token or "@" in token:  # símbolo de marca nunca faz parte de nome de pregão
        return True
    if not set(token) <= marcas | {"*"} or len(set(token)) != len(token):
        return False
    letras = token.replace("*", "")
    if len(letras) == 1 or set(letras) <= MARCAS_FORTES:
        return True
    # "ALLIANCE DAT": letras de marca no fim do nome; depois de classe ou de código, é marca
    return anterior in VIZINHOS_DE_MARCA or bool(RE_CODIGO.match(anterior))


def separar(
    tokens: list[str], da_nota: frozenset[str] = frozenset()
) -> tuple[list[str], str, bool]:
    """Pedaços do meio da linha -> (especificação, marcas, cotação por lote de mil). Lê do fim para
    o começo; o primeiro pedaço nunca é marca (é o nome ou o código do ativo)."""
    marcas = MARCAS_DA_LEGENDA | MARCAS_SEM_LEGENDA | da_nota
    resto = [t for t in tokens if t != "*"]  # "*" isolado: lote de mil, em qualquer posição
    obs = ""
    while len(resto) > 1 and _e_marca(resto[-1], resto[-2], marcas):
        obs = resto.pop() + obs
    return resto, obs.replace("*", ""), "*" in tokens or "*" in obs


def marca_suspeita(token: str) -> bool:
    """Pedaço curto com "D" que não é especificação conhecida: pode ser um day trade com marca nova.
    Quem o achar não pode seguir como operação comum em silêncio."""
    return (len(token) <= TAMANHO_SUSPEITO and MARCA_DAY_TRADE in token
            and token not in COM_D_NA_ESPECIFICACAO | VIZINHOS_DE_MARCA)  # fmt: skip
