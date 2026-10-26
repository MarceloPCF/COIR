"""CPF e nome do cliente no cabeçalho da nota (todos os layouts SINACOR os imprimem).

O CPF vem com máscara, com traço ou ponto antes dos dígitos verificadores (000.000.000-00 ou
000.000.000.00). O nome fica na mesma linha do CPF (Ágora, BTG) ou na primeira linha depois do
rodapé de atendimento ("... Ouvidoria ..."), às vezes precedido do código do cliente (XP, Rico).
"""

import re

from coir.domain.cpf import so_digitos
from coir.domain.nota import Cliente

RE_CPF = re.compile(r"(?<![\d./])\d{3}\.\d{3}\.\d{3}[-.]\d{2}(?![\d])")
RE_CODIGO = re.compile(r"^\s*\d[\d-]*\s+")


def cliente(texto: str) -> Cliente | None:
    """Cliente da página, ou None se nenhum CPF estiver impresso."""
    linhas = [linha for linha in texto.splitlines() if linha.strip()]
    for i, linha in enumerate(linhas):
        m = RE_CPF.search(linha)
        if m is None:
            continue
        nome = _nome(linha[: m.start()])
        if nome is None:
            rodape = [j for j in range(i) if "ouvidoria" in linhas[j].lower()]
            if rodape and rodape[-1] + 1 < i:
                nome = _nome(linhas[rodape[-1] + 1])
        return Cliente(cpf=so_digitos(m.group(0)), nome=nome)
    return None


def _nome(trecho: str) -> str | None:
    nome = " ".join(RE_CODIGO.sub("", trecho).split())
    return nome if re.search(r"[A-Za-zÀ-ÿ]{2}", nome) else None
