"""CPF e nome do cliente no cabeçalho da nota (todos os layouts SINACOR os imprimem).

O CPF vem com máscara, com traço ou ponto antes dos dígitos verificadores (000.000.000-00 ou
000.000.000.00). O nome fica na mesma linha do CPF (Ágora, BTG) ou na primeira linha depois do
rodapé de atendimento ("... Ouvidoria ..."), às vezes precedido do código do cliente (XP, Rico)
e do rótulo "Cliente" numa linha só dele. No layout BM&F novo da XP o nome vem antes da
ouvidoria, logo depois da linha do site ("Internet: ...").
"""

import re

from coir.domain.cpf import so_digitos
from coir.domain.nota import Cliente

RE_CPF = re.compile(r"(?<![\d./])\d{3}\.\d{3}\.\d{3}[-.]\d{2}(?![\d])")
RE_CODIGO = re.compile(r"^\s*\d[\d-]*\s+")
ROTULOS = {"cliente", "corretora"}


def cliente(texto: str) -> Cliente | None:
    """Cliente da página, ou None se nenhum CPF estiver impresso."""
    linhas = [linha for linha in texto.splitlines() if linha.strip()]
    for i, linha in enumerate(linhas):
        m = RE_CPF.search(linha)
        if m is None:
            continue
        nome = _nome(linha[: m.start()])
        if nome is None:
            nome = _depois_do_rotulo(linhas, i, "ouvidoria") or _depois_do_rotulo(
                linhas, i, "internet"
            )
        return Cliente(cpf=so_digitos(m.group(0)), nome=nome)
    return None


def _depois_do_rotulo(linhas: list[str], cpf: int, rotulo: str) -> str | None:
    """Nome na primeira linha depois da última que contém `rotulo`, pulando rótulos soltos."""
    antes = [j for j in range(cpf) if rotulo in linhas[j].lower()]
    if not antes:
        return None
    for linha in linhas[antes[-1] + 1 : cpf]:
        if linha.strip().lower() not in ROTULOS:
            return _nome(linha)
    return None


def _nome(trecho: str) -> str | None:
    nome = " ".join(RE_CODIGO.sub("", trecho).split())
    if re.search(r"\d", nome) or nome.lower() in ROTULOS:
        return None
    return nome if re.search(r"[A-Za-zÀ-ÿ]{2}", nome) else None
