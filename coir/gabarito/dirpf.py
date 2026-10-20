"""Leitura da declaração entregue (DIRPF), para comparar com as fichas que o COIR gera (F8).

Entrada: o PDF "Imprimir > Declaração" do programa IRPF, em ``COIR_DADOS/_dirpf/`` (fora do
repositório). Só o que o COIR calcula é lido: a ficha "Renda variável - operações comuns/day
trade" mês a mês, a de fundos imobiliários/Fiagro, Bens e Direitos, os rendimentos isentos e
exclusivos por fonte pagadora e o IR fonte da Lei 11.033/2004 de "Imposto pago/retido". Nome e
CPF do cabeçalho de cada página não são guardados.

Layout fora do esperado (seção que some, linha de mercado desconhecida) falha alto: a comparação
nunca roda sobre uma leitura pela metade.
"""

import re
from collections.abc import Iterator, Sequence
from decimal import Decimal
from pathlib import Path

from pydantic import BaseModel, ConfigDict

from coir.domain.cpf import so_digitos
from coir.erros import GabaritoError
from coir.parsing.texto_pdf import paginas

MESES = ("JAN", "FEV", "MAR", "ABR", "MAI", "JUN", "JUL", "AGO", "SET", "OUT", "NOV", "DEZ")
RE_VALOR = r"-?(?:\d{1,3}(?:\.\d{3})+|\d+),\d{2}"
RE_CNPJ = re.compile(r"\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}")
RE_EXERCICIO = re.compile(r"EXERCÍCIO (\d{4}) ANO-CALENDÁRIO (\d{4})")
RE_LINHA_VALORES = re.compile(
    rf"^(?P<rotulo>.+?) (?P<v1>{RE_VALOR}|\d+%)(?: (?P<v2>{RE_VALOR}|\d+%))?$"
)
RE_MES = re.compile(r"^GANHOS LÍQUIDOS OU PERDAS - (?P<mes>[A-Z]{3})$")
RE_BEM_VALORES = re.compile(rf"^(?P<anterior>{RE_VALOR}) (?P<atual>{RE_VALOR})$")
RE_BEM_CODIGO = re.compile(r"^(?P<grupo>\d{2}) (?P<codigo>\d{2}) (?P<texto>.*)$")
RE_NEGOCIACAO = re.compile(r"Código de Negociação: (?P<ticker>[A-Z0-9]+)")
RE_TOKEN_RENDIMENTO = re.compile(
    rf"(?P<cnpj>\d{{2}}\.\d{{3}}\.\d{{3}}/\d{{4}}-\d{{2}})|(?<![\d.,/])(?P<valor>{RE_VALOR})(?![\d,])"
)
RE_RENDIMENTO = re.compile(rf"^(?P<codigo>\d{{2}})\. (?P<nome>.+?)(?: (?P<total>{RE_VALOR}))?$")

# cabeçalho repetido em cada página (nome e CPF do titular ficam nas linhas 4 e 6)
CABECALHO = 6
CABECALHO_BENS = ("DECLARAÇÃO DE BENS E DIREITOS (Valores em Reais)", "SITUAÇÃO EM",
                  "GRUPO CÓDIGO DISCRIMINAÇÃO")  # fmt: skip
RV_TITULAR = "RENDA VARIÁVEL - OPERAÇÕES COMUNS/DAYTRADE - TITULAR"
RV_FIM = "RENDA VARIÁVEL - OPERAÇÕES COMUNS/DAYTRADE - DEPENDENTES"
FII = "FUNDOS DE INVESTIMENTO IMOBILIÁRIO OU NAS CADEIAS PRODUTIVAS AGROINDUSTRIAIS"
FII_TITULAR = f"{FII} - TITULAR"
FII_FIM = f"{FII} - DEPENDENTES"
BENS = "DECLARAÇÃO DE BENS E DIREITOS (Valores em Reais)"
ISENTOS = "RENDIMENTOS ISENTOS E NÃO TRIBUTÁVEIS (Valores em Reais)"
EXCLUSIVOS = "RENDIMENTOS SUJEITOS À TRIBUTAÇÃO EXCLUSIVA / DEFINITIVA (Valores em Reais)"
IMPOSTO_PAGO = "IMPOSTO PAGO / RETIDO (Valores em Reais)"
SEM_INFORMACOES = "Sem Informações"


class Modelo(BaseModel):
    model_config = ConfigDict(frozen=True, extra="forbid")


class MesFicha(Modelo):
    """Um mês de uma ficha de renda variável: rótulo impresso -> (comuns, day trade). As linhas
    de uma coluna só (consolidação do mês) ficam com day trade None."""

    mes: int
    valores: dict[str, tuple[Decimal, Decimal | None]]

    def valor(self, rotulo: str, coluna: int = 0) -> Decimal:
        try:
            v = self.valores[rotulo][coluna]
        except KeyError as erro:
            raise GabaritoError(
                f"DIRPF: linha {rotulo!r} ausente em {MESES[self.mes - 1]}"
            ) from erro
        if v is None:
            raise GabaritoError(f"DIRPF: linha {rotulo!r} sem a coluna {coluna}")
        return v


class BemDeclarado(Modelo):
    grupo: str
    codigo: str
    ticker: str | None  # "Código de Negociação" (bens negociados em bolsa)
    cnpj: str | None  # só dígitos
    discriminacao: str
    anterior: Decimal  # situação em 31/12 do ano anterior
    atual: Decimal  # situação em 31/12 do ano-calendário


class Rendimento(Modelo):
    """Um rendimento isento ou exclusivo por fonte pagadora (quando a leitura por fonte fecha)."""

    ficha: str  # "isentos" ou "exclusivos"
    codigo: str  # "09", "10", ...
    cnpj: str  # só dígitos
    valor: Decimal


class Dirpf(Modelo):
    exercicio: int
    ano: int  # ano-calendário
    renda_variavel: tuple[MesFicha, ...]
    fii: tuple[MesFicha, ...]
    bens: tuple[BemDeclarado, ...]
    totais: dict[tuple[str, str], Decimal]  # (ficha, código) -> total da linha
    rendimentos: tuple[Rendimento, ...]
    sem_fonte: tuple[tuple[str, str], ...]  # linhas cuja leitura por fonte não fechou
    ir_fonte_lei_11033: Decimal  # "Imposto pago/retido" 03


def _valor(texto: str) -> Decimal:
    if texto.endswith("%"):
        return Decimal(texto[:-1]) / 100
    return Decimal(texto.replace(".", "").replace(",", "."))


def _linhas(textos: Sequence[str]) -> tuple[int, int, list[str]]:
    """Linhas de todas as páginas, sem o cabeçalho (nome, CPF) nem o rodapé de cada página."""
    exercicio = ano = 0
    saida: list[str] = []
    for n, texto in enumerate(textos, 1):
        linhas = [x.strip() for x in texto.splitlines() if x.strip()]
        if len(linhas) < CABECALHO or linhas[0] != "NOME:":
            raise GabaritoError(f"DIRPF: página {n} sem o cabeçalho do programa IRPF")
        m = RE_EXERCICIO.search(linhas[2])
        if m is None:
            raise GabaritoError(f"DIRPF: página {n} sem exercício e ano-calendário")
        exercicio, ano = int(m[1]), int(m[2])
        saida += [x for x in linhas[CABECALHO:] if not x.startswith("Página ")]
    return exercicio, ano, saida


def _secao(linhas: list[str], inicio: str, fim: str | Sequence[str]) -> list[str]:
    try:
        i = linhas.index(inicio)
    except ValueError as erro:
        raise GabaritoError(f"DIRPF: seção {inicio!r} não encontrada") from erro
    fins = (fim,) if isinstance(fim, str) else tuple(fim)
    j = next((k for k in range(i + 1, len(linhas)) if linhas[k] in fins
              or any(linhas[k].startswith(f) for f in fins)), None)  # fmt: skip
    if j is None:
        raise GabaritoError(f"DIRPF: fim da seção {inicio!r} não encontrado")
    return linhas[i + 1 : j]


def _ficha(linhas: list[str], nome: str) -> tuple[MesFicha, ...]:
    if linhas[:1] == [SEM_INFORMACOES]:
        return ()
    meses: list[MesFicha] = []
    atual: dict[str, tuple[Decimal, Decimal | None]] | None = None
    mes = 0
    for linha in linhas:
        if m := RE_MES.match(linha):
            if atual is not None:
                meses.append(MesFicha(mes=mes, valores=atual))
            mes, atual = MESES.index(m["mes"]) + 1, {}
        elif atual is not None and (m := RE_LINHA_VALORES.match(linha)):
            v2 = _valor(m["v2"]) if m["v2"] else None
            atual[m["rotulo"]] = (_valor(m["v1"]), v2)
    if atual is not None:
        meses.append(MesFicha(mes=mes, valores=atual))
    if [m.mes for m in meses] != list(range(1, 13)):
        raise GabaritoError(f"DIRPF: ficha {nome} sem os 12 meses")
    return tuple(meses)


def _bens(linhas: list[str]) -> tuple[BemDeclarado, ...]:
    corpo = [x for x in linhas if x not in CABECALHO_BENS and not re.fullmatch(
        r"31/12/\d{4} 31/12/\d{4}", x)]  # fmt: skip
    bens: list[BemDeclarado] = []
    atual: dict[str, object] | None = None
    texto: list[str] = []
    for linha in corpo:
        if linha.startswith("TOTAL "):
            break
        if atual is None:
            if m := RE_BEM_VALORES.match(linha):
                atual, texto = {"anterior": _valor(m["anterior"]), "atual": _valor(m["atual"])}, []
            continue
        if (fim := linha.find("Bem com usufruto")) >= 0:  # às vezes colado na discriminação
            if "grupo" not in atual:
                raise GabaritoError("DIRPF: bem sem grupo e código")
            if linha[:fim].strip():
                texto.append(linha[:fim].strip())
            bens.append(BemDeclarado(discriminacao=" ".join(texto), **atual))  # type: ignore[arg-type]
            atual = None
        elif "grupo" in atual:
            texto.append(linha)
        elif linha.startswith("CNPJ: ") and "cnpj" not in atual:
            atual["cnpj"] = so_digitos(linha[6:])
        elif m := RE_NEGOCIACAO.search(linha):
            atual["ticker"] = m["ticker"]
        elif m := RE_BEM_CODIGO.match(linha):
            atual |= {"grupo": m["grupo"], "codigo": m["codigo"]}
            atual.setdefault("ticker", None)
            atual.setdefault("cnpj", None)
            texto = [m["texto"]]
    if atual is not None:
        raise GabaritoError("DIRPF: Bens e Direitos terminou no meio de um bem")
    return tuple(bens)


def _blocos(linhas: list[str]) -> Iterator[tuple[str, Decimal | None, list[str]]]:
    """Linhas numeradas ("09. Lucros e dividendos recebidos 22.727,56") e o que vem abaixo. Rótulo
    longo quebra a linha: o total vem sozinho numa linha antes do cabeçalho das fontes."""
    codigo, total, corpo = "", None, []
    for linha in linhas:
        if (m := RE_RENDIMENTO.match(linha)) and not RE_CNPJ.search(linha):
            if codigo:
                yield codigo, total, corpo
            codigo, total, corpo = m["codigo"], _valor(m["total"]) if m["total"] else None, []
        elif codigo:
            cabecalho = any(x.startswith(("Beneficiário", "CPF/CNPJ")) for x in corpo)
            if total is None and not cabecalho and not corpo_tem_fonte(corpo) and re.fullmatch(
                    RE_VALOR, linha):  # fmt: skip
                total = _valor(linha)
            else:
                corpo.append(linha)
    if codigo:
        yield codigo, total, corpo


def corpo_tem_fonte(corpo: list[str]) -> bool:
    return any(RE_CNPJ.search(x) or x.startswith("Titular ") for x in corpo)


def _rendimentos(
    ficha: str, linhas: list[str], totais: dict[tuple[str, str], Decimal],
    por_fonte: list[Rendimento], sem_fonte: list[tuple[str, str]],
) -> None:  # fmt: skip
    """Cada fonte tem um CNPJ e um valor; o nome quebra em várias linhas, mas o valor vem depois
    do CNPJ e antes do CNPJ da fonte seguinte. Fonte com zero ou dois valores nesse trecho é
    ambígua: não entra por fonte (a linha vai para ``sem_fonte``) e nada é chutado. A linha 99
    ("Outros") não imprime total ao lado do rótulo: vale a soma das fontes."""
    for codigo, total, corpo in _blocos(linhas):
        texto = "\n".join(x for x in corpo if not x.startswith(("Beneficiário", "TOTAL ")))
        fontes: list[tuple[str, list[Decimal]]] = []
        soltos: list[Decimal] = []  # valores antes do primeiro CNPJ
        for m in RE_TOKEN_RENDIMENTO.finditer(texto):
            if m["cnpj"]:
                fontes.append((so_digitos(m["cnpj"]), []))
            else:
                (fontes[-1][1] if fontes else soltos).append(_valor(m["valor"]))
        if not fontes and not any(soltos):
            if total is not None:
                totais[(ficha, codigo)] = total  # linha sem fonte detalhada (ex.: 25)
            continue
        certas = [(c, v[0]) for c, v in fontes if len(v) == 1]
        soma = sum((v for _, v in certas), Decimal(0))
        if total is None:
            total = soma + sum(soltos, Decimal(0)) + sum((sum(v, Decimal(0)) for _, v in fontes
                                                         if len(v) != 1), Decimal(0))  # fmt: skip
        totais[(ficha, codigo)] = total
        por_fonte += [Rendimento(ficha=ficha, codigo=codigo, cnpj=c, valor=v) for c, v in certas]
        if soltos or len(certas) != len(fontes) or soma != total:
            sem_fonte.append((ficha, codigo))


def ler_texto(textos: Sequence[str]) -> Dirpf:
    exercicio, ano, linhas = _linhas(textos)
    totais: dict[tuple[str, str], Decimal] = {}
    por_fonte: list[Rendimento] = []
    sem_fonte: list[tuple[str, str]] = []
    _rendimentos("isentos", _secao(linhas, ISENTOS, EXCLUSIVOS), totais, por_fonte, sem_fonte)
    _rendimentos("exclusivos", _secao(linhas, EXCLUSIVOS, ("RENDIMENTOS TRIBUTÁVEIS",
                                                           IMPOSTO_PAGO)),
                 totais, por_fonte, sem_fonte)  # fmt: skip
    pago = _secao(linhas, IMPOSTO_PAGO, ("RENDIMENTOS TRIBUTÁVEIS", "PAGAMENTOS EFETUADOS"))
    return Dirpf(
        exercicio=exercicio, ano=ano,
        renda_variavel=_ficha(_secao(linhas, RV_TITULAR, RV_FIM), "renda variável"),
        fii=_ficha(_secao(linhas, FII_TITULAR, FII_FIM), "FII"),
        bens=_bens(_secao(linhas, BENS, ("DÍVIDAS E ÔNUS REAIS", "DOAÇÕES A PARTIDOS"))),
        totais=totais, rendimentos=tuple(por_fonte), sem_fonte=tuple(sem_fonte),
        ir_fonte_lei_11033=_ir_fonte(pago),
    )  # fmt: skip


def _ir_fonte(linhas: list[str]) -> Decimal:
    """Linha 03 de "Imposto pago/retido". As linhas 02 e 03 não trazem o valor ao lado do rótulo:
    o programa os imprime em seguida, um por linha, na ordem (02, 03, detalhes da 02)."""
    rotulo = "03. Imposto sobre a renda na fonte (Lei 11.033/2004)"
    i = next((k for k, x in enumerate(linhas) if x.startswith(rotulo)), None)
    if i is None or i == 0 or not linhas[i - 1].startswith("02. "):
        raise GabaritoError("DIRPF: Imposto pago/retido fora do layout esperado")
    valores: list[str] = []
    for x in linhas[i + 1 :]:
        if not re.fullmatch(RE_VALOR, x):
            break
        valores.append(x)
    if len(valores) < 2:
        raise GabaritoError("DIRPF: Imposto pago/retido sem o valor da linha 03")
    return _valor(valores[1])


def ler_dirpf(caminho: Path) -> Dirpf:
    return ler_texto(paginas(caminho))
