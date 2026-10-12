"""Abas da planilha-gabarito → registros pseudonimizados.

Funções puras sobre listas de linhas (o I/O fica em ``leitor``), testáveis com linhas sintéticas.
"""

import hashlib
from collections.abc import Callable, Iterator, Sequence
from datetime import date
from decimal import Decimal
from functools import partial
from typing import Literal

from coir.domain.cpf import cpf_valido, so_digitos
from coir.erros import GabaritoError
from coir.gabarito import normalizar as nz
from coir.gabarito.leitor import Linhas
from coir.gabarito.modelo import (
    ApuracaoMensalGabarito,
    Categoria,
    CustodiaInicialGabarito,
    DarfGabarito,
    EventoGabarito,
    Gabarito,
    IsentoMensal,
    OperacaoGabarito,
    PosicaoGabarito,
)
from coir.gabarito.pseudonimos import MapaPseudonimos

ABAS_DADOS = {Categoria.COMUM: "Normais_Dados", Categoria.DAY_TRADE: "DayTrade_Dados"}
ABAS_OP = {Categoria.COMUM: "Normais_Op", Categoria.DAY_TRADE: "DayTrade_Op"}
ABAS_CONT = {Categoria.COMUM: "Normais_Cont", Categoria.DAY_TRADE: "DayTrade_Cont"}
ABAS = (
    *ABAS_DADOS.values(),
    *ABAS_OP.values(),
    *ABAS_CONT.values(),
    "DARF",
    "Isentos",
    "Eventos_B3",
    "Custodia_Inicial",
    "BTC_Outros",
)
MESES = (
    "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
    "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro",
)  # fmt: skip
CABECALHO_DADOS = ("Corretora", "CPF", "Conta", "Nota", "Data", "C/V", "Papel", "Operacao")
LINHAS_POR_ANO = 16


def _hash_curto(texto: str) -> str:
    return hashlib.sha256(texto.encode()).hexdigest()[:6]


def _indices(cabecalho: Sequence[object], exigidos: Sequence[str], aba: str) -> dict[str, int]:
    ix = {str(v).strip(): i for i, v in enumerate(cabecalho) if isinstance(v, str) and v.strip()}
    faltando = [c for c in exigidos if c not in ix]
    if faltando:
        raise GabaritoError(f"Aba {aba}: colunas ausentes no cabeçalho: {faltando}")
    return ix


def _celula(linha: Sequence[object], i: int) -> object:
    return linha[i] if i < len(linha) else None


def _linha_vazia(linha: Sequence[object]) -> bool:
    return all(nz.vazio(v) for v in linha)


def _por_nome(linha: Sequence[object], ix: dict[str, int], nome: str) -> object:
    return _celula(linha, ix[nome])


def _inteiro(valor: object, onde: str) -> int:
    if isinstance(valor, float) and valor.is_integer():
        return int(valor)
    raise GabaritoError(f"{onde}: número inteiro esperado, encontrado {valor!r}")


def _cv(valor: object, onde: str) -> Literal["C", "V"]:
    texto = str(valor).strip().upper() if not nz.vazio(valor) else ""
    if texto == "C":
        return "C"
    if texto == "V":
        return "V"
    raise GabaritoError(f"{onde}: C/V inválido: {valor!r}")


class Extrator:
    """Mantém o mapa de pseudônimos e os avisos de uma extração."""

    def __init__(self, mapa: MapaPseudonimos, titular_unico: bool = False) -> None:
        self.mapa = mapa
        self.titular_unico = titular_unico
        self.avisos: list[str] = []
        self._cpfs_validos: set[str] = set()
        self._cpfs_unificados: set[str] = set()

    # ------------------------------------------------------------------ titulares
    def registrar_cpfs(self, abas: dict[str, Linhas]) -> None:
        """Coleta os CPFs válidos (para corrigir erros de digitação nos demais)."""
        for aba, inicio in (
            (ABAS_DADOS[Categoria.COMUM], 0),
            (ABAS_DADOS[Categoria.DAY_TRADE], 0),
            ("Eventos_B3", 1),
        ):
            linhas = abas[aba]
            ix = _indices(linhas[inicio], ("CPF",), aba)
            for linha in linhas[inicio + 1 :]:
                cpf = _celula(linha, ix["CPF"])
                if not nz.vazio(cpf) and cpf_valido(str(cpf)):
                    self._cpfs_validos.add(so_digitos(str(cpf)))

    def _titular(self, cpf: object, onde: str) -> str:
        if nz.vazio(cpf):
            raise GabaritoError(f"{onde}: CPF vazio")
        digitos, corrigido = nz.corrigir_cpf(str(cpf), self._cpfs_validos)
        if self.titular_unico:
            # planilha de uma pessoa só com CPFs mistos (ex.: real e fictício): um titular
            pseudonimo = self.mapa.titulares.setdefault(digitos, "TITULAR-1")
            if pseudonimo != "TITULAR-1":
                raise GabaritoError("titular único pedido, mas o mapa já separa os CPFs")
            self._cpfs_unificados.add(digitos)
            return pseudonimo
        pseudonimo = self.mapa.titular(digitos)
        if corrigido:
            self.avisos.append(
                f"{onde}: CPF com dígito verificador inválido "
                f"(hash {_hash_curto(so_digitos(str(cpf)))}) tratado como {pseudonimo} "
                "(erro de digitação)."
            )
        return pseudonimo

    def _conta(self, corretora: str, conta: object, onde: str) -> str:
        texto = nz.texto_inteiro(conta)
        if texto is None:
            raise GabaritoError(f"{onde}: conta vazia")
        return self.mapa.conta(corretora, texto)

    # ------------------------------------------------------------------ operações
    def dados(
        self, linhas: Linhas, categoria: Categoria
    ) -> tuple[list[OperacaoGabarito], list[EventoGabarito]]:
        aba = ABAS_DADOS[categoria]
        ix = _indices(
            linhas[0],
            (*CABECALHO_DADOS, "Preço", "Quantidade", "Total", "Custos", "PM", "IRRF"),
            aba,
        )
        operacoes: list[OperacaoGabarito] = []
        eventos: list[EventoGabarito] = []
        for n, linha in enumerate(linhas[1:], start=2):
            if _linha_vazia(linha):
                continue
            onde = f"{aba} linha {n}"
            c = partial(_por_nome, linha, ix)
            corretora = nz.corretora(c("Corretora"))
            comum = {
                "categoria": categoria,
                "linha": n,
                "titular": self._titular(c("CPF"), onde),
                "corretora": corretora,
                "conta": self._conta(corretora, c("Conta"), onde),
                "data": nz.exigir_data(c("Data"), f"{onde} Data"),
                "cv": _cv(c("C/V"), onde),
                "papel": str(c("Papel")).strip().upper(),
            }
            nota = nz.texto_inteiro(c("Nota"))
            mercado = str(c("Operacao")).strip()
            if nota is None:
                eventos.append(
                    EventoGabarito(
                        **comum,
                        tipo=mercado,
                        preco=nz.para_decimal(c("Preço")),
                        quantidade=nz.para_decimal(c("Quantidade")),
                        total=nz.para_decimal(c("Total")),
                    )
                )
                continue
            operacoes.append(
                OperacaoGabarito(
                    **comum,
                    nota=nota,
                    mercado=mercado.upper(),
                    preco=nz.exigir_decimal(c("Preço"), f"{onde} Preço"),
                    quantidade=nz.exigir_decimal(c("Quantidade"), f"{onde} Quantidade"),
                    total=nz.exigir_decimal(c("Total"), f"{onde} Total"),
                    custos=nz.para_decimal(c("Custos")) or Decimal(0),
                    pm=nz.para_decimal(c("PM")),
                    irrf=nz.para_decimal(c("IRRF")) or Decimal(0),
                    prazo=nz.data_excel(c("Prazo")) if "Prazo" in ix else None,
                    ativo_objeto=(
                        str(c("Exercício")).strip().upper()
                        if "Exercício" in ix and not nz.vazio(c("Exercício"))
                        else None
                    ),
                )
            )
        return operacoes, eventos

    def eventos_cadastro(self, linhas: Linhas) -> list[EventoGabarito]:
        """Aba Eventos_B3 (cadastro de eventos; cabeçalho na 2ª linha)."""
        ix = _indices(linhas[1], (*CABECALHO_DADOS, "Preço", "Quantidade", "Total"), "Eventos_B3")
        eventos: list[EventoGabarito] = []
        for n, linha in enumerate(linhas[2:], start=3):
            if nz.vazio(_celula(linha, ix["Corretora"])):
                continue
            onde = f"Eventos_B3 linha {n}"
            c = partial(_por_nome, linha, ix)
            corretora = nz.corretora(c("Corretora"))
            eventos.append(
                EventoGabarito(
                    categoria=Categoria.COMUM,
                    linha=n,
                    titular=self._titular(c("CPF"), onde),
                    corretora=corretora,
                    conta=self._conta(corretora, c("Conta"), onde),
                    data=nz.exigir_data(c("Data"), f"{onde} Data"),
                    cv=_cv(c("C/V"), onde),
                    papel=str(c("Papel")).strip().upper(),
                    tipo=str(c("Operacao")).strip(),
                    preco=nz.para_decimal(c("Preço")),
                    quantidade=nz.para_decimal(c("Quantidade")),
                    total=nz.para_decimal(c("Total")),
                )
            )
        return eventos

    # ------------------------------------------------------------------ posições
    def posicoes(self, linhas: Linhas, categoria: Categoria) -> list[PosicaoGabarito]:
        aba = ABAS_OP[categoria]
        exigidos = (
            "Tipo",
            "Papel",
            "C/V",
            "Data",
            "Qtde",
            "PM",
            "PMA",
            "QtdeA",
            "Nota",
            "Custos",
            "IRRF",
            "Total",
            "Mês",
            "Ano",
            "Isento",
        )
        ix = _indices(linhas[0], (*exigidos, "Corretora", "Conta"), aba)
        mercado_col = "Mercado" if "Mercado" in ix else "Tipo"
        saida: list[PosicaoGabarito] = []
        for n, linha in enumerate(linhas[1:], start=2):
            c = partial(_por_nome, linha, ix)
            if nz.vazio(c("Tipo")):
                continue
            onde = f"{aba} linha {n}"
            corretora = nz.corretora(c("Corretora"))
            saida.append(
                PosicaoGabarito(
                    categoria=categoria,
                    ordem=len(saida) + 1,
                    corretora=corretora,
                    conta=self._conta(corretora, c("Conta"), onde),
                    papel=str(c("Papel")).strip().upper(),
                    mercado=str(c(mercado_col)).strip().upper(),
                    cv=_cv(c("C/V"), onde),
                    data=nz.exigir_data(c("Data"), f"{onde} Data"),
                    nota=nz.texto_inteiro(c("Nota")),
                    quantidade=nz.exigir_decimal(c("Qtde"), f"{onde} Qtde"),
                    pm_operacao=nz.para_decimal(c("PM")),
                    pm_apos=nz.para_decimal(c("PMA")),
                    quantidade_apos=nz.para_decimal(c("QtdeA")),
                    custos=nz.para_decimal(c("Custos")),
                    irrf=nz.para_decimal(c("IRRF")),
                    resultado=nz.para_decimal(c("Total")),
                    ano_competencia=_inteiro(c("Ano"), f"{onde} Ano"),
                    mes_competencia=_inteiro(c("Mês"), f"{onde} Mês"),
                    isento=str(c("Isento")).strip().upper() == "SIM",
                )
            )
        return saida


# ---------------------------------------------------------------------- blocos anuais
def blocos_anuais(
    linhas: Linhas, eh_cabecalho_ano: Callable[[Sequence[object]], int | None]
) -> Iterator[tuple[int, int, Sequence[object]]]:
    """Percorre blocos "ano → 12 meses" e devolve (ano, mês, linha) para cada linha de mês."""
    ano: int | None = None
    for linha in linhas:
        achado = eh_cabecalho_ano(linha)
        if achado is not None:
            ano = achado
            continue
        nome = _celula(linha, 2)
        if ano is not None and isinstance(nome, str) and nome.strip() in MESES:
            yield ano, MESES.index(nome.strip()) + 1, linha


def _ano(valor: object) -> int | None:
    if isinstance(valor, float) and valor.is_integer() and 2000 <= valor <= 2100:
        return int(valor)
    return None


def _cabecalho_cont(linha: Sequence[object]) -> int | None:
    return _ano(_celula(linha, 3)) if nz.vazio(_celula(linha, 2)) else None


def _cabecalho_darf(linha: Sequence[object]) -> int | None:
    if str(_celula(linha, 3)).strip() == "DARF":
        return _ano(_celula(linha, 4))
    return None


def apuracao(linhas: Linhas, categoria: Categoria, ate_ano: int) -> list[ApuracaoMensalGabarito]:
    """Abas *_Cont: Resultado, À Vista, Opções, Índice, Dólar, Commodity, IRPF, IRRF, A pagar,
    Acumulado (prejuízo a compensar) e Base de cálculo (vendas)."""
    campos = (
        "resultado", "a_vista", "opcoes", "indice", "dolar", "commodity",
        "irpf", "irrf", "a_pagar", "acumulado", "base_calculo",
    )  # fmt: skip
    saida: list[ApuracaoMensalGabarito] = []
    for ano, mes, linha in blocos_anuais(linhas, _cabecalho_cont):
        if ano > ate_ano:
            continue
        valores = {
            campo: nz.para_decimal(_celula(linha, 3 + i)) or Decimal(0)
            for i, campo in enumerate(campos)
        }
        saida.append(ApuracaoMensalGabarito(categoria=categoria, ano=ano, mes=mes, **valores))
    if not saida:
        raise GabaritoError(f"Aba {ABAS_CONT[categoria]}: nenhum bloco anual reconhecido")
    return saida


def darfs(linhas: Linhas) -> list[DarfGabarito]:
    """Aba DARF: Apuração, Vencimento, Pagamento, Código, Valor, Multa, Juros e Total.

    Só os meses com valor (DARFs efetivamente lançados)."""
    saida: list[DarfGabarito] = []
    for ano, mes, linha in blocos_anuais(linhas, _cabecalho_darf):
        valor = nz.para_decimal(_celula(linha, 7))
        if not valor:
            continue
        codigo = nz.texto_inteiro(_celula(linha, 6))
        saida.append(
            DarfGabarito(
                ano=ano,
                mes=mes,
                apuracao=_data_ou_none(_celula(linha, 3)),
                vencimento=_data_ou_none(_celula(linha, 4)),
                pagamento=_data_ou_none(_celula(linha, 5)),
                codigo=codigo,
                valor=valor,
                multa=nz.para_decimal(_celula(linha, 8)),
                juros=nz.para_decimal(_celula(linha, 9)),
                total=nz.para_decimal(_celula(linha, 10)),
            )
        )
    return saida


def _data_ou_none(valor: object) -> date | None:
    return (
        nz.data_excel(valor)
        if isinstance(valor, int | float) and not isinstance(valor, bool)
        else None
    )


def isentos(linhas: Linhas, ate_ano: int) -> list[IsentoMensal]:
    """Aba Isentos: anos nas colunas (linha de cabeçalho), meses nas linhas (coluna 1)."""
    colunas: dict[int, int] = {}
    saida: list[IsentoMensal] = []
    for linha in linhas:
        anos = {i: a for i, v in enumerate(linha) if (a := _ano(v)) is not None}
        if len(anos) >= 5 and not colunas:
            colunas = anos
            continue
        nome = _celula(linha, 1)
        if colunas and isinstance(nome, str) and nome.strip() in MESES:
            mes = MESES.index(nome.strip()) + 1
            for i, ano in colunas.items():
                if ano <= ate_ano:
                    saida.append(
                        IsentoMensal(
                            ano=ano, mes=mes, valor=nz.para_decimal(_celula(linha, i)) or Decimal(0)
                        )
                    )
    if not colunas:
        raise GabaritoError("Aba Isentos: linha de anos não encontrada")
    return sorted(saida, key=lambda r: (r.ano, r.mes))


def custodia_inicial(linhas: Linhas) -> list[CustodiaInicialGabarito]:
    exigidos = (
        "Corretora",
        "Data",
        "C/V",
        "Código_Papel",
        "Operacao",
        "Preço_Médio",
        "Quantidade",
        "Total",
    )
    ix = _indices(linhas[0], exigidos, "Custodia_Inicial")
    saida: list[CustodiaInicialGabarito] = []
    for n, linha in enumerate(linhas[1:], start=2):
        c = partial(_por_nome, linha, ix)
        if nz.vazio(c("Corretora")) or nz.vazio(c("Código_Papel")):
            continue
        onde = f"Custodia_Inicial linha {n}"
        prejuizo = ix.get("Prejuízo Acumulado")
        saida.append(
            CustodiaInicialGabarito(
                linha=n,
                corretora=nz.corretora(c("Corretora")),
                data=nz.exigir_data(c("Data"), f"{onde} Data"),
                cv=str(c("C/V")).strip().upper(),
                papel=str(c("Código_Papel")).strip().upper(),
                operacao=str(c("Operacao")).strip(),
                preco_medio=nz.exigir_decimal(c("Preço_Médio"), f"{onde} Preço_Médio"),
                quantidade=nz.exigir_decimal(c("Quantidade"), f"{onde} Quantidade"),
                total=nz.para_decimal(c("Total")),
                prejuizo_acumulado=nz.para_decimal(_celula(linha, prejuizo))
                if prejuizo is not None
                else None,
            )
        )
    return saida


def btc_outros_tem_valores(linhas: Linhas) -> bool:
    """Indica se a aba BTC_Outros tem algum valor numérico não nulo nos meses."""
    for *_, linha in blocos_anuais(linhas, _cabecalho_darf_ou_outros):
        if any(isinstance(v, int | float) and v for v in linha[3:]):
            return True
    return False


def _cabecalho_darf_ou_outros(linha: Sequence[object]) -> int | None:
    if str(_celula(linha, 3)).strip() == "Outros Custos":
        return _ano(_celula(linha, 4))
    return None


# ---------------------------------------------------------------------- extração completa
def extrair(
    abas: dict[str, Linhas], mapa: MapaPseudonimos, titular_unico: bool = False
) -> Gabarito:
    ext = Extrator(mapa, titular_unico)
    ext.registrar_cpfs(abas)
    operacoes: list[OperacaoGabarito] = []
    eventos: list[EventoGabarito] = []
    for categoria, aba in ABAS_DADOS.items():
        ops, evs = ext.dados(abas[aba], categoria)
        operacoes += ops
        eventos += evs
    if not operacoes:
        raise GabaritoError("Nenhuma operação encontrada nas abas *_Dados")
    ate_ano = max(op.data.year for op in operacoes)
    posicoes = [p for categoria, aba in ABAS_OP.items() for p in ext.posicoes(abas[aba], categoria)]
    cadastro = ext.eventos_cadastro(abas["Eventos_B3"])
    if len(cadastro) != len(eventos):
        ext.avisos.append(
            f"Eventos: {len(eventos)} lançados nas abas *_Dados e {len(cadastro)} na aba "
            "Eventos_B3 (conferir se algum evento cadastrado não foi lançado, ou vice-versa)."
        )
    if not btc_outros_tem_valores(abas["BTC_Outros"]):
        ext.avisos.append(
            "BTC_Outros: nenhum valor mensal encontrado (outros custos dedutíveis ficam vazios)."
        )
    custodia = custodia_inicial(abas["Custodia_Inicial"])
    if len(ext._cpfs_unificados) > 1:
        ext.avisos.append(
            f"Titular único: {len(ext._cpfs_unificados)} CPFs diferentes tratados como TITULAR-1."
        )
    return Gabarito(
        operacoes=operacoes,
        eventos=eventos,
        eventos_cadastro=cadastro,
        posicoes=posicoes,
        apuracao=[
            a
            for categoria, aba in ABAS_CONT.items()
            for a in apuracao(abas[aba], categoria, ate_ano)
        ],
        darfs=darfs(abas["DARF"]),
        isentos=isentos(abas["Isentos"], ate_ano),
        custodia_inicial=custodia,
        avisos=ext.avisos,
    )
