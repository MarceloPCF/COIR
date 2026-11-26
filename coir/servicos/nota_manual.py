"""Nota de corretagem digitada à mão: a nota perdida ou de corretora sem nota SINACOR.

Entra no livro como as notas lidas do PDF (``NotaCorretagem``, mesmas operações, custos e IRRF) e
passa pelas mesmas regras: day trade, rateio dos custos, IRRF, preço médio. O resumo da nota
(valor das operações, compras, vendas, subtotais e líquido) sai das linhas digitadas, para a nota
conciliar por construção; a conciliação roda do mesmo jeito.

Cada nota manual tem o seu ``ArquivoImportado`` (sem PDF, tamanho zero, nome com ``PREFIXO``): é
assim que o COIR sabe que ela é manual, sem coluna nova no banco. A reimportação de PDFs nunca a
toca. Corrigir ou remover guarda a versão anterior em ``nota_anterior`` e faz backup antes (regra
7 do projeto). Uma nota do PDF gravada na mesma conta, dia e segmento de uma manual gera aviso e
pendência: a manual não some sozinha.
"""

import hashlib
from collections.abc import Sequence
from dataclasses import dataclass, field
from datetime import date, datetime
from decimal import ROUND_HALF_UP, Decimal
from pathlib import Path

from sqlalchemy import select
from sqlalchemy.orm import Session

from coir.db import repositorio as repo
from coir.db.modelos import ArquivoImportado, Conta, NotaCorretagem, Titular
from coir.db.sessao import conectar, fazer_backup
from coir.domain.fiscal import CENTAVO, ZERO
from coir.domain.nota import (
    CV,
    CategoriaCusto,
    Custo,
    GrupoCusto,
    Montante,
    NegocioBMF,
    NegocioBovespa,
    NotaExtraida,
    OperacaoExtraida,
    ResumoBMF,
    ResumoBovespa,
    Segmento,
)
from coir.domain.parametros import ReferenciasFiscais
from coir.domain.referencias import Referencias
from coir.erros import ExtracaoError, ImportacaoError
from coir.parsing.conciliar import conciliar
from coir.parsing.normalizar import classes_das_series, operacoes
from coir.servicos import trilha
from coir.servicos.eventos import conferir_no_motor, pendencia_antes

PREFIXO = repo.PREFIXO_NOTA_MANUAL

# mercado como a nota imprime (data/mercados_bovespa.csv) -> rótulo do formulário
MERCADOS_BOVESPA = {
    "VISTA": "À vista",
    "FRACIONARIO": "Fracionário",
    "OPCAO DE COMPRA": "Opção de compra",
    "OPCAO DE VENDA": "Opção de venda",
    "EXERC OPC COMPRA": "Exercício de opção de compra",
    "EXERC OPC VENDA": "Exercício de opção de venda",
    "TERMO": "Termo",
}
OPCOES = ("OPCAO DE COMPRA", "OPCAO DE VENDA")

# subtotal da nota SINACOR em que cada custo entra (só a conciliação e a tela usam o grupo)
GRUPO_BOVESPA: dict[CategoriaCusto, GrupoCusto] = {
    CategoriaCusto.LIQUIDACAO: "cblc",
    CategoriaCusto.REGISTRO: "cblc",
    CategoriaCusto.TERMO_OPCOES: "bolsa",
    CategoriaCusto.ANA: "bolsa",
    CategoriaCusto.EMOLUMENTOS: "bolsa",
    CategoriaCusto.TRANSFERENCIA: "bolsa",
    CategoriaCusto.CORRETAGEM: "despesas",
    CategoriaCusto.EXECUCAO: "despesas",
    CategoriaCusto.CUSTODIA: "despesas",
    CategoriaCusto.IMPOSTOS: "despesas",
    CategoriaCusto.OUTROS: "despesas",
}
# BM&F: os impostos (ISS, PIS, COFINS) ficam no campo próprio do resumo, fora das linhas de custo
CUSTOS_BMF = (
    CategoriaCusto.TAXAS_BMF,
    CategoriaCusto.CORRETAGEM,
    CategoriaCusto.REGISTRO,
    CategoriaCusto.EMOLUMENTOS,
    CategoriaCusto.OUTROS,
    CategoriaCusto.IMPOSTOS,
)
ROTULOS = {
    CategoriaCusto.LIQUIDACAO: "Taxa de liquidação",
    CategoriaCusto.REGISTRO: "Taxa de registro",
    CategoriaCusto.TERMO_OPCOES: "Taxa de termo/opções",
    CategoriaCusto.ANA: "Taxa A.N.A.",
    CategoriaCusto.EMOLUMENTOS: "Emolumentos",
    CategoriaCusto.TRANSFERENCIA: "Taxa de transferência de ativos",
    CategoriaCusto.CORRETAGEM: "Corretagem",
    CategoriaCusto.EXECUCAO: "Taxa de execução",
    CategoriaCusto.CUSTODIA: "Taxa de custódia",
    CategoriaCusto.IMPOSTOS: "Impostos (ISS)",
    CategoriaCusto.OUTROS: "Outros",
    CategoriaCusto.TAXAS_BMF: "Taxas BM&F (emol + f.gar)",
}


@dataclass(frozen=True)
class LinhaBovespa:
    """``prazo``: vencimento da opção ("MM/AA") ou prazo do termo em dias; vazio no à vista."""

    cv: CV
    mercado: str
    ativo: str
    quantidade: Decimal
    preco: Decimal
    prazo: str = ""
    day_trade: bool = False


@dataclass(frozen=True)
class LinhaBMF:
    """``ajuste``: valor da linha na nota, crédito positivo e débito negativo. ``tipo``: NORMAL,
    DAY TRADE ou AJUPOS (ajuste da posição carregada, com a quantidade negativa se vendida)."""

    cv: CV
    contrato: str
    vencimento: date | None
    quantidade: Decimal
    preco: Decimal
    ajuste: Decimal
    tipo: str = "NORMAL"

    @property
    def day_trade(self) -> bool:
        return self.tipo == "DAY TRADE"


TIPOS_BMF = ("NORMAL", "DAY TRADE", "AJUPOS")


@dataclass(frozen=True)
class NotaManual:
    """``conta_id``: conta já cadastrada; sem ela, ``corretora`` e ``conta`` criam a conta no
    titular (corretora que nunca teve nota importada)."""

    conta_id: int | None
    data: date
    numero: str
    segmento: Segmento
    bovespa: tuple[LinhaBovespa, ...] = ()
    bmf: tuple[LinhaBMF, ...] = ()
    custos: dict[CategoriaCusto, Decimal] = field(default_factory=dict[CategoriaCusto, Decimal])
    irrf: Decimal = ZERO  # I.R.R.F. s/ operações (0,005%)
    irrf_day_trade: Decimal = ZERO
    corretora: str = ""
    conta: str = ""


@dataclass(frozen=True)
class Sobreposta:
    """Nota manual com nota do PDF na mesma conta, dia e segmento (pode ser a mesma, em dobro)."""

    nota_id: int
    numero: str
    data: date
    corretora: str


@dataclass(frozen=True)
class Gravada:
    nota_id: int
    avisos: list[str]


def _desconhecidos(ops: Sequence[OperacaoExtraida], refs: Referencias) -> list[str]:
    """Código digitado que nenhuma tabela conhece: o final basta para o COIR (como no código
    impresso pela corretora), mas pode ser erro de digitação. Avisa, sem recusar: ativo novo pode
    não estar nas tabelas ainda."""
    conhecidos = set(refs.ativos.values()) | {
        t for periodos in refs.nomes_historicos.values() for _, _, t in periodos
    }
    papeis = {o.ativo_objeto if o.mercado in ("CALL", "PUT") else o.papel for o in ops}
    return [f"{p}: código que as tabelas do COIR não conhecem; confira se foi digitado certo"
            for p in sorted(x for x in papeis if x and x not in conhecidos)]  # fmt: skip


class NotaParecida(ImportacaoError):
    """Já há nota na mesma conta, dia e segmento: pode ser a mesma nota com outro número (as
    operações contariam duas vezes). A gravação só segue com ``confirmar_parecidas``."""

    def __init__(self, parecidas: list[str]) -> None:
        super().__init__("já há nota nesta conta e pregão: " + "; ".join(parecidas))
        self.parecidas = parecidas


def _quantia(valor: Decimal) -> str:
    texto = format(valor.normalize(), "f")
    return texto.replace(".", ",")


def _parecidas(
    s: Session, conta: Conta, extraida: NotaExtraida, ops: Sequence[OperacaoExtraida],
    ignorar: int | None,
) -> list[str]:  # fmt: skip
    """As notas gravadas na mesma conta, dia e segmento, cada uma com os negócios iguais aos da
    nota digitada (ativo, compra ou venda, quantidade e preço)."""
    consulta = select(NotaCorretagem).where(
        NotaCorretagem.conta_id == conta.id,
        NotaCorretagem.data_pregao == extraida.data_pregao,
        NotaCorretagem.segmento == extraida.segmento.value,
    )
    digitados = {(o.papel, o.cv, o.quantidade, o.preco) for o in ops}
    saida: list[str] = []
    for n in s.scalars(consulta):
        if n.id == ignorar:
            continue
        iguais = [o for o in n.operacoes if (o.papel, o.cv, o.quantidade, o.preco) in digitados]
        origem = "manual" if eh_manual(n.arquivo.nome) else "do PDF"
        texto = (
            f"nota {n.numero} ({origem}) de {n.data_pregao:%d/%m/%Y}, {len(n.operacoes)} negócio(s)"
        )
        if iguais:
            exemplos = ", ".join(f"{o.papel} {o.cv} {_quantia(o.quantidade)} a "
                                 f"{_quantia(o.preco)}" for o in iguais[:3])  # fmt: skip
            texto += f", {len(iguais)} igual(is) aos digitados ({exemplos})"
        saida.append(texto)
    return saida


def _mesma_chave(s: Session, extraida: NotaExtraida, ignorar: int | None) -> None:
    """Mesmo número, conta, dia e segmento de uma nota gravada: recusa (a chave é única)."""
    existente = repo.nota_por_chave(s, extraida.chave)
    if existente is not None and existente.id != ignorar:
        raise ImportacaoError(
            f"já existe a nota {extraida.numero} de {extraida.data_pregao:%d/%m/%Y} nesta conta "
            f"(de {existente.arquivo.nome}); use outro número"
        )


def eh_manual(nome_arquivo: str) -> bool:
    return nome_arquivo.startswith(PREFIXO)


def _montante(valor: Decimal) -> Montante:
    return Montante(valor=abs(valor), dc="D" if valor < 0 else "C")


def _debito(valor: Decimal) -> Montante:
    return Montante(valor=abs(valor), dc="C" if valor < 0 else "D")


def _validar(nota: NotaManual) -> None:
    if not nota.numero.strip():
        raise ImportacaoError("informe o número da nota (o da corretora ou um seu, como M1)")
    linhas = nota.bovespa if nota.segmento is Segmento.BOVESPA else nota.bmf
    if not linhas:
        raise ImportacaoError("a nota precisa de pelo menos uma operação")
    if nota.bovespa and nota.bmf:
        raise ImportacaoError("uma nota é só Bovespa ou só BM&F")
    for i, linha in enumerate(linhas, start=1):
        ajuste = isinstance(linha, LinhaBMF) and linha.tipo == "AJUPOS"
        if linha.quantidade == 0 or (linha.quantidade < 0 and not ajuste):
            raise ImportacaoError(f"operação {i}: a quantidade deve ser maior que zero")
        if linha.preco <= 0:
            raise ImportacaoError(f"operação {i}: o preço deve ser maior que zero")
    for i, linha in enumerate(nota.bmf, start=1):
        if linha.tipo not in TIPOS_BMF:
            raise ImportacaoError(f"operação {i}: tipo {linha.tipo!r} desconhecido")
        if not linha.contrato.strip():
            raise ImportacaoError(f"operação {i}: informe o contrato (ex.: WINZ25)")
    for i, linha in enumerate(nota.bovespa, start=1):
        if linha.mercado not in MERCADOS_BOVESPA:
            raise ImportacaoError(f"operação {i}: mercado {linha.mercado!r} desconhecido")
        if not linha.ativo.strip():
            raise ImportacaoError(f"operação {i}: informe o código do ativo")
        if linha.mercado in OPCOES and not _vencimento_opcao(linha.prazo):
            raise ImportacaoError(f"operação {i}: informe o vencimento da opção (MM/AA)")
        if linha.mercado == "TERMO" and not linha.prazo.strip().isdigit():
            raise ImportacaoError(f"operação {i}: informe o prazo do termo, em dias")
    for categoria, valor in nota.custos.items():
        if valor < 0:
            raise ImportacaoError(f"{ROTULOS[categoria]}: custo negativo não existe")
        if nota.segmento is Segmento.BMF and categoria not in CUSTOS_BMF:
            raise ImportacaoError(f"{ROTULOS[categoria]}: custo que a nota BM&F não tem")
        if nota.segmento is Segmento.BOVESPA and categoria not in GRUPO_BOVESPA:
            raise ImportacaoError(f"{ROTULOS[categoria]}: custo que a nota Bovespa não tem")
    if nota.irrf < 0 or nota.irrf_day_trade < 0:
        raise ImportacaoError("o IRRF não pode ser negativo")


def _vencimento_opcao(prazo: str) -> bool:
    partes = prazo.strip().split("/")
    return (len(partes) == 2 and all(p.isdigit() and len(p) == 2 for p in partes)
            and 1 <= int(partes[0]) <= 12)  # fmt: skip


def _valor(quantidade: Decimal, preco: Decimal) -> Decimal:
    return (quantidade * preco).quantize(CENTAVO, ROUND_HALF_UP)


def nome_arquivo(corretora: str, nota: NotaManual) -> str:
    return f"{PREFIXO} {corretora} {nota.data:%d-%m-%Y} nº {nota.numero.strip()}"


def montar(nota: NotaManual, corretora: str, conta: str) -> NotaExtraida:
    """A nota digitada como ``NotaExtraida`` conciliada (ExtracaoError se não fechar)."""
    _validar(nota)
    custos = {c: v for c, v in nota.custos.items() if v != 0}
    comum = {"corretora": corretora, "conta": conta, "numero": nota.numero.strip(),
             "data_pregao": nota.data, "segmento": nota.segmento,
             "arquivo": nome_arquivo(corretora, nota), "paginas": ()}  # fmt: skip
    if nota.segmento is Segmento.BOVESPA:
        extraida = NotaExtraida(**comum, negocios_bovespa=tuple(_negocios_bovespa(nota)),  # type: ignore[arg-type]
                                resumo_bovespa=_resumo_bovespa(nota, custos))  # fmt: skip
    else:
        extraida = NotaExtraida(**comum, negocios_bmf=tuple(_negocios_bmf(nota, custos)),  # type: ignore[arg-type]
                                resumo_bmf=_resumo_bmf(nota, custos))  # fmt: skip
    return conciliar(extraida)


def _negocios_bovespa(nota: NotaManual) -> list[NegocioBovespa]:
    return [
        NegocioBovespa(
            cv=n.cv,
            mercado=n.mercado,
            prazo=n.prazo.strip(),
            especificacao=n.ativo.strip().upper(),
            observacao="D" if n.day_trade else "",
            quantidade=n.quantidade,
            preco=n.preco,
            valor=_valor(n.quantidade, n.preco),
            dc="D" if n.cv == "C" else "C",
        )
        for n in nota.bovespa
    ]


def _resumo_bovespa(nota: NotaManual, custos: dict[CategoriaCusto, Decimal]) -> ResumoBovespa:
    negocios = _negocios_bovespa(nota)
    termo = sum((n.valor for n in negocios if n.mercado == "TERMO"), ZERO)
    compras = sum((n.valor for n in negocios if n.mercado != "TERMO" and n.cv == "C"), ZERO)
    vendas = sum((n.valor for n in negocios if n.mercado != "TERMO" and n.cv == "V"), ZERO)
    linhas = tuple(Custo(rotulo=ROTULOS[c], categoria=c, grupo=GRUPO_BOVESPA[c], valor=v)
                   for c, v in custos.items())  # fmt: skip

    def soma(grupo: GrupoCusto) -> Decimal:
        return sum((c.valor for c in linhas if c.grupo == grupo), ZERO)

    cblc = vendas - compras - soma("cblc")
    liquido = cblc - soma("bolsa") - soma("despesas") - nota.irrf  # IRRF descontado na nota
    return ResumoBovespa(
        valor_operacoes=compras + vendas + termo, compras=compras, vendas=vendas, termo=termo,
        valor_liquido_operacoes=_montante(vendas - compras), custos=linhas,
        total_cblc=_montante(cblc), total_bovespa=_debito(soma("bolsa")),
        total_despesas=_debito(soma("despesas")), irrf=nota.irrf, irrf_base=None,
        irrf_no_liquido=nota.irrf > 0, irrf_day_trade=nota.irrf_day_trade,
        liquido=_montante(liquido), data_liquidacao=None,
    )  # fmt: skip


def _negocios_bmf(nota: NotaManual, custos: dict[CategoriaCusto, Decimal]) -> list[NegocioBMF]:
    corretagem = custos.get(CategoriaCusto.CORRETAGEM, ZERO)  # a nota a imprime por linha
    return [
        NegocioBMF(
            cv=n.cv,
            mercadoria=n.contrato.strip().upper().replace(" ", ""),
            vencimento=n.vencimento,
            quantidade=n.quantidade,
            preco=n.preco,
            tipo=n.tipo,
            valor=abs(n.ajuste),
            dc="D" if n.ajuste < 0 else "C",
            taxa_operacional=corretagem if i == 0 else ZERO,
        )
        for i, n in enumerate(nota.bmf)
    ]


def _resumo_bmf(nota: NotaManual, custos: dict[CategoriaCusto, Decimal]) -> ResumoBMF:
    impostos = custos.get(CategoriaCusto.IMPOSTOS, ZERO)
    linhas = tuple(Custo(rotulo=ROTULOS[c], categoria=c, grupo="bmf", valor=v)
                   for c, v in custos.items() if c is not CategoriaCusto.IMPOSTOS)  # fmt: skip
    total = sum((c.valor for c in linhas), ZERO)
    day_trade = sum((n.ajuste for n in nota.bmf if n.day_trade), ZERO)
    posicao = sum((n.ajuste for n in nota.bmf if not n.day_trade), ZERO)
    liquido = day_trade + posicao - total - impostos - nota.irrf - nota.irrf_day_trade
    return ResumoBMF(
        valor_negocios=_montante(day_trade + posicao), ajuste_posicao=_montante(posicao),
        ajuste_day_trade=_montante(day_trade), custos=linhas, total_custos=total,
        impostos=impostos, irrf=nota.irrf, irrf_day_trade=nota.irrf_day_trade,
        liquido=_montante(liquido),
    )  # fmt: skip


def _sha(extraida: NotaExtraida) -> str:
    chave = "|".join(str(x) for x in (PREFIXO, *extraida.chave))
    return hashlib.sha256(chave.encode()).hexdigest()


def _conta(s: Session, titular: int, nota: NotaManual) -> Conta:
    if nota.conta_id is not None:
        conta = s.get(Conta, nota.conta_id)
        if conta is None or conta.titular_id != titular:
            raise ImportacaoError("escolha uma conta do titular")
        return conta
    corretora, numero = nota.corretora.strip().upper(), nota.conta.strip()
    if not corretora or not numero:
        raise ImportacaoError("escolha a conta, ou informe a corretora e o número da conta nova")
    dono = s.get(Titular, titular)
    if dono is None:
        raise ImportacaoError("titular não encontrado")
    try:
        return repo.Contas(s, dono, explicito=True).conta_da_nota(corretora, numero, None)
    except repo.ContaRecusada as erro:
        raise ImportacaoError(str(erro)) from erro


def _manual(s: Session, titular: int, nota_id: int) -> NotaCorretagem:
    registro = s.get(NotaCorretagem, nota_id)
    if registro is None or registro.conta.titular_id != titular:
        raise ImportacaoError("nota não encontrada para o titular")
    if not eh_manual(registro.arquivo.nome):
        raise ImportacaoError("só a nota manual pode ser corrigida ou removida aqui; a do PDF "
                              "muda pelo reprocessamento na tela Importar")  # fmt: skip
    return registro


def _extrair(nota: NotaManual, conta: Conta) -> NotaExtraida:
    try:
        return montar(nota, conta.corretora.nome, conta.numero)
    except ExtracaoError as erro:
        raise ImportacaoError(f"a nota não fecha: {erro}") from erro


def gravar(
    banco: Path, titular: int, nota: NotaManual, refs: Referencias, fiscais: ReferenciasFiscais,
    agora: datetime, substituir: int | None = None, confirmar_parecidas: bool = False,
) -> Gravada:  # fmt: skip
    """Grava a nota (ou corrige a manual ``substituir``) e devolve o id. Nada muda se o ativo não
    está nas tabelas, se a nota não fecha ou se o livro passar a ter pendência por causa dela.
    Com nota já gravada na mesma conta, dia e segmento, levanta ``NotaParecida`` (pode ser a mesma
    nota com outro número), salvo com ``confirmar_parecidas``."""
    with conectar(banco) as fabrica, fabrica() as s:  # confere antes do backup (nada é gravado)
        conta = _conta(s, titular, nota)
        extraida = _extrair(nota, conta)
        _mesma_chave(s, extraida, substituir)
        if not confirmar_parecidas:
            try:
                previas = operacoes(extraida, refs, {
                    **classes_das_series([extraida], refs), **repo.series(s)})  # fmt: skip
            except ExtracaoError as erro:
                raise ImportacaoError(str(erro)) from erro
            parecidas = _parecidas(s, conta, extraida, previas, substituir)
            if parecidas:
                raise NotaParecida(parecidas)
    fazer_backup(banco, agora)
    with conectar(banco) as fabrica, fabrica.begin() as s:
        conta = _conta(s, titular, nota)
        extraida = _extrair(nota, conta)
        antes = pendencia_antes(s, titular, refs, fiscais)
        situacao, desde, arquivo = "manual", extraida.data_pregao, None
        if substituir is not None:
            velha = _manual(s, titular, substituir)
            # o mesmo registro de arquivo segue com a nota corrigida (mesmo que o número mude):
            # o histórico da tela Notas acha a versão anterior por ele
            arquivo, desde = velha.arquivo, min(desde, velha.data_pregao)
            repo.remover_nota(s, velha, "substituida", "nota manual corrigida", agora)
            s.flush()
            situacao = "manual_corrigida"
        _mesma_chave(s, extraida, None)
        conhecidas = repo.series(s)
        novas = {k: v for k, v in classes_das_series([extraida], refs).items()
                 if k not in conhecidas}  # fmt: skip
        repo.gravar_series(s, novas)
        ativos = repo.Ativos(s)
        ativos.sincronizar_trocas(refs.trocas)
        try:
            ops = operacoes(extraida, refs, {**novas, **conhecidas})
        except ExtracaoError as erro:
            raise ImportacaoError(str(erro)) from erro
        sha = _sha(extraida)
        do_sha = repo.arquivo_por_sha(s, sha)
        if arquivo is None:
            arquivo = do_sha
        elif do_sha is None:  # número ou data mudou: o registro passa a ter a chave nova
            arquivo.sha256 = sha
        if arquivo is None:
            arquivo = ArquivoImportado(sha256=sha, nome=extraida.arquivo, tamanho=0,
                                       importado_em=agora, completo=True)  # fmt: skip
            s.add(arquivo)
        arquivo.nome = extraida.arquivo
        registro = repo.gravar_nota(s, extraida, ops, arquivo, conta, ativos)
        s.flush()
        conferir_no_motor(s, titular, refs, fiscais, antes, "a nota")
        nota_id = registro.id
        avisos = _desconhecidos(ops, refs) if nota.segmento is Segmento.BOVESPA else []
        avisos += _darf_paga(s, titular, desde, "a nota gravada")
    trilha.nota(extraida, situacao, len(ops))
    return Gravada(nota_id, avisos)


def _darf_paga(s: Session, titular: int, desde: date, o_que: str) -> list[str]:
    """A nota muda o mês dela e pode mudar os seguintes (preço médio e prejuízo)."""
    pagas = repo.darfs_pagas_desde(s, titular, desde.year, desde.month)
    if not pagas:
        return []
    return [
        f"{pagas} DARF(s) paga(s) de {desde:%m/%Y} em diante: {o_que} pode mudar o imposto "
        "desses meses, pelo preço médio e pelo prejuízo. Veja cada mês na tela DARF"
    ]


def remover(
    banco: Path, titular: int, nota_id: int, refs: Referencias, fiscais: ReferenciasFiscais,
    agora: datetime,
) -> list[str]:  # fmt: skip
    """Remove a nota manual (a versão fica em ``nota_anterior``); devolve os avisos."""
    with conectar(banco) as fabrica, fabrica() as s:
        _manual(s, titular, nota_id)
    fazer_backup(banco, agora)
    avisos: list[str] = []
    with conectar(banco) as fabrica, fabrica.begin() as s:
        registro = _manual(s, titular, nota_id)
        extraida = NotaExtraida.model_validate_json(registro.json_nota)
        antes = pendencia_antes(s, titular, refs, fiscais)
        repo.remover_nota(s, registro, "removida", "nota manual removida", agora)
        s.flush()
        conferir_no_motor(s, titular, refs, fiscais, antes, "o livro sem a nota")
        avisos += _darf_paga(s, titular, extraida.data_pregao, "a nota removida")
    trilha.nota(extraida, "manual_removida")
    return avisos


def para_corrigir(banco: Path, titular: int, nota_id: int) -> NotaManual:
    """A nota manual gravada de volta no formulário."""
    with conectar(banco) as fabrica, fabrica() as s:
        registro = _manual(s, titular, nota_id)
        n = NotaExtraida.model_validate_json(registro.json_nota)
        conta_id = registro.conta_id
    return de_extraida(n, conta_id)


def de_extraida(n: NotaExtraida, conta_id: int | None) -> NotaManual:
    """Uma nota (do banco ou do PDF) como os campos do formulário; os custos somados por
    categoria (o rótulo e o subtotal impressos não mudam o cálculo)."""
    resumo = n.resumo_bovespa or n.resumo_bmf
    assert resumo is not None
    custos: dict[CategoriaCusto, Decimal] = {}
    for c in resumo.custos:
        custos[c.categoria] = custos.get(c.categoria, ZERO) + c.valor
    if n.resumo_bmf is not None and n.resumo_bmf.impostos:
        impostos = custos.get(CategoriaCusto.IMPOSTOS, ZERO) + n.resumo_bmf.impostos
        custos[CategoriaCusto.IMPOSTOS] = impostos
    return NotaManual(
        conta_id=conta_id, data=n.data_pregao, numero=n.numero, segmento=n.segmento,
        bovespa=tuple(LinhaBovespa(cv=x.cv, mercado=x.mercado, ativo=x.especificacao,
                                   quantidade=x.quantidade, preco=x.preco, prazo=x.prazo,
                                   day_trade=x.day_trade) for x in n.negocios_bovespa),
        bmf=tuple(LinhaBMF(cv=x.cv, contrato=x.mercadoria, vencimento=x.vencimento,
                           quantidade=x.quantidade, preco=x.preco,
                           ajuste=x.valor if x.dc != "D" else -x.valor,
                           tipo=_tipo_bmf(x.tipo)) for x in n.negocios_bmf),
        custos={c: v for c, v in custos.items() if v}, irrf=resumo.irrf,
        irrf_day_trade=resumo.irrf_day_trade,
        corretora="" if conta_id is not None else n.corretora,
        conta="" if conta_id is not None else n.conta,
    )  # fmt: skip


def proximo_numero(banco: Path, titular: int) -> str:
    """M1, M2...: o primeiro número livre entre as notas manuais do titular."""
    with conectar(banco) as fabrica, fabrica() as s:
        usados = set(s.scalars(
            select(NotaCorretagem.numero).join(Conta).join(ArquivoImportado)
            .where(Conta.titular_id == titular, ArquivoImportado.nome.startswith(PREFIXO))
        ))  # fmt: skip
    n = 1
    while f"M{n}" in usados:
        n += 1
    return f"M{n}"


def sobrepostas(banco: Path, titular: int) -> list[Sobreposta]:
    """Notas manuais do titular com nota do PDF na mesma conta, dia e segmento."""
    with conectar(banco) as fabrica, fabrica() as s:
        todas = list(s.scalars(select(NotaCorretagem).join(Conta)
                               .where(Conta.titular_id == titular)))  # fmt: skip
        do_pdf = {(n.conta_id, n.data_pregao, n.segmento) for n in todas
                  if not eh_manual(n.arquivo.nome)}  # fmt: skip
        return [
            Sobreposta(n.id, n.numero, n.data_pregao, n.conta.corretora.nome)
            for n in todas
            if eh_manual(n.arquivo.nome) and (n.conta_id, n.data_pregao, n.segmento) in do_pdf
        ]


def _tipo_bmf(tipo: str) -> str:
    compacto = tipo.replace(" ", "").upper()
    return {"DAYTRADE": "DAY TRADE", "AJUPOS": "AJUPOS"}.get(compacto, "NORMAL")


def contas_do_titular(banco: Path, titular: int) -> dict[int, str]:
    """Contas do titular para o formulário: id -> "XP 1234567"."""
    with conectar(banco) as fabrica, fabrica() as s:
        contas = s.scalars(select(Conta).where(Conta.titular_id == titular))
        rotulos = {c.id: f"{c.corretora.nome} {c.numero}" for c in contas}
    return dict(sorted(rotulos.items(), key=lambda par: par[1]))
