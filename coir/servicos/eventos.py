"""Eventos corporativos e de custódia: lançados na tela 7 (Eventos) ou num CSV mantido pelo
usuário (em COIR_DADOS) -> banco.

Colunas: data, corretora, conta, ticker, tipo, cv, quantidade, valor, corretora_destino,
conta_destino, observacao, origem e, opcionais, ticker_destino, quantidade_destino, natureza e
percentual_custo. Ajuste de posição: C soma quantidade e custo (valor), V subtrai (como na aba
Eventos_B3 do gabarito). Transferência: cv vazio, conta de destino e quantidade (vazia = toda a
posição). Incorporação e cisão: cv vazio, quantidade de origem, ativo e quantidade recebidos
(natureza na incorporação; percentual do custo na cisão). Leilão de fração: quantidade e valor
líquido. Reimportar o mesmo CSV não duplica (chave = hash da linha).
"""

import csv
import hashlib
import re
from collections import defaultdict
from dataclasses import dataclass, replace
from datetime import date, datetime, timedelta
from decimal import ROUND_FLOOR, Decimal, InvalidOperation
from pathlib import Path

from sqlalchemy import select
from sqlalchemy.orm import Session, selectinload

from coir import referencias
from coir.db import repositorio as repo
from coir.db.modelos import (
    Conta,
    Corretora,
    EventoCorporativo,
    MovimentacaoB3,
    PropostaIgnorada,
    SerieOpcao,
    TickerHistorico,
)
from coir.db.sessao import conectar
from coir.domain.fiscal import (
    TIPOS_COM_DESTINO,
    Movimento,
    NaturezaEvento,
    TipoAtivo,
    TipoEvento,
    Visao,
)
from coir.domain.parametros import ReferenciasFiscais
from coir.domain.referencias import Referencias
from coir.engine.calendario import proximo_pregao
from coir.engine.movimentos import movimentar
from coir.erros import ImportacaoError, MotorError, ReferenciaError
from coir.servicos.apuracao import Tickers, entradas

COLUNAS = ("data", "corretora", "conta", "ticker", "tipo", "cv", "quantidade", "valor",
           "corretora_destino", "conta_destino", "observacao", "origem")  # fmt: skip
EXTRAS = ("ticker_destino", "quantidade_destino", "natureza", "percentual_custo")  # opcionais
NUMERICAS = {"quantidade", "valor", "quantidade_destino", "percentual_custo"}


@dataclass
class RelatorioEventos:
    novos: int = 0
    ja_no_banco: int = 0
    fora_do_csv: int = 0  # no banco e não no CSV (linha corrigida ou apagada)
    removidos: int = 0


def _decimal(texto: str, campo: str, linha: int) -> Decimal | None:
    if not texto:
        return None
    try:
        valor = Decimal(texto)
    except InvalidOperation as erro:
        raise ImportacaoError(f"eventos linha {linha}: {campo} inválido ({texto!r})") from erro
    if not valor.is_finite() or valor < 0:  # o sentido vem da coluna C/V, nunca do sinal
        raise ImportacaoError(f"eventos linha {linha}: {campo} deve ser número >= 0 ({texto!r})")
    return valor


EXIGE_QUANTIDADE = {TipoEvento.DESDOBRAMENTO, TipoEvento.GRUPAMENTO, TipoEvento.BONIFICACAO,
                    TipoEvento.SUBSCRICAO, TipoEvento.CUSTODIA_INICIAL}  # fmt: skip


def _conta(sessao: Session, corretora: str, numero: str, linha: int) -> Conta:
    conta = sessao.scalar(
        select(Conta).join(Corretora).where(Corretora.nome == corretora, Conta.numero == numero)
    )
    if conta is None:
        raise ImportacaoError(
            f"eventos linha {linha}: conta da {corretora} sem notas no banco (importe as notas)"
        )
    return conta


def _ativo(sessao: Session, ticker: str, dia: date, linha: int) -> int:
    """Ativo já negociado (evento de ticker sem nota no banco é erro de digitação provável)."""
    ativo = repo.ativo_existente(sessao, ticker, dia) or repo.ativo_pelo_codigo_antigo(
        sessao, ticker, dia
    )
    if ativo is None:
        raise ImportacaoError(f"eventos linha {linha}: {ticker} sem negócio no banco em {dia}")
    return ativo


def _chave(linha: dict[str, str], n: int) -> str:
    """Hash dos campos que definem o evento, normalizados ("100" = "100.0"; tipo e C/V em
    maiúsculas); observação e origem não entram. Os campos opcionais só entram preenchidos (a
    chave dos eventos antigos não muda)."""
    nomes = [*COLUNAS[:10], *(c for c in EXTRAS if linha.get(c, ""))]
    campos: list[str] = []
    for nome in nomes:
        campo = linha.get(nome, "").upper()
        if nome in NUMERICAS:
            numero = _decimal(campo, nome, n)
            campo = "" if numero is None else format(numero.normalize(), "f")
        campos.append(campo if nome in COLUNAS else f"{nome}={campo}")
    return hashlib.sha256("|".join(campos).encode()).hexdigest()


def _chaves(linha: dict[str, str], n: int) -> tuple[str, str]:
    """A chave da linha e a da mesma linha com o valor zero escrito do outro jeito ("" ou "0"):
    o mesmo evento lançado na tela e no CSV não duplica."""
    valor = _decimal(linha.get("valor", ""), "valor", n)
    outra = "0" if not linha.get("valor", "") else ("" if valor == 0 else linha["valor"])
    return _chave(linha, n), _chave({**linha, "valor": outra}, n)


ORIGENS_TELA = frozenset({"tela", "b3"})  # lançados na tela 7: o CSV com --substituir não apaga


def ler_csv(caminho: Path) -> list[dict[str, str]]:
    """CSV em UTF-8 (com ou sem BOM), separado por vírgula ou ponto e vírgula (Excel pt-BR)."""
    with caminho.open(encoding="utf-8-sig", newline="") as arquivo:
        cabecalho = arquivo.readline()
        arquivo.seek(0)
        leitor = csv.DictReader(
            arquivo, delimiter=";" if cabecalho.count(";") > cabecalho.count(",") else ","
        )
        faltam = [c for c in COLUNAS if c not in (leitor.fieldnames or [])]
        if faltam:
            raise ImportacaoError(f"{caminho.name}: colunas ausentes: {', '.join(faltam)}")
        return [{k: (v or "").strip() for k, v in linha.items()} for linha in leitor]


def importar_eventos(
    sessao: Session, linhas: list[dict[str, str]], substituir: bool = False
) -> RelatorioEventos:
    """Grava as linhas novas numa transação (qualquer linha inválida desfaz tudo). Eventos do
    banco que não estão no CSV são contados; com `substituir`, removidos (o CSV é a fonte), menos
    os lançados na tela 7."""
    rel = RelatorioEventos()
    vistas: dict[str, int] = {}
    with sessao.begin():
        for n, linha in enumerate(linhas, start=2):
            chave, outra = _chaves(linha, n)
            if chave in vistas or outra in vistas:
                anterior = vistas.get(chave, vistas.get(outra))
                raise ImportacaoError(f"eventos linha {n}: repete a linha {anterior}")
            vistas[chave] = vistas[outra] = n
            if sessao.scalar(select(EventoCorporativo).where(
                    EventoCorporativo.chave.in_((chave, outra)))):  # fmt: skip
                rel.ja_no_banco += 1
                continue
            try:
                tipo = TipoEvento(linha["tipo"].upper())
                dia = date.fromisoformat(linha["data"])
                natureza = (NaturezaEvento(linha["natureza"].upper()) if linha.get("natureza")
                            else None)  # fmt: skip
            except ValueError as erro:
                raise ImportacaoError(f"eventos linha {n}: {erro}") from erro
            destino = None
            if tipo is TipoEvento.TRANSFERENCIA:
                destino = _conta(sessao, linha["corretora_destino"], linha["conta_destino"], n)
            conta = _conta(sessao, linha["corretora"], linha["conta"], n)
            d = DadosEvento(
                tipo=tipo, data=dia, conta=conta.id, ticker=linha["ticker"].upper(),
                cv=linha["cv"].upper() or None,
                quantidade=_decimal(linha["quantidade"], "quantidade", n),
                valor=_decimal(linha["valor"], "valor", n) or Decimal(0),
                conta_destino=None if destino is None else destino.id,
                ticker_destino=linha.get("ticker_destino", "").upper() or None,
                quantidade_destino=_decimal(linha.get("quantidade_destino", ""),
                                            "quantidade_destino", n),
                natureza=natureza,
                percentual_custo=_decimal(linha.get("percentual_custo", ""), "percentual_custo",
                                          n),
                observacao=linha["observacao"], origem=linha["origem"] or "manual",
            )  # fmt: skip
            try:
                _validar(d)
            except ImportacaoError as erro:
                raise ImportacaoError(f"eventos linha {n}: {erro}") from erro
            if destino is not None and destino.titular_id != conta.titular_id:
                raise ImportacaoError(
                    f"eventos linha {n}: transferência entre titulares diferentes"
                )
            registro = EventoCorporativo()
            _preencher(sessao, registro, d, chave, _ativo(sessao, d.ticker, dia, n))
            sessao.add(registro)
            rel.novos += 1
        fora = [e for e in sessao.scalars(select(EventoCorporativo))
                if e.chave not in vistas and e.origem not in ORIGENS_TELA]  # fmt: skip
        if substituir:
            for e in fora:
                sessao.delete(e)
            rel.removidos = len(fora)
        else:
            rel.fora_do_csv = len(fora)
    return rel


# ------------------------------------------------------------------------ tela 7 (Eventos)
NOTA_DATA_EX = ("a B3 informa a data do crédito: use a data ex (primeiro pregão sem o direito), "
                "como decidido para os eventos")  # fmt: skip
SEM_CV = frozenset({TipoEvento.TRANSFERENCIA, *TIPOS_COM_DESTINO, TipoEvento.LEILAO_FRACAO})
QUANTIDADE_OBRIGATORIA = EXIGE_QUANTIDADE | {*TIPOS_COM_DESTINO, TipoEvento.LEILAO_FRACAO}
AVISO_INCORPORACAO_ACOES = (
    "incorporação de ações (Lei 6.404 art. 252): o COIR mantém o custo (posição do Judiciário); a "
    "Receita trata como alienação com ganho de capital (SC Cosit 224/2014). Avalie com o seu "
    "contador.")  # fmt: skip


@dataclass(frozen=True)
class DadosEvento:
    """O que a tela grava (ou propõe a partir da movimentação da B3)."""

    tipo: TipoEvento
    data: date
    conta: int | None
    ticker: str
    cv: str | None = None
    quantidade: Decimal | None = None
    valor: Decimal = Decimal(0)
    conta_destino: int | None = None
    ticker_destino: str | None = None
    quantidade_destino: Decimal | None = None
    natureza: NaturezaEvento | None = None
    percentual_custo: Decimal | None = None
    observacao: str = ""
    origem: str = "tela"


@dataclass(frozen=True)
class EventoLancado:
    id: int
    conta_nome: str  # "XP 123456"
    conta_destino_nome: str | None
    dados: DadosEvento


@dataclass(frozen=True)
class Proposta:
    """Linha da movimentação da B3 que sugere um evento. ``situacao``: pendente, lançado (há
    evento do ativo perto da data, ou troca de ticker em data/) ou ignorada."""

    chave: str
    data: date
    movimentacao: str
    codigo: str
    quantidade: Decimal
    valor: Decimal | None
    corretora: str | None
    situacao: str
    dados: DadosEvento
    nota: str


def _nome_conta(conta: Conta) -> str:
    return f"{conta.corretora.nome} {conta.numero}"


def listar(banco: Path, titular: int) -> list[EventoLancado]:
    """Eventos do titular, mais recentes primeiro."""
    with conectar(banco) as fabrica, fabrica() as s:
        tickers = Tickers(s)
        saida: list[EventoLancado] = []
        for e in s.scalars(
            select(EventoCorporativo).join(Conta, EventoCorporativo.conta_id == Conta.id)
            .where(Conta.titular_id == titular)
            .options(selectinload(EventoCorporativo.conta).selectinload(Conta.corretora),
                     selectinload(EventoCorporativo.conta_destino).selectinload(Conta.corretora))
            .order_by(EventoCorporativo.data.desc(), EventoCorporativo.id.desc())
        ):  # fmt: skip
            destino = (None if e.ativo_destino_id is None
                       else tickers.ticker(e.ativo_destino_id, e.data))  # fmt: skip
            dados = DadosEvento(
                tipo=TipoEvento(e.tipo), data=e.data, conta=e.conta_id,
                ticker=tickers.ticker(e.ativo_id, e.data), cv=e.cv, quantidade=e.quantidade,
                valor=e.valor, conta_destino=e.conta_destino_id, ticker_destino=destino,
                quantidade_destino=e.quantidade_destino,
                natureza=None if e.natureza is None else NaturezaEvento(e.natureza),
                percentual_custo=e.percentual_custo, observacao=e.observacao, origem=e.origem,
            )  # fmt: skip
            destino_nome = None if e.conta_destino is None else _nome_conta(e.conta_destino)
            saida.append(EventoLancado(e.id, _nome_conta(e.conta), destino_nome, dados))
        return saida


def _validar(d: DadosEvento) -> None:
    def erro(motivo: str) -> ImportacaoError:
        return ImportacaoError(f"{d.tipo.value.lower().replace('_', ' ')}: {motivo}")

    if not d.ticker.strip():
        raise erro("informe o ativo")
    if d.conta is None:
        raise erro("informe a conta")
    if d.tipo in SEM_CV:
        if d.cv is not None:
            raise erro("este tipo não tem C/V")
    elif d.cv not in ("C", "V"):
        raise erro("informe o C/V: C soma, V subtrai")
    for nome, numero in (("quantidade", d.quantidade), ("valor", d.valor),
                         ("quantidade recebida", d.quantidade_destino)):  # fmt: skip
        if numero is not None and (not numero.is_finite() or numero < 0):
            raise erro(f"{nome} deve ser um número >= 0")
    if d.tipo in QUANTIDADE_OBRIGATORIA and not d.quantidade:
        raise erro("informe a quantidade")
    if d.tipo in (TipoEvento.REDUCAO_CAPITAL, TipoEvento.LEILAO_FRACAO) and not d.valor:
        raise erro("informe o valor")
    if d.tipo is TipoEvento.CUSTODIA_INICIAL:
        # decisões do usuário (25/10/2026): só em 31/12 e só posição comprada
        if (d.data.month, d.data.day) != (12, 31):
            raise erro("a custódia inicial é a posição em 31/12 (a do Bens e Direitos da última "
                       "declaração)")  # fmt: skip
        if d.cv != "C":
            raise erro("lance a posição comprada (C); posição vendida pede as notas")
        if not d.valor:
            raise erro("informe o custo de aquisição (o valor do Bens e Direitos)")
    if d.tipo is TipoEvento.TRANSFERENCIA and (d.conta_destino is None
                                               or d.conta_destino == d.conta):  # fmt: skip
        raise erro("informe a conta de destino (outra conta do titular)")
    if d.tipo in TIPOS_COM_DESTINO:
        if not d.ticker_destino or d.ticker_destino.strip().upper() == d.ticker.strip().upper():
            raise erro("informe o ativo recebido (diferente do de origem)")
        if not d.quantidade_destino:
            raise erro("informe a quantidade recebida")
    if d.tipo is TipoEvento.INCORPORACAO and d.natureza is None:
        raise erro("informe a natureza (incorporação, fusão, conversão, ajuste de série...)")
    if d.tipo not in TIPOS_COM_DESTINO and (d.ticker_destino or d.quantidade_destino):
        raise erro("ativo e quantidade recebidos só valem em incorporação e cisão")
    if d.tipo is not TipoEvento.INCORPORACAO and d.natureza is not None:
        raise erro("natureza só vale na incorporação")
    if d.tipo is not TipoEvento.CISAO and d.percentual_custo is not None:
        raise erro("percentual do custo só vale na cisão")
    if d.tipo is TipoEvento.CISAO and not (d.percentual_custo and 0 < d.percentual_custo < 1):
        # decisão do usuário (P4, 19/10/2026): sem o percentual divulgado, não calcula
        raise erro("informe o percentual do custo que vai para o ativo recebido (aviso aos "
                   "acionistas da companhia)")  # fmt: skip


def _campos_chave(conta: Conta, destino: Conta | None, d: DadosEvento) -> dict[str, str]:
    """Mesmos campos do CSV (a mesma linha no CSV e na tela tem a mesma chave)."""

    def numero(x: Decimal | None) -> str:
        return "" if x is None else format(x, "f")

    return {
        "data": d.data.isoformat(), "corretora": conta.corretora.nome, "conta": conta.numero,
        "ticker": d.ticker, "tipo": d.tipo.value, "cv": d.cv or "",
        "quantidade": numero(d.quantidade), "valor": numero(d.valor),
        "corretora_destino": "" if destino is None else destino.corretora.nome,
        "conta_destino": "" if destino is None else destino.numero,
        "ticker_destino": d.ticker_destino or "",
        "quantidade_destino": numero(d.quantidade_destino),
        "natureza": "" if d.natureza is None else d.natureza.value,
        "percentual_custo": numero(d.percentual_custo),
    }  # fmt: skip


def salvar(
    banco: Path, titular: int, d: DadosEvento, refs: Referencias, fiscais: ReferenciasFiscais,
    evento_id: int | None = None,
) -> int:  # fmt: skip
    """Inclui (ou altera) o evento numa transação e confere com o motor (visão por titular e por
    conta): evento que não fecha com as operações (sem posição, quantidade maior que a posição...)
    não é gravado e o erro diz por quê. Devolve o id."""
    [novo] = salvar_lote(banco, titular, [(d, evento_id)], refs, fiscais)
    return novo


def salvar_lote(
    banco: Path, titular: int, lote: list[tuple[DadosEvento, int | None]], refs: Referencias,
    fiscais: ReferenciasFiscais,
) -> list[int]:  # fmt: skip
    """Vários eventos (novos ou alterados) numa transação, conferidos juntos pelo motor no fim:
    a cisão de uma posição em várias contas só fecha com os eventos de todas."""
    feitos: list[int] = []
    with conectar(banco) as fabrica, fabrica.begin() as s:
        for d, evento_id in lote:
            feitos.append(_gravar(s, titular, d, evento_id))
        _conferir_no_motor(s, titular, refs, fiscais)
    return feitos


def _gravar(s: Session, titular: int, d: DadosEvento, evento_id: int | None) -> int:
    d = replace(d, ticker=d.ticker.strip().upper(),
                ticker_destino=(d.ticker_destino or "").strip().upper() or None,
                observacao=d.observacao.strip())  # fmt: skip
    _validar(d)
    conta = s.get(Conta, d.conta)
    destino = None if d.conta_destino is None else s.get(Conta, d.conta_destino)
    if conta is None or conta.titular_id != titular:
        raise ImportacaoError("conta de outro titular ou inexistente")
    if d.conta_destino is not None and (destino is None or destino.titular_id != titular):
        raise ImportacaoError("transferência entre titulares diferentes")
    chaves = _chaves(_campos_chave(conta, destino, d), 0)
    repetido = s.scalar(select(EventoCorporativo).where(EventoCorporativo.chave.in_(chaves)))
    if repetido is not None and repetido.id != evento_id:
        raise ImportacaoError(f"evento igual já lançado (id {repetido.id})")
    registro = EventoCorporativo() if evento_id is None else s.get(EventoCorporativo, evento_id)
    if registro is None or (evento_id is not None and registro.conta.titular_id != titular):
        raise ImportacaoError(f"evento {evento_id} não encontrado")
    ativo = repo.ativo_existente(s, d.ticker, d.data) or repo.ativo_pelo_codigo_antigo(
        s, d.ticker, d.data
    )
    if ativo is None:
        raise ImportacaoError(f"{d.ticker} sem negócio no banco em {d.data:%d/%m/%Y}")
    _preencher(s, registro, d, chaves[0], ativo)
    if evento_id is None:
        s.add(registro)
    s.flush()
    return registro.id


def _preencher(
    s: Session, registro: EventoCorporativo, d: DadosEvento, chave: str, ativo: int
) -> None:  # fmt: skip
    if d.conta is None:
        raise ImportacaoError("informe a conta")
    registro.chave, registro.data, registro.conta_id, registro.ativo_id = (
        chave, d.data, d.conta, ativo)  # fmt: skip
    registro.tipo, registro.cv, registro.quantidade, registro.valor = (
        d.tipo.value, d.cv, d.quantidade, d.valor)  # fmt: skip
    registro.conta_destino_id = d.conta_destino
    # ativo recebido que nunca foi negociado (ex.: AMOB3, série com código novo) é cadastrado
    registro.ativo_destino_id = (
        None if d.ticker_destino is None else repo.Ativos(s).ativo_id(d.ticker_destino, d.data)
    )
    registro.quantidade_destino, registro.percentual_custo = (
        d.quantidade_destino, d.percentual_custo)  # fmt: skip
    registro.natureza = None if d.natureza is None else d.natureza.value
    registro.observacao, registro.origem = d.observacao, d.origem


def _conferir_no_motor(
    s: Session, titular: int, refs: Referencias, fiscais: ReferenciasFiscais
) -> None:  # fmt: skip
    notas, eventos = entradas(s, refs, fiscais)
    notas = [n for n in notas if n.titular == titular]
    eventos = [e for e in eventos if e.titular == titular]
    try:
        for visao in (Visao.TITULAR, Visao.CONTA):
            movimentar(notas, eventos, fiscais.parametros, date.max, visao)
    except (MotorError, ReferenciaError) as erro:
        raise ImportacaoError(f"o evento não fecha com as operações: {erro}") from erro


def excluir(
    banco: Path, titular: int, evento_id: int, refs: Referencias, fiscais: ReferenciasFiscais
) -> None:  # fmt: skip
    """Apaga o evento se o livro continuar fechando sem ele (ex.: venda que dependia de um
    desdobramento impede apagar o desdobramento)."""
    with conectar(banco) as fabrica, fabrica.begin() as s:
        registro = s.get(EventoCorporativo, evento_id)
        if registro is None or registro.conta.titular_id != titular:
            raise ImportacaoError(f"evento {evento_id} não encontrado")
        s.delete(registro)
        s.flush()
        try:
            _conferir_no_motor(s, titular, refs, fiscais)
        except ImportacaoError as erro:
            raise ImportacaoError(f"sem este evento o livro não fecha: {erro.__cause__}") from erro


# movimentação da B3 (minúsculas) -> tipo e C/V do evento proposto
PROPOSTAS: dict[str, tuple[TipoEvento, str | None]] = {
    "desdobro": (TipoEvento.DESDOBRAMENTO, "C"),
    "bonificação em ativos": (TipoEvento.BONIFICACAO, "C"),
    "restituição de capital": (TipoEvento.REDUCAO_CAPITAL, "V"),
    "recibo de subscrição": (TipoEvento.SUBSCRICAO, "C"),
    "incorporação": (TipoEvento.INCORPORACAO, None),
    "cisão": (TipoEvento.CISAO, None),
    "leilão de fração": (TipoEvento.LEILAO_FRACAO, None),
}
NOTAS_PROPOSTA = {
    TipoEvento.BONIFICACAO: "informe o custo atribuído (aviso aos acionistas)",
    TipoEvento.SUBSCRICAO: "ações subscritas: informe o valor pago",
    TipoEvento.INCORPORACAO: "a B3 mostra o ativo recebido: informe o de origem e a quantidade",
    TipoEvento.CISAO: "informe o ativo cindido e o percentual do custo",
    TipoEvento.LEILAO_FRACAO: "fração vendida pela companhia: venda comum no mês do crédito",
}
RE_RECIBO = re.compile(r"^(?P<raiz>[A-Z]{4})(?P<classe>9|10)$")  # recibo de subscrição ON/PN
JANELA = 10  # dias entre a data da B3 (crédito) e a do evento lançado (data ex)


def propostas(
    banco: Path, titular: int, refs: Referencias, fiscais: ReferenciasFiscais | None = None,
    pasta: Path | None = None,
) -> list[Proposta]:  # fmt: skip
    """Eventos sugeridos, mais recentes primeiro: os do pacote público de referências que atingem
    a posição (com ``fiscais`` e ``pasta``; já com a data ex) e os da movimentação da B3 (data do
    crédito: a tela pede a data ex antes de gravar). O mesmo evento nas duas fontes aparece uma vez
    (a do pacote)."""
    with conectar(banco) as fabrica, fabrica() as s:
        ignoradas = set(s.scalars(select(PropostaIgnorada.chave)
                                  .where(PropostaIgnorada.titular_id == titular)))  # fmt: skip
        contas = list(s.scalars(select(Conta).where(Conta.titular_id == titular)
                                .options(selectinload(Conta.corretora))))  # fmt: skip
        lancados = [(e.ativo_id, e.ativo_destino_id, e.data) for e in s.scalars(
            select(EventoCorporativo).where(EventoCorporativo.conta_id.in_([c.id for c in contas]))
        )]  # fmt: skip
        saida: list[Proposta] = []
        for m in s.scalars(select(MovimentacaoB3).where(MovimentacaoB3.titular_id == titular)
                           .order_by(MovimentacaoB3.data.desc(), MovimentacaoB3.id)):  # fmt: skip
            mapa = PROPOSTAS.get(m.movimentacao.strip().lower())
            if mapa is None or (mapa[0] is TipoEvento.SUBSCRICAO and not m.entrada):
                continue
            tipo, cv = mapa
            corretora = refs.corretora_b3(m.instituicao)
            das_contas = [c for c in contas if c.corretora.nome == corretora]
            conta = das_contas[0].id if len(das_contas) == 1 else None
            codigo = m.codigo or m.produto.split(" - ")[0].strip()
            recibo = RE_RECIBO.match(codigo) if tipo is TipoEvento.SUBSCRICAO else None
            ticker = (f"{recibo['raiz']}{3 if recibo['classe'] == '9' else 4}" if recibo
                      else codigo)  # fmt: skip
            recebido = tipo is TipoEvento.INCORPORACAO  # a B3 mostra o crédito de B
            dados = DadosEvento(
                tipo=tipo, data=m.data, conta=conta, ticker="" if recebido else ticker, cv=cv,
                quantidade=None if recebido else m.quantidade,
                valor=(m.valor or Decimal(0)) if tipo in (TipoEvento.REDUCAO_CAPITAL,
                                                         TipoEvento.LEILAO_FRACAO) else Decimal(0),
                ticker_destino=ticker if recebido else None,
                quantidade_destino=m.quantidade if recebido else None,
                natureza=NaturezaEvento.INCORPORACAO_SOCIEDADE if recebido else None,
                observacao=f"proposto da movimentação da B3 ({m.movimentacao})", origem="b3",
            )  # fmt: skip
            if m.chave in ignoradas:
                situacao = "ignorada"
            elif _ja_lancado(s, refs, ticker, m.data, lancados):
                situacao = "lançado"
            else:
                situacao = "pendente"
            nota = "; ".join(x for x in (NOTAS_PROPOSTA.get(tipo, ""), NOTA_DATA_EX) if x)
            saida.append(Proposta(m.chave, m.data, m.movimentacao, codigo, m.quantidade,
                                  m.valor, corretora, situacao, dados, nota))  # fmt: skip
        if fiscais is None or pasta is None:
            return saida
        do_pacote = _do_pacote(s, titular, refs, fiscais, pasta, contas, ignoradas, lancados)
    perto = timedelta(days=JANELA)

    def mesma(p: Proposta, m: Proposta) -> bool:
        return (
            p.dados.tipo is m.dados.tipo and p.codigo == m.codigo and abs(p.data - m.data) <= perto
        )

    # ignorou a proposta da movimentação: a do pacote para o mesmo evento também fica ignorada
    do_pacote = [replace(p, situacao="ignorada") if p.situacao == "pendente" and any(
        mesma(p, m) and m.situacao == "ignorada" for m in saida) else p
        for p in do_pacote]  # fmt: skip
    repetidas = [m for m in saida if any(
        p.dados.tipo is m.dados.tipo and p.codigo == m.codigo and abs(p.data - m.data) <= perto
        for p in do_pacote)]  # fmt: skip
    return sorted([*do_pacote, *(m for m in saida if m not in repetidas)],
                  key=lambda p: p.data, reverse=True)  # fmt: skip


# eventos do pacote público (eventos_b3.csv: rótulos como a B3 escreve) -> tipo do evento
TIPOS_PACOTE = {
    "DESDOBRAMENTO": TipoEvento.DESDOBRAMENTO,
    "GRUPAMENTO": TipoEvento.GRUPAMENTO,
    "BONIFICACAO": TipoEvento.BONIFICACAO,
    "SUBSCRICAO": TipoEvento.SUBSCRICAO,
}
AJUSTA_SERIE = {t: t for t in (TipoEvento.DESDOBRAMENTO, TipoEvento.GRUPAMENTO,
                                TipoEvento.BONIFICACAO)}  # fmt: skip
GRUPAMENTO_EXTREMO = Decimal("0.01")  # 100 viram 1 ou mais forte
AVISO_GRUPAMENTO_EXTREMO = (
    "confira o aviso aos acionistas: grupamento tão forte costuma vir seguido de desdobramento "
    "no mesmo dia (ex.: Itaú em 2011), sem mudar a quantidade; a B3 publica só o grupamento"
)
NOTA_PACOTE = "evento do pacote público da B3: data ex = pregão seguinte à data com {data_com}"


def _ativos_do_ticker(s: Session, ticker: str) -> set[int]:
    return set(s.scalars(select(TickerHistorico.ativo_id).where(TickerHistorico.ticker == ticker)))


def _posicoes(movs: list[Movimento]) -> dict[tuple[int, int], list[tuple[date, Decimal]]]:
    """(conta, ativo) -> [(data, quantidade depois)], na visão por conta, em ordem."""
    saida: dict[tuple[int, int], list[tuple[date, Decimal]]] = defaultdict(list)
    for m in movs:
        if m.conta is not None and m.ativo is not None:
            saida[(m.conta, m.ativo)].append((m.data, m.quantidade_apos))
    return saida


def _na_data(historico: list[tuple[date, Decimal]], dia: date) -> Decimal:
    q = Decimal(0)
    for quando, depois in historico:
        if quando > dia:
            break
        q = depois
    return q


def _do_pacote(
    s: Session, titular: int, refs: Referencias, fiscais: ReferenciasFiscais, pasta: Path,
    contas: list[Conta], ignoradas: set[str], lancados: list[tuple[int, int | None, date]],
) -> list[Proposta]:  # fmt: skip
    """Para cada evento do pacote de ativo que o titular tinha na data com (por conta), o evento
    com a quantidade pelo fator e, nas séries de opção abertas do ativo, o ajuste de quantidade."""
    eventos = referencias.eventos_b3(pasta)
    if not eventos:
        return []
    por_isin = referencias.tickers_por_isin(pasta)
    notas, evs = entradas(s, refs, fiscais)
    movs = movimentar([n for n in notas if n.titular == titular],
                      [e for e in evs if e.titular == titular], fiscais.parametros, date.max,
                      Visao.CONTA)  # fmt: skip
    historico = _posicoes(movs)
    inicio = min((m.data for m in movs), default=date.max)
    tickers_serie = {m.ativo: m.ticker for m in movs if m.tipo is TipoAtivo.OPCAO and m.ativo}
    objeto_da_serie = {
        s_.serie: s_.serie[:4] + refs.classes[s_.classe]
        for s_ in s.scalars(select(SerieOpcao))
        if s_.classe in refs.classes
    }
    nome_conta = {c.id: c.corretora.nome for c in contas}
    saida: list[Proposta] = []
    for ev in eventos:
        tipo = TIPOS_PACOTE.get(ev["tipo"])
        cisao = ev["tipo"].startswith("CIS") and ev["isin_emitido"] != ev["isin"]
        ticker = por_isin.get(ev["isin"])
        if ticker is None or (tipo is None and not cisao):
            continue
        data_com = date.fromisoformat(ev["data_com"])  # linhas já validadas (eventos_b3)
        cal = fiscais.calendario
        if not cal.primeiro_ano <= data_com.year <= cal.ultimo_ano or data_com < inicio:
            continue  # antes do primeiro negócio (ou fora do calendário): não atinge a carteira
        try:
            data_ex = proximo_pregao(data_com, cal)
        except ReferenciaError:
            continue  # data ex depois do fim do calendário de feriados
        # a B3 escreve o grupamento como razão (0,1 = 10 viram 1) e os demais em % (200 = +200%)
        fator = Decimal(ev["fator"]) / (1 if tipo is TipoEvento.GRUPAMENTO else 100)
        ativos = _ativos_do_ticker(s, ticker)
        nomes = {ticker} | {t for a in ativos for t in s.scalars(
            select(TickerHistorico.ticker).where(TickerHistorico.ativo_id == a))}  # fmt: skip
        for (conta, ativo), hist in historico.items():
            q = _na_data(hist, data_com)
            if not q or conta not in nome_conta:
                continue
            serie = tickers_serie.get(ativo)
            if ativo in ativos and serie is None:
                dados, nota = _evento_acao(ev, tipo, cisao, ticker, conta, q, fator, data_ex,
                                           por_isin)  # fmt: skip
            elif (serie is not None
                  and (objeto_da_serie.get(serie) or refs.opcoes_raiz.get(serie[:4])) in nomes
                  and tipo is not None and tipo in AJUSTA_SERIE):  # fmt: skip
                dados, nota = _ajuste_serie(AJUSTA_SERIE[tipo], serie, conta, q, fator, data_ex)
            else:
                continue
            if dados is None:
                continue
            chave = hashlib.sha256(f"b3pub|{ev['isin']}|{ev['tipo']}|{ev['data_com']}|"
                                   f"{ev['isin_emitido']}|{ev['fator']}|{conta}|{dados.ticker}"
                                   .encode()).hexdigest()  # fmt: skip
            if chave in ignoradas:
                situacao = "ignorada"
            elif _ja_lancado(s, refs, dados.ticker, dados.data, lancados):
                situacao = "lançado"
            else:
                situacao = "pendente"
            saida.append(Proposta(
                chave, dados.data, f"{ev['tipo'].capitalize()} (eventos da B3)", dados.ticker, q,
                None, nome_conta[conta], situacao, dados,
                "; ".join(x for x in (NOTA_PACOTE.format(data_com=f"{data_com:%d/%m/%Y}"), nota)
                          if x)))  # fmt: skip
    return saida


def _fator(ev: dict[str, str], tipo: TipoEvento | None) -> str:
    """Como a B3 escreve: grupamento em razão (0,1 = 10 viram 1); os demais em %."""
    return f"razão {ev['fator']}" if tipo is TipoEvento.GRUPAMENTO else f"{ev['fator']}%"


def _fracao(total: Decimal) -> tuple[Decimal, str]:
    inteiro = total.to_integral_value(ROUND_FLOOR)
    nota = (f"fração de {total - inteiro} vai a leilão: lance o leilão de fração quando for "
            "creditado") if total != inteiro else ""  # fmt: skip
    return inteiro, nota


def _evento_acao(
    ev: dict[str, str], tipo: TipoEvento | None, cisao: bool, ticker: str, conta: int,
    q: Decimal, fator: Decimal, data_ex: date, por_isin: dict[str, str],
) -> tuple[DadosEvento | None, str]:  # fmt: skip
    if q < 0:
        return None, ""  # venda a descoberto na data com: sem proposta (o motor falha alto)
    base = DadosEvento(
        tipo=tipo or TipoEvento.DESDOBRAMENTO, data=data_ex, conta=conta, ticker=ticker,
        cv="C", origem="b3",
        observacao=f"pacote público da B3: {ev['tipo']} {_fator(ev, tipo)} (data com "
                   f"{ev['data_com']})",
    )  # fmt: skip
    if cisao:
        destino = por_isin.get(ev["isin_emitido"])
        if destino is None:
            return None, ""
        recebida, _ = _fracao(q * fator)
        nota = f"informe o percentual do custo que vai para {destino} (aviso aos acionistas)"
        return replace(base, tipo=TipoEvento.CISAO, cv=None, quantidade=q, ticker_destino=destino,
                       quantidade_destino=recebida), nota  # fmt: skip
    if tipo is TipoEvento.GRUPAMENTO:
        inteiro, nota = _fracao(q * fator)
        if fator <= GRUPAMENTO_EXTREMO:
            nota = "; ".join(x for x in (AVISO_GRUPAMENTO_EXTREMO, nota) if x)
        return replace(base, cv="V", quantidade=q - inteiro), nota
    inteiro, nota = _fracao(q * fator)
    if tipo is TipoEvento.SUBSCRICAO:
        preco = Decimal(ev["preco"] or 0)
        dia = date.fromisoformat(ev["data_subscricao"]) if ev["data_subscricao"] else data_ex
        return replace(base, data=dia, quantidade=inteiro,
                       valor=(inteiro * preco).quantize(Decimal("0.01"))), (
            f"subscrição é opcional (preço {preco}, período {ev['periodo']}): lance só o que "
            "você exerceu")  # fmt: skip
    if tipo is TipoEvento.BONIFICACAO:
        nota = "; ".join(x for x in ("informe o custo atribuído (aviso aos acionistas)", nota) if x)
    return replace(base, quantidade=inteiro), nota


def _ajuste_serie(
    tipo: TipoEvento, serie: str, conta: int, q: Decimal, fator: Decimal, data_ex: date
) -> tuple[DadosEvento, str]:  # fmt: skip
    novo = abs(q) * (fator if tipo is TipoEvento.GRUPAMENTO else 1 + fator)
    delta = novo.to_integral_value(ROUND_FLOOR) - abs(q)
    cv = ("C" if delta > 0 else "V") if q > 0 else ("V" if delta > 0 else "C")
    return DadosEvento(tipo=tipo, data=data_ex, conta=conta, ticker=serie, cv=cv,
                       quantidade=abs(delta), origem="b3",
                       observacao="ajuste da série pelo evento do ativo-objeto"), (
        "série de opção aberta: a B3 ajusta a quantidade e o preço de exercício (e às vezes o "
        "código; se mudou, use Incorporação / ajuste de série)")  # fmt: skip


def _ja_lancado(
    s: Session, refs: Referencias, ticker: str, dia: date,
    lancados: list[tuple[int, int | None, date]],
) -> bool:  # fmt: skip
    """Há evento do ativo (de origem ou recebido) a até JANELA dias, ou troca de código em
    data/trocas_ticker.csv (ex.: incorporação JSLG3 -> SIMH3 tratada como troca)."""
    perto = timedelta(days=JANELA)
    if any(t.ticker == ticker and abs(t.ate - dia) <= perto for t in refs.trocas):
        return True
    ativo = repo.ativo_existente(s, ticker, dia)
    if ativo is None:  # código atual de ativo que se chamava outra coisa na data
        candidatos = _ativos_do_ticker(s, ticker)
        ativo = next(iter(candidatos)) if len(candidatos) == 1 else None
    return ativo is not None and any(
        ativo in (origem, destino) and abs(quando - dia) <= perto
        for origem, destino, quando in lancados
    )


def ignorar(banco: Path, titular: int, chave: str, agora: datetime, sim: bool = True) -> None:
    """Marca (ou desmarca) a proposta como ignorada."""
    with conectar(banco) as fabrica, fabrica.begin() as s:
        registro = s.scalar(select(PropostaIgnorada).where(PropostaIgnorada.chave == chave))
        if sim and registro is None:
            s.add(PropostaIgnorada(titular_id=titular, chave=chave, ignorada_em=agora))
        elif not sim and registro is not None:
            s.delete(registro)


def contas_do_titular(banco: Path, titular: int) -> dict[int, str]:
    """Contas do titular para escolher na tela: id -> "XP 123456"."""
    with conectar(banco) as fabrica, fabrica() as s:
        contas = s.scalars(select(Conta).where(Conta.titular_id == titular)
                           .options(selectinload(Conta.corretora)))  # fmt: skip
        return dict(sorted(((c.id, _nome_conta(c)) for c in contas), key=lambda x: x[1]))
