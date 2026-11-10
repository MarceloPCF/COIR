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
from collections.abc import Sequence
from dataclasses import dataclass, replace
from datetime import date, datetime, timedelta
from decimal import ROUND_FLOOR, ROUND_HALF_UP, Decimal, InvalidOperation
from pathlib import Path

from sqlalchemy import func, select
from sqlalchemy.orm import Session, selectinload

from coir import auditoria, referencias
from coir.db import repositorio as repo
from coir.db.modelos import (
    Conta,
    Corretora,
    EventoCorporativo,
    MovimentacaoB3,
    NotaCorretagem,
    PropostaIgnorada,
    SerieOpcao,
    TickerHistorico,
)
from coir.db.sessao import conectar
from coir.domain.fiscal import (
    TIPOS_COM_DESTINO,
    AtivoIsolado,
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
from coir.erros import CoirError, ImportacaoError, MotorError, ReferenciaError
from coir.servicos import trilha
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
                    TipoEvento.SUBSCRICAO, TipoEvento.CUSTODIA_INICIAL,
                    TipoEvento.OFERTA_PUBLICA}  # fmt: skip


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
SEM_CV = frozenset({TipoEvento.TRANSFERENCIA, *TIPOS_COM_DESTINO, TipoEvento.LEILAO_FRACAO,
                    TipoEvento.DIREITO_NAO_EXERCIDO})  # fmt: skip
# direito não exercido sem quantidade baixa todo o saldo: a linha "não exercido" da B3 vem sem ela
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
    segura: bool = False  # lançada sozinha, sem pedir nada ao usuário (ver _segura)
    # depois da última nota: sem saber se o titular ainda tinha o ativo, fica fora da conferência
    depois_da_ultima_nota: bool = False
    custo_por_acao: Decimal | None = None  # bonificação: como veio do aviso aos acionistas
    referencia: str = ""  # de onde veio o custo (documento na CVM)


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
    if d.tipo is TipoEvento.OFERTA_PUBLICA:
        if d.cv != "C":
            raise erro("a oferta pública é uma compra (C)")
        if not d.valor:
            raise erro("informe o valor pago")
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
    outra_classe = (d.tipo in (TipoEvento.BONIFICACAO, TipoEvento.SUBSCRICAO)
                    and d.natureza is NaturezaEvento.OUTRA_CLASSE)  # fmt: skip
    if d.tipo is not TipoEvento.INCORPORACAO and d.natureza is not None and not outra_classe:
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
        antes = _pendencia_antes(s, titular, refs, fiscais)
        for d, evento_id in lote:
            feitos.append(_gravar(s, titular, d, evento_id, refs))
        _conferir_no_motor(s, titular, refs, fiscais, antes)
    for d, evento_id in lote:
        trilha.evento_salvo(titular, d, evento_id is not None)
    return feitos


RE_TICKER_CUSTODIA = re.compile(r"^[A-Z0-9]{4}\d{1,2}$")  # ação, unit, ETF, BDR, FII


def _gravar(
    s: Session, titular: int, d: DadosEvento, evento_id: int | None, refs: Referencias
) -> int:
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
    # código da época (CNTO3 em 2019): é o mesmo ativo do código novo (SBFG3), que é como as notas
    # podem ter entrado; sem isso, a custódia abriria um ativo separado das operações
    atual = refs.ticker_na_data(d.ticker, d.data)
    ativo = (
        (repo.ativo_existente(s, atual, d.data) if atual != d.ticker else None)
        or repo.ativo_existente(s, d.ticker, d.data)
        or repo.ativo_pelo_codigo_antigo(s, d.ticker, d.data)
    )
    outra_classe = d.origem == "b3" and d.natureza is NaturezaEvento.OUTRA_CLASSE
    if ativo is None and (outra_classe or d.tipo in (TipoEvento.CUSTODIA_INICIAL,
                                                     TipoEvento.OFERTA_PUBLICA)):  # fmt: skip
        # quem começa sem as notas antigas, compra na oferta ou recebe outra classe da companhia
        # lança ativo ainda sem negócio no banco
        if not RE_TICKER_CUSTODIA.match(d.ticker):
            raise ImportacaoError(f"{d.ticker}: código de negociação inválido (ex.: PETR4)")
        ativo = repo.Ativos(s).ativo_id(atual, d.data)
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


def _erro_do_livro(
    s: Session, titular: int, refs: Referencias, fiscais: ReferenciasFiscais
) -> MotorError | None:  # fmt: skip
    """Onde o motor para com o banco como está na sessão (None: o livro fecha): a pendência mais
    antiga das duas visões. Parar na de titular não pode esconder a de conta (evento lançado na
    conta errada, simulação da 3.1.0 com as notas reais)."""
    notas, eventos = entradas(s, refs, fiscais)
    notas = [n for n in notas if n.titular == titular]
    eventos = [e for e in eventos if e.titular == titular]
    erros: list[MotorError] = []
    for visao in (Visao.TITULAR, Visao.CONTA):
        try:
            movimentar(notas, eventos, fiscais.parametros, date.max, visao)
        except MotorError as erro:
            erros.append(erro)
    return min(erros, key=lambda e: e.data or date.min, default=None)


def _pendencia_antes(
    s: Session, titular: int, refs: Referencias, fiscais: ReferenciasFiscais
) -> MotorError | None:  # fmt: skip
    try:
        return _erro_do_livro(s, titular, refs, fiscais)
    except ReferenciaError:
        return None  # o motor recusa de novo depois da mudança, com a mensagem


def _conferir_no_motor(
    s: Session, titular: int, refs: Referencias, fiscais: ReferenciasFiscais,
    antes: MotorError | None,
) -> None:  # fmt: skip
    """A mudança não pode criar pendência no livro. Com uma pendência de antes (ex.: bonificação
    que falta noutro ativo), o livro pode continuar parando nela ou passar dela e parar numa data
    posterior (a mudança resolveu aquela, a próxima é outra); assim o usuário corrige uma por vez,
    sem a interface travar na primeira."""
    try:
        depois = _erro_do_livro(s, titular, refs, fiscais)
    except ReferenciaError as erro:
        raise ImportacaoError(f"o evento não fecha com as operações: {erro}") from erro
    if depois is None or (
        antes is not None
        and (
            str(depois) == str(antes)
            or (antes.data is not None and depois.data is not None and depois.data > antes.data)
        )
    ):
        return  # fmt: skip
    raise ImportacaoError(f"o evento não fecha com as operações: {depois}") from depois


def excluir(
    banco: Path, titular: int, evento_id: int, refs: Referencias, fiscais: ReferenciasFiscais,
    pasta: Path | None = None,
) -> None:  # fmt: skip
    """Apaga o evento se o livro continuar fechando sem ele (ex.: venda que dependia de um
    desdobramento impede apagar o desdobramento)."""
    excluir_lote(banco, titular, [evento_id], refs, fiscais, pasta=pasta)


def excluir_lote(
    banco: Path, titular: int, ids: Sequence[int], refs: Referencias, fiscais: ReferenciasFiscais,
    ignorar_ausentes: bool = False, pasta: Path | None = None,
) -> None:  # fmt: skip
    """Apaga os eventos numa transação, conferidos juntos pelo motor no fim (desfaz um lote; um
    evento do lote pode já ter sido excluído na tela: ``ignorar_ausentes``). Com ``pasta``, a
    proposta do evento apagado fica ignorada: o lançamento automático não o traz de volta."""
    apagados: list[tuple[int, str, date]] = []
    with conectar(banco) as fabrica, fabrica.begin() as s:
        antes = _pendencia_antes(s, titular, refs, fiscais)
        registros: list[EventoCorporativo] = []
        for evento_id in ids:
            registro = s.get(EventoCorporativo, evento_id)
            if registro is None and ignorar_ausentes:
                continue
            if registro is None or registro.conta.titular_id != titular:
                raise ImportacaoError(f"evento {evento_id} não encontrado")
            registros.append(registro)
        for registro in registros:
            s.delete(registro)
        s.flush()
        try:
            _conferir_no_motor(s, titular, refs, fiscais, antes)
        except ImportacaoError as erro:
            raise ImportacaoError(f"sem este evento o livro não fecha: {erro.__cause__}") from erro
        for registro in registros:
            auditoria.evento("eventos.excluido", titular=titular, tipo=str(registro.tipo),
                             data=registro.data, origem=registro.origem)  # fmt: skip
            if registro.origem == "b3":
                apagados.append((registro.ativo_id, str(registro.tipo), registro.data))
    if pasta is not None and apagados:
        _ignorar_as_apagadas(banco, titular, apagados, refs, fiscais, pasta)


def _ignorar_as_apagadas(
    banco: Path, titular: int, apagados: list[tuple[int, str, date]], refs: Referencias,
    fiscais: ReferenciasFiscais, pasta: Path,
) -> None:  # fmt: skip
    lista = propostas(banco, titular, refs, fiscais, pasta)
    with conectar(banco) as fabrica, fabrica.begin() as s:
        for ativo, tipo, dia in apagados:
            for p in lista:
                if (p.situacao == "pendente" and p.data == dia and str(p.dados.tipo) == tipo
                        and ativo in _ativos_do_ticker(s, p.dados.ticker)):  # fmt: skip
                    s.add(PropostaIgnorada(titular_id=titular, chave=p.chave,
                                           ignorada_em=datetime.now()))  # fmt: skip


# movimentação da B3 (minúsculas) -> tipo e C/V do evento proposto
PROPOSTAS: dict[str, tuple[TipoEvento, str | None]] = {
    "desdobro": (TipoEvento.DESDOBRAMENTO, "C"),
    "bonificação em ativos": (TipoEvento.BONIFICACAO, "C"),
    "restituição de capital": (TipoEvento.REDUCAO_CAPITAL, "V"),
    "recibo de subscrição": (TipoEvento.SUBSCRICAO, "C"),
    "incorporação": (TipoEvento.INCORPORACAO, None),
    "cisão": (TipoEvento.CISAO, None),
    "leilão de fração": (TipoEvento.LEILAO_FRACAO, None),
    # só em cotas de FI-Infra e FIP-IE (ver _mapa_proposta)
    "amortização": (TipoEvento.REDUCAO_CAPITAL, "V"),
}
NOTAS_PROPOSTA = {
    TipoEvento.BONIFICACAO: "informe o custo atribuído (aviso aos acionistas)",
    TipoEvento.SUBSCRICAO: "ações subscritas: informe o valor pago",
    TipoEvento.INCORPORACAO: "a B3 mostra o ativo recebido: informe o de origem e a quantidade",
    TipoEvento.CISAO: "informe o ativo cindido e o percentual do custo",
    TipoEvento.LEILAO_FRACAO: "fração vendida pela companhia: venda comum no mês do crédito",
    TipoEvento.DIREITO_NAO_EXERCIDO: "o custo pago pelos direitos comprados vira perda em "
                                     "operações comuns; os recebidos sem custo não dão resultado",
}  # fmt: skip
NAO_EXERCIDO = ("não exercid", "nao exercid")  # "Direitos de Subscrição - Não Exercido"
AMORTIZACAO_FUNDO = ("amortização de cotas de FI-Infra ou FIP-IE: reduz o custo das cotas (Lei "
                     "11.478/2007 art. 2º §2º e IN RFB 1.585/2015 art. 33, FIP-IE; Lei 12.431/2011 "
                     "art. 3º e IN art. 34, FI-Infra); o que passar do "
                     "custo é rendimento (confira o informe do administrador)")  # fmt: skip
RE_RECIBO = re.compile(r"^(?P<raiz>[A-Z]{4})(?P<classe>9|10)$")  # recibo de subscrição ON/PN
JANELA = 10  # dias entre a data da B3 (crédito) e a do evento lançado (data ex)


@auditoria.operacao("propostas")
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
        # a compra numa oferta não é o evento proposto: não esconde uma bonificação de dias depois
        lancados = [(e.ativo_id, e.ativo_destino_id, e.data) for e in s.scalars(
            select(EventoCorporativo).where(EventoCorporativo.conta_id.in_([c.id for c in contas]),
                                            EventoCorporativo.tipo != TipoEvento.OFERTA_PUBLICA)
        )]  # fmt: skip
        saida: list[Proposta] = []
        for m in s.scalars(select(MovimentacaoB3).where(MovimentacaoB3.titular_id == titular)
                           .order_by(MovimentacaoB3.data.desc(), MovimentacaoB3.id)):  # fmt: skip
            codigo = m.codigo or m.produto.split(" - ")[0].strip()
            mapa = _mapa_proposta(m, codigo, fiscais)
            if mapa is None:
                continue
            tipo, cv = mapa
            corretora = refs.corretora_b3(m.instituicao)
            das_contas = [c for c in contas if c.corretora.nome == corretora]
            conta = das_contas[0].id if len(das_contas) == 1 else None
            recibo = RE_RECIBO.match(codigo) if tipo is TipoEvento.SUBSCRICAO else None
            ticker = (f"{recibo['raiz']}{3 if recibo['classe'] == '9' else 4}" if recibo
                      else codigo)  # fmt: skip
            recebido = tipo is TipoEvento.INCORPORACAO  # a B3 mostra o crédito de B
            # restituição de capital e amortização: a quantidade da B3 é a posição sobre a qual o
            # dinheiro foi pago, não ações ou cotas canceladas; levada ao evento, zeraria a posição
            so_valor = tipo is TipoEvento.REDUCAO_CAPITAL
            data, nota_data = m.data, NOTA_DATA_EX
            if tipo is TipoEvento.DIREITO_NAO_EXERCIDO:
                data, nota_data = _fim_do_prazo(codigo, m.data, fiscais, pasta)
            elif m.movimentacao.strip().lower() == "amortização":
                nota_data = AMORTIZACAO_FUNDO
            dados = DadosEvento(
                tipo=tipo, data=data, conta=conta, ticker="" if recebido else ticker, cv=cv,
                quantidade=None if recebido or so_valor else m.quantidade,
                valor=(m.valor or Decimal(0)) if tipo in (TipoEvento.REDUCAO_CAPITAL,
                                                         TipoEvento.LEILAO_FRACAO) else Decimal(0),
                ticker_destino=ticker if recebido else None,
                quantidade_destino=m.quantidade if recebido else None,
                natureza=NaturezaEvento.INCORPORACAO_SOCIEDADE if recebido else None,
                observacao=f"proposto da movimentação da B3 ({m.movimentacao})", origem="b3",
            )  # fmt: skip
            if m.chave in ignoradas:
                situacao = "ignorada"
            elif _ja_lancado(s, refs, ticker, data, lancados):
                situacao = "lançado"
            else:
                situacao = "pendente"
            nota = "; ".join(x for x in (NOTAS_PROPOSTA.get(tipo, ""), nota_data) if x)
            saida.append(Proposta(m.chave, m.data, m.movimentacao, codigo, m.quantidade,
                                  m.valor, corretora, situacao, dados, nota))  # fmt: skip
        if fiscais is None or pasta is None:
            trilha.propostas(titular, saida)
            return _sem_pendencia_antes(saida)
        do_pacote = _do_pacote(s, titular, refs, fiscais, pasta, contas, ignoradas, lancados)
        ultima = s.scalar(select(func.max(NotaCorretagem.data_pregao)).join(Conta)
                          .where(Conta.titular_id == titular))  # fmt: skip
    if ultima is not None:
        do_pacote = [replace(p, depois_da_ultima_nota=True, segura=False) if p.data > ultima else p
                     for p in do_pacote]  # fmt: skip
    perto = timedelta(days=JANELA)

    def mesma(p: Proposta, m: Proposta) -> bool:
        janela = (timedelta(days=JANELA_RESGATE) if p.dados.tipo is TipoEvento.REDUCAO_CAPITAL
                  else perto)  # fmt: skip
        return (p.dados.tipo is m.dados.tipo and p.codigo == m.codigo
                and abs(p.data - m.data) <= janela)  # fmt: skip

    # ignorou a proposta da movimentação: a do pacote para o mesmo evento também fica ignorada
    do_pacote = [replace(p, situacao="ignorada") if p.situacao == "pendente" and any(
        mesma(p, m) and m.situacao == "ignorada" for m in saida) else p
        for p in do_pacote]  # fmt: skip
    repetidas = [m for m in saida if any(mesma(p, m) for p in do_pacote)]
    lista = sorted([*do_pacote, *(m for m in saida if m not in repetidas)],
                   key=lambda p: p.data, reverse=True)  # fmt: skip
    trilha.propostas(titular, lista)
    return _sem_pendencia_antes(lista)


FOLGA_COBERTURA = timedelta(days=45)  # movimentação que começa ou termina perto das notas basta
SO_NA_MOVIMENTACAO = ("subscrições, restituições de capital, ofertas públicas e eventos que o "
                      "pacote público não tem")  # fmt: skip


def lacuna_movimentacao(banco: Path, titular: int) -> str | None:
    """Aviso quando a movimentação da B3 importada não cobre o período das notas: só ela traz os
    eventos de cada conta, e sem ela as propostas dependem do pacote público. None se cobre."""
    with conectar(banco) as fabrica, fabrica() as s:
        notas = s.execute(select(func.min(NotaCorretagem.data_pregao),
                                 func.max(NotaCorretagem.data_pregao))
                          .join(Conta).where(Conta.titular_id == titular)).one()  # fmt: skip
        movs = s.execute(select(func.min(MovimentacaoB3.data), func.max(MovimentacaoB3.data))
                         .where(MovimentacaoB3.titular_id == titular)).one()  # fmt: skip
    primeira, ultima = notas
    if primeira is None or ultima is None:
        return None
    if movs[0] is None or movs[1] is None:
        return (
            f"Sem a movimentação da B3, {SO_NA_MOVIMENTACAO} não aparecem aqui. Importe a "
            "planilha de movimentação da Área do Investidor em Importar > Planilhas da B3."
        )
    faltas = []
    if movs[0] - primeira > FOLGA_COBERTURA:
        faltas.append(f"começa em {movs[0]:%m/%Y} e as notas em {primeira:%m/%Y}")
    if ultima - movs[1] > FOLGA_COBERTURA:
        faltas.append(f"termina em {movs[1]:%m/%Y} e as notas vão até {ultima:%m/%Y}")
    if not faltas:
        return None
    return (
        f"A movimentação da B3 importada {' e '.join(faltas)}: fora dela, "
        f"{SO_NA_MOVIMENTACAO} não aparecem aqui. Importe a movimentação do período que "
        "falta, se a Área do Investidor tiver."
    )


def _mapa_proposta(
    m: MovimentacaoB3, codigo: str, fiscais: ReferenciasFiscais | None
) -> tuple[TipoEvento, str | None] | None:
    """Tipo e C/V do evento que a linha da movimentação sugere (None: nenhum). A baixa do direito
    não exercido é uma saída; a amortização só reduz o custo em cotas de FI-Infra e FIP-IE (parecer
    do especialista de 09/11/2026; nos outros fundos fica como estava)."""
    movimentacao = m.movimentacao.strip().lower()
    if any(x in movimentacao for x in NAO_EXERCIDO):
        return None if m.entrada else (TipoEvento.DIREITO_NAO_EXERCIDO, None)
    mapa = PROPOSTAS.get(movimentacao)
    if mapa is None or (mapa[0] is TipoEvento.SUBSCRICAO and not m.entrada):
        return None
    if movimentacao == "amortização" and (
        not m.entrada or fiscais is None
        or fiscais.tipos.tipo_ou_sem(codigo) is not TipoAtivo.FUNDO_INFRA
    ):  # fmt: skip
        return None
    return mapa


JANELA_PRAZO_DIREITO = 60  # dias entre o fim do prazo de exercício e a baixa na custódia da B3


def _fim_do_prazo(
    codigo: str, baixa: date, fiscais: ReferenciasFiscais | None, pasta: Path | None
) -> tuple[date, str]:  # fmt: skip
    """Data do direito não exercido: o fim do prazo de exercício, que é decadencial (Lei 6.404
    art. 171 §4º), pelo pacote da B3 (a data de subscrição da companhia; sem ela, o fim do período
    de negociação do direito), nunca depois da baixa; senão, a data da baixa na B3. A data mais
    tarde nunca antecipa a perda (parâmetro direito_vencido_data; parecer do especialista de
    09/11/2026)."""
    sem = (baixa, "data da baixa na B3: se souber o fim do prazo de exercício (aviso aos "
                  "acionistas), use-o")  # fmt: skip
    if fiscais is None or pasta is None:
        return sem
    if fiscais.parametros.texto("direito_vencido_data", baixa) != "fim_do_prazo":
        return (baixa, "data da baixa na B3 (parâmetro direito_vencido_data)")
    fins: list[date] = []
    for ev in referencias.eventos_b3(pasta):
        if ev["tipo"] != "SUBSCRICAO" or ev["emissor"] != codigo[:4]:
            continue
        datas = [ev["data_subscricao"]] if ev["data_subscricao"] else []
        if " a " in ev["periodo"]:
            datas.append(ev["periodo"].split(" a ")[1].strip())
        try:
            fim = max(date.fromisoformat(x) if "-" in x else datetime.strptime(x, "%d/%m/%Y").date()
                      for x in datas)  # fmt: skip
        except ValueError:
            continue
        # a subscrição que terminou até JANELA_PRAZO_DIREITO dias antes da baixa (a liquidação
        # pode cair uns dias depois da baixa: a data fica na da baixa)
        if timedelta(days=-JANELA) <= baixa - fim <= timedelta(days=JANELA_PRAZO_DIREITO):
            fins.append(min(fim, baixa))
    if not fins:
        return sem
    fim = max(fins)
    return fim, (f"data: fim do prazo de subscrição pelo pacote da B3 ({fim:%d/%m/%Y}); a B3 "
                 f"baixou em {baixa:%d/%m/%Y}")  # fmt: skip


# eventos do pacote público (eventos_b3.csv: rótulos como a B3 escreve) -> tipo do evento
TIPOS_PACOTE = {
    "DESDOBRAMENTO": TipoEvento.DESDOBRAMENTO,
    "GRUPAMENTO": TipoEvento.GRUPAMENTO,
    "BONIFICACAO": TipoEvento.BONIFICACAO,
    "SUBSCRICAO": TipoEvento.SUBSCRICAO,
    # dinheiro junto de uma incorporação: "fator" é o % do custo que sai das ações recebidas,
    # "preco" é o dinheiro por ação e "data_subscricao" é o pagamento (PETZ3 -> AUAU3, resgate)
    "REDUCAO_CAPITAL": TipoEvento.REDUCAO_CAPITAL,
}
JANELA_RESGATE = 31  # dias entre a data ex e o pagamento em dinheiro que a B3 mostra
AJUSTA_SERIE = {t: t for t in (TipoEvento.DESDOBRAMENTO, TipoEvento.GRUPAMENTO,
                                TipoEvento.BONIFICACAO)}  # fmt: skip
GRUPAMENTO_EXTREMO = Decimal("0.01")  # 100 viram 1 ou mais forte
TIPOS_SEGUROS = frozenset({TipoEvento.DESDOBRAMENTO, TipoEvento.GRUPAMENTO})
AVISO_ANTERIOR = "lance antes o evento anterior deste ativo: a quantidade pode mudar"
FONTE_CUSTO_SEGURA = "cvm_ipe"  # coluna origem de eventos_b3.csv: comunicado da companhia na CVM
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


def _custos(movs: list[Movimento]) -> dict[tuple[int, int], list[tuple[date, Decimal]]]:
    """(conta, ativo) -> [(data, custo total depois)], na visão por conta, em ordem."""
    saida: dict[tuple[int, int], list[tuple[date, Decimal]]] = defaultdict(list)
    for m in movs:
        if m.conta is not None and m.ativo is not None:
            saida[(m.conta, m.ativo)].append((m.data, m.custo_apos))
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
    notas = [n for n in notas if n.titular == titular]
    evs = [e for e in evs if e.titular == titular]
    limite = date.max
    isolados: list[AtivoIsolado] = []
    try:
        # o ativo que não fecha sai do cálculo a partir do problema, e os outros seguem: um ativo
        # travado não pode esconder os eventos de toda a carteira
        movs = movimentar(notas, evs, fiscais.parametros, limite, Visao.CONTA, isolados=isolados)
    except MotorError as erro:
        if erro.data is None:
            raise
        # o livro parou (ex.: exercício com a quantidade já ajustada por uma bonificação que falta):
        # as posições até a véspera bastam para propor o evento que destrava o livro
        limite = erro.data
        movs = movimentar(notas, evs, fiscais.parametros, limite - timedelta(days=1), Visao.CONTA)
    parou_em = {i.ativo: i.desde for i in isolados}
    historico = _posicoes(movs)
    custos = _custos(movs)
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
        if data_ex > limite:
            continue  # depois de onde o livro parou: a posição na data com não é conhecida
        # a B3 escreve o grupamento como razão (0,1 = 10 viram 1) e os demais em % (200 = +200%)
        fator = Decimal(ev["fator"]) / (1 if tipo is TipoEvento.GRUPAMENTO else 100)
        ativos = _ativos_do_ticker(s, ticker)
        nomes = {ticker} | {t for a in ativos for t in s.scalars(
            select(TickerHistorico.ticker).where(TickerHistorico.ativo_id == a))}  # fmt: skip
        for (conta, ativo), hist in historico.items():
            q = _na_data(hist, data_com)
            if not q or conta not in nome_conta or data_ex > parou_em.get(ativo, date.max):
                continue  # depois de onde o ativo parou, a posição na data com não é conhecida
            serie = tickers_serie.get(ativo)
            if ativo in ativos and serie is None and tipo is TipoEvento.REDUCAO_CAPITAL:
                dados, nota = _resgate(ev, ticker, conta, q, _na_data(custos[(conta, ativo)],
                                                                      data_com), fator, data_ex,
                                       por_isin)  # fmt: skip
            elif ativo in ativos and serie is None:
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
            elif _ja_lancado(s, refs, dados.ticker, dados.data, lancados,
                             troca_conta=tipo is not TipoEvento.REDUCAO_CAPITAL):  # fmt: skip
                situacao = "lançado"
            else:
                situacao = "pendente"
            saida.append(Proposta(
                chave, dados.data, f"{ev['tipo'].capitalize()} (eventos da B3)", dados.ticker, q,
                None, nome_conta[conta], situacao, dados,
                "; ".join(x for x in (NOTA_PACOTE.format(data_com=f"{data_com:%d/%m/%Y}"), nota)
                          if x),
                segura=(serie is None and ev["isin_emitido"] in ("", ev["isin"])
                        and _segura(dados, fator, ev.get("origem", ""))),
                custo_por_acao=(Decimal(ev["custo_atribuido"]) if serie is None
                                and dados.tipo is TipoEvento.BONIFICACAO
                                and ev.get("custo_atribuido") else None),
                referencia=ev.get("referencia", "")))  # fmt: skip
    saida += _propor_restituicoes(s, refs, fiscais, pasta, por_isin, historico, nome_conta,
                                  tickers_serie, parou_em, inicio, limite, ignoradas,
                                  lancados)  # fmt: skip
    return saida


REST_CAP = "REST CAP DIN"  # restituição de capital em dinheiro nos proventos da B3
NOTA_RESTITUICAO = ("restituição de capital em dinheiro do pacote público da B3 (R$ {valor} por "
                    "ação, data com {data_com}): reduz o custo, sem resultado (IN RFB 1.585/2015 "
                    "art. 58 §8º); confira o valor creditado na movimentação da B3")  # fmt: skip


def _restituicoes(pasta: Path, refs: Referencias,
                  por_isin: dict[str, str]) -> list[tuple[str, date, Decimal]]:  # fmt: skip
    """(ticker, data com, valor por ação) das restituições de capital em dinheiro do pacote: as
    recentes (com ISIN) e as do histórico (empresa e classe); a mesma nas duas entra uma vez."""
    achadas: dict[tuple[str, str], Decimal] = {}
    for r in referencias.proventos_b3(pasta):
        ticker = por_isin.get(r["isin"])
        if r["tipo"] == REST_CAP and ticker and r.get("data_com"):
            achadas.setdefault((ticker, r["data_com"]), Decimal(r["valor"]))
    for r in referencias.proventos_historico_b3(pasta):
        digito = refs.classes.get(r["classe"])
        if r["tipo"] == REST_CAP and digito:
            achadas.setdefault((r["emissor"] + digito, r["data_com"]), Decimal(r["valor"]))
    saida: list[tuple[str, date, Decimal]] = []
    for (ticker, data_com), valor in sorted(achadas.items()):
        try:
            saida.append((ticker, date.fromisoformat(data_com), valor))
        except ValueError:
            continue
    return saida


def _propor_restituicoes(
    s: Session, refs: Referencias, fiscais: ReferenciasFiscais, pasta: Path,
    por_isin: dict[str, str], historico: dict[tuple[int, int], list[tuple[date, Decimal]]],
    nome_conta: dict[int, str], tickers_serie: dict[int, str], parou_em: dict[int, date],
    inicio: date, limite: date, ignoradas: set[str], lancados: list[tuple[int, int | None, date]],
) -> list[Proposta]:  # fmt: skip
    """Restituição de capital em dinheiro como redução de capital só com o valor (posição na data
    com vezes o valor por ação, truncado no centavo), na data ex. Pede conferência: a B3 credita
    o valor da conta, que a movimentação mostra."""
    cal = fiscais.calendario
    saida: list[Proposta] = []
    for ticker, data_com, valor in _restituicoes(pasta, refs, por_isin):
        if not cal.primeiro_ano <= data_com.year <= cal.ultimo_ano or data_com < inicio:
            continue
        try:
            data_ex = proximo_pregao(data_com, cal)
        except ReferenciaError:
            continue
        if data_ex > limite:
            continue
        ativos = _ativos_do_ticker(s, ticker)
        for (conta, ativo), hist in historico.items():
            q = _na_data(hist, data_com)
            if (
                ativo not in ativos
                or ativo in tickers_serie
                or q <= 0
                or conta not in nome_conta
                or data_ex > parou_em.get(ativo, date.max)
            ):
                continue  # fmt: skip
            total = (q * valor).quantize(Decimal("0.01"), rounding=ROUND_FLOOR)
            if total <= 0:
                continue
            dados = DadosEvento(
                tipo=TipoEvento.REDUCAO_CAPITAL, data=data_ex, conta=conta, ticker=ticker, cv="V",
                quantidade=None, valor=total, origem="b3",
                observacao="pacote público da B3: restituição de capital em dinheiro",
            )  # fmt: skip
            chave = hashlib.sha256(f"b3rest|{ticker}|{data_com}|{valor}|{conta}"
                                   .encode()).hexdigest()  # fmt: skip
            if chave in ignoradas:
                situacao = "ignorada"
            elif _ja_lancado(s, refs, ticker, data_ex, lancados, troca_conta=False):
                situacao = "lançado"
            else:
                situacao = "pendente"
            saida.append(Proposta(
                chave, data_ex, "Restituição de capital (proventos da B3)", ticker, q, total,
                nome_conta[conta], situacao, dados,
                NOTA_RESTITUICAO.format(valor=valor, data_com=f"{data_com:%d/%m/%Y}")))  # fmt: skip
    return saida


def _sem_pendencia_antes(lista: list[Proposta]) -> list[Proposta]:
    """A quantidade de uma proposta é calculada sem os eventos pendentes (100 ações, bonificação de
    10% e depois desdobramento de 100%: o desdobramento são 110, não 100). Por isso só a mais antiga
    de cada ativo é segura, e o número que o usuário confere é o que entra; lançada ela, a seguinte
    passa a ser a mais antiga, já com a quantidade certa. Empate na data também pede conferência."""
    pendentes = [p for p in lista if p.situacao == "pendente"]
    saida: list[Proposta] = []
    for p in lista:
        if p.situacao == "pendente":
            mesmas = [o for o in pendentes if o is not p and o.codigo == p.codigo]
            if any(o.data < p.data for o in mesmas):
                p = replace(p, segura=False, nota=f"{p.nota}; {AVISO_ANTERIOR}")
            elif any(o.data == p.data and (not o.segura or o.dados.conta == p.dados.conta)
                     for o in mesmas):  # fmt: skip
                p = replace(p, segura=False)
        saida.append(p)
    return saida


def _segura(d: DadosEvento, fator: Decimal, origem: str) -> bool:
    """Proposta do pacote público (já com a data ex) que não pede nada ao usuário: desdobramento e
    grupamento de ação, na conta certa, com quantidade, e bonificação com o custo por ação lido do
    documento da companhia na CVM. Fica de fora o grupamento extremo, que costuma vir seguido de
    desdobramento no mesmo dia (ex.: Itaú em 2011)."""
    if d.conta is None or not d.quantidade:
        return False
    if d.tipo is TipoEvento.BONIFICACAO:
        return bool(d.valor) and FONTE_CUSTO_SEGURA in origem
    if d.tipo not in TIPOS_SEGUROS:
        return False
    return d.tipo is not TipoEvento.GRUPAMENTO or fator > GRUPAMENTO_EXTREMO


@dataclass(frozen=True)
class ResultadoLote:
    ids: tuple[int, ...]
    recusadas: tuple[tuple[str, str], ...]  # (ativo, motivo)


def lancar_seguras(
    banco: Path, titular: int, chaves: Sequence[str], refs: Referencias,
    fiscais: ReferenciasFiscais, pasta: Path,
) -> ResultadoLote:  # fmt: skip
    """Lança de uma vez as propostas escolhidas que continuam seguras e pendentes (a tela pode estar
    atrás do banco). A que o motor recusa volta com o motivo e não trava as outras."""
    faltam = set(chaves)
    lista = [p for p in propostas(banco, titular, refs, fiscais, pasta) if p.chave in faltam]
    lote = [p for p in lista if p.segura and p.situacao == "pendente"]
    recusadas = [(p.codigo, "agora precisa de conferência: lance na lista") for p in lista
                 if p not in lote]  # fmt: skip
    ids: list[int] = []
    try:
        pedidos: list[tuple[DadosEvento, int | None]] = [(p.dados, None) for p in lote]
        ids.extend(salvar_lote(banco, titular, pedidos, refs, fiscais))
    except CoirError:  # uma não fecha: as outras entram uma a uma, e ela sai com o motivo
        for p in lote:
            try:
                ids.append(salvar(banco, titular, p.dados, refs, fiscais))
            except CoirError as erro:
                recusadas.append((p.codigo, str(erro)))
    return ResultadoLote(tuple(ids), tuple(recusadas))


def lancar_automaticos(
    banco: Path, titular: int, refs: Referencias, fiscais: ReferenciasFiscais, pasta: Path
) -> ResultadoLote:
    """Lança sozinho o que não pede nada ao usuário: as propostas seguras de ativos em carteira até
    a última nota dele. Em rodadas, porque só a mais antiga de cada ativo é segura (a seguinte
    depende da quantidade depois dela). O que o usuário apagou fica ignorado e não volta."""
    ids: list[int] = []
    recusadas: list[tuple[str, str]] = []
    tentadas: set[str] = set()
    while True:
        lista = [p.chave for p in propostas(banco, titular, refs, fiscais, pasta)
                 if p.segura and p.situacao == "pendente" and p.chave not in tentadas]  # fmt: skip
        if not lista:
            return ResultadoLote(tuple(ids), tuple(recusadas))
        tentadas.update(lista)
        rodada = lancar_seguras(banco, titular, lista, refs, fiscais, pasta)
        ids.extend(rodada.ids)
        recusadas.extend(rodada.recusadas)


def _fator(ev: dict[str, str], tipo: TipoEvento | None) -> str:
    """Como a B3 escreve: grupamento em razão (0,1 = 10 viram 1); os demais em %."""
    return f"razão {ev['fator']}" if tipo is TipoEvento.GRUPAMENTO else f"{ev['fator']}%"


def _fracao(total: Decimal) -> tuple[Decimal, str]:
    """Quantidade do evento com a fração: ela fica na posição até o leilão, que a vende pelo custo
    médio (IN RFB 1.585/2015 art. 58 caput; Lei 6.404/1976 art. 169 §3º; decisão do usuário de
    10/11/2026). Sem isso, o leilão tirava a fração das ações inteiras e a posição ficava
    fracionada."""
    inteiro = total.to_integral_value(ROUND_FLOOR)
    nota = (f"fração de {total - inteiro} fica na posição até o leilão: lance o leilão de fração "
            "quando for creditado") if total != inteiro else ""  # fmt: skip
    return total, nota


def _resgate(
    ev: dict[str, str], ticker: str, conta: int, q: Decimal, custo: Decimal, parte: Decimal,
    data_ex: date, por_isin: dict[str, str],
) -> tuple[DadosEvento | None, str]:  # fmt: skip
    """Dinheiro junto de uma incorporação: a parte do custo que era da ação resgatada sai das ações
    recebidas, sem resultado na bolsa (SC Cosit 37/2020; parecer do especialista de 07/11/2026). O
    valor do evento é esse custo, nunca o dinheiro. O resgate é alienação fora de bolsa."""
    if q <= 0:
        return None, ""
    destino = por_isin.get(ev["isin_emitido"], ticker)
    valor = (custo * parte).quantize(Decimal("0.01"), ROUND_HALF_UP)
    dinheiro = (q * Decimal(ev["preco"] or 0)).quantize(Decimal("0.01"), ROUND_HALF_UP)
    pago = date.fromisoformat(ev["data_subscricao"]) if ev["data_subscricao"] else data_ex
    nota = (
        f"o valor proposto é a parte do custo que sai de {destino} ({ev['fator']}%), não o "
        f"dinheiro. O dinheiro (R$ {dinheiro}, pago em {pago:%d/%m/%Y}) é resgate de ações: menos "
        "este custo, dá "
        "ganho ou perda fora da bolsa, no programa GCAP da Receita. O COIR não calcula esse "
        "resultado, a perda não compensa ganhos da bolsa, e há isenção até R$ 35 mil no mês"
    )
    return DadosEvento(
        tipo=TipoEvento.REDUCAO_CAPITAL, data=data_ex, conta=conta, ticker=destino, cv="V",
        quantidade=None, valor=valor, origem="b3",
        observacao=f"pacote público da B3: parte do custo na incorporação ({ev['fator']}%)",
    ), nota  # fmt: skip


def _evento_acao(
    ev: dict[str, str], tipo: TipoEvento | None, cisao: bool, ticker: str, conta: int,
    q: Decimal, fator: Decimal, data_ex: date, por_isin: dict[str, str],
) -> tuple[DadosEvento | None, str]:  # fmt: skip
    if q < 0:
        return None, ""  # venda a descoberto na data com: sem proposta (o motor falha alto)
    # subscrição e bonificação podem entregar outra classe (quem tem PN recebe ON): o crédito é na
    # classe emitida, não na ação que o titular tem
    outra = (tipo in (TipoEvento.BONIFICACAO, TipoEvento.SUBSCRICAO)
             and ev["isin_emitido"] not in ("", ev["isin"]))  # fmt: skip
    emitida = por_isin.get(ev["isin_emitido"]) if outra else None
    classe = ""
    if outra:
        classe = (f"a companhia entrega {emitida}, outra classe da ação que você tem" if emitida
                  else "a companhia entrega outra classe, que o COIR não identificou: confira o "
                       "ativo")  # fmt: skip
    base = DadosEvento(
        tipo=tipo or TipoEvento.DESDOBRAMENTO, data=data_ex, conta=conta, ticker=emitida or ticker,
        cv="C", origem="b3",
        natureza=NaturezaEvento.OUTRA_CLASSE if emitida else None,
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
        quantidade, nota = _fracao(q * fator)
        if fator <= GRUPAMENTO_EXTREMO:
            nota = "; ".join(x for x in (AVISO_GRUPAMENTO_EXTREMO, nota) if x)
        return replace(base, cv="V", quantidade=q - quantidade), nota
    quantidade, nota = _fracao(q * fator)
    if tipo is TipoEvento.SUBSCRICAO:
        quantidade = quantidade.to_integral_value(ROUND_FLOOR)  # só se subscrevem ações inteiras
        preco = Decimal(ev["preco"] or 0)
        dia = date.fromisoformat(ev["data_subscricao"]) if ev["data_subscricao"] else data_ex
        return replace(base, data=dia, quantidade=quantidade,
                       valor=(quantidade * preco).quantize(Decimal("0.01"))), (
            "; ".join(x for x in (classe, f"subscrição é opcional (preço {preco}, período "
                                          f"{ev['periodo']}): lance só o que você exerceu")
                      if x))  # fmt: skip
    if tipo is TipoEvento.BONIFICACAO:
        # o pacote traz o custo por ação quando o aviso aos acionistas informa: a proposta já sai
        # com o custo total; sem ele, o usuário informa (com custo zero o PM da venda fica errado)
        custo = Decimal(ev.get("custo_atribuido") or 0)
        if custo:
            valor = (quantidade * custo).quantize(Decimal("0.01"), ROUND_HALF_UP)
            pedido = (f"custo atribuído de R$ {custo} por ação, do aviso aos acionistas "
                      "(confira)")  # fmt: skip
            nota = "; ".join(x for x in (classe, pedido, nota) if x)
            return replace(base, quantidade=quantidade, valor=valor), nota
        nota = "; ".join(x for x in (classe, "informe o custo atribuído (aviso aos acionistas)",
                                     nota) if x)  # fmt: skip
    return replace(base, quantidade=quantidade), nota


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
    lancados: list[tuple[int, int | None, date]], troca_conta: bool = True,
) -> bool:  # fmt: skip
    """Há evento do ativo (de origem ou recebido) a até JANELA dias, ou troca de código em
    data/trocas_ticker.csv (ex.: incorporação JSLG3 -> SIMH3 tratada como troca). A repartição do
    custo de uma incorporação com dinheiro (``troca_conta=False``) vem junto com a troca."""
    perto = timedelta(days=JANELA)
    if troca_conta and any(t.ticker == ticker and abs(t.ate - dia) <= perto for t in refs.trocas):
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


def ignoradas(banco: Path, titular: int) -> set[str]:
    with conectar(banco) as fabrica, fabrica() as s:
        return set(s.scalars(select(PropostaIgnorada.chave)
                             .where(PropostaIgnorada.titular_id == titular)))  # fmt: skip


def contas_do_titular(banco: Path, titular: int) -> dict[int, str]:
    """Contas do titular para escolher na tela: id -> "XP 123456"."""
    with conectar(banco) as fabrica, fabrica() as s:
        contas = s.scalars(select(Conta).where(Conta.titular_id == titular)
                           .options(selectinload(Conta.corretora)))  # fmt: skip
        return dict(sorted(((c.id, _nome_conta(c)) for c in contas), key=lambda x: x[1]))
