"""Importação em lote de notas de corretagem para o banco (regra 8).

Um lote = uma transação: qualquer erro inesperado desfaz tudo. Idempotente:
- arquivo com SHA-256 já importado por completo é ignorado sem reextrair;
- nota com chave (corretora, conta, número, data, segmento) já gravada: conteúdo igual é ignorado,
  diferente é rejeitado com o motivo;
- arquivo com nota rejeitada fica incompleto e é relido na próxima importação (ex.: depois de
  cadastrar o ativo em data/ativos.csv); as notas dele já gravadas são ignoradas.
"""

import hashlib
import logging
from collections.abc import Callable, Iterator, Sequence
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

from sqlalchemy.orm import Session

from coir.db import repositorio as repo
from coir.db.modelos import ArquivoImportado, NotaCorretagem, Titular
from coir.db.sessao import conectar, fazer_backup, migrar
from coir.domain.cpf import cpf_valido
from coir.domain.nota import Cliente, NotaExtraida, OperacaoExtraida
from coir.domain.referencias import Referencias
from coir.erros import ExtracaoError, ImportacaoError
from coir.parsing.extrair import Rejeitada, ResultadoExtracao, extrair_pdf, sem_duplicadas
from coir.parsing.normalizar import classes_das_series, operacoes

log = logging.getLogger(__name__)
Extrator = Callable[[Path], ResultadoExtracao]


@dataclass
class RelatorioImportacao:
    arquivos_lidos: list[str] = field(default_factory=list)
    arquivos_ignorados: list[str] = field(default_factory=list)  # já importados por completo
    notas_gravadas: int = 0
    operacoes_gravadas: int = 0
    notas_ja_no_banco: int = 0
    notas_renormalizadas: int = 0  # gravadas antes, com operações regravadas (data/ mudou)
    rejeitadas: list[Rejeitada] = field(default_factory=list)
    avisos: list[str] = field(default_factory=list)
    titulares_criados: list[str] = field(default_factory=list)  # "nome (id N)", nunca o CPF

    @property
    def mudou(self) -> bool:
        return self.notas_gravadas > 0 or self.notas_renormalizadas > 0


def _titular(sessao: Session, cpf: str | None) -> Titular | None:
    if cpf is not None:
        titular = repo.titular_por_cpf(sessao, cpf)
        if titular is None:
            raise ImportacaoError("titular não cadastrado (coir titular adicionar)")
        return titular
    todos = repo.titulares(sessao)
    return todos[0] if len(todos) == 1 else None


def importar_arquivos(
    sessao: Session,
    arquivos: Sequence[Path],
    refs: Referencias,
    agora: datetime,
    titular: str | None = None,
    extrator: Extrator = extrair_pdf,
) -> RelatorioImportacao:
    """Importa os arquivos numa transação da sessão (que não pode ter uma transação aberta)."""
    rel = RelatorioImportacao()
    with sessao.begin():
        contas = repo.Contas(sessao, _titular(sessao, titular), explicito=titular is not None)
        ativos = repo.Ativos(sessao)
        ativos.sincronizar_trocas(refs.trocas)

        registros: dict[str, ArquivoImportado] = {}
        lote = ResultadoExtracao()
        for caminho in arquivos:
            dados = caminho.read_bytes()
            sha = hashlib.sha256(dados).hexdigest()  # mesmo cálculo de _sha256
            registro = repo.arquivo_por_sha(sessao, sha)
            if registro is not None and registro.completo:
                rel.arquivos_ignorados.append(caminho.name)
                continue
            if caminho.name in registros:
                raise ImportacaoError(f"dois arquivos diferentes com o nome {caminho.name} no lote")
            if registro is None:
                registro = ArquivoImportado(sha256=sha, nome=caminho.name, tamanho=len(dados),
                                            importado_em=agora, completo=False)  # fmt: skip
                sessao.add(registro)
            registros[caminho.name] = registro
            rel.arquivos_lidos.append(caminho.name)
            lote.juntar(extrator(caminho))
        lote = sem_duplicadas(lote)
        rel.avisos += lote.avisos
        rel.rejeitadas += lote.rejeitadas
        clientes = {k: c for k, c in lote.clientes.items() if cpf_valido(c.cpf)}
        invalidos = sum(1 for n in lote.notas if n.chave in lote.clientes) - len(clientes)
        if invalidos:
            rel.avisos.append(f"{invalidos} nota(s) com CPF impresso inválido (ex.: anonimizado): "
                              "titular pela conta ou por --titular")  # fmt: skip
        sem_cpf = [n for n in lote.notas if n.chave not in lote.clientes]
        if sem_cpf:
            rel.avisos.append(f"{len(sem_cpf)} nota(s) sem CPF do cliente reconhecido (ex.: nota "
                              f"{sem_cpf[0].numero} de {sem_cpf[0].arquivo}): titular pela conta "
                              "ou por --titular; confira se o layout mudou")  # fmt: skip

        conhecidas = repo.series(sessao)
        novas = {s: c for s, c in classes_das_series(lote.notas, refs).items()
                 if s not in conhecidas}  # fmt: skip
        repo.gravar_series(sessao, novas)
        series = {**novas, **conhecidas}
        # antes de gravar o lote: só as notas antigas podem ter normalização desatualizada
        rel.notas_renormalizadas = _renormalizar(sessao, refs, series, ativos)

        for nota in lote.notas:
            ops = _normalizar(nota, refs, series, rel)
            if ops is not None:
                _gravar(sessao, nota, ops, registros[nota.arquivo], contas, ativos, rel,
                        clientes.get(nota.chave))  # fmt: skip
        # série de opção gravada agora com a raiz antiga: segue a troca do ativo-objeto
        ativos.sincronizar_trocas(refs.trocas)
        rel.titulares_criados = [f"id {t.id}" for t in contas.criados]  # nome: titular listar
        incompletos = {r.arquivo for r in rel.rejeitadas}
        for nome, registro in registros.items():
            registro.completo = nome not in incompletos
    log.info(
        "importação: %d arquivos lidos, %d ignorados, %d notas gravadas, %d já no banco, "
        "%d rejeitadas",
        len(rel.arquivos_lidos), len(rel.arquivos_ignorados), rel.notas_gravadas,
        rel.notas_ja_no_banco, len(rel.rejeitadas),
    )  # fmt: skip
    return rel


def _renormalizar(
    sessao: Session, refs: Referencias, series: dict[str, str], ativos: repo.Ativos
) -> int:
    """As operações são derivadas da nota gravada (json_nota) e de data/: quando a normalização
    ou as tabelas mudam, as operações das notas antigas são regravadas. O livro não muda."""
    mudadas = 0
    for registro, novas in _desatualizadas(sessao, refs, series):
        repo.substituir_operacoes(sessao, registro, novas, ativos)
        mudadas += 1
    if mudadas:
        log.info("%d nota(s) com operações renormalizadas", mudadas)
    return mudadas


def _desatualizadas(
    sessao: Session, refs: Referencias, series: dict[str, str]
) -> Iterator[tuple[NotaCorretagem, list[OperacaoExtraida]]]:
    """Notas gravadas cujas operações mudam com a normalização e as tabelas atuais."""
    for registro in repo.notas_gravadas(sessao):
        nota = NotaExtraida.model_validate_json(registro.json_nota)
        try:
            novas = operacoes(nota, refs, series)
        except ExtracaoError as erro:
            raise ImportacaoError(
                f"nota {registro.numero} de {registro.data_pregao} (já no banco) não normaliza "
                f"com as tabelas atuais: {erro}"
            ) from erro
        if novas != repo.operacoes_gravadas(registro):
            yield registro, novas


def _normalizar(
    nota: NotaExtraida, refs: Referencias, series: dict[str, str], rel: RelatorioImportacao
) -> list[OperacaoExtraida] | None:
    try:
        return operacoes(nota, refs, series)
    except ExtracaoError as erro:
        rel.rejeitadas.append(Rejeitada(nota.arquivo, nota.paginas, f"nota {nota.numero}: {erro}",
                                        nota.numero, nota.data_pregao))  # fmt: skip
        return None


def _gravar(
    sessao: Session,
    nota: NotaExtraida,
    ops: list[OperacaoExtraida],
    arquivo: ArquivoImportado,
    contas: repo.Contas,
    ativos: repo.Ativos,
    rel: RelatorioImportacao,
    cliente: Cliente | None,
) -> None:
    try:  # antes de tudo: titular errado recusa a nota
        conta = contas.conta_da_nota(nota.corretora, nota.conta, cliente)
    except repo.ContaRecusada as erro:
        rel.rejeitadas.append(Rejeitada(nota.arquivo, nota.paginas, f"nota {nota.numero}: {erro}",
                                        nota.numero, nota.data_pregao))  # fmt: skip
        return
    existente = repo.nota_por_chave(sessao, nota.chave)
    if existente is None:
        repo.gravar_nota(sessao, nota, ops, arquivo, conta, ativos)
        rel.notas_gravadas += 1
        rel.operacoes_gravadas += len(ops)
    elif existente.hash_conteudo == repo.hash_conteudo(nota):
        rel.notas_ja_no_banco += 1
        if existente.arquivo.nome != nota.arquivo:
            rel.avisos.append(f"nota {nota.numero} de {nota.arquivo} já importada de "
                              f"{existente.arquivo.nome} (idêntica); ignorada")  # fmt: skip
    elif repo.mesmo_conteudo(NotaExtraida.model_validate_json(existente.json_nota), nota):
        rel.notas_ja_no_banco += 1  # o livro fica com a nota gravada primeiro (regra 7)
        rel.avisos.append(f"nota {nota.numero} de {nota.arquivo}: mesmos negócios e valores da "
                          f"já importada de {existente.arquivo.nome}, em outro layout; "
                          "ignorada")  # fmt: skip
    else:
        motivo = (f"nota {nota.numero}: já importada de {existente.arquivo.nome} com conteúdo "
                  "diferente")  # fmt: skip
        rel.rejeitadas.append(
            Rejeitada(nota.arquivo, nota.paginas, motivo, nota.numero, nota.data_pregao)
        )


def arquivos_da_pasta(pasta: Path, rel: RelatorioImportacao, sufixo: str = ".pdf") -> list[Path]:
    arquivos: list[Path] = []
    for caminho in sorted(pasta.iterdir()):
        if caminho.is_file() and caminho.suffix.lower() == sufixo:
            arquivos.append(caminho)
        else:
            rel.avisos.append(f"ignorado (não é PDF de nota): {caminho.name}")
    return arquivos


def importar_pasta(
    banco: Path,
    pasta: Path,
    refs: Referencias,
    agora: datetime,
    titular: str | None = None,
    extrator: Extrator = extrair_pdf,
    sufixo: str = ".pdf",
) -> RelatorioImportacao:
    """Migração (com backup, se houver o que migrar) e importação de todos os arquivos da pasta.
    Backup antes de importar só quando há arquivo novo ou incompleto ou operações a regravar:
    reimportar sem mudança não gira os backups."""
    if not pasta.is_dir():
        raise ImportacaoError(f"pasta de notas não encontrada: {pasta}")
    migrar(banco, agora)
    previos = RelatorioImportacao()
    arquivos = arquivos_da_pasta(pasta, previos, sufixo)
    with conectar(banco) as fabrica:
        with fabrica() as sessao:
            pendente = any(
                (r := repo.arquivo_por_sha(sessao, _sha256(a))) is None or not r.completo
                for a in arquivos
            ) or any(True for _ in _desatualizadas(sessao, refs, repo.series(sessao)))
            # (data/ mudou: as operações de notas antigas vão ser regravadas, regra 8)
        if pendente:
            fazer_backup(banco, agora)
        with fabrica() as sessao:
            rel = importar_arquivos(sessao, arquivos, refs, agora, titular, extrator)
    rel.avisos = previos.avisos + rel.avisos
    return rel


def _sha256(caminho: Path) -> str:
    return hashlib.sha256(caminho.read_bytes()).hexdigest()
