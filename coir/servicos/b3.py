"""Importação das planilhas da Área do Investidor da B3 para o banco (fatos, regra 7).

As planilhas não trazem o CPF: o titular é o escolhido na importação. Exportações que se
sobrepõem (a B3 exporta períodos) não duplicam: cada linha tem chave = hash dos campos + a
ocorrência daquela linha igual no arquivo (duas linhas idênticas legítimas, como dois créditos de
dividendo iguais no mesmo dia, continuam duas). A posição de uma data é substituída pela do
arquivo mais recente. Backup do banco antes de gravar (regra 8).
"""

import hashlib
import logging
from collections import Counter
from collections.abc import Callable, Sequence
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

from pydantic import BaseModel
from sqlalchemy import delete, select
from sqlalchemy.orm import Session

from coir.db.modelos import ArquivoB3, MovimentacaoB3, NegociacaoB3, PosicaoB3, Titular
from coir.db.sessao import conectar, fazer_backup, migrar
from coir.domain.b3 import TipoPlanilha
from coir.erros import ImportacaoError
from coir.importers import b3

log = logging.getLogger(__name__)
PASTA_B3 = "_relatorios-b3"  # padrão das planilhas da B3 dentro da pasta de dados


@dataclass(frozen=True)
class ResultadoPlanilha:
    arquivo: str
    tipo: TipoPlanilha | None  # None: não é planilha da B3 reconhecida
    linhas: int = 0
    novas: int = 0
    detalhe: str = ""


@dataclass
class RelatorioB3:
    arquivos: list[ResultadoPlanilha] = field(default_factory=list[ResultadoPlanilha])

    @property
    def novas(self) -> int:
        return sum(a.novas for a in self.arquivos)


def _chaves(titular: int, tipo: TipoPlanilha, linhas: Sequence[BaseModel]) -> list[str]:
    vistas: Counter[str] = Counter()
    chaves: list[str] = []
    for linha in linhas:
        base = f"{titular}|{tipo}|{linha.model_dump_json()}"
        vistas[base] += 1
        chaves.append(hashlib.sha256(f"{base}|{vistas[base]}".encode()).hexdigest())
    return chaves


def importar_planilhas(
    banco: Path, arquivos: Sequence[Path], titular: int, agora: datetime,
    progresso: Callable[[int, int, str], None] | None = None,
) -> RelatorioB3:  # fmt: skip
    """Importa as planilhas numa transação; arquivo já importado (mesmo SHA-256) é ignorado."""
    migrar(banco, agora)
    fazer_backup(banco, agora)
    rel = RelatorioB3()
    with conectar(banco) as fabrica, fabrica.begin() as s:
        if s.get(Titular, titular) is None:
            raise ImportacaoError(f"titular {titular} não cadastrado")
        for feitos, caminho in enumerate(arquivos):
            if progresso is not None:
                progresso(feitos, len(arquivos), caminho.name)
            rel.arquivos.append(_importar(s, caminho, titular, agora))
        if progresso is not None:
            progresso(len(arquivos), len(arquivos), "")
    log.info("planilhas B3: %d arquivo(s), %d linha(s) nova(s)", len(arquivos), rel.novas)
    return rel


def _importar(s: Session, caminho: Path, titular: int, agora: datetime) -> ResultadoPlanilha:
    try:
        dados = caminho.read_bytes()
    except OSError as erro:
        raise ImportacaoError(f"não foi possível ler {caminho.name}: {erro}") from erro
    sha = hashlib.sha256(dados).hexdigest()
    if s.scalar(select(ArquivoB3).where(ArquivoB3.sha256 == sha)) is not None:
        return ResultadoPlanilha(caminho.name, None, detalhe="já importado")
    try:
        tipo = b3.tipo(caminho)
    except ImportacaoError as erro:
        return ResultadoPlanilha(caminho.name, None, detalhe=str(erro))
    registro = ArquivoB3(sha256=sha, nome=caminho.name, tipo=tipo.value, titular_id=titular,
                         importado_em=agora, linhas=0)  # fmt: skip
    s.add(registro)
    s.flush()
    if tipo is TipoPlanilha.POSICAO:
        posicoes = b3.posicao(caminho)
        datas = {p.data for p in posicoes} or {b3.data_do_consolidado(caminho)}
        s.execute(delete(PosicaoB3).where(PosicaoB3.titular_id == titular,
                                          PosicaoB3.data.in_(datas)))  # fmt: skip
        s.add_all(PosicaoB3(arquivo_id=registro.id, titular_id=titular, **p.model_dump())
                  for p in posicoes)  # fmt: skip
        registro.linhas = len(posicoes)
        quando = ", ".join(f"{d:%d/%m/%Y}" for d in sorted(datas))
        return ResultadoPlanilha(caminho.name, tipo, len(posicoes), len(posicoes),
                                 f"posição em {quando}")  # fmt: skip
    linhas = b3.movimentacao(caminho) if tipo is TipoPlanilha.MOVIMENTACAO else b3.negociacao(
        caminho)  # fmt: skip
    modelo = MovimentacaoB3 if tipo is TipoPlanilha.MOVIMENTACAO else NegociacaoB3
    chaves = _chaves(titular, tipo, linhas)
    existentes = set(s.scalars(select(modelo.chave).where(modelo.chave.in_(chaves))))
    novas = [(c, x) for c, x in zip(chaves, linhas, strict=True) if c not in existentes]
    s.add_all(modelo(chave=c, arquivo_id=registro.id, titular_id=titular, **x.model_dump())
              for c, x in novas)  # fmt: skip
    registro.linhas = len(linhas)
    datas = [x.data for x in linhas]
    periodo = f"{min(datas):%d/%m/%Y} a {max(datas):%d/%m/%Y}" if datas else "vazio"
    return ResultadoPlanilha(caminho.name, tipo, len(linhas), len(novas), periodo)
