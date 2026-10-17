"""Estado da interface: o banco, as referências carregadas uma vez e o cache dos cálculos.

A interface só apresenta: tudo o que ela mostra sai daqui (ou de outros serviços), já calculado
pelo motor. O cache é refeito quando o arquivo do banco muda (importação, evento, pagamento de
DARF, ou a CLI gravando por fora com o app aberto).
"""

import logging
from collections.abc import Callable
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from threading import Lock
from typing import TypeVar

from sqlalchemy import func, select

from coir import referencias
from coir.config import dados_dir, salvar_dados_dir
from coir.db import repositorio as repo
from coir.db.modelos import EventoCorporativo, NotaCorretagem
from coir.db.sessao import NOME_BANCO, conectar, migrar
from coir.domain.fiscal import Apuracao, Movimento, PosicaoAtivo, Visao
from coir.domain.parametros import ReferenciasFiscais
from coir.domain.referencias import Referencias
from coir.domain.relatorio import DadosRelatorio
from coir.erros import ExportacaoError
from coir.parsing.extrair import extrair_pdf
from coir.servicos import carteira, darf, notas, operacoes
from coir.servicos.apuracao import apurar_banco, carteira_banco, movimentos_banco
from coir.servicos.importacao import Extrator
from coir.servicos.relatorios import dados_banco

log = logging.getLogger(__name__)
T = TypeVar("T")


@dataclass(frozen=True)
class TitularResumo:
    id: int
    nome: str


class Estado:
    """Um por processo (o app é local e de um usuário só)."""

    def __init__(
        self, banco: Path | None = None, refs: Referencias | None = None,
        fiscais: ReferenciasFiscais | None = None,
    ) -> None:  # fmt: skip
        self.banco = banco or dados_dir() / NOME_BANCO
        self.refs = refs or referencias.carregar()
        self.fiscais = fiscais or referencias.carregar_fiscais()
        self.titular: int | None = None
        self.ano: int | None = None
        self.agora: Callable[[], datetime] = datetime.now  # relógio (fixo nos testes)
        self.extrator: Extrator = extrair_pdf  # leitura dos PDFs (textos nos testes)
        self._cache: dict[tuple[object, ...], object] = {}
        self._trava = Lock()

    # ----------------------------------------------------------------------------- banco
    @property
    def existe(self) -> bool:
        return self.banco.is_file()

    @property
    def dados(self) -> Path:
        """Pasta de dados: a do banco (as notas, backups e relatórios ficam ao lado dele)."""
        return self.banco.parent

    @property
    def pasta_notas(self) -> Path:
        return self.dados / notas.PASTA_NOTAS

    @property
    def pronto(self) -> bool:
        """Banco criado e ao menos um titular: senão, a interface abre o assistente."""
        return self.existe and bool(self.titulares())

    def trocar_dados(
        self, pasta: Path, salvar: Callable[[Path], object] = salvar_dados_dir
    ) -> None:
        """Passa a usar outra pasta de dados (e a grava para as próximas aberturas)."""
        if pasta.resolve() == self.dados.resolve():
            return
        salvar(pasta)
        self.banco = pasta / NOME_BANCO
        self.titular = self.ano = None
        self.invalidar()

    def preparar(self, agora: datetime) -> None:
        """Cria ou migra o banco (com backup antes de migrar um banco existente)."""
        self.banco.parent.mkdir(parents=True, exist_ok=True)
        migrar(self.banco, agora)
        self.invalidar()

    def _marca(self) -> tuple[int, ...]:
        """Muda sempre que o SQLite grava (arquivo principal ou o journal WAL)."""
        marca: list[int] = []
        for arquivo in (self.banco, self.banco.with_name(self.banco.name + "-wal")):
            if arquivo.exists():
                st = arquivo.stat()
                marca += [st.st_mtime_ns, st.st_size]
        return tuple(marca)

    def invalidar(self) -> None:
        with self._trava:
            self._cache.clear()

    def _guardado(self, chave: tuple[object, ...], calcular: Callable[[], T]) -> T:
        chave = (*chave, self._marca())
        with self._trava:
            if chave in self._cache:
                return self._cache[chave]  # type: ignore[return-value]
        valor = calcular()
        with self._trava:
            self._cache[chave] = valor
        return valor

    # ------------------------------------------------------------------------- consultas
    def titulares(self) -> list[TitularResumo]:
        if not self.existe:
            return []

        def ler() -> list[TitularResumo]:
            with conectar(self.banco) as fabrica, fabrica() as s:
                return [TitularResumo(t.id, t.nome) for t in repo.titulares(s)]

        return self._guardado(("titulares",), ler)

    def anos(self, hoje: date) -> list[int]:
        """Anos com nota ou evento, do mais recente ao mais antigo (inclui o ano corrente)."""
        if not self.existe:
            return [hoje.year]

        def ler() -> list[int]:
            with conectar(self.banco) as fabrica, fabrica() as s:
                datas = [
                    *s.execute(select(func.min(NotaCorretagem.data_pregao),
                                      func.max(NotaCorretagem.data_pregao))).one(),
                    *s.execute(select(func.min(EventoCorporativo.data),
                                      func.max(EventoCorporativo.data))).one(),
                ]  # fmt: skip
            validas = [d for d in datas if d is not None]
            inicio = min((d.year for d in validas), default=hoje.year)
            fim = max([hoje.year, *(d.year for d in validas)])
            return list(range(fim, inicio - 1, -1))

        return self._guardado(("anos", hoje.year), ler)

    def escolher_padrao(self, hoje: date) -> None:
        """Titular único (ou o primeiro) e o ano mais recente com dados."""
        titulares = self.titulares()
        if self.titular not in {t.id for t in titulares}:
            self.titular = titulares[0].id if titulares else None
        anos = self.anos(hoje)
        if self.ano not in anos:
            self.ano = anos[0]

    def apuracao(self, ate: date) -> Apuracao:
        """Apuração de todos os titulares até `ate` (o motor recalcula do livro inteiro)."""
        return self._guardado(
            ("apuracao", ate), lambda: apurar_banco(self.banco, self.refs, self.fiscais, ate)
        )

    def lista_notas(self) -> list[notas.ResumoNota]:
        """Notas do titular e ano escolhidos."""
        if not self.existe:
            return []
        titular, ano = self.titular, self.ano
        return self._guardado(("notas", titular, ano),
                              lambda: notas.listar(self.banco, titular, ano))  # fmt: skip

    def livro(self) -> list[operacoes.LinhaOperacao]:
        """Livro de operações do titular e ano escolhidos."""
        if not self.existe or self.ano is None:
            return []
        titular, ano = self.titular, self.ano
        return self._guardado(("operacoes", titular, ano),
                              lambda: operacoes.livro(self.banco, titular, ano))  # fmt: skip

    def carteira(self, ate: date, visao: Visao) -> tuple[list[PosicaoAtivo], dict[int, str]]:
        """Posições de todos os titulares no fim de `ate`, na visão pedida."""
        return self._guardado(
            ("carteira", ate, visao),
            lambda: carteira_banco(self.banco, self.refs, self.fiscais, ate, visao),
        )

    def movimentos(self, ate: date, visao: Visao) -> list[Movimento]:
        return self._guardado(
            ("movimentos", ate, visao),
            lambda: movimentos_banco(self.banco, self.refs, self.fiscais, ate, visao),
        )

    def notas_das_operacoes(self) -> dict[int, int]:
        return self._guardado(("op->nota",), lambda: carteira.notas_das_operacoes(self.banco))

    def fim_do_ano(self) -> date:
        """Data de referência do ano escolhido: 31/12 ou hoje, o que vier antes."""
        hoje = self.agora().date()
        ano = self.ano or hoje.year
        return min(date(ano, 12, 31), hoje)

    def dados_relatorio(self) -> DadosRelatorio:
        """Dados dos relatórios do titular e ano escolhidos (gerados hoje)."""
        if self.ano is None:
            raise ExportacaoError("escolha o ano")
        agora, titular, ano = self.agora(), self.titular, self.ano
        return self._guardado(
            ("relatorios", titular, ano, agora.date()),
            lambda: dados_banco(self.banco, self.refs, self.fiscais, ano, agora, titular),
        )

    def darfs(self) -> list[darf.LinhaDarf]:
        """DARFs do titular e ano escolhidos, com situação e multa até hoje."""
        if not self.existe or self.titular is None or self.ano is None:
            return []
        hoje, titular, ano = self.agora().date(), self.titular, self.ano

        def calcular() -> list[darf.LinhaDarf]:
            meses = [m for m in self.apuracao(self.fim_do_ano()).meses
                     if m.titular == titular and m.ano == ano]  # fmt: skip
            return darf.darfs(meses, darf.pagamentos(self.banco, titular), hoje, self.fiscais)

        return self._guardado(("darfs", titular, ano, hoje), calcular)
