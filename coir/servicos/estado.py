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
from coir.config import dados_dir, dentro_do_repositorio, salvar_dados_dir
from coir.db import repositorio as repo
from coir.db.modelos import EventoCorporativo, NotaCorretagem
from coir.db.sessao import NOME_BANCO, conectar, migrar
from coir.domain.fiscal import Apuracao, Movimento, PosicaoAtivo, Visao
from coir.domain.parametros import ReferenciasFiscais
from coir.domain.proventos import Provento
from coir.domain.referencias import Referencias
from coir.domain.relatorio import DadosRelatorio
from coir.erros import CoirError, ConfiguracaoError, ExportacaoError
from coir.parsing.extrair import extrair_pdf
from coir.servicos import (
    atualizacao,
    carteira,
    conciliacao,
    darf,
    derivativos,
    eventos,
    notas,
    operacoes,
)
from coir.servicos.apuracao import apurar_banco, carteira_banco, movimentos_banco
from coir.servicos.importacao import Extrator, aplicar_tabelas
from coir.servicos.proventos import proventos_banco
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
        self._fixas = refs is not None or fiscais is not None  # testes passam as tabelas
        pasta = referencias.pasta_efetiva(self.banco.parent)  # pacote baixado, se mais novo
        self.pasta_refs = pasta  # tabelas em uso (eventos e instrumentos da B3 também)
        self.refs = refs or referencias.carregar(pasta)
        self.fiscais = fiscais or referencias.carregar_fiscais(pasta)
        self.titular: int | None = None
        self.aviso_tabelas: str | None = None  # erro ao aplicar data/ ao abrir (preparar)
        self.ano: int | None = None
        self.agora: Callable[[], datetime] = datetime.now  # relógio (fixo nos testes)
        self.extrator: Extrator = extrair_pdf  # leitura dos PDFs (textos nos testes)
        self.baixar: atualizacao.Baixar = atualizacao.baixar  # GitHub (falso nos testes)
        self.remoto: atualizacao.PacoteRemoto | None = None  # última verificação do pacote
        self.verificado = False
        self._cache: dict[tuple[object, ...], object] = {}
        self._trava = Lock()
        self._travas: dict[tuple[object, ...], Lock] = {}

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
        self, pasta: Path, agora: datetime, salvar: Callable[[Path], object] = salvar_dados_dir
    ) -> None:
        """Passa a usar outra pasta de dados: cria ou migra o banco nela e só então a grava na
        configuração (pasta que falha não fica gravada). Nunca dentro de um repositório."""
        if dentro_do_repositorio(pasta):
            raise ConfiguracaoError(f"a pasta de dados não pode ficar num repositório: {pasta}")
        banco = pasta / NOME_BANCO
        if pasta.resolve() != self.dados.resolve():
            pasta.mkdir(parents=True, exist_ok=True)
            migrar(banco, agora)
            salvar(pasta)
            self.banco = banco
            self.titular = self.ano = None
            self.recarregar_referencias()
        self.preparar(agora)

    def recarregar_referencias(self) -> None:
        """Depois de atualizar as tabelas (ou trocar a pasta de dados): relê o pacote em uso."""
        if not self._fixas:
            pasta = referencias.pasta_efetiva(self.dados)
            self.pasta_refs = pasta
            self.refs = referencias.carregar(pasta)
            self.fiscais = referencias.carregar_fiscais(pasta)
        self.invalidar()

    def preparar(self, agora: datetime) -> None:
        """Cria ou migra o banco (com backup antes de migrar um banco existente) e aplica as
        tabelas de data/ que mudaram depois da importação (ex.: troca de ticker nova)."""
        self.banco.parent.mkdir(parents=True, exist_ok=True)
        migrar(self.banco, agora)
        self.aviso_tabelas = None
        try:
            aplicar_tabelas(self.banco, self.refs, agora)
        except CoirError as erro:  # o app abre; a próxima importação mostra o erro de novo
            log.warning("tabelas de data/ não aplicadas ao banco: %s", erro)
            self.aviso_tabelas = f"As tabelas de referência não foram aplicadas ao banco: {erro}"
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
        """Calcula uma vez por marca do banco; telas que pedem a mesma chave ao mesmo tempo
        esperam o primeiro cálculo (uma trava por chave). Entradas de marcas antigas saem."""
        marca = self._marca()
        chave = (*chave, marca)
        with self._trava:
            if chave in self._cache:
                return self._cache[chave]  # type: ignore[return-value]
            trava = self._travas.setdefault(chave, Lock())
        with trava:
            with self._trava:
                if chave in self._cache:
                    return self._cache[chave]  # type: ignore[return-value]
            valor = calcular()
            with self._trava:
                for velha in [k for k in self._cache if k[-1] != marca]:
                    del self._cache[velha]
                self._cache[chave] = valor
                self._travas.pop(chave, None)
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

    def darfs_todos_os_anos(self) -> list[darf.LinhaDarf]:
        """DARFs do titular em todos os anos (pendências do painel), com a multa até hoje."""
        if not self.existe or self.titular is None:
            return []
        hoje, titular = self.agora().date(), self.titular

        def calcular() -> list[darf.LinhaDarf]:
            meses = [m for m in self.apuracao(hoje).meses if m.titular == titular]
            return darf.darfs(meses, darf.pagamentos(self.banco, titular), hoje, self.fiscais)

        return self._guardado(("darfs-todos", titular, hoje), calcular)

    def proventos(self) -> tuple[list[Provento], list[str]]:
        """Proventos do titular escolhido (todos os anos) e os avisos."""
        if not self.existe or self.titular is None:
            return [], []
        titular = self.titular
        return self._guardado(
            ("proventos", titular), lambda: proventos_banco(self.banco, self.fiscais, titular)
        )

    def propostas(self) -> list[eventos.Proposta]:
        """Eventos sugeridos pela movimentação da B3 e pelo pacote público (tela 7); recalcula
        quando o banco muda (lançar ou ignorar grava no banco)."""
        if not self.existe or self.titular is None:
            return []
        titular, pasta = self.titular, self.pasta_refs
        return self._guardado(
            ("propostas", titular, pasta),
            lambda: eventos.propostas(self.banco, titular, self.refs, self.fiscais, pasta),
        )

    def conciliacoes(self) -> list[conciliacao.Conciliacao]:
        """B3 x COIR em cada data com posição da B3 importada (titular escolhido)."""
        if not self.existe or self.titular is None:
            return []
        titular = self.titular

        def calcular() -> list[conciliacao.Conciliacao]:
            saida: list[conciliacao.Conciliacao] = []
            anterior = None
            for dia in conciliacao.datas_b3(self.banco, titular):
                saida.append(conciliacao.conciliar(self.banco, self.refs, self.fiscais, titular,
                                                   dia, anterior))  # fmt: skip
                anterior = dia
            return saida

        return self._guardado(("conciliacoes", titular), calcular)

    def opcoes(self) -> list[derivativos.CicloOpcao]:
        """Ciclos das séries de opção do titular até o fim do ano escolhido (ou hoje)."""
        if not self.existe or self.titular is None:
            return []
        ate, titular = self.fim_do_ano(), self.titular

        def calcular() -> list[derivativos.CicloOpcao]:
            posicoes, _ = self.carteira(ate, Visao.TITULAR)
            objetos = derivativos.objetos(self.banco, self.refs)
            return derivativos.ciclos_opcoes(self.movimentos(ate, Visao.TITULAR), posicoes,
                                             titular, objetos, self.refs)  # fmt: skip

        return self._guardado(("opcoes", titular, ate), calcular)

    def futuros(self) -> derivativos.Futuros | None:
        """Futuros do titular até o fim do ano escolhido (ou hoje)."""
        if not self.existe or self.titular is None:
            return None
        ate, titular = self.fim_do_ano(), self.titular
        return self._guardado(
            ("futuros", titular, ate),
            lambda: derivativos.futuros_banco(self.banco, self.refs, self.fiscais, titular, ate),
        )
