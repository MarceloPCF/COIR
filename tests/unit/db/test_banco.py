"""F3: banco SQLite, migração, resolução de ativos e importação idempotente em lote.

Usa as notas anonimizadas de tests/fixtures/notas (texto das páginas separado por \\f) no lugar
dos PDFs: o extrator é injetado.
"""

from datetime import date, datetime, timedelta
from decimal import Decimal
from pathlib import Path

import pytest
from alembic.autogenerate import compare_metadata
from alembic.runtime.migration import MigrationContext
from hypothesis import given
from hypothesis import strategies as st
from sqlalchemy import Column, Integer, MetaData, Table, create_engine, insert, select
from sqlalchemy.orm import Session, sessionmaker

from coir import referencias
from coir.db import repositorio as repo
from coir.db.modelos import Ativo, Base, NotaCorretagem, Operacao, TickerHistorico
from coir.db.sessao import abrir, criar_engine, fazer_backup, migrar, revisao_atual, revisao_final
from coir.db.tipos import DecimalTexto
from coir.domain.referencias import TrocaTicker
from coir.erros import BancoError, ImportacaoError
from coir.parsing.extrair import ResultadoExtracao, extrair_textos
from coir.parsing.normalizar import classes_das_series, operacoes
from coir.servicos.importacao import importar_arquivos, importar_pasta
from tests.apoio import cpf_de_teste

FIXTURES = Path(__file__).parents[2] / "fixtures" / "notas"
AGORA = datetime(2026, 10, 14, 12, 0, 0)
REFS = referencias.carregar()


def extrator_txt(caminho: Path) -> ResultadoExtracao:
    return extrair_textos(caminho.name, caminho.read_text(encoding="utf-8").split("\f"))


@pytest.fixture
def banco(tmp_path: Path) -> Path:
    caminho = tmp_path / "coir.db"
    migrar(caminho, AGORA)
    return caminho


@pytest.fixture
def fabrica(banco: Path) -> sessionmaker[Session]:
    f = abrir(banco)
    with f.begin() as s:
        repo.adicionar_titular(s, cpf_de_teste(), "Titular de Teste")
    return f


def _copiar(destino: Path, *nomes: str) -> list[Path]:
    destino.mkdir(exist_ok=True)
    saida = []
    for nome in nomes:
        alvo = destino / f"{nome}.txt"
        alvo.write_bytes((FIXTURES / f"{nome}.txt").read_bytes())
        saida.append(alvo)
    return saida


def _todas(destino: Path) -> list[Path]:
    return _copiar(destino, *sorted(p.stem for p in FIXTURES.glob("*.txt")))


def _importar(fabrica: sessionmaker[Session], arquivos: list[Path], **kw: object):
    with fabrica() as s:
        return importar_arquivos(s, arquivos, REFS, AGORA, extrator=extrator_txt, **kw)  # type: ignore[arg-type]


def _contagens(fabrica: sessionmaker[Session]) -> dict[str, int]:
    with fabrica() as s:
        return repo.contagens(s)


# ------------------------------------------------------------------------------------- tipos
@given(st.decimals(min_value=Decimal("-1e12"), max_value=Decimal("1e12"), places=8))
def test_decimal_texto_ida_e_volta_exata(valor: Decimal) -> None:
    engine = create_engine("sqlite://")
    meta = MetaData()
    t = Table("t", meta, Column("id", Integer, primary_key=True), Column("v", DecimalTexto()))
    meta.create_all(engine)
    with engine.begin() as c:
        c.execute(insert(t).values(v=valor))
        lido = c.execute(select(t.c.v)).scalar_one()
    assert lido == valor and str(lido) == str(valor)


def test_decimal_texto_recusa_float() -> None:
    with pytest.raises(BancoError, match="Decimal"):
        DecimalTexto().process_bind_param(0.1, create_engine("sqlite://").dialect)  # type: ignore[arg-type]


# ---------------------------------------------------------------------------------- migração
def test_migracao_cria_esquema_igual_aos_modelos(banco: Path) -> None:
    engine = criar_engine(banco)
    assert revisao_atual(engine) == revisao_final() == "0001"
    with engine.connect() as c:
        ctx = MigrationContext.configure(c, opts={"compare_type": True})
        assert compare_metadata(ctx, Base.metadata) == []
    engine.dispose()


def test_banco_ausente_ou_desatualizado_falha_alto(tmp_path: Path) -> None:
    with pytest.raises(BancoError, match="não encontrado"):
        abrir(tmp_path / "nao-existe.db")
    vazio = tmp_path / "vazio.db"
    criar_engine(vazio).connect().close()
    with pytest.raises(BancoError, match="desatualizado"):
        abrir(vazio)


def test_backup_mantem_os_mais_recentes(banco: Path) -> None:
    for minuto in range(12):
        assert fazer_backup(banco, datetime(2026, 1, 1, 0, minuto), manter=10) is not None
    copias = sorted((banco.parent / "backups").glob("*.db"))
    assert len(copias) == 10
    assert copias[0].name == "coir-20260101-000200.db"
    assert fazer_backup(banco.parent / "outro.db", AGORA) is None


# ----------------------------------------------------------------------------------- ativos
def test_resolucao_de_ativo_por_ticker_e_data(fabrica: sessionmaker[Session]) -> None:
    with fabrica.begin() as s:
        ativos = repo.Ativos(s)
        ativos.sincronizar_trocas([TrocaTicker(ticker_impresso="JSLG3", ticker="SIMH3",
                                               ate=date(2020, 9, 18))])  # fmt: skip
        antigo = ativos.ativo_id("JSLG3", date(2015, 3, 2))
        assert ativos.ativo_id("JSLG3", date(2020, 9, 18)) == antigo
        assert ativos.ativo_id("SIMH3", date(2021, 5, 3)) == antigo
        nova_jsl = ativos.ativo_id("JSLG3", date(2021, 5, 3))
        assert nova_jsl != antigo
        assert ativos.ativo_id("JSLG3", date(2026, 1, 2)) == nova_jsl
        petr = ativos.ativo_id("PETR4", date(2008, 4, 18))
        assert ativos.ativo_id("PETR4", date(2026, 4, 18)) == petr != antigo
    with fabrica.begin() as s:  # idempotente e persistido
        ativos = repo.Ativos(s)
        ativos.sincronizar_trocas(REFS.trocas)
        assert ativos.ativo_id("SIMH3", date(2022, 1, 3)) == antigo
        assert repo.contagens(s)["ticker_historico"] == 4


def _troca(de: str, para: str, ate: date) -> TrocaTicker:
    return TrocaTicker(ticker_impresso=de, ticker=para, ate=ate)


def test_troca_em_cadeia_e_vigencia_anterior(fabrica: sessionmaker[Session]) -> None:
    with fabrica.begin() as s:
        ativos = repo.Ativos(s)
        ativos.sincronizar_trocas([_troca("AAAA3", "BBBB3", date(2010, 6, 30)),
                                   _troca("BBBB3", "CCCC3", date(2015, 6, 30))])  # fmt: skip
        a = ativos.ativo_id("AAAA3", date(2009, 1, 5))
        assert ativos.ativo_id("BBBB3", date(2012, 1, 5)) == a
        assert ativos.ativo_id("CCCC3", date(2020, 1, 6)) == a
        # ticker usado antes da vigência conhecida: outro ativo, sem sobrepor
        ativos.sincronizar_trocas([_troca("DDDD3", "EEEE3", date(2018, 1, 2))])
        antes = ativos.ativo_id("EEEE3", date(2016, 1, 4))
        assert antes != ativos.ativo_id("EEEE3", date(2019, 1, 2))
        assert ativos.ativo_id("EEEE3", date(2017, 1, 2)) == antes


def test_dois_tickers_trocados_para_o_mesmo_codigo_falha_alto(
    fabrica: sessionmaker[Session],
) -> None:
    with fabrica.begin() as s:
        ativos = repo.Ativos(s)
        with pytest.raises(BancoError, match="conflita"):
            ativos.sincronizar_trocas([_troca("AAAA3", "BBBB3", date(2012, 1, 2)),
                                       _troca("CCCC3", "BBBB3", date(2014, 1, 2))])  # fmt: skip


def test_vigencias_sobrepostas_falham_alto(fabrica: sessionmaker[Session]) -> None:
    with fabrica.begin() as s:
        s.add_all([Ativo(id=1), Ativo(id=2)])
        s.flush()
        s.add_all([TickerHistorico(ativo_id=1, ticker="XPTO3"),
                   TickerHistorico(ativo_id=2, ticker="XPTO3")])  # fmt: skip
        s.flush()
        with pytest.raises(BancoError, match="sobrepostas"):
            repo.Ativos(s).ativo_id("XPTO3", date(2020, 1, 2))


def test_troca_acrescentada_depois_reatribui_operacoes(
    fabrica: sessionmaker[Session], tmp_path: Path
) -> None:
    """Notas já gravadas com o ticker antigo; a troca entra depois em trocas_ticker.csv: o ativo
    das operações é recalculado (derivado), o livro não muda."""
    arquivos = _todas(tmp_path / "notas")
    _importar(fabrica, arquivos)
    with fabrica() as s:
        op = s.scalars(select(Operacao).where(Operacao.mercado == "VISTA")).first()
        assert op is not None
        ticker, dia = op.ticker_impresso, s.get_one(NotaCorretagem, op.nota_id).data_pregao
        livro = repo.carregar_notas(s)
    troca = _troca(ticker, "NOVO3", dia)
    with fabrica() as s:
        importar_arquivos(s, arquivos, REFS.model_copy(update={"trocas": (*REFS.trocas, troca)}),
                          AGORA, extrator=extrator_txt)  # fmt: skip
    with fabrica.begin() as s:
        ativos = repo.Ativos(s)
        assert ativos.ativo_id("NOVO3", dia + timedelta(days=30)) == ativos.ativo_id(ticker, dia)
        assert s.get_one(Operacao, op.id).ativo_id == ativos.ativo_id(ticker, dia)
        assert repo.carregar_notas(s) == livro


# -------------------------------------------------------------------------------- titulares
def test_titular_com_cpf_invalido_ou_repetido(fabrica: sessionmaker[Session]) -> None:
    with fabrica() as s:
        with pytest.raises(ImportacaoError, match="inválido"):
            repo.adicionar_titular(s, "111.111.111-11", "X")
        with pytest.raises(ImportacaoError, match="já cadastrado"):
            repo.adicionar_titular(s, cpf_de_teste(), "Outro")


# ------------------------------------------------------------------------------- importação
def test_importar_duas_vezes_nao_muda_nada(fabrica: sessionmaker[Session], tmp_path: Path) -> None:
    arquivos = _todas(tmp_path / "notas")
    rel = _importar(fabrica, arquivos)
    assert (rel.notas_gravadas, len(rel.rejeitadas), len(rel.arquivos_lidos)) == (10, 0, 10)
    antes = _contagens(fabrica)
    with fabrica() as s:
        notas_antes = repo.carregar_notas(s)
    rel2 = _importar(fabrica, arquivos)
    assert not rel2.mudou
    assert len(rel2.arquivos_ignorados) == 10 and rel2.arquivos_lidos == []
    assert _contagens(fabrica) == antes
    with fabrica() as s:
        assert repo.carregar_notas(s) == notas_antes
    assert antes["nota_corretagem"] == 10 and antes["operacao"] > 10


def test_ida_e_volta_nota_e_operacoes_iguais_as_extraidas(
    fabrica: sessionmaker[Session], tmp_path: Path
) -> None:
    arquivos = _todas(tmp_path / "notas")
    _importar(fabrica, arquivos)
    extraidas = [n for a in arquivos for n in extrator_txt(a).notas]
    series = classes_das_series(extraidas, REFS)
    esperado = {n.chave: (n, operacoes(n, REFS, series)) for n in extraidas}
    with fabrica() as s:
        lidas = repo.carregar_notas(s)
    assert {n.chave: (n, ops) for n, ops in lidas} == esperado


def test_mesma_nota_em_outro_arquivo(fabrica: sessionmaker[Session], tmp_path: Path) -> None:
    [original] = _copiar(tmp_path / "a", "rico_2015_11_bovespa")
    _importar(fabrica, [original])
    copia = tmp_path / "b" / "copia.txt"
    copia.parent.mkdir()
    copia.write_bytes(original.read_bytes() + b"\n")  # outro SHA-256, mesma nota
    rel = _importar(fabrica, [copia])
    assert (rel.notas_gravadas, rel.notas_ja_no_banco, rel.rejeitadas) == (0, 1, [])
    assert any("já importada" in a and "idêntica" in a for a in rel.avisos)


def test_mesma_chave_com_conteudo_diferente_e_rejeitada(
    fabrica: sessionmaker[Session], tmp_path: Path
) -> None:
    [original] = _copiar(tmp_path / "a", "xp_2026_08_bovespa")
    _importar(fabrica, [original])
    alterada = tmp_path / "b" / "alterada.txt"
    alterada.parent.mkdir()
    texto = original.read_text(encoding="utf-8")
    alterada.write_text(_alterar_irrelevante(texto), encoding="utf-8")
    rel = _importar(fabrica, [alterada])
    assert rel.notas_gravadas == 0
    [rej] = rel.rejeitadas
    assert "conteúdo diferente" in rej.motivo
    with fabrica() as s:
        [registro] = s.scalars(select(NotaCorretagem)).all()
        assert registro.arquivo.nome == original.name


def _alterar_irrelevante(texto: str) -> str:
    """Muda a data de liquidação impressa: a nota ainda concila, mas o conteúdo difere."""
    nota = extrair_textos("x.pdf", texto.split("\f")).notas[0]
    assert nota.resumo_bovespa is not None and nota.resumo_bovespa.data_liquidacao is not None
    impressa = f"{nota.resumo_bovespa.data_liquidacao:%d/%m/%Y}"
    outra = f"{date(2099, 1, 2):%d/%m/%Y}"
    assert impressa in texto
    return texto.replace(impressa, outra)


def test_nota_que_nao_concilia_nao_entra_e_arquivo_fica_incompleto(
    fabrica: sessionmaker[Session], tmp_path: Path
) -> None:
    boa, ruim = _copiar(tmp_path / "n", "rico_2015_11_bovespa", "agora_2008_04_bovespa")
    texto = ruim.read_text(encoding="utf-8")
    assert "Corretagem 20,00 D" in texto
    ruim.write_text(texto.replace("Corretagem 20,00 D", "Corretagem 21,00 D", 1), encoding="utf-8")
    rel = _importar(fabrica, [boa, ruim])
    assert rel.notas_gravadas == 1 and len(rel.rejeitadas) == 1
    rel2 = _importar(fabrica, [boa, ruim])  # o arquivo com rejeição é relido; o bom, não
    assert rel2.arquivos_ignorados == [boa.name] and rel2.arquivos_lidos == [ruim.name]
    ruim.write_text(texto, encoding="utf-8")  # corrigido (outro SHA): entra
    rel3 = _importar(fabrica, [boa, ruim])
    assert rel3.notas_gravadas == 1 and rel3.rejeitadas == []


def test_erro_no_meio_do_lote_desfaz_tudo(fabrica: sessionmaker[Session], tmp_path: Path) -> None:
    arquivos = _todas(tmp_path / "notas")
    antes = _contagens(fabrica)
    lidos: list[str] = []

    def extrator_que_quebra(caminho: Path) -> ResultadoExtracao:
        lidos.append(caminho.name)
        if len(lidos) == 3:
            raise OSError("disco removido")
        return extrator_txt(caminho)

    with fabrica() as s, pytest.raises(OSError, match="disco"):
        importar_arquivos(s, arquivos, REFS, AGORA, extrator=extrator_que_quebra)
    assert _contagens(fabrica) == antes


def test_conta_de_outro_titular_e_titular_ausente(
    fabrica: sessionmaker[Session], tmp_path: Path
) -> None:
    [nota] = _copiar(tmp_path / "n", "rico_2015_11_bovespa")
    _importar(fabrica, [nota])
    outro = cpf_de_teste("123456789")
    with fabrica.begin() as s:
        repo.adicionar_titular(s, outro, "Segundo Titular")
    copia = tmp_path / "copia.txt"
    copia.write_bytes(nota.read_bytes() + b"\n")
    with pytest.raises(ImportacaoError, match="outro titular"):
        _importar(fabrica, [copia], titular=outro)
    [bmf] = _copiar(tmp_path / "m", "btg_2023_11_bmf")
    with pytest.raises(ImportacaoError, match="--titular"):
        _importar(fabrica, [bmf])  # dois titulares e conta nova: precisa dizer de quem é
    with pytest.raises(ImportacaoError, match="não cadastrado"):
        _importar(fabrica, [bmf], titular=cpf_de_teste("987654321"))


def test_serie_aberta_num_lote_resolve_exercicio_no_seguinte(
    fabrica: sessionmaker[Session], tmp_path: Path
) -> None:
    refs = REFS.model_copy(update={"opcoes_raiz": {}})  # só a série aberta resolve a classe
    [abertura] = _copiar(tmp_path / "a", "btg_2023_01_bovespa_opcoes")
    [exercicio] = _copiar(tmp_path / "b", "btg_2023_01_bovespa_exercicio")
    texto = exercicio.read_text(encoding="utf-8")
    exercicio.write_text(texto.replace("KLBNM199E", "BBASN380E"), encoding="utf-8")
    with fabrica() as s:
        importar_arquivos(s, [abertura], refs, AGORA, extrator=extrator_txt)
    with fabrica() as s:
        rel = importar_arquivos(s, [exercicio], refs, AGORA, extrator=extrator_txt)
        assert rel.rejeitadas == []
        ex = next(o for n, ops in repo.carregar_notas(s) for o in ops if o.mercado == "EXERC PUT")
    assert (ex.papel, ex.ativo_objeto) == ("BBAS3", "BBASN380")


def test_importar_pasta_faz_backup_e_avisa_outros_arquivos(tmp_path: Path) -> None:
    banco = tmp_path / "dados" / "coir.db"
    migrar(banco, AGORA)
    with abrir(banco).begin() as s:
        repo.adicionar_titular(s, cpf_de_teste(), "Titular de Teste")
    pasta = tmp_path / "notas"
    _copiar(pasta, "rico_2015_11_bovespa")
    (pasta / "leia-me.pdf.bak").write_text("x", encoding="utf-8")
    rel = importar_pasta(banco, pasta, REFS, AGORA, extrator=extrator_txt, sufixo=".txt")
    assert rel.notas_gravadas == 1
    assert rel.avisos == ["ignorado (não é PDF de nota): leia-me.pdf.bak"]
    backups = banco.parent / "backups"
    assert len(list(backups.glob("coir-*.db"))) == 1
    # reimportar sem nada novo não gera backup (não gira os backups úteis)
    depois = datetime(2026, 10, 15, 9, 0, 0)
    importar_pasta(banco, pasta, REFS, depois, extrator=extrator_txt, sufixo=".txt")
    assert len(list(backups.glob("coir-*.db"))) == 1
    with pytest.raises(ImportacaoError, match="não encontrada"):
        importar_pasta(banco, tmp_path / "nada", REFS, AGORA)
