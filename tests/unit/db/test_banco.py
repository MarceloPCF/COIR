"""F3: banco SQLite, migração, resolução de ativos e importação idempotente em lote.

Usa as notas anonimizadas de tests/fixtures/notas (texto das páginas separado por \\f) no lugar
dos PDFs: o extrator é injetado.
"""

import hashlib
from collections.abc import Callable
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
from coir.db.modelos import (
    Ativo,
    Base,
    Conta,
    EventoCorporativo,
    NotaCorretagem,
    Operacao,
    TickerHistorico,
)
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


def _com_cpf(texto: str, cpf: str) -> str:
    """Troca o CPF anonimizado das fixtures (000.000.000-00 / .00) por um CPF de teste."""
    mascara = f"{cpf[:3]}.{cpf[3:6]}.{cpf[6:9]}"
    return texto.replace("000.000.000-00", f"{mascara}-{cpf[9:]}").replace(
        "000.000.000.00", f"{mascara}.{cpf[9:]}"
    )


def extrator_cpf(cpf: str | None) -> Callable[[Path], ResultadoExtracao]:
    """Extrator das fixtures .txt com o CPF impresso trocado (None: o anonimizado, inválido)."""

    def extrair(caminho: Path) -> ResultadoExtracao:
        texto = caminho.read_text(encoding="utf-8")
        texto = texto if cpf is None else _com_cpf(texto, cpf)
        return extrair_textos(caminho.name, texto.split("\f"))

    return extrair


extrator_txt = extrator_cpf(cpf_de_teste())  # notas do titular da fábrica


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
    assert revisao_atual(engine) == revisao_final() == "0002"
    with engine.connect() as c:
        ctx = MigrationContext.configure(c, opts={"compare_type": True})
        assert compare_metadata(ctx, Base.metadata) == []


def test_banco_da_revisao_anterior_migra_com_backup(tmp_path: Path) -> None:
    """Banco da F3 (0001) com dados: migrar faz backup e chega à última revisão sem perder nada."""
    from alembic import command

    from coir.db.sessao import _config_alembic  # pyright: ignore[reportPrivateUsage]

    banco = tmp_path / "coir.db"
    engine = criar_engine(banco)
    with engine.begin() as c:
        config = _config_alembic()
        config.attributes["connection"] = c
        command.upgrade(config, "0001")
        titular = Base.metadata.tables["titular"]
        c.execute(insert(titular).values(cpf=cpf_de_teste(), nome="Titular de Teste"))
    engine.dispose()
    migrar(banco, AGORA)
    engine = criar_engine(banco)
    assert revisao_atual(engine) == "0002"
    engine.dispose()
    with abrir(banco)() as s:
        assert [t.nome for t in repo.titulares(s)] == ["Titular de Teste"]
    assert len(list((tmp_path / "backups").glob("coir-*.db"))) == 1
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
        ativos.sincronizar_trocas([t for t in REFS.trocas if t.ticker == "SIMH3"])
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


def test_serie_de_opcao_segue_a_troca_do_ativo_objeto(fabrica: sessionmaker[Session]) -> None:
    """Troca de código do ativo-objeto (ex.: EMBR3 -> EMBJ3): a série comprada com o código
    antigo e vendida com o novo é a mesma posição; a troca vale para a série (também em cadeia)
    e é idempotente."""
    with fabrica.begin() as s:
        ativos = repo.Ativos(s)
        comprada = ativos.ativo_id("AAAAE790", date(2025, 5, 20))
        vendida = ativos.ativo_id("BBBBE790", date(2026, 1, 26))
        assert comprada != vendida
        trocas = [_troca("AAAA3", "BBBB3", date(2025, 10, 31)),
                  _troca("BBBB3", "CCCC3", date(2026, 3, 31))]  # fmt: skip
        assert ativos.sincronizar_trocas(trocas)
        assert ativos.ativo_id("BBBBE790", date(2026, 1, 26)) == comprada
        assert ativos.ativo_id("CCCCE790", date(2026, 4, 1)) == comprada
        assert not ativos.sincronizar_trocas(trocas)


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
    """Notas já gravadas com o ticker antigo; a troca entra depois em trocas_ticker.csv: as
    operações (derivadas) são renormalizadas e o ativo recalculado; as notas não mudam."""
    arquivos = _todas(tmp_path / "notas")
    _importar(fabrica, arquivos)
    with fabrica() as s:
        op = s.scalars(select(Operacao).where(Operacao.mercado == "VISTA")).first()
        assert op is not None
        ticker, dia = op.ticker_impresso, s.get_one(NotaCorretagem, op.nota_id).data_pregao
        nota_id, ordem = op.nota_id, op.ordem
        notas = [n for n, _ in repo.carregar_notas(s)]
    troca = _troca(ticker, "NOVO3", dia)
    with fabrica() as s:
        rel = importar_arquivos(s, arquivos,
                                REFS.model_copy(update={"trocas": (*REFS.trocas, troca)}),
                                AGORA, extrator=extrator_txt)  # fmt: skip
    assert rel.notas_renormalizadas >= 1 and rel.mudou
    with fabrica.begin() as s:
        ativos = repo.Ativos(s)
        assert ativos.ativo_id("NOVO3", dia + timedelta(days=30)) == ativos.ativo_id(ticker, dia)
        nova = s.scalars(
            select(Operacao).where(Operacao.nota_id == nota_id, Operacao.ordem == ordem)
        ).one()
        assert (nova.papel, nova.ticker_impresso) == ("NOVO3", ticker)
        assert nova.ativo_id == ativos.ativo_id(ticker, dia)
        assert [n for n, _ in repo.carregar_notas(s)] == notas


def test_normalizacao_nova_regrava_operacoes_antigas(
    fabrica: sessionmaker[Session], tmp_path: Path
) -> None:
    """Operação gravada com normalização antiga (ex.: fracionário 'LREN3F' impresso na VISTA) é
    regravada na importação seguinte; reimportar de novo não muda mais nada."""
    arquivos = _todas(tmp_path / "notas")
    _importar(fabrica, arquivos)
    with fabrica.begin() as s:
        op = s.scalars(select(Operacao).where(Operacao.mercado == "VISTA")).first()
        assert op is not None
        certo, op.papel = op.papel, op.papel + "F"
    rel = _importar(fabrica, arquivos)
    assert rel.notas_renormalizadas == 1 and rel.notas_gravadas == 0
    with fabrica() as s:
        assert certo in {o.papel for o in s.scalars(select(Operacao))}
        assert certo + "F" not in {o.papel for o in s.scalars(select(Operacao))}
    assert not _importar(fabrica, arquivos).mudou


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


def test_mesma_nota_em_outro_layout_da_corretora_nao_e_rejeitada(
    fabrica: sessionmaker[Session], tmp_path: Path
) -> None:
    """PDF gerado de novo pela corretora: linhas de custo zeradas e ordem diferentes, mesmos
    valores. O arquivo fica completo e a nota gravada primeiro continua no livro."""
    [original] = _copiar(tmp_path / "a", "xp_2026_08_bovespa")
    _importar(fabrica, [original])
    [outro] = _copiar(tmp_path / "b", "xp_2026_08_bovespa")
    outro.write_bytes(original.read_bytes() + b"\n")

    def novo_layout(caminho: Path) -> ResultadoExtracao:
        res = extrator_txt(caminho)
        nota = res.notas[0]
        assert nota.resumo_bovespa is not None
        custos = nota.resumo_bovespa.custos
        zerada = custos[0].model_copy(update={"rotulo": "Taxa de Transf. de Ativos",
                                              "valor": Decimal(0)})  # fmt: skip
        resumo = nota.resumo_bovespa.model_copy(update={"custos": (zerada, *reversed(custos))})
        res.notas[0] = nota.model_copy(update={"resumo_bovespa": resumo})
        return res

    with fabrica() as s:
        rel = importar_arquivos(s, [outro], REFS, AGORA, extrator=novo_layout)
    assert (rel.notas_gravadas, rel.notas_ja_no_banco, rel.rejeitadas) == (0, 1, [])
    assert any("outro layout" in a for a in rel.avisos)
    with fabrica() as s:
        [registro] = s.scalars(select(NotaCorretagem)).all()
        assert registro.arquivo.nome == original.name
        arquivo = repo.arquivo_por_sha(s, hashlib.sha256(outro.read_bytes()).hexdigest())
        assert arquivo is not None and arquivo.completo


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


def test_lote_com_dois_cpfs_cria_os_titulares_pelo_cpf_da_nota(tmp_path: Path) -> None:
    """Contador com vários clientes: cada nota vai para o CPF impresso nela, sem --titular."""
    banco = tmp_path / "coir.db"
    migrar(banco, AGORA)
    f = abrir(banco)
    a, b = cpf_de_teste("111444777"), cpf_de_teste("123456789")
    [rico] = _copiar(tmp_path / "a", "rico_2015_11_bovespa")
    [bmf] = _copiar(tmp_path / "b", "btg_2023_11_bmf")
    extratores = {rico.name: extrator_cpf(a), bmf.name: extrator_cpf(b)}
    with f() as s:
        rel = importar_arquivos(s, [rico, bmf], REFS, AGORA,
                                extrator=lambda p: extratores[p.name](p))  # fmt: skip
    assert rel.rejeitadas == [] and rel.notas_gravadas == 2
    assert rel.titulares_criados == ["id 1", "id 2"]
    with f() as s:
        assert [t.nome for t in repo.titulares(s)] == ["FULANO DE TAL", "FULANO DE TAL"]
    with f() as s:
        donos = {c.corretora.nome: c.titular.cpf for c in s.scalars(select(Conta))}
    assert donos == {"RICO": a, "BTG": b}


def test_cpf_da_nota_de_outro_titular_e_recusado(
    fabrica: sessionmaker[Session], tmp_path: Path
) -> None:
    [nota] = _copiar(tmp_path / "n", "rico_2015_11_bovespa")
    _importar(fabrica, [nota])
    outro = cpf_de_teste("123456789")
    copia = tmp_path / "copia.txt"
    copia.write_bytes(nota.read_bytes() + b"\n")
    # conta já é do titular da fábrica; a cópia imprime outro CPF
    with fabrica() as s:
        rel = importar_arquivos(s, [copia], REFS, AGORA, extrator=extrator_cpf(outro))
    [rej] = rel.rejeitadas
    assert "outro titular" in rej.motivo and outro not in rej.motivo
    # --titular explícito: nota de outro CPF é recusada
    [bmf] = _copiar(tmp_path / "m", "btg_2023_11_bmf")
    with fabrica.begin() as s:
        repo.adicionar_titular(s, outro, "Segundo Titular")
    rel = _importar(fabrica, [bmf], titular=outro)
    [rej] = rel.rejeitadas
    assert "--titular" in rej.motivo and rel.notas_gravadas == 0


def test_nota_sem_cpf_valido_usa_a_conta_ou_o_titular_padrao(
    fabrica: sessionmaker[Session], tmp_path: Path
) -> None:
    """CPF impresso inválido (ex.: PDFs anonimizados): titular pela conta, pelo --titular ou pelo
    único do banco; um aviso por lote."""
    sem_cpf = extrator_cpf(None)
    [nota] = _copiar(tmp_path / "n", "rico_2015_11_bovespa")
    with fabrica() as s:
        rel = importar_arquivos(s, [nota], REFS, AGORA, extrator=sem_cpf)
    assert rel.notas_gravadas == 1 and rel.titulares_criados == []
    assert any("CPF impresso inválido" in a for a in rel.avisos)
    outro = cpf_de_teste("123456789")
    with fabrica.begin() as s:
        repo.adicionar_titular(s, outro, "Segundo Titular")
    copia = tmp_path / "copia.txt"
    copia.write_bytes(nota.read_bytes() + b"\n")
    with fabrica() as s:
        rel = importar_arquivos(s, [copia], REFS, AGORA, titular=outro, extrator=sem_cpf)
    assert "outro titular" in rel.rejeitadas[0].motivo
    [bmf] = _copiar(tmp_path / "m", "btg_2023_11_bmf")
    with fabrica() as s:  # dois titulares e conta nova sem CPF: precisa dizer de quem é
        rel = importar_arquivos(s, [bmf], REFS, AGORA, extrator=sem_cpf)
    assert "--titular" in rel.rejeitadas[0].motivo
    with pytest.raises(ImportacaoError, match="não cadastrado"):
        _importar(fabrica, [bmf], titular=cpf_de_teste("987654321"))


def test_cpf_fora_do_json_da_nota_e_do_repr(fabrica: sessionmaker[Session], tmp_path: Path) -> None:
    _importar(fabrica, _copiar(tmp_path / "n", "xp_2026_08_bovespa"))
    with fabrica() as s:
        [nota] = s.scalars(select(NotaCorretagem))
        assert cpf_de_teste() not in nota.json_nota
    res = extrator_txt(tmp_path / "n" / "xp_2026_08_bovespa.txt")
    [cliente] = res.clientes.values()
    assert cliente.cpf == cpf_de_teste() and cpf_de_teste() not in repr(cliente)
    assert "FULANO" not in repr(cliente)


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


def test_importar_pasta_faz_backup_antes_de_regravar_operacoes(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """Regra 8: data/ mudou e as operações das notas antigas vão ser regravadas (nenhum arquivo
    novo): backup antes."""
    import coir.servicos.importacao as imp

    banco = tmp_path / "dados" / "coir.db"
    migrar(banco, AGORA)
    with abrir(banco).begin() as s:
        repo.adicionar_titular(s, cpf_de_teste(), "Titular de Teste")
    pasta = tmp_path / "notas"
    _copiar(pasta, "rico_2015_11_bovespa")
    importar_pasta(banco, pasta, REFS, AGORA, extrator=extrator_txt, sufixo=".txt")
    real = imp._desatualizadas  # pyright: ignore[reportPrivateUsage]
    chamadas: list[int] = []

    def simulada(*args: object):  # type: ignore[no-untyped-def]
        chamadas.append(1)
        return iter([("nota", [])]) if len(chamadas) == 1 else real(*args)  # type: ignore[arg-type]

    monkeypatch.setattr(imp, "_desatualizadas", simulada)
    depois = datetime(2026, 10, 15, 9, 0, 0)
    importar_pasta(banco, pasta, REFS, depois, extrator=extrator_txt, sufixo=".txt")
    assert len(list((banco.parent / "backups").glob("coir-*.db"))) == 2


# ------------------------------------------------------------------------- motor fiscal (F4)
FISCAIS = referencias.carregar_fiscais()


def test_entradas_do_motor_a_partir_do_banco(
    fabrica: sessionmaker[Session], tmp_path: Path
) -> None:
    """Opção recebe o vencimento pelo prazo impresso + regra de data/ (02/23 -> 17/02/2023);
    linhas BM&F viram ajustes que somam o valor dos negócios da nota."""
    from coir.servicos.apuracao import entradas

    _importar(fabrica, _todas(tmp_path / "notas"))
    with fabrica() as s:
        notas, eventos = entradas(s, REFS, FISCAIS)
        valor_bmf = {n.id: n.valor_operacoes for n in s.scalars(select(NotaCorretagem))
                     if n.segmento == "bmf"}  # fmt: skip
    assert eventos == []
    opcoes = [o for n in notas for o in n.operacoes if o.mercado in ("CALL", "PUT")]
    assert opcoes and {o.vencimento for o in opcoes if o.ticker == "BBASN380"} == {
        date(2023, 2, 17)}  # fmt: skip
    for n in notas:
        if n.id in valor_bmf:
            assert sum(a.valor for a in n.ajustes) == valor_bmf[n.id]


def test_eventos_do_csv_entram_uma_vez_e_ticker_desconhecido_falha(
    fabrica: sessionmaker[Session], tmp_path: Path
) -> None:
    from coir.servicos.eventos import COLUNAS, importar_eventos

    _importar(fabrica, _copiar(tmp_path / "n", "rico_2015_11_bovespa"))
    with fabrica() as s:
        conta = s.scalars(select(NotaCorretagem)).one().conta.numero
    linha = dict.fromkeys(COLUNAS, "") | {
        "data": "2015-12-01", "corretora": "RICO", "conta": conta, "ticker": "CSNA3",
        "tipo": "DESDOBRAMENTO", "cv": "C", "quantidade": "100", "origem": "gabarito",
    }  # fmt: skip
    with fabrica() as s:
        assert importar_eventos(s, [linha]).novos == 1
    with fabrica() as s:
        assert importar_eventos(s, [linha]).ja_no_banco == 1
    with fabrica() as s, pytest.raises(ImportacaoError, match="sem negócio"):
        importar_eventos(s, [linha | {"ticker": "XPTO3"}])
    with fabrica() as s, pytest.raises(ImportacaoError, match="C/V"):
        importar_eventos(s, [linha | {"cv": "", "quantidade": "1"}])
    # o sentido vem de C/V: negativo, NaN e infinito falham na importação, com a linha
    for ruim in ("-100", "NaN", "Infinity"):
        with fabrica() as s, pytest.raises(ImportacaoError, match="linha 2: quantidade"):
            importar_eventos(s, [linha | {"data": "2015-12-02", "quantidade": ruim}])
    with fabrica() as s, pytest.raises(ImportacaoError, match="sem quantidade"):
        importar_eventos(s, [linha | {"data": "2015-12-02", "quantidade": ""}])
    with fabrica() as s, pytest.raises(ImportacaoError, match="redução de capital sem valor"):
        importar_eventos(s, [linha | {"data": "2015-12-02", "tipo": "REDUCAO_CAPITAL",
                                      "cv": "V", "quantidade": ""}])  # fmt: skip


def test_apuracao_do_banco_de_ponta_a_ponta(tmp_path: Path, banco: Path) -> None:
    """Notas das fixtures (sem o exercício, cuja opção não está nelas) -> motor -> meses."""
    from coir.domain.fiscal import Visao
    from coir.servicos.apuracao import apurar_banco, carteira_banco

    with abrir(banco).begin() as s:
        repo.adicionar_titular(s, cpf_de_teste(), "Titular de Teste")
    nomes = [p.stem for p in FIXTURES.glob("*.txt") if "exercicio" not in p.stem]
    with abrir(banco)() as s:
        importar_arquivos(s, _copiar(tmp_path / "n", *nomes), REFS, AGORA, extrator=extrator_txt)
    ap = apurar_banco(banco, REFS, FISCAIS, date(2026, 9, 30))
    assert ap.meses[0].ano == 2008 and ap.meses[-1].ano == 2026
    assert any(m.categoria.value == "DAY_TRADE" for m in ap.movimentos)
    # corretora de origem preservada em cada movimento da visão fiscal
    assert all(m.corretora for m in ap.movimentos if m.origem.startswith(("op:", "nota:")))
    # as três visões da carteira têm as mesmas quantidades por ativo (o custo difere por desenho)
    quantidades = []
    for visao in Visao:
        posicoes, contas = carteira_banco(banco, REFS, FISCAIS, date(2026, 9, 30), visao)
        soma: dict[tuple[int, int], Decimal] = {}
        for p in posicoes:
            soma[(p.titular, p.ativo)] = soma.get((p.titular, p.ativo), Decimal(0)) + p.quantidade
            assert (p.conta is None) is (visao is not Visao.CONTA)
            assert (p.corretora is None) is (visao is Visao.TITULAR)
        quantidades.append(soma)
        assert set(contas) and posicoes
    assert quantidades[0] == quantidades[1] == quantidades[2]


def test_eventos_repetidos_no_csv_ou_com_numero_reescrito(
    fabrica: sessionmaker[Session], tmp_path: Path
) -> None:
    from coir.servicos.eventos import COLUNAS, importar_eventos, ler_csv

    _importar(fabrica, _copiar(tmp_path / "n", "rico_2015_11_bovespa"))
    with fabrica() as s:
        conta = s.scalars(select(NotaCorretagem)).one().conta.numero
    linha = dict.fromkeys(COLUNAS, "") | {
        "data": "2015-12-01", "corretora": "RICO", "conta": conta, "ticker": "CSNA3",
        "tipo": "desdobramento", "cv": "c", "quantidade": "100",
    }  # fmt: skip
    with fabrica() as s, pytest.raises(ImportacaoError, match="repete"):
        importar_eventos(s, [linha, linha])
    with fabrica() as s:
        assert importar_eventos(s, [linha]).novos == 1
    with fabrica() as s:  # "100.0" e maiúsculas: o mesmo evento
        editada = linha | {"quantidade": "100.0", "tipo": "DESDOBRAMENTO", "observacao": "x"}
        assert importar_eventos(s, [editada]).ja_no_banco == 1
    csv_ptbr = tmp_path / "eventos.csv"
    csv_ptbr.write_text("\ufeff" + ";".join(COLUNAS) + "\n" + ";".join(linha[c] for c in COLUNAS)
                        + "\n", encoding="utf-8")  # fmt: skip
    assert ler_csv(csv_ptbr)[0]["ticker"] == "CSNA3"


def test_nota_sem_cpf_reconhecido_gera_aviso(
    fabrica: sessionmaker[Session], tmp_path: Path
) -> None:
    [nota] = _copiar(tmp_path / "n", "rico_2015_11_bovespa")
    nota.write_text(nota.read_text(encoding="utf-8").replace("000.000.000.00", "xxx"),
                    encoding="utf-8")  # fmt: skip
    with fabrica() as s:
        rel = importar_arquivos(s, [nota], REFS, AGORA, extrator=extrator_cpf(None))
    assert rel.notas_gravadas == 1 and any("sem CPF" in a for a in rel.avisos)


def test_csv_de_eventos_corrigido_substitui_o_evento_antigo(
    fabrica: sessionmaker[Session], tmp_path: Path
) -> None:
    from coir.servicos.eventos import COLUNAS, importar_eventos

    _importar(fabrica, _copiar(tmp_path / "n", "rico_2015_11_bovespa"))
    with fabrica() as s:
        conta = s.scalars(select(NotaCorretagem)).one().conta.numero
    linha = dict.fromkeys(COLUNAS, "") | {
        "data": "2015-12-01", "corretora": "RICO", "conta": conta, "ticker": "CSNA3",
        "tipo": "BONIFICACAO", "cv": "C", "quantidade": "100", "valor": "1257.99",
    }  # fmt: skip
    with fabrica() as s:
        importar_eventos(s, [linha])
    corrigida = linha | {"valor": "1003.20"}
    with fabrica() as s:
        rel = importar_eventos(s, [corrigida])
    assert (rel.novos, rel.fora_do_csv, rel.removidos) == (1, 1, 0)
    with fabrica() as s:
        rel = importar_eventos(s, [corrigida], substituir=True)
    assert (rel.novos, rel.ja_no_banco, rel.removidos) == (0, 1, 1)
    with fabrica() as s:
        [evento] = s.scalars(select(EventoCorporativo))
        assert evento.valor == Decimal("1003.20")


def test_serie_semanal_so_com_strike_antes_do_w() -> None:
    """Série semanal: strike + Wn (ex.: BOVAA120W1). Put de novembro (letra W) com strike de um
    dígito (ABEVW5) é mensal."""
    from coir.servicos.apuracao import RE_SEMANAL

    assert (m := RE_SEMANAL.search("BOVAA120W1")) is not None and m["semana"] == "1"
    assert RE_SEMANAL.search("ABEVW5") is None and RE_SEMANAL.search("PETRW290") is None
