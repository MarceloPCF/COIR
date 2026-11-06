"""Reprocessar notas já importadas (pedido do usuário de 05/11/2026): o leitor pode errar e a
corretora pode corrigir a nota. A releitura vale sobre o gravado, sem duplicar, com backup antes, e
o que ela não confirma continua gravado. Sem a opção, nada muda."""

from datetime import datetime
from decimal import Decimal
from pathlib import Path

import pytest
from sqlalchemy import func, select
from typer.testing import CliRunner

from coir import referencias
from coir.cli import app
from coir.config import VAR_DADOS
from coir.db.modelos import NotaCorretagem, Operacao
from coir.db.sessao import conectar
from coir.domain.nota import NotaExtraida
from coir.parsing.extrair import Rejeitada, ResultadoExtracao, extrair_pdf
from coir.servicos.importacao import RelatorioImportacao, importar_lista
from tools import notas_ficticias

AGORA = datetime(2026, 11, 5, 12, 0)
REFS = referencias.carregar()


@pytest.fixture
def pdf(tmp_path: Path) -> Path:
    nome, texto = notas_ficticias.textos(notas_ficticias.carregar())[0]
    return notas_ficticias.pdf(texto, tmp_path / f"{nome}.pdf")


def _lida_errada(pdf: Path, **troca: str) -> ResultadoExtracao:
    res = extrair_pdf(pdf)
    res.notas = [n.model_copy(update=troca) for n in res.notas]
    res.clientes = {n.chave: next(iter(res.clientes.values())) for n in res.notas}
    return res


def _quantidade_errada(pdf: Path) -> ResultadoExtracao:
    res = extrair_pdf(pdf)
    [nota] = res.notas
    primeiro = nota.negocios_bovespa[0].model_copy(update={"quantidade": Decimal(999)})
    errada = nota.model_copy(update={"negocios_bovespa": (primeiro, *nota.negocios_bovespa[1:])})
    res.notas, res.clientes = [errada], {errada.chave: next(iter(res.clientes.values()))}
    return res


def _notas(banco: Path) -> list[NotaExtraida]:
    with conectar(banco) as fabrica, fabrica() as s:
        return [NotaExtraida.model_validate_json(n.json_nota)
                for n in s.scalars(select(NotaCorretagem))]  # fmt: skip


def _operacoes(banco: Path) -> int:
    with conectar(banco) as fabrica, fabrica() as s:
        return s.scalar(select(func.count()).select_from(Operacao)) or 0


def _importar(
    banco: Path, pdf: Path, extrator: object = None, reprocessar: bool = False
) -> RelatorioImportacao:
    return importar_lista(banco, [pdf], REFS, AGORA, extrator=extrator,  # type: ignore[arg-type]
                          reprocessar=reprocessar)  # fmt: skip


def test_sem_a_opcao_o_arquivo_ja_importado_nao_e_relido(tmp_path: Path, pdf: Path) -> None:
    banco = tmp_path / "coir.db"
    _importar(banco, pdf, _quantidade_errada)
    rel = _importar(banco, pdf)
    assert rel.arquivos_ignorados == [pdf.name] and rel.notas_substituidas == 0
    assert _notas(banco)[0].negocios_bovespa[0].quantidade == Decimal(999)


def test_valor_lido_errado_e_corrigido_sem_duplicar(tmp_path: Path, pdf: Path) -> None:
    banco = tmp_path / "coir.db"
    _importar(banco, pdf, _quantidade_errada)
    certo = _operacoes(banco)
    rel = _importar(banco, pdf, reprocessar=True)
    assert (rel.notas_substituidas, rel.notas_gravadas, rel.notas_removidas) == (1, 0, 0)
    assert rel.mudou and any("substituída pela releitura (mudou: negócios" in a for a in rel.avisos)
    [nota] = _notas(banco)
    assert nota.negocios_bovespa == extrair_pdf(pdf).notas[0].negocios_bovespa
    assert _operacoes(banco) == certo


def test_reprocessar_faz_backup_mesmo_sem_mudanca_e_e_idempotente(
    tmp_path: Path, pdf: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    banco = tmp_path / "coir.db"
    _importar(banco, pdf)
    backups: list[Path] = []
    monkeypatch.setattr("coir.servicos.importacao.fazer_backup",
                        lambda b, _agora: backups.append(b))  # fmt: skip
    _importar(banco, pdf)
    assert backups == []
    rel = _importar(banco, pdf, reprocessar=True)
    assert backups == [banco]
    assert (rel.notas_substituidas, rel.notas_removidas, rel.notas_ja_no_banco) == (0, 0, 1)
    assert not rel.mudou and len(_notas(banco)) == 1


def test_dado_da_chave_lido_errado_nao_duplica_a_nota(tmp_path: Path, pdf: Path) -> None:
    """A conta foi lida errada: a chave nova entra e a antiga do mesmo arquivo sai."""
    banco = tmp_path / "coir.db"
    _importar(banco, pdf, lambda c: _lida_errada(c, numero="999999"))
    rel = _importar(banco, pdf, reprocessar=True)
    assert (rel.notas_gravadas, rel.notas_removidas) == (1, 1)
    assert any("registro antigo removido" in a for a in rel.avisos)
    [nota] = _notas(banco)
    assert nota.numero == extrair_pdf(pdf).notas[0].numero


def test_releitura_que_falha_mantem_a_nota_gravada(tmp_path: Path, pdf: Path) -> None:
    banco = tmp_path / "coir.db"
    _importar(banco, pdf)
    antes = _notas(banco)

    def falha(c: Path) -> ResultadoExtracao:
        return ResultadoExtracao(rejeitadas=[Rejeitada(c.name, (), "não fechou")])

    rel = _importar(banco, pdf, falha, reprocessar=True)
    assert rel.notas_removidas == 0 and len(rel.rejeitadas) == 1
    assert any("a gravada foi mantida" in a for a in rel.avisos)
    assert _notas(banco) == antes


def test_nota_corrigida_pela_corretora_em_outro_pdf(tmp_path: Path, pdf: Path) -> None:
    """Mesma chave com conteúdo diferente: sem a opção recusa e ensina; com ela, substitui."""
    banco = tmp_path / "coir.db"
    _importar(banco, pdf, _quantidade_errada)
    outro = tmp_path / "corrigida.pdf"
    outro.write_bytes(pdf.read_bytes() + b"\n%corrigida")
    recusada = _importar(banco, outro)
    assert "--reprocessar" in recusada.rejeitadas[0].motivo
    rel = _importar(banco, outro, reprocessar=True)
    assert rel.notas_substituidas == 1 and not rel.rejeitadas
    assert _notas(banco)[0].arquivo == outro.name


def test_cli_reprocessar(tmp_path: Path, pdf: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv(VAR_DADOS, str(tmp_path))
    pasta = tmp_path / "_notas-corretagens"
    pasta.mkdir()
    (pasta / pdf.name).write_bytes(pdf.read_bytes())
    runner = CliRunner()
    assert runner.invoke(app, ["banco", "migrar"]).exit_code == 0
    assert "Notas gravadas: 1" in runner.invoke(app, ["importar"]).output
    r = runner.invoke(app, ["importar", "--reprocessar"])
    assert r.exit_code == 0, r.output
    assert "Reprocessamento: 0 substituída(s), 0 removida(s)" in r.output
    assert "já importados: 0" in r.output
