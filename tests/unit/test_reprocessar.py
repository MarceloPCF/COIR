"""Reprocessar notas já importadas (pedido do usuário de 05/11/2026): o leitor pode errar e a
corretora pode corrigir a nota. A releitura vale sobre o gravado, sem duplicar, com backup antes, e
o que ela não confirma continua gravado. Sem a opção, nada muda."""

from collections.abc import Callable
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path

import pytest
from sqlalchemy import func, select
from typer.testing import CliRunner

from coir import referencias
from coir.cli import app
from coir.config import VAR_DADOS
from coir.db.modelos import Conta, DarfPagamento, NotaAnterior, NotaCorretagem, Operacao
from coir.db.sessao import conectar
from coir.domain.nota import NotaExtraida
from coir.parsing.extrair import Rejeitada, ResultadoExtracao, extrair_pdf
from coir.servicos.importacao import RelatorioImportacao, importar_lista
from coir.servicos.notas import detalhe
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


def _guardadas(banco: Path) -> list[tuple[str, str, NotaExtraida]]:
    with conectar(banco) as fabrica, fabrica() as s:
        return [(v.motivo, v.mudou, NotaExtraida.model_validate_json(v.json_nota))
                for v in s.scalars(select(NotaAnterior))]  # fmt: skip


def _id_nota(banco: Path) -> int:
    with conectar(banco) as fabrica, fabrica() as s:
        return s.scalars(select(NotaCorretagem.id)).one()


def test_a_versao_anterior_fica_guardada_e_aparece_na_nota(tmp_path: Path, pdf: Path) -> None:
    banco = tmp_path / "coir.db"
    _importar(banco, pdf, _quantidade_errada)
    _importar(banco, pdf, reprocessar=True)
    [(motivo, mudou, antiga)] = _guardadas(banco)
    assert (motivo, mudou) == ("substituida", "negócios")
    assert antiga.negocios_bovespa[0].quantidade == Decimal(999)
    [versao] = detalhe(banco, tmp_path, _id_nota(banco)).anteriores
    assert versao.alterada_em == AGORA and versao.mudou == "negócios" and versao.arquivo == pdf.name
    assert versao.nota == antiga


def test_releitura_igual_nao_guarda_versao(tmp_path: Path, pdf: Path) -> None:
    banco = tmp_path / "coir.db"
    _importar(banco, pdf)
    _importar(banco, pdf, reprocessar=True)
    assert _guardadas(banco) == []
    assert detalhe(banco, tmp_path, _id_nota(banco)).anteriores == ()


def test_registro_removido_aparece_na_nota_certa(tmp_path: Path, pdf: Path) -> None:
    """Outra chave, mesmas páginas do mesmo arquivo: é a mesma nota lida errada."""
    banco = tmp_path / "coir.db"
    _importar(banco, pdf, lambda c: _lida_errada(c, numero="999999"))
    _importar(banco, pdf, reprocessar=True)
    [(motivo, mudou, antiga)] = _guardadas(banco)
    assert (motivo, mudou, antiga.numero) == ("removida", "conta, número ou data", "999999")
    [versao] = detalhe(banco, tmp_path, _id_nota(banco)).anteriores
    assert versao.nota.numero == "999999"


def _duas_notas(numero: str | None = None) -> Callable[[Path], ResultadoExtracao]:
    """Duas notas no arquivo: a da página do PDF (``numero``: lido errado) e outra na página 2."""

    def extrair(c: Path) -> ResultadoExtracao:
        res = extrair_pdf(c)
        [nota] = res.notas
        cliente = next(iter(res.clientes.values()))
        res.notas = [nota.model_copy(update={"numero": numero or nota.numero}),
                     nota.model_copy(update={"numero": "777777", "paginas": (2,)})]  # fmt: skip
        res.clientes = {n.chave: cliente for n in res.notas}
        return res

    return extrair


def _id_do_numero(banco: Path, numero: str) -> int:
    with conectar(banco) as fabrica, fabrica() as s:
        return s.scalars(select(NotaCorretagem.id).where(NotaCorretagem.numero == numero)).one()


def test_versao_de_outra_nota_do_arquivo_nao_aparece(tmp_path: Path, pdf: Path) -> None:
    """Mesmo arquivo, páginas diferentes: são notas diferentes."""
    banco = tmp_path / "coir.db"
    _importar(banco, pdf, _duas_notas("999999"))
    rel = _importar(banco, pdf, _duas_notas(), reprocessar=True)
    assert rel.notas_removidas == 1
    assert [g[2].numero for g in _guardadas(banco)] == ["999999"]
    assert detalhe(banco, tmp_path, _id_do_numero(banco, "777777")).anteriores == ()
    numero = extrair_pdf(pdf).notas[0].numero
    [versao] = detalhe(banco, tmp_path, _id_do_numero(banco, numero)).anteriores
    assert versao.nota.numero == "999999"


def test_pagina_pulada_na_releitura_nao_apaga_a_nota(tmp_path: Path, pdf: Path) -> None:
    """Sem rejeição e sem outra nota nas páginas dela (ex.: pulada com aviso), a nota fica."""
    banco = tmp_path / "coir.db"
    _importar(banco, pdf, _duas_notas())

    def so_a_primeira(c: Path) -> ResultadoExtracao:
        res = _duas_notas()(c)
        res.notas = res.notas[:1]
        return res

    rel = _importar(banco, pdf, so_a_primeira, reprocessar=True)
    assert rel.notas_removidas == 0 and _guardadas(banco) == []
    assert any("nota 777777" in a and "a gravada foi mantida" in a for a in rel.avisos)
    assert len(_notas(banco)) == 2


def _pagar(banco: Path, ano: int, mes: int) -> None:
    with conectar(banco) as fabrica, fabrica.begin() as s:
        titular = s.scalars(select(Conta.titular_id)).first()
        assert titular is not None
        pago_em = date(ano, mes, 1)
        s.add(DarfPagamento(titular_id=titular, ano=ano, mes=mes, data_pagamento=pago_em,
                            valor_pago=Decimal(10), observacao="",
                            registrado_em=AGORA))  # fmt: skip


def test_aviso_de_darf_paga_do_mes_da_nota_em_diante(tmp_path: Path, pdf: Path) -> None:
    """Conta o mês da nota e os seguintes (preço médio e prejuízo), não os anteriores."""
    banco = tmp_path / "coir.db"
    _importar(banco, pdf, _quantidade_errada)
    data = _notas(banco)[0].data_pregao
    _pagar(banco, data.year - 1, data.month)
    _pagar(banco, data.year, data.month)
    _pagar(banco, data.year + 1, 1)
    rel = _importar(banco, pdf, reprocessar=True)
    assert any(f"2 DARF(s) paga(s) de {data:%m/%Y} em diante" in a for a in rel.avisos)
    assert any(f"confira a apuração de {data:%m/%Y}" in a for a in rel.avisos)


def test_sem_nota_alterada_nao_ha_aviso_de_darf(tmp_path: Path, pdf: Path) -> None:
    banco = tmp_path / "coir.db"
    _importar(banco, pdf)
    data = _notas(banco)[0].data_pregao
    _pagar(banco, data.year, data.month)
    rel = _importar(banco, pdf, reprocessar=True)
    assert not any("DARF" in a for a in rel.avisos)


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
