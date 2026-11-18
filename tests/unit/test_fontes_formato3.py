"""Formato 3 do pacote: eventos divididos por origem, ``saidas.csv``, lista de fontes em
``pacote.json`` e a chave ``fontes.b3`` (fonte da bolsa ligada ou desligada).

- a união dos arquivos divididos é a tabela de antes (uma linha por evento, origens somadas);
- o app lê os formatos 2 e 3 com o mesmo resultado;
- com a fonte da bolsa desligada, nenhum arquivo da bolsa é aberto.
"""

import json
import shutil
from collections.abc import Iterator
from datetime import date
from pathlib import Path

import pytest

from coir import fontes, referencias
from coir.config import modo_fonte_b3
from coir.erros import ConfiguracaoError
from coir.recursos import pasta_data
from tools.referencias import gerar

ARQUIVOS_B3 = {
    "ativos_b3.csv", "emissores_b3.csv", "eventos_b3.csv", "instrumentos_b3.csv",
    "nomes_pregao_b3.csv", "opcoes_raiz_b3.csv", "proventos_b3.csv",
    "proventos_historico_b3.csv", "tipos_ativo_b3.csv",
}  # fmt: skip


@pytest.fixture
def b3_desligada() -> Iterator[None]:
    referencias.definir_fonte_b3("desligada")
    try:
        yield
    finally:
        referencias.definir_fonte_b3("pacote")


def _formato_2(destino: Path) -> Path:
    """O data/ de hoje como era no formato 2: eventos juntos e ``saidas_b3.csv``."""
    shutil.copytree(pasta_data(), destino)
    juntos = gerar.ler_eventos_pacote(destino)
    (destino / "eventos_cvm.csv").unlink()
    (destino / "eventos_b3.csv").write_text(gerar.texto_csv(gerar.COLUNAS_EVENTOS, juntos),
                                            encoding="utf-8", newline="")  # fmt: skip
    (destino / "saidas.csv").rename(destino / "saidas_b3.csv")
    (destino / "pacote.json").write_text('{"formato": 2, "data": "2026-11-11"}', "utf-8")
    return destino


def test_dividir_e_juntar_devolve_a_mesma_tabela() -> None:
    juntos = gerar.ler_eventos_pacote(pasta_data())
    cvm, b3 = gerar.dividir_eventos(juntos)
    assert gerar.juntar_divididos(cvm, b3) == juntos
    assert all(set(r["origem"].split("+")) <= {"cvm_fre", "cvm_ipe"} for r in cvm)
    assert all(r["conferido_cotahist"] == ("sim" if "cvm_fre" in r["origem"] else "nao")
               for r in cvm)  # fmt: skip
    assert {r["origem"] for r in b3} == {"b3"}
    # o documento e o custo atribuído ficam do lado da CVM
    chaves_cvm = {tuple(r[c] for c in gerar.COLUNAS_CHAVE) for r in cvm}
    assert not any(r["referencia"] or r["custo_atribuido"] for r in b3
                   if tuple(r[c] for c in gerar.COLUNAS_CHAVE) in chaves_cvm)  # fmt: skip


def test_arquivos_do_data_sao_os_da_divisao() -> None:
    """data/ está dividido como o pipeline divide (gerar de novo não muda nada)."""
    arquivos = gerar.arquivos_eventos(gerar.ler_eventos_pacote(pasta_data()))
    for nome, texto in arquivos.items():
        assert (pasta_data() / nome).read_text(encoding="utf-8") == texto, nome


def test_pacote_json_lista_os_arquivos_por_fonte() -> None:
    info = json.loads((pasta_data() / "pacote.json").read_text(encoding="utf-8"))
    assert info["formato"] == referencias.FORMATO == gerar.FORMATO == 3
    fontes_ = info["fontes"]
    assert set(fontes_["b3"]["arquivos"]) == ARQUIVOS_B3
    assert fontes_["cvm"] == {"licenca": "ODbL-1.0", "arquivos": ["eventos_cvm.csv"]}
    assert "saidas.csv" in fontes_["curadoria"]["arquivos"]
    gerado = gerar.pacote_json(pasta_data(), date.fromisoformat(info["data"]))
    assert gerado == (pasta_data() / "pacote.json").read_text(encoding="utf-8")


def test_formatos_2_e_3_dao_o_mesmo_resultado(tmp_path: Path) -> None:
    antigo = _formato_2(tmp_path / "data2")
    novo = pasta_data()
    assert fontes.eventos(antigo) == fontes.eventos(novo)
    assert fontes.saidas.__wrapped__(antigo) == fontes.saidas.__wrapped__(novo)
    assert referencias.carregar.__wrapped__(antigo) == referencias.carregar.__wrapped__(novo)


def test_pasta_efetiva_aceita_formatos_2_e_3(tmp_path: Path) -> None:
    baixada = tmp_path / referencias.PASTA_BAIXADA
    baixada.mkdir(parents=True)
    for formato, aceita in ((2, True), (3, True), (4, False)):
        (baixada / "pacote.json").write_text(
            json.dumps({"formato": formato, "data": "2099-01-01"}), "utf-8")  # fmt: skip
        assert (referencias.pasta_efetiva(tmp_path) == baixada) is aceita, formato


def test_b3_desligada_nao_abre_arquivo_da_bolsa(
    b3_desligada: None, monkeypatch: pytest.MonkeyPatch
) -> None:
    abertos: list[str] = []
    original = Path.open

    def abrir(self: Path, *args: object, **kwargs: object):  # type: ignore[no-untyped-def]
        abertos.append(self.name)
        return original(self, *args, **kwargs)  # type: ignore[arg-type]

    monkeypatch.setattr(Path, "open", abrir)
    pasta = pasta_data()
    referencias.carregar(pasta)
    referencias.carregar_fiscais(pasta)
    eventos = fontes.eventos(pasta)
    lidos = (fontes.proventos(pasta), fontes.proventos_historicos(pasta),
             fontes.instrumentos(pasta), fontes.ultimos_negocios(pasta))  # fmt: skip
    assert not any(lidos)  # só a bolsa traz esses conceitos hoje
    assert abertos and not set(abertos) & ARQUIVOS_B3
    # o que sobra é o que a CVM e a curadoria trazem
    assert eventos and all("b3" not in e.origem.split("+") for e in eventos)
    assert all(t == "" for _, _, t in fontes.emissores(pasta).values())  # só a curadoria
    # ISIN -> ticker só pela tabela curada de trocas de código
    assert set(fontes.tickers_por_isin(pasta).values()) <= {
        t for r in gerar.ler_csv(pasta / "trocas_ticker.csv")
        for t in (r["ticker"], r["ticker_impresso"])}  # fmt: skip


def test_b3_desligada_no_formato_2_nao_le_a_tabela_junta(
    b3_desligada: None, tmp_path: Path
) -> None:
    """No pacote antigo os eventos da CVM estão no arquivo da bolsa: desligada, ficam de fora."""
    assert fontes.eventos(_formato_2(tmp_path / "data2")) == []


def test_desligar_e_religar_limpa_os_caches() -> None:
    pasta = pasta_data()
    com_b3 = referencias.carregar(pasta)
    try:
        assert referencias.definir_fonte_b3("desligada") is False
        sem_b3 = referencias.carregar(pasta)
        assert len(sem_b3.ativos) < len(com_b3.ativos) and not sem_b3.nomes_historicos
        assert fontes.ultimos_negocios(pasta) == {}
    finally:
        assert referencias.definir_fonte_b3("pacote") is True
    assert referencias.carregar(pasta) == com_b3


def test_modo_local_ainda_vale_como_desligada_e_invalido_falha() -> None:
    try:
        assert referencias.definir_fonte_b3("local") is False
    finally:
        referencias.definir_fonte_b3("pacote")
    with pytest.raises(ConfiguracaoError, match=r"fontes\.b3"):
        referencias.definir_fonte_b3("outro")


def test_chave_da_configuracao(tmp_path: Path) -> None:
    config = tmp_path / "config.json"
    ambiente = {"COIR_CONFIG": str(config)}
    assert modo_fonte_b3(ambiente) == "pacote"  # sem a chave: o padrão
    config.write_text('{"fontes.b3": "desligada"}', encoding="utf-8")
    assert modo_fonte_b3(ambiente) == "desligada"


def test_estado_avisa_o_que_fica_sem_cobertura(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    from coir.servicos import estado as modulo

    monkeypatch.setattr(modulo, "modo_fonte_b3", lambda: "desligada")
    try:
        e = modulo.Estado(tmp_path / "coir.db")
        assert e.fonte_b3 is False
        assert "CVM" in (e.sem_fonte_b3("eventos") or "")
        assert "JCP" in (e.sem_fonte_b3("proventos") or "")
        monkeypatch.setattr(modulo, "modo_fonte_b3", lambda: "outro")
        e.recarregar_referencias()  # valor inválido: vale o padrão, com aviso
        assert e.fonte_b3 is True and e.sem_fonte_b3("eventos") is None
        assert "inválido" in (e.aviso_fonte_b3 or "")
    finally:
        referencias.definir_fonte_b3("pacote")


def test_resumo_da_atualizacao_sem_a_fonte_da_bolsa(b3_desligada: None, tmp_path: Path) -> None:
    """Desligada a fonte da bolsa, o resumo do pacote novo não conta nem lista os arquivos dela."""
    from coir.servicos import atualizacao

    atual, nova = tmp_path / "atual", tmp_path / "nova"
    sem_irpf = shutil.ignore_patterns("irpf")  # o pacote baixado só tem os arquivos
    shutil.copytree(pasta_data(), atual, ignore=sem_irpf)
    shutil.copytree(pasta_data(), nova, ignore=sem_irpf)
    with (nova / "eventos_b3.csv").open("a", encoding="utf-8") as f:
        f.write("FICT,BRFICTACNOR0,DESDOBRAMENTO,100,2026-11-02,,BRFICTACNOR0,,,,b3,,\n")
    with (nova / "eventos_cvm.csv").open("a", encoding="utf-8") as f:
        f.write("FICT,BRFICTACNOR0,GRUPAMENTO,0.5,2026-11-03,,BRFICTACNOR0,,,,cvm_ipe,IPE 1,,nao\n")
    mudancas = atualizacao._mudancas(atual, nova)  # pyright: ignore[reportPrivateUsage]
    assert (mudancas.eventos_novos, mudancas.arquivos) == (1, ("eventos_cvm.csv",))


def test_aviso_do_pacote_antigo_com_a_fonte_desligada(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """Pacote do formato 2 com a fonte da bolsa desligada: não há eventos da CVM a propor."""
    from coir.servicos import estado as modulo

    monkeypatch.setattr(modulo, "modo_fonte_b3", lambda: "desligada")
    e = modulo.Estado(tmp_path / "coir.db")
    e.pasta_refs = _formato_2(tmp_path / "data2")
    assert e.sem_fonte_b3("eventos") == modulo.SEM_FONTE_B3_FORMATO_2
