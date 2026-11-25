"""F7b: "Atualizar tabelas" — o app só usa o pacote público se cada arquivo bater com o SHA-256 do
manifesto, o formato for um dos que esta versão lê e as tabelas carregarem; sem rede, nada muda."""

import json
import shutil
import zipfile
from datetime import date
from pathlib import Path

import pytest

from coir import fontes, referencias
from coir.erros import AtualizacaoError
from coir.recursos import pasta_data
from coir.servicos import atualizacao, versao
from tools.referencias import pacote
from tools.referencias.gerar import COLUNAS_EVENTOS, ler_eventos_pacote

URL_ZIP = "https://github.com/MarceloPCF/COIR/releases/download/referencias-2099.01.01/p.zip"


def _pacote(tmp: Path, data: str = "2099-01-01", formato: int = referencias.FORMATO,
            nomes_pregao: str = "") -> bytes:  # fmt: skip
    pasta = tmp / "pacote-data"
    shutil.copytree(pasta_data(), pasta)
    if nomes_pregao:  # linhas a mais em nomes_pregao_b3.csv
        with (pasta / "nomes_pregao_b3.csv").open("a", encoding="utf-8", newline="") as f:
            f.write(nomes_pregao)
    (pasta / "pacote.json").write_text(json.dumps({"formato": referencias.FORMATO, "data": data}))
    if not (pasta / "eventos_b3.csv").is_file():
        cabecalho = ",".join(COLUNAS_EVENTOS) + "\n"
        (pasta / "eventos_b3.csv").write_text(cabecalho, encoding="utf-8")
    with (pasta / "eventos_b3.csv").open("a", encoding="utf-8") as f:
        f.write("ZZZZ,BRZZZZACNOR0,DESDOBRAMENTO,100,2098-12-01,2098-11-01,BRZZZZACNOR0,,,\n")
    zip_ = pacote.montar(pasta, tmp / "dist")
    if formato != referencias.FORMATO:  # manifesto de outra versão do esquema
        with zipfile.ZipFile(zip_) as z:
            conteudo = {n: z.read(n) for n in z.namelist()}
        m = json.loads(conteudo["manifesto.json"])
        m["formato"] = formato
        conteudo["manifesto.json"] = json.dumps(m).encode()
        with zipfile.ZipFile(zip_, "w") as z:
            for n, b in conteudo.items():
                z.writestr(n, b)
    return zip_.read_bytes()


PAGINA_V3 = "https://github.com/MarceloPCF/COIR/releases/tag/v99.0.0"


def _github(zip_: bytes) -> atualizacao.Baixar:
    releases = [
        {"tag_name": "v99.0.0", "draft": False, "prerelease": False, "assets": []},
        {"tag_name": "referencias-2099.01.01", "draft": False, "prerelease": False,
         "assets": [{"name": "referencias-2099.01.01.zip", "browser_download_url": URL_ZIP}]},
        {"tag_name": "referencias-2099.02.01", "draft": True, "prerelease": False,
         "assets": [{"name": "x.zip", "browser_download_url": "https://x"}]},
    ]  # fmt: skip

    def obter(url: str) -> bytes:
        if url == atualizacao.RELEASES:
            return json.dumps(releases).encode()
        if url == versao.ULTIMO:  # o último release é o do programa
            return json.dumps(releases[0] | {"html_url": PAGINA_V3}).encode()
        assert url == URL_ZIP
        return zip_

    return obter


def test_verifica_baixa_confere_e_aplica(tmp_path: Path) -> None:
    dados = tmp_path / "dados"
    dados.mkdir()
    obter = _github(_pacote(tmp_path))
    remoto = atualizacao.verificar(obter)
    assert remoto is not None and remoto.nome == "referencias-2099.01.01"  # rascunho ignorado
    assert atualizacao.mais_novo(dados, remoto)
    mudancas = atualizacao.preparar(dados, remoto, obter)
    assert mudancas.eventos_novos == 1 and "eventos_b3.csv" in mudancas.arquivos
    assert referencias.pasta_efetiva(dados) == pasta_data()  # preparado, ainda não em uso
    atualizacao.aplicar(dados)
    s = atualizacao.situacao(dados)
    assert s.baixado and s.em_uso is not None and s.em_uso.year == 2099
    assert not atualizacao.mais_novo(dados, remoto)
    assert referencias.carregar(referencias.pasta_efetiva(dados)).ativos  # carrega
    with pytest.raises(AtualizacaoError, match="não há pacote anterior"):
        atualizacao.voltar(dados)  # o primeiro pacote baixado não tem anterior
    atualizacao.preparar(dados, remoto, obter)  # o mesmo de novo: o atual vira o anterior
    atualizacao.aplicar(dados)
    assert atualizacao.situacao(dados).anterior is not None
    atualizacao.voltar(dados)
    assert atualizacao.situacao(dados).baixado


def test_tabelas_lidas_seguem_o_pacote_aplicado_e_o_anterior(tmp_path: Path) -> None:
    """O pacote em uso fica sempre em <dados>/referencias/atual: aplicar outro (ou voltar ao
    anterior) troca o conteúdo no mesmo caminho, e o que foi lido do pacote antigo (último
    negócio de cada ticker, ativos que saíram, eventos) não pode valer até reabrir o programa."""
    dados = tmp_path / "dados"
    dados.mkdir()
    primeiro = _github(_pacote(tmp_path / "1"))
    novo_ticker = "ZZZZ TESTE ON,ZZZZ3,2098-01-02,2098-12-30,\n"
    segundo = _github(_pacote(tmp_path / "2", nomes_pregao=novo_ticker))
    remoto = atualizacao.verificar(primeiro)
    assert remoto is not None
    atualizacao.preparar(dados, remoto, primeiro)
    atualizacao.aplicar(dados)
    atual = referencias.pasta_efetiva(dados)
    assert "ZZZZ3" not in fontes.ultimos_negocios(atual)
    atualizacao.preparar(dados, remoto, segundo)
    atualizacao.aplicar(dados)
    assert referencias.pasta_efetiva(dados) == atual  # mesmo caminho, outro conteúdo
    assert fontes.ultimos_negocios(atual)["ZZZZ3"] == date(2098, 12, 30)
    atualizacao.voltar(dados)
    assert "ZZZZ3" not in fontes.ultimos_negocios(atual)


def test_hash_errado_nao_altera_nada(tmp_path: Path) -> None:
    dados = tmp_path / "dados"
    dados.mkdir()
    _pacote(tmp_path)
    caminho = tmp_path / "dist" / "referencias-2099.01.01.zip"
    with zipfile.ZipFile(caminho) as z:
        conteudo = {n: z.read(n) for n in z.namelist()}
    conteudo["ativos.csv"] = b"nome_pregao,ticker\nX ON,XXXX3\n"  # adulterado
    with zipfile.ZipFile(caminho, "w") as z:
        for nome, bytes_ in conteudo.items():
            z.writestr(nome, bytes_)
    obter = _github(caminho.read_bytes())
    remoto = atualizacao.verificar(obter)
    assert remoto is not None
    with pytest.raises(AtualizacaoError, match=r"SHA-256 não confere"):
        atualizacao.preparar(dados, remoto, obter)
    assert referencias.pasta_efetiva(dados) == pasta_data()


def test_formato_de_outra_versao_e_recusado(tmp_path: Path) -> None:
    dados = tmp_path / "dados"
    dados.mkdir()
    outro = referencias.FORMATO + 1
    obter = _github(_pacote(tmp_path, formato=outro))
    remoto = atualizacao.verificar(obter)
    assert remoto is not None
    with pytest.raises(AtualizacaoError, match=f"formato {outro}"):
        atualizacao.preparar(dados, remoto, obter)


def test_sem_rede_nada_muda(tmp_path: Path) -> None:
    def sem_rede(url: str) -> bytes:
        raise AtualizacaoError("sem conexão com o GitHub; nada foi alterado")

    with pytest.raises(AtualizacaoError, match="sem conexão"):
        atualizacao.verificar(sem_rede)
    assert referencias.pasta_efetiva(tmp_path) == pasta_data()


def _rezipar(caminho: Path, conteudo: dict[str, bytes], extra: list[tuple[str, bytes]]) -> bytes:
    with zipfile.ZipFile(caminho, "w") as z:
        for nome, bytes_ in [*conteudo.items(), *extra]:
            z.writestr(nome, bytes_)
    return caminho.read_bytes()


@pytest.mark.parametrize("nome", ["../fora.csv", "sub/a.csv", "a.csv:fluxo", "ativos.csv"])
def test_zip_com_nome_invalido_ou_repetido_e_recusado(tmp_path: Path, nome: str) -> None:
    import warnings

    _pacote(tmp_path)
    caminho = tmp_path / "dist" / "referencias-2099.01.01.zip"
    with zipfile.ZipFile(caminho) as z:
        conteudo = {n: z.read(n) for n in z.namelist()}
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")  # nome repetido de propósito
        obter = _github(_rezipar(caminho, conteudo, [(nome, b"x\n")]))
    remoto = atualizacao.verificar(obter)
    assert remoto is not None
    with pytest.raises(AtualizacaoError, match="nome de arquivo"):
        atualizacao.preparar(tmp_path / "dados", remoto, obter)


def test_segundo_pacote_preparado_e_conferido_de_novo(tmp_path: Path) -> None:
    """Revisão F7b A1: o cache de carregar() não pode aprovar um pacote ruim preparado depois
    de um bom na mesma pasta "novo"."""
    dados = tmp_path / "dados"
    dados.mkdir()
    bom = _github(_pacote(tmp_path / "a"))
    remoto = atualizacao.verificar(bom)
    assert remoto is not None
    atualizacao.preparar(dados, remoto, bom)
    ruim_pasta = tmp_path / "b"
    _pacote(ruim_pasta)
    pasta = ruim_pasta / "pacote-data"
    (pasta / "eventos_b3.csv").write_text(
        ",".join(COLUNAS_EVENTOS) + "\nXXXX,BRXXXX,DESDOBRAMENTO,0,,,,,,\n", encoding="utf-8"
    )
    ruim = _github(pacote.montar(pasta, ruim_pasta / "dist2").read_bytes())
    with pytest.raises(AtualizacaoError, match="linha"):
        atualizacao.preparar(dados, remoto, ruim)


def test_tabelas_embutidas_integras() -> None:
    """O data/ que vai na instalação passa na mesma conferência de um pacote baixado."""
    fontes.validar_pacote(pasta_data())
    linhas = ler_eventos_pacote(pasta_data())  # uma linha por evento, as duas fontes juntas
    # nenhuma descartada por inválida; só saem os pares grupamento + desdobramento que se anulam
    lidos = fontes.eventos(pasta_data())
    sairam = {(r["isin"], r["data_com"]) for r in linhas} - {(r.isin, r.data_com) for r in lidos}
    assert ("BRBBDCACNOR1", "2009-06-08") in sairam and len(sairam) == 12
    assert len(lidos) == len(linhas) - 2 * len(sairam)


def test_arquivo_que_some_durante_a_troca_do_pacote_nao_derruba_a_leitura(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """Aplicar ou voltar um pacote renomeia ``referencias/atual`` enquanto uma tela pode estar
    lendo: o arquivo listado que some antes de ser lido sai da assinatura, sem erro inesperado."""
    pasta = tmp_path / "atual"
    shutil.copytree(pasta_data(), pasta)
    listar = Path.iterdir

    def com_fantasma(self: Path):
        yield from listar(self)
        yield self / "apagado_no_meio.csv"

    monkeypatch.setattr(Path, "iterdir", com_fantasma)
    assert fontes.ultimos_negocios(pasta) == fontes.ultimos_negocios.__wrapped__(pasta)
