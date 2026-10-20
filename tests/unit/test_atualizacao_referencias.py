"""F7b: "Atualizar tabelas" — o app só usa o pacote público se cada arquivo bater com o SHA-256 do
manifesto, o formato for o desta versão e as tabelas carregarem; sem rede, nada muda."""

import json
import shutil
import zipfile
from pathlib import Path

import pytest

from coir import referencias
from coir.erros import AtualizacaoError
from coir.recursos import pasta_data
from coir.servicos import atualizacao
from tools.referencias import pacote
from tools.referencias.gerar import COLUNAS_EVENTOS

URL_ZIP = "https://github.com/MarceloPCF/COIR/releases/download/referencias-2099.01.01/p.zip"


def _pacote(tmp: Path, data: str = "2099-01-01", formato: int = referencias.FORMATO) -> bytes:
    pasta = tmp / "pacote-data"
    shutil.copytree(pasta_data(), pasta)
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


def _github(zip_: bytes) -> atualizacao.Baixar:
    releases = [
        {"tag_name": "v3.0.0", "draft": False, "prerelease": False, "assets": []},
        {"tag_name": "referencias-2099.01.01", "draft": False, "prerelease": False,
         "assets": [{"name": "referencias-2099.01.01.zip", "browser_download_url": URL_ZIP}]},
        {"tag_name": "referencias-2099.02.01", "draft": True, "prerelease": False,
         "assets": [{"name": "x.zip", "browser_download_url": "https://x"}]},
    ]  # fmt: skip

    def obter(url: str) -> bytes:
        if url == atualizacao.RELEASES:
            return json.dumps(releases).encode()
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
    referencias.validar_tabelas_b3(pasta_data())
    linhas = (pasta_data() / "eventos_b3.csv").read_text(encoding="utf-8").splitlines()
    assert len(referencias.eventos_b3(pasta_data())) == len(linhas) - 1  # nenhuma descartada
