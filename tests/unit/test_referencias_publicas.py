"""F7b: gerador das tabelas de referência públicas a partir de amostras reais (recortadas) do
Cadastro de Instrumentos e das empresas listadas da B3 e da série 4390 do BCB (dados públicos)."""

import json
from datetime import date
from pathlib import Path

import pytest

from tools.referencias import gerar
from tools.referencias.gerar import Coleta, GeracaoError

FIXTURES = Path(__file__).parents[1] / "fixtures" / "referencias"
DIA = date(2026, 10, 19)


def _coleta(dia: date = DIA, instrumentos: bytes | None = None) -> Coleta:
    eventos = {c: json.loads((FIXTURES / f"empresa-{c}.json").read_text(encoding="utf-8"))
               for c in ("DIRR", "MGLU", "ITSA", "VAMO")}  # fmt: skip
    return Coleta(
        data=dia,
        instrumentos=instrumentos or (FIXTURES / "instrumentos-2026-10-16.csv").read_bytes(),
        empresas=json.loads((FIXTURES / "empresas.json").read_text(encoding="utf-8"))["results"],
        eventos=eventos,
        selic=json.loads((FIXTURES / "selic.json").read_text(encoding="utf-8")),
    )


def _linhas(texto: str) -> list[str]:
    return texto.strip().splitlines()[1:]


def test_instrumentos_tipos_e_nomes_de_pregao(tmp_path: Path) -> None:
    arquivos, resumo = gerar.gerar(tmp_path, _coleta())
    tickers = [linha.split(",")[0] for linha in _linhas(arquivos["instrumentos_b3.csv"])]
    # só o mercado à vista (sem lote fracionário, índice nem opção), com ISIN
    assert tickers == ["AAPL34", "AATH11", "AMOB3", "BOVA11", "DIRR3", "HGLG11", "ITSA3",
                       "ITSA4", "MGLU3", "TAEE11", "VAMO3"]  # fmt: skip
    assert "DIRECIONAL ON,DIRR3" in _linhas(arquivos["ativos_b3.csv"])
    assert "MAGAZ LUIZA ON,MGLU3" in _linhas(arquivos["ativos_b3.csv"])  # como a nota imprime
    assert "TAESA UNT,TAEE11" in _linhas(arquivos["ativos_b3.csv"])
    tipos = {
        linha.split(",")[0]: linha.split(",")[1]
        for linha in _linhas(arquivos["tipos_ativo_b3.csv"])
    }
    # FIP de infraestrutura (AATH11) fica sem tipo: sem palpite
    assert tipos == {"AAPL34": "BDR", "BOVA11": "ETF", "HGLG11": "FII", "TAEE11": "UNIT"}
    assert _linhas(arquivos["opcoes_raiz_b3.csv"]) == ["ITSA,ITSA4"]
    assert len(resumo.ativos_novos) == 11 and not resumo.trocas_candidatas


def test_eventos_da_b3_com_data_com_e_fator(tmp_path: Path) -> None:
    arquivos, _ = gerar.gerar(tmp_path, _coleta())
    eventos = _linhas(arquivos["eventos_b3.csv"])
    assert "DIRR,BRDIRRACNOR0,DESDOBRAMENTO,200,2025-08-08,2025-08-08,BRDIRRACNOR0,,," in eventos
    assert "MGLU,BRMGLUACNOR2,GRUPAMENTO,0.1,2024-05-24,2024-04-24,BRMGLUACNOR2,,," in eventos
    cisao = [e for e in eventos if "CIS RED CAP" in e and "BRAMOBACNOR9" in e]
    assert cisao == [
        "VAMO,BRVAMOACNOR7,CIS RED CAP,28.290555,2024-12-13,2024-11-22,BRAMOBACNOR9,,,"
    ]
    subscricao = [e for e in eventos if e.startswith("VAMO") and "SUBSCRICAO" in e]
    assert subscricao and ",3.85,11/03/2026 a 27/04/2026,2026-04-30" in subscricao[0]


def test_selic_so_meses_encerrados(tmp_path: Path) -> None:
    (tmp_path / "selic_mensal.csv").write_text(
        "mes,taxa\n2026-06,1.12\n2026-07,1.22\n", encoding="utf-8"
    )
    arquivos, resumo = gerar.gerar(tmp_path, _coleta(date(2026, 10, 5)))
    assert _linhas(arquivos["selic_mensal.csv"]) == ["2026-06,1.12", "2026-07,1.22",
                                                    "2026-08,1.09", "2026-09,1.08"]  # fmt: skip
    assert resumo.selic_nova == ["2026-08: 1.09%", "2026-09: 1.08%"]  # outubro ainda parcial


def test_reprocessar_e_deterministico_e_nao_apaga_historico(tmp_path: Path) -> None:
    arquivos, _ = gerar.gerar(tmp_path, _coleta())
    gerar.gravar(tmp_path, arquivos)
    # coleta seguinte sem VAMO3 e com o mesmo ISIN em outro ticker: troca candidata
    bruto = (FIXTURES / "instrumentos-2026-10-16.csv").read_bytes().decode("latin-1")
    seguinte = bruto.replace(";VAMO3;", ";VAMN3;").encode("latin-1")
    novos, resumo = gerar.gerar(tmp_path, _coleta(date(2026, 10, 26), seguinte))
    linhas = _linhas(novos["instrumentos_b3.csv"])
    assert any(x.startswith("VAMO3,BRVAMOACNOR7") and x.endswith("2026-10-19,2026-10-19")
               for x in linhas)  # fmt: skip
    assert resumo.trocas_candidatas == [("VAMO3", "VAMN3", "2026-10-19")]
    assert gerar.gerar(tmp_path, _coleta())[0] == arquivos  # mesma entrada, mesmos bytes
    assert gerar.gravar(tmp_path, arquivos) == []  # nada mudou: nada é regravado


def test_resposta_fora_do_formato_falha_alto(tmp_path: Path) -> None:
    with pytest.raises(GeracaoError, match="cabeçalho"):
        gerar.gerar(tmp_path, _coleta(instrumentos=b"TckrSymb;ISIN\nPETR4;BR\n"))


def _copia_data(destino: Path) -> Path:
    import shutil

    from coir.recursos import pasta_data

    shutil.copytree(pasta_data(), destino)
    return destino


def test_tabela_gerada_entra_por_baixo_da_manual(tmp_path: Path) -> None:
    from coir import referencias

    pasta = _copia_data(tmp_path / "data")
    (pasta / "ativos_b3.csv").write_text(
        "nome_pregao,ticker\nNOVA EMPRESA ON,NOVA3\nDIRECIONAL ON,XXXX3\n", encoding="utf-8"
    )
    (pasta / "tipos_ativo_b3.csv").write_text(
        "ticker,tipo,observacao\nHGLG11,FII,B3\nBIDI11,ETF,B3\n", encoding="utf-8"
    )
    refs = referencias.carregar(pasta)
    assert refs.ativos["novaempresaon"] == "NOVA3"
    assert refs.ativos["direcionalon"] == "DIRR3"  # a manual vale
    tipos = referencias.carregar_fiscais(pasta).tipos
    assert tipos.tipo("HGLG11").value == "FII" and tipos.tipo("BIDI11").value == "UNIT"


@pytest.mark.parametrize(
    ("formato", "data", "baixada"),
    [(1, "2099-01-01", True), (2, "2099-01-01", False), (1, "2000-01-01", False)],
)
def test_pacote_baixado_so_vale_se_mais_novo_e_do_mesmo_formato(
    tmp_path: Path, formato: int, data: str, baixada: bool
) -> None:
    from coir import referencias
    from coir.recursos import pasta_data

    atual = tmp_path / "referencias" / "atual"
    atual.mkdir(parents=True)
    (atual / "pacote.json").write_text(json.dumps({"formato": formato, "data": data}))
    esperado = atual if baixada else pasta_data()
    assert referencias.pasta_efetiva(tmp_path) == esperado
    assert referencias.pasta_efetiva(None) == pasta_data()
    assert referencias.FORMATO == gerar.FORMATO


def test_pacote_deterministico_com_hash_de_cada_arquivo(tmp_path: Path) -> None:
    import hashlib
    import zipfile

    from coir.recursos import pasta_data
    from tools.referencias import pacote

    primeiro = pacote.montar(pasta_data(), tmp_path / "a")
    segundo = pacote.montar(pasta_data(), tmp_path / "b")
    assert primeiro.read_bytes() == segundo.read_bytes()  # mesmo data/, mesmo zip
    assert primeiro.name == pacote.nome_do_release(pasta_data()) + ".zip"
    with zipfile.ZipFile(primeiro) as z:
        manifesto = json.loads(z.read("manifesto.json"))
        for nome, sha in manifesto["arquivos"].items():
            assert hashlib.sha256(z.read(nome)).hexdigest() == sha
    assert {"ativos.csv", "eventos_b3.csv", "pacote.json"} <= set(manifesto["arquivos"])


def test_selic_e_eventos_fora_do_formato_falham_alto(tmp_path: Path) -> None:
    from dataclasses import replace

    with pytest.raises(GeracaoError, match="fora do formato"):
        gerar.gerar(tmp_path, replace(_coleta(), selic=[{"data": "outubro", "valor": "x"}]))
    with pytest.raises(GeracaoError, match="fora do formato"):
        gerar.gerar(tmp_path, replace(_coleta(), eventos={"XXXX": [{"stockDividends": [{}]}]}))
