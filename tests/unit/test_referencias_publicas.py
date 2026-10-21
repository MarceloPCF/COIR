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
        fundos=json.loads((FIXTURES / "fundos.json").read_text(encoding="utf-8")),
        historico={"ITSA": json.loads((FIXTURES / "historico-ITSA.json").read_text("utf-8"))},
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


def test_cnpj_de_companhias_e_fundos(tmp_path: Path) -> None:
    """CNPJ com zeros à esquerda restaurados (a B3 os tira: TAEE 7859971000130), fundos pelo
    detalhe e CNPJ com dígito errado fora, nos conflitos."""
    arquivos, resumo = gerar.gerar(tmp_path, _coleta())
    linhas = _linhas(arquivos["emissores_b3.csv"])
    assert (
        "TAEE,07859971000130,TRANSMISSORA ALIANÇA DE ENERGIA ELÉTRICA S.A.,COMPANHIA,"
        "2026-10-19,2026-10-19"
    ) in linhas
    assert any(x.startswith("HGLG,11728688000147,PÁTRIA LOG") and ",FII," in x for x in linhas)
    assert any(x.startswith("BOVA,10406511000161,") and ",ETF," in x for x in linhas)
    assert not any(x.startswith("XXXX") for x in linhas)
    assert resumo.conflitos == ["CNPJ inválido para XXXX: 11111111111111"]
    assert len(resumo.emissores_novos) == 8 and not resumo.cnpj_mudou


def test_cnpj_novo_de_emissor_conhecido_fica_no_historico(tmp_path: Path) -> None:
    arquivos, _ = gerar.gerar(tmp_path, _coleta())
    gerar.gravar(tmp_path, arquivos)
    coleta = _coleta(date(2026, 10, 26))
    fundos = [dict(f, cnpj="63375304000153") if f["acronym"] == "HGLG" else f
              for f in coleta.fundos]  # fmt: skip
    novos, resumo = gerar.gerar(tmp_path, Coleta(**{**coleta.__dict__, "fundos": fundos}))
    hglg = [x for x in _linhas(novos["emissores_b3.csv"]) if x.startswith("HGLG")]
    assert [x.split(",")[1] for x in hglg] == ["11728688000147", "63375304000153"]
    assert hglg[0].endswith("2026-10-19,2026-10-19") and hglg[1].endswith("2026-10-26,2026-10-26")
    assert resumo.cnpj_mudou and resumo.cnpj_mudou[0].startswith("HGLG: 63375304000153")


def test_emissores_no_app_curado_por_cima(tmp_path: Path) -> None:
    from coir import referencias
    from coir.erros import ConfiguracaoError

    pasta = _copia_data(tmp_path / "data")
    (pasta / "emissores_b3.csv").write_text(
        "emissor,cnpj,razao_social,tipo,visto_desde,visto_ate\n"
        "HGLG,11728688000147,ANTIGO,FII,2026-01-01,2026-05-01\n"
        "HGLG,63375304000153,NOVO,FII,2026-05-08,2026-10-19\n"
        "TAEE,07859971000130,TAESA,COMPANHIA,2026-10-19,2026-10-19\n",
        encoding="utf-8",
    )
    (pasta / "emissores.csv").write_text(
        "emissor,cnpj,razao_social,observacao\nTAEE,10406511000161,CURADO,teste\n", encoding="utf-8"
    )
    e = referencias.emissores(pasta)
    assert e["HGLG"] == ("63375304000153", "NOVO", "FII")  # o visto por último
    assert e["TAEE"] == ("10406511000161", "CURADO", "COMPANHIA")  # a curada vale
    referencias.validar_tabelas_b3(pasta)
    (pasta / "emissores.csv").write_text(
        "emissor,cnpj,razao_social,observacao\nTAEE,11111111111111,X,\n", encoding="utf-8"
    )
    with pytest.raises(ConfiguracaoError, match="CNPJ inválido"):
        referencias.emissores(pasta)


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
    [
        (2, "2099-01-01", True),
        (1, "2099-01-01", False),
        (3, "2099-01-01", False),
        (2, "2000-01-01", False),
    ],
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


def test_proventos_em_dinheiro_com_data_de_aprovacao(tmp_path: Path) -> None:
    """O crédito do JCP (data de aprovação) vem dos proventos em dinheiro da companhia; o
    histórico nunca apaga (a B3 só lista os últimos meses)."""
    arquivos, resumo = gerar.gerar(tmp_path, _coleta())
    linhas = _linhas(arquivos["proventos_b3.csv"])
    vamo = [x for x in linhas if x.startswith("VAMO,")]
    assert vamo == ["VAMO,BRVAMOACNOR7,JRS CAP PROPRIO,0.140701568,2025-12-15,2025-12-18,"
                    "2026-04-02,Anual/2025"]  # fmt: skip
    assert len(resumo.proventos_novos) == len(linhas)
    gerar.gravar(tmp_path, arquivos)
    seguinte = _coleta(date(2026, 10, 26))
    sem_vamo = Coleta(
        **{
            **seguinte.__dict__,
            "eventos": {k: v for k, v in seguinte.eventos.items() if k != "VAMO"},
        }
    )
    novos, resumo = gerar.gerar(tmp_path, sem_vamo)
    assert _linhas(novos["proventos_b3.csv"]) == linhas and not resumo.proventos_novos


def test_historico_de_jcp_com_classe_aprovacao_e_data_com(tmp_path: Path) -> None:
    """Aba "Proventos em dinheiro" da Itaúsa (amostra real): só JCP, com classe; dividendos e
    linhas anteriores a HISTORICO_DESDE (1996) ficam de fora; nunca apaga."""
    arquivos, resumo = gerar.gerar(tmp_path, _coleta())
    linhas = _linhas(arquivos["proventos_historico_b3.csv"])
    assert len(linhas) == 12 and len(resumo.historico_novos) == 12
    assert linhas[:3] == ["ITSA,PN,JRS CAP PROPRIO,0.0591,2025-06-16,2025-06-20",
                          "ITSA,PN,JRS CAP PROPRIO,0.1859,2025-08-11,2025-08-18",
                          "ITSA,PN,JRS CAP PROPRIO,0.0235295,2025-02-10,2025-08-29"]  # fmt: skip
    assert "ITSA,ON,JRS CAP PROPRIO,0.0242425,2026-02-09,2026-08-31" in linhas
    gerar.gravar(tmp_path, arquivos)
    seguinte = _coleta(date(2026, 10, 26))
    sem_itsa = Coleta(**{**seguinte.__dict__, "historico": {}})
    novos, resumo = gerar.gerar(tmp_path, sem_itsa)
    assert _linhas(novos["proventos_historico_b3.csv"]) == linhas and not resumo.historico_novos


def test_historico_sem_aprovacao_ou_por_lote_fica_de_fora() -> None:
    base = {"typeStock": "ON", "corporateAction": "JRS CAP PROPRIO", "valueCash": "0,10",
            "lastDatePriorEx": "02/01/2020", "dateApproval": "20/12/2019",
            "quotedPerShares": "1"}  # fmt: skip
    linhas = gerar.ler_historico("XXXX", [base, {**base, "dateApproval": None},
                                          {**base, "quotedPerShares": "1000"},
                                          {**base, "lastDatePriorEx": "30/12/2015"}])  # fmt: skip
    assert linhas == [{"emissor": "XXXX", "classe": "ON", "tipo": "JRS CAP PROPRIO",
                       "valor": "0.1", "aprovado_em": "2019-12-20",
                       "data_com": "2020-01-02"}]  # fmt: skip
    with pytest.raises(GeracaoError, match="ilegível"):
        gerar.ler_historico("XXXX", [{**base, "valueCash": "x"}])
