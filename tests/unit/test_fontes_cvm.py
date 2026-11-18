"""Cadastros da CVM (Fase 3): códigos de negociação das companhias pelo FCA e cotas de FII e
Fiagro pelos informes mensais, só para completar o que a bolsa não traz.

Os zips e o cadastro são fictícios, no layout dos dados abertos da CVM.
"""

import io
import zipfile
from pathlib import Path

from coir import fontes, referencias
from coir.domain.fiscal import TipoAtivo
from coir.fontes import composicao
from coir.fontes.provedores import cvm as provedor
from coir.recursos import pasta_data
from tools.referencias import cvm, gerar

CABECALHO_FCA = ("CNPJ_Companhia;Data_Referencia;Versao;ID_Documento;Nome_Empresarial;"
                 "Valor_Mobiliario;Sigla_Classe_Acao_Preferencial;Classe_Acao_Preferencial;"
                 "Codigo_Negociacao;Composicao_BDR_Unit;Mercado;Sigla_Entidade_Administradora;"
                 "Entidade_Administradora;Data_Inicio_Negociacao;Data_Fim_Negociacao;Segmento;"
                 "Data_Inicio_Listagem;Data_Fim_Listagem")  # fmt: skip
CNPJ_A = "11.222.333/0001-81"  # dígitos válidos, empresa fictícia
CNPJ_B = "11.444.777/0001-61"
CNPJ_C = "04.252.011/0001-10"


def _zip(nome: str, texto: str) -> bytes:
    bruto = io.BytesIO()
    with zipfile.ZipFile(bruto, "w") as z:
        z.writestr(nome, texto.encode("latin-1"))
    return bruto.getvalue()


def _fca(ano: int, *linhas: str) -> bytes:
    return _zip(f"fca_cia_aberta_valor_mobiliario_{ano}.csv",
                "\n".join((CABECALHO_FCA, *linhas)) + "\n")  # fmt: skip


def _linha_fca(cnpj: str, data: str, nome: str, especie: str, codigo: str,
               mercado: str = "Bolsa") -> str:  # fmt: skip
    return f"{cnpj};{data};1;1;{nome};{especie};;;{codigo};;{mercado};B3;B3;;;;;"


def test_fca_so_acao_e_unit_em_bolsa_e_a_declaracao_mais_recente() -> None:
    on, pn, nome = "Ações Ordinárias", "Ações Preferenciais", "FICTICIA SA"
    antigo = _fca(2019, _linha_fca(CNPJ_A, "2019-01-01", nome, on, "FICT3"),
                  _linha_fca(CNPJ_A, "2019-01-01", nome, "Debêntures", "FICT11"),
                  _linha_fca(CNPJ_A, "2019-01-01", nome, pn, "FICT4", mercado="Balcão Organizado"),
                  _linha_fca(CNPJ_B, "2019-01-01", "NUMERICA", on, "22853"))  # fmt: skip
    novo = _fca(2024, _linha_fca(CNPJ_A, "2024-01-01", "FICTICIA S.A.", on, "FICT3"),
                _linha_fca(CNPJ_A, "2024-01-01", "FICTICIA S.A.", "Units", "FICT11"))  # fmt: skip
    razoes = {"11222333000181": "FICTÍCIA PARTICIPAÇÕES S.A."}
    linhas = cvm.ler_valores([antigo, novo], razoes)
    assert [(r["ticker"], r["especie"], r["visto_ate"]) for r in linhas] == [
        ("FICT11", "UNIT", "2024-01-01"), ("FICT3", "ON", "2024-01-01")]  # fmt: skip
    # a razão social do cadastro de companhias vale sobre o nome do FCA
    assert {r["razao_social"] for r in linhas} == {"FICTÍCIA PARTICIPAÇÕES S.A."}
    assert {r["cnpj"] for r in linhas} == {"11222333000181"}


def test_cadastro_de_companhias_por_cnpj_so_digitos() -> None:
    texto = "CNPJ_CIA;DENOM_SOCIAL;DENOM_COMERC\n11.222.333/0001-81;FICTÍCIA S.A.;FICT\n"
    assert cvm.razoes_sociais(texto.encode("latin-1")) == {"11222333000181": "FICTÍCIA S.A."}


def _fii(ano: int, *linhas: tuple[str, str, str, str, str], antigo: bool = False) -> bytes:
    colunas = ("CNPJ_Fundo;Data_Referencia;Nome_Fundo;Codigo_ISIN;Mercado_Negociacao_Bolsa"
               if antigo else
               "CNPJ_Fundo_Classe;Data_Referencia;Nome_Fundo_Classe;Codigo_ISIN;"
               "Mercado_Negociacao_Bolsa")  # fmt: skip
    corpo = "\n".join(";".join(r) for r in linhas)
    return _zip(f"inf_mensal_fii_geral_{ano}.csv", f"{colunas}\n{corpo}\n")


def test_fundos_pelo_informe_nos_dois_layouts_e_isin_ambiguo_fora() -> None:
    antigo = _fii(2017, (CNPJ_B, "2017-05-01", "FICT RENDA FII", "BRFRNDCTF000", "S"),
                  (CNPJ_C, "2017-05-01", "FICT BALCAO FII", "BRFBLCCTF001", "N"),
                  antigo=True)  # fmt: skip
    novo = _fii(2026, (CNPJ_B, "2026-09-01", "FICT RENDA FII RESP LTDA", "BRFRNDCTF000", "S"),
                (CNPJ_A, "2026-09-01", "FICT LOGISTICA FII", "BRFLOGCTF002", "S"),
                (CNPJ_C, "2026-09-01", "OUTRO FUNDO FII", "BRFLOGCTF002", "S"),  # repete o ISIN
                (CNPJ_C, "2026-09-01", "FICT RECIBO", "BRFRNDR01003", "S"))  # fmt: skip
    fiagro = _zip("inf_mensal_fiagro_202609.csv",
                  "CNPJ_Classe;Data_Referencia;Nome_Classe;Codigo_ISIN;Mercado_Negociacao\n"
                  f"{CNPJ_A};2026-09-01;FICT AGRO FIAGRO;BRFAGRCTF003;BOLSA\n"
                  f"{CNPJ_C};2026-09-01;FICT AGRO BALCAO;BRFAGBCTF004;BALCAO\n")  # fmt: skip
    linhas, ambiguos = cvm.ler_fundos([antigo, novo], [fiagro])
    assert ambiguos == {"BRFLOGCTF002"}
    assert [(r["isin"], r["tipo"], r["nome"]) for r in linhas] == [
        ("BRFAGRCTF003", "FIAGRO", "FICT AGRO FIAGRO"),
        ("BRFRNDCTF000", "FII", "FICT RENDA FII RESP LTDA")]  # fmt: skip


def test_juntar_nunca_apaga_e_tira_o_ambiguo() -> None:
    anteriores = [{"isin": "A", "visto_ate": "2025-01-01", "nome": "velho"},
                  {"isin": "B", "visto_ate": "2025-01-01", "nome": "fica"},
                  {"isin": "C", "visto_ate": "2025-01-01", "nome": "ambíguo"}]  # fmt: skip
    novas = [{"isin": "A", "visto_ate": "2026-01-01", "nome": "novo"}]
    juntas = cvm.juntar(anteriores, novas, "isin", fora={"C"})
    assert [(r["isin"], r["nome"]) for r in juntas] == [("A", "novo"), ("B", "fica")]


def _pasta(tmp_path: Path, valores: str, fundos: str) -> Path:
    (tmp_path / provedor.VALORES).write_text(
        "ticker,cnpj,razao_social,especie,visto_ate\n" + valores, encoding="utf-8")  # fmt: skip
    (tmp_path / provedor.FUNDOS).write_text(
        "isin,cnpj,nome,tipo,visto_ate\n" + fundos, encoding="utf-8")  # fmt: skip
    return tmp_path


def test_codigo_ambiguo_na_cvm_fica_de_fora(tmp_path: Path) -> None:
    pasta = _pasta(
        tmp_path,
        "FICT3,11222333000181,FICTICIA SA,ON,2024-01-01\n"
        "FICT11,11222333000181,FICTICIA SA,UNIT,2024-01-01\n"
        "DUPL3,11222333000181,DUPLA SA,ON,2019-01-01\n"
        "MIST3,11444777000161,MISTA SA,ON,2019-01-01\n",
        "BRDUPLCTF000,04252011000110,DUPLA FII,FII,2026-09-01\n"  # código de companhia e de fundo
        "BRMISTCTF001,04252011000110,MISTA FII,FII,2022-01-01\n"
        "BRFRNDCTF000,11444777000161,FICT RENDA FII,FII,2025-01-01\n"
        "BRFRNDCTF018,11444777000161,FICT RENDA FII,FII,2026-09-01\n"  # classe nova, mesmo CNPJ
        "BR0PRQCTF003,04252011000110,SEM CODIGO,FII,2026-09-01\n",
    )  # fmt: skip
    emissores = {e.emissor: e for e in provedor.emissores(pasta)}
    assert set(emissores) == {"FICT", "FRND"}
    assert (emissores["FRND"].tipo, emissores["FRND"].visto_ate) == ("FII", "2026-09-01")
    assert provedor.tipos_ativo(pasta) == [("FICT11", "UNIT"), ("FRND11", "FII")]
    assert provedor.tickers(pasta) >= {"FICT3", "FICT11", "FRND11", "DUPL3"}


def test_bolsa_vale_sobre_a_cvm_e_a_cvm_completa(tmp_path: Path) -> None:
    """O código reaproveitado: a bolsa diz quem o tem hoje; a CVM só entra onde a bolsa não diz."""
    pasta = _pasta(tmp_path, "OLDC3,11222333000181,ANTIGA SA,ON,2020-01-01\n"
                             "NOVA3,11444777000161,SO NA CVM SA,ON,2024-01-01\n", "")  # fmt: skip
    (pasta / "emissores_b3.csv").write_text(
        "emissor,cnpj,razao_social,tipo,visto_desde,visto_ate\n"
        "OLDC,04252011000110,DONA DO CODIGO HOJE,ETF,2026-01-01,2026-11-10\n",
        encoding="utf-8",
    )
    (pasta / "emissores.csv").write_text(
        "emissor,cnpj,razao_social,observacao\nNOVA,04252011000110,CURADA SA,conferido\n",
        encoding="utf-8")  # fmt: skip
    emissores = fontes.emissores(pasta)
    assert emissores["OLDC"] == ("04252011000110", "DONA DO CODIGO HOJE", "ETF")
    assert emissores["NOVA"] == ("04252011000110", "CURADA SA", "COMPANHIA")
    referencias.definir_fonte_b3("desligada")
    try:
        assert fontes.emissores(pasta)["OLDC"] == ("11222333000181", "ANTIGA SA", "COMPANHIA")
    finally:
        referencias.definir_fonte_b3("pacote")


def test_tipo_pela_cvm_fica_por_baixo_da_bolsa_e_da_curadoria() -> None:
    pasta = pasta_data()
    fiscais = referencias.carregar_fiscais(pasta)
    curados = {r["ticker"] for r in gerar.ler_csv(pasta / "tipos_ativo.csv")}
    bolsa = {r["ticker"]: r["tipo"] for r in gerar.ler_csv(pasta / "tipos_ativo_b3.csv")}
    so_cvm = [(t, tipo) for t, tipo in provedor.tipos_ativo(pasta)
              if t not in bolsa and t not in curados]  # fmt: skip
    assert so_cvm and all(fiscais.tipos.tipo(t) is TipoAtivo(tipo) for t, tipo in so_cvm)
    # onde a bolsa diz outra coisa (código de FII antigo que hoje é de ETF), vale a bolsa
    for t, _ in provedor.tipos_ativo(pasta):
        if t in bolsa and t not in curados:
            assert fiscais.tipos.tipo(t) is TipoAtivo(bolsa[t])


def test_regra_do_isin_confere_com_todo_o_cadastro_da_bolsa() -> None:
    """Classe no ISIN -> final do código: nenhuma diferença no cadastro de instrumentos."""
    pasta = pasta_data()
    pares = [(i.isin, i.ticker) for i in composicao.b3.instrumentos(pasta)
             if composicao.RE_TICKER_PADRAO.match(i.ticker)]  # fmt: skip
    com_regra = [(t, composicao.ticker_pelo_isin(isin)) for isin, t in pares
                 if composicao.ticker_pelo_isin(isin)]  # fmt: skip
    assert len(com_regra) > 1000 and all(t == r for t, r in com_regra)
    assert composicao.ticker_pelo_isin("BRFICTACNPB0") == "FICT6"
    assert composicao.ticker_pelo_isin("BRFICTCDAM10") == "FICT11"
    assert composicao.ticker_pelo_isin("BRFICTDBS001") is None  # debênture: sem código de ação
    assert composicao.ticker_pelo_isin("USFICTACNOR0") is None


def test_isin_do_evento_pela_regra_so_com_ticker_confirmado(tmp_path: Path) -> None:
    _pasta(tmp_path, "FICT3,11222333000181,FICTICIA SA,ON,2024-01-01\n", "")
    (tmp_path / "eventos_cvm.csv").write_text(
        gerar.texto_csv(gerar.COLUNAS_EVENTOS_CVM, [
            {"emissor": "FICT", "isin": "BRFICTACNOR0", "tipo": "DESDOBRAMENTO", "fator": "100",
             "data_com": "2024-05-02", "isin_emitido": "BRFICTACNOR0", "origem": "cvm_ipe",
             "referencia": "IPE 1", "conferido_cotahist": "nao"},
            {"emissor": "OUTR", "isin": "BROUTRACNOR0", "tipo": "DESDOBRAMENTO", "fator": "100",
             "data_com": "2024-05-02", "isin_emitido": "BROUTRACNOR0", "origem": "cvm_ipe",
             "referencia": "IPE 2", "conferido_cotahist": "nao"},
        ]), encoding="utf-8")  # fmt: skip
    assert fontes.tickers_por_isin(tmp_path) == {"BRFICTACNOR0": "FICT3"}


def test_tabelas_da_cvm_no_pacote() -> None:
    pasta = pasta_data()
    valores = gerar.ler_csv(pasta / cvm.VALORES)
    fundos = gerar.ler_csv(pasta / cvm.FUNDOS)
    assert valores and tuple(valores[0]) == cvm.COLUNAS_VALORES
    assert fundos and tuple(fundos[0]) == cvm.COLUNAS_FUNDOS
    assert all(cvm.RE_CODIGO.fullmatch(r["ticker"]) for r in valores)
    assert all(cvm.RE_ISIN_COTA.fullmatch(r["isin"]) for r in fundos)
    assert len({r["isin"] for r in fundos}) == len(fundos)


def test_mapeamento_da_bolsa_vence_a_regra(tmp_path: Path) -> None:
    """O ISIN que o cadastro da bolsa já liga a um ticker não vai pela estrutura do código."""
    test_isin_do_evento_pela_regra_so_com_ticker_confirmado(tmp_path)
    (tmp_path / "instrumentos_b3.csv").write_text(
        "ticker,isin,categoria,especificacao,emissor,nome,visto_desde,visto_ate\n"
        "NOVO3,BRFICTACNOR0,SHARES,ON,NOVO,NOVA,2026-01-01,2026-11-10\n",
        encoding="utf-8",
    )
    assert fontes.tickers_por_isin(tmp_path) == {"BRFICTACNOR0": "NOVO3"}


def test_fca_sem_a_tabela_de_valores_mobiliarios() -> None:
    sem = _zip("fca_cia_aberta_geral_2018.csv", "CNPJ_Companhia;Data_Referencia\n")
    assert cvm.ler_valores([sem, _fca(2019)], {}) == []


def test_meses_de_fiagro_na_virada_do_ano_e_cache_so_de_ano_fechado(tmp_path: Path) -> None:
    from datetime import date

    from tools.referencias import fontes as coleta

    pedidos: list[str] = []

    def obter(url: str) -> bytes:
        pedidos.append(url)
        if "fiagro" in url:
            if url.endswith("202701.zip"):
                raise gerar.GeracaoError("ainda não saiu")
            return _zip("inf_mensal_fiagro_x.csv", "CNPJ_Classe;Data_Referencia;Nome_Classe;"
                        "Codigo_ISIN;Mercado_Negociacao\n")  # fmt: skip
        if url.endswith(".csv"):
            return b"CNPJ_CIA;DENOM_SOCIAL\n"
        return _fca(2018) if "fca" in url else _fii(2016)

    _, _, _, faltas = coleta.cadastros_cvm(date(2027, 1, 15), tmp_path, obter)
    meses = [u.rsplit("_", 1)[1][:6] for u in pedidos if "fiagro" in u]
    assert meses[:3] == ["202701", "202612", "202611"] and len(meses) == coleta.MESES_FIAGRO
    assert faltas == ["informe de Fiagro de 01/2027"]
    # o ano corrente muda: só os anos fechados ficam no cache
    assert (tmp_path / "fca_cia_aberta_2026.zip").is_file()
    assert not (tmp_path / "fca_cia_aberta_2027.zip").exists()


def test_cnpj_do_ticker_antigo_pela_companhia_que_o_declarou(tmp_path: Path) -> None:
    """O código foi reaproveitado (ou a bolsa traz a filial): o Bens e Direitos do ticker antigo
    sai com o CNPJ de quem o declarou no FCA; o ticker que a bolsa viu depois do último FCA e o
    código da curadoria ficam como estão."""
    pasta = _pasta(tmp_path,
                   "VELH3,11222333000181,VELHA SA,ON,2024-01-01\n"
                   "REUS3,11222333000181,VELHA SA,ON,2020-01-01\n"
                   "CURA3,11222333000181,VELHA SA,ON,2024-01-01\n"
                   "IGUA3,04252011000110,IGUAL SA,ON,2026-01-01\n", "")  # fmt: skip
    (pasta / "emissores_b3.csv").write_text(
        "emissor,cnpj,razao_social,tipo,visto_desde,visto_ate\n"
        "VELH,11444777000161,NOVA DONA,FII,2026-01-01,2026-11-10\n"
        "REUS,11444777000161,NOVA DONA,COMPANHIA,2026-01-01,2026-11-10\n"
        "CURA,11444777000161,NOVA DONA,COMPANHIA,2026-01-01,2026-11-10\n"
        "IGUA,04252011000110,IGUAL S.A.,COMPANHIA,2026-01-01,2026-11-10\n",
        encoding="utf-8",
    )
    (pasta / "instrumentos_b3.csv").write_text(
        "ticker,isin,categoria,especificacao,emissor,nome,visto_desde,visto_ate\n"
        "REUS3,BRREUSACNOR0,SHARES,ON,REUS,NOVA,2026-01-01,2026-11-10\n",
        encoding="utf-8",
    )
    (pasta / "emissores.csv").write_text(
        "emissor,cnpj,razao_social,observacao\nCURA,04252011000110,CURADA SA,conferido\n",
        encoding="utf-8")  # fmt: skip
    assert fontes.correcoes_por_ticker(pasta) == {
        "VELH3": ("11222333000181", "VELHA SA", "COMPANHIA")}  # fmt: skip
