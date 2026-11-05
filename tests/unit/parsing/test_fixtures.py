"""As 10 notas anonimizadas de tests/fixtures/notas extraídas e conciliadas = o esperado.json."""

import json
from datetime import date
from decimal import Decimal
from pathlib import Path

import pytest

from coir.domain.nota import Segmento
from coir.erros import ExtracaoError
from coir.parsing.cliente import cliente
from coir.parsing.extrair import extrair_textos
from coir.parsing.paginas import classificar

PASTA = Path(__file__).parents[2] / "fixtures" / "notas"
FIXTURES = sorted(p.name.removesuffix(".txt") for p in PASTA.glob("*.txt"))


def _ler(nome: str) -> tuple[list[str], dict]:
    texto = (PASTA / f"{nome}.txt").read_text(encoding="utf-8")
    esperado = json.loads((PASTA / f"{nome}.esperado.json").read_text(encoding="utf-8"))
    return texto.split("\f"), esperado


def test_ha_dez_fixtures() -> None:
    assert len(FIXTURES) == 10


@pytest.mark.parametrize("nome", FIXTURES)
def test_fixture_extraida_conciliada_e_igual_ao_esperado(nome: str) -> None:
    paginas, esperado = _ler(nome)
    res = extrair_textos(f"{nome}.pdf", paginas)
    assert res.rejeitadas == []
    [nota] = res.notas
    assert nota.corretora == esperado["corretora"].upper()
    assert nota.segmento is Segmento(esperado["segmento"])
    assert nota.numero == esperado["numero_nota"]
    assert nota.data_pregao == date.fromisoformat(esperado["data_pregao"])
    assert len(nota.paginas) == esperado.get("folhas", 1)
    ops = esperado["operacoes"]
    lidos = nota.negocios_bovespa or nota.negocios_bmf
    assert [(n.cv, n.quantidade, n.preco, n.valor) for n in lidos] == [
        (o["cv"], Decimal(o["quantidade"]), Decimal(o["preco"]), Decimal(o["valor"])) for o in ops
    ]
    liquido = esperado["resumo_financeiro"]["liquido"]
    resumo = nota.resumo_bovespa or nota.resumo_bmf
    assert resumo is not None
    assert (resumo.liquido.valor, resumo.liquido.dc) == (Decimal(liquido["valor"]), liquido["dc"])


def test_bovespa_marcas_e_especificacao() -> None:
    paginas, _ = _ler("xp_2016_11_bovespa_antigo_multipagina")
    [nota] = extrair_textos("x.pdf", paginas).notas
    assert {n.observacao for n in nota.negocios_bovespa} >= {"D", "#"}
    assert any(n.day_trade for n in nota.negocios_bovespa)
    r = nota.resumo_bovespa
    assert r is not None and r.irrf == Decimal("2.23") and r.irrf_no_liquido


def test_btg_opcoes_irrf_nao_descontado_e_exercicio() -> None:
    paginas, _ = _ler("btg_2023_01_bovespa_opcoes")
    [nota] = extrair_textos("x.pdf", paginas).notas
    r = nota.resumo_bovespa
    assert r is not None and (r.irrf, r.irrf_no_liquido) == (Decimal("0.01"), False)
    assert nota.negocios_bovespa[0].prazo == "02/23"
    paginas, _ = _ler("btg_2023_01_bovespa_exercicio")
    [ex] = extrair_textos("x.pdf", paginas).notas
    assert ex.negocios_bovespa[0].mercado == "EXERC OPC VENDA"
    assert ex.negocios_bovespa[0].especificacao == "KLBNM199E"


def test_bmf_custos_ajustes_e_irrf_day_trade() -> None:
    paginas, _ = _ler("xp_2019_02_bmf_daytrade")
    [nota] = extrair_textos("x.pdf", paginas).notas
    r = nota.resumo_bmf
    assert r is not None
    assert (r.ajuste_day_trade.valor, r.ajuste_day_trade.dc) == (Decimal("44.00"), "C")
    assert (r.total_custos, r.impostos, r.irrf_day_trade) == (
        Decimal("2.64"),
        Decimal("0.08"),
        Decimal("0.41"),
    )
    assert nota.custos_total == Decimal("2.72")
    paginas, _ = _ler("xp_2024_01_bmf_novo_ajupos_negativo")
    [ajupos] = extrair_textos("x.pdf", paginas).notas
    assert any(n.quantidade < 0 and n.tipo == "AJUPOS" for n in ajupos.negocios_bmf)


@pytest.mark.parametrize("nome", FIXTURES)
def test_cliente_impresso_na_nota(nome: str) -> None:
    """CPF (anonimizado) e nome do cliente em todos os layouts, 2008-2026; fora da NotaExtraida."""
    paginas, _ = _ler(nome)
    res = extrair_textos(f"{nome}.pdf", paginas)
    [nota] = res.notas
    cliente = res.clientes[nota.chave]
    assert (cliente.cpf, cliente.nome) == ("00000000000", "FULANO DE TAL")
    assert "cpf" not in nota.model_dump_json()


def test_pagina_sem_cpf_nao_tem_cliente() -> None:
    assert cliente("Cliente\nFULANO DE TAL\nRua Exemplo, 0\n") is None


def test_nome_depois_do_rotulo_cliente_xp_rico_bovespa() -> None:
    """XP/Rico Bovespa: o rótulo "Cliente" fica sozinho entre a ouvidoria e o nome."""
    texto = (
        "Ouvidoria: Tel. 0800 000 0000 e-mail ouvidoria:\n"
        "Cliente\n"
        "0012345 FULANO DE TAL\n"
        "RUA EXEMPLO, 0\n"
        "Participante destino do repasse\n"
        "000.000.000-00\n"
    )
    c = cliente(texto)
    assert c is not None and c.nome == "FULANO DE TAL"


def test_nome_antes_da_ouvidoria_xp_bmf_novo() -> None:
    """XP BM&F novo: nome logo depois da linha do site; depois da ouvidoria só há rótulos."""
    texto = (
        "Internet: https://www.exemplo.com.br/ SAC: 0800-000-0000 e-mail:\n"
        "FULANO DE TAL\n"
        "Rua Exemplo, 0 - apto 1\n"
        "Ouvidoria: 0800 000 0000 e-mail ouvidoria:\n"
        "Corretora\n"
        "Cliente\n"
        "C.N.P.J.\n"
        "00.000.000/0001-00\n"
        "Numero da Corretora\n"
        "C.N.P.J/C.P.F\n"
        "000.000.000-00\n"
    )
    c = cliente(texto)
    assert c is not None and c.nome == "FULANO DE TAL"


def test_nota_clear_com_razao_social_ctvm() -> None:
    """Clear (grupo XP) usa o layout da XP; a nota de 2026 traz "CLEAR CTVM S/A" e o CNPJ da Clear
    no cabeçalho, sem "Corretora". Antes a nota era recusada (corretora não reconhecida)."""
    paginas, esperado = _ler("xp_2026_08_bovespa")
    clear = [
        p.replace(
            "XP INVESTIMENTOS CORRETORA DE CÂMBIO, TÍTULOS E VALORES MOBILIÁRIOS S.A.",
            "CLEAR CTVM S/A",
        ).replace("02.332.886/0001-04", "15.107.963/0001-66")
        for p in paginas
    ]
    assert "XP INVESTIMENTOS" not in "".join(clear)
    res = extrair_textos("clear.pdf", clear)
    assert res.rejeitadas == []
    [nota] = res.notas
    assert nota.corretora == "CLEAR"
    assert nota.numero == esperado["numero_nota"]
    assert len(nota.negocios_bovespa) == len(esperado["operacoes"])


def _sem_nome_nem_cnpj(paginas: list[str], codigo: str) -> list[str]:
    """Nota da XP com a razão social e o CNPJ trocados por desconhecidos e o código da B3 dado."""
    return [
        p.replace("XP INVESTIMENTOS CORRETORA DE CÂMBIO, TÍTULOS E VALORES MOBILIÁRIOS S.A.",
                  "CORRETORA NOVA S.A.")
        .replace("02.332.886/0001-04", "00.000.000/0001-91")
        .replace(" 3-5 000000 ", f" {codigo} 000000 ")
        for p in paginas
    ]  # fmt: skip


def test_corretora_pelo_codigo_da_b3_sem_nome_nem_cnpj() -> None:
    """Segunda tentativa: sem nome nem CNPJ conhecidos, o código da corretora na B3 impresso antes
    da conta ("3-5" é a XP) identifica a corretora; a nota é lida normalmente."""
    paginas, esperado = _ler("xp_2026_08_bovespa")
    res = extrair_textos("nova.pdf", _sem_nome_nem_cnpj(paginas, "3-5"))
    assert res.rejeitadas == []
    [nota] = res.notas
    assert nota.corretora == "XP" and nota.numero == esperado["numero_nota"]


def test_corretora_desconhecida_diz_o_que_foi_tentado() -> None:
    paginas, _ = _ler("xp_2026_08_bovespa")
    res = extrair_textos("nova.pdf", _sem_nome_nem_cnpj(paginas, "999-9"))
    [rejeitada] = res.rejeitadas
    assert "código 999 na B3" in rejeitada.motivo and "CNPJ" in rejeitada.motivo


@pytest.mark.parametrize(
    ("linha", "impresso"),
    [
        ("Número da Corretora 88-4 Cliente", "(código 88 na B3)"),
        ("Numero da Corretora 88-4", "(código 88 na B3)"),
        ("NÚMERO DA CORRETORA\t88-4 Cliente", "(código 88 na B3)"),
        ("Númeroda Corretora 88-4 Cliente", "(código 88 na B3)"),
    ],
)
def test_corretora_desconhecida_com_o_codigo_na_linha_do_rotulo(linha: str, impresso: str) -> None:
    """A CM Capital imprime o código da B3 no próprio rótulo, sem a conta na mesma linha; a
    mensagem mostra o código lido, para o usuário saber qual corretora falta."""
    with pytest.raises(ExtracaoError, match="corretora não identificada") as erro:
        classificar(f"NOTA DE CORRETAGEM\nCorretora Nova S.A.\n{linha}\n")
    assert impresso in str(erro.value)


def test_codigo_no_rotulo_nao_identifica_a_corretora() -> None:
    """O código no rótulo só informa: "3-5" é o da XP, mas sem a conta ao lado o layout não é o
    que o perfil da XP sabe ler, e a nota continua recusada dizendo o código."""
    with pytest.raises(ExtracaoError, match=r"código 3 na B3"):
        classificar("NOTA DE CORRETAGEM\nCorretora Nova S.A.\nNúmero da Corretora 3-5 Cliente\n")


@pytest.mark.parametrize(
    "linha",
    ["Número da Corretora 123456-7 Cliente", "Número da Corretora", "Número da Corretora 88"],
)
def test_rotulo_sem_codigo_valido_nao_mostra_codigo(linha: str) -> None:
    """A conta do cliente (mais de 4 dígitos) nunca aparece na mensagem como se fosse o código."""
    with pytest.raises(ExtracaoError, match="corretora não identificada") as erro:
        classificar(f"NOTA DE CORRETAGEM\nCorretora Nova S.A.\n{linha}\n")
    assert "código" not in str(erro.value).split(":")[0]
