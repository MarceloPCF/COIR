"""Anonimizador do texto das notas (coir.parsing.anonimizar), só com dados fictícios.

Dois grupos de casos:
- as fixtures (notas completas, que conciliam) com os marcadores genéricos (FULANO DE TAL,
  000.000.000-00...) trocados por dados pessoais fictícios: a cópia extrai igual, fora a conta;
- cabeçalhos no layout real de cada corretora (BTG/Necton, XP antiga e nova, Rico, Ágora;
  Bovespa e BM&F), que o leitor recusa: a cópia é recusada igual e nenhum dado pessoal sobra.
"""

import re
from pathlib import Path

import pytest

from coir.domain.cpf import cpf_valido
from coir.parsing.anonimizar import (
    MapaAnonimizacao,
    anonimizar,
    conferir,
    divergencias_de_forma,
    vazamentos,
)
from coir.parsing.extrair import extrair_textos
from coir.parsing.numeros import compactar
from tests.apoio import cpf_de_teste

PASTA = Path(__file__).parents[2] / "fixtures" / "notas"
SEGREDO = "00" * 32
CPF = cpf_de_teste("314159265")
CPF_TRACO = f"{CPF[:3]}.{CPF[3:6]}.{CPF[6:9]}-{CPF[9:]}"
CPF_PONTO = f"{CPF[:3]}.{CPF[3:6]}.{CPF[6:9]}.{CPF[9:]}"
CONTA = "731942"  # fictícia, como a corretora imprime sem zeros
ASSESSOR = "51873"
NOME = "JOANA PEREIRA DOS SANTOS"
PESSOAIS = ("JOANA", "PEREIRA", "SANTOS", "ACACIAS", "GOIANIA", "BUENO", "98765", "74210",
            "765432", ASSESSOR, CONTA, CPF)  # fmt: skip


def _mapa() -> MapaAnonimizacao:
    return MapaAnonimizacao(segredo=SEGREDO)


def _sobrou(textos: list[str]) -> list[str]:
    """Dados pessoais fictícios que ainda aparecem (palavras sem espaço; números com ou sem
    separadores)."""
    texto = "\n".join(textos)
    achados = []
    for termo in PESSOAIS:
        if termo.isdigit():
            if re.search(r"(?<!\d)0*" + r"[.\- ]?".join(termo) + r"(?!\d)", texto):
                achados.append(termo)
        elif compactar(termo) in compactar(texto):
            achados.append(termo)
    return achados


# ----------------------------------------------------------------------------- mapa
def test_mapa_e_deterministico_e_cpf_ficticio_e_valido() -> None:
    a, b = _mapa(), _mapa()
    assert a.cpf(CPF) == b.cpf(CPF) and a.cpf(CPF) != CPF and cpf_valido(a.cpf(CPF))
    assert a.conta("XP", "0731942") == b.conta("XP", CONTA) != CONTA
    assert len(a.conta("XP", CONTA)) == len(CONTA) and not a.conta("XP", CONTA).startswith("0")
    assert a.conta("BTG", CONTA) != a.conta("XP", CONTA)  # outra corretora, outra conta
    assert CPF not in repr(a) and CONTA not in repr(a) and SEGREDO not in repr(a)


def test_mapa_mantem_a_conta_ja_atribuida() -> None:
    mapa = MapaAnonimizacao(segredo=SEGREDO, contas={f"XP|{CONTA}": "1234"})
    assert mapa.conta("XP", CONTA) == "1234"  # mais curta: sai com zeros à esquerda no texto


def test_trocar_mantem_formato() -> None:
    mapa = _mapa()
    novo = mapa.trocar("bloco", "Acácias-742")
    assert len(novo) == 11 and novo[0].isupper() and novo[1:7].isalpha() and novo[7] == "-"
    assert novo[8:].isdigit() and compactar(novo) != compactar("Acácias-742")


# ----------------------------------------------------------------------------- fixtures
def _pessoal(texto: str) -> str:
    """A fixture com dados pessoais fictícios no lugar dos marcadores genéricos."""
    trocas = [
        (r"FULANO DE TAL", NOME),
        (r"(?i)rua exemplo, 0 - apto 0", "Rua das Acacias, 742 - Apto 51"),
        (r"(?i)\bbairro\b", "SETOR BUENO"),
        (r"(?i)cidade - uf", "GOIANIA - GO"),
        (r"00000-000", "74210-030"),
        (r"\b00000000\b", "74210030"),
        (r"\(00\) 00000-(\s*)0000", r"(62) 98765-\g<1>4321"),
        (r"\(00\) 000000000", "(62) 987654321"),
        (r"000\.000\.000-00", CPF_TRACO),
        (r"000\.000\.000\.00", CPF_PONTO),
        (r"(?m)^(\s*\d{1,4}-\d) 0+(\s*)$", rf"\g<1> 0{CONTA}\g<2>"),
        (r"(?m)^0000000-0 ", f"0{CONTA}-4 "),
        (r"(?m)^000000 (?=" + NOME + ")", f"{CONTA} "),
        (r"(?m)^000000$", CONTA),  # sob "Codigo do Cliente" ou na última linha (XP BM&F)
        (r"(?m)^(CIDADE - UF|GOIANIA - GO) 000000$", rf"\g<1> {CONTA}"),
        (r"(?m)(Assessor\r?\n)0000$", rf"\g<1>{ASSESSOR}"),
        (r"(?m)^000 0000 000000$", "237 4321 765432"),
    ]
    for padrao, novo in trocas:
        texto = re.sub(padrao, novo, texto)
    return texto


FIXTURES = sorted(p.stem for p in PASTA.glob("*.txt"))


@pytest.mark.parametrize("nome", FIXTURES)
def test_fixture_copia_extrai_igual_e_sem_dado_pessoal(nome: str) -> None:
    textos = _pessoal((PASTA / f"{nome}.txt").read_text(encoding="utf-8")).split("\f")
    assert NOME in "".join(textos)  # a fixture recebeu os dados fictícios
    original = extrair_textos(f"{nome}.txt", textos)
    assert original.rejeitadas == [] and original.notas
    mapa = _mapa()
    copia = conferir(f"{nome}.txt", textos, mapa)
    assert copia.recusas == []
    assert _sobrou(copia.textos) == []
    anon = extrair_textos(f"{nome}.txt", copia.textos)
    for a, b in zip(original.notas, anon.notas, strict=True):
        assert b.conta == mapa.contas[f"{a.corretora}|{a.conta}"] != a.conta
        assert a.model_dump(exclude={"conta"}) == b.model_dump(exclude={"conta"})
    assert [len(t) for t in copia.textos] == [len(t) for t in textos]
    assert [t.count("\n") for t in copia.textos] == [t.count("\n") for t in textos]


def test_a_mesma_conta_vira_a_mesma_ficticia_em_notas_diferentes() -> None:
    mapa = _mapa()
    contas = set()
    for nome in (
        "xp_2026_08_bovespa",
        "xp_2024_01_bmf_novo",
        "xp_2016_11_bovespa_antigo_multipagina",
    ):
        textos = _pessoal((PASTA / f"{nome}.txt").read_text(encoding="utf-8")).split("\f")
        [nota] = extrair_textos(nome, conferir(nome, textos, mapa).textos).notas
        contas.add(nota.conta)
    assert contas == {mapa.contas[f"XP|{CONTA}"]}


# ----------------------------------------------------------------------------- portão recusa
def test_portao_de_fidelidade_recusa_e_diz_a_nota_e_o_campo() -> None:
    """Conta igual ao número da nota: trocar a conta trocaria o número; a cópia é recusada."""
    texto = _pessoal((PASTA / "xp_2026_08_bovespa.txt").read_text(encoding="utf-8"))
    [nota] = extrair_textos("x", texto.split("\f")).notas
    texto = texto.replace(CONTA, nota.numero)
    copia = conferir("x.txt", texto.split("\f"), _mapa())
    assert not copia.aceita
    assert any(f"nota {nota.numero}" in r and "campo numero" in r for r in copia.recusas)


def test_portao_de_privacidade_recusa_termo_da_lista_local() -> None:
    texto = _pessoal((PASTA / "xp_2026_08_bovespa.txt").read_text(encoding="utf-8"))
    copia = conferir("x.txt", texto.split("\f"), _mapa(), sensiveis=["sala 201"])
    assert copia.recusas == ["página 1: termo da lista local de termos sensíveis na cópia"]


def test_portao_de_privacidade_acha_termo_que_ficou() -> None:
    """Palavra trocada numa página e esquecida em outra; a parte pública não conta."""
    original = ["RUA DAS ACACIAS\nXP INVESTIMENTOS", "Acacias aqui"]
    copia = ["PEO BUS ZILOFUQ\nXP INVESTIMENTOS", "Acacias aqui"]
    termos = {"ACACIAS", "INVESTIMENTOS"}
    assert vazamentos(copia, termos, publico="XP INVESTIMENTOS") == [
        "página 2: palavra pessoal de 7 letras na cópia"
    ]
    assert divergencias_de_forma(original, copia) == []
    assert divergencias_de_forma(original, ["RUA", "x"]) == [
        "página 1: texto fora do lugar", "página 2: texto fora do lugar"
    ]  # fmt: skip


def test_nome_do_arquivo_com_dado_pessoal_e_recusado() -> None:
    texto = _pessoal((PASTA / "xp_2026_08_bovespa.txt").read_text(encoding="utf-8"))
    copia = conferir("nota joana pereira.txt", texto.split("\f"), _mapa())
    assert "o nome do arquivo tem dado pessoal" in copia.recusas


# ----------------------------------------------------------------------------- layouts reais
def _btg_bovespa() -> str:
    return f"""NOTA DE CORRETAGEM
2907027
 Nr. nota
1
 Folha
04/01/2023
 Data pregão
Avenida Brigadeiro Faria Lima, 3477 - 14º Andar ITAIM BIBI
C.N.P.J: 43.815.158/0001-22
04538-133 SÃO PAULO - SP
BTG Pactual CTVM S.A.
SAC: 0800-722-2827 Ouvidoria: Tel. 0800-722-0048
e-mail ouvidoria: ouvidoria@btgpactual.com
 Cliente
00{CONTA}
 C.P.F./C.N.P.J/C.V.M./C.O.B.
{NOME} {CPF_TRACO}
RUA DAS ACACIAS, 742 - APTO 51 74210030 GOIANIA
GOIAS
Tel. (62) 987654321 Código cliente
85-0 00{CONTA}
 Assessor
 Participante destino do repasse
{ASSESSOR}
 Cliente
-
 Banco Acionista Administrador Complemento nome P.Vinc
N
 Agência Conta Corrente
Negócios realizados
"""


def _btg_bmf() -> str:
    return f"""NOTA DE CORRETAGEM
5474
 Nr. nota
1
 Folha
02/01/2023
 Data pregão
 C.N.P.J.
Numero da Corretora
85-0
Corretora
Tel: (011) 3046 - 2100 / 3443 - 2100 43.815.158/0001-22
BTG Pactual CTVM S.A.
SAC: 0800-722-2827 Ouvidoria: Tel. 0800-722-0048
Cliente
74210030
C.N.P.J/C.P.F
{NOME} {CPF_TRACO}
RUA DAS ACACIAS, 742 - APTO 51 Código cliente
GOIANIA - GOIAS {CONTA}
C/V Mercadoria Vencimento Quantidade Preço / Ajuste Tipo Negócio Valor Operação / D/C Taxa Operacional
"""  # noqa: E501


def _xp_bovespa_antiga() -> str:
    return f"""NOTA DENEGOCIAÇÃO
Nr. nota
8060527
Folha
1
Data pregão
04/01/2017
XPINVESTIMENTOS CCTVMS/A
AVATAULFODEPAIVA, 153 - SALA201 LEBLON 22440-032 RIODEJANEIRO- RJ
C.N.P.J: 02.332.886/0001-04 Carta Patente:
Ouvidoria: Tel. 0800.722.0202 e-mail ouvidoria: ouvidoria@xpi.com.br
Cliente
0{CONTA} JOANAPEREIRADOSSANTOS
RuadasAcácias, 742 - Apto 51 SetorBueno Tel. (062) 98765-4321
74210-030 GOIANIA- GO
Participante destino do repasse
-
Cliente
0
Valor
Banco Agência Conta corrente Acionista Administrador
{CPF_PONTO}
C.P.F./C.N.P.J/C.V.M./C.O.B.
Código cliente
3-5 0{CONTA}
Assessor
{ASSESSOR}
"""


def _xp_bovespa_nova() -> str:
    return f"""NOTA DE NEGOCIAÇÃO
Nr. nota
99116670
XP INVESTIMENTOS CORRETORA DE CÂMBIO, TÍTULOS E VALORES MOBILIÁRIOS S.A.
C.N.P.J: 02.332.886/0001-04 Carta Patente: Ouvidoria: Tel. 0800 722 3730 e-mail ouvidoria:
Cliente
{CONTA} {NOME}
Rua das Acácias, 742 - apto 51 - Setor Bueno Tel. (62) 98765-
4321
74210-030 Goiania - GO
Participante destino do repasse
-
Cliente
0
Valor
Banco Agência Conta corrente Acionista Administrador
{CPF_TRACO}
C.P.F./C.N.P.J/C.V.M./C.O.B.
Código cliente
3-5 {CONTA}
Assessor
{ASSESSOR}
"""


def _xp_bmf(nome: str, ouvidoria_depois: bool) -> str:
    """BM&F da XP: antiga (palavras coladas) e a que vem junto de um PDF Bovespa, com a linha
    da ouvidoria depois do endereço."""
    ouvidoria = "Ouvidoria: 0800 722 3730 e-mail ouvidoria:"
    return f"""C/V Mercadoria Vencimento Quantidade Preço/Ajuste Tipo Negócio Vlr de Operação/Ajuste D/C Taxa Operacional
NOTA DE NEGOCIAÇÃO Nr. nota
410.490
XP INVESTIMENTOS CORRETORA DE CÂMBIO, TÍTULOS E VALORES
Internet: https://www.xpi.com.br/ SAC: 0800-772-0202 e-mail:{"" if ouvidoria_depois else " " + ouvidoria}
{nome}
Rua das Acacias, 742 - Apto 51
74210-030 Setor Bueno GOIANIA - GO
{ouvidoria if ouvidoria_depois else "Corretora"}
{"Corretora" if ouvidoria_depois else "Cliente"}
Cliente
C.N.P.J.
02.332.886/0001-04
Numero da Corretora
3-5
C.N.P.J/C.P.F
{CPF_TRACO}
Codigo do Cliente
0{CONTA}
"""  # noqa: E501


def _rico() -> str:
    return f"""NOTA DENEGOCIAÇÃO
Nr. nota
376118
Rico Investimentos -GrupoXP
C.N.P.J: 02.332.886/0016-82 Carta Patente:
Ouvidoria: Tel. 08007740402 e-mail ouvidoria:
Cliente
0{CONTA} JOANAPEREIRADOSSANTOS
BLRUADASAC?CIAS, BLOCOA, 742 - APTO51
SETORBUENO
Tel. (062) 98765-4321
74210-030 GOIANIA- GO
Participante destino do repasse
- 0
Cliente
0
Valor
237
Banco
4321
Agência
765432
Conta corrente Acionista Administrador
{CPF_PONTO}
C.P.F./C.N.P.J/C.V.M./C.O.B.
Código cliente
386-7 0{CONTA}
Assessor
{ASSESSOR}
"""


def _agora() -> str:
    return f"""NOTA DE CORRETAGEM
Nr.Nota Folha Data pregão
2588033 1 18/04/2008
AGORA CORRETORA DE TITULOS E VALORES MOBILIARIOS S/A
C.N.P.J.: 74.014.747/0001-35 Carta Patente: A-39
Ouvidoria : Tel. (55 800) 727-9933 e-mail ouvidoria:
Cliente C.P.F./C.N.P.J./C.V.M./C.O.B.
0{CONTA}-4 {NOME} {CPF_TRACO}
RUA DAS ACACIAS, 742 - AP 51 SETOR BUENO Código cliente Assessor
74210-030 GOIANIA - GO (062) 98765-4321 39 6- {CONTA} 4- 30
Agente de Compensação Cliente Valor Custodiante C.I.
 0- 0 - 0,00 C - N
Banco Agência Conta corrente Acionista Administrador Complemento nome P. Vinc
237 4321 765432 N
Negócios realizados
"""


CABECALHOS = {
    "btg_bovespa": (_btg_bovespa, ("BTG Pactual CTVM S.A.", "43.815.158/0001-22", "85-0 ")),
    "btg_bmf": (_btg_bmf, ("BTG Pactual CTVM S.A.", "Código cliente", "C/V Mercadoria")),
    "xp_bovespa_antiga": (_xp_bovespa_antiga, ("XPINVESTIMENTOS", "RIODEJANEIRO", "3-5 ")),
    "xp_bovespa_nova": (_xp_bovespa_nova, ("XP INVESTIMENTOS", "Assessor", "Código cliente")),
    "xp_bmf_antiga": (lambda: _xp_bmf("JOANAPEREIRADOSSANTOS", False), ("Codigo do Cliente",)),
    "xp_bmf_junto_da_bovespa": (lambda: _xp_bmf(NOME, True), ("Corretora\nCliente",)),
    "rico": (_rico, ("Rico Investimentos", "Banco\n", "Agência\n", "386-7 ")),
    "agora": (_agora, ("AGORA CORRETORA", "Código cliente Assessor", "Agente de Compensação")),
}


@pytest.mark.parametrize("layout", sorted(CABECALHOS))
@pytest.mark.parametrize("fim_de_linha", ["\n", "\r\n"])
def test_cabecalho_real_sem_dado_pessoal_e_corretora_intacta(
    layout: str, fim_de_linha: str
) -> None:
    gerar, publicos = CABECALHOS[layout]
    textos = [gerar().replace("\n", fim_de_linha)]
    copia = conferir(f"{layout}.txt", textos, _mapa())
    assert copia.recusas == []  # o leitor recusa o cabeçalho sozinho, igual no original
    assert _sobrou(copia.textos) == []
    for publico in publicos:
        assert publico.replace("\n", fim_de_linha) in copia.textos[0]
    assert len(copia.textos[0]) == len(textos[0])


def test_conta_no_formato_sinacor_vira_a_ficticia_do_mapa() -> None:
    mapa = _mapa()
    [texto] = anonimizar([_xp_bovespa_antiga()], mapa).textos
    assert f"3-5 0{mapa.contas['XP|' + CONTA]}" in texto
    assert f"0{mapa.contas['XP|' + CONTA]} " in texto  # a da linha do nome, igual


def test_nome_detectado_em_uma_pagina_e_trocado_nas_outras() -> None:
    """Página que não diz onde está o nome (sem CPF): ele é trocado pelo das outras páginas."""
    textos = [_xp_bovespa_nova(), f"Folha 2\n{NOME}\nRua das Acácias\n"]
    copia = anonimizar(textos, _mapa())
    assert all("JOANA" not in t for t in copia.textos)
