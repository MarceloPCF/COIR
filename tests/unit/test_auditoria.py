"""Registro de auditoria e portão de privacidade. Só dados fictícios."""

import ast
import json
import logging
import re
import zipfile
from collections.abc import Iterator
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path

import pytest
from hypothesis import given
from hypothesis import strategies as st
from typer.testing import CliRunner

from coir import auditoria
from coir.auditoria import portao, registro
from coir.auditoria.catalogo import CAMPOS, EVENTOS, OPERACOES, PREFIXOS, Classe
from coir.auditoria.limpeza import (
    diferenca,
    esqueleto,
    grandeza,
    impresso,
    limpar,
    texto_limpo,
    valido,
)
from coir.auditoria.vocabulario import PALAVRAS
from coir.cli import VAR_SEM_REGISTRO, app
from coir.config import VAR_DADOS
from coir.db.sessao import NOME_BANCO, migrar
from coir.erros import AnonimizacaoError, ConciliacaoError, ExtracaoError, ImportacaoError
from coir.parsing.anonimizar import DadosPessoais
from coir.parsing.extrair import Rejeitada, ResultadoExtracao
from coir.referencias import carregar, carregar_fiscais, pasta_efetiva
from coir.servicos import suporte, trilha
from coir.servicos.apuracao import apurar_banco
from coir.servicos.cadastros import adicionar_titular
from coir.servicos.importacao import importar_lista
from tests.apoio import cpf_de_teste
from tools import gerar_vocabulario, notas_ficticias

AGORA = datetime(2026, 1, 15, 10, 30, 0)
CPF = cpf_de_teste()
CPF_MASCARA = f"{CPF[:3]}.{CPF[3:6]}.{CPF[6:9]}-{CPF[9:]}"
NOME = "Zuleika Quixabeira Tabajara"
CONTA = "7654321"
ENDERECO = "Travessa Inventada 742"
EMAIL = "zuleika.tabajara@exemplo.com.br"
TELEFONE = "(11) 91234-5678"
CEP = "01310-100"
CAMINHO = r"C:\Users\zuleika\Documentos\notas"
SEGREDO = bytes(32)


@pytest.fixture
def dados(tmp_path: Path) -> Iterator[Path]:
    """Pasta de dados com o registro ligado."""
    pasta = tmp_path / "dados"
    pasta.mkdir()
    auditoria.configurar(pasta, "9.9.9", "cli")
    yield pasta
    auditoria.desligar()


def _eventos(pasta: Path, nome: str | None = None) -> list[dict[str, object]]:
    linhas = (pasta / auditoria.ARQUIVO).read_text(encoding="utf-8").splitlines()
    todos = [json.loads(x) for x in linhas]
    return [e for e in todos if nome is None or e["ev"] == nome]


def _dados_do(evento: dict[str, object]) -> dict[str, object]:
    conteudo = evento["d"]
    assert isinstance(conteudo, dict)
    return conteudo


def _linha(evento: str, /, **campos: object) -> str:
    return json.dumps(
        {"t": "2026-01-15T10:30:00.000-03:00", "s": "0a1b2c3d", "n": 1, "nv": "I", "ev": evento,
         "d": campos},
        ensure_ascii=False,
    )  # fmt: skip


TERMOS = portao.Termos.de(
    [NOME, CPF, CONTA, ENDERECO, EMAIL, TELEFONE, CEP], ["tabajara"], ["segredo-teste-xyz"]
)


# ------------------------------------------------------------------------------------ catálogo
def test_catalogo_e_coerente() -> None:
    """Todo campo de evento tem classe; todo pseudônimo tem prefixo."""
    for nome, campos in EVENTOS.items():
        assert campos <= set(CAMPOS), nome
    for campo in PREFIXOS:
        assert CAMPOS[campo] is Classe.PSEUDONIMO
    assert {c for c, k in CAMPOS.items() if k is Classe.PSEUDONIMO} == set(PREFIXOS)


def _chamadas_do_codigo() -> Iterator[tuple[Path, ast.Call]]:
    for arquivo in Path("coir").rglob("*.py"):
        if arquivo.parent.name == "auditoria":
            continue
        for no in ast.walk(ast.parse(arquivo.read_text(encoding="utf-8"))):
            if (
                isinstance(no, ast.Call)
                and isinstance(no.func, ast.Attribute)
                and isinstance(no.func.value, ast.Name)
                and no.func.value.id == "auditoria"
            ):
                yield arquivo, no


def test_todo_evento_do_codigo_esta_no_catalogo() -> None:
    """O que o programa registra está no catálogo: senão o campo seria recusado em silêncio no
    uso real e o portão reteria a linha."""
    vistos = 0
    for arquivo, chamada in _chamadas_do_codigo():
        assert isinstance(chamada.func, ast.Attribute)
        if chamada.func.attr not in ("evento", "operacao"):
            continue
        primeiro = chamada.args[0]
        if not isinstance(primeiro, ast.Constant):  # trilha.avisos: o nome vem de quem chama
            continue
        vistos += 1
        nome = str(primeiro.value)
        permitidos = EVENTOS[f"{nome}.inicio"] if chamada.func.attr == "operacao" else EVENTOS[nome]
        for k in chamada.keywords:
            if k.arg is not None:
                assert k.arg in permitidos, f"{arquivo}: {nome}.{k.arg}"
    assert vistos > 10


def test_campos_do_fim_de_operacao_estao_no_catalogo() -> None:
    for nome, (_, fim) in OPERACOES.items():
        assert set(fim.split()) <= EVENTOS[f"{nome}.fim"]


# ------------------------------------------------------------------------------------- limpeza
@pytest.mark.parametrize(
    ("original", "proibidos"),
    [
        (f"nota do {NOME}", ["Zuleika", "Quixabeira", "Tabajara"]),
        (f"erro no CPF {CPF_MASCARA}", [CPF[:3], CPF_MASCARA]),
        (f"erro no CPF {CPF}", [CPF]),
        (f"conta 85-0 {CONTA}", [CONTA]),
        (f"e-mail {EMAIL}", ["zuleika", "exemplo"]),
        (f"telefone {TELEFONE}", ["1234", "5678"]),
        (f"CEP {CEP}", ["01310"]),
        (f"não leu {CAMINHO}", ["zuleika", "Users"]),
        (f"arquivo Nota_{NOME.split()[0]}.pdf", ["Zuleika"]),
        (f"[parameters: ('{CPF}', '{NOME}')] falhou", [CPF, "Zuleika"]),
        ("líquido 1.234,56 e valor 98765.43", ["1.234,56", "98765.43", "234", "765"]),
        (f"https://exemplo.com/{NOME}", ["exemplo", "Zuleika"]),
    ],
)
def test_texto_limpo_tira_o_que_identifica(original: str, proibidos: list[str]) -> None:
    limpo = texto_limpo(original)
    for termo in proibidos:
        assert termo not in limpo


def test_texto_limpo_mantem_o_que_ajuda_a_corrigir() -> None:
    """A mensagem do programa e a forma da linha ficam: é o que mostra o layout."""
    limpo = texto_limpo("nota 12345: linha de negócio fora do padrão: 'C VISTA PETR4 100 28,50 D'")
    assert limpo == "nota 99999: linha de negócio fora do padrão: 'C VISTA PETR4 999 99,99 D'"
    assert texto_limpo("Inesperadíssima") == "«L15»"
    assert texto_limpo("página 3, folha 2") == "página 3, folha 2"
    assert texto_limpo("planilha da B3") == "planilha da B3"


def test_esqueleto_mostra_o_layout_sem_dado() -> None:
    pagina = f"NOTA DE CORRETAGEM\n{NOME}  CPF {CPF_MASCARA}\nConta 85-0 {CONTA}\nPETR4  100  28,50"
    linhas = esqueleto(pagina)
    assert linhas[0] == "NOTA DE CORRETAGEM"
    texto = "\n".join(linhas)
    assert "Zuleika" not in texto and CPF[:3] not in texto and CONTA not in texto
    assert "999.999.999-99" in texto and "PETR4" in texto


@given(st.text(max_size=200))
def test_a_limpeza_e_estavel(texto: str) -> None:
    """O portão exige que limpar de novo não mude nada: tem de valer para qualquer texto."""
    limpo = texto_limpo(texto)
    assert texto_limpo(limpo) == limpo
    linha = texto_limpo(texto, 200, espacos=True)
    assert texto_limpo(linha, 200, espacos=True) == linha


def test_grandeza_e_diferenca_nao_mostram_o_valor() -> None:
    assert grandeza(Decimal("1234.56")) == "+1e3"
    assert grandeza(Decimal("-98765.43")) == "-1e4"
    assert grandeza(Decimal("0.42")) == "+1e0"
    assert grandeza(0) == "0"
    assert diferenca(Decimal("0.06")) == "+0.06"
    assert diferenca(Decimal("-99.99")) == "-99.99"
    assert diferenca(Decimal("100.01")) == "+1e2"  # acima de 100, só a grandeza
    for ruim in ("x", None, True, Decimal("NaN")):
        with pytest.raises(ValueError):
            grandeza(ruim)


def test_impresso_so_em_maiusculas_e_sem_numero_longo() -> None:
    assert impresso("VALE ON NM") == "VALE ON NM"
    assert impresso("ITAÚ PN 12345") == "ITAU PN 99999"
    assert impresso(f"nome {NOME}") == texto_limpo(f"nome {NOME}", 60)


@pytest.mark.parametrize("valor", ["JOSE DA SILVA", "RUA DAS FLORES 12", CPF_MASCARA, "ZULEIKA"])
def test_impresso_sem_especie_nao_passa_como_nome_de_ativo(valor: str) -> None:
    """Se a leitura errar a coluna, o nome ou o endereço do cliente não têm a forma de um ativo."""
    assert impresso(valor) == texto_limpo(valor, 60)
    assert "SILVA" not in impresso(valor) and "FLORES" not in impresso(valor)
    assert CPF[-2:] not in impresso(valor)


def test_codigo_de_negociacao_e_nome_nao_se_confundem() -> None:
    for ticker in ("PETR4", "BOVA11", "B3SA3", "KLBNO193E", "WINZ24", "A1AP34", "PETRK300W1",
                   "5GTK11", "ITUB4F"):  # fmt: skip
        assert valido(Classe.ATIVO, ticker), ticker
    for nao in ("ZULEIKA", "Zuleika", CPF, "12345", "", "XP1234567", "BTG0012345",
                f"{CPF}A", "1234567890A", "AB12345"):  # fmt: skip
        assert not valido(Classe.ATIVO, nao), nao
        assert nao == "" or texto_limpo(f"conta {nao}") != f"conta {nao}"
    with pytest.raises(ValueError):
        limpar(Classe.ATIVO, "ZULEIKA")


# ----------------------------------------------------------------------------------- gravação
def test_sem_registro_ligado_nada_acontece(tmp_path: Path) -> None:
    assert not auditoria.ativa()
    auditoria.evento("ui.tela", rota="/")
    assert not auditoria.uma_vez("x")
    with auditoria.operacao("apuracao", ate=date(2026, 1, 31)) as op:
        op.dados(meses=1)
    assert not list(tmp_path.iterdir())


def test_inicio_de_sessao_diz_versao_e_sistema(dados: Path) -> None:
    [inicio] = _eventos(dados, "sessao.inicio")
    conteudo = _dados_do(inicio)
    assert conteudo["versao"] == "9.9.9" and conteudo["modo"] == "cli"
    assert conteudo["executavel"] is False and "python" in conteudo and "sistema" in conteudo
    assert inicio["nv"] == "I" and inicio["n"] == 1


def test_pseudonimo_e_estavel_e_nao_mostra_o_valor(dados: Path) -> None:
    for _ in range(2):
        auditoria.evento("extracao.nota", arquivo=f"Nota {NOME}.pdf", nota="123", conta=CONTA)
    primeiro, segundo = _eventos(dados, "extracao.nota")
    assert _dados_do(primeiro) == _dados_do(segundo)
    conteudo = _dados_do(primeiro)
    assert str(conteudo["arquivo"]).startswith("arq_") and str(conteudo["conta"]).startswith("cta_")
    bruto = (dados / auditoria.ARQUIVO).read_text(encoding="utf-8")
    assert "Zuleika" not in bruto and CONTA not in bruto


def test_segredo_fica_em_privado_e_vale_entre_sessoes(tmp_path: Path) -> None:
    auditoria.configurar(tmp_path, "9.9.9")
    um = auditoria.segredo()
    auditoria.desligar()
    auditoria.configurar(tmp_path, "9.9.9")
    assert um is not None and auditoria.segredo() == um
    auditoria.desligar()
    assert (tmp_path / auditoria.ARQUIVO_SEGREDO).is_file()
    assert (tmp_path / auditoria.ARQUIVO_SEGREDO).parent.name == "_privado"
    assert auditoria.pseudonimo("conta", CONTA, um) != auditoria.pseudonimo("conta", CONTA, SEGREDO)


def test_campo_fora_do_catalogo_nao_entra(dados: Path, caplog: pytest.LogCaptureFixture) -> None:
    with caplog.at_level(logging.WARNING):
        auditoria.evento("ui.tela", rota="/", nome=NOME, cpf=CPF)
    [e] = _eventos(dados, "ui.tela")
    assert _dados_do(e) == {"rota": "/"}
    assert "['cpf', 'nome']" in caplog.text and "Zuleika" not in caplog.text


def test_valor_que_nao_cabe_na_classe_nao_entra(dados: Path) -> None:
    auditoria.evento("extracao.nota", negocios=int(CPF), data="ontem", corretora="XP")
    [e] = _eventos(dados, "extracao.nota")
    assert _dados_do(e) == {"corretora": "XP"}
    assert CPF not in (dados / auditoria.ARQUIVO).read_text(encoding="utf-8")


def test_operacao_registra_tempo_e_junta_os_eventos(dados: Path) -> None:
    with auditoria.operacao("apuracao", ate=date(2026, 1, 31)) as op:
        auditoria.evento("apuracao.aviso", "A", texto="aviso")
        op.dados(meses=3)
    inicio, aviso, fim = (*_eventos(dados, "apuracao.inicio"), *_eventos(dados, "apuracao.aviso"),
                          *_eventos(dados, "apuracao.fim"))  # fmt: skip
    assert inicio["op"] == aviso["op"] == fim["op"]
    conteudo = _dados_do(fim)
    assert conteudo["ok"] is True and conteudo["meses"] == 3 and isinstance(conteudo["ms"], int)


def test_operacao_que_falha_registra_pilha_sem_caminho_e_sobe(dados: Path) -> None:
    with pytest.raises(ValueError, match="falhou"), auditoria.operacao("apuracao"):
        raise ValueError(f"falhou em {CAMINHO} para {NOME}")
    [fim] = _eventos(dados, "apuracao.fim")
    conteudo = _dados_do(fim)
    assert fim["nv"] == "E" and conteudo["ok"] is False and conteudo["classe"] == "ValueError"
    pilha = conteudo["pilha"]
    assert isinstance(pilha, list) and any("test_auditoria" in q for q in pilha)
    bruto = (dados / auditoria.ARQUIVO).read_text(encoding="utf-8")
    assert "zuleika" not in bruto.casefold() and "Users" not in bruto and "tests\\" not in bruto


def test_erro_com_texto_quebrado_nao_esconde_o_erro(dados: Path) -> None:
    class Quebrado(ImportacaoError):
        def __str__(self) -> str:
            raise RuntimeError("defeito")

    with pytest.raises(Quebrado), auditoria.operacao("apuracao"):
        raise Quebrado("x")
    assert _dados_do(_eventos(dados, "apuracao.fim")[-1])["classe"] == "Quebrado"


def test_operacao_tambem_e_decorador(dados: Path) -> None:
    @auditoria.operacao("restauracao")
    def fazer() -> int:
        return 7

    assert fazer() == 7
    assert fazer() == 7
    assert len(_eventos(dados, "restauracao.fim")) == 2


def test_erro_do_coir_fica_registrado_ao_ser_criado(dados: Path) -> None:
    """A tela só mostra a mensagem: sem isso o erro não ficaria em lugar nenhum."""
    try:
        raise ImportacaoError(f"conta {CONTA} de {NOME}")
    except ImportacaoError:
        pass
    [e] = _eventos(dados, "erro.dominio")
    conteudo = _dados_do(e)
    assert conteudo["classe"] == "ImportacaoError" and e["nv"] == "A"
    assert conteudo["modulo"] == __name__
    assert CONTA not in json.dumps(e) and "Zuleika" not in json.dumps(e)


def test_nota_rejeitada_nao_gera_erro_de_dominio(dados: Path) -> None:
    """Rejeição é caminho normal: é registrada pelo resultado, com mais detalhe."""
    ExtracaoError("linha fora do padrão")
    ConciliacaoError("valor", Decimal(1), Decimal(2))
    assert _eventos(dados, "erro.dominio") == []


def test_ponte_traz_avisos_e_erros_com_texto_limpo(dados: Path) -> None:
    log = logging.getLogger("coir.servicos.qualquer")
    log.info("backup de %s", NOME)
    log.warning("conta %s", CONTA)
    logging.getLogger("biblioteca.externa").info("não interessa")
    logging.getLogger("biblioteca.externa").warning("interessa")
    try:
        raise KeyError(CPF)
    except KeyError:
        log.exception("quebrou em %s", CAMINHO)
    eventos = _eventos(dados, "log.registro")
    assert [e["nv"] for e in eventos] == ["I", "A", "A", "E"]
    ultimo = _dados_do(eventos[-1])
    assert ultimo["classe"] == "KeyError" and ultimo["modulo"] == "coir.servicos.qualquer"
    bruto = (dados / auditoria.ARQUIVO).read_text(encoding="utf-8")
    for termo in (CPF, CONTA, "Zuleika", "zuleika"):
        assert termo not in bruto


def test_ponte_limita_a_mensagem_repetida(dados: Path) -> None:
    log = logging.getLogger("biblioteca.ruidosa")
    for _ in range(registro.REPETICOES_DA_PONTE + 20):
        log.warning("aviso a cada página")
    assert len(_eventos(dados, "log.registro")) == registro.REPETICOES_DA_PONTE


def test_erro_nao_tratado_fica_registrado(dados: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    """Direto no registro, com a pilha e sem repetir no console."""
    import sys

    chamado: list[str] = []
    auditoria.desligar()
    monkeypatch.setattr(sys, "excepthook", lambda *_: chamado.append("anterior"))
    auditoria.configurar(dados, "9.9.9")
    try:
        raise RuntimeError(NOME)
    except RuntimeError as erro:
        sys.excepthook(RuntimeError, erro, erro.__traceback__)
    assert chamado == ["anterior"]
    ultimo = _eventos(dados, "log.registro")[-1]
    assert ultimo["nv"] == "C" and _dados_do(ultimo)["classe"] == "RuntimeError"


def test_ganchos_ignoram_saida_normal_e_nao_desfazem_o_gancho_alheio(dados: Path) -> None:
    import sys

    sys.excepthook(SystemExit, SystemExit(0), None)
    assert _eventos(dados, "log.registro") == []
    alheio = sys.excepthook = lambda *_: None
    auditoria.desligar()
    assert sys.excepthook is alheio
    sys.excepthook = sys.__excepthook__


def test_rotacao_presa_por_outro_processo_nao_perde_linhas(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """Windows: a janela e a linha de comando abertas. Continua escrevendo; os antigos ficam."""
    monkeypatch.setattr(registro, "TAMANHO_MAXIMO", 500)
    auditoria.configurar(tmp_path, "9.9.9")

    def preso(_origem: object, _destino: object) -> None:
        raise PermissionError("arquivo em uso")

    monkeypatch.setattr(registro.os, "replace", preso)
    for _ in range(50):
        auditoria.evento("ui.tela", rota="/apuracao")
    monkeypatch.undo()
    auditoria.desligar()
    assert len(_eventos(tmp_path, "ui.tela")) == 50
    assert auditoria.arquivos_do_registro(tmp_path) == [tmp_path / auditoria.ARQUIVO]


@pytest.mark.parametrize("gravado", ["", "abc", "zz", "00" * 10])
def test_segredo_vazio_ou_curto_e_refeito(tmp_path: Path, gravado: str) -> None:
    """Um segredo vazio deixaria os pseudônimos fracos e travaria a exportação."""
    (tmp_path / "_privado").mkdir()
    (tmp_path / auditoria.ARQUIVO_SEGREDO).write_text(gravado, encoding="utf-8")
    auditoria.configurar(tmp_path, "9.9.9")
    chave = auditoria.segredo()
    auditoria.desligar()
    assert chave is not None and len(chave) == 32
    assert (tmp_path / auditoria.ARQUIVO_SEGREDO).read_text(encoding="utf-8") == chave.hex()


def test_segredo_vazio_nao_vira_termo() -> None:
    termos = portao.Termos.de([], segredos=[""])
    assert portao.conferir_linha(_linha("importacao.aviso", texto="nota"), termos)[0] is not None


def test_rotacao_guarda_os_antigos_em_ordem(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setattr(registro, "TAMANHO_MAXIMO", 2000)
    auditoria.configurar(tmp_path, "9.9.9")
    for _ in range(100):
        auditoria.evento("ui.tela", rota="/apuracao")
    auditoria.desligar()
    arquivos = auditoria.arquivos_do_registro(tmp_path)
    assert 2 <= len(arquivos) <= registro.ARQUIVOS_ANTIGOS + 1
    assert arquivos[-1] == tmp_path / auditoria.ARQUIVO
    numeros = [json.loads(x)["n"] for a in arquivos
               for x in a.read_text(encoding="utf-8").splitlines()]  # fmt: skip
    assert numeros == sorted(numeros)


def test_configurar_na_mesma_pasta_nao_repete_a_troca_muda(tmp_path: Path) -> None:
    um, dois = tmp_path / "um", tmp_path / "dois"
    um.mkdir()
    dois.mkdir()
    auditoria.configurar(um, "9.9.9")
    auditoria.configurar(um, "9.9.9")
    auditoria.configurar(dois, "9.9.9")
    auditoria.desligar()
    assert len(_eventos(um, "sessao.inicio")) == 1 and len(_eventos(dois, "sessao.inicio")) == 1


# -------------------------------------------------------------------- portão: linha a linha
@pytest.mark.parametrize(
    "valor",
    [NOME, NOME.upper(), "Quixabeira", CPF, CPF_MASCARA, CONTA, f"85-0 {CONTA}", EMAIL, TELEFONE,
     CEP, CAMINHO, ENDERECO, "TABAJARA", "segredo-teste-xyz"],
)  # fmt: skip
def test_portao_retem_dado_pessoal_em_qualquer_campo_de_texto(valor: str) -> None:
    """A linha foi adulterada ou veio de um programa com defeito: dado cru não passa."""
    for campo, conteudo in (("texto", valor), ("texto", f"erro com {valor} no meio")):
        objeto, motivo = portao.conferir_linha(_linha("importacao.aviso", **{campo: conteudo}),
                                               TERMOS)  # fmt: skip
        assert objeto is None, conteudo
        assert motivo and valor not in motivo


@pytest.mark.parametrize(
    "valor", [NOME.upper(), "QUIXABEIRA", CPF, f"85-0 {CONTA}", "ZULEIKA TABAJARA"]
)
def test_portao_retem_dado_pessoal_no_nome_impresso(valor: str) -> None:
    """Nome de ativo é a única classe fora do vocabulário: aí valem os termos do titular."""
    objeto, _ = portao.conferir_linha(_linha("extracao.rejeicao", lido=valor), TERMOS)
    assert objeto is None


@pytest.mark.parametrize(
    "linha",
    [
        "isto não é JSON",
        "[1, 2, 3]",
        '{"t":"2026"}',
        _linha("evento.inventado", texto="ok"),
        _linha("ui.tela", rota="/", nome=NOME),
        _linha("ui.tela", cpf="ok"),
        _linha("extracao.nota", arquivo=f"Nota {NOME}.pdf"),
        _linha("extracao.nota", conta=CONTA),
        _linha("extracao.nota", negocios=int(CPF)),
        _linha("extracao.nota", negocios="3"),
        _linha("extracao.nota", data="15/01/2026"),
        _linha("apuracao.mes", imposto="1234.56"),
        _linha("apuracao.mes", imposto="+1e3", darf=1234.56),
        _linha("conciliacao.divergencia", diferenca="1234.56"),
        _linha("conciliacao.divergencia", ativo="ZULEIKA"),
        _linha("log.registro", pilha=[r"C:\Users\zuleika\x.py:f:1"]),
        _linha("log.registro", classe="Erro de Zuleika"),
        _linha("sessao.inicio", sistema="Windows zuleika-pc"),
        _linha("importacao.arquivo", avisos=True),
        _linha("importacao.aviso", texto={"aninhado": NOME}),
    ],
)
def test_portao_retem_linha_fora_do_catalogo_ou_da_forma(linha: str) -> None:
    assert portao.conferir_linha(linha, portao.Termos())[0] is None


@pytest.mark.parametrize(
    ("evento", "campo"),
    [("log.registro", "classe"), ("log.registro", "modulo"), ("sessao.inicio", "sistema"),
     ("ui.tela", "rota")],
)  # fmt: skip
def test_portao_retem_nome_nos_campos_de_codigo(evento: str, campo: str) -> None:
    """Essas classes só conferem a forma: "Quixabeira" cabe nela. Os termos do titular valem."""
    valor = "/Quixabeira" if campo == "rota" else "Quixabeira"
    assert portao.conferir_linha(_linha(evento, **{campo: valor}), TERMOS)[0] is None


@pytest.mark.parametrize("mascara", ["«Maria»", "«Zuleika2»", "«L7x»"])
def test_mascara_so_na_forma_que_a_limpeza_produz(mascara: str) -> None:
    assert (
        portao.conferir_linha(_linha("importacao.aviso", texto=mascara), portao.Termos())[0] is None
    )
    assert "Maria" not in texto_limpo("«Maria»") and "Zuleika" not in texto_limpo("«Zuleika2»")


def test_portao_retem_nome_na_pilha() -> None:
    linha = _linha("log.registro", pilha=["Quixabeira:funcao:12"], causas=["Quixabeira"])
    assert portao.conferir_linha(linha, TERMOS)[0] is None
    # o número da linha da pilha não é confundido com uma conta
    termos = portao.Termos.de(["1234"])
    limpa = _linha("log.registro", pilha=["coir.cli:main:1234"])
    assert portao.conferir_linha(limpa, termos)[0] is not None


def test_hora_impossivel_fica_retida() -> None:
    base = json.loads(_linha("ui.tela", rota="/"))
    base["t"] = "2026-13-45T10:00:00.000-03:00"
    assert portao.conferir_linha(json.dumps(base), TERMOS) == (None, "hora fora do formato")


def test_portao_nao_aceita_envelope_com_dado() -> None:
    base = json.loads(_linha("ui.tela", rota="/"))
    for chave, valor in (("t", NOME), ("s", CPF), ("n", int(CPF)), ("nv", "Z"), ("op", NOME),
                         ("extra", 1), ("ev", NOME), ("d", NOME)):  # fmt: skip
        objeto, _ = portao.conferir_linha(json.dumps({**base, chave: valor}), portao.Termos())
        assert objeto is None, chave


def test_portao_aceita_o_que_o_registro_grava(dados: Path) -> None:
    auditoria.evento("extracao.pagina", arquivo="a.pdf", pagina=1, corretora="XP",
                     segmento="BOVESPA", layout="xp", tipo="nota", folha=1, nota="9",
                     data=date(2026, 1, 15))  # fmt: skip
    auditoria.evento("extracao.rejeicao", "A", texto=f"nota 9: linha de {NOME} {CPF}",
                     lido="VALE ON NM", diferenca=Decimal("0.06"), paginas=(1, 2))  # fmt: skip
    auditoria.evento("apuracao.mes", titular=1, mes=date(2026, 1, 1), assinatura="x",
                     imposto=Decimal("1234.56"), isento=True, incompletos=["PETR4"])  # fmt: skip
    linhas = (dados / auditoria.ARQUIVO).read_text(encoding="utf-8").splitlines()
    conferencia = portao.conferir_linhas(linhas, TERMOS)
    assert not conferencia.retidas
    assert len(conferencia.linhas) == len(linhas)


def test_portao_refaz_o_texto_e_conta_as_retidas() -> None:
    boa = _linha("ui.tela", rota="/")
    com_espacos = boa.replace(",", " ,  ")
    conferencia = portao.conferir_linhas([com_espacos, "lixo", _linha("ui.tela", cpf="1")], TERMOS)
    assert conferencia.linhas == [json.dumps(json.loads(boa), ensure_ascii=False,
                                             separators=(",", ":"))]  # fmt: skip
    assert conferencia.retidas == {"linha ilegível": 1, "campo fora do catálogo": 1}


def test_formato_ja_limpo_nao_e_confundido_com_dado() -> None:
    """A forma de um CPF, telefone ou CEP (só noves) é o que a limpeza produz: não é achado."""
    texto = "CPF 999.999.999-99 tel 9999-9999 CEP 99999-999 CNPJ 99.999.999/9999-99"
    assert portao.conferir_linha(_linha("importacao.aviso", texto=texto), TERMOS)[0] is not None


def test_palavra_do_programa_no_nome() -> None:
    """Sobrenome que também é palavra do programa não recusa toda linha com a palavra, mas não
    passa na forma de uma página de nota, nem duas palavras seguidas do nome em lugar nenhum."""
    assert {"nota", "valor"} <= PALAVRAS
    termos = portao.Termos.de(["Nota Valor"])
    aviso = _linha("importacao.aviso", texto="nota lida")
    assert portao.conferir_linha(aviso, termos)[0] is not None
    pagina = _linha("extracao.esqueleto", esqueleto=["Cliente Nota"])
    assert portao.conferir_linha(pagina, termos)[0] is None
    junto = _linha("importacao.aviso", texto="nota valor")
    assert portao.conferir_linha(junto, termos)[0] is None


# ---------------------------------------------------------------- portão: texto final inteiro
@pytest.mark.parametrize("valor", [CPF, CPF_MASCARA, EMAIL, TELEFONE, CEP, CAMINHO, NOME,
                                   "segredo-teste-xyz"])  # fmt: skip
def test_segunda_passagem_recusa_a_exportacao_toda(valor: str) -> None:
    """A rede para um defeito na primeira passagem: basta um achado para não sair nada."""
    with pytest.raises(AnonimizacaoError, match="portão") as erro:
        portao.exigir_texto_limpo(_linha("ui.tela", rota="/") + "\n"
                                  + _linha("importacao.aviso", texto=valor), TERMOS)  # fmt: skip
    assert valor not in str(erro.value)


def test_segunda_passagem_recusa_texto_fora_do_formato() -> None:
    with pytest.raises(AnonimizacaoError):
        portao.exigir_texto_limpo("texto solto\n", TERMOS)


def test_termos_nao_aparecem_no_repr() -> None:
    assert "Zuleika" not in repr(TERMOS) and "tabajara" not in repr(TERMOS)


# ----------------------------------------------------------------------------- vocabulário
def test_vocabulario_nao_tem_nomes_comuns() -> None:
    for nome in ("marcelo", "silva", "santos", "maria", "pedro", "campos"):
        assert nome not in PALAVRAS


def test_vocabulario_esta_em_dia() -> None:
    """Pode ficar um pouco defasado (palavra nova só vira forma), não muito. Se falhar, rode
    ``uv run python tools/gerar_vocabulario.py``."""
    assert len(gerar_vocabulario.vocabulario() ^ PALAVRAS) < 150


# ------------------------------------------------------------------- do domínio para o registro
def test_rejeicao_mostra_campo_e_diferenca_da_conciliacao(dados: Path) -> None:
    erro = ConciliacaoError("valor_liquido", Decimal("1000.00"), Decimal("1000.06"), "a.pdf: ")
    trilha.rejeicao(Rejeitada("Nota Zuleika.pdf", (1,), f"nota 77: {erro}", "77",
                              date(2026, 1, 15), erro=erro))  # fmt: skip
    [e] = _eventos(dados, "extracao.rejeicao")
    conteudo = _dados_do(e)
    assert conteudo["classe"] == "ConciliacaoError"
    assert conteudo["campo"] == "valor_liquido" and conteudo["diferenca"] == "+0.06"
    assert "1000" not in json.dumps(conteudo) and "Zuleika" not in json.dumps(conteudo)
    assert conteudo["data"] == "2026-01-15" and str(conteudo["nota"]).startswith("not_")


def test_mesmo_layout_rejeitado_guarda_um_esqueleto_so(dados: Path) -> None:
    for numero in ("111", "222", "333"):
        pagina = f"CORRETORA NOVA\nNota {numero}\nCPF {CPF_MASCARA}"
        trilha.rejeicao(Rejeitada("a.pdf", (1,), "sem perfil", numero, esqueleto=esqueleto(pagina)))
    assert len(_eventos(dados, "extracao.rejeicao")) == 3
    assert len(_eventos(dados, "extracao.esqueleto")) == 1


# ------------------------------------------------------------------------- exportação completa
def _notas_ficticias(pasta: Path) -> list[Path]:
    roteiro = notas_ficticias.carregar()
    pasta.mkdir(parents=True)
    textos = notas_ficticias.textos(roteiro)
    arquivos = [notas_ficticias.pdf(t, pasta / f"{n}.pdf") for n, t in textos[:6]]
    # uma nota com o total adulterado (não concilia) e uma de layout desconhecido, com o nome
    # do titular no nome do arquivo e no texto
    nome, texto = next((n, x) for n, x in textos if "quido para" in x)
    linhas = texto.splitlines()
    alvo = next(i for i, x in enumerate(linhas) if "quido para" in x)
    linhas[alvo] = re.sub(r"\d,\d\d(?!.*\d,\d\d)", "7,77", linhas[alvo])
    assert linhas[alvo] != texto.splitlines()[alvo]
    arquivos.append(notas_ficticias.pdf("\n".join(linhas), pasta / f"{nome}-outra.pdf"))
    desconhecida = (f"CORRETORA INVENTADA DE VALORES\nNota 1234 de {NOME}\nCPF {CPF_MASCARA}\n"
                    f"{ENDERECO} {CEP}\n{EMAIL} {TELEFONE}\nConta 85-0 {CONTA}\n"
                    "Total geral 1.234,56\n")  # fmt: skip
    arquivos.append(notas_ficticias.pdf(desconhecida, pasta / f"Nota de {NOME} {CPF}.pdf"))
    return arquivos


@pytest.fixture
def exportado(tmp_path: Path) -> Iterator[tuple[Path, suporte.ExportacaoRegistro]]:
    """Importa e apura notas fictícias, com rejeições, e exporta o registro."""
    pasta = tmp_path / "dados"
    pasta.mkdir()
    banco = pasta / NOME_BANCO
    auditoria.configurar(pasta, "9.9.9")
    roteiro = notas_ficticias.carregar()
    tabelas = pasta_efetiva(pasta)
    refs, fiscais = carregar(tabelas), carregar_fiscais(tabelas)
    migrar(banco, AGORA)
    adicionar_titular(banco, roteiro.cpf, roteiro.titular)
    suporte.salvar_titular(pasta, DadosPessoais(
        cpf=CPF, nome=NOME, contas=[CONTA], endereco=ENDERECO, email=EMAIL, telefones=TELEFONE,
        cep=CEP,
    ))  # fmt: skip
    arquivos = _notas_ficticias(tmp_path / f"notas de {NOME}")
    importar_lista(banco, arquivos, refs, AGORA)
    apurar_banco(banco, refs, fiscais, date(2025, 12, 31))
    logging.getLogger("coir.teste").warning("abrindo %s de %s", arquivos[-1], NOME)
    feito = suporte.exportar_registro(pasta, banco, AGORA)
    yield pasta, feito
    auditoria.desligar()


def _exportados(feito: suporte.ExportacaoRegistro) -> list[dict[str, object]]:
    return [json.loads(x) for x in feito.arquivo.read_text(encoding="utf-8").splitlines()]


def test_exportacao_nao_tem_dado_pessoal(
    exportado: tuple[Path, suporte.ExportacaoRegistro],
) -> None:
    _, feito = exportado
    roteiro = notas_ficticias.carregar()
    with zipfile.ZipFile(feito.zip) as z:
        assert z.namelist() == [feito.arquivo.name]
        conteudo = z.read(feito.arquivo.name).decode("utf-8")
    assert conteudo == feito.arquivo.read_text(encoding="utf-8")
    caixa = conteudo.casefold()
    proibidos = [
        CPF, CPF_MASCARA, CPF[:3] + "." + CPF[3:6], NOME, "zuleika", "quixabeira", "tabajara",
        CONTA, roteiro.conta, "inventada 742", EMAIL, "exemplo.com", "91234", "01310",
        roteiro.titular, "users", str(feito.pasta.parent.parent.parent).casefold(),
        "1.234,56", "1234.56",
    ]  # fmt: skip
    for termo in proibidos:
        assert termo.casefold() not in caixa, "achado um termo proibido"
    for nome, _ in notas_ficticias.textos(roteiro)[:6]:
        assert nome.casefold() not in caixa  # o nome do arquivo também não vai
    for nota in roteiro.notas[:6]:
        assert f'"{nota.numero}"' not in conteudo


def test_exportacao_permite_diagnosticar(
    exportado: tuple[Path, suporte.ExportacaoRegistro],
) -> None:
    """O objetivo do registro: dá para ver o que aconteceu sem nenhum outro dado."""
    _, feito = exportado
    assert feito.retidas == {}  # o uso normal não tem linha retida
    eventos = _exportados(feito)
    nomes = {e["ev"] for e in eventos}
    assert {"sessao.inicio", "importacao.inicio", "importacao.fim", "importacao.arquivo",
            "extracao.pagina", "extracao.nota", "extracao.rejeicao", "extracao.esqueleto",
            "importacao.nota", "log.registro", "suporte.exportacao"} <= nomes  # fmt: skip
    rejeicoes = [_dados_do(e) for e in eventos if e["ev"] == "extracao.rejeicao"]
    assert any(r.get("classe") == "ConciliacaoError" and r.get("campo") for r in rejeicoes)
    assert any(r.get("classe") == "ExtracaoError" for r in rejeicoes)  # a corretora desconhecida
    formas = [str(_dados_do(e)["esqueleto"]) for e in eventos if e["ev"] == "extracao.esqueleto"]
    assert any("CORRETORA" in f and "9.999,99" in f and "999.999.999-99" in f for f in formas)
    fim = next(_dados_do(e) for e in eventos if e["ev"] == "importacao.fim")
    assert fim["ok"] is True and fim["notas_gravadas"] and fim["rejeitadas"]
    assert isinstance(fim["ms"], int)
    assert any(e["ev"] == "apuracao.fim" for e in eventos)
    [cabecalho] = [_dados_do(e) for e in eventos if e["ev"] == "suporte.exportacao"]
    assert cabecalho["versao"] and cabecalho["termos_do_banco"] is True
    assert cabecalho["linhas"] == feito.linhas


def test_apuracao_mostra_grandeza_e_nunca_o_valor(
    exportado: tuple[Path, suporte.ExportacaoRegistro],
) -> None:
    _, feito = exportado
    meses = [_dados_do(e) for e in _exportados(feito) if e["ev"] == "apuracao.mes"]
    assert meses
    for m in meses:
        assert str(m["assinatura"]).startswith("res_") and str(m["titular"]).startswith("tit_")
        for campo in ("imposto", "irrf", "darf", "prejuizo"):
            assert valido(Classe.GRANDEZA, m[campo]), campo


def test_mes_vazio_nao_entra_no_registro(
    exportado: tuple[Path, suporte.ExportacaoRegistro],
) -> None:
    _, feito = exportado
    for e in _exportados(feito):
        if e["ev"] == "apuracao.mes":
            d = _dados_do(e)
            assert {d["resultado_comum"], d["resultado_day_trade"], d["resultado_fii"],
                    d["vendas_acoes"], d["prejuizo"], d["imposto"], d["irrf"]} != {"0"}  # fmt: skip


def test_recalcular_sem_mudanca_nao_repete_os_meses(tmp_path: Path) -> None:
    pasta = tmp_path / "dados"
    pasta.mkdir()
    banco = pasta / NOME_BANCO
    auditoria.configurar(pasta, "9.9.9")
    try:
        roteiro = notas_ficticias.carregar()
        tabelas = pasta_efetiva(pasta)
        refs, fiscais = carregar(tabelas), carregar_fiscais(tabelas)
        migrar(banco, AGORA)
        adicionar_titular(banco, roteiro.cpf, roteiro.titular)
        arquivos = _notas_ficticias(tmp_path / "notas")[:3]
        importar_lista(banco, arquivos, refs, AGORA)
        apurar_banco(banco, refs, fiscais, date(2025, 12, 31))
        primeira = len(_eventos(pasta, "apuracao.mes"))
        apurar_banco(banco, refs, fiscais, date(2025, 12, 31))
        assert primeira > 0 and len(_eventos(pasta, "apuracao.mes")) == primeira
    finally:
        auditoria.desligar()


def test_exportacao_retem_linha_adulterada_e_avisa(dados: Path) -> None:
    hoje = datetime.now().date().isoformat()
    adulterada = _linha("importacao.aviso", texto=f"nota de {NOME}").replace("2026-01-15", hoje)
    with (dados / auditoria.ARQUIVO).open("a", encoding="utf-8") as saida:
        saida.write(f"{adulterada}\n")
        saida.write(f"linha crua com {CPF_MASCARA}\n")
    feito = suporte.exportar_registro(dados, dados / NOME_BANCO, datetime.now())
    assert feito.total_retidas == 2
    conteudo = feito.arquivo.read_text(encoding="utf-8")
    assert CPF_MASCARA not in conteudo and "Quixabeira" not in conteudo
    motivos = [_dados_do(e)["motivo"] for e in _exportados(feito) if e["ev"] == "portao.retidas"]
    assert motivos


def test_exportacao_recusa_quando_a_segunda_passagem_acha(
    dados: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """Defeito simulado na primeira passagem: nada é gravado."""

    def sem_conferir(linhas: list[str], _termos: portao.Termos) -> portao.Conferencia:
        return portao.Conferencia(linhas=[*linhas, _linha("importacao.aviso", texto=CPF_MASCARA)])

    monkeypatch.setattr(portao, "conferir_linhas", sem_conferir)
    with pytest.raises(AnonimizacaoError):
        suporte.exportar_registro(dados, dados / NOME_BANCO, datetime.now())
    assert not (dados / suporte.PASTA_SUPORTE).exists()


def test_periodo_deixa_de_fora_o_antigo(dados: Path) -> None:
    with (dados / auditoria.ARQUIVO).open("a", encoding="utf-8") as saida:
        saida.write(_linha("ui.tela", rota="/antiga").replace("2026-01-15", "2020-01-15") + "\n")
    feito = suporte.exportar_registro(dados, dados / NOME_BANCO, datetime.now(), dias=30)
    assert "/antiga" not in feito.arquivo.read_text(encoding="utf-8")


def test_limite_de_tamanho_guarda_as_mais_recentes(
    dados: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    for i in range(200):
        auditoria.evento("ui.tela", rota=f"/r{i}")
    monkeypatch.setattr(suporte, "MAXIMO_DO_REGISTRO", 2000)
    feito = suporte.exportar_registro(dados, dados / NOME_BANCO, datetime.now())
    conteudo = feito.arquivo.read_text(encoding="utf-8")
    assert feito.cortadas > 0 and '"/r199"' in conteudo and '"/r0"' not in conteudo


def test_conciliacao_nao_mostra_a_quantidade_exata(dados: Path) -> None:
    """A diferença pode ser a posição inteira do cliente naquele ativo."""
    auditoria.evento("conciliacao.divergencia", "A", ativo="PETR4", quantidade=Decimal(73))
    [e] = _eventos(dados, "conciliacao.divergencia")
    assert _dados_do(e)["quantidade"] == "+1e1"


def test_tudo_retido_diz_o_motivo(tmp_path: Path) -> None:
    (tmp_path / "logs").mkdir()
    hoje = datetime.now().date().isoformat()
    linha = _linha("ui.tela", cpf=CPF).replace("2026-01-15", hoje)
    (tmp_path / auditoria.ARQUIVO).write_text(f"{linha}\n", encoding="utf-8")
    with pytest.raises(AnonimizacaoError, match="campo fora do catálogo") as erro:
        suporte.exportar_registro(tmp_path, tmp_path / NOME_BANCO, datetime.now())
    assert CPF not in str(erro.value)


def test_motor_so_levanta_erro_fora_do_registro() -> None:
    """O motor não escreve em disco (regra 6): os erros dele não são registrados ao nascer."""
    permitidos = {"MotorError", "ReferenciaError", "ConciliacaoError", "ExtracaoError"}
    from coir import erros

    dominio = {n for n, c in vars(erros).items() if isinstance(c, type)
               and issubclass(c, erros.CoirError)}  # fmt: skip
    for arquivo in Path("coir/engine").rglob("*.py"):
        for no in ast.walk(ast.parse(arquivo.read_text(encoding="utf-8"))):
            if isinstance(no, ast.Raise) and isinstance(no.exc, ast.Call):
                nome = getattr(no.exc.func, "id", "")
                assert nome not in dominio or nome in permitidos, f"{arquivo}: {nome}"


def test_exportar_sem_registro_explica(tmp_path: Path) -> None:
    with pytest.raises(AnonimizacaoError, match="ainda não há registro"):
        suporte.exportar_registro(tmp_path, tmp_path / NOME_BANCO, AGORA)


def test_portao_funciona_com_o_banco_inacessivel(
    dados: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """O registro é mais útil justamente quando o banco não abre."""
    from coir.erros import BancoError

    def banco_quebrado(_banco: Path) -> list[DadosPessoais]:
        raise BancoError("banco em uso")

    suporte.salvar_titular(dados, DadosPessoais(cpf=CPF, nome=NOME, contas=[CONTA]))
    monkeypatch.setattr(suporte, "do_banco", banco_quebrado)
    termos, do_banco = suporte.termos_do_titular(dados, dados / NOME_BANCO)
    assert do_banco is False and "quixabeira" in termos.palavras
    feito = suporte.exportar_registro(dados, dados / NOME_BANCO, datetime.now())
    assert feito.termos_do_banco is False
    [cabecalho] = [_dados_do(e) for e in _exportados(feito) if e["ev"] == "suporte.exportacao"]
    assert cabecalho["termos_do_banco"] is False


def test_email_nao_leva_a_pasta(exportado: tuple[Path, suporte.ExportacaoRegistro]) -> None:
    pasta, feito = exportado
    endereco = suporte.endereco_email_registro(feito, "9.9.9")
    assert endereco.startswith(f"mailto:{suporte.EMAIL_SUPORTE}?")
    assert feito.zip.name in endereco and str(pasta) not in endereco and NOME not in endereco


def test_segredo_e_privado_nao_vao_no_envio(
    exportado: tuple[Path, suporte.ExportacaoRegistro],
) -> None:
    pasta, feito = exportado
    segredo = (pasta / auditoria.ARQUIVO_SEGREDO).read_text(encoding="utf-8")
    assert segredo not in feito.arquivo.read_text(encoding="utf-8")
    with zipfile.ZipFile(feito.zip) as z:
        assert z.namelist() == [feito.arquivo.name]


# ------------------------------------------------------------------------------------- CLI
def test_cli_exporta_e_registra_o_comando(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv(VAR_DADOS, str(tmp_path))
    monkeypatch.delenv(VAR_SEM_REGISTRO, raising=False)
    try:
        resultado = CliRunner().invoke(app, ["registro", "exportar", "--dias", "1"])
        assert resultado.exit_code == 0, resultado.output
        assert "Para anexar:" in resultado.output and suporte.EMAIL_SUPORTE in resultado.output
        comandos = [_dados_do(e)["comando"] for e in _eventos(tmp_path, "cli.comando")]
        assert comandos == ["registro"]
        assert list((tmp_path / suporte.PASTA_SUPORTE).glob("registro-*/*.zip"))
    finally:
        auditoria.desligar()


def test_cli_recusa_periodo_invalido(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv(VAR_DADOS, str(tmp_path))
    assert CliRunner().invoke(app, ["registro", "exportar", "--dias", "0"]).exit_code != 0


def test_cli_sem_pasta_de_dados_nao_cria_nada(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    faltando = tmp_path / "nao-existe"
    monkeypatch.setenv(VAR_DADOS, str(faltando))
    monkeypatch.delenv(VAR_SEM_REGISTRO, raising=False)
    assert CliRunner().invoke(app, ["versao"]).exit_code == 0
    assert not faltando.exists() and not auditoria.ativa()


def test_testes_da_cli_nao_gravam_na_pasta_da_maquina(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setenv(VAR_DADOS, str(tmp_path))
    assert CliRunner().invoke(app, ["versao"]).exit_code == 0
    assert not auditoria.ativa() and not (tmp_path / "logs").exists()


def test_resultado_vazio_nao_quebra_a_trilha() -> None:
    trilha.extracao("a.pdf", ResultadoExtracao())
    trilha.rejeicao(Rejeitada("a.pdf", (), "x"))
