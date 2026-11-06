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

from coir import auditoria
from coir.db.sessao import NOME_BANCO, migrar
from coir.erros import (
    AnonimizacaoError,
    BancoError,
    ConciliacaoError,
    ExtracaoError,
    ImportacaoError,
)
from coir.parsing.anonimizar import DadosPessoais
from coir.parsing.extrair import Rejeitada
from coir.referencias import carregar, carregar_fiscais, pasta_efetiva
from coir.servicos import suporte, trilha
from coir.servicos.apuracao import apurar_banco
from coir.servicos.cadastros import adicionar_titular
from coir.servicos.importacao import importar_lista
from tests.apoio import cpf_de_teste
from tools import notas_ficticias

CPF = cpf_de_teste()
CPF_MASCARA = f"{CPF[:3]}.{CPF[3:6]}.{CPF[6:9]}-{CPF[9:]}"
NOME = "Zuleika Quixabeira Tabajara"
CONTA = "7654321"
ENDERECO = "Travessa Inventada 742"
EMAIL = "zuleika.tabajara@exemplo.com.br"
TELEFONE = "91234-5678"
CEP = "01310-100"
CAMINHO = r"C:\Users\zuleika\Documentos"
VALORES = ("resultado", "vendas", "prejuizo", "imposto", "irrf", "darf")
TERMOS = auditoria.Termos.de([NOME, CPF, CONTA, ENDERECO, TELEFONE], ["tabajara"], "segredo-xyz")


@pytest.fixture
def dados(tmp_path: Path) -> Iterator[Path]:
    pasta = tmp_path / "dados"
    pasta.mkdir()
    auditoria.configurar(pasta, "9.9.9")
    yield pasta
    auditoria.desligar()


def _eventos(pasta: Path, nome: str | None = None) -> list[dict[str, object]]:
    todos = [json.loads(x) for x in (pasta / auditoria.ARQUIVO).read_text("utf-8").splitlines()]
    return [e for e in todos if nome is None or e["ev"] == nome]


def _d(evento: dict[str, object]) -> dict[str, object]:
    conteudo = evento["d"]
    assert isinstance(conteudo, dict)
    return conteudo


def _linha(evento: str, /, **campos: object) -> str:
    return json.dumps({"t": f"{date.today()}T10:30:00.000-03:00", "s": "0a1b2c3d", "n": 1,
                       "nv": "I", "ev": evento, "d": campos}, ensure_ascii=False)  # fmt: skip


def test_toda_chamada_do_codigo_esta_na_lista_fechada() -> None:
    """Senão o campo seria recusado em silêncio no uso real."""
    vistos = 0
    for arquivo in Path("coir").rglob("*.py"):
        for no in ast.walk(ast.parse(arquivo.read_text(encoding="utf-8"))):
            if (isinstance(no, ast.Call) and isinstance(no.func, ast.Attribute)
                    and isinstance(no.func.value, ast.Name) and no.func.value.id == "auditoria"
                    and no.func.attr == "evento"):  # fmt: skip
                assert isinstance(no.args[0], ast.Constant), arquivo
                assert no.args[0].value in auditoria.EVENTOS, arquivo
                assert {k.arg for k in no.keywords if k.arg} <= set(auditoria.CAMPOS), arquivo
                vistos += 1
    assert vistos > 10


def test_valor_vira_ordem_de_grandeza() -> None:
    assert auditoria.grandeza(Decimal("1234.56")) == "+1e3"
    assert auditoria.grandeza(Decimal("-98765.43")) == "-1e4"
    assert auditoria.grandeza(Decimal("0.42")) == "+1e0"
    assert auditoria.grandeza(0) == "0"
    assert auditoria.diferenca(Decimal("0.06")) == "+0.06"
    assert auditoria.diferenca(Decimal("100.01")) == "+1e2"
    for ruim in ("x", True, Decimal("NaN")):
        with pytest.raises(ValueError):
            auditoria.grandeza(ruim)


def test_sem_registro_ligado_nada_acontece(tmp_path: Path) -> None:
    auditoria.evento("ui.tela", rota="/")
    assert not auditoria.ativa() and not auditoria.uma_vez("x") and not list(tmp_path.iterdir())


def test_campo_fora_da_lista_ou_da_forma_nao_entra(dados: Path) -> None:
    auditoria.evento("ui.tela", rota="/", nome_cliente=NOME, cpf=CPF)
    auditoria.evento("importacao.nota", negocios=int(CPF), corretora="XP", data="ontem",
                     ativo="ZULEIKA", imposto="x")  # fmt: skip
    auditoria.evento(
        "importacao.nota", corretora=NOME + " " + CPF_MASCARA, tipo=EMAIL, campo=CAMINHO
    )
    tela, nota, vazia = (*_eventos(dados, "ui.tela"), *_eventos(dados, "importacao.nota"))
    assert _d(tela) == {"rota": "/"} and _d(nota) == {"corretora": "XP"} and _d(vazia) == {}
    bruto = (dados / auditoria.ARQUIVO).read_text("utf-8")
    assert CPF not in bruto and "zuleika" not in bruto.casefold()


def test_pseudonimo_e_estavel_e_nao_mostra_o_valor(dados: Path) -> None:
    for _ in range(2):
        auditoria.evento("importacao.nota", arquivo=f"Nota {NOME}.pdf", conta=CONTA)
    um, dois = _eventos(dados, "importacao.nota")
    assert _d(um) == _d(dois) and str(_d(um)["conta"]).startswith("cta_")
    bruto = (dados / auditoria.ARQUIVO).read_text("utf-8")
    assert "Zuleika" not in bruto and CONTA not in bruto


@pytest.mark.parametrize("gravado", ["", "abc", "zz", "00" * 10])
def test_segredo_vazio_ou_curto_e_refeito(tmp_path: Path, gravado: str) -> None:
    (tmp_path / "_privado").mkdir()
    (tmp_path / auditoria.ARQUIVO_SEGREDO).write_text(gravado, encoding="utf-8")
    auditoria.configurar(tmp_path, "9.9.9")
    chave = auditoria.segredo()
    auditoria.desligar()
    assert len(chave) == 32
    assert (tmp_path / auditoria.ARQUIVO_SEGREDO).read_text("utf-8") == chave.hex()
    auditoria.configurar(tmp_path, "9.9.9")
    assert auditoria.segredo() == chave  # e vale entre sessões


def test_operacao_que_falha_registra_a_pilha_sem_o_texto_e_sobe(dados: Path) -> None:
    @auditoria.operacao("apuracao")
    def falha() -> None:
        raise ValueError(f"falhou em {CAMINHO} para {NOME}")

    with pytest.raises(ValueError, match="falhou"):
        falha()
    inicio, fim = _eventos(dados, "operacao.inicio")[0], _eventos(dados, "operacao.fim")[0]
    assert _d(inicio) == {"nome": "apuracao"} and fim["nv"] == "E"
    d = _d(fim)
    assert d["ok"] is False and d["classe"] == "ValueError" and isinstance(d["ms"], int)
    assert any("test_auditoria" in q for q in d["pilha"])  # type: ignore[attr-defined]
    bruto = (dados / auditoria.ARQUIVO).read_text("utf-8")
    assert "zuleika" not in bruto.casefold() and "Users" not in bruto


def test_erro_do_coir_fica_registrado_ao_ser_criado_sem_o_texto(dados: Path) -> None:
    """A tela só mostra a mensagem: sem isso o erro não ficaria em lugar nenhum."""
    ImportacaoError(f"conta {CONTA} de {NOME}")
    [e] = _eventos(dados, "erro.dominio")
    assert _d(e)["classe"] == "ImportacaoError" and _d(e)["modulo"] == __name__
    assert CONTA not in json.dumps(e) and "Zuleika" not in json.dumps(e)


def test_nota_rejeitada_nao_gera_erro_de_dominio(dados: Path) -> None:
    ExtracaoError("linha fora do padrão")
    ConciliacaoError("valor", Decimal(1), Decimal(2))
    assert _eventos(dados, "erro.dominio") == []


def test_ponte_traz_avisos_e_erros_uma_vez_e_sem_o_texto(dados: Path) -> None:
    log = logging.getLogger("biblioteca.externa")
    log.info("não interessa")
    for _ in range(3):
        log.warning("conta %s", CONTA)  # a mesma linha do código: uma vez só
    try:
        raise KeyError(CPF)
    except KeyError:
        log.exception("quebrou em %s", CAMINHO)
    aviso, erro = _eventos(dados, "log.registro")
    assert aviso["nv"] == "A" and erro["nv"] == "E" and _d(erro)["classe"] == "KeyError"
    bruto = (dados / auditoria.ARQUIVO).read_text("utf-8")
    for termo in (CPF, CONTA, "zuleika"):
        assert termo not in bruto


# ------------------------------------------------------------------------------------- portão
@pytest.mark.parametrize(
    "linha",
    [
        "isto não é JSON",
        "[1, 2, 3]",
        '{"t":"2026"}',
        _linha("evento.inventado", rota="/"),
        _linha("ui.tela", nome_cliente=NOME),
        _linha("ui.tela", texto="qualquer texto"),
        _linha("importacao.nota", arquivo=f"Nota {NOME}.pdf"),
        _linha("importacao.nota", conta=CONTA),
        _linha("importacao.nota", negocios=int(CPF)),
        _linha("importacao.nota", negocios="3"),
        _linha("importacao.nota", data="15/01/2026"),
        _linha("apuracao.mes", imposto="1234.56"),
        _linha("apuracao.mes", darf=1234.56),
        _linha("conciliacao.divergencia", ativo="ZULEIKA"),
        _linha("conciliacao.divergencia", ativo="XP1234567"),
        _linha("conciliacao.divergencia", ativo=f"{CPF}A"),
        _linha("log.registro", pilha=[r"C:\Users\zuleika\x.py:f:1"]),
        _linha("log.registro", classe="Erro de Zuleika"),
        _linha("sessao.inicio", python=CPF_MASCARA),
        _linha("importacao.arquivo", tamanho=True),
        *(_linha("importacao.nota", corretora=v) for v in (CPF_MASCARA, CPF, EMAIL, TELEFONE, CEP,
                                                           CAMINHO, "/home/zuleika")),
    ],
)  # fmt: skip
def test_portao_retem_o_que_esta_fora_da_lista_ou_da_forma(linha: str) -> None:
    assert auditoria.conferir_linha(linha, auditoria.Termos()) is None


@pytest.mark.parametrize("valor", [NOME, "Quixabeira", "TABAJARA", "Inventada"])
def test_portao_retem_termo_do_titular_nos_rotulos(valor: str) -> None:
    linha = _linha("importacao.nota", corretora=valor)
    assert auditoria.conferir_linha(linha, auditoria.Termos()) is not None  # cabe na forma
    assert auditoria.conferir_linha(linha, TERMOS) is None


def test_portao_acha_termo_em_codigo_e_pilha_sem_confundir_a_linha() -> None:
    assert auditoria.conferir_linha(_linha("log.registro", classe="Quixabeira"), TERMOS) is None
    assert (
        auditoria.conferir_linha(_linha("log.registro", pilha=["Quixabeira:f:1"]), TERMOS) is None
    )
    termos = auditoria.Termos.de(["1234"])
    assert auditoria.conferir_linha(_linha("log.registro", pilha=["coir.cli:main:1234"]), termos)


def test_portao_nao_aceita_envelope_com_dado() -> None:
    base = json.loads(_linha("ui.tela", rota="/"))
    for chave, valor in (("t", NOME), ("t", "2026-13-45T10:00:00.000-03:00"), ("s", CPF),
                         ("n", int(CPF)), ("nv", "Z"), ("extra", 1), ("ev", NOME), ("ev", [1]),
                         ("d", NOME)):  # fmt: skip
        assert auditoria.conferir_linha(json.dumps({**base, chave: valor}), TERMOS) is None


def test_portao_aceita_o_que_o_registro_grava_e_refaz_o_texto(dados: Path) -> None:
    auditoria.evento(
        "importacao.nota", arquivo="a.pdf", nota="9", corretora="XP", data=date.today()
    )
    auditoria.evento("apuracao.mes", mes=date(2026, 1, 1), assinatura="x", incompletos=("PETR4",),
                     imposto=Decimal("1234.56"), isento=True)  # fmt: skip
    linhas = (dados / auditoria.ARQUIVO).read_text("utf-8").splitlines()
    assert auditoria.conferir_linhas(linhas, TERMOS) == (linhas, 0)
    assert auditoria.conferir_linhas([linhas[0].replace(",", " , "), "lixo"], TERMOS)[1] == 1


@pytest.mark.parametrize("valor", [CPF, CPF_MASCARA, EMAIL, TELEFONE, CEP, CAMINHO, "segredo-xyz"])
def test_segunda_passagem_acha_o_que_a_primeira_deixou_passar(valor: str) -> None:
    assert auditoria.tem_dado_pessoal(_linha("ui.tela", rota="/", x=valor), TERMOS)


def test_segredo_vazio_nao_vira_termo() -> None:
    termos = auditoria.Termos.de([], segredo="")
    assert not auditoria.tem_dado_pessoal(_linha("ui.tela", rota="/"), termos)


def test_termos_nao_aparecem_no_repr() -> None:
    assert "uixabeira" not in repr(TERMOS) and "tabajara" not in repr(TERMOS)


# --------------------------------------------------------------------------------- do domínio
def test_rejeicao_mostra_campo_e_diferenca(dados: Path) -> None:
    erro = ConciliacaoError("líquido", Decimal("1000.00"), Decimal("1000.06"), "a.pdf: ")
    trilha.rejeicao(Rejeitada(f"Nota {NOME}.pdf", (1,), f"nota 77: {erro}", "77", date(2026, 1, 15),
                              erro))  # fmt: skip
    [e] = _eventos(dados, "importacao.rejeicao")
    d = _d(e)
    assert d["classe"] == "ConciliacaoError" and d["campo"] == "líquido"
    assert d["diferenca"] == "+0.06" and d["data"] == "2026-01-15"
    assert "1000" not in json.dumps(d) and "Zuleika" not in json.dumps(d)


def test_conciliacao_nao_mostra_a_quantidade_exata(dados: Path) -> None:
    """A diferença pode ser a posição inteira do cliente naquele ativo."""
    auditoria.evento("conciliacao.divergencia", "A", ativo="PETR4", quantidade=Decimal(73))
    assert _d(_eventos(dados, "conciliacao.divergencia")[0])["quantidade"] == "+1e1"


def _notas_ficticias(pasta: Path) -> list[Path]:
    """Notas boas, uma que não concilia e uma de corretora desconhecida com dados no nome."""
    pasta.mkdir(parents=True)
    textos = notas_ficticias.textos(notas_ficticias.carregar())
    arquivos = [notas_ficticias.pdf(t, pasta / f"{n}.pdf") for n, t in textos[:6]]
    nome, texto = next((n, x) for n, x in textos if "quido para" in x)
    linhas = texto.splitlines()
    alvo = next(i for i, x in enumerate(linhas) if "quido para" in x)
    linhas[alvo] = re.sub(r"\d,\d\d(?!.*\d,\d\d)", "7,77", linhas[alvo])
    arquivos.append(notas_ficticias.pdf("\n".join(linhas), pasta / f"{nome}-outra.pdf"))
    desconhecida = (f"CORRETORA INVENTADA\nNota 1234 de {NOME}\nCPF {CPF_MASCARA}\n"
                    f"{ENDERECO} {CEP}\n{EMAIL} {TELEFONE}\nConta 85-0 {CONTA}\n"
                    "Total 1.234,56\n")  # fmt: skip
    arquivos.append(notas_ficticias.pdf(desconhecida, pasta / f"Nota de {NOME} {CPF}.pdf"))
    return arquivos


@pytest.fixture
def exportado(tmp_path: Path) -> Iterator[tuple[Path, suporte.ExportacaoRegistro]]:
    pasta = tmp_path / "dados"
    pasta.mkdir()
    banco = pasta / NOME_BANCO
    auditoria.configurar(pasta, "9.9.9")
    roteiro = notas_ficticias.carregar()
    agora = datetime.now()
    tabelas = pasta_efetiva(pasta)
    refs, fiscais = carregar(tabelas), carregar_fiscais(tabelas)
    migrar(banco, agora)
    adicionar_titular(banco, roteiro.cpf, roteiro.titular)
    suporte.salvar_titular(pasta, DadosPessoais(cpf=CPF, nome=NOME, contas=[CONTA]))
    importar_lista(banco, _notas_ficticias(tmp_path / f"notas de {NOME}"), refs, agora)
    apurar_banco(banco, refs, fiscais, date(2025, 12, 31))
    apurar_banco(banco, refs, fiscais, date(2025, 12, 31))
    logging.getLogger("coir.teste").warning("abrindo a nota de %s", NOME)
    yield pasta, suporte.exportar_registro(pasta, banco, agora)
    auditoria.desligar()


def test_exportacao_sem_dado_pessoal_e_que_permite_diagnosticar(
    exportado: tuple[Path, suporte.ExportacaoRegistro],
) -> None:
    pasta, feito = exportado
    roteiro = notas_ficticias.carregar()
    with zipfile.ZipFile(feito.zip) as z:
        assert z.namelist() == [feito.arquivo.name]
        conteudo = z.read(feito.arquivo.name).decode("utf-8")
    assert conteudo == feito.arquivo.read_text("utf-8")
    caixa = conteudo.casefold()
    for termo in (CPF, CPF_MASCARA, "zuleika", "quixabeira", "tabajara", CONTA, roteiro.conta,
                  "inventada", "exemplo.com", "users", roteiro.titular, "1.234,56",
                  (pasta / auditoria.ARQUIVO_SEGREDO).read_text("utf-8")):  # fmt: skip
        assert termo.casefold() not in caixa, "termo proibido na exportação"
    for nome, _ in notas_ficticias.textos(roteiro)[:6]:
        assert nome.casefold() not in caixa  # o nome do arquivo também não vai

    assert feito.retidas == 0  # o uso normal não tem linha retida
    eventos = [json.loads(x) for x in conteudo.splitlines()]
    assert {"sessao.inicio", "operacao.inicio", "operacao.fim", "importacao.arquivo",
            "importacao.nota", "importacao.rejeicao", "importacao.resumo", "apuracao.resumo",
            "apuracao.mes", "log.registro",
            "suporte.exportacao"} <= {e["ev"] for e in eventos}  # fmt: skip
    rejeicoes = [_d(e) for e in eventos if e["ev"] == "importacao.rejeicao"]
    assert any(r.get("classe") == "ConciliacaoError" and r.get("campo") for r in rejeicoes)
    assert any(r.get("classe") == "CorretoraDesconhecida" and r.get("pilha") for r in rejeicoes)
    meses = [_d(e) for e in eventos if e["ev"] == "apuracao.mes"]
    assert meses and len({(m["titular"], m["mes"]) for m in meses}) == len(meses)  # sem repetir
    for m in meses:
        assert {v for c, v in m.items() if c.startswith(VALORES)} != {"0"}
        assert str(m["assinatura"]).startswith("res_")


def test_email_nao_leva_a_pasta(exportado: tuple[Path, suporte.ExportacaoRegistro]) -> None:
    pasta, feito = exportado
    endereco = suporte.endereco_email_registro(feito.zip, "9.9.9")
    assert endereco.startswith(f"mailto:{suporte.EMAIL_SUPORTE}?")
    assert feito.zip.name in endereco and str(pasta) not in endereco


def test_exportacao_retem_linha_adulterada(dados: Path) -> None:
    with (dados / auditoria.ARQUIVO).open("a", encoding="utf-8") as saida:
        saida.write(_linha("importacao.nota", corretora=f"X {CPF_MASCARA}") + "\n")
        saida.write(f"linha crua com {CPF_MASCARA}\n")
    feito = suporte.exportar_registro(dados, dados / NOME_BANCO, datetime.now())
    assert feito.retidas == 2 and CPF_MASCARA not in feito.arquivo.read_text("utf-8")


def test_exportacao_recusa_quando_a_segunda_passagem_acha(
    dados: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """Defeito simulado na primeira passagem: nada é gravado."""

    def sem_conferir(linhas: list[str], _termos: object) -> tuple[list[str], int]:
        return [
            json.dumps(json.loads(_linha("ui.tela", rota=CPF_MASCARA)), separators=(",", ":"))
        ], 0

    monkeypatch.setattr(auditoria, "conferir_linhas", sem_conferir)
    with pytest.raises(AnonimizacaoError):
        suporte.exportar_registro(dados, dados / NOME_BANCO, datetime.now())
    assert not (dados / suporte.PASTA_SUPORTE).exists()


def test_periodo_deixa_de_fora_o_antigo(dados: Path) -> None:
    antiga = json.loads(_linha("ui.tela", rota="/antiga")) | {"t": "2020-01-15T10:00:00.000-03:00"}
    with (dados / auditoria.ARQUIVO).open("a", encoding="utf-8") as saida:
        saida.write(json.dumps(antiga, separators=(",", ":")) + "\n")
    feito = suporte.exportar_registro(dados, dados / NOME_BANCO, datetime.now(), dias=30)
    assert "/antiga" not in feito.arquivo.read_text("utf-8")


def test_exportar_sem_registro_explica(tmp_path: Path) -> None:
    with pytest.raises(AnonimizacaoError, match="ainda não há registro"):
        suporte.exportar_registro(tmp_path, tmp_path / NOME_BANCO, datetime.now())


def test_exporta_com_o_banco_inacessivel(dados: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    """O registro é mais útil justamente quando o banco não abre."""

    def banco_quebrado(_banco: Path) -> list[DadosPessoais]:
        raise BancoError("banco em uso")

    suporte.salvar_titular(dados, DadosPessoais(cpf=CPF, nome=NOME, contas=[CONTA]))
    monkeypatch.setattr(suporte, "do_banco", banco_quebrado)
    feito = suporte.exportar_registro(dados, dados / NOME_BANCO, datetime.now())
    cabecalho = json.loads(feito.arquivo.read_text("utf-8").splitlines()[0])
    assert cabecalho["ev"] == "suporte.exportacao" and _d(cabecalho)["termos_do_banco"] is False


def test_motor_so_levanta_erro_fora_do_registro() -> None:
    """O motor não escreve em disco (regra 6): os erros dele não são registrados ao nascer."""
    from coir import erros

    permitidos = {"MotorError", "ReferenciaError", "ConciliacaoError", "ExtracaoError"}
    dominio = {n for n, c in vars(erros).items() if isinstance(c, type)
               and issubclass(c, erros.CoirError)}  # fmt: skip
    for arquivo in Path("coir/engine").rglob("*.py"):
        for no in ast.walk(ast.parse(arquivo.read_text(encoding="utf-8"))):
            if isinstance(no, ast.Raise) and isinstance(no.exc, ast.Call):
                nome = getattr(no.exc.func, "id", "")
                assert nome not in dominio or nome in permitidos, f"{arquivo}: {nome}"
