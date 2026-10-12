"""Testes do hook que bloqueia dados pessoais (CLAUDE.md, regras 9 e 13).

Os CPFs de teste são montados em tempo de execução a partir de uma base fictícia,
para que este arquivo não contenha nenhum número que o próprio hook bloquearia.
"""

import subprocess
from pathlib import Path

import pytest

from tools.verificar_dados_pessoais import (
    carregar_termos,
    cpf_valido,
    main,
    verificar_arquivo,
    verificar_historico,
    verificar_texto,
)


def _com_dv(base9: str) -> str:
    nums = [int(c) for c in base9]
    for pos in (9, 10):
        soma = sum(n * p for n, p in zip(nums, range(pos + 1, 1, -1), strict=False))
        nums.append((soma * 10) % 11 % 10)
    return "".join(map(str, nums))


CPF_FICTICIO = _com_dv("123456789")
CPF_FORMATADO = f"{CPF_FICTICIO[:3]}.{CPF_FICTICIO[3:6]}.{CPF_FICTICIO[6:9]}-{CPF_FICTICIO[9:]}"


@pytest.fixture(autouse=True)
def _sem_termos_reais(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    """Isola os testes da lista de termos sensíveis real do usuário."""
    monkeypatch.setenv("COIR_DADOS", str(tmp_path / "dados-inexistentes"))


def test_dv_conhecido() -> None:
    # base 123456789: DV calculado à mão (soma 210 -> 0; soma 255 -> 9)
    assert CPF_FICTICIO.endswith("09")
    assert cpf_valido(CPF_FICTICIO)


def test_cpf_invalido_e_repetido_nao_validam() -> None:
    assert not cpf_valido(CPF_FICTICIO[:-1] + "8")
    assert not cpf_valido("1" * 11)
    assert not cpf_valido("123")


def test_detecta_cpf_formatado() -> None:
    achados = verificar_texto(f"titular: {CPF_FORMATADO}")
    assert [a.motivo for a in achados] == ["padrão de CPF formatado"]


@pytest.mark.parametrize(
    "texto",
    [
        f"cpf={CPF_FICTICIO}",
        f"nome,{CPF_FICTICIO},XP",  # CSV
        f"CPF {CPF_FICTICIO}.",  # fim de frase
        f"cpf {CPF_FICTICIO[:9]}-{CPF_FICTICIO[9:]}",
        f"cpf {CPF_FICTICIO[:3]} {CPF_FICTICIO[3:6]} {CPF_FICTICIO[6:9]} {CPF_FICTICIO[9:]}",
    ],
)
def test_detecta_cpf_sem_mascara_valido(texto: str) -> None:
    assert [a.motivo for a in verificar_texto(texto)] == ["11 dígitos com DV de CPF válido"]


@pytest.mark.parametrize(
    "texto",
    [
        "exemplo 000.000.000-00",
        "nota 12345678901 e valor 1.234.567,89",  # 11 dígitos sem DV válido
        "conta 85-0 000000000",
        "preço 12.345.678,90",
        f"valor R$ {CPF_FICTICIO},00",  # parte inteira de valor monetário
        f"total 1.{CPF_FICTICIO}",
    ],
)
def test_nao_acusa_exemplos_e_numeros_comuns(texto: str) -> None:
    assert verificar_texto(texto) == []


@pytest.mark.parametrize("separador", [" ", "  ", "\t"])
def test_detecta_conta_de_corretora(separador: str) -> None:
    achados = verificar_texto("Conta: 85-0" + separador + "004201234")
    assert [a.motivo for a in achados] == ["número de conta de corretora"]


def test_marca_de_aceite_ignora_linha() -> None:
    assert verificar_texto(f"{CPF_FORMATADO}  # dados-pessoais: ok") == []


def test_termos_sensiveis_locais(tmp_path: Path) -> None:
    lista = tmp_path / "termos.txt"
    lista.write_text("# comentário\nFulano de Tal\n\n", encoding="utf-8")
    termos = carregar_termos(lista)
    assert termos == ["fulano de tal"]
    achados = verificar_texto("autor: FULANO DE TAL", termos=termos)
    # a mensagem não repete o termo encontrado
    assert [str(a) for a in achados] == ["<texto>:1: termo da lista local de termos sensíveis"]


def test_termos_ausentes_viram_lista_vazia(tmp_path: Path) -> None:
    assert carregar_termos(tmp_path / "nao-existe.txt") == []


def test_bloqueia_extensao_proibida_fora_de_fixtures(tmp_path: Path) -> None:
    pdf = tmp_path / "nota.pdf"
    pdf.write_bytes(b"%PDF-1.4")
    assert [a.motivo for a in verificar_arquivo(pdf)] == ["extensão proibida (.pdf)"]


def test_permite_so_pdf_e_xlsx_em_fixtures() -> None:
    assert verificar_arquivo(Path("tests/fixtures/sintetica.pdf")) == []
    assert [a.motivo for a in verificar_arquivo(Path("tests/fixtures/x.db"))] == [
        "extensão proibida (.db)"
    ]


@pytest.mark.parametrize("nome", ["notas.zip", "coir.db-wal", "coir.bak", "extrato.ofx"])
def test_bloqueia_compactados_e_backups(nome: str) -> None:
    assert len(verificar_arquivo(Path(nome))) == 1


@pytest.mark.parametrize(
    ("conteudo", "motivo"),
    [
        (b"%PDF-1.7\n...", "conteúdo de PDF"),
        (b"PK\x03\x04...", "conteúdo de zip/xlsx"),
        (b"SQLite format 3\x00...", "conteúdo de banco SQLite"),
    ],
)
def test_detecta_conteudo_pela_assinatura(tmp_path: Path, conteudo: bytes, motivo: str) -> None:
    disfarcado = tmp_path / "leia.txt"
    disfarcado.write_bytes(conteudo)
    assert [a.motivo for a in verificar_arquivo(disfarcado)] == [motivo]


def test_binarios(tmp_path: Path) -> None:
    imagem = tmp_path / "tela.JPG"
    imagem.write_bytes(b"\xff\xd8\x00\x01")
    assert verificar_arquivo(imagem) == []
    desconhecido = tmp_path / "dump.bin"
    desconhecido.write_bytes(b"\x01\x00\x02")
    assert [a.motivo for a in verificar_arquivo(desconhecido)] == [
        "binário de tipo não permitido (.bin)"
    ]


def test_uv_lock_nao_e_varrido(tmp_path: Path) -> None:
    lock = tmp_path / "uv.lock"
    lock.write_text(f"hash = {CPF_FICTICIO}\n", encoding="utf-8")
    assert verificar_arquivo(lock, limite_kb=0) == []


def test_bloqueia_arquivo_grande(tmp_path: Path) -> None:
    grande = tmp_path / "grande.txt"
    grande.write_bytes(b"a" * 3 * 1024)
    assert [a.motivo for a in verificar_arquivo(grande, limite_kb=2)] == [
        "arquivo grande (3 KB > 2 KB)"
    ]


def test_le_utf16(tmp_path: Path) -> None:
    arq = tmp_path / "req.txt"
    arq.write_text(CPF_FORMATADO, encoding="utf-16")
    assert len(verificar_arquivo(arq)) == 1


def test_main_codigo_de_saida(tmp_path: Path) -> None:
    limpo = tmp_path / "limpo.py"
    limpo.write_text("x = 1\n", encoding="utf-8")
    sujo = tmp_path / "sujo.py"
    sujo.write_text(f"cpf = '{CPF_FORMATADO}'\n", encoding="utf-8")
    assert main([str(limpo)]) == 0
    assert main([str(limpo), str(sujo)]) == 1
    assert main(["--minimo", "1"]) == 2


def _git(repo: Path, *args: str) -> None:
    subprocess.run(
        ["git", "-c", "user.name=t", "-c", "user.email=t@example.invalid", *args],
        cwd=repo,
        check=True,
        capture_output=True,
    )


def test_historico_acha_dado_removido_depois(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    repo = tmp_path / "repo"
    repo.mkdir()
    _git(repo, "init", "-q")
    (repo / "a.txt").write_text("linha limpa\n", encoding="utf-8")
    _git(repo, "add", ".")
    _git(repo, "commit", "-q", "-m", "base")
    (repo / "a.txt").write_text(f"linha limpa\ncpf {CPF_FORMATADO}\n", encoding="utf-8")
    (repo / "nota.pdf").write_bytes(b"%PDF-1.4")
    _git(repo, "add", ".")
    _git(repo, "commit", "-q", "-m", "vaza")
    (repo / "a.txt").write_text("linha limpa\n", encoding="utf-8")
    _git(repo, "rm", "-q", "nota.pdf")
    _git(repo, "commit", "-q", "-am", "apaga")
    monkeypatch.chdir(repo)
    achados = verificar_historico(["HEAD~2..HEAD"])
    assert sorted((a.arquivo.split(":")[-1], a.linha, a.motivo) for a in achados) == [
        ("a.txt", 2, "padrão de CPF formatado"),
        ("nota.pdf", 0, "extensão proibida (.pdf)"),
    ]
    assert verificar_historico(["HEAD~1..HEAD"]) == []
