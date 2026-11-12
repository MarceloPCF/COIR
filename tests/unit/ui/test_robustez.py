"""Robustez da interface: valores errados em todos os campos, arquivos fora do padrão em todos os
envios, cliques em todos os botões e troca rápida de telas, titular e ano.

Nada disso pode virar "Erro inesperado" (exceção que chega ao NiceGUI) nem aviso de elemento
apagado no Registro de atividades. Erro de uso tem de virar mensagem na própria tela.

Efeitos fora do programa ficam desligados: abrir pasta ou arquivo no Explorer, navegador, diálogo
de pastas do Windows e acesso à internet (a verificação de versão e de tabelas recebe "sem
conexão").
"""

import asyncio
import io
import os
import random
import threading
import webbrowser
import zipfile
from collections.abc import Iterator
from datetime import datetime
from pathlib import Path
from typing import Any

import pytest
from nicegui import app, background_tasks, run, ui
from nicegui import helpers as nicegui_helpers
from nicegui.elements.upload_files import SmallFileUpload
from nicegui.testing.user import User
from nicegui.testing.user_interaction import UserInteraction
from nicegui.testing.user_simulation import user_simulation

from coir.servicos import atualizacao
from coir.servicos.cadastros import adicionar_titular
from coir.servicos.estado import Estado
from coir.ui import dialogos
from coir.ui.app import INICIO, TELAS, montar
from tests.apoio import cpf_de_teste
from tools import notas_ficticias
from tools.banco_demo import FIXTURES, criar, extrator

pytestmark = pytest.mark.slow

AGORA = datetime(2026, 10, 16, 12, 0, 0)

_CPF = cpf_de_teste()
VALORES_RUINS = (
    "", "   ", "abc", "-1", "0", "-0,01", "1e9", "1e400", "99999999999999999999,99", "NaN",
    "Infinity", "1,2,3", "1.2.3", "31/02/2026", "2026-02-31", "2026-13-01", "9999-12-31",
    "0001-01-01", "1900-01-01", "00/00/0000", "<b>x</b>", "😀", "'; DROP TABLE nota; --",
    "x" * 300, "C:\\nao\\existe\\pasta", "Z:\\", "\\\\servidor\\pasta", "CON", "a/b:c*?",
    # CPF com dígito errado e CPF de teste com máscara (montado aqui: nenhum CPF escrito no código)
    "1" * 11, f"{_CPF[:3]}.{_CPF[3:6]}.{_CPF[6:9]}-{_CPF[9:]}",
)  # fmt: skip

# botões que mexem fora do teste mesmo com os desvios (nenhum hoje) ficam aqui
BOTOES_PULADOS: frozenset[str] = frozenset()


class Registro:
    """O que o usuário veria como erro: exceções e avisos do NiceGUI."""

    def __init__(self) -> None:
        self.erros: list[tuple[str, Exception]] = []
        self.avisos: list[tuple[str, str]] = []
        self.passos: list[str] = []  # cada ação feita na tela (para conferir que houve ação)
        self._onde = ""

    @property
    def onde(self) -> str:
        return self._onde

    @onde.setter
    def onde(self, valor: str) -> None:
        self._onde = valor
        self.passos.append(valor)
        if destino := os.environ.get("COIR_ROBUSTEZ_PASSOS"):  # depuração: o que foi feito
            with open(destino, "a", encoding="utf-8") as f:
                print(valor, file=f)

    def __str__(self) -> str:
        linhas = [f"{o}: {type(e).__name__}: {e}" for o, e in self.erros]
        linhas += [f"{o}: aviso: {m}" for o, m in self.avisos]
        return "\n".join(linhas)


@pytest.fixture
def registro(monkeypatch: pytest.MonkeyPatch) -> Iterator[Registro]:
    r = Registro()
    tratar = app.handle_exception

    def guardar(e: Exception) -> None:
        r.erros.append((r.onde, e))
        tratar(e)

    monkeypatch.setattr(app, "handle_exception", guardar)
    monkeypatch.setattr(nicegui_helpers, "warn_once",
                        lambda msg, **_: r.avisos.append((r.onde, msg)))  # fmt: skip
    yield r


@pytest.fixture
def sem_efeitos_externos(monkeypatch: pytest.MonkeyPatch) -> list[str]:
    abertos: list[str] = []
    if hasattr(os, "startfile"):
        monkeypatch.setattr(os, "startfile", lambda p, *a, **k: abertos.append(str(p)))
    monkeypatch.setattr(webbrowser, "open", lambda url, *a, **k: abertos.append(url) or True)
    monkeypatch.setattr(dialogos, "_powershell", lambda inicial: None)
    # o reload da simulação tira o endereço do histórico e só o devolve ao reabrir: dois reloads
    # seguidos (dois cliques rápidos) davam IndexError no próprio NiceGUI de teste, não no COIR
    from nicegui.testing.user_navigate import UserNavigate

    def recarregar(self: UserNavigate) -> None:
        if self.user.back_history:
            alvo = self.user.back_history[-1]
            self.user.back_history.pop()
            background_tasks.create(self.user.open(alvo, clear_forward_history=False),
                                    name=f"reload {alvo}")  # fmt: skip

    monkeypatch.setattr(UserNavigate, "reload", recarregar)
    return abertos


class Andamento:
    """Conta os cálculos em thread (run.io_bound) para esperar a tela assentar; com ``atraso``,
    cada cálculo demora um pouco (simula a máquina lenta do usuário)."""

    def __init__(self, monkeypatch: pytest.MonkeyPatch, atraso: float = 0.0) -> None:
        self.ativos = 0
        self.trava = threading.Lock()
        self.atraso = atraso
        self.sorteio = random.Random(11)
        original = run.io_bound

        async def contado(f: Any, *args: Any, **kwargs: Any) -> Any:
            espera = self.sorteio.uniform(0, self.atraso) if self.atraso else 0.0
            # a chamada cancelada ainda na fila do executor nunca roda: quem baixa a conta é
            # quem chegar por último (a thread ou a corrotina)
            estado = {"comecou": False, "baixado": False}

            def baixar() -> None:
                with self.trava:
                    if not estado["baixado"]:
                        estado["baixado"] = True
                        self.ativos -= 1

            def chamar() -> Any:
                with self.trava:
                    estado["comecou"] = True
                try:
                    if espera:
                        threading.Event().wait(espera)
                    return f(*args, **kwargs)
                finally:
                    baixar()

            with self.trava:
                self.ativos += 1
            try:
                return await original(chamar)
            finally:
                with self.trava:
                    comecou = estado["comecou"]
                if not comecou:
                    baixar()

        monkeypatch.setattr(run, "io_bound", contado)

    async def assentar(self, folga: float = 0.15) -> None:
        for _ in range(600):
            if self.ativos == 0:
                await asyncio.sleep(folga)
                if self.ativos == 0:
                    return
            await asyncio.sleep(0.05)
        raise AssertionError("cálculos não terminaram em 30 s")


def _sem_rede(url: str) -> bytes:
    raise atualizacao.AtualizacaoError("sem conexão (teste)")


def _estado(banco: Path, falso: bool = True) -> Estado:
    estado = Estado(banco)
    estado.agora = lambda: AGORA
    estado.baixar = _sem_rede
    if falso:
        estado.extrator = extrator(cpf_de_teste())
    return estado


def _visivel(e: ui.element) -> bool:
    while e is not None:
        if not e.visible or e.is_deleted:
            return False
        e = e.parent_slot.parent if e.parent_slot else None  # type: ignore[assignment]
    return True


def _marca(e: ui.element) -> str:
    return ",".join(e._markers) or e.props.get("label", "") or getattr(e, "text", "")


def _da_tela(user: User, tipo: type[ui.element]) -> list[ui.element]:
    """Os elementos do tipo na tela atual, menos o topo e o menu lateral."""
    try:
        achados = user.find(kind=tipo).elements
    except AssertionError:
        return []
    fora = {"alternar-menu", "sino"}
    return [e for e in achados if _visivel(e) and not (set(e._markers) & fora)
            and not _no_topo(e)]  # fmt: skip


def _no_topo(e: ui.element) -> bool:
    while e is not None:
        if isinstance(e, (ui.header, ui.left_drawer)):
            return True
        e = e.parent_slot.parent if e.parent_slot else None  # type: ignore[assignment]
    return False


async def _clicar_tudo(user: User, andamento: Andamento, r: Registro, rota: str,
                       vistos: set[int]) -> None:  # fmt: skip
    for b in _da_tela(user, ui.button):
        if id(b) in vistos or b.is_deleted or not _visivel(b) or not getattr(b, "enabled", True):
            continue
        vistos.add(id(b))
        if set(b._markers) & BOTOES_PULADOS:
            continue
        r.onde = f"{rota} botão {_marca(b)!r}"
        UserInteraction(user, {b}, None).click()
        await andamento.assentar(0.05)


async def _mexer_nas_escolhas(user: User, andamento: Andamento, r: Registro, rota: str) -> None:
    """Cada lista de escolha (select, toggle, rádio) passa por todas as opções; cada caixa de
    marcar e chave liga e desliga; cada item clicável e aba é clicado."""
    for tipo in (ui.select, ui.toggle, ui.radio):
        for s in _da_tela(user, tipo):
            opcoes = s.options  # type: ignore[attr-defined]
            valores = list(opcoes) if isinstance(opcoes, (list, dict)) else []
            for v in valores:
                if s.is_deleted or not _visivel(s):
                    break
                r.onde = f"{rota} escolha {_marca(s)!r} = {v!r}"
                s.set_value(v)  # type: ignore[attr-defined]
                await andamento.assentar(0.02)
    for tipo in (ui.checkbox, ui.switch):
        for c in _da_tela(user, tipo):
            for _ in range(2):
                if c.is_deleted or not _visivel(c):
                    break
                r.onde = f"{rota} marcar {_marca(c)!r}"
                c.set_value(not c.value)  # type: ignore[attr-defined]
                await andamento.assentar(0.02)
    for tipo in (ui.tab, ui.item):
        for it in _da_tela(user, tipo):
            if it.is_deleted or not _visivel(it):
                continue
            r.onde = f"{rota} item {_marca(it) or type(it).__name__!r}"
            UserInteraction(user, {it}, None).click()
            await andamento.assentar(0.02)


async def _preencher(user: User, andamento: Andamento, r: Registro, rota: str, valor: str,
                     todos: bool = False) -> None:  # fmt: skip
    """Põe o valor em cada campo da tela; com ``todos``, passa antes por toda a lista de valores
    errados em cada campo (cada digitação dispara o on_change da tela)."""
    for campo in _da_tela(user, ui.input):
        for v in (*VALORES_RUINS, valor) if todos else (valor,):
            if campo.is_deleted or not _visivel(campo) or not getattr(campo, "enabled", True):
                break
            r.onde = f"{rota} campo {_marca(campo)!r} = {v[:20]!r}"
            campo.set_value(v)  # type: ignore[attr-defined]
            await andamento.assentar(0.01)


ROTAS = [t.rota for t in TELAS] + ["/notas/1", "/notas/999999", "/notas/abc", INICIO.rota]
# telas sem campo nem botão próprios no banco de demonstração (a nota sem PDF e os endereços de
# nota inexistente): só abrem
SEM_CAMPOS_NEM_BOTOES = frozenset({"/notas/1", "/notas/999999", "/notas/abc"})
# telas que têm lista de escolha, aba, item clicável ou caixa de marcar fora do topo
COM_ESCOLHAS = frozenset({
    "/importar", "/notas", "/operacoes", "/declaracao", "/relatorios", "/carteira", "/eventos",
    "/cadastros", "/configuracoes", "/suporte", "/notas/1", "/notas/999999", "/inicio",
})  # fmt: skip
# um de cada família de erro (a lista inteira em todos os botões de todas as telas leva horas)
AMOSTRA = (
    "", "abc", "-1", "1e400", "NaN", "2026-02-31", "9999-12-31", "0001-01-01",
    "😀'; DROP TABLE nota; --", "x" * 300, r"Z:\nao\existe", "CON",
)  # fmt: skip


@pytest.mark.asyncio
@pytest.mark.parametrize("rota", ROTAS)
async def test_valores_errados_em_todos_os_campos_e_cliques_em_todos_os_botoes(
    rota: str, tmp_path: Path, monkeypatch: pytest.MonkeyPatch, registro: Registro,
    sem_efeitos_externos: list[str],
) -> None:  # fmt: skip
    estado = _estado(criar(tmp_path / "dados"))
    adicionar_titular(estado.banco, cpf_de_teste("111444777"), "Segundo Titular")
    andamento = Andamento(monkeypatch)
    async with user_simulation(lambda: montar(estado)) as user:
        for i, valor in enumerate(AMOSTRA):
            registro.onde = f"{rota} abrir"
            await user.open(rota)
            await andamento.assentar(0.05)
            # cada campo com o valor errado (na 1ª volta, com todos os da lista); depois os botões
            await _preencher(user, andamento, registro, rota, valor, todos=i == 0)
            vistos: set[int] = set()
            await _clicar_tudo(user, andamento, registro, f"{rota} [{valor[:12]!r}]", vistos)
            # diálogos e formulários abertos pelos cliques: preenche e clica de novo
            await _preencher(user, andamento, registro, rota, valor)
            await _clicar_tudo(user, andamento, registro, f"{rota} [{valor[:12]!r}]", vistos)
    if rota not in SEM_CAMPOS_NEM_BOTOES:  # a tela foi mexida de fato, não só aberta
        assert len(registro.passos) > 2 * len(AMOSTRA), registro.passos[:5]
    assert not registro.erros and not registro.avisos, str(registro)


# ------------------------------------------------------------------ arquivos fora do padrão


def _xlsx(linhas: list[list[Any]], aba: str = "Movimentação") -> bytes:
    from openpyxl import Workbook

    livro = Workbook()
    folha = livro.active
    assert folha is not None
    folha.title = aba
    for linha in linhas:
        folha.append(linha)
    saida = io.BytesIO()
    livro.save(saida)
    return saida.getvalue()


def _arquivos_ruins(pasta: Path) -> list[tuple[str, bytes]]:
    pasta.mkdir(parents=True, exist_ok=True)
    nota = (FIXTURES / "xp_2026_08_bovespa.txt").read_text(encoding="utf-8")
    pdf_nota = notas_ficticias.pdf(nota, pasta / "nota.pdf").read_bytes()
    pdf_outro = notas_ficticias.pdf("Extrato qualquer\nSem nada de nota", pasta / "o.pdf")
    pdf_branco = notas_ficticias.pdf("", pasta / "b.pdf")
    pdf_senha = notas_ficticias.pdf(nota, pasta / "s.pdf", senha="segredo")
    zip_ = io.BytesIO()
    with zipfile.ZipFile(zip_, "w") as z:
        z.writestr("nota.pdf", pdf_nota)
    from coir.importers.b3 import COLUNAS, TipoPlanilha

    mov = list(COLUNAS[TipoPlanilha.MOVIMENTACAO])
    neg = list(COLUNAS[TipoPlanilha.NEGOCIACAO])
    return [
        ("vazio.pdf", b""),
        ("texto.pdf", "isto não é um PDF".encode()),
        ("truncado.pdf", pdf_nota[:400]),
        ("lixo.pdf", bytes(random.Random(3).randbytes(5000))),
        ("outro-documento.pdf", pdf_outro.read_bytes()),
        ("em-branco.pdf", pdf_branco.read_bytes()),
        ("com-senha.pdf", pdf_senha.read_bytes()),
        ("nota.pdf", pdf_nota),
        ("nota-repetida.pdf", pdf_nota),
        ("NOTA.PDF", pdf_nota),
        ("nome ç ã 😀 ;&$(x).pdf", pdf_nota),
        (("n" * 200) + ".pdf", pdf_nota),
        ("sem-extensao", pdf_nota),
        ("programa.exe", b"MZ\x90\x00"),
        ("notas.zip", zip_.getvalue()),
        ("vazio.xlsx", b""),
        ("lixo.xlsx", bytes(random.Random(5).randbytes(3000))),
        ("pdf-renomeado.xlsx", pdf_nota),
        ("texto.csv", b"a;b;c\n1;2;3\n"),
        ("so-cabecalho.xlsx", _xlsx([mov])),
        ("colunas-erradas.xlsx", _xlsx([["A", "B", "C"], [1, 2, 3]])),
        ("aba-vazia.xlsx", _xlsx([])),
        (
            "movimentacao-ruim.xlsx",
            _xlsx([mov, ["x"] * len(mov), [None] * len(mov), [12345] * len(mov)]),
        ),
        (
            "negociacao-ruim.xlsx",
            _xlsx([neg, ["2026-99-99"] + ["?"] * (len(neg) - 1)], aba="Negociação"),
        ),
        ("grande.pdf", b"%PDF-1.4\n" + b"0" * (12 * 1024 * 1024)),
    ]


ROTAS_COM_ENVIO = ("/importar", "/darf", "/suporte")


@pytest.mark.asyncio
async def test_arquivos_fora_do_padrao_em_todos_os_envios(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, registro: Registro,
    sem_efeitos_externos: list[str],
) -> None:  # fmt: skip
    estado = _estado(criar(tmp_path / "dados"), falso=False)  # extrator real
    andamento = Andamento(monkeypatch)
    ruins = _arquivos_ruins(tmp_path / "ruins")
    async with user_simulation(lambda: montar(estado)) as user:
        for rota in ROTAS_COM_ENVIO:
            await user.open(rota)
            await andamento.assentar()
            for aba in ("aba-notas", "aba-b3", "ver-darf"):
                if rota == "/importar" or aba == "ver-darf":
                    try:
                        user.find(marker=aba).click()
                    except AssertionError:
                        continue
                    await andamento.assentar()
                for envio in _da_tela(user, ui.upload):
                    # um por vez e todos juntos (o envio de vários arquivos de uma vez)
                    for lote in [*([x] for x in ruins), ruins]:
                        registro.onde = f"{rota} {aba} envio {[n for n, _ in lote][:3]}"
                        if envio.is_deleted:
                            break
                        await envio.handle_uploads(  # type: ignore[attr-defined]
                            [SmallFileUpload(n, "application/octet-stream", c) for n, c in lote]
                        )
                        await andamento.assentar(0.3)
    envios = {p.split(" envio ")[0] for p in registro.passos if " envio " in p}
    # as notas e as planilhas (Importar), o comprovante (DARF) e as notas do suporte
    assert len(envios) >= 4, envios
    assert not registro.erros and not registro.avisos, str(registro)


@pytest.mark.asyncio
async def test_pastas_com_arquivos_fora_do_padrao(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, registro: Registro,
    sem_efeitos_externos: list[str],
) -> None:  # fmt: skip
    """Importar a pasta das notas e a das planilhas com lixo dentro (e subpastas)."""
    estado = _estado(criar(tmp_path / "dados"), falso=False)
    andamento = Andamento(monkeypatch)
    ruins = _arquivos_ruins(tmp_path / "ruins")
    estado.preparar(AGORA)
    for pasta in (estado.pasta_notas, estado.pasta_b3):
        pasta.mkdir(parents=True, exist_ok=True)
        (pasta / "subpasta").mkdir(exist_ok=True)
        for nome, conteudo in ruins:
            (pasta / nome).write_bytes(conteudo)
            (pasta / "subpasta" / nome).write_bytes(conteudo)
    async with user_simulation(lambda: montar(estado)) as user:
        await user.open("/importar")
        await andamento.assentar()
        for aba, botao in (("aba-notas", "importar-pasta"), ("aba-b3", "importar-pasta-b3")):
            registro.onde = f"/importar {botao}"
            user.find(marker=aba).click()
            await andamento.assentar()
            user.find(marker=botao).click()
            await andamento.assentar(0.5)
    assert not registro.erros and not registro.avisos, str(registro)


# ------------------------------------------------------------------ troca rápida de telas


@pytest.mark.asyncio
async def test_trocar_de_tela_titular_e_ano_muito_rapido(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, registro: Registro,
    sem_efeitos_externos: list[str],
) -> None:  # fmt: skip
    """Cliques no menu sem esperar a tela calcular, trocando titular e ano no meio; cada cálculo
    demora um tempo sorteado. No fim, a última tela tem de terminar de desenhar."""
    estado = _estado(criar(tmp_path / "dados"))
    adicionar_titular(estado.banco, cpf_de_teste("111444777"), "Segundo Titular")
    andamento = Andamento(monkeypatch, atraso=0.4)
    sorteio = random.Random(7)
    rotas = [t.rota for t in TELAS] + ["/notas/1"]
    async with user_simulation(lambda: montar(estado)) as user:
        await user.open("/")
        for passo in range(120):
            acao = sorteio.random()
            registro.onde = f"passo {passo}"
            if acao < 0.6:
                rota = sorteio.choice(rotas)
                registro.onde += f" abre {rota}"
                ui.navigate.to(rota)  # o clique no menu
            else:
                rotulos = ("Ano", "Titular")
                seletor = [s for s in user.find(kind=ui.select).elements
                           if s.props.get("label") in rotulos and _no_topo(s)]  # fmt: skip
                if seletor:
                    s = sorteio.choice(seletor)
                    opcoes = list(s.options) if isinstance(s.options, (list, dict)) else []  # type: ignore[attr-defined]
                    if opcoes:
                        v = sorteio.choice(opcoes)
                        registro.onde += f" {s.props.get('label')}={v}"
                        s.set_value(v)  # type: ignore[attr-defined]
            await asyncio.sleep(sorteio.uniform(0, 0.08))
        await andamento.assentar(0.5)
    assert not registro.erros and not registro.avisos, str(registro)


# ------------------------------------------------------------------ banco e relatórios


@pytest.mark.parametrize(
    "conteudo", [b"isto nao e um banco", b"", b"SQLite format 3\x00" + b"\xff" * 200]
)
def test_banco_corrompido_vira_mensagem_e_nao_erro_inesperado(
    tmp_path: Path, conteudo: bytes
) -> None:
    """Um coir.db corrompido (ou outro arquivo com esse nome) para a abertura com BancoError, que
    a janela explica ("o banco não abriu", com o caminho dos backups). Antes: DatabaseError do
    SQLAlchemy, e o COIR não abria."""
    from coir.erros import BancoError

    banco = tmp_path / "coir.db"
    banco.write_bytes(conteudo)
    estado = _estado(banco)
    if not conteudo:  # arquivo vazio: o SQLite trata como banco novo e o COIR o cria
        from coir.db.sessao import criar_engine, revisao_atual, revisao_final

        estado.preparar(AGORA)
        engine = criar_engine(banco)
        try:
            assert revisao_atual(engine) == revisao_final()
        finally:
            engine.dispose()
        return
    with pytest.raises(BancoError, match="não é um banco do COIR ou está corrompido"):
        estado.preparar(AGORA)


@pytest.mark.parametrize("visao", ["titular", "corretora", "conta"])
def test_todos_os_relatorios_de_todos_os_anos_em_todos_os_formatos(
    tmp_path: Path, visao: str
) -> None:
    from coir.domain.fiscal import Visao
    from coir.servicos.relatorios import FORMATOS, dados_banco, exportar

    estado = _estado(criar(tmp_path / "dados"))
    estado.preparar(AGORA)
    for ano in [*estado.anos(AGORA.date()), 2007]:
        dados = dados_banco(estado.banco, estado.refs, estado.fiscais, ano, AGORA,
                            estado.titular, visao=Visao(visao))  # fmt: skip
        arquivos = exportar(dados, tmp_path / "saida" / str(ano), FORMATOS)
        assert arquivos and all(a.stat().st_size > 0 for a in arquivos), ano


def test_exportar_com_o_arquivo_aberto_no_excel_vira_mensagem(tmp_path: Path) -> None:
    """O xlsx do ano anterior aberto no Excel (arquivo travado): mensagem para fechar o arquivo,
    não erro inesperado (o xlsxwriter levanta FileCreateError, que não é OSError)."""
    import stat

    from coir.erros import ExportacaoError
    from coir.servicos.relatorios import dados_banco, exportar

    estado = _estado(criar(tmp_path / "dados"))
    estado.preparar(AGORA)
    dados = dados_banco(estado.banco, estado.refs, estado.fiscais, 2026, AGORA, estado.titular)
    destino = tmp_path / "saida"
    [xlsx] = exportar(dados, destino, ("xlsx",))
    os.chmod(xlsx, stat.S_IREAD)  # como o Excel deixa o arquivo aberto
    try:
        with pytest.raises(ExportacaoError, match=f"{xlsx.name} está aberto em outro programa"):
            exportar(dados, destino, ("xlsx",))
    finally:
        os.chmod(xlsx, stat.S_IREAD | stat.S_IWRITE)
    with pytest.raises(ExportacaoError, match="não foi possível usar a pasta"):
        exportar(dados, xlsx, ("csv",))  # a "pasta" é um arquivo


@pytest.mark.asyncio
@pytest.mark.parametrize("rota", ROTAS)
async def test_todas_as_escolhas_abas_e_caixas_de_cada_tela(
    rota: str, tmp_path: Path, monkeypatch: pytest.MonkeyPatch, registro: Registro,
    sem_efeitos_externos: list[str],
) -> None:  # fmt: skip
    estado = _estado(criar(tmp_path / "dados"))
    andamento = Andamento(monkeypatch)
    async with user_simulation(lambda: montar(estado)) as user:
        registro.onde = f"{rota} abrir"
        await user.open(rota)
        await andamento.assentar(0.05)
        await _mexer_nas_escolhas(user, andamento, registro, rota)
        # e de novo, já com as escolhas mexidas (abas e listas mostram outros campos)
        await _mexer_nas_escolhas(user, andamento, registro, rota)
    if rota in COM_ESCOLHAS:  # a tela foi mexida de fato, não só aberta
        assert len(registro.passos) > 3, registro.passos
    assert not registro.erros and not registro.avisos, str(registro)


@pytest.mark.asyncio
async def test_trocar_de_tela_e_de_ano_durante_a_importacao(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, registro: Registro,
    sem_efeitos_externos: list[str],
) -> None:  # fmt: skip
    """A importação continua numa thread enquanto o usuário vai para outras telas e troca o ano;
    quando termina, a tela Importar já não existe. As notas têm de ser gravadas e nada pode
    reclamar de elemento apagado."""
    from coir.servicos.cadastros import adicionar_titular as cadastrar
    from coir.servicos.notas import listar
    from tools.banco_demo import notas as fixtures

    estado = _estado(tmp_path / "dados" / "coir.db")
    estado.preparar(AGORA)
    cadastrar(estado.banco, cpf_de_teste(), "Titular")
    lento = estado.extrator
    assert lento is not None

    def extrair_devagar(caminho: Path) -> Any:
        threading.Event().wait(0.4)
        return lento(caminho)

    estado.extrator = extrair_devagar
    andamento = Andamento(monkeypatch)
    arquivos = fixtures()
    async with user_simulation(lambda: montar(estado)) as user:
        await user.open("/importar")
        await andamento.assentar(0.05)
        [envio] = user.find(marker="enviar-pdfs").elements
        registro.onde = "envio"
        tarefa = asyncio.create_task(envio.handle_uploads(  # type: ignore[attr-defined]
            [SmallFileUpload(f"{a.stem}.pdf", "application/pdf", a.read_bytes())
             for a in arquivos]))  # fmt: skip
        await asyncio.sleep(0.3)
        for i, rota in enumerate(("/", "/carteira", "/importar", "/apuracao", "/eventos", "/")):
            registro.onde = f"durante a importação: {rota}"
            ui.navigate.to(rota)
            await asyncio.sleep(0.25)
            anos = [s for s in user.find(kind=ui.select).elements
                    if s.props.get("label") == "Ano" and _no_topo(s)]  # fmt: skip
            if anos:
                anos[0].set_value(2025 if i % 2 else 2026)  # type: ignore[attr-defined]
        await tarefa
        await andamento.assentar(0.5)
    assert not registro.erros and not registro.avisos, str(registro)
    assert len(listar(estado.banco, None, None)) >= len(arquivos)
