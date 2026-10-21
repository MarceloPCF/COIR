"""F8: mapeamento das fichas do programa IRPF por exercício (data/irpf/<exercício>.yaml, R17)."""

import shutil
from pathlib import Path

import pytest
import yaml

from coir import referencias
from coir.domain.fiscal import Grupo
from coir.erros import ConfiguracaoError
from coir.recursos import pasta_data


def _copia(tmp: Path) -> Path:
    pasta = tmp / "data"
    shutil.copytree(pasta_data() / "irpf", pasta / "irpf")
    return pasta


def _troca(pasta: Path, exercicio: int, mudar: object) -> None:
    caminho = pasta / "irpf" / f"{exercicio}.yaml"
    dados = yaml.safe_load(caminho.read_text(encoding="utf-8"))
    mudar(dados)  # type: ignore[operator]
    caminho.write_text(yaml.safe_dump(dados, allow_unicode=True), encoding="utf-8")


def test_exercicios_publicados() -> None:
    m = referencias.carregar_irpf(2025)
    assert (m.exercicio, m.ano_calendario, m.provisorio) == (2026, 2025, False)
    assert m.linha_do_grupo(Grupo.A_VISTA) == "Mercado à Vista - Ações"
    assert m.linha_do_grupo(Grupo.OPCOES_OUTROS) == "Mercado de Opções - Outros"
    assert m.linha_do_grupo(Grupo.TERMO) == "Mercado a Termo - Ações/Ouro"
    assert [x.rotulo for x in m.renda_variavel.linhas][:2] == [
        "Mercado à Vista - Ações", "Mercado à Vista - Ouro"]  # fmt: skip
    assert (m.bens.item("ACAO").grupo, m.bens.item("ACAO").codigo) == ("03", "01")
    assert m.bens.item("OPCAO").codigo == "04" and m.bens.item("JCP_A_RECEBER").grupo == "99"
    assert m.isentos.linha("lucro_acoes_ate_limite").codigo == "20"
    assert referencias.carregar_irpf(2026).provisorio  # programa de 2027 ainda não saiu


def test_exercicio_sem_arquivo_falha_alto(tmp_path: Path) -> None:
    with pytest.raises(ConfiguracaoError, match=r"falta data/irpf/2024.yaml"):
        referencias.carregar_irpf(2023, _copia(tmp_path))


@pytest.mark.parametrize(
    ("mudar", "erro"),
    [
        (lambda d: d["renda_variavel"]["linhas"].pop(0), "cada grupo"),
        (lambda d: d["bens"]["itens"].append(d["bens"]["itens"][0]), "repetido"),
        (lambda d: d.update(exercicio=2030), "ano-calendário"),
        (lambda d: d["isentos"]["linhas"][0].update(extra="x"), "extra"),
    ],
)
def test_mapeamento_invalido_falha_alto(tmp_path: Path, mudar: object, erro: str) -> None:
    pasta = _copia(tmp_path)
    _troca(pasta, 2026, mudar)
    with pytest.raises(ConfiguracaoError, match=erro):
        referencias.carregar_irpf.__wrapped__(2025, pasta)
