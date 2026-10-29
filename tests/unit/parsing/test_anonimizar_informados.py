"""Dados informados pelo titular (``DadosPessoais``) e a corretora que nenhum perfil reconhece.
Só dados fictícios."""

import re
from pathlib import Path

from coir.parsing.anonimizar import (
    DadosPessoais,
    MapaAnonimizacao,
    conferir,
    divergencias_de_forma,
    informados_na_copia,
)
from coir.parsing.numeros import compactar
from tests.apoio import cpf_de_teste

PASTA = Path(__file__).parents[2] / "fixtures" / "notas"
CPF = cpf_de_teste("271828182")
CPF_TRACO = f"{CPF[:3]}.{CPF[3:6]}.{CPF[6:9]}-{CPF[9:]}"
CONTA = "7319421"
PESSOAIS = ("JOANA", "PEREIRA", "SANTOS", "ACACIAS", "GOIANIA", "BUENO", "98765", "74210130",
            "62998765432", "51873", "731942", CPF)  # fmt: skip


def _mapa() -> MapaAnonimizacao:
    return MapaAnonimizacao(segredo="00" * 32)


def _sobrou(textos: list[str]) -> list[str]:
    texto = "\n".join(textos)
    achados = []
    for termo in PESSOAIS:
        if termo.isdigit():
            if re.search(r"(?<!\d)0*" + r"[.\-() ]{0,2}".join(termo) + r"(?!\d)", texto):
                achados.append(termo)
        elif compactar(termo) in compactar(texto):
            achados.append(termo)
    return achados


def _desconhecida() -> str:
    """Nota de uma corretora que nenhum perfil reconhece, com a titular fictícia espalhada."""
    return (
        "CORRETORA EXEMPLO S.A.  Fone (11) 3333-4444  atendimento@corretora-exemplo.com.br\n"
        "Nota de corretagem 123   Pregão 05/03/2026\n"
        f"Titular: Joana Pereira dos Santos   CPF {CPF_TRACO}\n"
        "Endereco: R. DAS ACACIAS 98765 APTO 12 - SETOR BUENO\n"
        "CEP 74210-130 GOIANIA GO  Tel (62) 99876-5432  joana.santos@exemplo.com\n"
        "Conta: 731942-1   Assessor 51873\n"
        "C VALE3 VALE ON 100 10,00 1.000,00\n"
        "C ABCD3 7319 10,00 73.194,21\n"
    )


def _joana() -> DadosPessoais:
    return DadosPessoais(
        cpf=CPF, nome="Joana Pereira dos Santos", contas=[CONTA],
        endereco="Rua das Acácias, 98765, apto 12", bairro="Setor Bueno\nVale do Sol",
        cidade="Goiânia", cep="74210-130", telefones="(62) 99876-5432",
        email="joana.santos@exemplo.com", assessor="51873",
    )  # fmt: skip


def test_corretora_desconhecida_sem_dados_informados_e_recusada() -> None:
    c = conferir("nota.pdf", [_desconhecida()], _mapa())
    assert any("CPF impresso na nota não tem dados para ocultar" in r for r in c.recusas)
    outra = _joana().model_copy(update={"nome": "Maria Souza"})
    c = conferir("nota.pdf", [_desconhecida()], _mapa(), informados=[outra])
    assert any("nome informado" in r and "não aparece" in r for r in c.recusas)


def test_corretora_desconhecida_com_dados_informados() -> None:
    texto = _desconhecida()
    c = conferir("nota.pdf", [texto], _mapa(), informados=[_joana()])
    assert c.recusas == [], c.recusas
    [copia] = c.textos
    assert _sobrou(c.textos) == [] and "joana.santos" not in copia
    assert divergencias_de_forma([texto], c.textos) == []
    # o que é da nota fica: negócios, valores, datas (a quantidade 7319 não é a conta)
    assert "C VALE3 VALE ON 100 10,00 1.000,00" in copia
    assert "C ABCD3 7319 10,00 73.194,21" in copia and "05/03/2026" in copia
    # o bairro "Vale do Sol" à mostra na ação vira alerta (contorno na imagem), não troca
    assert [texto[a:b] for a, b in c.alertas[0]] == ["VALE", "VALE"]
    assert c.cpfs == frozenset({CPF}) and any(c.trocados[0])


def test_corretora_desconhecida_troca_formatos_de_contato() -> None:
    c = conferir("nota.pdf", [_desconhecida()], _mapa(), informados=[_joana()])
    [copia] = c.textos
    assert "3333-4444" not in copia and "corretora-exemplo" not in copia  # da corretora também


def test_dado_informado_que_sobra_na_copia_e_recusado() -> None:
    """Nome, CPF, conta e e-mail informados nunca podem sobrar (sem exceção para o texto da
    corretora), mesmo que a troca falhe."""
    achados = informados_na_copia([_desconhecida()], [_joana()])
    assert {a.split(": ", 1)[1] for a in achados} == {
        "o nome informado ficou na cópia", "o CPF informado ficou na cópia",
        "uma conta informada ficou na cópia", "o e-mail informado ficou na cópia",
    }  # fmt: skip
    # o que o anonimizador escreveu não conta: a fictícia pode ser igual a uma conta informada
    texto = _desconhecida()
    i = texto.index("731942-1")
    mascara = bytearray(len(texto))
    mascara[i : i + 8] = b"\x01" * 8
    so_conta = DadosPessoais(cpf="", contas=[CONTA])
    assert informados_na_copia([texto], [so_conta], [bytes(mascara)]) == []
    assert informados_na_copia([texto], [so_conta]) != []


def test_dados_informados_valem_tambem_nas_corretoras_conhecidas() -> None:
    """Um telefone antigo informado, fora do bloco do cliente, some também numa nota da XP."""
    texto = (PASTA / "xp_2026_08_bovespa.txt").read_text(encoding="utf-8")
    extra = texto.replace("Negócios realizados", "Contato (62) 3222-1144\nNegócios realizados", 1)
    assert extra != texto
    c = conferir("nota.pdf", [extra], _mapa(),
                 informados=[DadosPessoais(cpf=CPF, telefones="(62) 3222-1144")])  # fmt: skip
    assert c.recusas == [] and "3222-1144" not in c.textos[0]
    assert c.alertas == [[]]  # alerta só na corretora desconhecida


def test_dados_pessoais_nao_aparecem_no_repr() -> None:
    assert "Joana" not in repr(_joana()) and "Joana" not in str(_joana())
    assert _joana().valores()[:3] == [
        "Rua das Acácias, 98765, apto 12", "Setor Bueno", "Vale do Sol",
    ]  # fmt: skip
