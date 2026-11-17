"""Arquivos públicos da bolsa, gerados pelo pipeline (``tools/referencias``): cadastro de
instrumentos e de emissores, nomes de pregão, eventos e proventos."""

from datetime import date
from decimal import Decimal
from pathlib import Path

from coir.domain.cpf import cnpj_valido
from coir.erros import ConfiguracaoError
from coir.fontes.leitura import linhas, mapa_opcional, opcionais
from coir.fontes.modelo import (
    B3,
    Emissor,
    Evento,
    Instrumento,
    NomePregao,
    Provento,
    ProventoHistorico,
    evento_valido,
    partes_da_origem,
)

EVENTOS = "eventos_b3.csv"
INSTRUMENTOS = "instrumentos_b3.csv"
EMISSORES = "emissores_b3.csv"
PROVENTOS = "proventos_b3.csv"
HISTORICO = "proventos_historico_b3.csv"
NOMES_PREGAO = "nomes_pregao_b3.csv"
ATIVOS = "ativos_b3.csv"
OPCOES_RAIZ = "opcoes_raiz_b3.csv"
TIPOS_ATIVO = "tipos_ativo_b3.csv"

COLUNAS = {
    EVENTOS: ("emissor", "isin", "tipo", "fator", "data_com", "aprovado_em", "isin_emitido",
              "preco", "periodo", "data_subscricao"),
    INSTRUMENTOS: ("ticker", "isin", "categoria", "especificacao", "emissor", "nome",
                   "visto_desde", "visto_ate"),
    EMISSORES: ("emissor", "cnpj", "razao_social", "tipo", "visto_desde", "visto_ate"),
    PROVENTOS: ("emissor", "isin", "tipo", "valor", "aprovado_em", "data_com", "pagamento",
                "referente"),
    HISTORICO: ("emissor", "classe", "tipo", "valor", "aprovado_em", "data_com"),
}  # fmt: skip


def eventos(pasta: Path) -> list[Evento]:
    """Eventos utilizáveis que a bolsa lista (origem vazia, em pacote antigo, vale bolsa)."""
    return [Evento.de_linha(r) for r in opcionais(pasta / EVENTOS)
            if B3 in partes_da_origem(r.get("origem", "")) and evento_valido(r)]  # fmt: skip


def _provento_valido(r: dict[str, str]) -> bool:
    try:
        date.fromisoformat(r["aprovado_em"])
        if r["pagamento"]:
            date.fromisoformat(r["pagamento"])
        return Decimal(r["valor"]).is_finite()
    except (KeyError, ValueError, ArithmeticError):
        return False


def proventos(pasta: Path) -> list[Provento]:
    """Proventos em dinheiro com aprovação válida, pagamento válido ou ainda a definir (vazio) e
    valor numérico."""
    return [Provento.de_linha(r) for r in opcionais(pasta / PROVENTOS) if _provento_valido(r)]


def _historico_valido(r: dict[str, str]) -> bool:
    try:
        date.fromisoformat(r["aprovado_em"])
        date.fromisoformat(r["data_com"])
        return bool(r["classe"]) and Decimal(r["valor"]).is_finite()
    except (KeyError, ValueError, ArithmeticError):
        return False


def proventos_historicos(pasta: Path) -> list[ProventoHistorico]:
    """JCP e restituição de capital de anos antigos, com datas válidas e valor numérico."""
    return [ProventoHistorico.de_linha(r) for r in opcionais(pasta / HISTORICO)
            if _historico_valido(r)]  # fmt: skip


def instrumentos(pasta: Path) -> list[Instrumento]:
    """Cadastro de instrumentos; só ticker, ISIN e ``visto_ate`` são obrigatórios na leitura (o
    pacote novo passa antes por ``validar``)."""
    return [
        Instrumento(ticker=r["ticker"], isin=r["isin"], especificacao=r.get("especificacao", ""),
                    emissor=r.get("emissor", ""), nome=r.get("nome", ""),
                    visto_ate=r["visto_ate"], origem=B3)
        for r in opcionais(pasta / INSTRUMENTOS)
    ]  # fmt: skip


def emissores(pasta: Path) -> list[Emissor]:
    """Emissores com CNPJ válido (um por CNPJ visto; o mais recente vale na composição)."""
    return [
        Emissor(emissor=r["emissor"], cnpj=r["cnpj"], razao_social=r["razao_social"],
                tipo=r["tipo"], visto_ate=r["visto_ate"], origem=B3)
        for r in opcionais(pasta / EMISSORES) if cnpj_valido(r["cnpj"])
    ]  # fmt: skip


def ativos(pasta: Path) -> dict[str, str]:
    """Nome de pregão (compacto) -> ticker, do cadastro da bolsa."""
    return mapa_opcional(pasta / ATIVOS, "nome_pregao", "ticker", compacta=True)


def opcoes_raiz(pasta: Path) -> dict[str, str]:
    """Raiz da opção -> ativo-objeto, do cadastro da bolsa."""
    return mapa_opcional(pasta / OPCOES_RAIZ, "raiz", "ticker")


def tipos_ativo(pasta: Path) -> list[tuple[str, str]]:
    """(ticker, tipo) do cadastro da bolsa, na ordem da tabela."""
    return [(t["ticker"], t["tipo"]) for t in opcionais(pasta / TIPOS_ATIVO)]


def nomes_pregao(pasta: Path) -> list[NomePregao]:
    return [NomePregao(nome_pregao=r.get("nome_pregao", ""), ticker=r["ticker"],
                       desde=r.get("desde", ""), ate=r["ate"], origem=B3)
            for r in opcionais(pasta / NOMES_PREGAO)]  # fmt: skip


def validar(pasta: Path) -> None:
    """Pacote novo: as tabelas que só as propostas da tela Eventos e os proventos leem também têm
    de estar íntegras (a linha ruim que a leitura pularia calada falha aqui)."""
    for nome, colunas in COLUNAS.items():
        caminho = pasta / nome
        if not caminho.is_file():
            continue
        tabela = linhas(caminho)
        if tabela and set(colunas) - set(tabela[0]):
            raise ConfiguracaoError(f"{nome}: colunas ausentes")
        if nome == EVENTOS:
            ruins = [r for r in tabela if not evento_valido(r)]
            if ruins:
                raise ConfiguracaoError(f"{nome}: {len(ruins)} linha(s) inválida(s)")
        elif nome == EMISSORES:
            if any(not r["emissor"] or not cnpj_valido(r["cnpj"]) for r in tabela):
                raise ConfiguracaoError(f"{nome}: linha sem emissor ou com CNPJ inválido")
        elif nome == PROVENTOS:
            if not all(_provento_valido(r) for r in tabela):
                raise ConfiguracaoError(f"{nome}: linha com data ou valor inválido")
        elif nome == HISTORICO:
            if not all(_historico_valido(r) for r in tabela):
                raise ConfiguracaoError(f"{nome}: linha com data ou valor inválido")
        elif any(not r["ticker"] or not r["isin"] for r in tabela):
            raise ConfiguracaoError(f"{nome}: linha sem ticker ou ISIN")
