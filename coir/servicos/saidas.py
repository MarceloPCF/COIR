"""Ativo na carteira que parou de negociar na bolsa: costuma ser evento que falta (incorporação,
troca, OPA, cisão). Vale para qualquer cliente e para o que acontecer amanhã: a rede pega o caso que
a tabela de trocas ainda não tem, como foi o de PETZ3 -> AUAU3.

Só avisa e, se souber, diz o que houve (``data/saidas_b3.csv``); o custo e a quantidade só mudam
quando o usuário lança o evento, conferido com as notas dele.
"""

from collections.abc import Collection, Iterable
from dataclasses import dataclass
from datetime import date, timedelta
from decimal import Decimal
from pathlib import Path

from coir import referencias
from coir.domain.fiscal import PosicaoAtivo, TipoAtivo

# ativo com pouca liquidez fica semanas sem negócio: abaixo disso não dá para dizer que saiu
DIAS_SEM_NEGOCIO = timedelta(days=90)
TIPOS_VISTA = (TipoAtivo.ACAO, TipoAtivo.UNIT, TipoAtivo.ETF, TipoAtivo.BDR, TipoAtivo.FII)
DESCRICOES = {
    "INCORPORACAO": "incorporada",
    "TROCA_CODIGO": "troca de código",
    "CISAO": "cisão",
    "MIGRACAO_CLASSE": "migração de classe",
    "OPA_CANCELAMENTO": "OPA ou cancelamento de registro",
    "FALENCIA_LIQUIDACAO": "falência ou liquidação",
    "FUNDO_ENCERRADO": "fundo encerrado",
}


@dataclass(frozen=True)
class AtivoQueSaiu:
    ticker: str
    quantidade: Decimal
    ultimo_negocio: date
    texto: str

    @property
    def chave(self) -> str:
        """Dispensa o aviso (``eventos.ignorar``) até o ativo voltar a negociar e parar de novo."""
        return f"saiu|{self.ticker}|{self.ultimo_negocio}"


def _texto(ticker: str, quantidade: Decimal, ultimo: date, saida: dict[str, str] | None) -> str:
    inicio = (f"{ticker}: último negócio na bolsa em {ultimo:%d/%m/%Y}, e ainda há {quantidade} na "
              "carteira.")  # fmt: skip
    descricao = DESCRICOES.get(saida["situacao"]) if saida else None
    if saida is None or descricao is None:
        return (f"{inicio} Veja o que aconteceu com o ativo (incorporação, troca, OPA, "
                "cancelamento) e lance o evento aqui; se ainda tem o ativo, ignore este "
                "aviso.")  # fmt: skip
    partes = [f"{inicio} O que houve: {descricao}"]
    if saida["destino"]:
        relacao = f", {saida['relacao']} por ação" if saida["relacao"] else ""
        partes.append(f", recebendo {saida['destino']}{relacao}")
    if saida["dinheiro_por_acao"]:
        partes.append(f", mais R$ {saida['dinheiro_por_acao']} por ação")
    confianca = (f", levantamento de confiança {saida['confianca']}"
                 if saida["confianca"] != "alta" else "")  # fmt: skip
    partes.append(f". Confira na fonte e no seu extrato e lance o evento{confianca}")
    if saida["fonte"]:
        partes.append(f" (fonte: {saida['fonte']})")
    return "".join(partes) + "."


def ativos_que_sairam(
    carteira: Iterable[PosicaoAtivo], titular: int, pasta: Path, ignoradas: Collection[str] = ()
) -> list[AtivoQueSaiu]:
    """Posições abertas do titular em ativo sem negócio há mais de ``DIAS_SEM_NEGOCIO``, contados
    da última data do pacote de referências (não do relógio: sem pacote novo, nada muda). Ativo
    que a tabela já diz que saiu avisa sem esperar esse prazo."""
    ultimos = referencias.ultimos_negocios(pasta)
    if not ultimos:
        return []
    pacote = max(ultimos.values())
    corte = pacote - DIAS_SEM_NEGOCIO
    conhecidas = referencias.saidas_b3(pasta)
    quantidades: dict[str, Decimal] = {}
    for p in carteira:
        if p.titular == titular and p.quantidade > 0 and p.tipo in TIPOS_VISTA:
            quantidades[p.ticker] = quantidades.get(p.ticker, Decimal(0)) + p.quantidade
    achados = [
        AtivoQueSaiu(t, q, ultimos[t], _texto(t, q, ultimos[t], conhecidas.get(t)))
        for t, q in sorted(quantidades.items())
        if t in ultimos and (ultimos[t] < corte or (ultimos[t] < pacote and t in conhecidas))
    ]
    return [a for a in achados if a.chave not in ignoradas]
