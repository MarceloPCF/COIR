"""Registro de auditoria: o que o programa fez, para o suporte corrigir um problema só com o log.

Grava ``<dados>/logs/auditoria.jsonl`` (JSON por linha, rotativo). Nada de dado pessoal entra, nem
no arquivo local: cada evento só aceita os campos do catálogo (``catalogo``), e cada campo é
limpo pela sua classe (``limpeza``): identificador vira pseudônimo, valor vira ordem de grandeza,
texto livre só mantém as palavras do vocabulário do próprio programa. Na exportação, o portão de
privacidade (``portao``) confere tudo de novo, sem confiar no arquivo.

Pode ser usado de qualquer camada menos do motor fiscal, que não registra nada (regra 6): quem
registra é o serviço que o chama, a partir do resultado. Veja docs/auditoria.md.
"""

from coir.auditoria.registro import (
    ARQUIVO,
    ARQUIVO_SEGREDO,
    Operacao,
    RotacaoTolerante,
    arquivos_do_registro,
    ativa,
    configurar,
    dados_do_erro,
    desligar,
    erro_de_dominio,
    evento,
    linha_de_evento,
    operacao,
    pseudonimo,
    segredo,
    uma_vez,
)

__all__ = [
    "ARQUIVO",
    "ARQUIVO_SEGREDO",
    "Operacao",
    "RotacaoTolerante",
    "arquivos_do_registro",
    "ativa",
    "configurar",
    "dados_do_erro",
    "desligar",
    "erro_de_dominio",
    "evento",
    "linha_de_evento",
    "operacao",
    "pseudonimo",
    "segredo",
    "uma_vez",
]
