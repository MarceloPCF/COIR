"""O que pode entrar no registro de auditoria: lista de permissão, não de proibição.

Cada campo tem uma classe, que diz como o valor é limpo na gravação e conferido na exportação.
Evento ou campo fora daqui não entra. Para registrar algo novo, acrescente aqui, escolhendo a
classe mais restrita que resolva o diagnóstico (docs/auditoria.md).
"""

from enum import StrEnum


class Classe(StrEnum):
    PSEUDONIMO = "pseudonimo"  # identificador: HMAC com segredo local (arq_3fa9c1...)
    DATA = "data"  # AAAA-MM-DD
    MES = "mes"  # AAAA-MM
    INTEIRO = "inteiro"  # contagens, páginas, milissegundos, bytes (até 10 dígitos)
    INTEIROS = "inteiros"
    LOGICO = "logico"
    GRANDEZA = "grandeza"  # só sinal e ordem de grandeza (+1e4)
    DIFERENCA = "diferenca"  # exata até 100; acima, ordem de grandeza
    ATIVO = "ativo"  # código de negociação (PETR4)
    ATIVOS = "ativos"
    ROTULO = "rotulo"  # texto curto, limpo pelo vocabulário
    TEXTO = "texto"  # texto livre, limpo pelo vocabulário
    TEXTOS = "textos"
    IMPRESSO = "impresso"  # o que a nota imprime como nome de ativo, mercado ou contrato
    CODIGO = "codigo"  # nome vindo do código: classe de erro, módulo, função
    PILHA = "pilha"  # modulo:funcao:linha
    VERSAO = "versao"
    ROTA = "rota"


PREFIXOS = {"arquivo": "arq", "nota": "not", "conta": "cta", "titular": "tit", "assinatura": "res"}


def _classe(classe: Classe, nomes: str) -> dict[str, Classe]:
    return dict.fromkeys(nomes.split(), classe)


CAMPOS: dict[str, Classe] = {
    **dict.fromkeys(PREFIXOS, Classe.PSEUDONIMO),
    **_classe(Classe.DATA, "data desde ate tabelas"),
    **_classe(Classe.MES, "mes"),
    **_classe(
        Classe.INTEIRO,
        "aceitas alertas arquivos avisos contas cortadas dias divergentes eventos folha ignorados "
        "ignoradas isolados ja_no_banco lancadas lidos linha linhas meses ms negocios notas "
        "notas_gravadas notas_ja_no_banco notas_renormalizadas novas novos ocultados operacoes "
        "operacoes_gravadas pagina pendentes rejeitadas retidas tamanho titulares "
        "titulares_criados total",
    ),
    **_classe(Classe.INTEIROS, "paginas meses_fora"),
    **_classe(
        Classe.LOGICO,
        "alterado com_juros com_multa com_pendencia executavel isento ok provisorio sim "
        "termos_do_banco titular_informado",
    ),
    **_classe(
        Classe.GRANDEZA,
        "darf darf_acumulado imposto irrf prejuizo quantidade resultado_comum resultado_day_trade "
        "resultado_fii valor vendas_acoes",
    ),
    **_classe(Classe.DIFERENCA, "diferenca"),
    **_classe(Classe.ATIVO, "ativo ativo_destino"),
    **_classe(Classe.ATIVOS, "incompletos dependentes"),
    **_classe(
        Classe.ROTULO,
        "campo categoria comando corretora layout modo origem segmento situacao tipo",
    ),
    **_classe(Classe.TEXTO, "causa motivo texto"),
    **_classe(Classe.TEXTOS, "esqueleto recusas"),
    **_classe(Classe.IMPRESSO, "lido"),
    **_classe(Classe.CODIGO, "classe funcao modulo"),
    **_classe(Classe.PILHA, "causas pilha"),
    **_classe(Classe.VERSAO, "python revisao sistema versao"),
    **_classe(Classe.ROTA, "rota"),
}


def _campos(nomes: str) -> frozenset[str]:
    return frozenset(nomes.split())


_DO_ERRO = "ok ms classe texto pilha causas data"
# operações: "<nome>.inicio" leva a primeira coluna; "<nome>.fim", a segunda e o resultado
OPERACOES: dict[str, tuple[str, str]] = {
    "importacao": (
        "arquivos titular_informado",
        "lidos ignorados notas_gravadas operacoes_gravadas notas_ja_no_banco "
        "notas_renormalizadas rejeitadas avisos titulares_criados",
    ),
    "planilhas_b3": ("arquivos titular", "novas"),
    "apuracao": ("ate", "meses avisos isolados"),
    "conciliacao": ("titular data", "linhas divergentes avisos"),
    "propostas": ("titular", "total pendentes ignoradas lancadas"),
    "eventos_importar": ("total", "novos ja_no_banco"),
    "anonimizacao": ("arquivos", "aceitas rejeitadas"),
    "exportacao_registro": ("dias", "linhas retidas cortadas"),
    "atualizacao_tabelas": ("", ""),
    "restauracao": ("", ""),
}

EVENTOS: dict[str, frozenset[str]] = {
    "sessao.inicio": _campos("versao python sistema executavel modo"),
    "sessao.banco": _campos("revisao tabelas titulares contas notas operacoes eventos arquivos"),
    "cli.comando": _campos("comando"),
    "ui.tela": _campos("rota"),
    "log.registro": _campos("modulo funcao linha texto classe pilha causas data"),
    "erro.dominio": _campos("classe texto modulo funcao linha"),
    "importacao.arquivo": _campos("arquivo tamanho situacao total notas rejeitadas avisos"),
    "extracao.pagina": _campos("arquivo pagina corretora segmento layout tipo folha nota data"),
    "extracao.nota": _campos("arquivo nota conta corretora segmento data paginas negocios"),
    "extracao.rejeicao": _campos("arquivo paginas nota data classe campo diferenca texto lido"),
    "extracao.esqueleto": _campos("arquivo pagina assinatura esqueleto"),
    "importacao.nota": _campos("arquivo nota situacao operacoes"),
    "importacao.aviso": _campos("texto"),
    "b3.planilha": _campos("arquivo tipo situacao linhas novas desde ate texto"),
    "apuracao.mes": _campos(
        "titular mes assinatura resultado_comum resultado_day_trade resultado_fii vendas_acoes "
        "isento prejuizo imposto irrf darf darf_acumulado provisorio incompletos dependentes"
    ),
    "apuracao.aviso": _campos("texto"),
    "apuracao.isolado": _campos("titular ativo categoria desde meses_fora motivo"),
    "conciliacao.divergencia": _campos("titular data corretora ativo quantidade causa"),
    "conciliacao.aviso": _campos("titular data texto"),
    "eventos.proposta": _campos("titular tipo ativo data origem situacao corretora"),
    "eventos.salvo": _campos("titular tipo ativo ativo_destino data origem alterado"),
    "eventos.excluido": _campos("titular tipo ativo data origem"),
    "eventos.proposta_ignorada": _campos("titular sim"),
    "darf.linha": _campos("titular mes situacao valor com_multa com_juros com_pendencia"),
    "darf.pagamento": _campos("titular mes"),
    "darf.pagamento_excluido": _campos("titular mes"),
    "suporte.nota": _campos("arquivo situacao paginas ocultados alertas recusas"),
    "suporte.exportacao": _campos("versao dias desde ate linhas retidas cortadas termos_do_banco"),
    "portao.retidas": _campos("motivo retidas"),
    **{
        f"{nome}.{fase}": _campos(campos)
        for nome, (inicio, fim) in OPERACOES.items()
        for fase, campos in (("inicio", inicio), ("fim", f"{fim} {_DO_ERRO}"))
    },
}
