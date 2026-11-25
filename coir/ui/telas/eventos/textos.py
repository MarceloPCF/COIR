"""Rótulos, ajuda e colunas da tela Eventos (o texto que o usuário lê em cada tipo de evento)."""

from coir.domain.fiscal import NaturezaEvento, TipoEvento
from coir.reports.modelo import Coluna, TipoColuna

T = TipoColuna


TIPOS = {
    TipoEvento.DESDOBRAMENTO: "Desdobramento",
    TipoEvento.GRUPAMENTO: "Grupamento",
    TipoEvento.BONIFICACAO: "Bonificação",
    TipoEvento.SUBSCRICAO: "Subscrição",
    TipoEvento.OFERTA_PUBLICA: "Oferta pública (IPO)",
    TipoEvento.CUSTODIA_INICIAL: "Custódia inicial (31/12)",
    TipoEvento.REDUCAO_CAPITAL: "Redução de capital",
    TipoEvento.TRANSFERENCIA: "Transferência de custódia",
    TipoEvento.INCORPORACAO: "Incorporação / conversão / ajuste de série",
    TipoEvento.CISAO: "Cisão",
    TipoEvento.LEILAO_FRACAO: "Leilão de fração",
    TipoEvento.DIREITO_NAO_EXERCIDO: "Direito não exercido",
    TipoEvento.DIREITO_RECEBIDO: "Direito de subscrição recebido",
    TipoEvento.EXERCICIO_DIREITO: "Subscrição com direito (exercício)",
    TipoEvento.DIVIDENDO_EM_BENS: "Dividendo pago em bens",
}


NATUREZAS = {
    NaturezaEvento.CONVERSAO: "Conversão de classe (ex.: PN em ON)",
    NaturezaEvento.INCORPORACAO_SOCIEDADE: "Incorporação de sociedade",
    NaturezaEvento.INCORPORACAO_ACOES: "Incorporação de ações (art. 252)",
    NaturezaEvento.FUSAO: "Fusão",
    NaturezaEvento.AJUSTE_SERIE: "Ajuste de série de opção pela B3",
    NaturezaEvento.RECIBO_SUBSCRICAO: "Recibo de subscrição que virou ação",
}


AJUDA = {
    TipoEvento.DESDOBRAMENTO: "Informe as ações recebidas. O custo total não muda e o preço médio "
                              "cai.",
    TipoEvento.GRUPAMENTO: "Informe as ações que deixaram de existir. O custo total não muda e o "
                           "preço médio sobe.",
    TipoEvento.BONIFICACAO: "Informe as ações recebidas e o custo atribuído pela companhia, que "
                            "está no aviso aos acionistas.",
    TipoEvento.SUBSCRICAO: "Informe as ações subscritas e o valor pago.",
    TipoEvento.OFERTA_PUBLICA: "Informe as ações compradas na oferta e o valor pago (quantidade "
                               "vezes preço). Entra no preço médio como qualquer compra.",
    TipoEvento.CUSTODIA_INICIAL: "Para quem não tem as notas antigas: a quantidade e o custo em "
                                 "31/12 do Bens e Direitos da última declaração (um por ativo e "
                                 "conta). Opções e futuros em aberto pedem as notas.",
    TipoEvento.REDUCAO_CAPITAL: "Informe o valor restituído pela companhia. Ele reduz o custo, sem "
                                "resultado.",
    TipoEvento.TRANSFERENCIA: "Move a posição para outra conta do titular, com o mesmo preço "
                              "médio. Deixe a quantidade vazia para mover toda a posição.",
    TipoEvento.INCORPORACAO: "As ações de origem saem sem resultado e as recebidas entram com o "
                             "mesmo custo (IN 1.585 art. 58 §6º). Série de opção: herda o "
                             "vencimento.",
    TipoEvento.CISAO: "O percentual do custo divulgado pela companhia vai para o ativo recebido. "
                      "A quantidade de origem é a posição da conta (lance cada conta).",
    TipoEvento.LEILAO_FRACAO: "Venda da fração pelo valor líquido creditado, na data do crédito. "
                              "Entra no limite de isenção do mês.",
    TipoEvento.DIREITO_NAO_EXERCIDO: "Direito, recibo ou bônus de subscrição comprado que venceu "
                                     "sem exercício nem venda: o custo pago vira perda em "
                                     "operações comuns, fora do limite de isenção. Use o fim do "
                                     "prazo de exercício. Sem quantidade, baixa todo o saldo; "
                                     "os recebidos sem custo não dão resultado.",
    TipoEvento.DIREITO_RECEBIDO: "Direitos que a companhia ou o fundo creditou para você, sem "
                                 "custo. Use a data do crédito (a movimentação da B3 traz). O "
                                 "direito comprado em bolsa já vem na nota de corretagem.",
    TipoEvento.EXERCICIO_DIREITO: "Você usou os direitos para subscrever. Informe os direitos "
                                  "usados, o recibo (ou a ação) recebido e o valor pago. O custo "
                                  "dos direitos comprados soma ao valor pago; o exercício não gera "
                                  "imposto. Quando o recibo virar ação, lance a conversão na aba "
                                  "Conversões.",
    TipoEvento.DIVIDENDO_EM_BENS: "A empresa pagou o dividendo com ações ou cotas de outra "
                                  "empresa. Informe a empresa que pagou, o ativo e a quantidade "
                                  "recebidos, o valor atribuído a eles no aviso aos acionistas "
                                  "(é o dividendo, rendimento isento) e o valor de mercado na "
                                  "data da entrega (quantidade vezes a cotação de fechamento). "
                                  "O custo fica pelo menor dos dois. O dividendo em dinheiro já "
                                  "vem na movimentação da B3.",
}  # fmt: skip


ROTULO_DATA = {
    TipoEvento.SUBSCRICAO: "Data da subscrição",
    TipoEvento.OFERTA_PUBLICA: "Data da liquidação",
    TipoEvento.TRANSFERENCIA: "Data da transferência",
    TipoEvento.LEILAO_FRACAO: "Data do crédito",
    TipoEvento.DIREITO_NAO_EXERCIDO: "Fim do prazo de exercício",
    TipoEvento.DIREITO_RECEBIDO: "Data do crédito",
    TipoEvento.EXERCICIO_DIREITO: "Data do exercício",
    TipoEvento.DIVIDENDO_EM_BENS: "Data da entrega",
    TipoEvento.CUSTODIA_INICIAL: "Data",
}


ROTULO_QUANTIDADE = {
    TipoEvento.DESDOBRAMENTO: "Ações recebidas",
    TipoEvento.GRUPAMENTO: "Ações que deixaram de existir",
    TipoEvento.BONIFICACAO: "Ações recebidas",
    TipoEvento.SUBSCRICAO: "Ações subscritas",
    TipoEvento.OFERTA_PUBLICA: "Ações compradas",
    TipoEvento.INCORPORACAO: "Quantidade de origem",
    TipoEvento.CISAO: "Quantidade de origem (posição da conta)",
    TipoEvento.LEILAO_FRACAO: "Fração vendida",
    TipoEvento.DIREITO_NAO_EXERCIDO: "Direitos que venceram (vazio: todo o saldo)",
    TipoEvento.DIREITO_RECEBIDO: "Direitos recebidos",
    TipoEvento.EXERCICIO_DIREITO: "Direitos exercidos",
}


ROTULO_VALOR = {
    TipoEvento.BONIFICACAO: "Custo atribuído",
    TipoEvento.SUBSCRICAO: "Valor pago",
    TipoEvento.OFERTA_PUBLICA: "Valor pago",
    TipoEvento.CUSTODIA_INICIAL: "Custo de aquisição",
    TipoEvento.REDUCAO_CAPITAL: "Valor restituído",
    TipoEvento.LEILAO_FRACAO: "Valor líquido recebido",
    TipoEvento.EXERCICIO_DIREITO: "Valor pago na subscrição",
    TipoEvento.DIVIDENDO_EM_BENS: "Valor atribuído no aviso (total)",
}


ROTULO_ATIVO = {TipoEvento.DIVIDENDO_EM_BENS: "Empresa que pagou (código)"}


# incorporação e cisão: "Ativo recebido" e "Quantidade recebida"
ROTULO_RECEBIDO = {
    TipoEvento.EXERCICIO_DIREITO: ("Recibo (ou ação) recebido", "Quantidade subscrita")
}


# se o evento soma ou tira da posição: só aparece para escolher quando foge do normal (ajuste de
# série de opção vendida, evento antigo lançado ao contrário)
SENTIDO = {
    TipoEvento.DESDOBRAMENTO: "C",
    TipoEvento.GRUPAMENTO: "V",
    TipoEvento.BONIFICACAO: "C",
    TipoEvento.SUBSCRICAO: "C",
    TipoEvento.OFERTA_PUBLICA: "C",
    TipoEvento.CUSTODIA_INICIAL: "C",
    TipoEvento.REDUCAO_CAPITAL: "V",
    TipoEvento.DIREITO_RECEBIDO: "C",
}


SEM_CUSTO = frozenset(
    {TipoEvento.DESDOBRAMENTO, TipoEvento.GRUPAMENTO, TipoEvento.DIREITO_RECEBIDO}
)


COM_OUTRA_CLASSE = frozenset({TipoEvento.BONIFICACAO, TipoEvento.SUBSCRICAO})


COM_DESTINO = frozenset({TipoEvento.INCORPORACAO, TipoEvento.CISAO})


# têm o ativo recebido no formulário; o exercício do direito fica na aba Subscrições, com valor
COM_RECEBIDO = COM_DESTINO | {TipoEvento.EXERCICIO_DIREITO, TipoEvento.DIVIDENDO_EM_BENS}


SEM_QUANTIDADE = frozenset({TipoEvento.DIVIDENDO_EM_BENS})  # só a quantidade recebida


SUBSCRICOES = frozenset(
    {TipoEvento.SUBSCRICAO, TipoEvento.EXERCICIO_DIREITO, TipoEvento.DIREITO_RECEBIDO}
)


NOS_LANCADOS = frozenset(TIPOS) - {TipoEvento.CUSTODIA_INICIAL}  # a custódia fica em Cadastros


COLUNAS_PROPOSTAS = (
    Coluna("Data na B3", T.DATA),
    Coluna("Evento"),
    Coluna("Ativo"),
    Coluna("Corretora"),
    Coluna("Qtd.", T.QUANTIDADE),
    Coluna("Valor", T.DINHEIRO),
    Coluna("Situação"),
)


COLUNAS_SEGURAS = (
    Coluna("Data ex", T.DATA),
    Coluna("Evento"),
    Coluna("Ativo"),
    Coluna("Corretora"),
    Coluna("Ações", T.QUANTIDADE),
    Coluna("Custo total", T.DINHEIRO),
    Coluna("R$ por ação", T.QUANTIDADE),
    Coluna("Fonte do custo"),
)


COLUNAS_COMPRA = (
    Coluna("Data", T.DATA),
    Coluna("Ativo"),
    Coluna("Conta"),
    Coluna("Ações", T.QUANTIDADE),
    Coluna("Valor pago", T.DINHEIRO),
)
