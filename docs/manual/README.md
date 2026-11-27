# Manual do COIR

O COIR lê as suas notas de corretagem e as planilhas da B3 e calcula, no seu computador, o preço
médio, o imposto de renda mensal, a DARF e as fichas da declaração anual. As imagens deste manual
usam o **banco de demonstração**: um ano (2025) de um investidor fictício, gerado por
`tools/notas_ficticias.py`. Nenhum dado real.

## Comece por aqui
- [Primeiro uso](primeiro-uso.md): baixar, abrir, escolher a pasta de dados, cadastrar o titular e
  importar as primeiras notas.
- [Atualizar para uma versão nova](atualizacao.md): onde ficam os dados, backup e migração.
- [Se o COIR não abrir](se-nao-abrir.md): aviso do Windows, antivírus e o que fazer.

## Telas
O menu à esquerda agrupa as telas. O topo traz o titular e o ano em uso (valem para todas as
telas), o **olho** que esconde os valores, o sino das pendências, o tema claro/escuro e o botão de
recolher o menu. Em cada tela, o botão **Como faço?** abre a página deste manual ao lado, sem cobrir
o que você está preenchendo.

**Esconder os valores (olho do topo):** para mostrar a tela a alguém ou gravar um vídeo, clique no
olho. Os valores em R$ de todas as telas viram `•••••`: cartões, tabelas, totais, avisos e os
números dos gráficos (as barras continuam). Quantidades, preço médio, preços e percentuais
continuam à mostra. Outro clique mostra os valores de novo. A escolha fica gravada: o COIR abre do
jeito que você deixou. É só na tela: o cálculo, os relatórios exportados, o PDF da DARF e o que os
botões de copiar da Declaração copiam continuam com o valor real. Os campos que já vêm com um valor
(valor pago da DARF, outros custos, saldos iniciais) mostram pontos até você clicar neles para
digitar; a nota manual, que você digita inteira, fica à mostra. As tabelas continuam ordenando
pelo valor e pintando de vermelho o negativo, e as mensagens de erro de uma nota que não fecha
mostram os números lidos, para achar o problema.

| Grupo | Tela | Para quê |
|---|---|---|
| | [Painel](telas/painel.md) | resumo do ano: imposto, DARFs, pendências |
| Lançamentos | [Importar](telas/importar.md) | notas em PDF e planilhas da B3 |
| | [Notas](telas/notas.md) | cada nota lida, com o PDF ao lado |
| | [Operações](telas/operacoes.md) | o livro de operações, com filtros |
| | [Simulação](telas/simulacao.md) | o efeito no imposto de vendas e compras antes de operar |
| Imposto | [Apuração mensal](telas/apuracao.md) | o imposto mês a mês, com a memória de cálculo |
| | [DARF](telas/darf.md) | o que pagar, quando, multa e juros; pagamento e comprovante |
| | [Proventos](telas/proventos.md) | dividendos, JCP e rendimentos de FII |
| | [Outros custos](telas/outros-custos.md) | custos dedutíveis fora das notas |
| | [Declaração anual](telas/declaracao.md) | as fichas do programa IRPF, prontas para copiar |
| | [Relatórios](telas/relatorios.md) | relatórios do ano em Excel, CSV e PDF |
| Carteira | [Carteira e preço médio](telas/carteira.md) | posições e a trilha do preço médio |
| | [Eventos](telas/eventos.md) | desdobramentos, bonificações, incorporações |
| | [Opções](telas/opcoes.md) | ciclos de cada série: abertura, exercício, pó |
| | [Futuros](telas/futuros.md) | resultado de futuros por dia e posições |
| | [Conciliação com a B3](telas/conciliacao.md) | a custódia do COIR contra a da B3 |
| Sistema | [Cadastros](telas/cadastros.md) | titulares, contas, ativos, parâmetros e saldos iniciais |
| | [Configurações](telas/configuracoes.md) | pasta de dados, tema, backups, tabelas, registro |
| | [Suporte](telas/suporte.md) | enviar uma nota que o COIR não lê, sem os seus dados |

> O COIR ajuda a calcular e a conferir; ele não substitui a orientação de um contador.
