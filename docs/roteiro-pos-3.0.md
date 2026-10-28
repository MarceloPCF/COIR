# O que fica para depois da versão 3.0

Lista do que ficou de fora da 3.0, com a origem de cada item. Ela serve para quem usa saber o que
ainda não existe e para orientar as próximas versões. Ordem dentro de cada grupo: do mais provável
ao menos provável na 3.1.

## Decidido para a 3.1 ou depois

| Item | Situação na 3.0 | Origem |
|---|---|---|
| Anonimizador de notas dentro do programa (texto no formato das fixtures e imagens das páginas com tarjas), para mandar notas de corretoras novas ou de erros | não existe; a nota é enviada à mão, sem dados pessoais | planejamento da F9 |
| Registro de auditoria (uma linha por importação, pagamento, evento e restauração, com tela própria) | o registro de atividades mostra importações, backups, migrações e erros | planejamento da F9 |
| Envio de notas e sugestões de dentro do programa | exige um servidor; na 3.0 os pedidos de corretora vão pela página de pedidos (issues) do GitHub | planejamento da F9 |
| Assinatura digital do `COIR.exe` (o Windows deixa de mostrar o aviso do SmartScreen) | sem assinatura, por decisão: volta à pauta se muitos usuários tiverem problema para abrir. A opção gratuita para projeto de código aberto é a SignPath Foundation; o manual ([Se o COIR não abrir](manual/se-nao-abrir.md)) explica o aviso e como conferir o SHA-256 | planejamento da F9; revisto na 3.0.1 |
| Enviar cada pacote novo à Microsoft como falso positivo (formulário de envio de arquivos do Defender), para o antivírus parar de pôr o programa em quarentena | não é feito; o manual ensina a restaurar o arquivo da quarentena | versão 3.0.1 (27/10/2026) |
| Link para os guias na página inicial do site | os guias só aparecem no rodapé e no manual | versão 3.0.1 (27/10/2026) |
| Troca do nome do programa | continua COIR | decisão do mantenedor |
| Importar a planilha de resultado da v2 (`.xlsb`) como ponto de partida | a v3 refaz tudo a partir dos PDFs (os mesmos da pasta `Entrada` da v2); quem não tem as notas antigas informa a custódia e o prejuízo em 31/12 em Cadastros › Saldos iniciais | rodada do beta (26/10/2026) |
| Vídeo narrado do programa no site | o topo do site mostra as telas em sequência (imagem animada) | rodada do beta (26/10/2026) |
| Formulário de mensagens no site e espaço para publicar elogios e depoimentos | o site só tem o formulário de aviso de versão nova por e-mail | rodada de uso (25/10/2026) |

## Pendências fiscais e do motor

Registradas também em [`regras-fiscais.md`](regras-fiscais.md).

| Item | Situação na 3.0 |
|---|---|
| Rateio dos custos de futuros por contrato e por categoria, configurável | rateio pelo volume (a soma dos custos é a mesma) |
| Day trade entre corretoras antes de 28/07/2010 | ordem dos negócios pelo número da nota (as notas não trazem horário) |
| Lote de ação de PME isento separado do tributado (IN RFB 1.585/2015 art. 69-A §4º) | o cálculo para com erro explícito |
| Venda a termo | o cálculo para com erro explícito (a compra a termo funciona) |
| Feriados depois de 2027 | `data/feriados.csv` vai até 2027 e precisa ser estendido antes de apurar dezembro de 2027 |
| Dividendos de BDR e de ETF, rendimentos pagos em ações | listados com aviso "fora da v3.0" |

## Conferência contra o gabarito

O comando `validar-gabarito` compara as notas lidas com uma planilha externa já conferida. Na
3.0.1 o IRRF passou a ir só para as vendas tributáveis (sem exercício de opção e sem as ações de
PME isentas); sobraram estas diferenças, que não são erros do cálculo:

| Item | Situação na 3.0.1 |
|---|---|
| Isenção de PME julgada pela data da venda, não pelo lote | a ferramenta de conferência não separa o lote isento do tributado (o motor para com erro nesse caso, ver acima) |
| 14 diferenças de IRRF: notas com day trade e operação comum juntas e notas de futuros em que a planilha externa põe o IRRF na compra | a revisar com o mantenedor |
| 160 diferenças de custos em notas de futuros | a investigar; a origem provável é o rateio (ver "Rateio dos custos de futuros" acima) |

## Fora do escopo da versão 3

- Imposto de renda de ativos no exterior.
- Opção flexível (o leitor reconhece a nota e avisa; não calcula).
- Aluguel de ações (BTC) como operação; na 3.0 os custos do aluguel entram como "outros custos" e o
  reembolso de proventos como rendimento isento.
- Renda fixa e Tesouro Direto.
- Carteira: alocação, balanceamento, rentabilidade e manejo de risco (fase futura, F10).
- Senha para abrir o programa.
