# O que fica para depois da versão 3.0

Lista do que ficou de fora da 3.0, com a origem de cada item. Ela serve para quem usa saber o que
ainda não existe e para orientar as próximas versões. Ordem dentro de cada grupo: do mais provável
ao menos provável na 3.1.

## Decidido para a 3.1 ou depois

| Item | Situação na 3.0 | Origem |
|---|---|---|
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

## Anonimizador de notas (fidelidade comprovada)

Feito depois da 3.0.1 (28/10/2026), para sair junto com as outras correções.

**Por que.** As cópias anonimizadas de desenvolvimento foram feitas mexendo no PDF, e isso mudou a
ordem do texto: nas notas BM&F da BTG a conta foi parar na última linha. O leitor foi ajustado às
cópias e errou nas notas reais. A regra agora é anonimizar o texto que o COIR lê, não o PDF.

**Como funciona** (`coir/parsing/anonimizar.py`):

- Cada dado pessoal é trocado no próprio lugar por outro do mesmo formato e comprimento: dígito por
  dígito, letra por letra, caixa e pontuação mantidas. Nenhuma linha muda de lugar.
- CPF vira um CPF fictício com dígitos verificadores válidos. A conta vira a conta fictícia do
  mapa, em todas as ocorrências, com ou sem zeros à esquerda. O nome é trocado onde aparecer,
  inclusive com as palavras coladas do layout antigo da XP.
- Também são trocados o bloco de endereço que segue o nome (rua, bairro, CEP, cidade, telefone,
  e-mail), o código do assessor, o CEP solto sob "Cliente" (BTG BM&F) e banco, agência e conta
  corrente.
- O mapa (segredo e contas/CPFs fictícios) fica em `<dados>/_privado/mapa-anonimizacao.json`,
  fora do repositório. A mesma conta real vira sempre a mesma fictícia.

**Três portões.** Qualquer achado recusa a cópia, com a nota e o campo, nunca o valor:

- Forma: mesmo comprimento e, posição a posição, o mesmo caractere ou outro da mesma classe.
- Fidelidade: `extrair_textos` do original e da cópia dão o mesmo resultado (notas, recusas e
  avisos). Só a conta muda, e só pela do mapa.
- Privacidade: cada termo trocado no original (palavras e números de 4 ou mais caracteres) e os
  termos da lista local (`termos-sensiveis.txt`, a mesma do `tools/verificar_dados_pessoais.py`)
  são procurados na cópia. A parte pública (dados da corretora) só vale de páginas em que o bloco
  do cliente foi achado. Assim, uma página em que o nome não foi localizado não esconde um
  endereço esquecido.

**Saída.** É texto no formato das fixtures (páginas separadas por `\f`, fins de linha `\r\n`
preservados). O leitor aceita esse texto como entrada: `texto_pdf.paginas` lê `.txt`, e
`extrair_pasta`/`arquivos_da_pasta` aceitam `sufixo=".txt"`. Assim os testes de desenvolvimento
rodam sobre o mesmo texto que os PDFs originais dariam. Sem pedido, as pastas continuam lendo só
PDF.

**Corretora desconhecida e dados informados (29/10/2026).** O primeiro desenho só achava os dados
pessoais nos layouts conhecidos: numa corretora nova (o motivo principal do envio) o nome, o
endereço e a conta passavam para a cópia, e o portão aceitava. Agora:

- Cada titular informa, na tela Suporte, os dados para ocultar (`DadosPessoais`: nome, CPF e
  contas do cadastro, mais endereço, bairro, cidade, CEP, telefones, e-mail, assessor, dados
  bancários e outros termos). Ficam em `<dados>/_privado/dados-para-ocultar.json`.
- Os valores inteiros (nome, CPF, conta, e-mail, número de 6 ou mais dígitos, frase de várias
  palavras) são trocados em qualquer lugar e em qualquer layout. A palavra solta ou o número
  curto, só nas linhas pessoais (as que já tiveram troca e, na corretora desconhecida, as três
  depois da do nome). Fora delas, na corretora desconhecida, vira alerta: contorno laranja na
  imagem, para o usuário conferir (o bairro "Vale" e a ação VALE ON).
- Na corretora desconhecida também é trocado o que tem formato de e-mail, CEP ou telefone, e o
  bloco de endereço não é procurado (o layout é desconhecido). A cópia é recusada se o CPF
  impresso não tiver dados informados ou se o nome informado não aparecer na nota.
- Nome, CPF, conta e e-mail informados nunca podem sobrar na cópia, sem a exceção do texto da
  corretora.
- A imagem de cada página sai com tarja preta sobre cada caractere trocado
  (`coir/parsing/imagens.py`). A posição vem do próprio PDF, na mesma ordem do texto que o COIR
  lê; se o texto não casar, a imagem não é gerada e a nota é recusada.

**Onde usar:**

- No programa: tela Suporte. O usuário escolhe os PDFs; o programa grava numa pasta nova em
  `<dados>/suporte/<data-hora>` o texto e as imagens de cada nota aceita, com nomes neutros
  (`nota-01.txt`, `nota-01-pagina-1.png`...), e lista de quem é cada nota (pelo CPF), quantos
  trechos foram ocultados e as recusas. O `.zip` só é gravado depois que o usuário marca que
  conferiu as imagens; o botão "Escrever e-mail" abre o programa de e-mail com a mensagem para
  suporte.coir@gmail.com (só a versão e o nome do `.zip` vão no endereço). O usuário anexa e
  envia. Envio direto pelo programa exige servidor e continua fora.
- Na CLI: `coir notas anonimizar <PDFs ou pastas> [--saida] [--mapa] [--imagens]`, com os dados
  para ocultar da pasta de dados; sai com código 1 se alguma nota for recusada.

**Base de desenvolvimento refeita.**

- Origem: as 273 notas reais (2.034 notas). Destino: `COIR_DADOS/_notas-anonimizadas` (`.txt`).
  Todas passaram nos portões.
- Os testes golden passaram a ler dessa pasta. `tests/golden/test_notas_reais_x_anonimizadas.py`
  compara as reais com a base nova e confere a conta pelo mapa.
- O mapa foi semeado com as contas fictícias que o gabarito já usava, para o gabarito continuar
  fechando sem mudança. A segunda conta da XP (BM&F), que não tinha sido anonimizada nas cópias
  antigas, ganhou uma fictícia nova. Ela foi acrescentada com o mesmo pseudônimo no mapa privado
  do gabarito, e o arquivo anterior ficou como `.antes-anonimizador`.
- A pasta antiga `_notas-corretagens` ficou como estava; nenhum teste a lê mais.

**Pendências deste item:**

- A nota de opção flexível da BTG estava só na base antiga: o original não está entre as notas
  reais. Para voltar a cobrir o aviso de opção flexível com nota real, ponha o PDF original na
  pasta das reais e rode `coir notas anonimizar` de novo.
- As fixtures de `tests/fixtures/notas` vieram das cópias antigas. A ordem do texto delas pode não
  ser a do PDF original (a `btg_2023_11_bmf` certamente não é) e precisa ser refeita a partir da
  base nova.
- ~~`coir.parsing.cliente` devolve "Cliente" como nome do cliente nas notas Bovespa da XP e da
  Rico (e "Corretora" numa BM&F da XP).~~ Corrigido: pula os rótulos soltos depois da ouvidoria
  e, no layout BM&F novo da XP, pega o nome depois da linha do site. Na base nova todas as
  páginas com CPF trazem o nome.
