# Histórico de versões

Este projeto segue o [Versionamento Semântico](https://semver.org/lang/pt-BR/) (`MAIOR.MENOR.CORREÇÃO`):

* **MAIOR** – mudança que quebra a compatibilidade (ex.: novo formato da planilha de resultado);
* **MENOR** – funcionalidade nova que continua compatível;
* **CORREÇÃO** – correção de erro, sem funcionalidade nova.

O formato deste arquivo segue o [Keep a Changelog](https://keepachangelog.com/pt-BR/).

## [Não lançado]

### Adicionado
- **Restituição de capital pelo pacote público**: a coleta guarda as restituições de capital em dinheiro do histórico da B3 (antes só JCP), e a tela Eventos as propõe como redução de capital, só com o valor, para quem não importou a movimentação.
- **Aviso de movimentação da B3 ausente**: a tela Eventos avisa quando a movimentação importada não cobre o período das notas (sem ela, subscrições, restituições e ofertas não aparecem).

- **Eventos em abas:** A conferir, Lançados, Subscrições, Transferências, Conversões e Ofertas (IPO), cada uma com um formulário curto e sem o C/V.
- **Lançar as seguras:** lança de uma vez os desdobramentos e grupamentos das tabelas da B3 que não pedem nada do usuário, com lista para conferir antes e opção de desfazer.
- **Oferta pública (IPO e follow-on):** a compra, que não aparece em nota, entra no preço médio, mesmo em ativo ainda sem nota.
- **Termo em aberto na Carteira:** a compra a termo ainda não liquidada aparece numa linha própria, logo abaixo do ativo, com a data da liquidação; não entra na posição, no PM nem no total até liquidar, como na B3.
- **Direito de subscrição não exercido:** o direito, recibo ou bônus comprado que vence sem ser exercido nem vendido vira perda em operações comuns, como a opção que vira pó, na data do fim do prazo de exercício. A tela Eventos propõe o evento "Direito não exercido" a partir da baixa na movimentação da B3, já com a data do prazo de subscrição das tabelas da B3; ele também pode ser lançado à mão. Os direitos recebidos sem custo não dão resultado.
- **Como faço?:** em cada tela, um botão abre a página do manual ao lado, sem cobrir o que você está preenchendo.

### Corrigido

- **Erro no Painel ao abrir**: trocar o titular ou o ano, ou a tela se refazer, enquanto o Painel ainda calculava dava "ValueError: list.remove(x): x not in list" no registro de atividades. O cálculo antigo agora termina sem mexer na tela nova; o mesmo vale para Apuração, Conciliação, Futuros, Proventos e Declaração.
- **Restituição de capital e amortização propostas pela B3**: a proposta vinha com a quantidade da linha da B3, que é a posição, e o evento tirava essas ações da carteira; a venda seguinte virava venda a descoberto. Agora a proposta leva só o valor.
- **Direito não exercido sem quantidade**: a linha "não exercido" da B3 vem sem quantidade; o evento sem ela baixa todo o saldo.
- **Notas da Clear de 2018 a 2020:** eram recusadas com "rótulo 'Data pregão' não encontrado". O COIR agora lê esse layout, com negócios "B3 RV LISTADO" e o resumo em blocos, em qualquer corretora.
- **Fundos imobiliários sem tipo:** mais de 300 FII e Fiagro, como HGBS11 e BRCR11, estavam fora da tabela de tipos porque o nome abreviado não dizia "imobiliário". A tabela agora usa a lista oficial de fundos da B3 e o COTAHIST.
- **Direito, recibo e bônus de subscrição:** ficavam sem tipo, e o ganho não era apurado; alguns direitos de FII (BRCR12, XPHT14) estavam como FII, a 20%. Agora são tributados a 15% (20% em day trade), sem isenção e fora do limite de R$ 20 mil (Soluções de Consulta Cosit 145 e 166/2021). O direito vendido sem nota de compra, recebido da empresa ou do fundo, entra com custo zero.
- **FI-Infra e FIP-IE:** ficavam sem tipo. Agora são reconhecidos pela lista de fundos da B3 (41 FI-Infra e 21 FIP-IE) e têm alíquota zero para pessoa física (Lei 12.431/2011 art. 3º; Lei 11.478/2007 art. 2º): ficam fora da apuração mensal, do limite de R$ 20 mil e do DARF, e o prejuízo não compensa outros ganhos. Na declaração, o ganho vai em Tributação exclusiva (código 06) e as cotas em Bens e Direitos 07.10 (FI-Infra) e 07.07 (FIP-IE). O FIP comum continua isolado, porque paga 15%.
- **Rendimentos de FI-Infra e FIP-IE:** desde que esses fundos ganharam tipo próprio, os rendimentos mensais saíam como "fora do escopo"; antes eram somados aos de FII, em isentos, o que também estava errado. Agora vão para Tributação exclusiva, código 06, pelo valor líquido, com aviso se a B3 mostrar imposto retido. A amortização dessas cotas reduz o custo e é proposta na tela Eventos como redução de capital.
- **Aviso de venda maior que a posição:** só falava em evento corporativo; agora lembra também da nota de compra que falta (de outra corretora ou anterior às notas importadas) e da custódia inicial.
- **Ativo sem tipo não para mais o programa:** ele sai do cálculo, com aviso e os meses marcados como incompletos, e o resto da carteira segue apurado. A venda dele conta no limite de R$ 20 mil como se fosse ação, o lado seguro, e o mês em que isso decide a isenção ganha um aviso.
- **Aluguel de ações (BTC):** a tela Outros custos mandava deduzir a taxa de aluguel, que, segundo a Receita, pessoa física não pode abater (Solução de Consulta Cosit 42/2020). Se você lançou essa taxa, revise: ficam só a corretagem e as tarifas da B3 do empréstimo, no mês da recompra. A tela avisa quando a descrição fala em aluguel.
- **Custo em mês de ganho isento:** reduz primeiro o ganho isento, em vez de virar prejuízo a compensar.
- **Subscrição com posição vendida:** agora dá erro, em vez de zerar a posição e perder o ganho da venda.
- **Evento que entrega outra classe:** subscrição e bonificação creditam a classe emitida, não a ação que você tem.
- **Marcas da coluna Obs.:** qualquer combinação e ordem é aceita (`D@`, `D#@`, `@D`...), e vale também a legenda impressa na própria nota. Antes, `D@` e `D#@` travavam a nota.
- **Day trade perdido em silêncio:** em notas com o código do ativo na linha, `PETR4 D@` entrava como operação comum. Quem importou notas desse tipo deve usar "Reprocessar notas já importadas".
- **Day trade parcial:** 500 D x 200 D no mesmo dia são 200 de day trade e 300 comuns, sem aviso (IN RFB 1.585/2015 art. 65).
- **Lote de cotação diferente:** RAIZ55 cota por 100; o valor de quantidade x preço agora aceita fator 1, 10, 100 e 1.000.
- **Coleta das tabelas públicas:** feriado não trava mais a coleta, a falha abre uma issue e o dia seguinte tenta de novo.
- **Mensagem de ativo não encontrado:** diz o que fazer e só sugere nome muito parecido.
- **Tela Importar:** arquivo com parte das notas recusada mostra "N notas gravadas; outras recusadas".

- **Propostas de eventos sumiam:** um ativo que o motor não consegue apurar parava a busca, e os eventos seguintes de toda a carteira deixavam de ser propostos. Agora só as propostas daquele ativo esperam.
- **Ações de PME na Apuração mensal:** o resultado isento (e a perda) agora aparece na coluna própria e na memória de cálculo do mês.

- **Tabela de ações de PME incompleta:** só tinha PRIO3 e AGRO3. Entraram SQIA3 (compras até 05/09/2019), CRDE3, FRTA3, GSHP3, NUTR3 e PRNR3, e a data de corte de AGRO3 passou a 04/02/2021. Vendas dessas ações até 31/12/2023 podem mudar de tributadas para isentas: reprocesse a apuração e confira.
- **Ações de PME com lote isento e tributado:** o COIR agora controla os dois lotes em separado (IN RFB 1.585/2015 art. 69-A §4º), e a custódia inicial e os eventos entram no lote certo. Antes, o ativo saía do cálculo.
- **Ajuste de custo do fim da isenção de PME:** em 01/01/2024, o lote isento sobe à média dos 30 pregões, quando maior que o custo pago (Lei 13.043/2014 art. 17 §6º). Muda o preço médio e o imposto de vendas depois de 2023.

- **Custódia inicial com código antigo:** CNTO3 lançado em 31/12/2019 criava um ativo separado do SBFG3 das notas. Agora a custódia segue a tabela de trocas. Quem já lançou deve apagar a linha e lançar de novo.

### Alterado

- **Eventos entram sozinhos:** desdobramento, grupamento e bonificação com custo publicado são lançados automaticamente para as ações em carteira na data, ao importar notas e ao abrir a tela Eventos. O que você apagar não volta. Subscrição, bonificação sem custo, cisão e incorporação continuam na conferência.
- **Eventos depois da última nota** não aparecem como pendência: o COIR não sabe se você ainda tinha o ativo. Com as notas seguintes, eles passam a valer.
- **A janela abre maximizada**, em vez do tamanho fixo que não cabia nas telas menores.
- **Centavos de arredondamento não recusam mais a nota.** A diferença entre os valores impressos vira a linha de custo "Ajuste de arredondamento", e a nota fecha com o líquido impresso. Só o que passar do que o arredondamento explica consome a tolerância de R$ 0,05 por nota, com aviso; acima dela, a nota é recusada.

### Adicionado

- **Histórico de nomes de pregão da B3 desde 2008** (`nomes_pregao_b3.csv`), pela data do negócio, e todos os ETFs de renda fixa listados.
- **RAIZ4 -> RAIZ55, PETZ3 -> AUAU3** (com a repartição do custo do resgate em dinheiro, em Eventos), trocas de nome confirmadas pelo histórico (IMCH3, OGSA3, BRIT3, NINJ3, REAG3, FICT3) e tipo do NFTS11.
- **Aviso de ativo que parou de negociar:** posição em ativo sem negócio há mais de 90 dias aparece no sino e na tela Eventos, com o que houve, quando se sabe.
- A coleta semanal lista no PR todo ativo que saiu do cadastro, para revisão.
- **Operações por ativo, mês a mês:** em Operações finalizadas (também nos arquivos exportados) e na Apuração mensal, uma linha por dia e ativo, com custo médio, preço final, lucratividade e resultado, em blocos por mês e mercado. Clicar numa linha abre as operações dela.
- **Reprocessar notas já importadas**: marque a opção na tela Importar (ou use
  `coir importar --reprocessar`) para ler os PDFs de novo e ficar com a releitura, quando o COIR
  leu algo errado ou a corretora corrigiu a nota. Há backup antes, a lista diz o que mudou e a nota
  que a releitura não confirma continua como está. A versão anterior fica guardada e aparece na
  tela Notas, e a lista avisa quando há DARF paga do mês da nota em diante.
- **Corretora ainda sem perfil**: o COIR tenta ler a nota pelos layouts que já conhece. Se os totais
  da nota fecharem, ela entra com um aviso para você conferir; se não, ela é recusada dizendo o que
  não deu certo, e nada dela é gravado.
- **Registro de auditoria para o suporte**: o COIR registra em `logs/auditoria.jsonl` o que faz
  (importações, notas rejeitadas, conciliação com a B3, propostas de eventos, apuração por mês, DARF
  e erros). Nem no seu computador ele tem CPF, nome, conta, endereço, e-mail, nomes de arquivos,
  textos ou valores exatos; mostra os códigos dos ativos e as datas dos problemas. Tela **Suporte >
  Enviar o registro ao suporte**: grava o arquivo para você conferir e o `.zip`, depois de uma
  conferência de privacidade, e abre o e-mail; você anexa e envia.

### Corrigido
- **PDF com senha**: a mensagem agora também ensina o caminho da linha de comando,
  `coir importar --senha`, além do campo "Senha dos PDFs" da tela Importar.
- **Corretora não identificada**: a mensagem mostra o código da B3 também quando ele vem na mesma
  linha do rótulo ("Número da Corretora 88-4 Cliente"), como na CM Capital.
- **Avisos da apuração por titular**: o sino, o Painel e a tela Apuração mostram só os avisos do
  titular escolhido. Os arquivos com nota rejeitada continuam contados para todos, porque um
  arquivo rejeitado pode não ter titular conhecido.

## [3.1.0] - 2026-11-04

### Adicionado
- **Tela Suporte**, para mandar uma nota que o COIR não lê ou lê errado (uma corretora nova,
  por exemplo) sem os seus dados pessoais:
  - **Dados para ocultar**, por titular: nome, CPF e contas vêm do cadastro; você completa com
    endereço, telefones, e-mail, assessor e dados bancários, como saem nas notas. Ficam só no seu
    computador. O COIR procura esses dados em qualquer corretora.
  - **Enviar notas para suporte**: o COIR grava numa pasta o texto de cada nota, com os dados
    pessoais trocados por fictícios do mesmo tamanho, e a imagem de cada página com tarja preta
    sobre eles, mostrando o progresso nota a nota. A cópia só sai se o COIR ler nela o mesmo que
    no original e não sobrar nenhum dado pessoal. Você confere as imagens, grava o `.zip` e o
    botão abre o seu e-mail com a mensagem para o suporte; você anexa e envia.
  - Pela linha de comando: `coir notas anonimizar [--imagens]`.

- **Um evento que falta não trava mais a apuração**: quando o cálculo de um ativo não fecha (em
  geral por um desdobramento, bonificação ou ajuste de série que não foi lançado), só esse ativo
  sai do cálculo a partir do dia do problema; os demais continuam. O Painel mostra quantos
  ativos estão fora, a Apuração diz qual, desde quando e por quê, e a DARF dos meses afetados sai
  como **provisória**, inclusive a dos meses seguintes cujo prejuízo a compensar depende dele. As
  vendas do ativo continuam contando no limite de isenção de R$ 20 mil. Lance o evento na tela
  Eventos e o ativo volta ao cálculo.
- **Aviso de evento possivelmente faltando**: venda maior que a posição comprada que deixa o
  ativo vendido no fim do mês gera um aviso na Apuração (não muda o cálculo).
- **Notas em PDF com senha** (como as enviadas por e-mail): o COIR tenta sozinho o começo do CPF
  dos titulares cadastrados. Para a primeira importação, ou se a senha for outra, informe-a no
  campo **Senha dos PDFs** da tela Importar (vale para todo o lote e fica só na memória até
  fechar o programa; nunca é gravada) ou com `coir importar --senha`. A tela Notas mostra a
  página do PDF protegido, e a tela Suporte também o aceita. Sem senha que sirva, a nota é
  recusada com "PDF protegido por senha", não mais "PDF ilegível".
- **Sino de pendências no topo da janela**, com a quantidade e um atalho para cada uma: DARF em
  atraso, nota rejeitada, ativo fora do cálculo, avisos da apuração e a novidade **propostas de
  evento para conferir**. Vale para todos os anos, não só o escolhido: o atalho abre a tela já no
  ano do item (a DARF em atraso mais antiga, o ano mais recente com aviso). O cartão Pendências do
  Painel mostra a mesma lista.

### Corrigido
- **Notas da Clear de 2026 recusadas** com "página sem perfil de corretora conhecido": o
  cabeçalho traz "CLEAR CTVM S/A" em vez de "Clear Corretora". O COIR agora reconhece também
  esse nome e o CNPJ da Clear.
- **Corretora identificada também pelo código da B3**: sem nome nem CNPJ conhecidos no
  cabeçalho, vale o código da corretora impresso antes da conta do cliente (3 = XP, 85 = BTG,
  308 = Clear, 386 = Rico). Se nada bater, a mensagem diz o código lido e o que foi tentado.
- **Evento lançado na conta errada era aceito** quando já havia outra pendência no livro (por
  exemplo, uma bonificação que faltava em outro ativo): a conferência parava na visão por
  titular e não olhava a visão por conta. Com isso a tela Eventos deixava de propor os eventos
  seguintes. Agora vale a pendência mais antiga das duas visões.
- **Nota fotografada ou escaneada**: a mensagem agora diz que o PDF é só imagem e pede o PDF
  baixado do site da corretora (o COIR lê o texto da nota, não faz OCR).
- **Bonificação proposta com custo zero**: quando o pacote de eventos traz o custo atribuído pela
  companhia (aviso aos acionistas), a proposta da tela Eventos já vem com o custo total. Antes vinha
  zerada e, se o usuário não preenchesse, o preço médio das vendas seguintes ficava errado.
- **Motivo de um ativo fora do cálculo**: uma série de opção que saiu do cálculo junto com o
  ativo-objeto mostrava o motivo de outro ativo do titular; agora mostra o do ativo que a puxou.
- **Ícone da janela**: a barra de título e a barra de tarefas mostram o símbolo do COIR.
- **Nome do titular nas notas da XP e da Rico**: nas notas Bovespa o COIR lia o rótulo
  "Cliente" como nome do cliente (e "Corretora" nas BM&F novas da XP); o titular criado na
  importação ficava com esse nome. Agora lê o nome impresso.
- **Conta das notas BM&F da BTG**: o COIR lia como conta o último número da página, que nas
  notas BTG/Necton pode ser o telefone da ouvidoria. Com isso criava contas que não existem. Nas
  notas de 2025 em diante a última linha é o endereço do banco, e a nota era rejeitada ("conta
  do cliente não encontrada"). Agora a conta é lida depois de "Código cliente". Se em
  Cadastros aparecer uma conta BTG estranha (4 dígitos, como o fim de um telefone), importe as
  notas de novo numa pasta de dados nova. As notas já gravadas não são relidas.
- **Telas travadas por um evento que faltava**: exercício de opção com a quantidade ajustada por
  uma bonificação não lançada (ex.: 1.000 opções vendidas e 1.050 exercidas) parava todas as
  telas. Também impedia lançar o próprio evento que resolveria o problema. Agora a tela de
  eventos propõe os eventos até a data em que o livro parou. Dá para lançar os eventos um por
  vez: um evento só é recusado se criar pendência nova.
- Erro no registro ao sair da tela de opções ou de eventos durante o cálculo.

## [3.0.1] - 2026-10-27

### Corrigido
- **A janela não abria** depois de extrair o `.zip` baixado: o Windows marca cada arquivo
  extraído como "baixado da internet", e com essa marca a DLL da janela não carregava. O
  programa ficava rodando sem janela e sem nenhum aviso. Agora o COIR tira essa marca das próprias
  DLLs ao abrir. Na 3.0.0, o contorno é desbloquear o `.zip` (Propriedades › Desbloquear) antes
  de extrair.

### Adicionado
- Página do manual **Se o COIR não abrir**: aviso do Windows, antivírus que põe o programa em
  quarentena e Controle Inteligente de Aplicativos. Link no LEIA-ME e no rodapé do site.

## [3.0.0] - 2026-10-26

Nova versão escrita do zero, sem dependência do Excel. É um programa para Windows que se baixa,
extrai e abre (`COIR.exe`), sem instalar Python nem Java. Os dados ficam numa pasta sua
(`COIR_DADOS`); nada sai do computador. O código da v2 (2.2.0) fica em `legacy/` como referência.

### Em resumo
- **Telas**: painel do ano, importação de notas e planilhas da B3, notas com o PDF ao lado,
  operações, apuração mensal com memória de cálculo, DARF (com multa e juros de mora, pagamento e
  comprovante), proventos, outros custos, declaração anual ficha a ficha, relatórios, carteira e
  preço médio, eventos corporativos, opções, futuros, conciliação com a B3, cadastros e
  configurações. Tema claro e escuro.
- **Leitura das notas**: um só extrator SINACOR (Bovespa e BM&F) para Ágora, BTG/Necton, Clear,
  Rico e XP, com conciliação ao centavo: nota que não fecha não entra.
- **Imposto**: preço médio por titular, day trade pela marcação da corretora, isenção de
  R$ 20 mil, compensação de prejuízo por categoria, FII, opções (exercício e pó), termo, futuros
  (inclusive em US$), IRRF por ficha, DARF com mínimo de R$ 10 e vencimento no último dia útil.
- **Declaração**: Bens e Direitos (um item por ativo, CNPJ e discriminação prontos), rendimentos
  isentos e exclusivos por fonte pagadora, renda variável mês a mês, conferência com a declaração
  entregue.
- **Tabelas públicas** (eventos, proventos, feriados, Selic) atualizadas pelo botão "Atualizar
  tabelas"; aviso de versão nova do programa (só avisa, não baixa).
- **Atualização segura**: o banco é migrado automaticamente com backup; banco criado por versão
  mais nova é recusado com explicação.
- O que ficou para depois está em [`docs/roteiro-pos-3.0.md`](docs/roteiro-pos-3.0.md).

### Por que meus números mudaram em relação à v2
A v3 refaz o cálculo a partir das notas, pela norma. Diferenças esperadas em relação à planilha
da v2 (todas são correções da v2):
- **Preço médio por titular**, somando todas as corretoras (a v2 usava o PM por conta).
- **Custos da nota inteiros**: a v3 usa o total que concilia com o resumo da nota; a v2 lançava
  parte dos custos.
- **IRRF de day trade (1%) deduzido** no mês; a v2 não o deduzia.
- **Saldo de IRRF transportado** para os meses seguintes do ano; a v2 o perdia no próprio mês.
- **Isenção de R$ 20 mil pelo valor das vendas de ações à vista do mês**; a v2 a inferia pelo IRRF,
  o que falhava com duas ou mais corretoras e podia isentar ETF.
- **Prejuízo de mês isento compensável**; a v2 o descartava.
- **FII como categoria própria** (20%, prejuízo separado); a v2 não tinha FII.
- **DARF**: mínimo de R$ 10 (abaixo disso acumula para o mês seguinte) e vencimento no último dia
  útil do mês seguinte; a v2 não aplicava o mínimo e usava o último dia corrido.
- **Eventos corporativos** (desdobramento, grupamento, bonificação, incorporação, cisão, troca de
  ticker) aplicados pelas tabelas públicas da B3, sem ajuste manual.
- **Exercício de opção** na linha do ativo-objeto, sem isenção; opção que virou pó apurada no
  vencimento.
- **IRRF até R$ 1,00** (dispensado, não retido) não vira crédito.

### Adicionado
- Licença da v3: AGPL-3.0-or-later (a v2 em `legacy/` continua MIT).
- Fase F0 (fundação): projeto uv com Python 3.13, estrutura `coir/`, configuração da pasta de dados (`COIR_DADOS`), CLI mínima, verificador de dados pessoais (pre-commit, pre-push e CI) e testes automatizados em Windows e Linux.
- Fase F8 (declaração anual): fichas do programa IRPF por exercício (`data/irpf/<exercício>.yaml`) com Bens e Direitos (um item por ativo, CNPJ do pacote público da B3, discriminação pronta), Dívidas (posição vendida), rendimentos isentos e exclusivos por fonte pagadora, renda variável mês a mês e imposto pago; tela "Declaração anual" com botão de copiar e marcação do que já foi lançado; leitura e conferência da declaração entregue (PDF do programa); data do crédito do JCP (a mais tarde entre a aprovação e a data com) pela B3, inclusive pelo histórico público de JCP desde 2016 (`data/proventos_historico_b3.csv`), pela declaração entregue ou informada; JCP pago mais de um ano depois do crédito a receber em cada 31/12; origem dos reembolsos de BTC; venda no dia da liquidação do termo pelo preço a termo, com o excedente saindo das ações antigas pelo custo médio; lista gerada das fontes fiscais (`docs/fontes-fiscais.md`).
- Fase F9 (versão 3.0): programa para Windows em pasta (`COIR.exe` e `coir-cli.exe`, PyInstaller) com `tools/empacotar.py` e workflow que cria o release como rascunho; identidade visual da logomarca (topo marinho, menu retrátil); aviso de versão nova; banco mais novo que o programa recusado com tela de erro; day trade pela marcação D da corretora (parâmetro `day_trade_criterio`); saldos iniciais em 31/12 (custódia e prejuízo por categoria); origem do reembolso de BTC informada pelo usuário; leitura da DIRPF 2024; registro de atividades legível; gerador de notas fictícias (`tools/notas_ficticias.py`) e banco de demonstração; capturas de tela automáticas (`tools/capturas.py`); manual do usuário em `docs/manual/`, guia de como adicionar corretora e site novo.

## [2.2.0] - 2026-10-21

Versão da v2 (planilha), em `legacy/`.

Atualização para o Python 3.15 e para as versões atuais das bibliotecas (pandas 3, numpy 2.5, tabula-py 2.10, xlwings 0.37, openpyxl 3.1). Compatibilidade com o pandas atual ([#11](https://github.com/MarceloPCF/COIR/issues/11)), com base na contribuição de [@tommmdl](https://github.com/tommmdl) ([#34](https://github.com/MarceloPCF/COIR/pull/34)).

Os resultados foram comparados com os da versão 2.1.1 em todas as notas de teste (XP, BTG e Rico), valor por valor, e no Excel em uma amostra de notas de todos os anos e corretoras.

### Corrigido
- Preenchimento de células vazias com `fillna(..., inplace=True)` em uma coluna trocado por atribuição. No pandas 3 a forma antiga não altera nada e os valores vazios das taxas continuariam vazios. (@tommmdl)
- Agrupamentos (`agrupar`, `agrupar_btg`, `agrupar_bmf`) usam `'sum'` em vez da função `sum` do Python, depreciada no pandas 2. (@tommmdl)
- `agrupar_operacoes` e `agrupar_operacoes_correcao`: no pandas 3, `groupby(Series)` seguido de `get_group` falha quando o grupo tem uma só linha, e o erro era engolido, descartando a operação sem aviso. Agora o agrupamento é feito pelo nome da coluna.
- Notas XP/Rico/Clear de várias páginas em que algumas linhas trazem só a quantidade (sem marca de observação) davam erro (`TypeError`) ou, no pandas 3, quantidade 0. Exemplo: XP 2024-11.
- Python 3.15 usa UTF-8 por padrão; a saída do `taskkill` (usado para fechar o Excel) é lida agora com tolerância a caracteres inválidos, em vez de interromper a gravação no Excel.

### Alterado
- Python mínimo: 3.12 (exigido pelo numpy 2.5). Pacote para Windows com Python 3.15.
- `requirements.txt` com as versões atuais testadas.
- `tools/montar_pacote.py` passou a calcular as dependências do pacote a partir das bibliotecas instaladas (arquivo `RECORD` de cada uma), em vez de uma lista fixa de arquivos, e aceita a versão do Python embutido informada pelo arquivo `.zip`.

## [2.1.1] - 2026-10-09

Correção de instalação: uma instalação nova da 2.1.0 falhava por incompatibilidade de bibliotecas. Nenhuma alteração no código do programa.

### Corrigido
- `requirements.txt` fixa `numpy==1.23.5` e `setuptools==69.5.1`. Sem isso o pip instalava o numpy 2.x, incompatível com o pandas 1.3.3, e o setuptools 81 ou superior, que não traz o `pkg_resources` usado pelo tabula-py 2.3.0.

### Alterado
- O pacote para Windows passa a levar o **Python 3.9.7 e o Java 17 (Temurin) embutidos**. Não é preciso instalar Python, bibliotecas nem Java: basta extrair o `.zip` e dar dois cliques em `COIR.bat`.
- Estrutura do pacote simplificada: na pasta principal ficam `COIR.bat`, `Entrada`, `Saida`, `Resultado`, `dados`, `modelos` e `sistema` (programa, Python e Java).
- A instalação manual (Python próprio, `pip install -r requirements.txt`) continua possível para quem usa o código-fonte.
- Novo `tools/montar_pacote.py`, que monta o pacote para Windows.
- Planilha-modelo `modelos/COIR.xlsb` revisada: uma linha que estava oculta foi corrigida.

## [2.1.0] - 2026-10-08

Primeira versão do repositório reorganizado. O histórico anterior foi reiniciado (veja "Removido").

### Alterado
- Pastas reorganizadas: código em `coir/` (antes `Utils/`), planilha-modelo em `modelos/`, tabelas auxiliares em `dados/` (antes `Utils/Tickets/`) e capturas de tela em `docs/imagens/` (antes `Utils/Screenshots/`).
- `COIR.py` refatorado: despacho por corretora em uma tabela, validações em funções auxiliares e exigência explícita de Python 3.9.2 ou superior.
- A planilha de resultado passa a ser gravada uma única vez, ao final do processamento de todas as notas: os extratores das corretoras agora devolvem os dados em vez de gravar a cada nota.
- Ajustes nas áreas de leitura dos PDFs das corretoras BTG e XP/Rico/Clear e nas funções comuns (`coir/funcoes.py`).
- `requirements.txt` convertido para UTF-8 (as versões das bibliotecas não foram alteradas).

### Adicionado
- Leitura da conta do cliente por nota, quando um mesmo PDF reúne notas de contas diferentes.
- Correções manuais opcionais em `dados/ajustes_manuais.csv` (mudança de ticker, conta, corretora, quantidade ou preço); modelo em `dados/ajustes_manuais.modelo.csv`.
- Novos tickers em `dados/acoes.csv` (Lojas Renner, Moura Dubeux, PRIO, Porto Seguro).
- Número da versão exibido na abertura do programa (`VERSION` e `coir/__init__.py`).
- `.gitignore` e `.gitattributes`.
- Verificação contra dados pessoais (`tools/verificar_dados_pessoais.py`) e hook de pre-commit em `.githooks/`.
- Este `CHANGELOG.md`.

### Removido
- Executáveis e pacotes `.zip` das versões antigas, que existiam no histórico anterior do repositório.
- Link para um `style.css` inexistente em `index.html`.

## Versões anteriores (histórico reiniciado)

Até 2.0.4 o repositório usava tags fora do padrão e sem Releases. Elas foram removidas quando o histórico foi reiniciado. A tabela abaixo é apenas um registro das datas; os detalhes técnicos de cada uma não foram reconstruídos.

| Tag antiga | Data | Equivale a |
|---|---|---|
| v1.0.0 | 2024-01-05 | 1.0.0 |
| V1.0.1 | 2024-01-07 | 1.0.1 |
| v1.0.2 | 2024-02-08 | 1.0.2 |
| v1.0.3 | 2024-02-13 | 1.0.3 |
| v1.0.4 | 2024-09-19 | 1.0.4 |
| v.2.0 | 2024-11-13 | 2.0.0 (código dividido em pacote com um módulo por corretora) |
| v.2.01 | 2024-11-29 | 2.0.1 (página do projeto no GitHub Pages) |
| v.2.02 | 2024-12-25 | 2.0.2 |
| v2.03 | 2025-04-05 | 2.0.3 |
| v2.04 | 2025-04-05 | 2.0.4 |

Entre abril de 2025 e junho de 2026 houve atualizações sem tag (`xp_rico_clear.py`, `acoes.csv`, `opcoes.csv` e a planilha-modelo); elas estão incorporadas à 2.1.0.

[3.1.0]: https://github.com/MarceloPCF/COIR/releases/tag/v3.1.0
[3.0.1]: https://github.com/MarceloPCF/COIR/releases/tag/v3.0.1
[3.0.0]: https://github.com/MarceloPCF/COIR/releases/tag/v3.0.0
[2.2.0]: https://github.com/MarceloPCF/COIR/releases/tag/v2.2.0
[2.1.1]: https://github.com/MarceloPCF/COIR/releases/tag/v2.1.1
[2.1.0]: https://github.com/MarceloPCF/COIR/releases/tag/v2.1.0
