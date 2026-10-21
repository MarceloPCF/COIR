# Regras fiscais — renda variável pessoa física (base para coir/engine)

Situação em 05/10/2026. A MP 1.303/2025 caducou (Câmara, 08/10/2025): regras de renda variável seguem as
anteriores. Novidade vigente desde 01/01/2026: Lei 15.270/2025 (dividendos e IRPF mínimo).
Toda regra aqui vira parâmetro com vigência + teste. Confiança: A = alta, M = média (confirmar), ? = decisão do usuário.

| # | Regra | Implementação | Fonte | Conf. |
|---|-------|---------------|-------|-------|
| R1 | Operações comuns (ações, units, ETF, BDR, opções, futuros, termo): 15% sobre ganho líquido mensal | Categoria COMUM, base única com subtotais por mercado | IN RFB 1.585/2015 | A |
| R2 | Day trade: 20%; prejuízo só compensa com ganho DT | Categoria DAY_TRADE; excedente de quantidade vira COMUM | IN RFB 1.585/2015 | A |
| R3 | FII/Fiagro: 20%; sem isenção 20 mil; prejuízo só com ganho FII | Categoria FII (não existe no v2.02) | IN RFB 1.585/2015 | A |
| R4 | Isenção: vendas de AÇÕES no mercado à vista ≤ R$ 20.000 no mês (todas as corretoras/contas do CPF) | Soma do valor **bruto** das alienações por titular/mês (fracionário e venda por exercício incluídos); fora: units, ETF, BDR, FII, opções, futuros, termo e day trade | Lei 11.033/2004 art. 3º I | A |
| R4a | Units NÃO entram na isenção | Parâmetro `units_na_isencao = false` | decisão do usuário (05/10/2026) | A |
| R5 | Prejuízo de mês isento continua compensável | Parâmetro `prejuizo_mes_isento_compensavel = true` | Lei 8.981/1995 art. 72 §8º; Manual MIR 2026; decisão do usuário (14/10/2026) | A |
| R6 | IRRF 0,005% s/ vendas comuns (dispensado se ≤ R$ 1; não incide no exercício nem em PME) e 1% s/ ganho de DT | Ler da nota; os dois abatem o imposto total do mês (o de DT consumido primeiro); saldo aos meses seguintes do ano; em dezembro, o de 0,005% vai para a declaração anual e o de DT só restituição | Lei 9.959/2000 art. 8º §§4º, 8º; IN RFB 1.585/2015 arts. 63 §§8º e 9º e 65 §§8º, 9º e 12; P&R IRPF 2026 perguntas 714 e 715; a página "Retenções" da RFB (19/11/2025) restringe o DT ao imposto de DT: divergência de orientação registrada | M |
| R7 | Custo de aquisição inclui corretagem, emolumentos e taxas; na venda os custos reduzem a alienação | Rateio proporcional ao valor da operação na nota | IN RFB 1.585/2015 | A |
| R8 | Bonificação: soma quantidade com custo atribuído; desdobramento/grupamento: só quantidade; redução de capital com restituição reduz o custo. Incorporação/fusão/cisão/conversão: o custo das ações entregues vai para as recebidas na proporção da assembleia, sem resultado (cisão: percentual divulgado pela companhia, sobre o custo de antes do evento; sem percentual, bloqueia); incorporação de ações (art. 252): custo mantido, com aviso da divergência da RFB (parâmetro `incorporacao_acoes_tratamento`); ajuste de série de opção pela B3: sem efeito fiscal (código e quantidade mudam, custo e vencimento ficam); fração vendida em leilão: alienação pelo PM no mês do crédito, entra no limite de isenção (parâmetro `leilao_fracao_isentavel`). Data do evento: data ex | evento_corporativo na linha do tempo do ativo (tela 7) | IN RFB 1.585/2015 art. 58 caput, §§6º e 7º, art. 60 §2º; P&R IRPF (substituição de ações; incorporação de ações); SC Cosit 224/2014 (divergência); Lei 6.404 arts. 227-229 e 252 | A |
| R9 | Exercício de opções: prêmio integra custo (titular de call) / valor de venda (lançador de call); casos simétricos para put | Vincular opção, ativo-objeto e negócio de exercício; 4 casos com exemplo. A nota traz só o código da opção + "E" (ex.: KLBNO193E), às vezes + classe; ativo-objeto = 4 letras + classe (ON→3, PN→4, PNA→5, PNB→6, UNT→11); sem classe: classe da abertura da série, depois opcoes.csv | IN RFB 1.585/2015 | M |
| R10 | Opção que vira pó: prêmio pago = perda (titular); recebido = ganho (lançador) no vencimento | Evento automático de vencimento (data como parâmetro, não "hoje") | IN RFB 1.585/2015 | A |
| R11 | Futuros: resultado no encerramento ou liquidação da posição = soma dos ajustes desde a abertura − custos; encerramento parcial pelo preço médio; day trade no dia | Ajustes da posição aberta acumulados por contrato e realizados no encerramento ou vencimento | Lei 11.051/2004 art. 32 (caput e §2º II); IN RFB 1.585/2015 art. 61; P&R IRPF 2026 pergunta 728 | A |
| R12 | DARF código 6015; vencimento último dia útil (bancário) do mês seguinte; valor < R$ 10 acumula, inclusive para o ano seguinte | `data/feriados.csv` (o último dia útil do ano não é útil: Res. CMN 4.880/2020 art. 3º) + acúmulo | RIR/2018 art. 915; Lei 9.430/1996 art. 68 caput e §1º | A |
| R13 | Atraso: multa 0,33%/dia limitada a 20% + juros Selic (do mês seguinte ao vencimento até o mês anterior ao pagamento + 1% no mês do pagamento) | `data/selic_mensal.csv` (BCB SGS 4390) | Lei 9.430/1996 art. 61 | A |
| R14 | Dividendos > R$ 50 mil/mês da mesma PJ: IRRF 10% sobre o total; IRPFM para renda anual > R$ 600 mil | Registrar IRRF do provento; relatório informativo (IRPFM fora do escopo v3.0) | Lei 15.270/2025 | A |
| R15 | JCP: IRRF exclusivo 15% até 31/12/2025; 17,5% para JCP pago ou creditado desde 01/01/2026 | Provento com IRRF; alíquota com vigência | LC 224/2025 | A |
| R16 | Rendimentos de FII isentos se o fundo cumpre requisitos | Usuário marca conforme informe do fundo | Lei 11.033/2004 art. 3º (alt. Lei 14.754/2023) | M |
| R17 | Fichas/códigos da DIRPF mudam por ano | Mapeamento por ano em `data/irpf/<ano>.yaml` | programa IRPF do ano | revisar anualmente |
| R18 | Termo: custo/alienação pelo preço a termo; ganho apurado na liquidação/encerramento (operação comum) | Contrato a termo vinculado à liquidação | IN RFB 1.585/2015 | M (validar com as operações a termo da planilha) |

## Outros custos dedutíveis (aba BTC_Outros do v2)
Lançamento manual mensal (ex.: custos de BTC) que reduz o resultado tributável do mês, como no v2.02. BTC como
operação está fora do escopo.

## Decisões da validação do gabarito (F1, 14/10/2026)
Respostas do usuário aos pontos do `resumo.md`. Onde o gabarito difere da regra por desenho, a comparação
da F4 trata como divergência conhecida, nunca ajusta o motor para bater.

- **PM global por titular (R7/R8):** para apuração e declaração vale o PM do titular somando todas as
  corretoras/contas. A tela também mostra o PM por corretora/conta. O gabarito calcula por papel + conta,
  então diverge por desenho quando o mesmo ticker está em mais de uma conta.
- **Opções lançadas (R9/R10):** resultado apurado no mês do encerramento/vencimento, não no do pregão, como
  no gabarito. Confirmado pelo usuário; revisar com o `especialista-fiscal-b3` na F4.
- **IRRF dispensado (R6):** IRRF ≤ R$ 1,00 não retido (o PDF não o desconta) não vira crédito a deduzir; só
  conta o que foi efetivamente retido (a BTG acumula no mês e retém quando o total passa de R$ 1,00). O
  gabarito registra esses valores: divergência conhecida.
- **Prejuízo em mês isento (R5):** em aberto para a F4. O gabarito o exclui do resultado (ex.: fracionário
  06/2008). R5 (P&R IRPF) diz que continua compensável; uma resposta de IA consultada pelo usuário diz o
  contrário, sem fonte primária. Decidir com o `especialista-fiscal-b3` e a pergunta da edição vigente do
  P&R IRPF antes de implementar; enquanto isso, parâmetro configurável.
- **Isentos zerados de 2021 em diante:** correto. **Custódia inicial vazia:** correto (histórico começa
  zerado em 2008).
- **Outros custos dedutíveis (BTC_Outros):** o usuário não os lançou. Verificar nas fases de telas se a
  outros sistemas de IR têm equivalente (`docs/referencias/`); sem equivalente, não implementar.

## Decisões da validação da extração (F2, 14/10/2026)
- **Termo (R18), regra do usuário:** durante o contrato, as ações compradas a termo ficam numa posição
  separada (como na B3); na liquidação (vencimento ou antecipada), quando o contrato é pago e as ações
  são recebidas, entram na carteira à vista e o PM se junta ao da posição existente. A nota deixa a
  compra fora do líquido ("Operações à termo"); o prazo (dias) fica na operação extraída. O gabarito v2
  soma na data do pregão: diferença de PM esperada entre pregão e liquidação. Decisão do usuário
  (14/10/2026): data da liquidação = data do pregão + prazo do contrato; os custos da nota de compra
  (corretagem, emolumentos, taxa de termo) entram no custo das ações liquidadas. Validar com o
  `especialista-fiscal-b3` na F4.
- **ETF de renda fixa (ex.: B5P211):** tributação exclusiva na fonte, retida pelo administrador na venda
  com lucro. Registrar a posição e o custo, mas fora da apuração mensal e do DARF; vai para a DIRPF
  como rendimento de tributação exclusiva. Precisa de tipo de ativo próprio em `data/` (F4, confirmar
  alíquotas e base com o `especialista-fiscal-b3`).
- **Transferência de custódia entre corretoras do mesmo titular:**
  evento próprio que move quantidade e PM de uma conta para outra, sem resultado. A nota continua na
  corretora que a emitiu; a planilha v2 relabelava as notas.
- **Opção flexível (BTG):** fica para uma versão seguinte à v3.0 (parser do layout "Mercadoria: FCA"
  e operação de opção flexível); na v3.0 a extração reconhece e avisa.

## Decisões da F4 (motor fiscal, 14-15/10/2026)
Respostas do usuário e pareceres do `especialista-fiscal-b3`; todos os valores estão em
`data/parametros_fiscais.csv` com vigência e fonte, e os casos em `tests/fiscal_cases/`.
- **Regra geral (15/10/2026):** a norma oficial prevalece sobre preferência do usuário; havendo
  divergência ou ambiguidade, o usuário é consultado antes de aplicar.
- **Day trade:** casamento pela ordem dos negócios (IN 1.585 art. 65 §3º: 1ª compra com a 1ª venda),
  pela ordem impressa nas notas do dia (número da nota e linha; a nota não traz horário); o
  excedente vira comum pelo preço do próprio negócio. Substitui a "média do dia". Antes de
  28/07/2010 conta em qualquer corretora (MP 497/2010 art. 21). Exercício de opção não é day trade
  (IN 1.585 art. 65 §13). Day trade de FII fica na categoria FII. Sem compensação entre comum e day
  trade (Lei 9.959/2000 art. 8º §§6º e 7º).
- **IRRF:** saldo único (0,005% e 1%) abate o imposto total do mês; o saldo passa aos meses
  seguintes do ano e não passa para janeiro. Em dezembro, o de 0,005% vai para a declaração anual
  (IN 1.585 art. 63 §8º III) e o de day trade só pode ser pedido em restituição (art. 65 §§9º e
  12 II; P&R 715), como em R6. Só conta o IRRF impresso na nota.
- **IRRF por ficha do IRPF (decisão de 16/10/2026: separar só na declaração):** a DARF continua no
  saldo único (a lei e a IN não separam o IRRF de FII; Lei 11.033/2004 art. 2º §7º). As fichas do
  programa (comuns/day trade e FII) recebem o IRRF que cada uma consumiu nesse saldo, mais o
  retido nela que sobrou; os saldos de dezembro das fichas somam o saldo único (sem uso duplo).
  - Ordem de consumo: comuns usa o day trade, o 0,005% dela e depois o de FII; FII usa o 0,005%
    dela, o de comuns e, só se faltar, o day trade (vira aviso: a ficha FII não tem essa linha).
  - O 0,005% de uma nota mista é repartido pelas vendas comuns de cada ficha (sem day trade nem
    venda por exercício), por titular, corretora, conta e mês; o centavo residual fica em comuns.
  - "Imposto pago (calculado)": o a pagar do mês quando a DARF que o cobriu (inclusive acumulada
    abaixo de R$ 10) já foi emitida.
- **Exercício de opção:** o ganho da venda de ações por exercício nunca é isento (IN 1.585 art. 59
  §2º IV). O valor dessa venda entra na soma do limite das demais vendas (decisão de 15/10/2026,
  abaixo).
- **Decisões de 15/10/2026 (especialista + usuário, norma oficial primeiro):**
  - IRRF de day trade abate o imposto total do mês (leitura da lei, da IN e do P&R 714); o de day
    trade é consumido primeiro, porque o saldo dele em dezembro só pode ser pedido em restituição.
  - Venda por exercício (e na liquidação de termo): ganho sempre tributável; o valor entra na
    soma do limite (nenhuma fonte decide; leitura conservadora do usuário).
  - Futuros no encerramento (Lei 11.051/2004 art. 32), encerramento parcial pelo preço médio.
  - Units fora da isenção: coincide com a RFB (SC Cosit 145/2021: só "ativos da espécie ações").
  - Venda a descoberto (sem norma específica, leitura A1): resultado na recompra; o limite de
    isenção é o do mês da venda.
  - Isenção de PME (Lei 13.043/2014 art. 16; IN arts. 59 III e 66 a 69-A), `data/isencao_pme.csv`:
    PRIO3 (compras de 10/07/2014 a 29/01/2021) e AGRO3 (até 03/02/2021), vendas até 31/12/2023;
    ganho isento, fora do limite, perdas não compensam; day trade dessas ações também isento
    (leitura literal: o art. 59 §2º I só afasta o inciso I). Lotes isento e tributado juntos
    falham alto (controle separado do art. 69-A §4º não implementado; não ocorre no livro). Depois
    de 31/12/2023 a isenção acabou e o lote antigo se junta às compras novas.
- **PM e carteira (princípios do usuário, 15/10/2026; IN 1.585 art. 58):**
  - PM fiscal consolidado por titular e ativo, somando todas as corretoras e contas, na posição
    da categoria de operação aplicável: carteira comum à vista; opção por série e vencimento;
    termo em posição separada até a liquidação; lote de PME separado do tributado. É o PM da
    apuração mensal e da declaração (os dois precisam ser o mesmo, senão a soma dos custos não
    fecha com o que foi pago).
  - Corretora e conta de origem preservadas em cada movimento, em qualquer visão. Ficam vazias só
    quando o passo junta contas: day trade entre contas diferentes (antes de 28/07/2010, entre
    corretoras) ou o pó de uma série comprada em mais de uma conta.
  - Day trade apurado à parte (ordem dos negócios, mesma corretora), sem tocar no PM da carteira.
  - Carteira por corretora e por conta (`coir carteira --visao corretora|conta`): visão
    operacional, igual ao que a corretora e a B3 mostram. Não entra no imposto. As quantidades
    fecham com a visão fiscal; o custo pode diferir quando o mesmo ativo esteve em mais de uma
    conta.
- **Opções:** a posição é por série e vencimento (o mesmo código de série volta em outro ano).
  Vencimento: 3ª segunda-feira (sem pregão: o seguinte) até abr/2021; 3ª
  sexta-feira (sem pregão: o anterior) desde mai/2021; semanais Wn na n-ésima sexta do mês (B3,
  notícia de 20/04/2021). Exercício antecipado é aceito em qualquer data. A série segue a troca de
  ticker do ativo-objeto (RRRPT32 -> BRAVT32): na normalização e também no ativo da série, de
  modo que a série comprada com o código antigo e vendida com o novo é uma posição só (EMBRE790 ->
  EMBJE790, troca EMBR3 -> EMBJ3). Bonificação na ação ajusta a quantidade da série na data ex
  (evento sem custo, `origem=ajuste_serie` no eventos.csv).
- **Futuros:** resultado no encerramento ou no vencimento (ver "Decisões de 15/10/2026"); ajustes e
  custos rateados pelo volume.
- **Termo:** entra na carteira na liquidação (pregão + prazo em dias corridos; sem pregão, o
  seguinte), pelo preço a termo + custos da nota de compra. Ações vendidas antes da liquidação
  deixam a posição vendida; a liquidação a cobre e realiza o resultado nessa data. Esse ganho nunca
  é isento (IN RFB 1.585/2015 art. 59 §2º IV); a venda entra na soma do limite do mês dela.
- **ETF de renda fixa:** fora da apuração e do DARF; ganho listado para a DIRPF (Lei 13.043/2014).
- **Calendário:** dois calendários em `data/feriados.csv` (bancário para o DARF; pregão B3 para
  opções e termo): feriados de São Paulo fecharam a B3 até 2021; 24/12 sem pregão; exceções de 2014
  (Copa) e 2020 (feriados de SP antecipados).
- **Comparação com o gabarito (DARFs pagos):** cada DARF confere (até R$ 1,00) ou tem a diferença
  explicada em `COIR_DADOS/gabarito/darf-explicados.csv` (fora do repositório). Causas encontradas,
  todas defeitos da planilha v2: saldo de prejuízo pelo PM por conta (o COIR usa o PM por titular);
  custos da nota lançados só em parte (o COIR usa o total, que concilia com o resumo); IRRF de day
  trade que a v2 não deduziu; vendas tratadas como isentas em meses acima de R$ 20 mil (explicadas
  pela isenção de PME); prejuízo de mês isento descartado; IRRF não transportado.

- **Comparação com uma planilha externa (`coir gabarito externo`, operações finalizadas exportadas
  de outro sistema de IR):** compara o resultado por dia, corretora e ativo e por mês e mercado.
  A planilha calcula o PM por corretora (`--visao corretora` confere mais). Divergências
  analisadas ficam em
  `COIR_DADOS/gabarito/divergencias-externo.csv` (fora do repositório). Diferenças de método, em que
  o COIR segue a norma: day trade pela ordem dos negócios; exercício e venda no mesmo dia pelo
  custo médio, não casados; prêmio da call lançada e exercida na alienação. Opção flexível fica
  para depois da v3.0. `--desde AAAA-MM` compara a partir de um marco (ex.: o mês seguinte a uma
  carteira zerada), quando o PM inicial lançado à mão no outro sistema não é confiável. Desde o marco,
  as diferenças encontradas foram:
  - centavos: o outro sistema rateia os custos da nota sem arredondar a parcela (com rateio exato, confere).
    Nenhuma norma fixa o critério de rateio nem o arredondamento da parcela (a IN RFB 1.585/2015 só
    manda somar os custos ao custo de aquisição e deduzi-los da alienação); o COIR mantém
    centavos com maiores restos, que somam exatamente a nota (decisão de 16/10/2026);
  - futuros: o outro sistema distribui os custos de cada nota BM&F entre as linhas por outro critério; a
    soma de todos os futuros é igual (diferença total R$ 0,00 fora a opção flexível). Pendência
    pós-v3.0: ratear os custos de futuros por contrato e por categoria (a B3 cobra emolumentos e
    registro por contrato, com tabela própria de day trade), configurável com vigência; a v3.0
    rateia pelo volume;
  - notas de exercício: outro critério de rateio, até R$ 2 por ativo (às vezes o outro sistema deixa
    de fora a taxa de transferência de ativos);
  - dados do outro sistema: custos de uma nota de exercício e de uma venda de opção não lançados;
  - dados do COIR: séries de opção com custo na bonificação (só a quantidade muda) e custo
    atribuído de uma bonificação diferente do oficial, corrigidos no `eventos.csv`.

## Relatórios (F5, 16/10/2026)
`coir exportar` monta os relatórios de um titular e um ano só com os valores do motor e das notas
(regra 12), sem cálculo novo. Saem em xlsx, csv (padrão brasileiro) e PDF (ADR 0003).
- Escolhas do usuário:
  - os relatórios mapeados e gratuitos das plataformas de referência, também para a interface (F6);
  - nome e CPF completo no cabeçalho, mas nunca no nome do arquivo;
  - csv com `;` e vírgula decimal.
- A apuração vai até o fim do ano ou até a data de geração: opção que vence depois de hoje não vira
  pó no relatório do ano em curso.
- "Renda variável mês a mês" segue a ordem dos blocos do programa da Receita. O IRRF aparece em
  duas partes, day trade e Lei 11.033, com mês, meses anteriores e a compensar; o de day trade é
  usado primeiro. As fichas por ano ficam na F8.
- Fichas do programa: o total de comuns e day trade não leva o imposto de FII, que tem ficha
  própria. O IRRF de cada ficha segue "IRRF por ficha do IRPF" (acima), com imposto pago.
- Futuros: cada mercado de `data/contratos_futuros.csv` tem linha na apuração (índice, dólar,
  commodities em "outros"). Um mercado sem linha (ex.: juros) falha alto.
- Para a F8 (revisão fiscal da F5):
  - Bens e Direitos com situação em 31/12 do ano anterior, grupo e código do bem por ano, CNPJ,
    código de negociação e discriminação;
  - opções sobre ETF ou índice na linha "outros";
  - linha "Mercado a termo" para o ganho realizado na liquidação (o usuário manteve o à vista em
    16/10/2026; rever na F8);
  - conferir no programa IRPF do ano: a ficha FII sem linha de IR fonte de day trade; ordem de
    abatimento (day trade antes da Lei 11.033) na ficha de comuns.
- Ficam para outras fases:
  - patrimônio e rentabilidade, que precisam de cotação (F10);
  - proventos (F7);
  - os relatórios que as plataformas vendem à parte.

## Declaração anual (F8, 20/10/2026)
Parecer do `especialista-fiscal-b3` e decisões do usuário; casos em
`tests/fiscal_cases/test_irpf_linhas.py`. Critério de saída: reproduzir a DIRPF 2026 (ano-calendário
2025) entregue, lida do PDF do programa (`coir/gabarito/dirpf.py`, fora do repositório).
- Linhas da ficha "Renda variável - operações comuns/day trade": opção cujo ativo-objeto é ação ou
  unit em "opções - ações"; ETF, índice e outros em "opções - outros" (parâmetro
  `opcoes_linha_acoes_tipos_objeto`; P&R IRPF 2026 pergunta 707 separa cotas de fundo de índice de
  ações; unit como ação, como no Bens e Direitos 03.01 da DIRPF entregue). Objeto não identificado
  pela raiz fica em ações. Futuro de juros (DI1) tem linha própria (`grupo = juros` em
  `data/contratos_futuros.csv`; o contrato ainda não está cadastrado).
- Termo (IN RFB 1.585/2015 art. 62 I e §1º): a venda à vista na data da liquidação vai para
  "Mercado a termo - ações/ouro", com ganho pelo preço a termo + custos da compra, e nunca é isenta
  (art. 59 §2º IV). Se o lote não for vendido nessa data, o custo de aquisição é o preço a termo +
  custos (art. 62 §1º) e entra no custo médio (art. 58); a venda posterior é à vista comum e pode
  ser isenta (decisão do usuário de 20/10/2026, mantida em 21/10 apesar do "ou posteriormente" do
  P&R IRPF 2026 pergunta 723, que trata de lote isolado).
- Termo com ações antigas em carteira (decisões do usuário de 21/10/2026, parecer do especialista):
  no dia da liquidação, as vendas consomem primeiro o lote do termo, pela ordem dos negócios (por
  analogia com o art. 65 §3º; nenhuma norma decide); o excedente sai da posição anterior pelo PM e
  pode ser isento; a venda inteira entra na soma do limite de R$ 20 mil (leitura de 15/10). O que
  sobra do lote entra no custo médio pelo preço a termo. Exemplo: 500 a 9,00 + termo de 1.000 a
  10,00, venda de 600 a 11,00: ganho 600,00 no termo e 900 ações com custo 8.500,00 (pelo PM
  misturado seriam 800,00 e 8.700,00). A venda feita antes da liquidação e coberta por ela continua
  na linha à vista (nenhuma fonte decide).
- Prejuízo herdado em janeiro: o recalculado pelo COIR, com aviso da diferença para o declarado
  (decisão do usuário; retificação dentro de 5 anos, se for o caso).
- Bens e Direitos: um item por ativo, custo médio do titular somando corretoras (P&R IRPF 2026
  pergunta 720); a discriminação lista a quantidade em cada corretora. CNPJ do emissor pelo pacote
  público da B3 (`data/emissores_b3.csv`).
- Bens e Direitos pelo custo de aquisição (IN RFB 1.585/2015 art. 58), também no ano da compra; a
  DIRPF 2026 entregue trouxe valor de mercado nos itens novos de 2025: divergência explicada, o COIR
  segue a norma (decisão do usuário de 20/10/2026).
- Posição vendida em 31/12 (opção lançada, venda a descoberto): não é bem; vai para Dívidas e ônus
  reais, código 16, pelo valor recebido (prêmio ou venda), até o encerramento. Sem norma
  específica; decisão do usuário de 20/10/2026 para fechar a variação patrimonial.
- JCP: declarado no ano do crédito (P&R IRPF 2026 pergunta 320; Lei 9.249/1995 art. 9º §2º), com
  a alíquota da data do crédito ou do pagamento, o que vier primeiro (LC 224/2025 art. 8º). A data
  do crédito vem da correção do usuário na tela de Proventos (informe da empresa, tabela
  `credito_jcp`) ou da data de aprovação no pacote público da B3 (`proventos_b3.csv`: mesma
  empresa, data de pagamento e valor por ação dentro de 1%). Sem nenhuma das duas, o ano anterior
  só é percebido pela mudança de alíquota. A B3 só lista os pagos desde mai/2025: os créditos de
  2024 pagos no começo de 2025 precisam da correção do usuário. Crédito depois do pagamento falha
  alto (decisão do usuário de 20/10/2026).

## DARF paga em atraso: multa de mora (F6, 17/10/2026)
`coir/engine/darf.py`; casos conferidos à mão em `tests/fiscal_cases/test_multa_mora.py`.
- 0,33% por dia de atraso sobre o principal, com teto de 20% (Lei 9.430/1996 art. 61 caput e
  §2º). Os percentuais ficam em `data/parametros_fiscais.csv`, com vigência pela competência.
- Contagem: a lei (§1º) conta do dia seguinte ao vencimento; a Receita (Sicalc e "Como calcular
  multa de mora") conta do primeiro dia útil seguinte. Decisão do usuário: como a Receita
  (`multa_mora_inicio_contagem = rfb`); `lei` continua disponível. As duas só diferem quando o
  vencimento cai antes de fim de semana ou feriado.
- Pagamento informado em dia não útil vale no próximo dia útil (glossário do Sicalc; decisão do
  usuário). A tela avisa o ajuste.
- Centavos: a lei é omissa; o Sicalc trunca (`multa_mora_arredondamento = ROUND_DOWN`).
  Conferido pelo usuário em 17/10/2026: 6015, competência 08/2026, R$ 1.234,57, vencimento
  30/09/2026, consolidação em 07/10/2026 → multa 28,51 (7 dias, 2,31% = 28,518567) e juros
  12,34 (1% do mês do pagamento; caso para a F7).
- DARF acumulada abaixo do mínimo não tem multa; vence com o período em que o total chega ao
  mínimo (art. 68 §1º), e a multa conta desse vencimento.
- O pagamento é um fato lançado pelo usuário (tabela `darf_pagamento`, comprovante em
  `<dados>/comprovantes/<ano>/`, nome sem CPF) e não muda a apuração. Pago a menor ou a maior: a
  tela mostra a diferença (DARF complementar no Sicalc; restituição pelo PER/DCOMP), sem
  imputação automática.
- A denúncia espontânea (CTN art. 138; STJ REsp 1.149.022/SP; Súmula 360) fica como nota na tela:
  o COIR calcula como o Sicalc.
- Juros (art. 61 §3º c/c art. 5º §3º; F7a, 19/10/2026): soma simples da Selic mensal (BCB SGS 4390,
  a mesma do Sicalc, em `data/selic_mensal.csv`) dos meses entre o do vencimento e o do pagamento, sem
  contar nenhum dos dois, mais 1% no mês do pagamento; pago no mesmo mês do vencimento, zero. Base: o
  principal. Centavos truncados (Sicalc, 17/10/2026: 12,3457 → 12,34). Data não útil vale no próximo
  dia útil, o que pode mudar o mês. DARF acumulada: juros desde o vencimento do período que fechou o
  mínimo (art. 68 §1º). Mês posterior ao último da tabela: estimativa avisada; buraco na tabela: erro.
  A tabela confere com as tabelas de acréscimos da RFB (`tests/fiscal_cases/test_juros_mora.py`).

## Futuros de commodities e conferência das notas BM&F (F7a, 19/10/2026)
`coir/parsing/futuros.py`; casos em `tests/fiscal_cases/test_futuros_commodities.py`.
- O resultado vem dos ajustes em R$ impressos na nota (IN RFB 1.585/2015 art. 61). O multiplicador só
  entra no volume do rateio de custos e na conferência abaixo. Café (ICF) e soja CME (SJC) são cotados
  em US$; o ajuste da nota já vem em R$.
- Conferência do câmbio implícito (decisão do usuário: erro, a nota não entra): com o preço de ajuste
  do dia (linha AJUPOS) ou com os negócios que se fecham no dia, valor = (PA − preço) × quantidade ×
  multiplicador × câmbio; contrato em R$ tem câmbio 1 (confere o multiplicador nota a nota), em US$ o
  mesmo câmbio em todas as linhas. As 624 notas BM&F reais passam.
- Rateio de custos numa nota que mistura contratos em R$ e em US$: o volume em US$ é convertido pelo
  câmbio implícito da própria nota; sem ele, a nota não entra. Nota de uma moeda só: sem conversão.
- Fora do escopo, com falha alta: tipo de negócio além de NORMAL, DAY TRADE, AJUPOS e taxa de
  permanência (entrega física, exercício); opção sobre futuro; contrato sem cadastro. Café com posição
  aberta no vencimento (liquidação física): o resultado é o dos ajustes e o COIR avisa.
- A linha da apuração de cada contrato (índice, dólar, "outros") está em `data/contratos_futuros.csv`.

## Proventos (F7a, 19/10/2026)
`coir/engine/proventos.py`, a partir da movimentação da B3; casos em `tests/fiscal_cases/test_proventos.py`.
- Dividendo de ação ou unit: isento (Lei 9.249/1995 art. 10), código 09. BDR, ETF e FII com
  "Dividendo": fora da v3.0, listados com aviso (BDR não é isento, Lei 14.754/2023).
- JCP: exclusivo (código 10, pelo líquido). O preço unitário da B3 é o bruto por ação, o que mostra a
  alíquota retida (15% até 2025; 17,5% desde 2026, LC 224/2025 art. 8º, pela data do pagamento ou do
  crédito). Pago em 2026 com 15% = creditado em 2025: declarado no ano do crédito, com o crédito a
  receber em 31/12 (Bens e Direitos 99.07) — decisão do usuário. Bruto e IRRF estimados (vale o
  informe).
- Rendimento de FII/Fiagro: isento (Lei 11.033 art. 3º III).
- "Rendimento" em ação pago junto com dividendo ou JCP (BB, BB Seguridade, Caixa Seguridade): é a
  atualização do provento pela Selic; tributação exclusiva como renda fixa (Lei 8.981/1995 art. 65
  §4º c; Lei 11.033 art. 1º), como as fontes retêm (22,5% nos dados reais); código 06 ou 12 conforme
  o informe. Alternativa (acessório segue o principal, SC Cosit 67/2024) não adotada.
- Reembolso de BTC ao doador: isento até 2025 (Lei 13.043/2014 art. 7º, reproduzido na IN 1.585
  art. 74); desde 2026 a Lei 15.265/2025 é omissa para dividendos: isento por equivalência, com
  aviso (parâmetro; decisão do usuário). O reembolso de JCP vem descontado do IR que a companhia
  reteria (art. 7º §1º II; Lei 15.265 art. 20), mas o desconto não é IRRF do doador; é isento até
  2025 e, desde 2026, sem incidência expressa (Lei 15.265 art. 21 I). Declarado no ano do
  recebimento, em isentos "Outros" (99), com o CNPJ da companhia, sem JCP a receber (99.07): o
  crédito da companhia foi para o tomador (parecer do especialista e decisões do usuário de
  21/10/2026). A origem (dividendo ou JCP) sai do valor: quantidade emprestada inteira vezes o
  valor por ação de um provento do mesmo dia e empresa (movimentação do titular e pacote da B3),
  com o fator 1 - alíquota no JCP; sem origem única, aviso. Desde 2026 o aviso da Lei 15.270
  soma os reembolsos de dividendos.
- Restituição de capital: reduz o custo (IN 1.585 art. 58 §8º), proposta como evento; não é
  rendimento. Acima do custo: o motor falha alto.
- "- Transferido": pares que se anulam (provento acompanhando a transferência de custódia).
- Lei 15.270/2025 (desde 2026): só aviso, somando os dividendos líquidos da mesma empresa (raiz do
  código) no mês: acima de R$ 50 mil, ou acima de R$ 45 mil como "possível retenção".

## Pendências do motor (revisão final da F4, 16/10/2026)
Decididos na revisão: o IRRF de day trade impresso como "projeção" nas notas da XP é o 1% retido
sobre o ganho do dia e é deduzido no mês (confirmado pelo usuário); day trade "na mesma instituição
intermediadora" (Lei 9.959/2000 art. 8º §1º) compara a corretora da nota: XP, Clear e Rico são
corretoras diferentes, com código e contas próprios (decisão do usuário). Ficam para depois:
- Rateio dos custos de futuros por contrato e por categoria, configurável (hoje pelo volume).
- Day trade entre corretoras antes de 28/07/2010: a ordem dos negócios entre notas de corretoras
  diferentes segue o número da nota (as notas não trazem horário).
- Controle separado de lote de PME isento e tributado (IN 1.585 art. 69-A §4º): falha alto.
- Venda a termo: falha alto (só a compra a termo está implementada).
- `data/feriados.csv` cobre até 2027: estender antes de apurar dezembro de 2027.

## Pontos do v2.02 a corrigir (ver docs/PLANO.md, Diagnóstico)
- Isenção inferida por "IRRF < R$ 1" (falha com 2+ corretoras; pode isentar ETF).
- Sem categoria FII.
- Saldo de IRRF não transportado (L = J − K no próprio mês) — conferir.
- DARF sem mínimo de R$ 10 e vencimento no último dia corrido.
- Taxas de futuros e trocas de ticker fixas no código.

## Aviso
Este documento orienta a implementação; não substitui contador/advogado tributarista. Divergências de
entendimento devem virar parâmetro configurável e decisão explícita do usuário.
