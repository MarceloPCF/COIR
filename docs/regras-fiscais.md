# Regras fiscais — renda variável pessoa física (base para coir/engine)

Situação em 05/10/2026. A MP 1.303/2025 caducou (Câmara, 08/10/2025): regras de renda variável seguem as
anteriores. Novidade vigente desde 01/01/2026: Lei 15.270/2025 (dividendos e IRPF mínimo).
Toda regra aqui vira parâmetro com vigência + teste. Confiança: A = alta, M = média (confirmar), ? = decisão do usuário.

| # | Regra | Implementação | Fonte | Conf. |
|---|-------|---------------|-------|-------|
| R1 | Operações comuns (ações, units, ETF, BDR, opções, futuros, termo): 15% sobre ganho líquido mensal | Categoria COMUM, base única com subtotais por mercado | IN RFB 1.585/2015 | A |
| R2 | Day trade: 20%; prejuízo só compensa com ganho DT | Categoria DAY_TRADE; excedente de quantidade vira COMUM | IN RFB 1.585/2015 | A |
| R3 | FII/Fiagro: 20%; sem isenção 20 mil; prejuízo só com ganho FII | Categoria FII (não existe no v2.02) | IN RFB 1.585/2015 | A |
| R4 | Isenção: vendas de AÇÕES no mercado à vista ≤ R$ 20.000 no mês (todas as corretoras/contas do CPF) | Somar valor de alienação por titular/mês; excluir ETF, BDR, FII, opções, futuros | Lei 9.250/1995 art. 22 I; IN 1.585 art. 25 §1 | A |
| R4a | Units NÃO entram na isenção | Parâmetro `units_na_isencao = false` | decisão do usuário (05/10/2026) | A |
| R5 | Prejuízo de mês isento continua compensável | Saldo de prejuízo atualizado mesmo em mês isento | Perguntas e Respostas IRPF (edição vigente) | M |
| R6 | IRRF 0,005% s/ vendas comuns (dispensado se ≤ R$ 1) e 1% s/ ganho de DT | Ler da nota; deduzir do imposto do mês; transportar saldo aos meses seguintes do ano; saldo de dezembro vai para a DAA (restituição/compensação), não para o ano seguinte | IN RFB 1.585/2015; decisão do usuário | A |
| R7 | Custo de aquisição inclui corretagem, emolumentos e taxas; na venda os custos reduzem a alienação | Rateio proporcional ao valor da operação na nota | IN RFB 1.585/2015 | A |
| R8 | Bonificação: soma quantidade com custo atribuído; desdobramento/grupamento: só quantidade; redução de capital com restituição reduz o custo | evento_corporativo na linha do tempo do ativo | RIR/P&R IRPF | A |
| R9 | Exercício de opções: prêmio integra custo (titular de call) / valor de venda (lançador de call); casos simétricos para put | Vincular opção, ativo-objeto e negócio de exercício; 4 casos com exemplo. A nota traz só o código da opção + "E" (ex.: KLBNO193E), às vezes + classe; ativo-objeto = 4 letras + classe (ON→3, PN→4, PNA→5, PNB→6, UNT→11); sem classe: classe da abertura da série, depois opcoes.csv | IN RFB 1.585/2015 | M |
| R10 | Opção que vira pó: prêmio pago = perda (titular); recebido = ganho (lançador) no vencimento | Evento automático de vencimento (data como parâmetro, não "hoje") | IN RFB 1.585/2015 | A |
| R11 | Futuros: resultado = soma dos ajustes no mês − custos | Ler "Ajuste de posição"/"Ajuste day trade" da nota BM&F; multiplicador por contrato com vigência | IN RFB 1.585/2015 | M |
| R12 | DARF código 6015; vencimento último dia útil do mês seguinte; valor < R$ 10 acumula | Calendário de feriados + acúmulo | Lei 9.430/1996 art. 68 | A |
| R13 | Atraso: multa 0,33%/dia limitada a 20% + juros Selic (acumulada a partir do mês seguinte ao vencimento + 1% no mês do pagamento) | Tabela Selic mensal editável | Lei 9.430/1996 art. 61 | A |
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
  Sencon ou o myProfit têm equivalente (`docs/referencias/`); sem equivalente, não implementar.

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
- **Transferência de custódia entre corretoras do mesmo titular (ex.: VALE5 Rico → Ágora, 2016):**
  evento próprio que move quantidade e PM de uma conta para outra, sem resultado. A nota continua na
  corretora que a emitiu; a planilha v2 relabelava as notas.
- **Opção flexível (BTG):** fica para uma versão seguinte à v3.0 (parser do layout "Mercadoria: FCA"
  e operação de opção flexível); na v3.0 a extração reconhece e avisa.

## Pontos do v2.02 a corrigir (ver docs/PLANO.md, Diagnóstico)
- Isenção inferida por "IRRF < R$ 1" (falha com 2+ corretoras; pode isentar ETF).
- Sem categoria FII.
- Saldo de IRRF não transportado (L = J − K no próprio mês) — conferir.
- DARF sem mínimo de R$ 10 e vencimento no último dia corrido.
- Taxas de futuros e trocas de ticker fixas no código.

## Aviso
Este documento orienta a implementação; não substitui contador/advogado tributarista. Divergências de
entendimento devem virar parâmetro configurável e decisão explícita do usuário.
