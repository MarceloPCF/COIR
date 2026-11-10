# Manual do Meu Imposto de Renda (MIR) — páginas da RFB

| Campo | Valor |
|---|---|
| Ementa (resumo) | Manual de preenchimento da declaração online (MIR): rendimentos, patrimônio, eventos do patrimônio, resumo; e páginas de "Pagamento > Renda variável". |
| Fonte oficial | https://www.gov.br/receitafederal/pt-br/assuntos/meu-imposto-de-renda/preenchimento/manual-mir e .../pagamento/renda-variavel |
| Texto lido | as 40 páginas do acervo `mir/` (preenchimento: tela inicial, pessoas, rendimentos do trabalho, do capital, outros, decisão judicial, patrimônio, bens financeiros, imóveis e móveis, dívidas, eventos do patrimônio, pagamentos/doações, resumo, entrega, destinação, glossário; pagamento: renda variável, bolsa de valores, cálculo e pagamento, compensações, deduções, demonstrativo, ganho líquido, isenções, retenções, fundos no Brasil) |
| Lido em | 2026-11-10 (páginas atualizadas entre 04/06/2024 e 31/08/2026) |
| Leitura | íntegra das páginas de rendimentos do capital, outros rendimentos, rendimentos do trabalho (dividendos), bens financeiros, eventos do patrimônio, resumo e as páginas de pagamento de renda variável; as de deduções, imóveis, pessoas e doações percorridas (não tocam renda variável) |
| Confiança | média: é orientação da RFB (não norma) e o MIR não dá os **números** dos grupos/códigos de Bens e Direitos nem os códigos das fichas do PGD |
| Temas | ações; isenção 20 mil; FII/Fiagro; ETF; BDR; opções; proventos/JCP; FI-Infra/FIP-IE; IRRF; DARF/multa; fichas do IRPF |

## Vigência e alterações

| Dispositivo | Redação | Vigente de | até | Alterado por |
|---|---|---|---|---|
| páginas de preenchimento | exercício 2026 (publicadas 17/03/2026, atualizadas até 31/08/2026) | 2026 | — | revisão anual da RFB |
| páginas de pagamento (retenções, cálculo) | atualizadas 19/11/2025; isenções 04/06/2024 | — | — | — |

Não há versão histórica: para exercícios anteriores valem o PGD e o P&R do ano.

## Dispositivos relevantes para o COIR

Rendimentos (nomes do MIR; o código do PGD não aparece):
- **Isentos**: "Lucros e dividendos" (pág. rendimentos do trabalho, PJ no Brasil, lucros desde 1996);
  "Incorporação de reservas ao capital / Bonificações em ações" (valor conforme comprovante da empresa ou
  custodiante); "Ganhos líquidos em operações no mercado à vista de ações… até R$ 20.000,00 em cada mês"
  (isenção não alcança day trade, ETF de ações, resgate de FIA, exercício de opção, vencimento/liquidação
  antecipada de termo); "Ganhos líquidos em ouro… até R$ 20.000"; **rendimentos isentos de FII e Fiagro vão
  em "Outros rendimentos isentos"** (com informação complementar obrigatória).
- **Exclusivos/definitivos**: "Juros sobre capital próprio"; "Rendimentos de aplicações financeiras" (valor
  líquido; inclui rendimentos de fundos, distribuição **não isenta** de FII/Fiagro, renda fixa negociada em
  bolsa e operações conjugadas/box/termo de financiamento); "Operações de swap" (perda de swap não compensa
  com outra renda variável).
- **Renda Variável** (ficha mensal): ganhos/perdas por mês em operações comuns, day trade e FII/Fiagro, IRRF
  de comuns, IRRF de day trade e DARF 6015 pago; prejuízo do ano anterior por modalidade. Entram: à vista de
  ações, ouro, termo, opções e futuros (em bolsa e fora, inclusive opções flexíveis), FII/Fiagro em bolsa.
  Não entram: box e outras operações de renda predeterminada, mútuo de ouro, operações isentas (salvo para
  compensar perdas) e FII/Fiagro **fora de bolsa** (vai para o programa de Ganho de Capital).
- O resumo soma o "Imposto retido na fonte - Renda Variável" ao imposto pago da DAA.

Patrimônio (nomes do MIR; sem números de grupo/código):
- Participações societárias: "Ações (inclusive as listadas em Bolsa)" — declarar se custo ≥ R$ 1.000.
- Aplicações e investimentos: "Ativos Negociados em Bolsa no Brasil, Exceto Ações e Fundos" (onde cabem BDR,
  opções — descrição com quantidade, série e vencimento — e direitos) e "Ouro, Ativo Financeiro" — declarar se custo ≥ R$ 1.000.
- Fundos (declarar se saldo > R$ 140): come-cotas; Fiagro; FII; FIA e FMP-FGTS; "FIP - Entidade de
  investimento, FIDC - Entidade de investimento sem tributação periódica, ETF - Entidade de investimento"
  (um item só); FIP-IE e FIP-PD&I; Fundos de Índice de Renda Fixa; "Fundos de Infraestrutura, FIDC e outros
  Fundos de Investimento (alíquota 0%)"; FIEE; Multimercado; Fundos no exterior.
- Outros bens e direitos: "Juros Sobre Capital Próprio a receber" — declarar sempre.
- Eventos do patrimônio: JCP e dividendos podem ser lançados a partir da ação; rendimentos de aplicações e
  "Outros rendimentos isentos" a partir do fundo.

Pagamento (renda variável):
- Cálculo: os custos da compra somam ao custo e os da venda reduzem o valor de venda; o IRRF de 0,005% não
  reduz o valor de venda, só o imposto (exemplo: 1.495,05 − 3,00 = 1.492,05); DARF 6015 até o último dia útil
  do mês seguinte ao do ganho.
- Retenções: 0,005% sobre a venda (comuns) deduz no mês, nos meses seguintes e, se sobrar, na DAA; 1% sobre
  o resultado de day trade deduz só do imposto de day trade, só até dezembro do mesmo ano, e o saldo final
  pode ser pedido em restituição.
- Isenções: ações e ouro até R$ 20 mil de vendas no mês, cada um com seu limite; fora: day trade, ETF de ações,
  resgate de FIA. Exemplo: IRRF de R$ 0,15 não retido (abaixo de R$ 1).

## Onde o COIR usa

- `data/irpf/2025.yaml`, `2026.yaml`, `2027.yaml` — citam o "Manual do MIR" em bens, isentos e exclusivos:
  - isentos: dividendos 09, bonificação 18, lucro até R$ 20 mil 20, FII/Fiagro e reembolso de BTC em 99
    ("Outros"): os nomes conferem com o MIR (o MIR põe FII/Fiagro isento em "Outros rendimentos isentos").
    Os números 09/18/20/99 não estão no MIR (vêm do PGD).
  - exclusivos: JCP 10, ETF RF 06, FI-Infra/FIP-IE 06 ("Rendimentos de aplicações financeiras"): nomes conferem.
  - bens: ações 03.01, BDR/opção/direito 04.04, FII 07.03, Fiagro 07.02, ETF 07.06, FIP-IE 07.07, ETF RF 07.08,
    FI-Infra 07.10, JCP a receber 99.07: os **nomes** conferem com o MIR; os **números** não constam do MIR.
- `data/parametros_fiscais.csv` `prejuizo_mes_isento_compensavel` e R5 citam "Manual MIR 2026": o MIR
  confirma que operações isentas não vão à ficha, "exceto no caso de pretender compensar as perdas apuradas
  com ganhos… sujeitos à incidência do imposto". Confere.
- R6 (saldo de day trade em dezembro só restituição): confere com a página "Retenções".
- `coir/reports/irpf.py:335` (FIP-IE na exclusiva 06 "pelo Manual MIR"): o MIR lista FIP-IE como patrimônio que
  aceita "Rendimentos de aplicações financeiras", não "Outros rendimentos isentos"; confere com a leitura do projeto.

## Pontos em aberto

- **Citação imprecisa** (`data/irpf/2026.yaml` e `2027.yaml`, nota do FIP_IE): "o manual MIR dá só os nomes, na
  ordem 06 ETF, 07 FIP-IE, 08 ETF RF, 10 Infra". O MIR atual não traz números e junta "FIP / FIDC / ETF -
  Entidade de investimento" num item só; a ordem não permite deduzir os códigos. A R20 diz que 07.07 e 07.10 foram
  conferidos no programa: a nota deve citar o programa IRPF 2026, não o MIR. Gravidade baixa.
- O MIR põe a distribuição **não isenta** de FII/Fiagro em "Rendimentos de aplicações financeiras" (exclusiva);
  o COIR só tem a linha isenta (R16: o usuário marca). Sugestão: linha exclusiva para FII não isento.
- Bens abaixo do limite (R$ 1.000 em ações/ativos de bolsa, R$ 140 em fundos) podem ser omitidos; o COIR lista tudo.
