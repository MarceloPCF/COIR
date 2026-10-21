# Fontes fiscais do COIR

Gerado por `uv run python -m tools.fontes_fiscais` a partir das tabelas de `data/`, de
`docs/regras-fiscais.md` e das docstrings dos casos fiscais e do motor. Não edite à mão:
corrija a citação na origem e gere de novo. O documento orienta; não substitui a norma.

| Norma | Usos |
|---|---|
| LC 224/2025 | 8 |
| Lei 6.404 | 3 |
| Lei 8.668/1993 | 1 |
| Lei 8.981/1995 | 10 |
| Lei 9.249/1995 | 15 |
| Lei 9.250 | 3 |
| Lei 9.430/1996 | 17 |
| Lei 9.779/1999 | 1 |
| Lei 9.959/2000 | 5 |
| Lei 11.033/2004 | 25 |
| Lei 11.051/2004 | 5 |
| Lei 12.350/2010 | 1 |
| Lei 13.043/2014 | 13 |
| Lei 14.754/2023 | 3 |
| Lei 15.265/2025 | 10 |
| Lei 15.270/2025 | 10 |
| MP 497/2010 | 5 |
| IN RFB 1.585/2015 | 61 |
| Res. CMN 4.880/2020 | 1 |
| SC Cosit 67/2024 | 1 |
| SC Cosit 145/2021 | 1 |
| SC Cosit 224/2014 | 3 |
| P&R IRPF | 5 |
| P&R IRPF 2026 | 21 |
| Manual MIR 2026 | 3 |
| Programa IRPF | 7 |
| CTN | 1 |
| Súmula STJ 360 | 1 |
| REsp 1.149.022/SP | 1 |
| RIR/2018 | 6 |

## LC 224/2025

- `coir/engine/proventos.py` — _jcp_com_credito: …receber em 31/12 se pago depois), com a alíquota da data do crédito ou do pagamento, o que vier primeiro (LC 224/2025 art. 8º). Crédito depois do pagamento falha alto.
- `coir/engine/proventos.py` — módulo: …retida: 15% até 2025, 17,5% desde 2026, pela data do pagamento ou do crédito, o que vier primeiro (LC 224/2025 art. 8º). Pago num ano com a alíquota do ano anterior = creditado no ano anterior: declarado no ano…
- `data/parametros_fiscais.csv` — aliquota_irrf_jcp: LC 224/2025 art. 8º e art. 14 III (17,5% pago ou creditado desde 01/01/2026)
- `docs/regras-fiscais.md` — Declaração anual (F8, 20/10/2026): a alíquota da data do crédito ou do pagamento, o que vier primeiro (LC 224/2025 art. 8º). A data
- `docs/regras-fiscais.md` — Proventos (F7a, 19/10/2026): alíquota retida (15% até 2025; 17,5% desde 2026, LC 224/2025 art. 8º, pela data do pagamento ou do
- `docs/regras-fiscais.md` — R15: R15 | JCP: IRRF exclusivo 15% até 31/12/2025; 17,5% para JCP pago ou creditado desde 01/01/2026 | Provento com IRRF; alíquota com vigência | LC 224/2025 | A
- `tests/fiscal_cases/test_proventos.py` — módulo: …conferidos à mão, 19/10/2026). Fontes: Lei 9.249/1995 arts. 9º e 10 (JCP exclusivo; dividendos isentos); LC 224/2025 art. 8º e art. 14 III (JCP a 17,5% pago ou creditado desde 01/01/2026, na data do pagamento ou do…
- `tests/fiscal_cases/test_proventos.py` — test_jcp_creditado_em_2025_e_pago_em_2026_usa_a_aliquota_do_credito: LC 224/2025 art. 8º: 17,5% para JCP pago ou creditado desde 01/01/2026, o que vier primeiro. Creditado em 20/12/2025 e pago em 15/04/2026: 15%, declarado em 2025.

## Lei 6.404

- `data/parametros_fiscais.csv` — incorporacao_acoes_tratamento: Lei 6.404 art. 252: custo das ações entregues mantido nas recebidas (TRF1/TRF3/TRF4; decisão do usuário 19/10/2026). A RFB diverge (SC Cosit 224/2014: alienação, ganho de capital no GCAP); a tela de Eventos avisa
- `docs/regras-fiscais.md` — R8: …art. 60 §2º; P&R IRPF (substituição de ações; incorporação de ações); SC Cosit 224/2014 (divergência); Lei 6.404 arts. 227-229 e 252 | A
- `tests/fiscal_cases/test_eventos_societarios.py` — módulo: …(substituição de ações): não é alienação, a data de aquisição é a original. - Incorporação de ações (Lei 6.404 art. 252): custo mantido (TRF1, TRF3, TRF4; decisão do usuário de 19/10/2026); a RFB diverge (SC Cosit…

## Lei 8.668/1993

- `data/parametros_fiscais.csv` — aliquota_fii: Lei 8.668/1993 art. 18 II (red. Lei 9.779/1999); inclui day trade de FII

## Lei 8.981/1995

- `coir/engine/apuracao.py` — módulo: …sem prazo (IN RFB 1.585/2015); prejuízo de mês isento compensa se `prejuizo_mes_isento_compensavel` (Lei 8.981/1995 art. 72 §8º). - IRRF (0,005% e 1%): deduz do imposto do mês; o saldo passa aos meses seguintes do…
- `coir/engine/proventos.py` — módulo: …Seguridade); não é lucro distribuído (Lei 9.249 art. 10): tributação exclusiva na fonte como renda fixa (Lei 8.981/1995 art. 65 §4º c; Lei 11.033/2004 art. 1º), como as fontes retêm (22,5% nos dados reais); bruto =…
- `data/irpf/2025.yaml` — exclusivos: Programa IRPF 2026, ficha Rendimentos sujeitos à tributação exclusiva/definitiva; Lei 9.249/1995 art. 9º; Lei 13.043/2014 art. 2º; Lei 8.981/1995 art. 65 §4º c
- `data/irpf/2026.yaml` — exclusivos: Programa IRPF 2026, ficha Rendimentos sujeitos à tributação exclusiva/definitiva; Lei 9.249/1995 art. 9º; Lei 13.043/2014 art. 2º; Lei 8.981/1995 art. 65 §4º c
- `data/irpf/2027.yaml` — exclusivos: Programa IRPF 2026, ficha Rendimentos sujeitos à tributação exclusiva/definitiva; Lei 9.249/1995 art. 9º; Lei 13.043/2014 art. 2º; Lei 8.981/1995 art. 65 §4º c
- `data/parametros_fiscais.csv` — prejuizo_mes_isento_compensavel: Lei 8.981/1995 art. 72 §8º; Manual MIR 2026; decisão do usuário (14/10/2026)
- `docs/regras-fiscais.md` — Proventos (F7a, 19/10/2026): atualização do provento pela Selic; tributação exclusiva como renda fixa (Lei 8.981/1995 art. 65
- `docs/regras-fiscais.md` — R5: R5 | Prejuízo de mês isento continua compensável | Parâmetro `prejuizo_mes_isento_compensavel = true` | Lei 8.981/1995 art. 72 §8º; Manual MIR 2026; decisão do usuário (14/10/2026) | A
- `tests/fiscal_cases/test_casos_fiscais.py` — test_t1_prejuizo_em_mes_isento_compensa: R5: prejuízo de ações em mês isento continua compensável (Lei 8.981/1995 art. 72 §8º; Manual MIR 2026; decisão do usuário). Mar/24: ações -1.200 (vendas 15.000), opções…
- `tests/fiscal_cases/test_proventos.py` — test_atualizacao_do_provento_pela_selic_e_tributacao_exclusiva: …por ação, pago no mesmo dia (B3). Para 100 ações: bruto 5,684 -> 5,68; retidos 22,5% como renda fixa (Lei 8.981/1995 art. 65 §4º c; Lei 11.033/2004 art. 1º) -> líquido 4,40, IRRF 1,28. Não é dividendo: fora da Lei…

## Lei 9.249/1995

- `coir/engine/proventos.py` — _jcp_com_credito: …empresa ou data de aprovação no pacote da B3): declarado no ano do crédito (P&R IRPF 2026 pergunta 320; Lei 9.249/1995 art. 9º §2º; a receber em 31/12 se pago depois), com a alíquota da data do crédito ou do…
- `coir/engine/proventos.py` — módulo: …da B3 (decisões do usuário de 19/10/2026). - Dividendo de empresa brasileira (ação, unit): isento (Lei 9.249/1995 art. 10). De BDR, ETF ou FII: fora da v3.0, listado com aviso (BDR: Lei 14.754/2023, não é…
- `data/irpf/2025.yaml` — exclusivos: Programa IRPF 2026, ficha Rendimentos sujeitos à tributação exclusiva/definitiva; Lei 9.249/1995 art. 9º; Lei 13.043/2014 art. 2º; Lei 8.981/1995 art. 65 §4º c
- `data/irpf/2025.yaml` — isentos: Programa IRPF 2026, ficha Rendimentos isentos e não tributáveis; Manual do MIR (rendimentos do capital e do patrimônio); Lei 9.249/1995 art. 10; Lei 11.033/2004 art. 3º I e III; IN RFB 1.585/2015 arts. 58 §1º e 74
- `data/irpf/2026.yaml` — exclusivos: Programa IRPF 2026, ficha Rendimentos sujeitos à tributação exclusiva/definitiva; Lei 9.249/1995 art. 9º; Lei 13.043/2014 art. 2º; Lei 8.981/1995 art. 65 §4º c
- `data/irpf/2026.yaml` — isentos: Programa IRPF 2026, ficha Rendimentos isentos e não tributáveis; Manual do MIR (rendimentos do capital e do patrimônio); Lei 9.249/1995 art. 10; Lei 11.033/2004 art. 3º I e III; IN RFB 1.585/2015 arts. 58 §1º e 74
- `data/irpf/2027.yaml` — exclusivos: Programa IRPF 2026, ficha Rendimentos sujeitos à tributação exclusiva/definitiva; Lei 9.249/1995 art. 9º; Lei 13.043/2014 art. 2º; Lei 8.981/1995 art. 65 §4º c
- `data/irpf/2027.yaml` — isentos: Programa IRPF 2026, ficha Rendimentos isentos e não tributáveis; Manual do MIR (rendimentos do capital e do patrimônio); Lei 9.249/1995 art. 10; Lei 11.033/2004 art. 3º I e III; IN RFB 1.585/2015 arts. 58 §1º e 74
- `data/parametros_fiscais.csv` — aliquota_irrf_jcp: Lei 9.249/1995 art. 9º §2º (15% na data do pagamento ou crédito)
- `docs/regras-fiscais.md` — Declaração anual (F8, 20/10/2026): JCP: declarado no ano do crédito (P&R IRPF 2026 pergunta 320; Lei 9.249/1995 art. 9º §2º), com
- `docs/regras-fiscais.md` — Proventos (F7a, 19/10/2026): Dividendo de ação ou unit: isento (Lei 9.249/1995 art. 10), código 09. BDR, ETF e FII com
- `tests/fiscal_cases/test_casos_fiscais.py` — test_bonificacao_ajusta_a_quantidade_da_serie_sem_custo: R8 (Lei 9.249/1995 art. 10 parágrafo único e IN RFB 1.585/2015 art. 58: o custo atribuído da bonificação vai para as…
- `tests/fiscal_cases/test_irpf_declaracao.py` — módulo: …63 §8º III (saldo de IRRF de dezembro na declaração) e 74; Lei 11.033/2004 art. 3º I (código 20); Lei 9.249/1995 arts. 9º e 10; Manual do MIR (Bens e Direitos, rendimentos do capital); data/irpf/2026.yaml.
- `tests/fiscal_cases/test_proventos.py` — módulo: Proventos a partir da movimentação da B3 (casos fictícios conferidos à mão, 19/10/2026). Fontes: Lei 9.249/1995 arts. 9º e 10 (JCP exclusivo; dividendos isentos); LC 224/2025 art. 8º e art. 14 III (JCP a 17,5%…
- `tests/fiscal_cases/test_proventos.py` — test_jcp_creditado_no_ano_anterior_com_a_mesma_aliquota: …ela (pacote da B3 ou informe), declarado em 2024 e a receber em 31/12/2024 (P&R IRPF 2026 pergunta 320; Lei 9.249/1995 art. 9º §2º).

## Lei 9.250

- `data/parametros_fiscais.csv` — aliquota_irrf_dividendos_acima_limite: Lei 15.270/2025 (Lei 9.250 art. 6º-A caput e §§1º e 2º)
- `data/parametros_fiscais.csv` — dividendos_limite_mensal_mesma_pj: Lei 15.270/2025 (Lei 9.250 art. 6º-A): acima de R$ 50 mil no mês da mesma PJ há retenção de 10% sobre o total
- `tests/fiscal_cases/test_proventos.py` — módulo: …e 21 (reembolso de BTC); IN 1.585 art. 58 §8º (restituição de capital reduz o custo); Lei 15.270/2025 (Lei 9.250 art. 6º-A: IRRF de 10% sobre o total dos dividendos da mesma PJ no mês acima de R$ 50 mil). Decisões do…

## Lei 9.430/1996

- `coir/engine/apuracao.py` — módulo: …cada uma consumiu (`fichas.py`). - DARF 6015: abaixo de R$ 10 acumula, inclusive para o ano seguinte (Lei 9.430/1996 art. 68 §1º); vence no último dia útil do mês seguinte (RIR/2018 art. 915).
- `coir/engine/calendario.py` — módulo: …por parâmetro). - DARF: último dia útil (bancário) do mês seguinte ao da apuração (RIR/2018 art. 915; Lei 9.430/1996 art. 68 §1º para o acúmulo abaixo do mínimo). - Opções sobre ações: regra de…
- `coir/engine/darf.py` — módulo: Multa e juros de mora da DARF paga em atraso (Lei 9.430/1996 art. 61). Juros (art. 61 §3º c/c art. 5º §3º): a Selic acumulada de cada mês, da do mês seguinte…
- `data/parametros_fiscais.csv` — darf_minimo: Lei 9.430/1996 art. 68 caput e §1º (abaixo do mínimo acumula)
- `data/parametros_fiscais.csv` — juros_mora_mes_pagamento: Lei 9.430/1996 art. 61 §3º c/c art. 5º §3º (1% no mês do pagamento; Selic acumulada nos meses entre o vencimento e o pagamento)
- `data/parametros_fiscais.csv` — multa_mora_dia: Lei 9.430/1996 art. 61 caput (0,33% por dia de atraso sobre o principal)
- `data/parametros_fiscais.csv` — multa_mora_teto: Lei 9.430/1996 art. 61 §2º (limite de 20%)
- `docs/regras-fiscais.md` — DARF paga em atraso: multa de mora (F6, 17/10/2026): 0,33% por dia de atraso sobre o principal, com teto de 20% (Lei 9.430/1996 art. 61 caput e
- `docs/regras-fiscais.md` — R12: …(o último dia útil do ano não é útil: Res. CMN 4.880/2020 art. 3º) + acúmulo | RIR/2018 art. 915; Lei 9.430/1996 art. 68 caput e §1º | A
- `docs/regras-fiscais.md` — R13: R13 | Atraso: multa 0,33%/dia limitada a 20% + juros Selic (do mês seguinte ao vencimento até o mês anterior ao pagamento + 1% no mês do pagamento) | `data/selic_mensal.csv` (BCB SGS 4390) | Lei 9.430/1996 art. 61 | A
- `tests/fiscal_cases/test_casos_fiscais.py` — test_t4_darf_abaixo_de_10_acumula_e_vira_o_ano: R12: DARF abaixo de R$ 10 acumula, inclusive para o ano seguinte (Lei 9.430/1996 art. 68 §1º); vence no último dia útil do mês seguinte (RIR/2018 art. 915). Dez/23 7,00 + Jan/24 4,00 = 11,00, vencimento 29/02/2024.
- `tests/fiscal_cases/test_irrf_fichas.py` — test_darf_acumulada_marca_os_meses_que_ela_pagou: Lei 9.430/1996 art. 68 §1º: fevereiro/2026 com imposto 6,00 (ganho 40 em ações; vendas acima de 20.000) acumula;…
- `tests/fiscal_cases/test_irrf_fichas.py` — test_darf_de_janeiro_paga_o_acumulado_de_dezembro: Lei 9.430/1996 art. 68 §1º: dezembro/2025 com imposto 6,00 acumula e entra na DARF de janeiro/2026 (6,00 + 8,25 = 14,25): o "pago" de dezembro é 6,00 no relatório de 2025.
- `tests/fiscal_cases/test_juros_mora.py` — módulo: Juros de mora da DARF 6015 paga em atraso (números conferidos à mão em 19/10/2026). Fontes: Lei 9.430/1996 art. 61 §3º c/c art. 5º §3º (Selic acumulada mensalmente + 1% no mês do pagamento) e art. 68 §1º;…
- `tests/fiscal_cases/test_multa_mora.py` — módulo: Multa de mora da DARF 6015 paga em atraso (números conferidos à mão em 17/10/2026). Fontes: Lei 9.430/1996 art. 61 caput (0,33% por dia), §1º (contagem a partir do dia seguinte ao vencimento até o…
- `tests/fiscal_cases/test_multa_mora.py` — test_contagem_literal_da_lei_por_parametro: Lei 9.430/1996 art. 61 §1º: do dia seguinte ao vencimento, em dias corridos.
- `tests/fiscal_cases/test_multa_mora.py` — test_multa_de_mora_como_a_receita: Lei 9.430/1996 art. 61; contagem e centavos como decididos (rfb, ROUND_DOWN).

## Lei 9.779/1999

- `data/parametros_fiscais.csv` — aliquota_fii: Lei 8.668/1993 art. 18 II (red. Lei 9.779/1999); inclui day trade de FII

## Lei 9.959/2000

- `data/parametros_fiscais.csv` — day_trade_mesma_corretora: Lei 9.959/2000 art. 8º §1º I a (redação original): day trade em qualquer corretora
- `docs/regras-fiscais.md` — Decisões da F4 (motor fiscal, 14-15/10/2026): trade (Lei 9.959/2000 art. 8º §§6º e 7º).
- `docs/regras-fiscais.md` — Pendências do motor (revisão final da F4, 16/10/2026): intermediadora" (Lei 9.959/2000 art. 8º §1º) compara a corretora da nota: XP, Clear e Rico são
- `docs/regras-fiscais.md` — R6: …meses seguintes do ano; em dezembro, o de 0,005% vai para a declaração anual e o de DT só restituição | Lei 9.959/2000 art. 8º §§4º, 8º; IN RFB 1.585/2015 arts. 63 §§8º e 9º e 65 §§8º, 9º e 12; P&R IRPF 2026 perguntas…
- `tests/fiscal_cases/test_casos_fiscais.py` — test_irrf_de_day_trade_e_consumido_primeiro: IRRF em saldo único (Lei 9.959/2000 art. 8º §4º; IN 1.585 art. 65 §8º; P&R 714): abate o imposto total do mês. Consome primeiro o de…

## Lei 11.033/2004

- `coir/engine/apuracao.py` — módulo: …e limites vêm de data/parametros_fiscais.csv, com a fonte de cada um): - Alíquotas: 15% comum (Lei 11.033/2004 art. 2º II), 20% day trade e 20% FII (IN RFB 1.585/2015, arts. 56 a 65). - Isenção: ganho líquido…
- `coir/engine/fichas.py` — módulo: …IRPF (comuns/day trade e FII), com a DARF no saldo único. A DARF mensal deduz o IRRF de um saldo único (Lei 11.033/2004 art. 2º §7º; IN RFB 1.585/2015 arts. 63 §8º e 65 §8º; decisão do usuário de 16/10/2026: separar…
- `coir/engine/proventos.py` — módulo: …(1 - alíquota), centavos arredondados (estimativa: vale o informe). - Rendimento de FII/Fiagro: isento (Lei 11.033/2004 art. 3º III). "Rendimento" em ação pago junto com um dividendo ou JCP do mesmo código e dia: a…
- `data/irpf/2025.yaml` — isentos: Programa IRPF 2026, ficha Rendimentos isentos e não tributáveis; Manual do MIR (rendimentos do capital e do patrimônio); Lei 9.249/1995 art. 10; Lei 11.033/2004 art. 3º I e III; IN RFB 1.585/2015 arts. 58 §1º e 74
- `data/irpf/2026.yaml` — isentos: Programa IRPF 2026, ficha Rendimentos isentos e não tributáveis; Manual do MIR (rendimentos do capital e do patrimônio); Lei 9.249/1995 art. 10; Lei 11.033/2004 art. 3º I e III; IN RFB 1.585/2015 arts. 58 §1º e 74
- `data/irpf/2027.yaml` — isentos: Programa IRPF 2026, ficha Rendimentos isentos e não tributáveis; Manual do MIR (rendimentos do capital e do patrimônio); Lei 9.249/1995 art. 10; Lei 11.033/2004 art. 3º I e III; IN RFB 1.585/2015 arts. 58 §1º e 74
- `data/parametros_fiscais.csv` — aliquota_comum: Lei 11.033/2004 art. 2º II; RIR/2018 art. 839
- `data/parametros_fiscais.csv` — aliquota_day_trade: Lei 11.033/2004 art. 2º I
- `data/parametros_fiscais.csv` — limite_isencao_acoes: Lei 11.033/2004 art. 3º I (valor bruto das alienações de ações à vista no mês)
- `docs/regras-fiscais.md` — Decisões da F4 (motor fiscal, 14-15/10/2026): saldo único (a lei e a IN não separam o IRRF de FII; Lei 11.033/2004 art. 2º §7º). As fichas do
- `docs/regras-fiscais.md` — Proventos (F7a, 19/10/2026): Rendimento de FII/Fiagro: isento (Lei 11.033 art. 3º III).
- `docs/regras-fiscais.md` — Proventos (F7a, 19/10/2026): §4º c; Lei 11.033 art. 1º), como as fontes retêm (22,5% nos dados reais); código 06 ou 12 conforme
- `docs/regras-fiscais.md` — R16: R16 | Rendimentos de FII isentos se o fundo cumpre requisitos | Usuário marca conforme informe do fundo | Lei 11.033/2004 art. 3º (alt. Lei 14.754/2023) | M
- `docs/regras-fiscais.md` — R4: …e venda por exercício incluídos); fora: units, ETF, BDR, FII, opções, futuros, termo e day trade | Lei 11.033/2004 art. 3º I | A
- `docs/regras-fiscais.md` — Relatórios (F5, 16/10/2026): abatimento (day trade antes da Lei 11.033) na ficha de comuns.
- `docs/regras-fiscais.md` — Relatórios (F5, 16/10/2026): duas partes, day trade e Lei 11.033, com mês, meses anteriores e a compensar; o de day trade é
- `tests/fiscal_cases/test_casos_fiscais.py` — test_t2_limite_pelo_valor_bruto_das_vendas: R4: o limite de R$ 20 mil é sobre o valor bruto das alienações (Lei 11.033/2004 art. 3º I); custos de venda não reduzem o valor para o limite. 20.000,00 com custos 10 é isento; 20.000,01 é tributável.
- `tests/fiscal_cases/test_eventos_societarios.py` — módulo: …(decisão do usuário de 19/10/2026: entra no limite de R$ 20 mil do mês do crédito, Lei 11.033/2004 art. 3º I).
- `tests/fiscal_cases/test_irpf_declaracao.py` — módulo: …arts. 58 §1º (bonificação: custo atribuído), 63 §8º III (saldo de IRRF de dezembro na declaração) e 74; Lei 11.033/2004 art. 3º I (código 20); Lei 9.249/1995 arts. 9º e 10; Manual do MIR (Bens e Direitos, rendimentos…
- `tests/fiscal_cases/test_irpf_declaracao.py` — test_ficha_mes_a_mes_e_imposto_pago_de_dezembro: …1.000,00. Dezembro: V 4.000 PETR4 a 30 (resultado 0) com IRRF 6,00 que sobra: ficha "Imposto pago", Lei 11.033 = 6,00.
- `tests/fiscal_cases/test_irpf_linhas.py` — test_venda_no_dia_acima_do_termo_sai_das_antigas_pelo_pm: …das antigas pelo PM: 5.500,00 - 4.500,00 = 1.000,00, isento (vendas de 16.500,00, abaixo de R$ 20 mil; Lei 11.033/2004 art. 3º I). Imposto 150,00; posição zerada (decisão do usuário de 21/10/2026).
- `tests/fiscal_cases/test_irrf_fichas.py` — módulo: …separar o IRRF de FII só na declaração). Casos do especialista fiscal, conferidos à mão. Fontes: Lei 11.033/2004 art. 2º §§1º, 2º, 5º e 7º; IN RFB 1.585/2015 arts. 63 §8º e 65 §§8º e 9º; fichas "Renda variável…
- `tests/fiscal_cases/test_irrf_fichas.py` — test_base_da_reparticao_usa_o_bruto_e_ignora_venda_por_exercicio: Lei 11.033/2004 art. 2º §2º: venda por exercício fora da base do 0,005%. Agosto/2026, uma nota: V 1.000 AAAA3 a…
- `tests/fiscal_cases/test_proventos.py` — módulo: …8º e art. 14 III (JCP a 17,5% pago ou creditado desde 01/01/2026, na data do pagamento ou do crédito); Lei 11.033/2004 art. 3º III (rendimentos de FII isentos); IN RFB 1.585/2015 art. 74 e Lei 15.265/2025 arts. 20 e…
- `tests/fiscal_cases/test_proventos.py` — test_atualizacao_do_provento_pela_selic_e_tributacao_exclusiva: …(B3). Para 100 ações: bruto 5,684 -> 5,68; retidos 22,5% como renda fixa (Lei 8.981/1995 art. 65 §4º c; Lei 11.033/2004 art. 1º) -> líquido 4,40, IRRF 1,28. Não é dividendo: fora da Lei 15.270.

## Lei 11.051/2004

- `coir/engine/futuros.py` — módulo: Futuros (índice, dólar, commodities): resultado no encerramento da posição. Lei 11.051/2004 art. 32 (caput e §2º II): os resultados em mercados de liquidação futura sujeitos a ajustes…
- `docs/regras-fiscais.md` — Decisões da F4 (motor fiscal, 14-15/10/2026): Futuros no encerramento (Lei 11.051/2004 art. 32), encerramento parcial pelo preço médio.
- `docs/regras-fiscais.md` — R11: …no dia | Ajustes da posição aberta acumulados por contrato e realizados no encerramento ou vencimento | Lei 11.051/2004 art. 32 (caput e §2º II); IN RFB 1.585/2015 art. 61; P&R IRPF 2026 pergunta 728 | A
- `tests/fiscal_cases/test_casos_fiscais.py` — test_futuros_inversao_de_posicao_e_falhas_altas: Lei 11.051/2004 art. 32: comprado em 1, a venda de 3 encerra 1 (resultado: ajuste acumulado 10 + 1/3 do ajuste do…
- `tests/fiscal_cases/test_casos_fiscais.py` — test_futuros_resultado_no_encerramento_da_posicao: R11: Lei 11.051/2004 art. 32; IN 1.585 art. 61; P&R 728 (resultado na liquidação ou encerramento). Compra de 1 WINM24…

## Lei 12.350/2010

- `data/parametros_fiscais.csv` — day_trade_mesma_corretora: MP 497/2010 art. 21 (conv. Lei 12.350/2010): mesma corretora

## Lei 13.043/2014

- `coir/engine/proventos.py` — _origem_reembolso: …é a quantidade emprestada vezes o valor por ação do provento do mesmo dia e empresa, menos o IR do JCP (Lei 13.043/2014 art. 7º §1º II; Lei 15.265/2025 art. 20). Só uma origem com quantidade inteira identifica;…
- `data/irpf/2025.yaml` — exclusivos: Programa IRPF 2026, ficha Rendimentos sujeitos à tributação exclusiva/definitiva; Lei 9.249/1995 art. 9º; Lei 13.043/2014 art. 2º; Lei 8.981/1995 art. 65 §4º c
- `data/irpf/2026.yaml` — exclusivos: Programa IRPF 2026, ficha Rendimentos sujeitos à tributação exclusiva/definitiva; Lei 9.249/1995 art. 9º; Lei 13.043/2014 art. 2º; Lei 8.981/1995 art. 65 §4º c
- `data/irpf/2027.yaml` — exclusivos: Programa IRPF 2026, ficha Rendimentos sujeitos à tributação exclusiva/definitiva; Lei 9.249/1995 art. 9º; Lei 13.043/2014 art. 2º; Lei 8.981/1995 art. 65 §4º c
- `data/isencao_pme.csv` — AGRO3: Lei 13.043/2014 art. 16; IN RFB 1.585/2015 arts. 66 a 69-A; lista B3 (AGRO3); oferta subsequente precificada em 04/02/2021
- `data/isencao_pme.csv` — PRIO3: Lei 13.043/2014 art. 16; IN RFB 1.585/2015 arts. 66 a 69-A; lista B3 (HRT Participações, HRTP3); RI PRIO: isentas as compradas entre 10/07/2014 e 29/01/2021 (oferta subsequente de jan/2021)
- `data/parametros_fiscais.csv` — reembolso_btc_dividendos_isento: IN RFB 1.585/2015 art. 74 §1º I a (Lei 13.043/2014 art. 7º)
- `docs/regras-fiscais.md` — Decisões da F4 (motor fiscal, 14-15/10/2026): **ETF de renda fixa:** fora da apuração e do DARF; ganho listado para a DIRPF (Lei 13.043/2014).
- `docs/regras-fiscais.md` — Decisões da F4 (motor fiscal, 14-15/10/2026): Isenção de PME (Lei 13.043/2014 art. 16; IN arts. 59 III e 66 a 69-A), `data/isencao_pme.csv`:
- `docs/regras-fiscais.md` — Proventos (F7a, 19/10/2026): Reembolso de BTC ao doador: isento até 2025 (Lei 13.043/2014 art. 7º, reproduzido na IN 1.585
- `tests/fiscal_cases/test_casos_fiscais.py` — test_etf_de_renda_fixa_fica_fora_da_apuracao: Lei 13.043/2014 arts. 2º a 4º: ETF de renda fixa tem tributação exclusiva na fonte; o ganho vai para a lista da DIRPF, não para o DARF nem para a isenção.
- `tests/fiscal_cases/test_casos_fiscais.py` — test_pme_ganho_isento_e_fora_do_limite: Lei 13.043/2014 art. 16; IN 1.585 arts. 66 e 67 §3º: ganho na venda de ações de PME é isento e o valor não entra…
- `tests/fiscal_cases/test_proventos.py` — test_reembolso_de_jcp_identificado_pelo_valor_por_acao: …o reembolso é a quantidade emprestada vezes o valor por ação do provento do dia, menos o IR do JCP (Lei 13.043/2014 art. 7º §1º II). JCP de 0,520530743 por ação pago em 14/03/2025 (pacote da B3), 1.000 ações…

## Lei 14.754/2023

- `coir/engine/proventos.py` — módulo: …(ação, unit): isento (Lei 9.249/1995 art. 10). De BDR, ETF ou FII: fora da v3.0, listado com aviso (BDR: Lei 14.754/2023, não é isento). - JCP: tributação exclusiva; a B3 credita o líquido e o preço unitário é o bruto…
- `docs/regras-fiscais.md` — Proventos (F7a, 19/10/2026): "Dividendo": fora da v3.0, listados com aviso (BDR não é isento, Lei 14.754/2023).
- `docs/regras-fiscais.md` — R16: R16 | Rendimentos de FII isentos se o fundo cumpre requisitos | Usuário marca conforme informe do fundo | Lei 11.033/2004 art. 3º (alt. Lei 14.754/2023) | M

## Lei 15.265/2025

- `coir/engine/proventos.py` — _origem_reembolso: …o valor por ação do provento do mesmo dia e empresa, menos o IR do JCP (Lei 13.043/2014 art. 7º §1º II; Lei 15.265/2025 art. 20). Só uma origem com quantidade inteira identifica; nenhuma ou mais de uma: ("", 0), com…
- `coir/engine/proventos.py` — módulo: …com aviso. - Reembolso de BTC (doador): isento até 2025 (IN RFB 1.585/2015 art. 74); desde 2026 a Lei 15.265/2025 é omissa para dividendos: isento por equivalência, com aviso (parâmetro). - Restituição de…
- `data/parametros_fiscais.csv` — reembolso_btc_dividendos_isento: Lei 15.265/2025 arts. 20 e 21 omissos para dividendos: isento por equivalência (decisão do usuário 19/10/2026; rever se a RFB se manifestar)
- `data/parametros_fiscais.csv` — reembolso_btc_norma_omissa: Lei 15.265/2025 arts. 20 e 21 omissos para dividendos: a tela de proventos avisa (isento por equivalência)
- `docs/regras-fiscais.md` — Proventos (F7a, 19/10/2026): 2025 e, desde 2026, sem incidência expressa (Lei 15.265 art. 21 I). Declarado no ano do
- `docs/regras-fiscais.md` — Proventos (F7a, 19/10/2026): art. 74); desde 2026 a Lei 15.265/2025 é omissa para dividendos: isento por equivalência, com
- `docs/regras-fiscais.md` — Proventos (F7a, 19/10/2026): reteria (art. 7º §1º II; Lei 15.265 art. 20), mas o desconto não é IRRF do doador; é isento até
- `tests/fiscal_cases/test_proventos.py` — módulo: …ou do crédito); Lei 11.033/2004 art. 3º III (rendimentos de FII isentos); IN RFB 1.585/2015 art. 74 e Lei 15.265/2025 arts. 20 e 21 (reembolso de BTC); IN 1.585 art. 58 §8º (restituição de capital reduz o custo);…
- `tests/fiscal_cases/test_proventos.py` — test_p4_reembolso_de_btc_isento_com_aviso_desde_2026: Doador de 1.000 ações: reembolso de dividendo de 1.200,00. Em 2025 isento (IN 1.585 art. 74); em 2026 isento por equivalência (Lei 15.265/2025 omissa), com aviso; o PM não muda.
- `tests/fiscal_cases/test_proventos.py` — test_reembolso_de_jcp_identificado_pelo_valor_por_acao: …(não é dividendo de 2,141847479: 206,57 ações). Isento, no ano do recebimento, sem 99.07; desde 2026, Lei 15.265/2025 art. 21 I.

## Lei 15.270/2025

- `coir/engine/proventos.py` — módulo: …ignoradas, mas cada par tem de fechar (senão, aviso). Débito de provento fora desse padrão falha alto. - Lei 15.270/2025 (desde 2026): dividendos da mesma empresa no mês acima de R$ 50 mil têm IRRF de 10% sobre o…
- `data/parametros_fiscais.csv` — aliquota_irrf_dividendos_acima_limite: Lei 15.270/2025 (Lei 9.250 art. 6º-A caput e §§1º e 2º)
- `data/parametros_fiscais.csv` — dividendos_limite_mensal_mesma_pj: Lei 15.270/2025 (Lei 9.250 art. 6º-A): acima de R$ 50 mil no mês da mesma PJ há retenção de 10% sobre o total
- `docs/regras-fiscais.md` — Proventos (F7a, 19/10/2026): Lei 15.270/2025 (desde 2026): só aviso, somando os dividendos líquidos da mesma empresa (raiz do
- `docs/regras-fiscais.md` — Proventos (F7a, 19/10/2026): com o fator 1 - alíquota no JCP; sem origem única, aviso. Desde 2026 o aviso da Lei 15.270
- `docs/regras-fiscais.md` — R14: R14 | Dividendos > R$ 50 mil/mês da mesma PJ: IRRF 10% sobre o total; IRPFM para renda anual > R$ 600 mil | Registrar IRRF do provento; relatório informativo (IRPFM fora do escopo v3.0) | Lei 15.270/2025 | A
- `docs/regras-fiscais.md` — Regras fiscais — renda variável pessoa física (base para coir/engine): anteriores. Novidade vigente desde 01/01/2026: Lei 15.270/2025 (dividendos e IRPF mínimo).
- `tests/fiscal_cases/test_proventos.py` — módulo: …arts. 20 e 21 (reembolso de BTC); IN 1.585 art. 58 §8º (restituição de capital reduz o custo); Lei 15.270/2025 (Lei 9.250 art. 6º-A: IRRF de 10% sobre o total dos dividendos da mesma PJ no mês acima de R$ 50…
- `tests/fiscal_cases/test_proventos.py` — test_atualizacao_do_provento_pela_selic_e_tributacao_exclusiva: …8.981/1995 art. 65 §4º c; Lei 11.033/2004 art. 1º) -> líquido 4,40, IRRF 1,28. Não é dividendo: fora da Lei 15.270.
- `tests/fiscal_cases/test_proventos.py` — test_reembolso_de_dividendo_pela_movimentacao_e_lei_15270: …(como JCP a 15% seriam 588,24 ações: não). Com outro dividendo de 5.000,00 no mês, em 2026 o aviso da Lei 15.270 soma os três (55.000,00 no mês, acima de R$ 50 mil; decisão do usuário de 21/10/2026).

## MP 497/2010

- `coir/engine/movimentos.py` — módulo: …(pó). Fontes: - Day trade: IN RFB 1.585/2015 art. 65 (mesmo ativo, mesmo dia, mesma corretora desde a MP 497/2010; antes, qualquer corretora); §13: exercício de opção não é day trade. - Exercício (art. 58 §§):…
- `data/parametros_fiscais.csv` — day_trade_mesma_corretora: MP 497/2010 art. 21 (conv. Lei 12.350/2010): mesma corretora
- `docs/regras-fiscais.md` — Decisões da F4 (motor fiscal, 14-15/10/2026): 28/07/2010 conta em qualquer corretora (MP 497/2010 art. 21). Exercício de opção não é day trade
- `tests/fiscal_cases/test_casos_fiscais.py` — test_day_trade_entre_corretoras_antes_de_2010_fica_sem_conta_de_origem: MP 497/2010: antes de 28/07/2010 o day trade junta corretoras; o movimento fica sem conta e sem corretora (a origem lista as duas operações). O pó de uma série comprada numa conta só fica nessa conta.
- `tests/fiscal_cases/test_casos_fiscais.py` — test_t7_day_trade_em_corretoras_diferentes_so_conta_antes_de_2010: MP 497/2010: desde 28/07/2010 o day trade é na mesma corretora; antes, em qualquer.

## IN RFB 1.585/2015

- `coir/engine/apuracao.py` — módulo: …a fonte de cada um): - Alíquotas: 15% comum (Lei 11.033/2004 art. 2º II), 20% day trade e 20% FII (IN RFB 1.585/2015, arts. 56 a 65). - Isenção: ganho líquido em ações no mercado à vista quando as alienações do mês…
- `coir/engine/fichas.py` — módulo: …com a DARF no saldo único. A DARF mensal deduz o IRRF de um saldo único (Lei 11.033/2004 art. 2º §7º; IN RFB 1.585/2015 arts. 63 §8º e 65 §8º; decisão do usuário de 16/10/2026: separar só na declaração). O programa…
- `coir/engine/futuros.py` — módulo: …"serão reconhecidos por ocasião da liquidação do contrato, cessão ou encerramento da posição". IN RFB 1.585/2015 art. 61 e P&R IRPF 2026 pergunta 728: resultado de cada contrato = soma algébrica dos ajustes…
- `coir/engine/movimentos.py` — _cisao: Cisão: o percentual do custo de A divulgado pela companhia vai para B, sem resultado (IN RFB 1.585/2015 art. 58 §6º: na proporção fixada pela assembleia). A quantidade de A não muda. Os eventos das…
- `coir/engine/movimentos.py` — _fatiar: IN RFB 1.585/2015 art. 65 §3º: casa pela ordem (1º negócio de compra com o 1º de venda). Os primeiros negócios até a…
- `coir/engine/movimentos.py` — _incorporacao: …fusão, conversão ou ajuste de série: a quantidade de A sai sem resultado e B entra com o mesmo custo (IN RFB 1.585/2015 art. 58 §6º; ajuste de série da B3: sem efeito fiscal, o custo é o dos prêmios, art. 60 §2º). A…
- `coir/engine/movimentos.py` — _leilao_fracao: Fração vendida em leilão pela companhia: alienação pelo valor líquido recebido, custo pelo PM (IN RFB 1.585/2015 art. 58 caput). Sem norma sobre a isenção: parâmetro leilao_fracao_isentavel (decisão do usuário…
- `coir/engine/movimentos.py` — módulo: …pelo preço do próprio negócio); compras; exercícios; vendas; vencimento (pó). Fontes: - Day trade: IN RFB 1.585/2015 art. 65 (mesmo ativo, mesmo dia, mesma corretora desde a MP 497/2010; antes, qualquer corretora);…
- `coir/engine/posicao.py` — módulo: Posição de um ativo pelo custo médio ponderado (IN RFB 1.585/2015 art. 58: custo médio de aquisição; custos da compra integram o custo, os da venda reduzem a…
- `coir/engine/proventos.py` — módulo: …em ação sem esse par: fora, listado com aviso. - Reembolso de BTC (doador): isento até 2025 (IN RFB 1.585/2015 art. 74); desde 2026 a Lei 15.265/2025 é omissa para dividendos: isento por equivalência, com…
- `coir/engine/rateio.py` — módulo: Rateio dos custos da nota pelo valor das operações (IN RFB 1.585/2015 art. 58 §2º: custos de corretagem e taxas integram o custo de aquisição e reduzem o valor de…
- `data/irpf/2025.yaml` — imposto_pago: Programa IRPF 2026, ficha Imposto pago/retido (leiaute DIRPF 2023 reg. 20); IN RFB 1.585/2015 art. 63 §8º III; o saldo de day trade de dezembro só por restituição (art. 65 §§9º e 12 II; P&R IRPF 2026 pergunta 715)
- `data/irpf/2025.yaml` — isentos: Programa IRPF 2026, ficha Rendimentos isentos e não tributáveis; Manual do MIR (rendimentos do capital e do patrimônio); Lei 9.249/1995 art. 10; Lei 11.033/2004 art. 3º I e III; IN RFB 1.585/2015 arts. 58 §1º e 74
- `data/irpf/2025.yaml` — renda_variavel: Programa IRPF 2026, ficha "Renda variável - operações comuns/day trade" (leiaute DIRPF 2023 reg. 40); P&R IRPF 2026 pergunta 707; IN RFB 1.585/2015 art. 62 I
- `data/irpf/2026.yaml` — imposto_pago: Programa IRPF 2026, ficha Imposto pago/retido (leiaute DIRPF 2023 reg. 20); IN RFB 1.585/2015 art. 63 §8º III; o saldo de day trade de dezembro só por restituição (art. 65 §§9º e 12 II; P&R IRPF 2026 pergunta 715)
- `data/irpf/2026.yaml` — isentos: Programa IRPF 2026, ficha Rendimentos isentos e não tributáveis; Manual do MIR (rendimentos do capital e do patrimônio); Lei 9.249/1995 art. 10; Lei 11.033/2004 art. 3º I e III; IN RFB 1.585/2015 arts. 58 §1º e 74
- `data/irpf/2026.yaml` — renda_variavel: Programa IRPF 2026, ficha "Renda variável - operações comuns/day trade" (leiaute DIRPF 2023 reg. 40); P&R IRPF 2026 pergunta 707; IN RFB 1.585/2015 art. 62 I
- `data/irpf/2027.yaml` — imposto_pago: Programa IRPF 2026, ficha Imposto pago/retido (leiaute DIRPF 2023 reg. 20); IN RFB 1.585/2015 art. 63 §8º III; o saldo de day trade de dezembro só por restituição (art. 65 §§9º e 12 II; P&R IRPF 2026 pergunta 715)
- `data/irpf/2027.yaml` — isentos: Programa IRPF 2026, ficha Rendimentos isentos e não tributáveis; Manual do MIR (rendimentos do capital e do patrimônio); Lei 9.249/1995 art. 10; Lei 11.033/2004 art. 3º I e III; IN RFB 1.585/2015 arts. 58 §1º e 74
- `data/irpf/2027.yaml` — renda_variavel: Programa IRPF 2026, ficha "Renda variável - operações comuns/day trade" (leiaute DIRPF 2023 reg. 40); P&R IRPF 2026 pergunta 707; IN RFB 1.585/2015 art. 62 I
- `data/isencao_pme.csv` — AGRO3: Lei 13.043/2014 art. 16; IN RFB 1.585/2015 arts. 66 a 69-A; lista B3 (AGRO3); oferta subsequente precificada em 04/02/2021
- `data/isencao_pme.csv` — PRIO3: Lei 13.043/2014 art. 16; IN RFB 1.585/2015 arts. 66 a 69-A; lista B3 (HRT Participações, HRTP3); RI PRIO: isentas as compradas entre 10/07/2014 e 29/01/2021 (oferta subsequente de jan/2021)
- `data/parametros_fiscais.csv` — leilao_fracao_isentavel: sem norma específica: fração vendida em leilão é alienação de ação no mercado à vista (IN RFB 1.585/2015 art. 58 caput); entra no limite de R$ 20 mil do mês do crédito (decisão do usuário 19/10/2026)
- `data/parametros_fiscais.csv` — reembolso_btc_dividendos_isento: IN RFB 1.585/2015 art. 74 §1º I a (Lei 13.043/2014 art. 7º)
- `data/parametros_fiscais.csv` — reembolso_btc_norma_omissa: IN RFB 1.585/2015 art. 74 §1º I a (reembolso de dividendos de BTC expressamente isento)
- `docs/regras-fiscais.md` — Decisões da F4 (motor fiscal, 14-15/10/2026): Nenhuma norma fixa o critério de rateio nem o arredondamento da parcela (a IN RFB 1.585/2015 só
- `docs/regras-fiscais.md` — Decisões da F4 (motor fiscal, 14-15/10/2026): é isento (IN RFB 1.585/2015 art. 59 §2º IV); a venda entra na soma do limite do mês dela.
- `docs/regras-fiscais.md` — Declaração anual (F8, 20/10/2026): Bens e Direitos pelo custo de aquisição (IN RFB 1.585/2015 art. 58), também no ano da compra; a
- `docs/regras-fiscais.md` — Declaração anual (F8, 20/10/2026): Termo (IN RFB 1.585/2015 art. 62 I e §1º): a venda à vista na data da liquidação vai para
- `docs/regras-fiscais.md` — Futuros de commodities e conferência das notas BM&F (F7a, 19/10/2026): O resultado vem dos ajustes em R$ impressos na nota (IN RFB 1.585/2015 art. 61). O multiplicador só
- `docs/regras-fiscais.md` — R1: R1 | Operações comuns (ações, units, ETF, BDR, opções, futuros, termo): 15% sobre ganho líquido mensal | Categoria COMUM, base única com subtotais por mercado | IN RFB 1.585/2015 | A
- `docs/regras-fiscais.md` — R10: R10 | Opção que vira pó: prêmio pago = perda (titular); recebido = ganho (lançador) no vencimento | Evento automático de vencimento (data como parâmetro, não "hoje") | IN RFB 1.585/2015 | A
- `docs/regras-fiscais.md` — R11: …por contrato e realizados no encerramento ou vencimento | Lei 11.051/2004 art. 32 (caput e §2º II); IN RFB 1.585/2015 art. 61; P&R IRPF 2026 pergunta 728 | A
- `docs/regras-fiscais.md` — R18: R18 | Termo: custo/alienação pelo preço a termo; ganho apurado na liquidação/encerramento (operação comum) | Contrato a termo vinculado à liquidação | IN RFB 1.585/2015 | M (validar com as operações a termo da planilha)
- `docs/regras-fiscais.md` — R2: R2 | Day trade: 20%; prejuízo só compensa com ganho DT | Categoria DAY_TRADE; excedente de quantidade vira COMUM | IN RFB 1.585/2015 | A
- `docs/regras-fiscais.md` — R3: R3 | FII/Fiagro: 20%; sem isenção 20 mil; prejuízo só com ganho FII | Categoria FII (não existe no v2.02) | IN RFB 1.585/2015 | A
- `docs/regras-fiscais.md` — R6: …o de 0,005% vai para a declaração anual e o de DT só restituição | Lei 9.959/2000 art. 8º §§4º, 8º; IN RFB 1.585/2015 arts. 63 §§8º e 9º e 65 §§8º, 9º e 12; P&R IRPF 2026 perguntas 714 e 715; a página "Retenções" da…
- `docs/regras-fiscais.md` — R7: R7 | Custo de aquisição inclui corretagem, emolumentos e taxas; na venda os custos reduzem a alienação | Rateio proporcional ao valor da operação na nota | IN RFB 1.585/2015 | A
- `docs/regras-fiscais.md` — R8: …Data do evento: data ex | evento_corporativo na linha do tempo do ativo (tela 7) | IN RFB 1.585/2015 art. 58 caput, §§6º e 7º, art. 60 §2º; P&R IRPF (substituição de ações; incorporação de ações); SC…
- `docs/regras-fiscais.md` — R9: …(ON→3, PN→4, PNA→5, PNB→6, UNT→11); sem classe: classe da abertura da série, depois opcoes.csv | IN RFB 1.585/2015 | M
- `tests/fiscal_cases/test_casos_fiscais.py` — test_bonificacao_ajusta_a_quantidade_da_serie_sem_custo: R8 (Lei 9.249/1995 art. 10 parágrafo único e IN RFB 1.585/2015 art. 58: o custo atribuído da bonificação vai para as ações) e ajuste de proventos da B3 nas…
- `tests/fiscal_cases/test_casos_fiscais.py` — test_compensacao_separada_entre_comum_e_day_trade: R2: prejuízo de day trade só compensa com ganho de day trade (IN RFB 1.585/2015).
- `tests/fiscal_cases/test_casos_fiscais.py` — test_exercicio_parcial_e_sobra_vira_po_no_vencimento: R9/R10 (IN RFB 1.585/2015 art. 58; P&R IRPF, opções): titular de 1.000 calls exerce 600 no vencimento; as 400 restantes viram pó (perde 40% do prêmio de 1.000 = -400); o custo das ações leva 600 do prêmio.
- `tests/fiscal_cases/test_casos_fiscais.py` — test_ganho_na_liquidacao_do_termo_nunca_e_isento: IN RFB 1.585/2015 art. 59 §2º IV: a isenção de R$ 20 mil não se aplica à alienação no vencimento ou na liquidação…
- `tests/fiscal_cases/test_casos_fiscais.py` — test_mesmo_codigo_de_serie_em_anos_diferentes_sao_posicoes_separadas: IN RFB 1.585/2015 art. 58 (custo médio por ativo) e regulamento de opções da B3 (o código da série se repete entre anos): TAEEL382 venceu em dez/2023 e em dez/2025; a de 2023 vira pó (R10), a de 2025 continua.
- `tests/fiscal_cases/test_casos_fiscais.py` — test_pm_fiscal_unico_por_titular_e_a_conta_de_origem_no_movimento: IN RFB 1.585/2015 art. 58: custo médio de aquisição do contribuinte, sem separar por corretora (decisões da F1 e do…
- `tests/fiscal_cases/test_casos_fiscais.py` — test_t5_ganho_da_venda_por_exercicio_nao_e_isento: IN RFB 1.585/2015 art. 59 §2º IV: a isenção não se aplica à alienação de ações em exercício de opções. Lançador de…
- `tests/fiscal_cases/test_casos_fiscais.py` — test_t5_titular_de_call_soma_o_premio_ao_custo: R9 (IN RFB 1.585/2015 art. 58): prêmio 1.000 + custo 5; exercício 1.000 a 30 com custos 20 -> custo 31.025, PM 31,025.
- `tests/fiscal_cases/test_casos_fiscais.py` — test_t6_opcao_que_vira_po_no_vencimento: R10: 1.000 opções a 0,50 com custo 5, não exercidas: titular perde 505; lançador ganha 495, no mês do vencimento (IN RFB 1.585/2015).
- `tests/fiscal_cases/test_casos_fiscais.py` — test_t7_day_trade_pela_ordem_dos_negocios_e_excedente_comum: R2: IN RFB 1.585/2015 art. 65 §3º ("pela ordem, o 1º negócio de compra com o 1º de venda") e §2º (o estoque anterior não…
- `tests/fiscal_cases/test_casos_fiscais.py` — test_t7_exercicio_e_venda_no_mesmo_dia_nao_e_day_trade: IN RFB 1.585/2015 art. 65 §13: exercício de opção não caracteriza day trade.
- `tests/fiscal_cases/test_casos_fiscais.py` — test_t8_day_trade_de_fii_fica_na_categoria_fii: R3: FII (inclusive day trade) a 20%, compensação só entre FII (IN RFB 1.585/2015). Ganho 100: IR 20,00 - IRRF 1,00 = 19,00.
- `tests/fiscal_cases/test_casos_fiscais.py` — test_termo_liquidado_depois_da_venda_das_acoes_realiza_o_resultado: R18 + venda a descoberto (leitura A1, decisão do usuário de 15/10/2026; IN RFB 1.585/2015 art. 58: custo médio): sem posição, compra a termo de 1.000 a 31,00 (custos 40) em 03/01 e venda…
- `tests/fiscal_cases/test_eventos_societarios.py` — módulo: …de fração (números conferidos à mão no parecer do especialista fiscal de 19/10/2026). Fontes: - IN RFB 1.585/2015 art. 58 §6º: na substituição de ações por incorporação, fusão ou cisão, o custo das ações…
- `tests/fiscal_cases/test_futuros_commodities.py` — módulo: …e na conferência do câmbio implícito. Casos fechados à mão pelo especialista (19/10/2026). Fontes: IN RFB 1.585/2015 art. 61 (resultado = soma dos ajustes no encerramento), art. 65 (15% e 20% day trade, inclusive…
- `tests/fiscal_cases/test_irpf_declaracao.py` — módulo: …Fontes: P&R IRPF 2026 perguntas 707, 720 (um item por ação, total das corretoras, custo médio) e 715; IN RFB 1.585/2015 arts. 58 §1º (bonificação: custo atribuído), 63 §8º III (saldo de IRRF de dezembro na declaração)…
- `tests/fiscal_cases/test_irpf_linhas.py` — módulo: …"Mercado a termo - ações/ouro"); P&R IRPF 2026 pergunta 707 (cotas de fundo de índice não são ações); IN RFB 1.585/2015 art. 62 I e §1º (ganho do comprador a termo pela venda à vista na data da liquidação; vendido…
- `tests/fiscal_cases/test_irrf_fichas.py` — módulo: …Casos do especialista fiscal, conferidos à mão. Fontes: Lei 11.033/2004 art. 2º §§1º, 2º, 5º e 7º; IN RFB 1.585/2015 arts. 63 §8º e 65 §§8º e 9º; fichas "Renda variável - operações comuns/day trade" e "Operações -…
- `tests/fiscal_cases/test_irrf_fichas.py` — test_day_trade_usado_no_imposto_de_fii_gera_aviso: IN RFB 1.585/2015 art. 65 §8º I: o IRRF de day trade deduz o imposto do mês (saldo único da DARF), mas a ficha FII…
- `tests/fiscal_cases/test_irrf_fichas.py` — test_irrf_de_day_trade_nao_abate_fii_antes_do_0005: IN RFB 1.585/2015 art. 65 §8º I e ficha FII sem linha de day trade: dezembro/2026 com imposto só de FII (V 100…
- `tests/fiscal_cases/test_proventos.py` — módulo: …na data do pagamento ou do crédito); Lei 11.033/2004 art. 3º III (rendimentos de FII isentos); IN RFB 1.585/2015 art. 74 e Lei 15.265/2025 arts. 20 e 21 (reembolso de BTC); IN 1.585 art. 58 §8º (restituição de…

## Res. CMN 4.880/2020

- `docs/regras-fiscais.md` — R12: …acumula, inclusive para o ano seguinte | `data/feriados.csv` (o último dia útil do ano não é útil: Res. CMN 4.880/2020 art. 3º) + acúmulo | RIR/2018 art. 915; Lei 9.430/1996 art. 68 caput e §1º | A

## SC Cosit 67/2024

- `docs/regras-fiscais.md` — Proventos (F7a, 19/10/2026): o informe. Alternativa (acessório segue o principal, SC Cosit 67/2024) não adotada.

## SC Cosit 145/2021

- `docs/regras-fiscais.md` — Decisões da F4 (motor fiscal, 14-15/10/2026): Units fora da isenção: coincide com a RFB (SC Cosit 145/2021: só "ativos da espécie ações").

## SC Cosit 224/2014

- `data/parametros_fiscais.csv` — incorporacao_acoes_tratamento: Lei 6.404 art. 252: custo das ações entregues mantido nas recebidas (TRF1/TRF3/TRF4; decisão do usuário 19/10/2026). A RFB diverge (SC Cosit 224/2014: alienação, ganho de capital no GCAP); a tela de Eventos avisa
- `docs/regras-fiscais.md` — R8: …art. 58 caput, §§6º e 7º, art. 60 §2º; P&R IRPF (substituição de ações; incorporação de ações); SC Cosit 224/2014 (divergência); Lei 6.404 arts. 227-229 e 252 | A
- `tests/fiscal_cases/test_eventos_societarios.py` — módulo: …6.404 art. 252): custo mantido (TRF1, TRF3, TRF4; decisão do usuário de 19/10/2026); a RFB diverge (SC Cosit 224/2014), parâmetro incorporacao_acoes_tratamento. - Ajuste de série de opção pela B3 (Manual de…

## P&R IRPF

- `docs/regras-fiscais.md` — Decisões da validação do gabarito (F1, 14/10/2026): 06/2008). R5 (P&R IRPF) diz que continua compensável; uma resposta de IA consultada pelo usuário diz o
- `docs/regras-fiscais.md` — Decisões da validação do gabarito (F1, 14/10/2026): P&R IRPF antes de implementar; enquanto isso, parâmetro configurável.
- `docs/regras-fiscais.md` — R8: …na linha do tempo do ativo (tela 7) | IN RFB 1.585/2015 art. 58 caput, §§6º e 7º, art. 60 §2º; P&R IRPF (substituição de ações; incorporação de ações); SC Cosit 224/2014 (divergência); Lei 6.404 arts. 227-229…
- `tests/fiscal_cases/test_casos_fiscais.py` — test_exercicio_parcial_e_sobra_vira_po_no_vencimento: R9/R10 (IN RFB 1.585/2015 art. 58; P&R IRPF, opções): titular de 1.000 calls exerce 600 no vencimento; as 400 restantes viram pó (perde 40% do prêmio de 1.000 = -400); o custo das ações leva 600 do prêmio.
- `tests/fiscal_cases/test_eventos_societarios.py` — módulo: …ou cisão, o custo das ações originais é atribuído às recebidas na proporção fixada pela assembleia; P&R IRPF (substituição de ações): não é alienação, a data de aquisição é a original. - Incorporação de ações (Lei…

## P&R IRPF 2026

- `coir/engine/futuros.py` — módulo: …ocasião da liquidação do contrato, cessão ou encerramento da posição". IN RFB 1.585/2015 art. 61 e P&R IRPF 2026 pergunta 728: resultado de cada contrato = soma algébrica dos ajustes diários entre a abertura e o…
- `coir/engine/proventos.py` — _jcp_com_credito: …conhecida (informe da empresa ou data de aprovação no pacote da B3): declarado no ano do crédito (P&R IRPF 2026 pergunta 320; Lei 9.249/1995 art. 9º §2º; a receber em 31/12 se pago depois), com a alíquota da data do…
- `data/irpf/2025.yaml` — bens: Programa IRPF 2026, ficha Bens e Direitos; Manual do MIR (bens financeiros); P&R IRPF 2026 perguntas 459 e 720 (um item por ação, total das corretoras); leiaute DIRPF 2023 reg. 27
- `data/irpf/2025.yaml` — imposto_pago: Programa IRPF 2026, ficha Imposto pago/retido (leiaute DIRPF 2023 reg. 20); IN RFB 1.585/2015 art. 63 §8º III; o saldo de day trade de dezembro só por restituição (art. 65 §§9º e 12 II; P&R IRPF 2026 pergunta 715)
- `data/irpf/2025.yaml` — renda_variavel: Programa IRPF 2026, ficha "Renda variável - operações comuns/day trade" (leiaute DIRPF 2023 reg. 40); P&R IRPF 2026 pergunta 707; IN RFB 1.585/2015 art. 62 I
- `data/irpf/2026.yaml` — bens: Programa IRPF 2026, ficha Bens e Direitos; Manual do MIR (bens financeiros); P&R IRPF 2026 perguntas 459 e 720 (um item por ação, total das corretoras); leiaute DIRPF 2023 reg. 27
- `data/irpf/2026.yaml` — imposto_pago: Programa IRPF 2026, ficha Imposto pago/retido (leiaute DIRPF 2023 reg. 20); IN RFB 1.585/2015 art. 63 §8º III; o saldo de day trade de dezembro só por restituição (art. 65 §§9º e 12 II; P&R IRPF 2026 pergunta 715)
- `data/irpf/2026.yaml` — renda_variavel: Programa IRPF 2026, ficha "Renda variável - operações comuns/day trade" (leiaute DIRPF 2023 reg. 40); P&R IRPF 2026 pergunta 707; IN RFB 1.585/2015 art. 62 I
- `data/irpf/2027.yaml` — bens: Programa IRPF 2026, ficha Bens e Direitos; Manual do MIR (bens financeiros); P&R IRPF 2026 perguntas 459 e 720 (um item por ação, total das corretoras); leiaute DIRPF 2023 reg. 27
- `data/irpf/2027.yaml` — imposto_pago: Programa IRPF 2026, ficha Imposto pago/retido (leiaute DIRPF 2023 reg. 20); IN RFB 1.585/2015 art. 63 §8º III; o saldo de day trade de dezembro só por restituição (art. 65 §§9º e 12 II; P&R IRPF 2026 pergunta 715)
- `data/irpf/2027.yaml` — renda_variavel: Programa IRPF 2026, ficha "Renda variável - operações comuns/day trade" (leiaute DIRPF 2023 reg. 40); P&R IRPF 2026 pergunta 707; IN RFB 1.585/2015 art. 62 I
- `data/parametros_fiscais.csv` — opcoes_linha_acoes_tipos_objeto: programa IRPF, ficha Renda variável: "Mercado de opções - ações" para opção sobre ação; ETF, índice e outros em "Mercado de opções - outros" (P&R IRPF 2026 pergunta 707 distingue cotas de fundo de índice de ações); unit como ação, coerente com o Bens e Direitos do programa (código 03.01, ações inclusive units)
- `docs/regras-fiscais.md` — Declaração anual (F8, 20/10/2026): Bens e Direitos: um item por ativo, custo médio do titular somando corretoras (P&R IRPF 2026
- `docs/regras-fiscais.md` — Declaração anual (F8, 20/10/2026): JCP: declarado no ano do crédito (P&R IRPF 2026 pergunta 320; Lei 9.249/1995 art. 9º §2º), com
- `docs/regras-fiscais.md` — Declaração anual (F8, 20/10/2026): P&R IRPF 2026 pergunta 723, que trata de lote isolado).
- `docs/regras-fiscais.md` — Declaração anual (F8, 20/10/2026): `opcoes_linha_acoes_tipos_objeto`; P&R IRPF 2026 pergunta 707 separa cotas de fundo de índice de
- `docs/regras-fiscais.md` — R11: …encerramento ou vencimento | Lei 11.051/2004 art. 32 (caput e §2º II); IN RFB 1.585/2015 art. 61; P&R IRPF 2026 pergunta 728 | A
- `docs/regras-fiscais.md` — R6: …| Lei 9.959/2000 art. 8º §§4º, 8º; IN RFB 1.585/2015 arts. 63 §§8º e 9º e 65 §§8º, 9º e 12; P&R IRPF 2026 perguntas 714 e 715; a página "Retenções" da RFB (19/11/2025) restringe o DT ao imposto de DT:…
- `tests/fiscal_cases/test_irpf_declaracao.py` — módulo: …Direitos, rendimentos e fichas mês a mês que o COIR preenche. Casos do especialista fiscal (parecer de 20/10/2026), conferidos à mão; custos zero. Fontes: P&R IRPF 2026 perguntas 707, 720 (um item por ação, total das…
- `tests/fiscal_cases/test_irpf_linhas.py` — módulo: …- operações comuns/day trade" do programa IRPF (F8). Casos do especialista fiscal (pareceres de 20 e 21/10/2026), conferidos à mão. Fontes: programa IRPF, ficha de renda variável (linhas "Mercado de opções -…
- `tests/fiscal_cases/test_proventos.py` — test_jcp_creditado_no_ano_anterior_com_a_mesma_aliquota: Caso C do especialista (F8, 20/10/2026): crédito em 16/12/2024, pago em 10/03/2025, 15% nos dois anos (1.000 x 0,50 bruto, líquido 425,00). Sem…

## Manual MIR 2026

- `data/parametros_fiscais.csv` — prejuizo_mes_isento_compensavel: Lei 8.981/1995 art. 72 §8º; Manual MIR 2026; decisão do usuário (14/10/2026)
- `docs/regras-fiscais.md` — R5: R5 | Prejuízo de mês isento continua compensável | Parâmetro `prejuizo_mes_isento_compensavel = true` | Lei 8.981/1995 art. 72 §8º; Manual MIR 2026; decisão do usuário (14/10/2026) | A
- `tests/fiscal_cases/test_casos_fiscais.py` — test_t1_prejuizo_em_mes_isento_compensa: R5: prejuízo de ações em mês isento continua compensável (Lei 8.981/1995 art. 72 §8º; Manual MIR 2026; decisão do usuário). Mar/24: ações -1.200 (vendas 15.000), opções -300. Abr/24: ganho 2.000 -> base…

## Programa IRPF

- `coir/engine/apuracao.py` — módulo: …a declaração anual (IN RFB 1.585/2015; decisão do usuário). A DARF usa um saldo único; as fichas do programa IRPF recebem o IRRF que cada uma consumiu (`fichas.py`). - DARF 6015: abaixo de R$ 10 acumula, inclusive para…
- `coir/engine/fichas.py` — módulo: IRRF por ficha do programa IRPF (comuns/day trade e FII), com a DARF no saldo único. A DARF mensal deduz o IRRF de um saldo único (Lei…
- `data/parametros_fiscais.csv` — opcoes_linha_acoes_tipos_objeto: programa IRPF, ficha Renda variável: "Mercado de opções - ações" para opção sobre ação; ETF, índice e outros em "Mercado de opções - outros" (P&R IRPF 2026 pergunta 707 distingue cotas de fundo de índice de ações); unit como ação, coerente com o Bens e Direitos do programa (código 03.01, ações inclusive units)
- `docs/regras-fiscais.md` — R17: R17 | Fichas/códigos da DIRPF mudam por ano | Mapeamento por ano em `data/irpf/<ano>.yaml` | programa IRPF do ano | revisar anualmente
- `docs/regras-fiscais.md` — Relatórios (F5, 16/10/2026): conferir no programa IRPF do ano: a ficha FII sem linha de IR fonte de day trade; ordem de
- `tests/fiscal_cases/test_irpf_linhas.py` — módulo: Linha da ficha "Renda variável - operações comuns/day trade" do programa IRPF (F8). Casos do especialista fiscal (pareceres de 20 e 21/10/2026), conferidos à mão. Fontes: programa…
- `tests/fiscal_cases/test_irrf_fichas.py` — módulo: IRRF por ficha do programa IRPF com a DARF no saldo único (decisão do usuário de 16/10/2026: separar o IRRF de FII só na declaração).…

## CTN

- `docs/regras-fiscais.md` — DARF paga em atraso: multa de mora (F6, 17/10/2026): A denúncia espontânea (CTN art. 138; STJ REsp 1.149.022/SP; Súmula 360) fica como nota na tela:

## Súmula STJ 360

- `docs/regras-fiscais.md` — DARF paga em atraso: multa de mora (F6, 17/10/2026): A denúncia espontânea (CTN art. 138; STJ REsp 1.149.022/SP; Súmula 360) fica como nota na tela:

## REsp 1.149.022/SP

- `docs/regras-fiscais.md` — DARF paga em atraso: multa de mora (F6, 17/10/2026): A denúncia espontânea (CTN art. 138; STJ REsp 1.149.022/SP; Súmula 360) fica como nota na tela:

## RIR/2018

- `coir/engine/apuracao.py` — módulo: …inclusive para o ano seguinte (Lei 9.430/1996 art. 68 §1º); vence no último dia útil do mês seguinte (RIR/2018 art. 915).
- `coir/engine/calendario.py` — módulo: …de termo (datas sempre por parâmetro). - DARF: último dia útil (bancário) do mês seguinte ao da apuração (RIR/2018 art. 915; Lei 9.430/1996 art. 68 §1º para o acúmulo abaixo do mínimo). - Opções sobre ações: regra…
- `data/parametros_fiscais.csv` — aliquota_comum: Lei 11.033/2004 art. 2º II; RIR/2018 art. 839
- `docs/regras-fiscais.md` — R12: …| `data/feriados.csv` (o último dia útil do ano não é útil: Res. CMN 4.880/2020 art. 3º) + acúmulo | RIR/2018 art. 915; Lei 9.430/1996 art. 68 caput e §1º | A
- `tests/fiscal_cases/test_casos_fiscais.py` — test_eventos_ajustam_quantidade_e_custo_sem_resultado: R8: bonificação soma quantidade com custo atribuído; desdobramento e grupamento só a quantidade; redução de capital reduz o custo (RIR/2018; Perguntas e Respostas IRPF).
- `tests/fiscal_cases/test_casos_fiscais.py` — test_t4_darf_abaixo_de_10_acumula_e_vira_o_ano: R12: DARF abaixo de R$ 10 acumula, inclusive para o ano seguinte (Lei 9.430/1996 art. 68 §1º); vence no último dia útil do mês seguinte (RIR/2018 art. 915). Dez/23 7,00 + Jan/24 4,00 = 11,00, vencimento 29/02/2024.
