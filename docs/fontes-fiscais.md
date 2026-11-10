# Fontes fiscais do COIR

Gerado por `uv run python -m tools.fontes_fiscais` a partir das tabelas de `data/`, de
`docs/regras-fiscais.md` e das docstrings dos casos fiscais e do motor. Não edite à mão:
corrija a citação na origem e gere de novo. O documento orienta; não substitui a norma.
O que cada norma diz (vigência, artigos, pontos em aberto) está na base normativa:
[`docs/normas/`](normas/README.md).

| Norma | Usos | Ficha |
|---|---|---|
| LC 224/2025 | 8 | [lc-224-2025.md](normas/lc-224-2025.md) |
| Lei 6.404/1976 | 7 | [lei-6404-1976.md](normas/lei-6404-1976.md) |
| Lei 7.713/1988 | 2 | [lei-7713-1988.md](normas/lei-7713-1988.md) |
| Lei 8.668/1993 | 1 | [lei-8668-1993.md](normas/lei-8668-1993.md) |
| Lei 8.981/1995 | 10 | [lei-8981-1995.md](normas/lei-8981-1995.md) |
| Lei 9.249/1995 | 17 | [lei-9249-1995.md](normas/lei-9249-1995.md) |
| Lei 9.250/1995 | 4 | [lei-9250-1995.md](normas/lei-9250-1995.md) |
| Lei 9.430/1996 | 17 | [lei-9430-1996.md](normas/lei-9430-1996.md) |
| Lei 9.779/1999 | 1 | [lei-9779-1999.md](normas/lei-9779-1999.md) |
| Lei 9.959/2000 | 5 | [lei-9959-2000.md](normas/lei-9959-2000.md) |
| Lei 11.033/2004 | 28 | [lei-11033-2004.md](normas/lei-11033-2004.md) |
| Lei 11.051/2004 | 5 | [lei-11051-2004.md](normas/lei-11051-2004.md) |
| Lei 11.478/2007 | 6 | [lei-11478-2007.md](normas/lei-11478-2007.md) |
| Lei 12.350/2010 | 1 | [lei-12350-2010.md](normas/lei-12350-2010.md) |
| Lei 12.431/2011 | 8 | [lei-12431-2011.md](normas/lei-12431-2011.md) |
| Lei 13.043/2014 | 29 | [lei-13043-2014.md](normas/lei-13043-2014.md) |
| Lei 14.754/2023 | 3 | [lei-14754-2023.md](normas/lei-14754-2023.md) |
| Lei 15.265/2025 | 15 | [lei-15265-2025.md](normas/lei-15265-2025.md) |
| Lei 15.270/2025 | 10 | [lei-15270-2025.md](normas/lei-15270-2025.md) |
| MP 497/2010 | 5 | [mp-497-2010.md](normas/mp-497-2010.md) |
| IN RFB 1.585/2015 | 94 | [in-rfb-1585-2015.md](normas/in-rfb-1585-2015.md) |
| IN RFB 1.700/2017 | 2 | [in-rfb-1700-2017.md](normas/in-rfb-1700-2017.md) |
| IN RFB 1.916/2019 | 1 | [in-rfb-1916-2019.md](normas/in-rfb-1916-2019.md) |
| Res. CMN 4.880/2020 | 1 | [res-cmn-4880-2020.md](normas/res-cmn-4880-2020.md) |
| SC Cosit 37/2020 | 3 | [sc-cosit-37-2020.md](normas/sc-cosit-37-2020.md) |
| SC Cosit 42/2020 | 2 | [sc-cosit-42-2020.md](normas/sc-cosit-42-2020.md) |
| SC Cosit 63/2020 | 2 | [sc-cosit-63-2020.md](normas/sc-cosit-63-2020.md) |
| SC Cosit 67/2024 | 1 | [sc-cosit-67-2024.md](normas/sc-cosit-67-2024.md) |
| SC Cosit 145/2021 | 3 | [sc-cosit-145-2021.md](normas/sc-cosit-145-2021.md) |
| SC Cosit 166/2021 | 3 | [sc-cosit-166-2021.md](normas/sc-cosit-166-2021.md) |
| SC Cosit 224/2014 | 4 | [sc-cosit-224-2014.md](normas/sc-cosit-224-2014.md) |
| P&R IRPF | 5 | [pr-irpf.md](normas/pr-irpf.md) |
| P&R IRPF 2026 | 21 | [pr-irpf.md](normas/pr-irpf.md) |
| Manual MIR | 2 | [manual-mir.md](normas/manual-mir.md) |
| Manual MIR 2026 | 3 | [manual-mir.md](normas/manual-mir.md) |
| Programa IRPF | 8 | — |
| CTN | 5 | [lei-5172-1966-ctn.md](normas/lei-5172-1966-ctn.md) |
| Súmula STJ 360 | 1 | — |
| REsp 1.149.022/SP | 1 | — |
| RIR/2018 | 6 | [decreto-9580-2018-rir.md](normas/decreto-9580-2018-rir.md) |
| Instrução Normativa SRF 41/1998 | 5 | [in-srf-41-1998.md](normas/in-srf-41-1998.md) |

## Fontes dos dados públicos (pacote de referências)

As tabelas geradas em `data/` (o pacote que vai no zip do programa e no release
`referencias-AAAA.MM.DD`) vêm só de fontes públicas, sem login. Cada linha de
`data/eventos_b3.csv` diz de onde veio (coluna `origem`, e `referencia` com o documento).

| Fonte | URL | O que cobre | Onde entra |
|---|---|---|---|
| B3, consulta de empresas listadas (`GetListedSupplementCompany`) | `https://sistemaswebb3-listados.b3.com.br/listedCompaniesProxy/CompanyCall/GetListedSupplementCompany/` | Desdobramento, grupamento, bonificação, cisão, resgate e subscrição com data com e fator exatos; proventos em dinheiro recentes. **Só o último evento de cada tipo por ISIN**: o histórico se perde quando a companhia repete o tipo de evento. | coleta semanal; `origem` = `b3` |
| B3, proventos em dinheiro (`GetListedCashDividends`) | mesmo endereço, método `GetListedCashDividends` | JCP desde 2016, com data de aprovação e data com. | `proventos_historico_b3.csv` |
| B3, cadastro de instrumentos | `https://arquivos.b3.com.br/api/download/requestname?fileName=InstrumentsConsolidatedFile` | Tickers, ISIN, espécie e emissor do dia (o portal guarda só o último mês). | `instrumentos_b3.csv` e derivados |
| CVM, Formulário de Referência (dados abertos), tabela `capital_social_desdobramento` (e `..._classe_acao`) | `https://dados.cvm.gov.br/dados/CIA_ABERTA/DOC/FRE/DADOS/fre_cia_aberta_AAAA.zip` (2010 a 2023) | Desdobramentos, grupamentos e bonificações aprovados de 2007 a 2022 (cada FRE traz os três exercícios anteriores): data de aprovação e quantidade de ações antes e depois, por espécie (ON, PN, PNA, PNB). O FRE de 2024 em diante não tem mais a tabela. | `python -m tools.referencias historico`; `origem` = `cvm_fre` |
| B3, série histórica de cotações (COTAHIST) | `https://bvmf.bmfbovespa.com.br/InstDados/SerHist/COTAHIST_AAAA.ZIP` | Por pregão e ISIN: preço e número de distribuição (muda no pregão seguinte à data com de todo evento). Confirma a data com dos eventos do FRE e dos avisos. | idem (não é publicado: só confere) |
| CVM, documentos IPE (dados abertos) | `https://dados.cvm.gov.br/dados/CIA_ABERTA/DOC/IPE/DADOS/ipe_cia_aberta_AAAA.zip` e o documento em `https://www.rad.cvm.gov.br/ENET/frmDownloadDocumento.aspx` | Fatos relevantes e avisos aos acionistas: proporção, data-base, data ex e custo atribuído da bonificação. Usados para 2023 em diante, quando o FRE parou e a B3 já tinha substituído o evento, e para o que o FRE não resolve (fator ruidoso, papel sem negócio para conferir). | `tools/referencias/eventos_cvm_ipe.csv` (transcrito, com o protocolo); `origem` = `cvm_ipe` |
| CVM, Formulário Cadastral (FCA, dados abertos), tabelas `geral` e `valor_mobiliario` | `https://dados.cvm.gov.br/dados/CIA_ABERTA/DOC/FCA/DADOS/fca_cia_aberta_AAAA.zip` (2010 em diante) | Por CNPJ: nome empresarial com o nome anterior e a data da mudança; códigos de negociação (só de 2018 em diante). Confirma que o código antigo e o novo de uma troca são a mesma companhia, e liga ao FRE a companhia que saiu da bolsa. | `trocas_ticker.csv` (coluna `fonte`; curada) e `python -m tools.referencias historico` |
| Banco Central, série 4390 (Selic mensal) | `https://api.bcb.gov.br/dados/serie/bcdata.sgs.4390/dados?formato=json` | Taxa Selic de cada mês encerrado. | `selic_mensal.csv` |

Regras da combinação (`tools/referencias/gerar.py`, `historico.py`):

- nada é apagado sem registro: a linha que uma fonte deixou de mostrar continua no pacote. O que
  sai vai para `tools/referencias/eventos_excluidos.csv`, com as colunas da chave do evento
  (`isin`, `tipo`, `data_com`, `isin_emitido`, `fator`), o `motivo` e a `fonte`; a junção não
  reinsere o que está na lista, mesmo que a B3 o devolva de novo, e exclusão sem motivo ou sem
  fonte faz a geração falhar. Hoje: o par de ITUB 2011 (abaixo) e GSHP 26/03/2019, que a B3 lista
  como bonificação de 100% e é dividendo pago em dinheiro e em cotas do FII Top Center (nenhuma
  ação emitida; provento isento, Lei 9.249/95, art. 10). As cotas entram à mão como entrada, com
  custo de R$ 9,2089623 por cota (R$ 621.716.835,00 / 67.512.149 ações; CVM, Reg. 1311/19, e
  aviso aos acionistas de 26/03/2019). O COIR não modela dividendo pago em ativo de outra empresa
  (pendência para versões futuras);
- o mesmo evento vindo de duas fontes soma a origem (`b3+cvm_fre`); fator diferente na mesma
  data com, ou o mesmo fator em data com próxima, é **divergência**: o pacote fica como estava e
  o caso vai para o PR semanal, para revisão humana;
- do FRE só entra fator que é fração simples (p/q com q até 20, 1/n ou n) e data com confirmada
  pela cotação (a única mudança de distribuição, depois da aprovação, em que o preço andou como o
  fator manda; ou a data com já confirmada de outra espécie do mesmo ato). Ambíguo, sem negócio
  no pregão ou com fator ruidoso vai para revisão, nunca para o pacote;
- bonificação e desdobramento mexem na quantidade do mesmo jeito: o mesmo ato chamado de um numa
  fonte (ou versão do FRE) e de outro noutra não entra duas vezes (dobraria a posição); vai para
  revisão e o tipo certo sai do aviso aos acionistas;
- grupamento e desdobramento simultâneos ficam como duas linhas na mesma data com, como a B3
  publica; quando a B3 só guardou uma delas, a outra vem do documento (BEES 2012). O par que se
  anula (100 para 1 e 1 para 100) não entra em proposta nem em controle nenhum, porque a quantidade
  não muda: o de ITUB 2011 saiu do pacote pela lista de exclusões, e os demais (BBDC 2009, MERC
  2023, TIMS 2025...) o COIR descarta ao ler o pacote (`coir/referencias.eventos_b3`), sem
  depender da lista;
- a unit entra com o fator das ações só quando todas as espécies tiveram o mesmo fator e a mesma
  data com e a cotação da unit confirma.

Trocas de código (`data/trocas_ticker.csv`, curada à mão, uma linha por espécie):

- a troca de código quase sempre troca o ISIN junto (KROT3 era `BRKROTACNOR9`, COGN3 é
  `BRCOGNACNOR2`): a tabela guarda os dois (`isin_antigo`, `isin_novo`) e a `fonte` de cada linha;
- só é troca o que é a mesma companhia (mesmo CNPJ no FCA: os dois códigos no mesmo CNPJ, ou o
  nome anterior no histórico do CNPJ) com o preço contínuo de um código para o outro (1 para 1,
  último negócio do antigo e primeiro do novo no COTAHIST). Incorporação com relação de troca ou
  com CNPJ diferente (IGTA3 -> IGTI3) não é troca: é evento corporativo, e fica para revisão
  humana. A exceção é a incorporação 1 para 1 em que o acionista só trocou de código, revista e
  confirmada caso a caso (TIMP3 -> TIMS3, SMLE3 -> SMLS3, JPSA3 -> IGTI3: `fonte` "troca
  curada");
- `ate` é o pregão anterior ao primeiro negócio do código novo (o mesmo dia para todas as espécies
  da companhia, mesmo a que ficou sem negócio);
- código que voltou (CTAX3 -> LIQO3 em 2018, de novo CTAX3 no ISIN antigo em 2020) não entra: a
  tabela não representa a volta, e as séries de cotação se misturariam;
- o histórico (`python -m tools.referencias historico`) encadeia as séries do ISIN antigo e do
  novo: o evento do FRE (que é por CNPJ) de antes da troca entra no ISIN antigo, e o que tem a data
  ex no dia da troca (LLIS3 -> VSTE3 com grupamento 8:1 em 09/02/2023) entra no ISIN que vale na
  data ex. Sem isso o evento se perdia, porque o ISIN antigo não tem o código do emissor atual;
- o app (`coir/referencias.tickers_por_isin`) leva o ISIN antigo ao ticker da época
  (`BRKROTACNOR9` -> KROT3), que o cadastro de instrumentos de hoje não tem, e a troca liga os dois
  tickers ao mesmo ativo.

Companhias que saíram da bolsa:

- a consulta de empresas da B3 e o cadastro de emissores (`emissores_b3.csv`) só trazem as
  listadas hoje: a companhia deslistada ficava sem nenhum evento no pacote;
- o histórico liga o CNPJ do FRE ao código de negociação pelo FCA (2018 em diante: o código com
  negócio mais recente no COTAHIST; CNPJ com dois códigos empatados fica de fora). Antes de 2018
  o FCA não traz código, e a companhia deslistada antes disso continua sem o FRE;
- de 2023 em diante (sem FRE), o evento vem do aviso aos acionistas transcrito em
  `eventos_cvm_ipe.csv`, conferido pela cotação; a varredura do COTAHIST (mudança de distribuição
  com o preço andando 8% ou mais, sem evento no pacote) acha os candidatos, e a queda com aviso de
  dividendo na mesma janela não entra (é provento em dinheiro, não evento em ações).

## LC 224/2025

- `coir/engine/proventos.py` — _jcp_com_credito: …depois: art. 3º parágrafo único), com a alíquota da data do crédito ou do pagamento, o que vier primeiro (LC 224/2025 art. 8º). Retenção implícita com a alíquota do ano anterior: só aviso (a empresa tratou o crédito…
- `coir/engine/proventos.py` — módulo: …retida: 15% até 2025, 17,5% desde 2026, pela data do pagamento ou do crédito, o que vier primeiro (LC 224/2025 art. 8º). Pago num ano com a alíquota do ano anterior = creditado no ano anterior: declarado no ano…
- `data/parametros_fiscais.csv` — aliquota_irrf_jcp: LC 224/2025 art. 8º e art. 14 III (17,5% pago ou creditado desde 01/01/2026)
- `docs/regras-fiscais.md` — Declaração anual (F8, 20/10/2026): 9º §2º), com a alíquota da data do crédito ou do pagamento, o que vier primeiro (LC 224/2025 art.
- `docs/regras-fiscais.md` — Proventos (F7a, 19/10/2026): alíquota retida (15% até 2025; 17,5% desde 2026, LC 224/2025 art. 8º, pela data do pagamento ou do
- `docs/regras-fiscais.md` — R15: R15 | JCP: IRRF exclusivo 15% até 31/12/2025; 17,5% para JCP pago ou creditado desde 01/01/2026 | Provento com IRRF; alíquota com vigência | LC 224/2025 | A
- `tests/fiscal_cases/test_proventos.py` — módulo: …conferidos à mão, 19/10/2026). Fontes: Lei 9.249/1995 arts. 9º e 10 (JCP exclusivo; dividendos isentos); LC 224/2025 art. 8º e art. 14 III (JCP a 17,5% pago ou creditado desde 01/01/2026, na data do pagamento ou do…
- `tests/fiscal_cases/test_proventos.py` — test_jcp_creditado_em_2025_e_pago_em_2026_usa_a_aliquota_do_credito: LC 224/2025 art. 8º: 17,5% para JCP pago ou creditado desde 01/01/2026, o que vier primeiro. Creditado em 20/12/2025 e pago em 15/04/2026: 15%, declarado em 2025.

## Lei 6.404/1976

- `data/parametros_fiscais.csv` — direito_vencido_data: IN RFB 1.585/2015 art. 60 §3º (analogia: perda na data do vencimento); Lei 6.404 art. 171 §4º (prazo decadencial): o direito comprado não exercido vira perda no fim do prazo de subscrição do pacote da B3, ou na data da baixa na B3 sem ele; alternativa baixa_b3 (parecer do especialista de 09/11/2026 seguido pelo usuário)
- `data/parametros_fiscais.csv` — incorporacao_acoes_tratamento: Lei 6.404 art. 252: custo das ações entregues mantido nas recebidas (TRF1/TRF3/TRF4; decisão do usuário 19/10/2026). A RFB diverge (SC Cosit 224/2014: alienação, ganho de capital no GCAP); a tela de Eventos avisa
- `docs/regras-fiscais.md` — Declaração anual (F8, 20/10/2026): 75 §10), o que pede a deliberação e a data com (Lei 6.404/1976 art. 205). A data do crédito vem
- `docs/regras-fiscais.md` — R21: …RFB 1.585/2015 art. 60 §3º (analogia com a opção que vira pó), arts. 56 §3º, 58 e 64; SC Cosit 166/2021; Lei 6.404 art. 171 §4º; P&R IRPF 2026 pergunta 727; parecer de 09/11/2026, seguido pelo usuário | M (sem norma…
- `docs/regras-fiscais.md` — R8: …de ações; incorporação de ações); SC Cosit 224/2014 (divergência); SC Cosit 37/2020 (resgate); Lei 6.404 arts. 227-229 e 252 | A
- `tests/fiscal_cases/test_direito_nao_exercido.py` — módulo: …1.585/2015 art. 60 §3º, por analogia com a opção que vira pó; arts. 56 §3º, 58 e 64; SC Cosit 166/2021; Lei 6.404 art. 171 §4º). Parecer do especialista de 09/11/2026, seguido pelo usuário.
- `tests/fiscal_cases/test_eventos_societarios.py` — módulo: …(substituição de ações): não é alienação, a data de aquisição é a original. - Incorporação de ações (Lei 6.404 art. 252): custo mantido (TRF1, TRF3, TRF4; decisão do usuário de 19/10/2026); a RFB diverge (SC Cosit…

## Lei 7.713/1988

- `tests/fiscal_cases/test_petz_auau.py` — módulo: …pela companhia: 0,71 / 4,39 = 16,17312073%. - O resgate da PN é alienação: dinheiro menos o custo da PN (Lei 7.713/1988 art. 3º §3º; SC Cosit 37/2020, CETIP/BM&FBOVESPA, com a mesma estrutura de ON + PN resgatável).…
- `tests/fiscal_cases/test_raiz55.py` — módulo: …e primeiro de RAIZ55 em 03/11/2026. Nada muda na quantidade nem nos direitos. Regra: não há alienação (Lei 7.713/1988 art. 3º §3º), e o custo e a quantidade continuam os mesmos. Na nota, a quantidade é a de ações e o…

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

- `coir/engine/proventos.py` — _jcp_com_credito: …B3, a mais tarde entre a aprovação e a data com): declarado no ano do crédito (IN SRF 41/1998 art. 3º; Lei 9.249/1995 art. 9º §2º; a receber em 31/12 se pago depois: art. 3º parágrafo único), com a alíquota da data…
- `coir/engine/proventos.py` — módulo: …da B3 (decisões do usuário de 19/10/2026). - Dividendo de empresa brasileira (ação, unit): isento (Lei 9.249/1995 art. 10). De BDR, ETF ou FII: fora da v3.0, listado com aviso (BDR: Lei 14.754/2023, não é…
- `data/irpf/2025.yaml` — exclusivos: Programa IRPF 2026, ficha Rendimentos sujeitos à tributação exclusiva/definitiva; Lei 9.249/1995 art. 9º; Lei 13.043/2014 art. 2º; Lei 8.981/1995 art. 65 §4º c
- `data/irpf/2025.yaml` — isentos: Programa IRPF 2026, ficha Rendimentos isentos e não tributáveis; Manual do MIR (rendimentos do capital e do patrimônio); Lei 9.249/1995 art. 10; Lei 11.033/2004 art. 3º I e III; IN RFB 1.585/2015 arts. 58 §1º e 74
- `data/irpf/2026.yaml` — exclusivos: Programa IRPF 2026, ficha Rendimentos sujeitos à tributação exclusiva/definitiva; Lei 9.249/1995 art. 9º; Lei 13.043/2014 art. 2º; Lei 8.981/1995 art. 65 §4º c
- `data/irpf/2026.yaml` — isentos: Programa IRPF 2026, ficha Rendimentos isentos e não tributáveis; Manual do MIR (rendimentos do capital e do patrimônio); Lei 9.249/1995 art. 10; Lei 11.033/2004 art. 3º I e III; IN RFB 1.585/2015 arts. 58 §1º e 74
- `data/irpf/2027.yaml` — exclusivos: Programa IRPF 2026, ficha Rendimentos sujeitos à tributação exclusiva/definitiva; Lei 9.249/1995 art. 9º; Lei 13.043/2014 art. 2º; Lei 8.981/1995 art. 65 §4º c
- `data/irpf/2027.yaml` — isentos: Programa IRPF 2026, ficha Rendimentos isentos e não tributáveis; Manual do MIR (rendimentos do capital e do patrimônio); Lei 9.249/1995 art. 10; Lei 11.033/2004 art. 3º I e III; IN RFB 1.585/2015 arts. 58 §1º e 74
- `data/parametros_fiscais.csv` — aliquota_irrf_jcp: Lei 9.249/1995 art. 9º §2º (15% na data do pagamento ou crédito)
- `docs/regras-fiscais.md` — Declaração anual (F8, 20/10/2026): como crédito contra a companhia (IN SRF 41/1998 art. 3º e parágrafo único; Lei 9.249/1995 art.
- `docs/regras-fiscais.md` — Oferta pública (IPO e follow-on; parecer do especialista, 08/11/2026): com o custo por ação lido do comunicado da companhia na CVM (`cvm_ipe`), mostrado antes de lançar (Lei 9.249
- `docs/regras-fiscais.md` — Proventos (F7a, 19/10/2026): Dividendo de ação ou unit: isento (Lei 9.249/1995 art. 10), código 09. BDR, ETF e FII com
- `tests/fiscal_cases/test_casos_fiscais.py` — test_bonificacao_ajusta_a_quantidade_da_serie_sem_custo: R8 (Lei 9.249/1995 art. 10 parágrafo único e IN RFB 1.585/2015 art. 58: o custo atribuído da bonificação vai para as…
- `tests/fiscal_cases/test_irpf_declaracao.py` — módulo: …63 §8º III (saldo de IRRF de dezembro na declaração) e 74; Lei 11.033/2004 art. 3º I (código 20); Lei 9.249/1995 arts. 9º e 10; Manual do MIR (Bens e Direitos, rendimentos do capital); data/irpf/2026.yaml.
- `tests/fiscal_cases/test_oferta_publica.py` — test_bonificacao_noutra_classe_abre_posicao_com_o_custo: Lei 9.249/1995 art. 10 §1º: o custo é a parcela capitalizada que corresponde ao acionista. Quem tem PN recebe 100…
- `tests/fiscal_cases/test_proventos.py` — módulo: Proventos a partir da movimentação da B3 (casos fictícios conferidos à mão, 19/10/2026). Fontes: Lei 9.249/1995 arts. 9º e 10 (JCP exclusivo; dividendos isentos); LC 224/2025 art. 8º e art. 14 III (JCP a 17,5%…
- `tests/fiscal_cases/test_proventos.py` — test_jcp_creditado_no_ano_anterior_com_a_mesma_aliquota: …B3 ou informe), declarado em 2024 e a receber em 31/12/2024 (IN SRF 41/1998 art. 3º e parágrafo único; Lei 9.249/1995 art. 9º §2º).

## Lei 9.250/1995

- `data/parametros_fiscais.csv` — aliquota_irrf_dividendos_acima_limite: Lei 15.270/2025 (Lei 9.250 art. 6º-A caput e §§1º e 2º)
- `data/parametros_fiscais.csv` — dividendos_limite_mensal_mesma_pj: Lei 15.270/2025 (Lei 9.250 art. 6º-A): acima de R$ 50 mil no mês da mesma PJ há retenção de 10% sobre o total
- `tests/fiscal_cases/test_petz_auau.py` — módulo: …mensal; a perda não compensa ganhos da bolsa. - Isenção do resgate: até R$ 35 mil alienados no mês (Lei 9.250/1995 art. 22 II). - A incorporação em si segue o parâmetro incorporacao_acoes_tratamento (custo…
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
- `docs/regras-fiscais.md` — Decisões da F4 (motor fiscal, 14-15/10/2026): venda de ação no limite (Lei 11.033/2004 art. 3º I; parecer do especialista de 09/11/2026,
- `docs/regras-fiscais.md` — Proventos (F7a, 19/10/2026): Rendimento de FII/Fiagro: isento (Lei 11.033 art. 3º III).
- `docs/regras-fiscais.md` — Proventos (F7a, 19/10/2026): §4º c; Lei 11.033 art. 1º), como as fontes retêm (22,5% nos dados reais); código 06 ou 12 conforme
- `docs/regras-fiscais.md` — R16: R16 | Rendimentos de FII isentos se o fundo cumpre requisitos | Usuário marca conforme informe do fundo | Lei 11.033/2004 art. 3º (alt. Lei 14.754/2023) | M
- `docs/regras-fiscais.md` — R4: …e venda por exercício incluídos); fora: units, ETF, BDR, FII, opções, futuros, termo e day trade | Lei 11.033/2004 art. 3º I | A
- `docs/regras-fiscais.md` — Relatórios (F5, 16/10/2026): abatimento (day trade antes da Lei 11.033) na ficha de comuns.
- `docs/regras-fiscais.md` — Relatórios (F5, 16/10/2026): duas partes, day trade e Lei 11.033, com mês, meses anteriores e a compensar; o de day trade é
- `tests/fiscal_cases/test_ativo_isolado.py` — test_venda_do_ativo_sem_tipo_conta_no_limite_e_avisa_quando_decide: …conta como venda de ação no limite (parecer do especialista de 09/11/2026, seguido pelo usuário; Lei 11.033/2004 art. 3º I): se for ação, não contar isentaria vendas acima do limite. Vendas de fevereiro: 19.000…
- `tests/fiscal_cases/test_casos_fiscais.py` — test_t2_limite_pelo_valor_bruto_das_vendas: R4: o limite de R$ 20 mil é sobre o valor bruto das alienações (Lei 11.033/2004 art. 3º I); custos de venda não reduzem o valor para o limite. 20.000,00 com custos 10 é isento; 20.000,01 é tributável.
- `tests/fiscal_cases/test_direitos_subscricao.py` — módulo: …sem isenção e fora da soma do limite de R$ 20 mil (SC Cosit 145/2021 e 166/2021, ementas oficiais; Lei 11.033/2004 art. 2º; IN RFB 1.585/2015 arts. 58, 59 e 65). O direito recebido sem nota de compra entra com…
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

## Lei 11.478/2007

- `coir/engine/proventos.py` — _rendimento_fundo_infra: …do líquido indica retenção, o que só acontece com o fundo desenquadrado (Lei 12.431 art. 3º §§5º e 6º; Lei 11.478 art. 2º §§4º e 5º): a retenção é definitiva e a ficha continua a mesma, pelo líquido; o COIR avisa.
- `coir/engine/proventos.py` — módulo: …e FIP-IE (tipo FUNDO_INFRA): alíquota zero, exclusiva na fonte (Lei 12.431/2011 art. 3º §§1º, 2º e 2º-A; Lei 11.478/2007 art. 2º §3º); ficha exclusiva 06 pelo valor líquido. Preço x quantidade acima do líquido sugere…
- `docs/regras-fiscais.md` — R20: …(proposta de redução de capital); sem a retenção de 0,005% | Lei 12.431/2011 art. 3º §§1º b, 2º e 2º-A; Lei 11.478/2007 art. 2º §§1º I, 2º e 3º; IN RFB 1.585/2015 art. 16 §2º; IN RFB 1.585/2015 arts. 33 (§8º: sem…
- `tests/fiscal_cases/test_fundo_infra.py` — módulo: FI-Infra (Lei 12.431/2011 art. 3º) e FIP-IE (Lei 11.478/2007 art. 2º §1º I), tipo FUNDO_INFRA pela lista de fundos da B3: o ganho de pessoa física tem…
- `tests/fiscal_cases/test_fundo_infra.py` — test_rendimentos_distribuidos_vao_para_a_exclusiva_06: Parecer de 09/11/2026: rendimento de FI-Infra (Lei 12.431 art. 3º §§1º e 2º) e de FIP-IE (Lei 11.478 art. 2º §3º diz isento, mas o programa só aceita a exclusiva 06) na exclusiva 06, em linhas separadas;…
- `tests/fiscal_cases/test_proventos.py` — test_amortizacao_de_fi_infra_reduz_o_custo_e_a_de_fii_fica_fora: Lei 11.478/2007 art. 2º §2º; IN 1.585 art. 16 §2º: a amortização reduz o custo das cotas. A de FII continua fora (não foi objeto do parecer).

## Lei 12.350/2010

- `data/parametros_fiscais.csv` — day_trade_mesma_corretora: MP 497/2010 art. 21 (conv. Lei 12.350/2010): mesma corretora

## Lei 12.431/2011

- `coir/engine/proventos.py` — _rendimento_fundo_infra: …IRRF. Preço x quantidade acima do líquido indica retenção, o que só acontece com o fundo desenquadrado (Lei 12.431 art. 3º §§5º e 6º; Lei 11.478 art. 2º §§4º e 5º): a retenção é definitiva e a ficha continua a mesma,…
- `coir/engine/proventos.py` — módulo: …(parâmetro). - Rendimento de FI-Infra e FIP-IE (tipo FUNDO_INFRA): alíquota zero, exclusiva na fonte (Lei 12.431/2011 art. 3º §§1º, 2º e 2º-A; Lei 11.478/2007 art. 2º §3º); ficha exclusiva 06 pelo valor líquido.…
- `docs/regras-fiscais.md` — R20: …mostra retenção; amortização reduz o custo (proposta de redução de capital); sem a retenção de 0,005% | Lei 12.431/2011 art. 3º §§1º b, 2º e 2º-A; Lei 11.478/2007 art. 2º §§1º I, 2º e 3º; IN RFB 1.585/2015 art. 16…
- `tests/fiscal_cases/test_fundo_infra.py` — módulo: FI-Infra (Lei 12.431/2011 art. 3º) e FIP-IE (Lei 11.478/2007 art. 2º §1º I), tipo FUNDO_INFRA pela lista de fundos da B3: o…
- `tests/fiscal_cases/test_fundo_infra.py` — test_bens_e_direitos_e_ficha_do_ganho_na_dirpf: …em 07.10 e FIP-IE em 07.07, com o CNPJ do fundo; o ganho na ficha Exclusiva/Definitiva, código 06 (Lei 12.431 art. 3º §2º: "exclusivamente na fonte"), não em isentos.
- `tests/fiscal_cases/test_fundo_infra.py` — test_rendimentos_distribuidos_vao_para_a_exclusiva_06: Parecer de 09/11/2026: rendimento de FI-Infra (Lei 12.431 art. 3º §§1º e 2º) e de FIP-IE (Lei 11.478 art. 2º §3º diz isento, mas o programa só aceita a…
- `tests/fiscal_cases/test_proventos.py` — test_rendimento_de_fi_infra_com_retencao_avisa: Fundo desenquadrado retém na fonte (Lei 12.431 art. 3º §§5º e 6º): preço x quantidade acima do líquido vira aviso; o lançamento continua pelo líquido.
- `tests/fiscal_cases/test_proventos.py` — test_rendimento_de_fi_infra_vai_para_a_exclusiva_sem_irrf: Lei 12.431/2011 art. 3º §§1º e 2º: alíquota zero, exclusivamente na fonte. 100 cotas x 1,2345 = 123,45 líquido = bruto; não é rendimento de FII (isentos).

## Lei 13.043/2014

- `coir/engine/movimentos.py` — _fim_da_isencao: Lei 13.043/2014 art. 17 §6º e IN 1.585 art. 68: no dia seguinte ao fim da isenção, o lote isento sobe à média dos…
- `coir/engine/posicao.py` — fim_da_isencao: Lei 13.043/2014 art. 17 §6º: o lote isento sobe ao maior entre o custo pago e quantidade x média dos 30 pregões, e os lotes se fundem num PM só (art. 58). Devolve o aumento.
- `coir/engine/proventos.py` — _origem_reembolso: …é a quantidade emprestada vezes o valor por ação do provento do mesmo dia e empresa, menos o IR do JCP (Lei 13.043/2014 art. 7º §1º II; Lei 15.265/2025 art. 20). Só uma origem com quantidade inteira identifica;…
- `data/irpf/2025.yaml` — exclusivos: Programa IRPF 2026, ficha Rendimentos sujeitos à tributação exclusiva/definitiva; Lei 9.249/1995 art. 9º; Lei 13.043/2014 art. 2º; Lei 8.981/1995 art. 65 §4º c
- `data/irpf/2026.yaml` — exclusivos: Programa IRPF 2026, ficha Rendimentos sujeitos à tributação exclusiva/definitiva; Lei 9.249/1995 art. 9º; Lei 13.043/2014 art. 2º; Lei 8.981/1995 art. 65 §4º c
- `data/irpf/2027.yaml` — exclusivos: Programa IRPF 2026, ficha Rendimentos sujeitos à tributação exclusiva/definitiva; Lei 9.249/1995 art. 9º; Lei 13.043/2014 art. 2º; Lei 8.981/1995 art. 65 §4º c
- `data/isencao_pme.csv` — AGRO3: Lei 13.043/2014 art. 16; IN RFB 1.585/2015 arts. 66 a 69-A; lista B3 (BrasilAgro); oferta subsequente precificada em 03/02/2021, anunciada em 04/02/2021, novas ações negociadas a partir de 05/02/2021; compras até 04/02/2021 (imprensa e o mesmo critério da PRIO; sem documento aberto da companhia); média de fim: COTAHIST da B3, 30 pregões de 16/11 a 28/12/2023, fechamento ponderado pela quantidade (cálculo do COIR; sem valor oficial publicado)
- `data/isencao_pme.csv` — CRDE3: Lei 13.043/2014 art. 16; IN RFB 1.585/2015 arts. 66 a 69-A; lista B3 (CR2 Empreendimentos); sem oferta subsequente conhecida: isenta até o fim do prazo (confirmação por fontes secundárias); sem negócio nos 30 pregões finais de 2023
- `data/isencao_pme.csv` — FRTA3: Lei 13.043/2014 art. 16; IN RFB 1.585/2015 arts. 66 a 69-A; lista B3 (Renar Maçãs, RNAR3, hoje Pomifrutas); sem oferta subsequente conhecida: isenta até o fim do prazo (confirmação por fontes secundárias); média de fim: COTAHIST da B3, 30 pregões de 16/11 a 28/12/2023, fechamento ponderado pela quantidade (cálculo do COIR; sem valor oficial publicado)
- `data/isencao_pme.csv` — GSHP3: Lei 13.043/2014 art. 16; IN RFB 1.585/2015 arts. 66 a 69-A; lista B3 (General Shopping); sem oferta subsequente conhecida: isenta até o fim do prazo (confirmação por fontes secundárias); média de fim: COTAHIST da B3, 30 pregões de 16/11 a 28/12/2023, fechamento ponderado pela quantidade (cálculo do COIR; sem valor oficial publicado); só 4 pregões com negócio
- `data/isencao_pme.csv` — NUTR3: Lei 13.043/2014 art. 16; IN RFB 1.585/2015 arts. 66 a 69-A; lista B3 (Nutriplant); sem oferta subsequente conhecida: isenta até o fim do prazo (confirmação por fontes secundárias); média de fim: COTAHIST da B3, 30 pregões de 16/11 a 28/12/2023, fechamento ponderado pela quantidade (cálculo do COIR; sem valor oficial publicado)
- `data/isencao_pme.csv` — PRIO3: Lei 13.043/2014 art. 16; IN RFB 1.585/2015 arts. 66 a 69-A; lista B3 (HRT Participações, HRTP3); RI PRIO: isentas as compradas entre 10/07/2014 e 29/01/2021, e as adquiridas a partir de 01/02/2021 não (follow-on precificado em 28/01/2021); média de fim: COTAHIST da B3, 30 pregões de 16/11 a 28/12/2023, fechamento ponderado pela quantidade (cálculo do COIR; sem valor oficial publicado)
- `data/isencao_pme.csv` — PRNR3: Lei 13.043/2014 art. 16; IN RFB 1.585/2015 arts. 66 a 69-A; fora da tabela de 2014 da B3; citada por fontes secundárias como enquadrada; sem oferta subsequente conhecida: isenta até o fim do prazo (confirmação por fontes secundárias); média de fim: COTAHIST da B3, 30 pregões de 16/11 a 28/12/2023, fechamento ponderado pela quantidade (cálculo do COIR; sem valor oficial publicado)
- `data/isencao_pme.csv` — SQIA3: Lei 13.043/2014 art. 16; IN RFB 1.585/2015 arts. 66 a 69-A; lista B3 (Senior Solution, SNSL3); RI Sinqia, citado por fontes secundárias: isentas as compradas entre 10/07/2014 e 05/09/2019 (oferta deliberada em 05/09/2019, Formulário de Referência 2019); sem negócio nos 30 pregões finais de 2023
- `data/parametros_fiscais.csv` — reembolso_btc_dividendos_isento: IN RFB 1.585/2015 art. 74 §1º I a (Lei 13.043/2014 art. 7º)
- `docs/regras-fiscais.md` — Decisões da F4 (motor fiscal, 14-15/10/2026): **ETF de renda fixa:** fora da apuração e do DARF; ganho listado para a DIRPF (Lei 13.043/2014).
- `docs/regras-fiscais.md` — Decisões da F4 (motor fiscal, 14-15/10/2026): Isenção de PME (Lei 13.043/2014 art. 16; IN arts. 59 III e 66 a 69-A), `data/isencao_pme.csv`:
- `docs/regras-fiscais.md` — Decisões da F4 (motor fiscal, 14-15/10/2026): isenta, com aviso de que a compra antes de 10/07/2014 não é isenta (Lei 13.043 art. 17 caput);
- `docs/regras-fiscais.md` — Decisões da F4 (motor fiscal, 14-15/10/2026): nos 30 pregões anteriores a 31/12/2023 (Lei 13.043 art. 17 §6º; IN art. 68; texto A), e os lotes
- `docs/regras-fiscais.md` — Decisões da validação da extração (F2, 14/10/2026): **ETF de renda fixa (ex.: B5P211):** tributação exclusiva na fonte (Lei 13.043/2014 art. 2º §4º). Na
- `docs/regras-fiscais.md` — Outros custos dedutíveis (aba BTC_Outros do v2): **Corretagem e emolumentos do tomador:** até 2025, custo da operação (Lei 13.043/2014 art. 9º par. único; IN 1.585
- `docs/regras-fiscais.md` — Proventos (F7a, 19/10/2026): Reembolso de BTC ao doador: isento até 2025 (Lei 13.043/2014 art. 7º, reproduzido na IN 1.585
- `tests/fiscal_cases/test_casos_fiscais.py` — test_etf_de_renda_fixa_fica_fora_da_apuracao: Lei 13.043/2014 arts. 2º a 4º: ETF de renda fixa tem tributação exclusiva na fonte; o ganho vai para a lista da DIRPF, não para o DARF nem para a isenção.
- `tests/fiscal_cases/test_casos_fiscais.py` — test_pme_ganho_isento_e_fora_do_limite: Lei 13.043/2014 art. 16; IN 1.585 arts. 66 e 67 §3º: ganho na venda de ações de PME é isento e o valor não entra…
- `tests/fiscal_cases/test_oferta_publica.py` — test_oferta_de_pme_tem_o_ganho_isento_e_fora_do_limite: IN RFB 1.585/2015 art. 67 I e Lei 13.043/2014 art. 16: ações de PME adquiridas em oferta. 1.000 a 27,00 (27.000,00) vendidas a 30,00: ganho de…
- `tests/fiscal_cases/test_outros_custos.py` — módulo: …ao doador não é dedutível por pessoa física (SC Cosit 42/2020); corretagem e emolumentos do tomador: Lei 13.043/2014 art. 9º par. único, IN RFB 1.585/2015 art. 76 par. único e, desde 2026, art. 56 §3º.
- `tests/fiscal_cases/test_pme_lotes.py` — módulo: Lotes isento e tributado de ações de PME (parecer do especialista de 08/11/2026). Norma lida: - Lei 13.043/2014 art. 16 caput (isenção até 31/12/2023, mercado à vista de bolsa) e art. 17: caput (aquisição a…
- `tests/fiscal_cases/test_proventos.py` — test_reembolso_com_origem_informada_pelo_usuario: …usuário, pelo informe ou extrato, prevalece sobre a identificada pelo valor; o tratamento é o da origem (Lei 13.043/2014 art. 7º; Lei 15.265/2025 arts. 20 e 21). Mesmo caso do JCP de 0,520530743 por ação (442,45 =…
- `tests/fiscal_cases/test_proventos.py` — test_reembolso_de_jcp_identificado_pelo_valor_por_acao: …o reembolso é a quantidade emprestada vezes o valor por ação do provento do dia, menos o IR do JCP (Lei 13.043/2014 art. 7º §1º II). JCP de 0,520530743 por ação pago em 14/03/2025 (pacote da B3), 1.000 ações…

## Lei 14.754/2023

- `coir/engine/proventos.py` — módulo: …(ação, unit): isento (Lei 9.249/1995 art. 10). De BDR, ETF ou FII: fora da v3.0, listado com aviso (BDR: Lei 14.754/2023, não é isento). - JCP: tributação exclusiva; a B3 credita o líquido e o preço unitário é o bruto…
- `docs/regras-fiscais.md` — Proventos (F7a, 19/10/2026): "Dividendo": fora da v3.0, listados com aviso (BDR não é isento, Lei 14.754/2023).
- `docs/regras-fiscais.md` — R16: R16 | Rendimentos de FII isentos se o fundo cumpre requisitos | Usuário marca conforme informe do fundo | Lei 11.033/2004 art. 3º (alt. Lei 14.754/2023) | M

## Lei 15.265/2025

- `coir/engine/proventos.py` — _origem_reembolso: …o valor por ação do provento do mesmo dia e empresa, menos o IR do JCP (Lei 13.043/2014 art. 7º §1º II; Lei 15.265/2025 art. 20). Só uma origem com quantidade inteira identifica; nenhuma ou mais de uma: ("", 0), com…
- `coir/engine/proventos.py` — módulo: …com aviso. - Reembolso de BTC (doador): isento até 2025 (IN RFB 1.585/2015 art. 74); desde 2026 a Lei 15.265/2025 é omissa para dividendos: isento por equivalência, com aviso (parâmetro). - Rendimento de…
- `data/parametros_fiscais.csv` — reembolso_btc_dividendos_isento: Lei 15.265/2025 arts. 20 e 21 omissos para dividendos: isento por equivalência (decisão do usuário 19/10/2026; rever se a RFB se manifestar)
- `data/parametros_fiscais.csv` — reembolso_btc_norma_omissa: Lei 15.265/2025 arts. 20 e 21 omissos para dividendos: a tela de proventos avisa (isento por equivalência)
- `docs/regras-fiscais.md` — Outros custos dedutíveis (aba BTC_Outros do v2): **Doador:** não lança nada; a remuneração tem IRRF definitivo (Lei 15.265 art. 19 caput e §4º I).
- `docs/regras-fiscais.md` — Outros custos dedutíveis (aba BTC_Outros do v2): **Mês:** o da recompra, ou o da liquidação do empréstimo sem recompra (Lei 15.265 art. 29 §§2º a 4º); estender ao custo é
- `docs/regras-fiscais.md` — Outros custos dedutíveis (aba BTC_Outros do v2): **Remuneração paga ao doador:** não dedutível por pessoa física (SC Cosit 42/2020; o art. 19 §2º da Lei 15.265/2025
- `docs/regras-fiscais.md` — Outros custos dedutíveis (aba BTC_Outros do v2): art. 76 par. único). Desde 2026 a Lei 15.265 revogou essa frase; decisão: manter pela regra geral (IN 1.585 art. 56 §3º).
- `docs/regras-fiscais.md` — Proventos (F7a, 19/10/2026): 2025 e, desde 2026, sem incidência expressa (Lei 15.265 art. 21 I). Declarado no ano do
- `docs/regras-fiscais.md` — Proventos (F7a, 19/10/2026): art. 74); desde 2026 a Lei 15.265/2025 é omissa para dividendos: isento por equivalência, com
- `docs/regras-fiscais.md` — Proventos (F7a, 19/10/2026): reteria (art. 7º §1º II; Lei 15.265 art. 20), mas o desconto não é IRRF do doador; é isento até
- `tests/fiscal_cases/test_proventos.py` — módulo: …ou do crédito); Lei 11.033/2004 art. 3º III (rendimentos de FII isentos); IN RFB 1.585/2015 art. 74 e Lei 15.265/2025 arts. 20 e 21 (reembolso de BTC); IN 1.585 art. 58 §8º (restituição de capital reduz o custo);…
- `tests/fiscal_cases/test_proventos.py` — test_p4_reembolso_de_btc_isento_com_aviso_desde_2026: Doador de 1.000 ações: reembolso de dividendo de 1.200,00. Em 2025 isento (IN 1.585 art. 74); em 2026 isento por equivalência (Lei 15.265/2025 omissa), com aviso; o PM não muda.
- `tests/fiscal_cases/test_proventos.py` — test_reembolso_com_origem_informada_pelo_usuario: …extrato, prevalece sobre a identificada pelo valor; o tratamento é o da origem (Lei 13.043/2014 art. 7º; Lei 15.265/2025 arts. 20 e 21). Mesmo caso do JCP de 0,520530743 por ação (442,45 = 1.000 ações): (a) informada…
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
- `coir/engine/movimentos.py` — _antecipar_ofertas: …entra antes dela, e o ganho fica no mês do pregão da venda (decisão do usuário de 08/11/2026; IN RFB 1.585/2015 art. 58). Antecipa só o que cobre a venda, da oferta mais próxima para a mais distante. A posição…
- `coir/engine/movimentos.py` — _cisao: Cisão: o percentual do custo de A divulgado pela companhia vai para B, sem resultado (IN RFB 1.585/2015 art. 58 §6º: na proporção fixada pela assembleia). A quantidade de A não muda. Os eventos das…
- `coir/engine/movimentos.py` — _day_trade: …marcou D formam o day trade quando a marcação é consistente; senão, a ordem dos negócios impressa (IN RFB 1.585/2015 art. 65 §3º).
- `coir/engine/movimentos.py` — _direito_nao_exercido: …emolumentos, pelo PM) vira perda em operações comuns na data do evento, o fim do prazo de exercício (IN RFB 1.585/2015 art. 60 §3º, por analogia com a opção que vira pó; art. 64 caput; SC Cosit 166/2021; parecer do…
- `coir/engine/movimentos.py` — _direitos_recebidos: …pagar nada. A parte da venda que passa da posição (e das compras do dia) entra antes, com custo zero (IN RFB 1.585/2015 art. 58 §7º, por analogia: a norma não trata do direito recebido; parecer do especialista de…
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
- `data/isencao_pme.csv` — AGRO3: Lei 13.043/2014 art. 16; IN RFB 1.585/2015 arts. 66 a 69-A; lista B3 (BrasilAgro); oferta subsequente precificada em 03/02/2021, anunciada em 04/02/2021, novas ações negociadas a partir de 05/02/2021; compras até 04/02/2021 (imprensa e o mesmo critério da PRIO; sem documento aberto da companhia); média de fim: COTAHIST da B3, 30 pregões de 16/11 a 28/12/2023, fechamento ponderado pela quantidade (cálculo do COIR; sem valor oficial publicado)
- `data/isencao_pme.csv` — CRDE3: Lei 13.043/2014 art. 16; IN RFB 1.585/2015 arts. 66 a 69-A; lista B3 (CR2 Empreendimentos); sem oferta subsequente conhecida: isenta até o fim do prazo (confirmação por fontes secundárias); sem negócio nos 30 pregões finais de 2023
- `data/isencao_pme.csv` — FRTA3: Lei 13.043/2014 art. 16; IN RFB 1.585/2015 arts. 66 a 69-A; lista B3 (Renar Maçãs, RNAR3, hoje Pomifrutas); sem oferta subsequente conhecida: isenta até o fim do prazo (confirmação por fontes secundárias); média de fim: COTAHIST da B3, 30 pregões de 16/11 a 28/12/2023, fechamento ponderado pela quantidade (cálculo do COIR; sem valor oficial publicado)
- `data/isencao_pme.csv` — GSHP3: Lei 13.043/2014 art. 16; IN RFB 1.585/2015 arts. 66 a 69-A; lista B3 (General Shopping); sem oferta subsequente conhecida: isenta até o fim do prazo (confirmação por fontes secundárias); média de fim: COTAHIST da B3, 30 pregões de 16/11 a 28/12/2023, fechamento ponderado pela quantidade (cálculo do COIR; sem valor oficial publicado); só 4 pregões com negócio
- `data/isencao_pme.csv` — NUTR3: Lei 13.043/2014 art. 16; IN RFB 1.585/2015 arts. 66 a 69-A; lista B3 (Nutriplant); sem oferta subsequente conhecida: isenta até o fim do prazo (confirmação por fontes secundárias); média de fim: COTAHIST da B3, 30 pregões de 16/11 a 28/12/2023, fechamento ponderado pela quantidade (cálculo do COIR; sem valor oficial publicado)
- `data/isencao_pme.csv` — PRIO3: Lei 13.043/2014 art. 16; IN RFB 1.585/2015 arts. 66 a 69-A; lista B3 (HRT Participações, HRTP3); RI PRIO: isentas as compradas entre 10/07/2014 e 29/01/2021, e as adquiridas a partir de 01/02/2021 não (follow-on precificado em 28/01/2021); média de fim: COTAHIST da B3, 30 pregões de 16/11 a 28/12/2023, fechamento ponderado pela quantidade (cálculo do COIR; sem valor oficial publicado)
- `data/isencao_pme.csv` — PRNR3: Lei 13.043/2014 art. 16; IN RFB 1.585/2015 arts. 66 a 69-A; fora da tabela de 2014 da B3; citada por fontes secundárias como enquadrada; sem oferta subsequente conhecida: isenta até o fim do prazo (confirmação por fontes secundárias); média de fim: COTAHIST da B3, 30 pregões de 16/11 a 28/12/2023, fechamento ponderado pela quantidade (cálculo do COIR; sem valor oficial publicado)
- `data/isencao_pme.csv` — SQIA3: Lei 13.043/2014 art. 16; IN RFB 1.585/2015 arts. 66 a 69-A; lista B3 (Senior Solution, SNSL3); RI Sinqia, citado por fontes secundárias: isentas as compradas entre 10/07/2014 e 05/09/2019 (oferta deliberada em 05/09/2019, Formulário de Referência 2019); sem negócio nos 30 pregões finais de 2023
- `data/parametros_fiscais.csv` — day_trade_criterio: IN RFB 1.585/2015 art. 65 §3º: casar pela ordem cronológica dos negócios; a nota não traz horário e a corretora (que o tem e retém o IRRF) marca D: vale a marcação D quando consistente (quantidade D de compra = de venda = menor das duas no dia); sem marcação ou inconsistente, a ordem impressa (decisão do usuário de 25/10/2026, parecer do especialista). Alternativa: ordem_impressa
- `data/parametros_fiscais.csv` — direito_recebido_custo_zero: IN RFB 1.585/2015 art. 58 §7º I a III (analogia: a norma não trata do direito recebido); parecer do especialista de 09/11/2026 seguido pelo usuário
- `data/parametros_fiscais.csv` — direito_vencido_data: IN RFB 1.585/2015 art. 60 §3º (analogia: perda na data do vencimento); Lei 6.404 art. 171 §4º (prazo decadencial): o direito comprado não exercido vira perda no fim do prazo de subscrição do pacote da B3, ou na data da baixa na B3 sem ele; alternativa baixa_b3 (parecer do especialista de 09/11/2026 seguido pelo usuário)
- `data/parametros_fiscais.csv` — janela_liquidacao_oferta_dias: Sem norma: critério do COIR para venda antes da liquidação da oferta (decisão do usuário de 08/11/2026); a aquisição vem antes da venda (IN RFB 1.585/2015 art. 58)
- `data/parametros_fiscais.csv` — leilao_fracao_isentavel: sem norma específica: fração vendida em leilão é alienação de ação no mercado à vista (IN RFB 1.585/2015 art. 58 caput); entra no limite de R$ 20 mil do mês do crédito (decisão do usuário 19/10/2026)
- `data/parametros_fiscais.csv` — pme_lote_ordem_venda: IN RFB 1.585/2015 art. 69-A §4º manda controlar os lotes em separado e não fixa a ordem; isento_primeiro ou tributado_primeiro (decisão do usuário de 08/11/2026)
- `data/parametros_fiscais.csv` — reembolso_btc_dividendos_isento: IN RFB 1.585/2015 art. 74 §1º I a (Lei 13.043/2014 art. 7º)
- `data/parametros_fiscais.csv` — reembolso_btc_norma_omissa: IN RFB 1.585/2015 art. 74 §1º I a (reembolso de dividendos de BTC expressamente isento)
- `data/trocas_ticker.csv` — PETZ3: B3 COTAHIST: último negócio de PETZ3 em 02/01/2026; AUAU3 desde 05/01/2026 (RI: relação 1:1, abertura deduzida de R$ 0,71); CNPJ muda (18328118000109 -> 53153938000108): incorporação tratada como troca curada, custo mantido (IN RFB 1.585/2015 art. 58 §6º); R$ 0,71 por ação à parte: resgate de PN, SC Cosit 37/2020
- `docs/regras-fiscais.md` — Decisões da F4 (motor fiscal, 14-15/10/2026): A IN RFB 1.585/2015 art. 56 §§3º e 4º admite deduzir os custos necessários e permite ratear os do
- `docs/regras-fiscais.md` — Decisões da F4 (motor fiscal, 14-15/10/2026): é isento (IN RFB 1.585/2015 art. 59 §2º IV); a venda entra na soma do limite do mês dela.
- `docs/regras-fiscais.md` — Declaração anual (F8, 20/10/2026): Bens e Direitos pelo custo de aquisição (IN RFB 1.585/2015 art. 58), também no ano da compra; a
- `docs/regras-fiscais.md` — Declaração anual (F8, 20/10/2026): Termo (IN RFB 1.585/2015 art. 62 I e §1º): a venda à vista na data da liquidação vai para
- `docs/regras-fiscais.md` — Futuros de commodities e conferência das notas BM&F (F7a, 19/10/2026): O resultado vem dos ajustes em R$ impressos na nota (IN RFB 1.585/2015 art. 61). O multiplicador só
- `docs/regras-fiscais.md` — Oferta pública (IPO e follow-on; parecer do especialista, 08/11/2026): efetivamente pago, com preço médio ponderado, sem resultado e sem IRRF (IN RFB 1.585/2015 art. 58
- `docs/regras-fiscais.md` — R1: R1 | Operações comuns (ações, units, ETF, BDR, opções, futuros, termo): 15% sobre ganho líquido mensal | Categoria COMUM, base única com subtotais por mercado | IN RFB 1.585/2015 | A
- `docs/regras-fiscais.md` — R10: R10 | Opção que vira pó: prêmio pago = perda (titular); recebido = ganho (lançador) no vencimento | Evento automático de vencimento (data como parâmetro, não "hoje") | IN RFB 1.585/2015 | A
- `docs/regras-fiscais.md` — R11: …por contrato e realizados no encerramento ou vencimento | Lei 11.051/2004 art. 32 (caput e §2º II); IN RFB 1.585/2015 art. 61; P&R IRPF 2026 pergunta 728 | A
- `docs/regras-fiscais.md` — R18: R18 | Termo: custo/alienação pelo preço a termo; ganho apurado na liquidação/encerramento (operação comum) | Contrato a termo vinculado à liquidação | IN RFB 1.585/2015 | M (validar com as operações a termo da planilha)
- `docs/regras-fiscais.md` — R19: …nunca FII; parâmetro `direito_recebido_custo_zero` | SC Cosit 145/2021 e 166/2021 (ementas oficiais); IN RFB 1.585/2015 arts. 58 §7º (custo zero por analogia), 59 e 65; recibo: posição B do parecer de 09/11/2026,…
- `docs/regras-fiscais.md` — R2: R2 | Day trade: 20%; prejuízo só compensa com ganho DT | Categoria DAY_TRADE; excedente de quantidade vira COMUM | IN RFB 1.585/2015 | A
- `docs/regras-fiscais.md` — R20: …de 0,005% | Lei 12.431/2011 art. 3º §§1º b, 2º e 2º-A; Lei 11.478/2007 art. 2º §§1º I, 2º e 3º; IN RFB 1.585/2015 art. 16 §2º; IN RFB 1.585/2015 arts. 33 (§8º: sem retenção), 34, 63 e 64; Manual MIR; P&R IRPF…
- `docs/regras-fiscais.md` — R21: …ou `baixa_b3`); baixa o menor entre a quantidade e a posição, com aviso se sobrar posição | IN RFB 1.585/2015 art. 60 §3º (analogia com a opção que vira pó), arts. 56 §3º, 58 e 64; SC Cosit 166/2021; Lei…
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
- `tests/fiscal_cases/test_day_trade_marcacao.py` — módulo: Day trade pela marcação D da corretora (F9, decisão do usuário de 25/10/2026). IN RFB 1.585/2015 art. 65 §3º: o day trade casa "pela ordem" a 1ª compra com a 1ª venda, isto é, na ordem…
- `tests/fiscal_cases/test_day_trade_parcial.py` — módulo: …dela fecha no dia (parecer do especialista fiscal de 07/11/2026; números conferidos à mão). Fontes (IN RFB 1.585/2015): - art. 65 §1º: day trade é a operação iniciada e encerrada no mesmo dia, com o mesmo ativo, na…
- `tests/fiscal_cases/test_direito_nao_exercido.py` — módulo: …vira perda em operações comuns na data do fim do prazo de exercício, fora do limite de R$ 20 mil (IN RFB 1.585/2015 art. 60 §3º, por analogia com a opção que vira pó; arts. 56 §3º, 58 e 64; SC Cosit 166/2021; Lei…
- `tests/fiscal_cases/test_direitos_subscricao.py` — módulo: …soma do limite de R$ 20 mil (SC Cosit 145/2021 e 166/2021, ementas oficiais; Lei 11.033/2004 art. 2º; IN RFB 1.585/2015 arts. 58, 59 e 65). O direito recebido sem nota de compra entra com custo zero (IN 1.585 art. 58…
- `tests/fiscal_cases/test_eventos_societarios.py` — módulo: …de fração (números conferidos à mão no parecer do especialista fiscal de 19/10/2026). Fontes: - IN RFB 1.585/2015 art. 58 §6º: na substituição de ações por incorporação, fusão ou cisão, o custo das ações…
- `tests/fiscal_cases/test_fundo_infra.py` — módulo: …fora da apuração mensal, do limite de R$ 20 mil e do DARF, e o prejuízo não compensa outros ganhos (IN RFB 1.585/2015 art. 64 só admite as perdas dos arts. 27, 58 e 60 a 62). Parecer do especialista de 09/11/2026,…
- `tests/fiscal_cases/test_futuros_commodities.py` — módulo: …e na conferência do câmbio implícito. Casos fechados à mão pelo especialista (19/10/2026). Fontes: IN RFB 1.585/2015 art. 61 (resultado = soma dos ajustes no encerramento), art. 65 (15% e 20% day trade, inclusive…
- `tests/fiscal_cases/test_irpf_declaracao.py` — módulo: …Fontes: P&R IRPF 2026 perguntas 707, 720 (um item por ação, total das corretoras, custo médio) e 715; IN RFB 1.585/2015 arts. 58 §1º (bonificação: custo atribuído), 63 §8º III (saldo de IRRF de dezembro na declaração)…
- `tests/fiscal_cases/test_irpf_linhas.py` — módulo: …"Mercado a termo - ações/ouro"); P&R IRPF 2026 pergunta 707 (cotas de fundo de índice não são ações); IN RFB 1.585/2015 art. 62 I e §1º (ganho do comprador a termo pela venda à vista na data da liquidação; vendido…
- `tests/fiscal_cases/test_irrf_fichas.py` — módulo: …Casos do especialista fiscal, conferidos à mão. Fontes: Lei 11.033/2004 art. 2º §§1º, 2º, 5º e 7º; IN RFB 1.585/2015 arts. 63 §8º e 65 §§8º e 9º; fichas "Renda variável - operações comuns/day trade" e "Operações -…
- `tests/fiscal_cases/test_irrf_fichas.py` — test_day_trade_usado_no_imposto_de_fii_gera_aviso: IN RFB 1.585/2015 art. 65 §8º I: o IRRF de day trade deduz o imposto do mês (saldo único da DARF), mas a ficha FII…
- `tests/fiscal_cases/test_irrf_fichas.py` — test_irrf_de_day_trade_nao_abate_fii_antes_do_0005: IN RFB 1.585/2015 art. 65 §8º I e ficha FII sem linha de day trade: dezembro/2026 com imposto só de FII (V 100…
- `tests/fiscal_cases/test_oferta_publica.py` — módulo: Compra em oferta pública (IPO e follow-on), que não aparece em nota de corretagem. IN RFB 1.585/2015 art. 58 caput: o custo é o valor pago, com preço médio ponderado; art. 56 §5º: ganho apurado na…
- `tests/fiscal_cases/test_oferta_publica.py` — test_oferta_de_pme_tem_o_ganho_isento_e_fora_do_limite: IN RFB 1.585/2015 art. 67 I e Lei 13.043/2014 art. 16: ações de PME adquiridas em oferta. 1.000 a 27,00 (27.000,00)…
- `tests/fiscal_cases/test_outros_custos.py` — módulo: …física (SC Cosit 42/2020); corretagem e emolumentos do tomador: Lei 13.043/2014 art. 9º par. único, IN RFB 1.585/2015 art. 76 par. único e, desde 2026, art. 56 §3º.
- `tests/fiscal_cases/test_petz_auau.py` — módulo: …em 23/01/2026. Valor exato pago: R$ 0,7109 por ação (imprensa, sobre o resultado do 4T25). Regra: - IN RFB 1.585/2015 art. 58 §6º: o custo das ações originais vai para as recebidas, e a ON e a PN foram recebidas.…
- `tests/fiscal_cases/test_pme_lotes.py` — módulo: …pelo volume negociado nos últimos 30 (trinta) pregões anteriores a 31 de dezembro de 2023"). - IN RFB 1.585/2015 art. 58 caput (média ponderada), arts. 66 a 68 (repetem a lei) e art. 69-A (incluído pela IN RFB…
- `tests/fiscal_cases/test_pme_lotes.py` — test_data_de_corte_das_ofertas_subsequentes: IN RFB 1.585/2015 art. 69-A: compras a partir da oferta subsequente são tributadas; §2º: as anteriores continuam isentas. A data é a que a companhia publicou, quando há.
- `tests/fiscal_cases/test_proventos.py` — módulo: …na data do pagamento ou do crédito); Lei 11.033/2004 art. 3º III (rendimentos de FII isentos); IN RFB 1.585/2015 art. 74 e Lei 15.265/2025 arts. 20 e 21 (reembolso de BTC); IN 1.585 art. 58 §8º (restituição de…
- `tests/fiscal_cases/test_raiz55.py` — módulo: …os mesmos. Na nota, a quantidade é a de ações e o preço é por 100 ações; o custo é o valor impresso (IN RFB 1.585/2015 art. 58 caput). A PNA continua ação para a isenção de R$ 20 mil (art. 59 caput).
- `tests/fiscal_cases/test_saldos_iniciais.py` — módulo: …parecer do especialista). Custódia inicial: a posição em 31/12 entra pelo custo de aquisição (IN RFB 1.585/2015 art. 58), o mesmo do Bens e Direitos, sem resultado e sem efeito no day trade (art. 65 §2º).…

## IN RFB 1.700/2017

- `docs/regras-fiscais.md` — Declaração anual (F8, 20/10/2026): 8º). Creditado = registrado como dívida com o acionista individualizado (IN RFB 1.700/2017 art.
- `tests/fiscal_cases/test_proventos.py` — test_jcp_com_data_com_em_janeiro_e_retencao_do_ano_anterior_so_avisa: …x 0,50 bruto. Crédito na data com (a mais tarde entre aprovação e data com: IN SRF 41/1998 art. 3º; IN RFB 1.700/2017 art. 75 §10). 3a: retenção de 17,5% (líquido 412,50): ano 2026, bruto 500,00, sem aviso. 3b:…

## IN RFB 1.916/2019

- `tests/fiscal_cases/test_pme_lotes.py` — módulo: …1.585/2015 art. 58 caput (média ponderada), arts. 66 a 68 (repetem a lei) e art. 69-A (incluído pela IN RFB 1.916/2019): caput (sem isenção para as ações adquiridas a partir da data da oferta subsequente), §1º…

## Res. CMN 4.880/2020

- `docs/regras-fiscais.md` — R12: …acumula, inclusive para o ano seguinte | `data/feriados.csv` (o último dia útil do ano não é útil: Res. CMN 4.880/2020 art. 3º) + acúmulo | RIR/2018 art. 915; Lei 9.430/1996 art. 68 caput e §1º | A

## SC Cosit 37/2020

- `data/trocas_ticker.csv` — PETZ3: B3 COTAHIST: último negócio de PETZ3 em 02/01/2026; AUAU3 desde 05/01/2026 (RI: relação 1:1, abertura deduzida de R$ 0,71); CNPJ muda (18328118000109 -> 53153938000108): incorporação tratada como troca curada, custo mantido (IN RFB 1.585/2015 art. 58 §6º); R$ 0,71 por ação à parte: resgate de PN, SC Cosit 37/2020
- `docs/regras-fiscais.md` — R8: …custo, não o dinheiro) e o resgate é alienação fora de bolsa, no GCAP, fora da apuração mensal (SC Cosit 37/2020; critério de rateio é inferência, confiança média; decisão delegada ao especialista em 07/11/2026).…
- `tests/fiscal_cases/test_petz_auau.py` — módulo: …- O resgate da PN é alienação: dinheiro menos o custo da PN (Lei 7.713/1988 art. 3º §3º; SC Cosit 37/2020, CETIP/BM&FBOVESPA, com a mesma estrutura de ON + PN resgatável). Fora de bolsa, vai para o GCAP e…

## SC Cosit 42/2020

- `docs/regras-fiscais.md` — Outros custos dedutíveis (aba BTC_Outros do v2): **Remuneração paga ao doador:** não dedutível por pessoa física (SC Cosit 42/2020; o art. 19 §2º da Lei 15.265/2025
- `tests/fiscal_cases/test_outros_custos.py` — módulo: …"Outros custos dedutíveis"). A remuneração paga ao doador não é dedutível por pessoa física (SC Cosit 42/2020); corretagem e emolumentos do tomador: Lei 13.043/2014 art. 9º par. único, IN RFB 1.585/2015 art. 76…

## SC Cosit 63/2020

- `docs/regras-fiscais.md` — Decisões da F4 (motor fiscal, 14-15/10/2026): a partir dela, "que serão tributadas"; §2º e SC Cosit 63/2020: as de antes continuam isentas. A
- `tests/fiscal_cases/test_pme_lotes.py` — módulo: …§1º (inclusive preferência e bonificação), §2º (as anteriores permanecem isentas), §3º (IRRF do art. 63 na venda isenta é compensável) e §4º ("O investidor deverá controlar, de forma separada, as ações…

## SC Cosit 67/2024

- `docs/regras-fiscais.md` — Proventos (F7a, 19/10/2026): o informe. Alternativa (acessório segue o principal, SC Cosit 67/2024) não adotada.

## SC Cosit 145/2021

- `docs/regras-fiscais.md` — Decisões da F4 (motor fiscal, 14-15/10/2026): Units fora da isenção: coincide com a RFB (SC Cosit 145/2021: só "ativos da espécie ações").
- `docs/regras-fiscais.md` — R19: …tipo do papel); categoria COMUM ou DAY_TRADE, nunca FII; parâmetro `direito_recebido_custo_zero` | SC Cosit 145/2021 e 166/2021 (ementas oficiais); IN RFB 1.585/2015 arts. 58 §7º (custo zero por analogia), 59 e 65;…
- `tests/fiscal_cases/test_direitos_subscricao.py` — módulo: …15% sobre o ganho líquido, 20% em day trade, sem isenção e fora da soma do limite de R$ 20 mil (SC Cosit 145/2021 e 166/2021, ementas oficiais; Lei 11.033/2004 art. 2º; IN RFB 1.585/2015 arts. 58, 59 e 65). O…

## SC Cosit 166/2021

- `coir/engine/movimentos.py` — _direito_nao_exercido: …de exercício (IN RFB 1.585/2015 art. 60 §3º, por analogia com a opção que vira pó; art. 64 caput; SC Cosit 166/2021; parecer do especialista de 09/11/2026). Fora do limite de R$ 20 mil: não há valor de alienação.…
- `docs/regras-fiscais.md` — R21: …posição | IN RFB 1.585/2015 art. 60 §3º (analogia com a opção que vira pó), arts. 56 §3º, 58 e 64; SC Cosit 166/2021; Lei 6.404 art. 171 §4º; P&R IRPF 2026 pergunta 727; parecer de 09/11/2026, seguido pelo usuário | M…
- `tests/fiscal_cases/test_direito_nao_exercido.py` — módulo: …20 mil (IN RFB 1.585/2015 art. 60 §3º, por analogia com a opção que vira pó; arts. 56 §3º, 58 e 64; SC Cosit 166/2021; Lei 6.404 art. 171 §4º). Parecer do especialista de 09/11/2026, seguido pelo usuário.

## SC Cosit 224/2014

- `data/parametros_fiscais.csv` — incorporacao_acoes_tratamento: Lei 6.404 art. 252: custo das ações entregues mantido nas recebidas (TRF1/TRF3/TRF4; decisão do usuário 19/10/2026). A RFB diverge (SC Cosit 224/2014: alienação, ganho de capital no GCAP); a tela de Eventos avisa
- `docs/regras-fiscais.md` — R8: …art. 58 caput, §§6º e 7º, art. 60 §2º; P&R IRPF (substituição de ações; incorporação de ações); SC Cosit 224/2014 (divergência); SC Cosit 37/2020 (resgate); Lei 6.404 arts. 227-229 e 252 | A
- `tests/fiscal_cases/test_eventos_societarios.py` — módulo: …6.404 art. 252): custo mantido (TRF1, TRF3, TRF4; decisão do usuário de 19/10/2026); a RFB diverge (SC Cosit 224/2014), parâmetro incorporacao_acoes_tratamento. - Ajuste de série de opção pela B3 (Manual de…
- `tests/fiscal_cases/test_petz_auau.py` — módulo: …incorporação em si segue o parâmetro incorporacao_acoes_tratamento (custo mantido); a RFB diverge (SC Cosit 224/2014), e a tela de Eventos avisa. O motor recebe só a repartição do custo, na data ex: REDUCAO_CAPITAL…

## P&R IRPF

- `docs/regras-fiscais.md` — Decisões da validação do gabarito (F1, 14/10/2026): 06/2008). R5 (P&R IRPF) diz que continua compensável; uma resposta de IA consultada pelo usuário diz o
- `docs/regras-fiscais.md` — Decisões da validação do gabarito (F1, 14/10/2026): P&R IRPF antes de implementar; enquanto isso, parâmetro configurável.
- `docs/regras-fiscais.md` — R8: …na linha do tempo do ativo (tela 7) | IN RFB 1.585/2015 art. 58 caput, §§6º e 7º, art. 60 §2º; P&R IRPF (substituição de ações; incorporação de ações); SC Cosit 224/2014 (divergência); SC Cosit 37/2020…
- `tests/fiscal_cases/test_casos_fiscais.py` — test_exercicio_parcial_e_sobra_vira_po_no_vencimento: R9/R10 (IN RFB 1.585/2015 art. 58; P&R IRPF, opções): titular de 1.000 calls exerce 600 no vencimento; as 400 restantes viram pó (perde 40% do prêmio de 1.000 = -400); o custo das ações leva 600 do prêmio.
- `tests/fiscal_cases/test_eventos_societarios.py` — módulo: …ou cisão, o custo das ações originais é atribuído às recebidas na proporção fixada pela assembleia; P&R IRPF (substituição de ações): não é alienação, a data de aquisição é a original. - Incorporação de ações (Lei…

## P&R IRPF 2026

- `coir/engine/futuros.py` — módulo: …ocasião da liquidação do contrato, cessão ou encerramento da posição". IN RFB 1.585/2015 art. 61 e P&R IRPF 2026 pergunta 728: resultado de cada contrato = soma algébrica dos ajustes diários entre a abertura e o…
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
- `docs/regras-fiscais.md` — Declaração anual (F8, 20/10/2026): P&R IRPF 2026 pergunta 723, que trata de lote isolado).
- `docs/regras-fiscais.md` — Declaração anual (F8, 20/10/2026): `opcoes_linha_acoes_tipos_objeto`; P&R IRPF 2026 pergunta 707 separa cotas de fundo de índice de
- `docs/regras-fiscais.md` — Outros custos dedutíveis (aba BTC_Outros do v2): **A conferir:** o PDF oficial da SC 42/2020 e o P&R IRPF 2026 não puderam ser lidos; as leis foram lidas em compiladores.
- `docs/regras-fiscais.md` — R11: …encerramento ou vencimento | Lei 11.051/2004 art. 32 (caput e §2º II); IN RFB 1.585/2015 art. 61; P&R IRPF 2026 pergunta 728 | A
- `docs/regras-fiscais.md` — R20: …1.585/2015 art. 16 §2º; IN RFB 1.585/2015 arts. 33 (§8º: sem retenção), 34, 63 e 64; Manual MIR; P&R IRPF 2026 perguntas 736 e 738; pareceres de 09/11/2026 | A (alíquota, retenção do FIP-IE, rendimento do FI-Infra…
- `docs/regras-fiscais.md` — R21: …com a opção que vira pó), arts. 56 §3º, 58 e 64; SC Cosit 166/2021; Lei 6.404 art. 171 §4º; P&R IRPF 2026 pergunta 727; parecer de 09/11/2026, seguido pelo usuário | M (sem norma específica)
- `docs/regras-fiscais.md` — R6: …| Lei 9.959/2000 art. 8º §§4º, 8º; IN RFB 1.585/2015 arts. 63 §§8º e 9º e 65 §§8º, 9º e 12; P&R IRPF 2026 perguntas 714 e 715; a página "Retenções" da RFB (19/11/2025) restringe o DT ao imposto de DT:…
- `tests/fiscal_cases/test_irpf_declaracao.py` — módulo: …Direitos, rendimentos e fichas mês a mês que o COIR preenche. Casos do especialista fiscal (parecer de 20/10/2026), conferidos à mão; custos zero. Fontes: P&R IRPF 2026 perguntas 707, 720 (um item por ação, total das…
- `tests/fiscal_cases/test_irpf_linhas.py` — módulo: …- operações comuns/day trade" do programa IRPF (F8). Casos do especialista fiscal (pareceres de 20 e 21/10/2026), conferidos à mão. Fontes: programa IRPF, ficha de renda variável (linhas "Mercado de opções -…

## Manual MIR

- `docs/regras-fiscais.md` — R20: …2º e 3º; IN RFB 1.585/2015 art. 16 §2º; IN RFB 1.585/2015 arts. 33 (§8º: sem retenção), 34, 63 e 64; Manual MIR; P&R IRPF 2026 perguntas 736 e 738; pareceres de 09/11/2026 | A (alíquota, retenção do FIP-IE, rendimento…
- `tests/fiscal_cases/test_fundo_infra.py` — test_bens_e_direitos_e_ficha_do_ganho_na_dirpf: Parecer de 09/11/2026 (Manual MIR; Guia IR 2026): FI-Infra em 07.10 e FIP-IE em 07.07, com o CNPJ do fundo; o ganho na ficha…

## Manual MIR 2026

- `data/parametros_fiscais.csv` — prejuizo_mes_isento_compensavel: Lei 8.981/1995 art. 72 §8º; Manual MIR 2026; decisão do usuário (14/10/2026)
- `docs/regras-fiscais.md` — R5: R5 | Prejuízo de mês isento continua compensável | Parâmetro `prejuizo_mes_isento_compensavel = true` | Lei 8.981/1995 art. 72 §8º; Manual MIR 2026; decisão do usuário (14/10/2026) | A
- `tests/fiscal_cases/test_casos_fiscais.py` — test_t1_prejuizo_em_mes_isento_compensa: R5: prejuízo de ações em mês isento continua compensável (Lei 8.981/1995 art. 72 §8º; Manual MIR 2026; decisão do usuário). Mar/24: ações -1.200 (vendas 15.000), opções -300. Abr/24: ganho 2.000 -> base…

## Programa IRPF

- `coir/engine/apuracao.py` — módulo: …a declaração anual (IN RFB 1.585/2015; decisão do usuário). A DARF usa um saldo único; as fichas do programa IRPF recebem o IRRF que cada uma consumiu (`fichas.py`). - DARF 6015: abaixo de R$ 10 acumula, inclusive para…
- `coir/engine/fichas.py` — módulo: IRRF por ficha do programa IRPF (comuns/day trade e FII), com a DARF no saldo único. A DARF mensal deduz o IRRF de um saldo único (Lei…
- `data/parametros_fiscais.csv` — opcoes_linha_acoes_tipos_objeto: programa IRPF, ficha Renda variável: "Mercado de opções - ações" para opção sobre ação; ETF, índice e outros em "Mercado de opções - outros" (P&R IRPF 2026 pergunta 707 distingue cotas de fundo de índice de ações); unit como ação, coerente com o Bens e Direitos do programa (código 03.01, ações inclusive units)
- `docs/regras-fiscais.md` — Decisões da F4 (motor fiscal, 14-15/10/2026): 03/11/2026):** um PDF por ano, na ordem da ficha de renda variável do programa IRPF, em
- `docs/regras-fiscais.md` — R17: R17 | Fichas/códigos da DIRPF mudam por ano | Mapeamento por ano em `data/irpf/<ano>.yaml` | programa IRPF do ano | revisar anualmente
- `docs/regras-fiscais.md` — Relatórios (F5, 16/10/2026): conferir no programa IRPF do ano: a ficha FII sem linha de IR fonte de day trade; ordem de
- `tests/fiscal_cases/test_irpf_linhas.py` — módulo: Linha da ficha "Renda variável - operações comuns/day trade" do programa IRPF (F8). Casos do especialista fiscal (pareceres de 20 e 21/10/2026), conferidos à mão. Fontes: programa…
- `tests/fiscal_cases/test_irrf_fichas.py` — módulo: IRRF por ficha do programa IRPF com a DARF no saldo único (decisão do usuário de 16/10/2026: separar o IRRF de FII só na declaração).…

## CTN

- `docs/regras-fiscais.md` — DARF paga em atraso: multa de mora (F6, 17/10/2026): A denúncia espontânea (CTN art. 138; STJ REsp 1.149.022/SP; Súmula 360) fica como nota na tela:
- `docs/regras-fiscais.md` — Declaração anual (F8, 20/10/2026): crédito decaiu (CTN art. 150 §4º, havendo pagamento; art. 173 I se houver dolo) e não se retifica
- `docs/regras-fiscais.md` — Declaração anual (F8, 20/10/2026): nem se paga (seria pagamento indevido, CTN art. 156 V); restituição de pagamento a maior, 5 anos do
- `docs/regras-fiscais.md` — Declaração anual (F8, 20/10/2026): pagamento (CTN art. 168 I). Decisão do usuário de 03/11/2026, com parecer do especialista, num
- `tests/fiscal_cases/test_pme_lotes.py` — test_data_da_custodia_decide_a_marca: A custódia inicial segue a janela do CSV pela data do evento. Dentro dela o COIR presume aquisição a partir de 10/07/2014 (com aviso); fora, não presume isenção (CTN art. 111 II).

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

## Instrução Normativa SRF 41/1998

- `coir/engine/proventos.py` — _jcp_com_credito: …da empresa ou, pela B3, a mais tarde entre a aprovação e a data com): declarado no ano do crédito (IN SRF 41/1998 art. 3º; Lei 9.249/1995 art. 9º §2º; a receber em 31/12 se pago depois: art. 3º parágrafo único), com…
- `docs/regras-fiscais.md` — Declaração anual (F8, 20/10/2026): como crédito contra a companhia (IN SRF 41/1998 art. 3º e parágrafo único; Lei 9.249/1995 art.
- `tests/fiscal_cases/test_irpf_declaracao.py` — test_jcp_pago_mais_de_um_ano_depois_do_credito_fica_a_receber_nos_dois_31_12: …com 26/12/2023 pago em 09/01/2025. Creditado e não pago: a receber em 31/12/2023 e em 31/12/2024 (IN SRF 41/1998 art. 3º parágrafo único); em 2025, de 425,00 para 0,00, sem rendimento novo.
- `tests/fiscal_cases/test_proventos.py` — test_jcp_com_data_com_em_janeiro_e_retencao_do_ano_anterior_so_avisa: …pago em 30/01/2026, 1.000 x 0,50 bruto. Crédito na data com (a mais tarde entre aprovação e data com: IN SRF 41/1998 art. 3º; IN RFB 1.700/2017 art. 75 §10). 3a: retenção de 17,5% (líquido 412,50): ano 2026, bruto…
- `tests/fiscal_cases/test_proventos.py` — test_jcp_creditado_no_ano_anterior_com_a_mesma_aliquota: …para perceber); com ela (pacote da B3 ou informe), declarado em 2024 e a receber em 31/12/2024 (IN SRF 41/1998 art. 3º e parágrafo único; Lei 9.249/1995 art. 9º §2º).
