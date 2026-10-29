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
| CVM, documentos IPE (dados abertos) | `https://dados.cvm.gov.br/dados/CIA_ABERTA/DOC/IPE/DADOS/ipe_cia_aberta_AAAA.zip` e o documento em `https://www.rad.cvm.gov.br/ENET/frmDownloadDocumento.aspx` | Fatos relevantes e avisos aos acionistas: proporção, data-base, data ex e custo atribuído da bonificação. Usados para 2023 em diante, quando o FRE parou e a B3 já tinha substituído o evento. | `tools/referencias/eventos_cvm_ipe.csv` (transcrito, com o protocolo); `origem` = `cvm_ipe` |
| Banco Central, série 4390 (Selic mensal) | `https://api.bcb.gov.br/dados/serie/bcdata.sgs.4390/dados?formato=json` | Taxa Selic de cada mês encerrado. | `selic_mensal.csv` |

Regras da combinação (`tools/referencias/gerar.py`, `historico.py`):

- nada é apagado: a linha que uma fonte deixou de mostrar continua no pacote;
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
  publica; quando a B3 só guardou uma delas, a outra vem do documento (ITUB 2011, BEES 2012);
- a unit entra com o fator das ações só quando todas as espécies tiveram o mesmo fator e a mesma
  data com e a cotação da unit confirma.
