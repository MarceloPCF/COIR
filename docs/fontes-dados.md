## Fontes dos dados públicos (pacote de referências)

As tabelas geradas em `data/` (o pacote que vai no zip do programa e no release
`referencias-AAAA.MM.DD`) vêm só de fontes públicas, sem login. Cada linha de
`data/eventos_b3.csv` diz de onde veio (coluna `origem`, e `referencia` com o documento).

| Fonte | URL | O que cobre | Onde entra |
|---|---|---|---|
| B3, consulta de empresas listadas (`GetListedSupplementCompany`) | `https://sistemaswebb3-listados.b3.com.br/listedCompaniesProxy/CompanyCall/GetListedSupplementCompany/` | Desdobramento, grupamento, bonificação, cisão, resgate e subscrição com data com e fator exatos; proventos em dinheiro recentes. **Só o último evento de cada tipo por ISIN**: o histórico se perde quando a companhia repete o tipo de evento. | coleta semanal; `origem` = `b3` |
| B3, proventos em dinheiro (`GetListedCashDividends`) | mesmo endereço, método `GetListedCashDividends` | JCP e restituição de capital em dinheiro (`REST CAP DIN`) desde 2016, com data de aprovação e data com. A restituição vira proposta de redução de capital na tela Eventos (só o valor, na data ex). | `proventos_historico_b3.csv` |
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
