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
| B3, série histórica de cotações (COTAHIST) | `https://bvmf.bmfbovespa.com.br/InstDados/SerHist/COTAHIST_AAAA.ZIP` | Por pregão e ISIN, no mercado à vista de qualquer faixa (BDI 02 e também 05 a 08 e 58, das companhias em recuperação judicial ou extrajudicial, concordata ou sanção): preço, número de distribuição (muda no pregão seguinte à data com de todo evento) e a marca de evento na especificação, do dia ex em diante (`EB` bonificação ou desdobramento, `EG` grupamento, `ED`/`EJ`/`ER`/`ES`/`EC` dividendo, juros, restituição ou rendimento, subscrição, cisão). Confirma a data com dos eventos do FRE e dos avisos e é a varredura do mercado inteiro. | idem (não é publicado: só confere) |
| CVM, documentos IPE (dados abertos) | `https://dados.cvm.gov.br/dados/CIA_ABERTA/DOC/IPE/DADOS/ipe_cia_aberta_AAAA.zip` e o documento em `https://www.rad.cvm.gov.br/ENET/frmDownloadDocumento.aspx` | Fatos relevantes e avisos aos acionistas: proporção, data-base, data ex e custo atribuído da bonificação. Usados para 2023 em diante, quando o FRE parou e a B3 já tinha substituído o evento, para o que o FRE não resolve (fator ruidoso, papel sem negócio para conferir, quantidade "antes" repetida de outro ato: BIDI 2021), para o que a varredura acha e para o custo atribuído das bonificações. | `tools/referencias/eventos_cvm_ipe.csv` (transcrito, com o protocolo); `origem` = `cvm_ipe` |
| CVM, Formulário Cadastral (FCA, dados abertos), tabelas `geral` e `valor_mobiliario` | `https://dados.cvm.gov.br/dados/CIA_ABERTA/DOC/FCA/DADOS/fca_cia_aberta_AAAA.zip` (2010 em diante) | Por CNPJ: nome empresarial com o nome anterior e a data da mudança; códigos de negociação (só de 2018 em diante). Confirma que o código antigo e o novo de uma troca são a mesma companhia, e liga ao FRE a companhia que saiu da bolsa. | `trocas_ticker.csv` (coluna `fonte`; curada) e `python -m tools.referencias historico` |
| CVM, cadastro de companhias abertas | `https://dados.cvm.gov.br/dados/CIA_ABERTA/CAD/DADOS/cad_cia_aberta.csv` | CNPJ, nome e data do cancelamento do registro. Liga ao FRE e ao IPE a companhia que saiu da bolsa antes de 2018 (o FCA não tem o código dela). | `tools/referencias/emissores_antigos.csv` (curado, com a fonte) |
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
  data com e a cotação da unit confirma; o custo atribuído da unit é o da ação vezes a composição
  (ALUP11 e SULA11: 1 ON + 2 PN; KLBN11: 1 ON + 4 PN), ou o que o aviso informa (ENGI11);
- a marca da B3 decide a data com quando o preço não decide (bonificação pequena, duas mudanças
  de distribuição na janela): conta a marca que aparece no pregão (a B3 a mantém alguns dias, e um
  dividendo nesse meio-tempo não é evento novo: BBDC 2010). A marca dá a data, não o fator: se o
  preço consegue conferir o fator, tem de confirmar (o FRE dizia 5722% para CTSA 2020 e 811% para
  PATI 2014; a B3 e o aviso, 183,21% e 12,32%). Fator que arredonda para 1 é ruído, não evento;
- papel sem negócio entre o último pregão antes do evento e o primeiro depois dele: qualquer data
  com do intervalo dá a mesma posição, porque ninguém negociou. Vale a véspera do primeiro pregão
  com a marca (ou com a mudança coerente com o fator), e a referência guarda o intervalo ("data
  com: sem negócio de A a B"); a data de outra fonte dentro do intervalo é o mesmo evento (soma a
  origem, sem divergência);
- o custo atribuído da bonificação vem do aviso aos acionistas (Lei 9.249/1995, art. 10, §1º; IN
  RFB 1.585/2015, art. 58, §1º). Quando há re-ratificação ou aviso final, vale o último (RAPT 2010:
  R$ 1,99 corrigido para R$ 3,98; VITT 2026). Sem o valor por ação no documento, é a parcela
  capitalizada dividida pelas ações bonificadas, calculada do próprio aviso (SULA 2013: R$
  1.000.000.000 / 163.666.122 ações). Bonificação sem custo no pacote não é lançada sozinha: a
  tela Eventos pede o custo.

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
- antes de 2018, o CNPJ do FRE chega ao código pela tabela curada
  `tools/referencias/emissores_antigos.csv` (CNPJ, código, razão social e a fonte: cadastro de
  companhias abertas da CVM e o nome no COTAHIST). Ex.: a Ambev antiga (AMBV, desdobramento 5:1
  em 2010), Aços Villares, Confab, Fosfertil, AES Tietê, Oi/Brasil Telecom, ALL;
- de 2023 em diante (sem FRE), o evento vem do aviso aos acionistas transcrito em
  `eventos_cvm_ipe.csv`, conferido pela cotação.

Varredura do mercado inteiro e casos conferidos (`historico.saltos`, `sem_evento.csv`):

- toda mudança de distribuição com a marca de bonificação, desdobramento ou grupamento da B3 (B ou
  G aparecendo no pregão), ou com o preço saltando perto de uma fração simples (2, 1/10, 3/2...,
  com pelo menos 40% de variação), tem de ter evento no pacote para o emissor (ou para quem trocou
  de código com ele) com data com entre 7 pregões antes do último negócio sem a mudança e o
  primeiro com ela. Isso independe do FRE e da consulta da B3: é o que pegou o desdobramento 1:3
  do Banco Inter em 2021 (o FRE repetiu a quantidade "antes" de 2019, e o fator não fechou);
- o que não é evento em ações vai para `tools/referencias/sem_evento.csv` (colunas `origem`,
  `isin`, `data`, `motivo`, `fonte`): `origem` "cotacao" para o salto (data com do salto) e "fre"
  para o ato do FRE (data de aprovação). Motivos: a marca da B3 é de dividendo, juros,
  restituição, subscrição ou cisão; o aviso é de dividendo, redução de capital ou aumento por
  subscrição; o ato do FRE é anterior à listagem, posterior à saída, já está no pacote por outra
  fonte (com outro tipo ou data), ou a B3 não marcou nada nem o preço saltou; data impossível do
  FRE (VIVR "2077-03-08" é 08/03/2007, antes da listagem). Quando a única evidência é a marca da
  B3 e não há documento acessível (companhia com registro cancelado, fora do IPE e do FRE), a
  linha diz "sem fonte oficial acessível" e o evento fica fora até haver o aviso. Linha sem motivo,
  sem fonte ou com outra origem faz a geração falhar;
- o resumo do histórico lista o que sobra (atos do FRE sem evento, saltos sem evento,
  divergências); o teste `tests/golden/test_cobertura_mercado_real.py` (com os arquivos públicos
  na pasta de `COIR_CACHE_B3`) exige que tudo isso saia vazio.

Restituição de capital em dinheiro: a aba de proventos da B3 (`REST CAP DIN`) traz só as
companhias listadas hoje; a de companhia que já saiu da bolsa (CR2 2015 e 2020, Tarpon 2019, GTD
2013...) fica registrada em `sem_evento.csv` como redução de capital, sem proposta no app.
