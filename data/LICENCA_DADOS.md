# Origem e licença dos dados desta pasta

O código do COIR é AGPL-3.0-or-later. As tabelas desta pasta são **dados**, e parte delas vem de
bases de terceiros, cada uma com as suas condições. Este arquivo acompanha as tabelas no
repositório, no pacote do programa e no pacote de referências.

## Comissão de Valores Mobiliários (CVM)

Contém informações do Portal de Dados Abertos da CVM (https://dados.cvm.gov.br), acessadas pelo
portal e disponibilizadas sob a Open Database License (ODbL) 1.0
(https://opendatacommons.org/licenses/odbl/1-0/): Formulário de Referência (FRE), Informações
Periódicas e Eventuais (IPE), Formulário Cadastral (FCA), cadastro de companhias abertas e
informes mensais de FII e de Fiagro.

- Onde entram: `eventos_cvm.csv` (coluna `origem`: `cvm_fre` ou `cvm_ipe`),
  `valores_mobiliarios_cvm.csv`, `fundos_cvm.csv`, as linhas de
  `trocas_ticker.csv` e de `saidas.csv` cuja `fonte` cita a CVM, e as tabelas de
  `tools/referencias/` (`eventos_cvm_ipe.csv`, `emissores_antigos.csv`, `sem_evento.csv`).
- Essas linhas, como base de dados derivada, ficam disponíveis sob a mesma ODbL 1.0.
- A CVM fornece os dados como estão e não responde pelo uso feito deles; a seleção, a
  transcrição e os erros desta pasta são do COIR.

## Banco Central do Brasil

`selic_mensal.csv`: série 4390 do Sistema Gerenciador de Séries Temporais (SGS), Selic acumulada
no mês, do Portal de Dados Abertos do Banco Central (https://dadosabertos.bcb.gov.br), sob a Open
Database License (ODbL) 1.0. Fonte: Banco Central do Brasil, Departamento de Operações do Mercado
Aberto.

## B3 S.A. – Brasil, Bolsa, Balcão

Fonte das tabelas com `_b3` no nome (cadastro de instrumentos, nomes de pregão, tipos de ativo,
emissores, eventos e proventos):
páginas e arquivos públicos da B3 (https://www.b3.com.br), consultados sem login. As tabelas não
trazem cotações.

O COIR não tem vínculo com a B3, não é endossado por ela e não fala em nome dela. Os dados
pertencem à B3 e estão sujeitos aos termos de uso dela; não os use para fins comerciais nem os
redistribua sem consultar esses termos.

## Tabelas do próprio COIR

`ativos.csv`, `opcoes_raiz.csv`, `tipos_ativo.csv`, `parametros_fiscais.csv`, `feriados.csv`,
`irpf/` e as demais tabelas pequenas são curadoria do projeto, a partir de leis, normas, notas de
corretagem e avisos das companhias, e seguem a licença do COIR.
