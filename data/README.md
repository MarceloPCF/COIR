# data/

Tabelas de referência **públicas** e versionadas (ativos, aliases, corretoras, contratos futuros, feriados,
parâmetros fiscais com vigência). Nunca colocar aqui dados pessoais: esses ficam na pasta `COIR_DADOS`.

| Arquivo | Uso | Colunas |
|---|---|---|
| `tolerancias.toml` | tolerâncias da comparação com o gabarito | `operacao`, `pm`, `mes` |
| `ativos.csv` | nome de pregão → ticker (comparado sem espaços e sem acentos; sufixos de governança removidos) | `nome_pregao`, `ticker` |
| `sufixos_especificacao.csv` | sufixos removidos do fim do nome de pregão (NM, N1, EDJ, ...) | `sufixo` |
| `opcoes_raiz.csv` | raiz de 4 letras → ativo-objeto (último recurso no exercício, depois da classe da linha e da abertura da série) | `raiz`, `ticker` |
| `classes_acao.csv` | classe → dígito do ticker (ON → 3, PN → 4, ..., UNT → 11) | `classe`, `digito` |
| `mercados_bovespa.csv` | tipo de mercado impresso → mercado normalizado | `tipo_mercado`, `mercado` |
| `contratos_futuros.csv` | raiz do contrato → mercado, multiplicador, moeda da cotação (BRL/USD), liquidação (FINANCEIRA/FISICA) e linha da apuração (indice, dolar, commodity), com vigência | `raiz`, `mercado`, `multiplicador`, `moeda`, `liquidacao`, `grupo`, `inicio`, `fim` |
| `selic_mensal.csv` | Selic acumulada no mês (% a.m.), para os juros de mora da DARF; só meses completos (a API devolve o mês corrente parcial), contíguos, 2 casas | `mes` (AAAA-MM), `taxa` |
| `trocas_ticker.csv` | ticker impresso que valia como outro até uma data | `ticker_impresso`, `ticker`, `ate` |

`ativos.csv` e `opcoes_raiz.csv` vieram de `legacy/dados/` (v2.1.1). Na migração, nomes repetidos com dois
tickers ficaram com o primeiro, como a v2 fazia: `B TECH EQI ON` (JBDU3; descartado BLUT3), `B TECH EQI PN`
(JBDU4; descartado BLUT4) e `WILSON SONS ON` (WSON33; descartado PORT3). Corrigido: `WEG ON ED` → WEGE3
(a v2 tinha WEGE4).

Ativo sem cadastro não é chutado: a extração falha com sugestões de nomes parecidos para incluir em `ativos.csv`.

`selic_mensal.csv`: série 4390 do Banco Central (SGS), a mesma da consulta Selic do Sicalc. Para
atualizar, baixar `https://api.bcb.gov.br/dados/serie/bcdata.sgs.4390/dados?formato=csv` e manter só os
meses já encerrados; o carregamento recusa mês repetido, buraco na sequência e taxa com mais de 2 casas.
