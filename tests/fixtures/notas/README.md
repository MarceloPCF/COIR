# Fixtures de notas de corretagem (texto anonimizado)

Dez notas reais em **texto extraído** (pypdfium2, `get_textpage().get_text_range()`), uma nota por arquivo, para
testar o extrator SINACOR (F2). Quando a nota tem mais de uma folha (página), o `.txt` traz o texto de todas
as folhas, separadas por uma linha com `` (form feed). Cada `<id>.txt` tem um `<id>.esperado.json` com o
resultado esperado (operações, resumo financeiro, IRRF, líquido). Valores monetários são strings decimais com
ponto (`"18660.00"`); datas em ISO.

| id | Corretora / layout | Segmento | O que cobre |
|---|---|---|---|
| `agora_2008_04_bovespa` | Ágora, layout antigo (2008) | Bovespa | 1 compra à vista (obs. `H`), corretagem + ISS, sem IRRF, D/C antes do valor |
| `xp_2026_08_bovespa` | XP, layout novo (2026) | Bovespa | 13 vendas à vista (obs. `@` e `@#`), 2 ativos, IRRF de R$ 4,20 incluído em "Total Custos / Despesas", rótulos e valores do resumo separados |
| `btg_2023_01_bovespa_opcoes` | BTG/Necton (2023) | Bovespa | 6 opções (CALL e PUT, compra e venda, PN/ON/UNT), IRRF de R$ 0,01 não debitado (<= R$ 1) |
| `btg_2023_01_bovespa_exercicio` | BTG/Necton (2023) | Bovespa | exercício de opção de venda (`EXERC OPC VENDA`, código com sufixo `E`, sem classe), D/C = D |
| `xp_2019_02_bmf_daytrade` | XP BM&F (2019) | BM&F | 8 negócios de mini-índice (`WING19`), `DAYTRADE`, obs. `@`, IRRF Day Trade (proj.) |
| `xp_2016_11_bovespa_antigo_multipagina` | XP, layout antigo (2016-2023) | Bovespa | nota de 2 folhas (folha 1 com 18 negócios e "CONTINUA...", folha 2 com 1 negócio e o resumo), 19 operações em 2 ativos (PN), day trade (obs. `D` e `D#`) e comuns (`#`), IRRF de R$ 2,23 sobre vendas, incluído em "Total Custos / Despesas", linha "IRRF Day Trade: Base ... Projeção ...", palavras coladas |
| `rico_2015_11_bovespa` | Rico (2015-2016, template XP antigo) | Bovespa | 6 vendas à vista (obs. `H` e `H#`), IRRF de R$ 0,53 não debitado (<= R$ 1), corretagem (Taxa Operacional) + Impostos, palavras coladas (`baseR$10.660,00`), conta `386-7` |
| `xp_2024_01_bmf_novo` | XP BM&F (2024+) | BM&F | 6 negócios de milho (`CCM H24`): 1 `AJUPOS` (ajuste de posição) + 5 `DAY TRADE` (obs. `@`), valor dos negócios devedor, Ajuste de posição x Ajuste day trade separados, "NOTA DE NEGOCIAÇÃO Nr. nota" na mesma linha |
| `xp_2024_01_bmf_novo_ajupos_negativo` | XP BM&F (2024+) | BM&F | `AJUPOS` com **quantidade negativa** (`-4`, venda de posição) + 2 linhas `NORMAL` (obs. `@`), IRRF de R$ 0,08 D que não entra no líquido, "Total liquido (#)" 0,00 |
| `btg_2023_11_bmf` | BTG/Necton BM&F (2023) | BM&F | 2 negócios `DAY TRADE` de mini-índice (`WINZ23`), valor antes do rótulo e D/C soltos no texto, IRRF Day Trade (proj.) de R$ 0,57 (1% de 57,50 truncado) |

`btg_2023_01_bovespa_exercicio` e `xp_2024_01_bmf_novo_ajupos_negativo` são extras além das notas pedidas
nas tarefas (a segunda porque nenhuma nota XP BM&F de 2024+ tem, ao mesmo tempo, `AJUPOS` negativo e day trade).

## Origem e anonimização

As notas vêm de PDFs reais do mantenedor (guardados fora do repositório, em `COIR_DADOS`; os PDFs
**nunca** são copiados para cá). Em cada texto foram substituídos:

- nome do cliente por `FULANO DE TAL`; endereço, bairro, cidade, CEP e telefone por valores genéricos
  (`Rua Exemplo, 0`, `CIDADE - UF`, `00000-000`, `(00) 00000-0000`);
- CPF por `000.000.000-00`; código do cliente/conta por `000000` (formato de conta SINACOR por
  `85-0 000000000`, ou `3-5 000000` na XP); código do assessor por `0000`;
- número da nota por número fictício (`1000001` a `1000010`), em todas as ocorrências (inclusive em todas as
  folhas de uma nota multipágina).

O CPF mantém a pontuação do original: as notas XP/Rico/BTG de 2015-2024 trazem o CPF com ponto (e não hífen)
antes dos dois últimos dígitos, que virou `000.000.000.00`. Nas BM&F da XP o "Codigo do Cliente" (sozinho numa linha)
virou `000000`; na BTG BM&F o CEP do cliente virou `00000000` e o código no fim do endereço, `000000`. O código
do assessor (`Código cliente Assessor`) virou `0000`. A Rico traz ainda uma linha extra de 3 grupos de números
(antes do CPF), trocada por `000 0000 00000`.

Foram mantidos ativos, datas, quantidades, preços, valores, custos, IRRF e líquido (para os totais continuarem
conciliando). Dados da corretora (razão social, CNPJ, endereço, telefones de SAC) são públicos e ficam como estão.
Na `xp_2019_02_bmf_daytrade` o número da nota aparece no texto como `1.000.005` (formato com ponto de milhar da
XP) e no JSON como `"1000005"` (`numero_nota_no_texto` guarda o formato original). A data do pregão desta nota é
01/02/2019; `13/02/2019` nas linhas é o vencimento do contrato `WING19`.

Antes de versionar, rode `uv run python tools/verificar_dados_pessoais.py tests/fixtures/notas/*`.

## Conferência da conciliação (feita na criação)

Para cada fixture foi conferido que:

- a soma das operações é igual ao "Valor das operações" (Bovespa) ou, no BM&F, a soma com sinal (C positivo,
  D negativo) é igual ao "Valor dos negócios" e ao "Ajuste day trade";
- compras e vendas batem com o "Resumo dos Negócios"; quantidade x preço = valor em cada linha Bovespa;
- Bovespa: Total CBLC = valor líquido das operações -/+ taxa de liquidação e de registro; Total Bovespa/Soma =
  emolumentos + termo/opções + ANA (+ transferência de ativos); Total corretagem/despesas (ou Total
  Custos/Despesas) = soma de seus itens; líquido da nota = Total CBLC -/+ Total Bovespa -/+ corretagem/despesas;
- BM&F: Taxa operacional por linha (8 x 0,20) = campo do resumo; custos somam 2,64; Total líquido (#) = 44,00 - 2,64;
  Total Conta Normal = líquido (#) - IRRF Day Trade; Total líquido da nota = líquido (#) - IRRF Day Trade - impostos.

Conferência das cinco fixtures acrescentadas depois (feita com script descartável fora do repositório, com
`Decimal`; todas fecharam sem diferença):

- `xp_2016_11_bovespa_antigo_multipagina`: as 19 operações (18 na folha 1, 1 na folha 2) somam 123.910,00 = "Valor
  das operações" = vendas 83.610,00 + compras 40.300,00; quantidade x preço = valor em todas as linhas; valor
  líquido 43.310,00 C = vendas - compras; Total CBLC 43.281,88 = 43.310,00 - 28,12 (liquidação) - 0,00 (registro);
  Total Bovespa/Soma 6,19 = emolumentos; Total Custos/Despesas 2,23 = taxa operacional 0,00 + impostos 0,00 + IRRF
  2,23; líquido 43.273,46 = 43.281,88 - 6,19 - 2,23.
- `rico_2015_11_bovespa`: as 6 vendas somam 10.660,00 = "Valor das operações"; Total CBLC 10.657,07 = 10.660,00 -
  2,93; Total Bovespa/Soma 0,53 = emolumentos; Total Custos/Despesas 10,31 = taxa operacional 9,80 + impostos 0,51
  (o IRRF de 0,53 não entra); líquido 10.646,23 = 10.657,07 - 0,53 - 10,31. IRRF 0,53 = 0,005% de 10.660,00
  truncado.
- `xp_2024_01_bmf_novo`: soma com sinal das 6 linhas = -652,50 = "Valor dos negócios" (D); `AJUPOS` +756,00 = "Ajuste
  de posição" (C); `DAY TRADE` 202,50 + 1.206,00 = -1.408,50 = "Ajuste day trade" (D); taxa operacional por linha
  (4 x 4,00 + 16,00) = 32,00; custos 32,00 + 1,84 + 1,04 = 34,88; Total Conta Normal = -652,50 - 34,88 = 687,38 D;
  Total líquido da nota = 687,38 + impostos 3,38 = 690,76 D.
- `xp_2024_01_bmf_novo_ajupos_negativo`: `AJUPOS` 4.590,00 C - 2 x `NORMAL` 738,00 D = 3.114,00 C = "Valor dos
  negócios" = "Ajuste de posição" (as linhas `NORMAL` entram no ajuste de posição; ajuste day trade 0,00);
  taxa operacional 2 x 4,00 = 8,00; custos 8,00 + 0,94 + 0,50 = 9,44; Total Conta Normal = 3.114,00 - 9,44 =
  3.104,56 C; Total líquido da nota = 3.104,56 - impostos 0,84 = 3.103,72 C (o IRRF de 0,08 D não é deduzido).
- `btg_2023_11_bmf`: 634,00 C - 574,00 D = 60,00 C = "Valor dos negócios" = "Ajuste day trade"; custos 0,00 +
  1,60 + 0,90 = 2,50 = "Total das despesas"; Total Conta Normal = Total liquido (#) = 60,00 - 2,50 = 57,50 C; IRRF
  Day Trade (proj.) = 1% de 57,50 = 0,575 truncado em 0,57; Total líquido da nota = 57,50 - 0,57 = 56,93 C. Na BTG o
  valor de cada campo foi associado ao rótulo pela posição na página (pdfplumber), não pela ordem do texto.

Todas as conferências fecharam sem diferença.

## Particularidades do layout (úteis para o extrator)

- Nas notas XP novas e BTG o texto extraído separa os rótulos do resumo dos seus valores (ficam em blocos
  distintos). Nas BTG, os três valores sem rótulo antes de "I.R.R.F. s/ operações" são, pela ordem, ISS, IRRF e
  "Outras Bovespa" (associação inferida: o IRRF 0,01 é 0,005% de 391,00 truncado). No JSON, ISS e "Outras
  Bovespa" constam como `0.00`.
- XP 2026: o IRRF acima de R$ 1,00 vem com `D` e entra em "Total Custos / Despesas"; o campo
  `incluido_em_total_custos_despesas` do JSON registra isso.
- BM&F: `valor` das linhas é o ajuste de day trade com D/C; "Total líquido da nota" desconta também o IRRF Day
  Trade (proj.) e os impostos.
- `conta` e `numero_nota` do JSON são os valores fictícios; `conta: null` quando a nota não traz conta legível.
- XP antigo (2016-2023) com mais de uma folha: a folha 1 traz os negócios e só os rótulos do resumo, com
  "CONTINUA..." junto a "Total Bovespa / Soma"; os valores do resumo e o líquido ficam só na última folha. O número da
  nota é o mesmo nas folhas, e "Folha" muda (1, 2). `folhas` e `operacoes_por_folha` do JSON registram isso. O
  texto antigo é colado ("Resumo dosNegócios", "Vendasà vista", "baseR$44.610,00 2,23 D").
- Rico 2015-2016: mesmo template do XP antigo (CNPJ da Rico 02.332.886/0016-82, conta `386-7`). Com IRRF <= R$ 1,00 o
  valor vem sem `D` e não entra em "Total Custos / Despesas"; acima disso vem com `D` (ver notas de 2016 da Rico).
  Os endereços ficam em linhas separadas (bairro, telefone, uma linha extra com 3 grupos de números antes do CPF).
- XP BM&F 2024+: o tipo de negócio é `DAY TRADE` (com espaço; em 2019 era `DAYTRADE`), `AJUPOS` ou `NORMAL`; o
  mercado tem espaço (`CCM H24`), o vencimento pode vir colado a `@`, e a quantidade do `AJUPOS` pode ser negativa
  (posição vendida). O IRRF da linha de custos pode vir como `0,08|D`; nesse caso (nota de ajuste, sem day trade)
  ele não entra em "Total Conta Normal" nem em "Total líquido da nota". Quando o resultado é devedor ou a nota
  só tem ajuste, "Total liquido (#)" vem `0,00`; "Total líquido da nota" = Total Conta Normal +/- impostos.
- BTG BM&F: no texto do pypdfium2 cada valor vem antes do seu rótulo (ordem invertida em relação à tabela) e os D/C
  ficam em linhas soltas (` | C`, ` | D`) no fim; usar pdfplumber (coordenadas) ou o par rótulo/valor com cuidado.
  "Total Conta Normal" e "Total liquido (#)" são iguais ao resultado menos custos; só "Total líquido da nota"
  desconta o IRRF Day Trade (proj.). O número da nota vem logo abaixo de "NOTA DE CORRETAGEM" (sem pontuação).
- `conta` é `null` nas BM&F: o código da corretora (`3-5`, `85-0`) e o código do cliente vêm em linhas separadas.
