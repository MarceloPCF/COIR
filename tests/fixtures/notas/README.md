# Fixtures de notas de corretagem (texto anonimizado)

Cinco notas reais em **texto extraído** (pypdfium2, `get_textpage().get_text_range()`), uma nota (uma página)
por arquivo, para testar o extrator SINACOR (F2). Cada `<id>.txt` tem um `<id>.esperado.json` com o resultado
esperado (operações, resumo financeiro, IRRF, líquido). Valores monetários são strings decimais com ponto
(`"18660.00"`); datas em ISO.

| id | Corretora / layout | Segmento | O que cobre |
|---|---|---|---|
| `agora_2008_04_bovespa` | Ágora, layout antigo (2008) | Bovespa | 1 compra à vista (obs. `H`), corretagem + ISS, sem IRRF, D/C antes do valor |
| `xp_2026_08_bovespa` | XP, layout novo (2026) | Bovespa | 13 vendas à vista (obs. `@` e `@#`), 2 ativos, IRRF de R$ 4,20 incluído em "Total Custos / Despesas", rótulos e valores do resumo separados |
| `btg_2023_01_bovespa_opcoes` | BTG/Necton (2023) | Bovespa | 6 opções (CALL e PUT, compra e venda, PN/ON/UNT), IRRF de R$ 0,01 não debitado (<= R$ 1) |
| `btg_2023_01_bovespa_exercicio` | BTG/Necton (2023) | Bovespa | exercício de opção de venda (`EXERC OPC VENDA`, código com sufixo `E`, sem classe), D/C = D |
| `xp_2019_02_bmf_daytrade` | XP BM&F (2019) | BM&F | 8 negócios de mini-índice (`WING19`), `DAYTRADE`, obs. `@`, IRRF Day Trade (proj.) |

`btg_2023_01_bovespa_exercicio` é um extra além das quatro notas pedidas na tarefa.

## Origem e anonimização

As notas vêm de PDFs reais do mantenedor (guardados fora do repositório, em `COIR_DADOS`; os PDFs
**nunca** são copiados para cá). Em cada texto foram substituídos:

- nome do cliente por `FULANO DE TAL`; endereço, bairro, cidade, CEP e telefone por valores genéricos
  (`Rua Exemplo, 0`, `CIDADE - UF`, `00000-000`, `(00) 00000-0000`);
- CPF por `000.000.000-00`; código do cliente/conta por `000000` (formato de conta SINACOR por
  `85-0 000000000`, ou `3-5 000000` na XP); código do assessor por `0000`;
- número da nota por número fictício (`1000001` a `1000005`), em todas as ocorrências.

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
