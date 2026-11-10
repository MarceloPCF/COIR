# Solução de Consulta Cosit nº 37, de 31 de março de 2020

| Campo | Valor |
|---|---|
| Ementa (resumo) | Resgate em dinheiro de ações preferenciais resgatáveis (operação CETIP/BM&FBOVESPA via holding) é alienação sujeita a ganho de capital: valor recebido menos custo, sem atualização do custo posterior a 1996 |
| Fonte oficial | https://normas.receita.fazenda.gov.br/sijut2consulta/link.action?idAto=108303 |
| Texto lido | https://normasinternet2.receita.fazenda.gov.br/api/consulta-externa/ato/108303/visao/multivigente e anexo 56086 (íntegra) |
| Lido em | 2026-11-10 |
| Leitura | íntegra |
| Confiança | alta (conteúdo); média (aplicação a outros casos de resgate) |
| Temas | eventos societários; ações; resgate |

## Vigência e alterações

| Dispositivo | Redação | Vigente de | até | Alterado por |
|---|---|---|---|---|
| SC inteira | original | 02/04/2020 (DOU) | — | não reformada (SIJUT: vigente em 10/11/2026) |

## Dispositivos relevantes para o COIR

- **Relatório §§2-6** — fatos: cada ação CETIP virou 1 ON + 3 PN resgatáveis da holding; as ON foram trocadas por ações da BM&FBOVESPA e as PN resgatadas em dinheiro (com correção pelo CDI, sobre a qual houve IRRF como rendimento). O consulente não questionou a troca das ON.
- **§15** — ganho de capital de pessoa física fora de bolsa segue a IN SRF 84/2001 (art. 16 §§3º-5º: custo médio ponderado por espécie).
- **§18 e conclusão** — o resgate é alienação "a qualquer título" (IN 84 art. 3º I); ganho = valor de alienação menos custo; custo adquirido desde 1996 não se atualiza (IN 84 art. 8º).
- A SC **não** diz como repartir o custo entre ON e PN, nem trata da parcela de correção pelo CDI (só registra que houve retenção como rendimento).

## Onde o COIR usa

- `docs/regras-fiscais.md` R8 (PETZ3 -> AUAU3: parte resgatada sai por `REDUCAO_CAPITAL`, resgate é alienação fora de bolsa, no GCAP): **confere** com a SC; o critério de rateio do custo é inferência do projeto (o R8 já diz "confiança média").
- `data/trocas_ticker.csv` (PETZ3), `data/saidas_b3.csv` (CTIP3 -> BVMF3), `coir/servicos/eventos.py`, `tests/fiscal_cases/test_petz_auau.py`, `tests/unit/test_servicos_eventos.py`: citação **confere** (CTIP3 é exatamente o caso da SC).

## Pontos em aberto

- A correção pelo CDI paga junto do resgate (caso CETIP) é rendimento com IRRF, fora do ganho de capital (relatório §6); o COIR não a trata. Registrar se aparecer em outro evento.
