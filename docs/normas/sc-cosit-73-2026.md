# Solução de Consulta Cosit nº 73, de 24 de abril de 2026

| Campo | Valor |
|---|---|
| Ementa (resumo) | ETF negociado em bolsa cujo índice de renda fixa seja de debêntures incentivadas (ou de fundos delas) é Fundo de Índice de Renda Fixa (Lei 13.043 art. 2º), tributado pela IN 1.585 arts. 28-30, e não fundo de debêntures incentivadas (Lei 12.431 art. 3º) |
| Fonte oficial | https://normas.receita.fazenda.gov.br/sijut2consulta/link.action?idAto=150826 |
| Texto lido | https://normasinternet2.receita.fazenda.gov.br/api/consulta-externa/ato/150826/visao/multivigente e anexo 84284 (íntegra) |
| Lido em | 2026-11-10 |
| Leitura | íntegra |
| Confiança | alta |
| Temas | ETF; FI-Infra/FIP-IE |

## Vigência e alterações

| Dispositivo | Redação | Vigente de | até | Alterado por |
|---|---|---|---|---|
| SC inteira | original | 30/04/2026 (DOU) | — | não reformada (SIJUT: vigente). Consulta prévia, antes de o fundo existir; interpreta a Lei 13.043 art. 2º, vigente desde 2015 |

## Dispositivos relevantes para o COIR

- **§§12-13** — FIRF = fundo com cotas em bolsa, que busca refletir índice de renda fixa, com no mínimo 75% em ativos do índice (Lei 13.043 art. 2º); é o ETF de renda fixa.
- **§§14-15** — FICDI = fundo com no mínimo 85% em debêntures incentivadas/de infraestrutura, ou fundo de fundos com 95% (Lei 12.431 art. 3º, redação da Lei 14.801/2024).
- **§§16-17 e conclusão** — prevalece a forma de constituição: o fundo constituído como fundo de índice é FIRF para todos os efeitos, mesmo que o índice seja de debêntures incentivadas.

## Onde o COIR usa

- Não citada. `docs/regras-fiscais.md` R20: tipo `FUNDO_INFRA` só pela lista de fundos da B3 (FI-INFRA, FIP IE), nunca pelo nome; ETF de renda fixa é `ETF_RF` (`coir/engine/movimentos.py`, `SEM_IRRF_DA_VENDA`). **Confere**: um ETF de debêntures incentivadas fica como ETF de renda fixa, sem a alíquota zero do FI-Infra.

## Pontos em aberto

- Conferir se algum ETF de debêntures aparece na lista da B3 com o tipo FI-INFRA. Se aparecer, a lista da B3 levaria ao tipo errado e precisaria de exceção.
