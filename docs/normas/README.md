# Base normativa do COIR

Fichas das normas que o COIR aplica, uma por norma: vigência de cada redação desde 2016, dispositivos
relevantes resumidos com as próprias palavras, onde o COIR usa a norma e pontos em aberto. Cada ficha
diz de onde e quando a íntegra foi lida. O documento orienta; não substitui a norma nem um profissional
habilitado.

- Montada em 10/11/2026 pelo especialista fiscal, lendo a íntegra nas fontes oficiais: Câmara (Legin) e
  Senado para leis (o Planalto recusa conexões de linha de comando; o link dele fica na ficha), SIJUT
  (API da Receita, visão multivigente, com a íntegra das Soluções de Consulta), gov.br e BCB.
- A íntegra não vai para o Git: `uv run python -m tools.normas_acervo` baixa as normas de
  [`acervo.csv`](acervo.csv) para `.cache/normas/`.
- Modelo de ficha: [`_modelo.md`](_modelo.md). Soluções de Consulta lidas e descartadas:
  [`sc-cosit-lidas-sem-relevancia.md`](sc-cosit-lidas-sem-relevancia.md).
- Lacunas: os Perguntas e Respostas IRPF 2023 e 2025 não estão publicados na página da Receita
  (as URLs devolvem 404 em 10/11/2026); a correspondência de perguntas pula esses anos.
- Revisão do projeto contra a base: [`revisao-2026-11-10.md`](revisao-2026-11-10.md).
- Mantida na rotina semanal do especialista (`.claude/agents/especialista-fiscal-b3.md`).
- [`docs/fontes-fiscais.md`](../fontes-fiscais.md) lista onde cada norma é citada no código; as fichas
  dizem o que a norma diz.

## Índice por tema

Uma ficha aparece em todos os temas que trata. Confiança e tipo de leitura vêm do cabeçalho da ficha.

### Ações (ganho líquido, custo médio, custos)

- [Decreto nº 9.580, de 22 de novembro de 2018 (Regulamento do Imposto sobre a Renda — RIR/2018)](decreto-9580-2018-rir.md) — leitura parcial; confiança alta
- [Instrução Normativa RFB nº 1.022, de 5 de abril de 2010 (revogada)](in-rfb-1022-2010.md) — leitura parcial; confiança alta
- [Instrução Normativa RFB nº 1.500, de 29 de outubro de 2014](in-rfb-1500-2014.md) — leitura íntegra; confiança alta
- [Instrução Normativa RFB nº 1.585, de 31 de agosto de 2015](in-rfb-1585-2015.md) — leitura íntegra; confiança alta
- [Instrução Normativa RFB nº 1.916, de 18 de dezembro de 2019](in-rfb-1916-2019.md) — leitura íntegra; confiança alta
- [Instrução Normativa RFB nº 2.164, de 25 de outubro de 2023](in-rfb-2164-2023.md) — leitura íntegra; confiança alta
- [Lei nº 11.033, de 21 de dezembro de 2004](lei-11033-2004.md) — leitura parcial; confiança alta
- [Lei nº 15.265, de 21 de novembro de 2025](lei-15265-2025.md) — leitura íntegra; confiança alta
- [Lei nº 7.713, de 22 de dezembro de 1988](lei-7713-1988.md) — leitura parcial; confiança alta
- [Lei nº 8.981, de 20 de janeiro de 1995](lei-8981-1995.md) — leitura parcial; confiança alta
- [Manual do Meu Imposto de Renda (MIR) — páginas da RFB](manual-mir.md) — leitura íntegra; confiança média
- [Medida Provisória nº 1.303, de 11 de junho de 2025 (perdeu a eficácia; não vale)](mp-1303-2025.md) — leitura íntegra; confiança alta
- [Perguntas e Respostas IRPF 2016 a 2020 — renda variável (por edição)](pr-irpf-2016-2020.md) — leitura íntegra; confiança alta
- [Perguntas e Respostas IRPF 2021 a 2026 — renda variável (por edição)](pr-irpf-2021-2026.md) — leitura íntegra; confiança alta
- [Perguntas e Respostas IRPF (RFB), edições 2016 a 2026 — renda variável](pr-irpf.md) — leitura íntegra; confiança alta
- [Manual de utilização do ReVar (RFB)](revar.md) — leitura íntegra; confiança média
- [Solução de Consulta Cosit nº 10, de 3 de fevereiro de 2016](sc-cosit-10-2016.md) — leitura íntegra; confiança alta
- [Solução de Consulta Cosit nº 127, de 8 de maio de 2024](sc-cosit-127-2024.md) — leitura íntegra; confiança alta
- [Solução de Consulta Cosit nº 130, de 26 de junho de 2023](sc-cosit-130-2023.md) — leitura íntegra; confiança alta
- [Solução de Consulta Cosit nº 148, de 15 de agosto de 2025](sc-cosit-148-2025.md) — leitura parcial; confiança alta
- [Solução de Consulta Cosit nº 153, de 18 de novembro de 2016](sc-cosit-153-2016.md) — leitura íntegra; confiança alta
- [Solução de Consulta Cosit nº 224, de 14 de agosto de 2014](sc-cosit-224-2014.md) — leitura íntegra; confiança alta
- [Solução de Consulta Cosit nº 227, de 4 de dezembro de 2018](sc-cosit-227-2018.md) — leitura íntegra; confiança alta
- [Solução de Consulta Cosit nº 37, de 31 de março de 2020](sc-cosit-37-2020.md) — leitura íntegra; confiança alta
- [Solução de Consulta Cosit nº 42, de 25 de maio de 2020](sc-cosit-42-2020.md) — leitura íntegra; confiança alta
- [Solução de Consulta Cosit nº 50, de 22 de junho de 2020](sc-cosit-50-2020.md) — leitura íntegra; confiança alta
- [Solução de Consulta Cosit nº 63, de 23 de junho de 2020](sc-cosit-63-2020.md) — leitura íntegra; confiança alta

### Isenção de R$ 20 mil e PME

- [Decreto nº 9.580, de 22 de novembro de 2018 (Regulamento do Imposto sobre a Renda — RIR/2018)](decreto-9580-2018-rir.md) — leitura parcial; confiança alta
- [Instrução Normativa RFB nº 1.022, de 5 de abril de 2010 (revogada)](in-rfb-1022-2010.md) — leitura parcial; confiança alta
- [Instrução Normativa RFB nº 1.500, de 29 de outubro de 2014](in-rfb-1500-2014.md) — leitura íntegra; confiança alta
- [Instrução Normativa RFB nº 1.585, de 31 de agosto de 2015](in-rfb-1585-2015.md) — leitura íntegra; confiança alta
- [Instrução Normativa RFB nº 1.916, de 18 de dezembro de 2019](in-rfb-1916-2019.md) — leitura íntegra; confiança alta
- [Instrução Normativa RFB nº 2.305, de 31 de dezembro de 2025 (com as INs RFB 2.306/2026 e 2.307/2026)](in-rfb-2305-2025.md) — leitura íntegra; confiança alta
- [Instruções Normativas RFB da DIRPF de cada exercício (2016 a 2026)](in-rfb-dirpf-anuais.md) — leitura íntegra; confiança alta
- [Lei Complementar nº 224, de 26 de dezembro de 2025](lc-224-2025.md) — leitura íntegra; confiança alta
- [Lei nº 11.033, de 21 de dezembro de 2004](lei-11033-2004.md) — leitura parcial; confiança alta
- [Lei nº 11.196, de 21 de novembro de 2005 ("Lei do Bem")](lei-11196-2005.md) — leitura parcial; confiança alta
- [Lei nº 13.043, de 13 de novembro de 2014](lei-13043-2014.md) — leitura parcial; confiança alta
- [Lei nº 15.270, de 26 de novembro de 2025](lei-15270-2025.md) — leitura íntegra; confiança alta
- [Lei nº 5.172, de 25 de outubro de 1966 (Código Tributário Nacional)](lei-5172-1966-ctn.md) — leitura parcial; confiança alta
- [Lei nº 8.981, de 20 de janeiro de 1995](lei-8981-1995.md) — leitura parcial; confiança alta
- [Lei nº 9.250, de 26 de dezembro de 1995](lei-9250-1995.md) — leitura parcial; confiança alta
- [Manual do Meu Imposto de Renda (MIR) — páginas da RFB](manual-mir.md) — leitura íntegra; confiança média
- [Medida Provisória nº 1.303, de 11 de junho de 2025 (perdeu a eficácia; não vale)](mp-1303-2025.md) — leitura íntegra; confiança alta
- [Perguntas e Respostas IRPF 2016 a 2020 — renda variável (por edição)](pr-irpf-2016-2020.md) — leitura íntegra; confiança alta
- [Perguntas e Respostas IRPF 2021 a 2026 — renda variável (por edição)](pr-irpf-2021-2026.md) — leitura íntegra; confiança alta
- [Perguntas e Respostas IRPF (RFB), edições 2016 a 2026 — renda variável](pr-irpf.md) — leitura íntegra; confiança alta
- [Solução de Consulta Cosit nº 145, de 21 de setembro de 2021](sc-cosit-145-2021.md) — leitura íntegra; confiança alta
- [Solução de Consulta Cosit nº 148, de 15 de agosto de 2025](sc-cosit-148-2025.md) — leitura parcial; confiança alta
- [Solução de Consulta Cosit nº 166, de 27 de setembro de 2021](sc-cosit-166-2021.md) — leitura íntegra; confiança alta
- [Solução de Consulta Cosit nº 39, de 12 de setembro de 2022](sc-cosit-39-2022.md) — leitura íntegra; confiança alta
- [Solução de Consulta Cosit nº 63, de 23 de junho de 2020](sc-cosit-63-2020.md) — leitura íntegra; confiança alta

### Day trade

- [Decreto nº 9.580, de 22 de novembro de 2018 (Regulamento do Imposto sobre a Renda — RIR/2018)](decreto-9580-2018-rir.md) — leitura parcial; confiança alta
- [Instrução Normativa RFB nº 1.022, de 5 de abril de 2010 (revogada)](in-rfb-1022-2010.md) — leitura parcial; confiança alta
- [Instrução Normativa RFB nº 1.585, de 31 de agosto de 2015](in-rfb-1585-2015.md) — leitura íntegra; confiança alta
- [Lei nº 11.033, de 21 de dezembro de 2004](lei-11033-2004.md) — leitura parcial; confiança alta
- [Lei nº 12.350, de 20 de dezembro de 2010](lei-12350-2010.md) — leitura parcial; confiança alta
- [Lei nº 8.981, de 20 de janeiro de 1995](lei-8981-1995.md) — leitura parcial; confiança alta
- [Lei nº 9.959, de 27 de janeiro de 2000](lei-9959-2000.md) — leitura íntegra; confiança alta
- [Medida Provisória nº 1.303, de 11 de junho de 2025 (perdeu a eficácia; não vale)](mp-1303-2025.md) — leitura íntegra; confiança alta
- [Medida Provisória nº 497, de 27 de julho de 2010](mp-497-2010.md) — leitura parcial; confiança alta
- [Perguntas e Respostas IRPF 2016 a 2020 — renda variável (por edição)](pr-irpf-2016-2020.md) — leitura íntegra; confiança alta
- [Perguntas e Respostas IRPF 2021 a 2026 — renda variável (por edição)](pr-irpf-2021-2026.md) — leitura íntegra; confiança alta
- [Perguntas e Respostas IRPF (RFB), edições 2016 a 2026 — renda variável](pr-irpf.md) — leitura íntegra; confiança alta
- [Solução de Consulta Cosit nº 127, de 8 de maio de 2024](sc-cosit-127-2024.md) — leitura íntegra; confiança alta

### FII e Fiagro

- [Instrução Normativa RFB nº 1.022, de 5 de abril de 2010 (revogada)](in-rfb-1022-2010.md) — leitura parcial; confiança alta
- [Instrução Normativa RFB nº 1.500, de 29 de outubro de 2014](in-rfb-1500-2014.md) — leitura íntegra; confiança alta
- [Instrução Normativa RFB nº 1.585, de 31 de agosto de 2015](in-rfb-1585-2015.md) — leitura íntegra; confiança alta
- [Instrução Normativa RFB nº 2.164, de 25 de outubro de 2023](in-rfb-2164-2023.md) — leitura íntegra; confiança alta
- [Instrução Normativa RFB nº 2.305, de 31 de dezembro de 2025 (com as INs RFB 2.306/2026 e 2.307/2026)](in-rfb-2305-2025.md) — leitura íntegra; confiança alta
- [Instruções Normativas RFB da DIRPF de cada exercício (2016 a 2026)](in-rfb-dirpf-anuais.md) — leitura íntegra; confiança alta
- [Lei Complementar nº 224, de 26 de dezembro de 2025](lc-224-2025.md) — leitura íntegra; confiança alta
- [Lei nº 11.033, de 21 de dezembro de 2004](lei-11033-2004.md) — leitura parcial; confiança alta
- [Lei nº 11.196, de 21 de novembro de 2005 ("Lei do Bem")](lei-11196-2005.md) — leitura parcial; confiança alta
- [Lei nº 14.130, de 29 de março de 2021](lei-14130-2021.md) — leitura íntegra; confiança alta
- [Lei nº 14.754, de 12 de dezembro de 2023](lei-14754-2023.md) — leitura íntegra; confiança alta
- [Lei nº 8.668, de 25 de junho de 1993](lei-8668-1993.md) — leitura íntegra; confiança alta
- [Lei nº 9.779, de 19 de janeiro de 1999](lei-9779-1999.md) — leitura íntegra; confiança alta
- [Manual do Meu Imposto de Renda (MIR) — páginas da RFB](manual-mir.md) — leitura íntegra; confiança média
- [Medida Provisória nº 1.303, de 11 de junho de 2025 (perdeu a eficácia; não vale)](mp-1303-2025.md) — leitura íntegra; confiança alta
- [Perguntas e Respostas IRPF 2021 a 2026 — renda variável (por edição)](pr-irpf-2021-2026.md) — leitura íntegra; confiança alta
- [Perguntas e Respostas IRPF (RFB), edições 2016 a 2026 — renda variável](pr-irpf.md) — leitura íntegra; confiança alta
- [Manual de utilização do ReVar (RFB)](revar.md) — leitura íntegra; confiança média
- [Solução de Consulta Cosit nº 166, de 27 de setembro de 2021](sc-cosit-166-2021.md) — leitura íntegra; confiança alta
- [Solução de Consulta Cosit nº 223, de 4 de dezembro de 2018](sc-cosit-223-2018.md) — leitura íntegra; confiança alta

### ETF

- [Instrução Normativa RFB nº 1.022, de 5 de abril de 2010 (revogada)](in-rfb-1022-2010.md) — leitura parcial; confiança alta
- [Instrução Normativa RFB nº 1.585, de 31 de agosto de 2015](in-rfb-1585-2015.md) — leitura íntegra; confiança alta
- [Instrução Normativa RFB nº 1.637, de 9 de maio de 2016](in-rfb-1637-2016.md) — leitura íntegra; confiança alta
- [Instrução Normativa RFB nº 2.164, de 25 de outubro de 2023](in-rfb-2164-2023.md) — leitura íntegra; confiança alta
- [Instrução Normativa RFB nº 2.180, de 11 de março de 2024](in-rfb-2180-2024.md) — leitura parcial; confiança média
- [Lei nº 13.043, de 13 de novembro de 2014](lei-13043-2014.md) — leitura parcial; confiança alta
- [Lei nº 14.754, de 12 de dezembro de 2023](lei-14754-2023.md) — leitura íntegra; confiança alta
- [Manual do Meu Imposto de Renda (MIR) — páginas da RFB](manual-mir.md) — leitura íntegra; confiança média
- [Perguntas e Respostas IRPF 2016 a 2020 — renda variável (por edição)](pr-irpf-2016-2020.md) — leitura íntegra; confiança alta
- [Perguntas e Respostas IRPF 2021 a 2026 — renda variável (por edição)](pr-irpf-2021-2026.md) — leitura íntegra; confiança alta
- [Perguntas e Respostas IRPF (RFB), edições 2016 a 2026 — renda variável](pr-irpf.md) — leitura íntegra; confiança alta
- [Manual de utilização do ReVar (RFB)](revar.md) — leitura íntegra; confiança média
- [Solução de Consulta Cosit nº 145, de 21 de setembro de 2021](sc-cosit-145-2021.md) — leitura íntegra; confiança alta
- [Solução de Consulta Cosit nº 73, de 24 de abril de 2026](sc-cosit-73-2026.md) — leitura íntegra; confiança alta

### BDR

- [Instrução Normativa RFB nº 1.585, de 31 de agosto de 2015](in-rfb-1585-2015.md) — leitura íntegra; confiança alta
- [Instrução Normativa RFB nº 2.164, de 25 de outubro de 2023](in-rfb-2164-2023.md) — leitura íntegra; confiança alta
- [Instrução Normativa RFB nº 2.180, de 11 de março de 2024](in-rfb-2180-2024.md) — leitura parcial; confiança média
- [Manual do Meu Imposto de Renda (MIR) — páginas da RFB](manual-mir.md) — leitura íntegra; confiança média
- [Manual de utilização do ReVar (RFB)](revar.md) — leitura íntegra; confiança média
- [Solução de Consulta Cosit nº 39, de 12 de setembro de 2022](sc-cosit-39-2022.md) — leitura íntegra; confiança alta

### Opções

- [Decreto nº 9.580, de 22 de novembro de 2018 (Regulamento do Imposto sobre a Renda — RIR/2018)](decreto-9580-2018-rir.md) — leitura parcial; confiança alta
- [Instrução Normativa RFB nº 1.022, de 5 de abril de 2010 (revogada)](in-rfb-1022-2010.md) — leitura parcial; confiança alta
- [Instrução Normativa RFB nº 1.585, de 31 de agosto de 2015](in-rfb-1585-2015.md) — leitura íntegra; confiança alta
- [Instrução Normativa RFB nº 2.164, de 25 de outubro de 2023](in-rfb-2164-2023.md) — leitura íntegra; confiança alta
- [Lei nº 11.033, de 21 de dezembro de 2004](lei-11033-2004.md) — leitura parcial; confiança alta
- [Lei nº 8.981, de 20 de janeiro de 1995](lei-8981-1995.md) — leitura parcial; confiança alta
- [Manual do Meu Imposto de Renda (MIR) — páginas da RFB](manual-mir.md) — leitura íntegra; confiança média
- [Perguntas e Respostas IRPF 2016 a 2020 — renda variável (por edição)](pr-irpf-2016-2020.md) — leitura íntegra; confiança alta
- [Perguntas e Respostas IRPF 2021 a 2026 — renda variável (por edição)](pr-irpf-2021-2026.md) — leitura íntegra; confiança alta
- [Perguntas e Respostas IRPF (RFB), edições 2016 a 2026 — renda variável](pr-irpf.md) — leitura íntegra; confiança alta

### Futuros

- [Decreto nº 9.580, de 22 de novembro de 2018 (Regulamento do Imposto sobre a Renda — RIR/2018)](decreto-9580-2018-rir.md) — leitura parcial; confiança alta
- [Instrução Normativa RFB nº 1.022, de 5 de abril de 2010 (revogada)](in-rfb-1022-2010.md) — leitura parcial; confiança alta
- [Instrução Normativa RFB nº 1.585, de 31 de agosto de 2015](in-rfb-1585-2015.md) — leitura íntegra; confiança alta
- [Instrução Normativa RFB nº 2.164, de 25 de outubro de 2023](in-rfb-2164-2023.md) — leitura íntegra; confiança alta
- [Lei nº 11.033, de 21 de dezembro de 2004](lei-11033-2004.md) — leitura parcial; confiança alta
- [Lei nº 11.051, de 29 de dezembro de 2004](lei-11051-2004.md) — leitura parcial; confiança alta
- [Lei nº 8.981, de 20 de janeiro de 1995](lei-8981-1995.md) — leitura parcial; confiança alta
- [Perguntas e Respostas IRPF 2016 a 2020 — renda variável (por edição)](pr-irpf-2016-2020.md) — leitura íntegra; confiança alta
- [Perguntas e Respostas IRPF 2021 a 2026 — renda variável (por edição)](pr-irpf-2021-2026.md) — leitura íntegra; confiança alta
- [Perguntas e Respostas IRPF (RFB), edições 2016 a 2026 — renda variável](pr-irpf.md) — leitura íntegra; confiança alta

### Termo

- [Decreto nº 9.580, de 22 de novembro de 2018 (Regulamento do Imposto sobre a Renda — RIR/2018)](decreto-9580-2018-rir.md) — leitura parcial; confiança alta
- [Instrução Normativa RFB nº 1.022, de 5 de abril de 2010 (revogada)](in-rfb-1022-2010.md) — leitura parcial; confiança alta
- [Instrução Normativa RFB nº 1.585, de 31 de agosto de 2015](in-rfb-1585-2015.md) — leitura íntegra; confiança alta
- [Lei nº 11.033, de 21 de dezembro de 2004](lei-11033-2004.md) — leitura parcial; confiança alta
- [Lei nº 8.981, de 20 de janeiro de 1995](lei-8981-1995.md) — leitura parcial; confiança alta
- [Perguntas e Respostas IRPF 2016 a 2020 — renda variável (por edição)](pr-irpf-2016-2020.md) — leitura íntegra; confiança alta
- [Perguntas e Respostas IRPF 2021 a 2026 — renda variável (por edição)](pr-irpf-2021-2026.md) — leitura íntegra; confiança alta
- [Perguntas e Respostas IRPF (RFB), edições 2016 a 2026 — renda variável](pr-irpf.md) — leitura íntegra; confiança alta

### BTC (empréstimo de ações)

- [Decreto nº 9.580, de 22 de novembro de 2018 (Regulamento do Imposto sobre a Renda — RIR/2018)](decreto-9580-2018-rir.md) — leitura parcial; confiança alta
- [Instrução Normativa RFB nº 1.022, de 5 de abril de 2010 (revogada)](in-rfb-1022-2010.md) — leitura parcial; confiança alta
- [Instrução Normativa RFB nº 1.585, de 31 de agosto de 2015](in-rfb-1585-2015.md) — leitura íntegra; confiança alta
- [Instrução Normativa RFB nº 2.164, de 25 de outubro de 2023](in-rfb-2164-2023.md) — leitura íntegra; confiança alta
- [Lei nº 13.043, de 13 de novembro de 2014](lei-13043-2014.md) — leitura parcial; confiança alta
- [Lei nº 15.265, de 21 de novembro de 2025](lei-15265-2025.md) — leitura íntegra; confiança alta
- [Medida Provisória nº 1.303, de 11 de junho de 2025 (perdeu a eficácia; não vale)](mp-1303-2025.md) — leitura íntegra; confiança alta
- [Manual de utilização do ReVar (RFB)](revar.md) — leitura íntegra; confiança média
- [Solução de Consulta Cosit nº 153, de 18 de novembro de 2016](sc-cosit-153-2016.md) — leitura íntegra; confiança alta
- [Solução de Consulta Cosit nº 42, de 25 de maio de 2020](sc-cosit-42-2020.md) — leitura íntegra; confiança alta

### Proventos e JCP

- [Decreto nº 9.580, de 22 de novembro de 2018 (Regulamento do Imposto sobre a Renda — RIR/2018)](decreto-9580-2018-rir.md) — leitura parcial; confiança alta
- [Instrução Normativa RFB nº 1.500, de 29 de outubro de 2014](in-rfb-1500-2014.md) — leitura íntegra; confiança alta
- [Instrução Normativa RFB nº 1.585, de 31 de agosto de 2015](in-rfb-1585-2015.md) — leitura íntegra; confiança alta
- [Instrução Normativa RFB nº 1.700, de 14 de março de 2017](in-rfb-1700-2017.md) — leitura parcial; confiança média
- [Instrução Normativa RFB nº 2.299, de 17 de dezembro de 2025](in-rfb-2299-2025.md) — leitura íntegra; confiança alta
- [Instrução Normativa SRF nº 12, de 10 de fevereiro de 1999](in-srf-12-1999.md) — leitura íntegra; confiança média
- [Instrução Normativa SRF nº 41, de 22 de abril de 1998](in-srf-41-1998.md) — leitura íntegra; confiança alta
- [Lei Complementar nº 224, de 26 de dezembro de 2025](lc-224-2025.md) — leitura íntegra; confiança alta
- [Lei nº 11.478, de 29 de maio de 2007](lei-11478-2007.md) — leitura íntegra; confiança alta
- [Lei nº 13.043, de 13 de novembro de 2014](lei-13043-2014.md) — leitura parcial; confiança alta
- [Lei nº 14.130, de 29 de março de 2021](lei-14130-2021.md) — leitura íntegra; confiança alta
- [Lei nº 14.754, de 12 de dezembro de 2023](lei-14754-2023.md) — leitura íntegra; confiança alta
- [Lei nº 15.265, de 21 de novembro de 2025](lei-15265-2025.md) — leitura íntegra; confiança alta
- [Lei nº 15.270, de 26 de novembro de 2025](lei-15270-2025.md) — leitura íntegra; confiança alta
- [Lei nº 6.404, de 15 de dezembro de 1976 (Lei das S.A.)](lei-6404-1976.md) — leitura parcial; confiança alta
- [Lei nº 7.713, de 22 de dezembro de 1988](lei-7713-1988.md) — leitura parcial; confiança alta
- [Lei nº 8.668, de 25 de junho de 1993](lei-8668-1993.md) — leitura íntegra; confiança alta
- [Lei nº 9.249, de 26 de dezembro de 1995](lei-9249-1995.md) — leitura parcial; confiança alta
- [Lei nº 9.250, de 26 de dezembro de 1995](lei-9250-1995.md) — leitura parcial; confiança alta
- [Manual do Meu Imposto de Renda (MIR) — páginas da RFB](manual-mir.md) — leitura íntegra; confiança média
- [Medida Provisória nº 1.303, de 11 de junho de 2025 (perdeu a eficácia; não vale)](mp-1303-2025.md) — leitura íntegra; confiança alta
- [Perguntas e Respostas — Tributação de Altas Rendas: lucros e dividendos (RFB, versão de 19/12/2025)](pr-altas-rendas-2025.md) — leitura íntegra; confiança alta
- [Perguntas e Respostas IRPF 2016 a 2020 — renda variável (por edição)](pr-irpf-2016-2020.md) — leitura íntegra; confiança alta
- [Perguntas e Respostas IRPF 2021 a 2026 — renda variável (por edição)](pr-irpf-2021-2026.md) — leitura íntegra; confiança alta
- [Perguntas e Respostas IRPF (RFB), edições 2016 a 2026 — renda variável](pr-irpf.md) — leitura íntegra; confiança alta
- [Solução de Consulta Cosit nº 153, de 18 de novembro de 2016](sc-cosit-153-2016.md) — leitura íntegra; confiança alta
- [Solução de Consulta Cosit nº 67, de 1º de abril de 2024](sc-cosit-67-2024.md) — leitura íntegra; confiança alta

### Eventos societários e direitos de subscrição

- [Instrução Normativa RFB nº 1.585, de 31 de agosto de 2015](in-rfb-1585-2015.md) — leitura íntegra; confiança alta
- [Instrução Normativa RFB nº 1.916, de 18 de dezembro de 2019](in-rfb-1916-2019.md) — leitura íntegra; confiança alta
- [Lei nº 6.404, de 15 de dezembro de 1976 (Lei das S.A.)](lei-6404-1976.md) — leitura parcial; confiança alta
- [Lei nº 7.713, de 22 de dezembro de 1988](lei-7713-1988.md) — leitura parcial; confiança alta
- [Lei nº 9.249, de 26 de dezembro de 1995](lei-9249-1995.md) — leitura parcial; confiança alta
- [Perguntas e Respostas — Tributação de Altas Rendas: lucros e dividendos (RFB, versão de 19/12/2025)](pr-altas-rendas-2025.md) — leitura íntegra; confiança alta
- [Perguntas e Respostas IRPF 2016 a 2020 — renda variável (por edição)](pr-irpf-2016-2020.md) — leitura íntegra; confiança alta
- [Perguntas e Respostas IRPF 2021 a 2026 — renda variável (por edição)](pr-irpf-2021-2026.md) — leitura íntegra; confiança alta
- [Perguntas e Respostas IRPF (RFB), edições 2016 a 2026 — renda variável](pr-irpf.md) — leitura íntegra; confiança alta
- [Solução de Consulta Cosit nº 10, de 3 de fevereiro de 2016](sc-cosit-10-2016.md) — leitura íntegra; confiança alta
- [Solução de Consulta Cosit nº 130, de 26 de junho de 2023](sc-cosit-130-2023.md) — leitura íntegra; confiança alta
- [Solução de Consulta Cosit nº 145, de 21 de setembro de 2021](sc-cosit-145-2021.md) — leitura íntegra; confiança alta
- [Solução de Consulta Cosit nº 148, de 15 de agosto de 2025](sc-cosit-148-2025.md) — leitura parcial; confiança alta
- [Solução de Consulta Cosit nº 166, de 27 de setembro de 2021](sc-cosit-166-2021.md) — leitura íntegra; confiança alta
- [Solução de Consulta Cosit nº 224, de 14 de agosto de 2014](sc-cosit-224-2014.md) — leitura íntegra; confiança alta
- [Solução de Consulta Cosit nº 37, de 31 de março de 2020](sc-cosit-37-2020.md) — leitura íntegra; confiança alta

### FI-Infra e FIP-IE

- [Instrução Normativa RFB nº 1.585, de 31 de agosto de 2015](in-rfb-1585-2015.md) — leitura íntegra; confiança alta
- [Instrução Normativa RFB nº 1.637, de 9 de maio de 2016](in-rfb-1637-2016.md) — leitura íntegra; confiança alta
- [Instrução Normativa RFB nº 2.164, de 25 de outubro de 2023](in-rfb-2164-2023.md) — leitura íntegra; confiança alta
- [Instrução Normativa RFB nº 2.305, de 31 de dezembro de 2025 (com as INs RFB 2.306/2026 e 2.307/2026)](in-rfb-2305-2025.md) — leitura íntegra; confiança alta
- [Lei Complementar nº 224, de 26 de dezembro de 2025](lc-224-2025.md) — leitura íntegra; confiança alta
- [Lei nº 11.478, de 29 de maio de 2007](lei-11478-2007.md) — leitura íntegra; confiança alta
- [Lei nº 12.431, de 24 de junho de 2011](lei-12431-2011.md) — leitura parcial; confiança alta
- [Lei nº 14.754, de 12 de dezembro de 2023](lei-14754-2023.md) — leitura íntegra; confiança alta
- [Manual do Meu Imposto de Renda (MIR) — páginas da RFB](manual-mir.md) — leitura íntegra; confiança média
- [Medida Provisória nº 1.303, de 11 de junho de 2025 (perdeu a eficácia; não vale)](mp-1303-2025.md) — leitura íntegra; confiança alta
- [Perguntas e Respostas IRPF 2021 a 2026 — renda variável (por edição)](pr-irpf-2021-2026.md) — leitura íntegra; confiança alta
- [Perguntas e Respostas IRPF (RFB), edições 2016 a 2026 — renda variável](pr-irpf.md) — leitura íntegra; confiança alta
- [Manual de utilização do ReVar (RFB)](revar.md) — leitura íntegra; confiança média
- [Solução de Consulta Cosit nº 103, de 25 de março de 2019](sc-cosit-103-2019.md) — leitura íntegra; confiança alta
- [Solução de Consulta Cosit nº 73, de 24 de abril de 2026](sc-cosit-73-2026.md) — leitura íntegra; confiança alta

### IRRF

- [Decreto nº 9.580, de 22 de novembro de 2018 (Regulamento do Imposto sobre a Renda — RIR/2018)](decreto-9580-2018-rir.md) — leitura parcial; confiança alta
- [Instrução Normativa RFB nº 1.022, de 5 de abril de 2010 (revogada)](in-rfb-1022-2010.md) — leitura parcial; confiança alta
- [Instrução Normativa RFB nº 1.500, de 29 de outubro de 2014](in-rfb-1500-2014.md) — leitura íntegra; confiança alta
- [Instrução Normativa RFB nº 1.585, de 31 de agosto de 2015](in-rfb-1585-2015.md) — leitura íntegra; confiança alta
- [Instrução Normativa RFB nº 1.637, de 9 de maio de 2016](in-rfb-1637-2016.md) — leitura íntegra; confiança alta
- [Instrução Normativa RFB nº 1.720, de 20 de julho de 2017](in-rfb-1720-2017.md) — leitura íntegra; confiança alta
- [Instrução Normativa RFB nº 1.916, de 18 de dezembro de 2019](in-rfb-1916-2019.md) — leitura íntegra; confiança alta
- [Instrução Normativa RFB nº 2.166, de 15 de dezembro de 2023](in-rfb-2166-2023.md) — leitura íntegra; confiança alta
- [Instrução Normativa RFB nº 2.253, de 21 de fevereiro de 2025](in-rfb-2253-2025.md) — leitura íntegra; confiança alta
- [Instrução Normativa RFB nº 2.299, de 17 de dezembro de 2025](in-rfb-2299-2025.md) — leitura íntegra; confiança alta
- [Instrução Normativa SRF nº 12, de 10 de fevereiro de 1999](in-srf-12-1999.md) — leitura íntegra; confiança média
- [Lei nº 11.033, de 21 de dezembro de 2004](lei-11033-2004.md) — leitura parcial; confiança alta
- [Lei nº 11.196, de 21 de novembro de 2005 ("Lei do Bem")](lei-11196-2005.md) — leitura parcial; confiança alta
- [Lei nº 12.350, de 20 de dezembro de 2010](lei-12350-2010.md) — leitura parcial; confiança alta
- [Lei nº 13.043, de 13 de novembro de 2014](lei-13043-2014.md) — leitura parcial; confiança alta
- [Lei nº 15.270, de 26 de novembro de 2025](lei-15270-2025.md) — leitura íntegra; confiança alta
- [Lei nº 8.668, de 25 de junho de 1993](lei-8668-1993.md) — leitura íntegra; confiança alta
- [Lei nº 8.981, de 20 de janeiro de 1995](lei-8981-1995.md) — leitura parcial; confiança alta
- [Lei nº 9.249, de 26 de dezembro de 1995](lei-9249-1995.md) — leitura parcial; confiança alta
- [Lei nº 9.779, de 19 de janeiro de 1999](lei-9779-1999.md) — leitura íntegra; confiança alta
- [Lei nº 9.959, de 27 de janeiro de 2000](lei-9959-2000.md) — leitura íntegra; confiança alta
- [Manual do Meu Imposto de Renda (MIR) — páginas da RFB](manual-mir.md) — leitura íntegra; confiança média
- [Medida Provisória nº 497, de 27 de julho de 2010](mp-497-2010.md) — leitura parcial; confiança alta
- [Perguntas e Respostas — Tributação de Altas Rendas: lucros e dividendos (RFB, versão de 19/12/2025)](pr-altas-rendas-2025.md) — leitura íntegra; confiança alta
- [Perguntas e Respostas IRPF 2016 a 2020 — renda variável (por edição)](pr-irpf-2016-2020.md) — leitura íntegra; confiança alta
- [Perguntas e Respostas IRPF 2021 a 2026 — renda variável (por edição)](pr-irpf-2021-2026.md) — leitura íntegra; confiança alta
- [Perguntas e Respostas IRPF (RFB), edições 2016 a 2026 — renda variável](pr-irpf.md) — leitura íntegra; confiança alta
- [Manual de utilização do ReVar (RFB)](revar.md) — leitura íntegra; confiança média
- [Solução de Consulta Cosit nº 63, de 23 de junho de 2020](sc-cosit-63-2020.md) — leitura íntegra; confiança alta

### DARF, multa e juros

- [Lei Complementar nº 236, de 4 de setembro de 2026](lc-236-2026.md) — leitura íntegra; confiança alta (média para a data do DOU)
- [Decreto nº 9.580, de 22 de novembro de 2018 (Regulamento do Imposto sobre a Renda — RIR/2018)](decreto-9580-2018-rir.md) — leitura parcial; confiança alta
- [Instrução Normativa RFB nº 1.500, de 29 de outubro de 2014](in-rfb-1500-2014.md) — leitura íntegra; confiança alta
- [Instrução Normativa RFB nº 1.585, de 31 de agosto de 2015](in-rfb-1585-2015.md) — leitura íntegra; confiança alta
- [Instrução Normativa RFB nº 2.164, de 25 de outubro de 2023](in-rfb-2164-2023.md) — leitura íntegra; confiança alta
- [Lei nº 14.759, de 21 de dezembro de 2023](lei-14759-2023.md) — leitura íntegra; confiança alta
- [Lei nº 5.172, de 25 de outubro de 1966 (Código Tributário Nacional)](lei-5172-1966-ctn.md) — leitura parcial; confiança alta
- [Lei nº 9.430, de 27 de dezembro de 1996](lei-9430-1996.md) — leitura parcial; confiança alta
- [Manual do Meu Imposto de Renda (MIR) — páginas da RFB](manual-mir.md) — leitura íntegra; confiança média
- [Perguntas e Respostas IRPF 2016 a 2020 — renda variável (por edição)](pr-irpf-2016-2020.md) — leitura íntegra; confiança alta
- [Perguntas e Respostas IRPF 2021 a 2026 — renda variável (por edição)](pr-irpf-2021-2026.md) — leitura íntegra; confiança alta
- [Perguntas e Respostas IRPF (RFB), edições 2016 a 2026 — renda variável](pr-irpf.md) — leitura íntegra; confiança alta
- [Resolução CMN nº 2.932, de 28 de fevereiro de 2002](res-cmn-2932-2002.md) — leitura íntegra; confiança alta
- [Resolução CMN nº 4.880, de 23 de dezembro de 2020](res-cmn-4880-2020.md) — leitura íntegra; confiança alta
- [Manual de utilização do ReVar (RFB)](revar.md) — leitura íntegra; confiança média
- [Solução de Consulta Cosit nº 80, de 4 de abril de 2024](sc-cosit-80-2024.md) — leitura íntegra; confiança alta

### Fichas do IRPF e obrigatoriedade

- [Instrução Normativa RFB nº 1.500, de 29 de outubro de 2014](in-rfb-1500-2014.md) — leitura íntegra; confiança alta
- [Instrução Normativa RFB nº 1.756, de 31 de outubro de 2017](in-rfb-1756-2017.md) — leitura íntegra; confiança alta
- [Instrução Normativa RFB nº 2.299, de 17 de dezembro de 2025](in-rfb-2299-2025.md) — leitura íntegra; confiança alta
- [Instruções Normativas RFB da DIRPF de cada exercício (2016 a 2026)](in-rfb-dirpf-anuais.md) — leitura íntegra; confiança alta
- [Instrução Normativa SRF nº 41, de 22 de abril de 1998](in-srf-41-1998.md) — leitura íntegra; confiança alta
- [Lei nº 11.478, de 29 de maio de 2007](lei-11478-2007.md) — leitura íntegra; confiança alta
- [Lei nº 12.431, de 24 de junho de 2011](lei-12431-2011.md) — leitura parcial; confiança alta
- [Lei nº 15.270, de 26 de novembro de 2025](lei-15270-2025.md) — leitura íntegra; confiança alta
- [Lei nº 9.250, de 26 de dezembro de 1995](lei-9250-1995.md) — leitura parcial; confiança alta
- [Manual do Meu Imposto de Renda (MIR) — páginas da RFB](manual-mir.md) — leitura íntegra; confiança média
- [Perguntas e Respostas IRPF (RFB), edições 2016 a 2026 — renda variável](pr-irpf.md) — leitura íntegra; confiança alta
- [Solução de Consulta Cosit nº 103, de 25 de março de 2019](sc-cosit-103-2019.md) — leitura íntegra; confiança alta
- [Solução de Consulta Cosit nº 80, de 4 de abril de 2024](sc-cosit-80-2024.md) — leitura íntegra; confiança alta

### Prejuízo, herança e outros

- [Lei nº 5.172, de 25 de outubro de 1966 (Código Tributário Nacional)](lei-5172-1966-ctn.md) — leitura parcial; confiança alta
- [Solução de Consulta Cosit nº 130, de 26 de junho de 2023](sc-cosit-130-2023.md) — leitura íntegra; confiança alta
- [Solução de Consulta Cosit nº 182, de 24 de junho de 2024](sc-cosit-182-2024.md) — leitura íntegra; confiança alta
- [Solução de Consulta Cosit nº 223, de 4 de dezembro de 2018](sc-cosit-223-2018.md) — leitura íntegra; confiança alta
- [Solução de Consulta Cosit nº 50, de 22 de junho de 2020](sc-cosit-50-2020.md) — leitura íntegra; confiança alta

## Todas as fichas

| Ficha | Leitura | Confiança |
|---|---|---|
| [Decreto nº 9.580, de 22 de novembro de 2018 (Regulamento do Imposto sobre a Renda — RIR/2018)](decreto-9580-2018-rir.md) | parcial | alta |
| [Instrução Normativa RFB nº 1.022, de 5 de abril de 2010 (revogada)](in-rfb-1022-2010.md) | parcial | alta |
| [Instrução Normativa RFB nº 1.500, de 29 de outubro de 2014](in-rfb-1500-2014.md) | íntegra | alta |
| [Instrução Normativa RFB nº 1.585, de 31 de agosto de 2015](in-rfb-1585-2015.md) | íntegra | alta |
| [Instrução Normativa RFB nº 1.637, de 9 de maio de 2016](in-rfb-1637-2016.md) | íntegra | alta |
| [Instrução Normativa RFB nº 1.700, de 14 de março de 2017](in-rfb-1700-2017.md) | parcial | média |
| [Instrução Normativa RFB nº 1.720, de 20 de julho de 2017](in-rfb-1720-2017.md) | íntegra | alta |
| [Instrução Normativa RFB nº 1.756, de 31 de outubro de 2017](in-rfb-1756-2017.md) | íntegra | alta |
| [Instrução Normativa RFB nº 1.916, de 18 de dezembro de 2019](in-rfb-1916-2019.md) | íntegra | alta |
| [Instrução Normativa RFB nº 2.164, de 25 de outubro de 2023](in-rfb-2164-2023.md) | íntegra | alta |
| [Instrução Normativa RFB nº 2.166, de 15 de dezembro de 2023](in-rfb-2166-2023.md) | íntegra | alta |
| [Instrução Normativa RFB nº 2.180, de 11 de março de 2024](in-rfb-2180-2024.md) | parcial | média |
| [Instrução Normativa RFB nº 2.253, de 21 de fevereiro de 2025](in-rfb-2253-2025.md) | íntegra | alta |
| [Instrução Normativa RFB nº 2.299, de 17 de dezembro de 2025](in-rfb-2299-2025.md) | íntegra | alta |
| [Instrução Normativa RFB nº 2.305, de 31 de dezembro de 2025 (com as INs RFB 2.306/2026 e 2.307/2026)](in-rfb-2305-2025.md) | íntegra | alta |
| [Instruções Normativas RFB da DIRPF de cada exercício (2016 a 2026)](in-rfb-dirpf-anuais.md) | íntegra | alta |
| [Instrução Normativa SRF nº 12, de 10 de fevereiro de 1999](in-srf-12-1999.md) | íntegra | média |
| [Instrução Normativa SRF nº 41, de 22 de abril de 1998](in-srf-41-1998.md) | íntegra | alta |
| [Lei Complementar nº 224, de 26 de dezembro de 2025](lc-224-2025.md) | íntegra | alta |
| [Lei nº 11.033, de 21 de dezembro de 2004](lei-11033-2004.md) | parcial | alta |
| [Lei nº 11.051, de 29 de dezembro de 2004](lei-11051-2004.md) | parcial | alta |
| [Lei nº 11.196, de 21 de novembro de 2005 ("Lei do Bem")](lei-11196-2005.md) | parcial | alta |
| [Lei nº 11.478, de 29 de maio de 2007](lei-11478-2007.md) | íntegra | alta |
| [Lei nº 12.350, de 20 de dezembro de 2010](lei-12350-2010.md) | parcial | alta |
| [Lei nº 12.431, de 24 de junho de 2011](lei-12431-2011.md) | parcial | alta |
| [Lei nº 13.043, de 13 de novembro de 2014](lei-13043-2014.md) | parcial | alta |
| [Lei nº 14.130, de 29 de março de 2021](lei-14130-2021.md) | íntegra | alta |
| [Lei nº 14.754, de 12 de dezembro de 2023](lei-14754-2023.md) | íntegra | alta |
| [Lei Complementar nº 236, de 4 de setembro de 2026](lc-236-2026.md) | íntegra | alta |
| [Lei nº 14.759, de 21 de dezembro de 2023](lei-14759-2023.md) | íntegra | alta |
| [Lei nº 15.265, de 21 de novembro de 2025](lei-15265-2025.md) | íntegra | alta |
| [Lei nº 15.270, de 26 de novembro de 2025](lei-15270-2025.md) | íntegra | alta |
| [Lei nº 5.172, de 25 de outubro de 1966 (Código Tributário Nacional)](lei-5172-1966-ctn.md) | parcial | alta |
| [Lei nº 6.404, de 15 de dezembro de 1976 (Lei das S.A.)](lei-6404-1976.md) | parcial | alta |
| [Lei nº 7.713, de 22 de dezembro de 1988](lei-7713-1988.md) | parcial | alta |
| [Lei nº 8.668, de 25 de junho de 1993](lei-8668-1993.md) | íntegra | alta |
| [Lei nº 8.981, de 20 de janeiro de 1995](lei-8981-1995.md) | parcial | alta |
| [Lei nº 9.249, de 26 de dezembro de 1995](lei-9249-1995.md) | parcial | alta |
| [Lei nº 9.250, de 26 de dezembro de 1995](lei-9250-1995.md) | parcial | alta |
| [Lei nº 9.430, de 27 de dezembro de 1996](lei-9430-1996.md) | parcial | alta |
| [Lei nº 9.779, de 19 de janeiro de 1999](lei-9779-1999.md) | íntegra | alta |
| [Lei nº 9.959, de 27 de janeiro de 2000](lei-9959-2000.md) | íntegra | alta |
| [Manual do Meu Imposto de Renda (MIR) — páginas da RFB](manual-mir.md) | íntegra | média |
| [Medida Provisória nº 1.303, de 11 de junho de 2025 (perdeu a eficácia; não vale)](mp-1303-2025.md) | íntegra | alta |
| [Medida Provisória nº 497, de 27 de julho de 2010](mp-497-2010.md) | parcial | alta |
| [Perguntas e Respostas — Tributação de Altas Rendas: lucros e dividendos (RFB, versão de 19/12/2025)](pr-altas-rendas-2025.md) | íntegra | alta |
| [Perguntas e Respostas IRPF 2016 a 2020 — renda variável (por edição)](pr-irpf-2016-2020.md) | íntegra | alta |
| [Perguntas e Respostas IRPF 2021 a 2026 — renda variável (por edição)](pr-irpf-2021-2026.md) | íntegra | alta |
| [Perguntas e Respostas IRPF (RFB), edições 2016 a 2026 — renda variável](pr-irpf.md) | íntegra | alta |
| [Resolução CMN nº 2.932, de 28 de fevereiro de 2002](res-cmn-2932-2002.md) | íntegra | alta |
| [Resolução CMN nº 4.880, de 23 de dezembro de 2020](res-cmn-4880-2020.md) | íntegra | alta |
| [Manual de utilização do ReVar (RFB)](revar.md) | íntegra | média |
| [Solução de Consulta Cosit nº 10, de 3 de fevereiro de 2016](sc-cosit-10-2016.md) | íntegra | alta |
| [Solução de Consulta Cosit nº 103, de 25 de março de 2019](sc-cosit-103-2019.md) | íntegra | alta |
| [Solução de Consulta Cosit nº 127, de 8 de maio de 2024](sc-cosit-127-2024.md) | íntegra | alta |
| [Solução de Consulta Cosit nº 130, de 26 de junho de 2023](sc-cosit-130-2023.md) | íntegra | alta |
| [Solução de Consulta Cosit nº 145, de 21 de setembro de 2021](sc-cosit-145-2021.md) | íntegra | alta |
| [Solução de Consulta Cosit nº 148, de 15 de agosto de 2025](sc-cosit-148-2025.md) | parcial | alta |
| [Solução de Consulta Cosit nº 153, de 18 de novembro de 2016](sc-cosit-153-2016.md) | íntegra | alta |
| [Solução de Consulta Cosit nº 166, de 27 de setembro de 2021](sc-cosit-166-2021.md) | íntegra | alta |
| [Solução de Consulta Cosit nº 182, de 24 de junho de 2024](sc-cosit-182-2024.md) | íntegra | alta |
| [Solução de Consulta Cosit nº 223, de 4 de dezembro de 2018](sc-cosit-223-2018.md) | íntegra | alta |
| [Solução de Consulta Cosit nº 224, de 14 de agosto de 2014](sc-cosit-224-2014.md) | íntegra | alta |
| [Solução de Consulta Cosit nº 227, de 4 de dezembro de 2018](sc-cosit-227-2018.md) | íntegra | alta |
| [Solução de Consulta Cosit nº 37, de 31 de março de 2020](sc-cosit-37-2020.md) | íntegra | alta |
| [Solução de Consulta Cosit nº 39, de 12 de setembro de 2022](sc-cosit-39-2022.md) | íntegra | alta |
| [Solução de Consulta Cosit nº 42, de 25 de maio de 2020](sc-cosit-42-2020.md) | íntegra | alta |
| [Solução de Consulta Cosit nº 50, de 22 de junho de 2020](sc-cosit-50-2020.md) | íntegra | alta |
| [Solução de Consulta Cosit nº 63, de 23 de junho de 2020](sc-cosit-63-2020.md) | íntegra | alta |
| [Solução de Consulta Cosit nº 67, de 1º de abril de 2024](sc-cosit-67-2024.md) | íntegra | alta |
| [Solução de Consulta Cosit nº 73, de 24 de abril de 2026](sc-cosit-73-2026.md) | íntegra | alta |
| [Solução de Consulta Cosit nº 80, de 4 de abril de 2024](sc-cosit-80-2024.md) | íntegra | alta |
