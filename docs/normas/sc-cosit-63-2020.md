# Solução de Consulta Cosit nº 63, de 23 de junho de 2020

| Campo | Valor |
|---|---|
| Ementa (resumo) | Ações de PME (Lei 13.043/2014 arts. 16-17) compradas quando a companhia cumpria os requisitos continuam isentas até 31/12/2023, mesmo que oferta subsequente a desenquadre; as compradas a partir da oferta subsequente são tributadas |
| Fonte oficial | https://normas.receita.fazenda.gov.br/sijut2consulta/link.action?idAto=110693 |
| Texto lido | https://normasinternet2.receita.fazenda.gov.br/api/consulta-externa/ato/110693/visao/multivigente e anexo 57010 (íntegra) |
| Lido em | 2026-11-10 |
| Leitura | íntegra |
| Confiança | alta |
| Temas | ações; isenção PME; IRRF |

## Vigência e alterações

| Dispositivo | Redação | Vigente de | até | Alterado por |
|---|---|---|---|---|
| SC inteira | original | 29/06/2020 (DOU) | — | não reformada; a isenção que interpreta acabou em 31/12/2023 (Lei 13.043 art. 16 caput) |

## Dispositivos relevantes para o COIR

- **§§13-16** — os requisitos do art. 16 se medem em 10/07/2014, na oferta inicial ou em cada oferta subsequente; quem comprou a partir de 10/07/2014 com a companhia enquadrada tem isenção até 31/12/2023.
- **§§18-20** — transcreve a IN 1.585 arts. 59, 66, 67, 68 e 69-A (IN 1.916/2019): 69-A §2º mantém isentas as ações anteriores à oferta subsequente; §1º inclui direito de preferência e bonificação recebidos depois dela entre as tributadas; §3º o IRRF retido pode ser compensado; §4º controle separado.
- **§§21-22** — recomenda controlar as aquisições por data; a ação comprada depois da oferta subsequente pode ainda ter a isenção dos R$ 20 mil (art. 59 I).
- Transcreve também art. 17 §§2º e 3º (perdas não compensam; valor não entra no limite) e §8º/art. 67 §6º (sem retenção de 0,005%).
- A SC **não** fixa ordem de saída na venda (isento ou tributado primeiro).

## Onde o COIR usa

- `docs/regras-fiscais.md` (lotes de PME) e `tests/fiscal_cases/test_pme_lotes.py`: "o lote anterior continua isento mesmo com a companhia desenquadrada" — **confere**.

## Pontos em aberto

- Ordem de venda entre lotes: ver `sc-cosit-148-2025.md` (a RFB, em caso análogo do DL 1.510/1976, rejeitou o PEPS e adotou a proporcionalidade).
