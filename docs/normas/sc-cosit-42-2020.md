# Solução de Consulta Cosit nº 42, de 25 de maio de 2020

| Campo | Valor |
|---|---|
| Ementa (resumo) | Tomador pessoa física no BTC não pode deduzir a remuneração paga ao doador (nem no ganho líquido, nem no ajuste anual); a dedução é só para PJ do lucro real |
| Fonte oficial | https://normas.receita.fazenda.gov.br/sijut2consulta/link.action?idAto=110035 |
| Texto lido | https://normasinternet2.receita.fazenda.gov.br/api/consulta-externa/ato/110035/visao/multivigente e anexo 56722 (íntegra) |
| Lido em | 2026-11-10 |
| Leitura | íntegra |
| Confiança | alta |
| Temas | BTC; custos |

## Vigência e alterações

| Dispositivo | Redação | Vigente de | até | Alterado por |
|---|---|---|---|---|
| SC inteira | original | 03/06/2020 (DOU) | — | não reformada. Base legal (Lei 13.043/2014 arts. 6º, 9º e 10) foi alterada pela Lei 15.265/2025 a partir de 2026: a conclusão segue pela mesma lógica (dedução só no lucro real), mas a SC não interpreta a lei nova |

## Dispositivos relevantes para o COIR

- **§19** — o tomador (PF ou PJ) pode computar como custo corretagens e emolumentos da operação (IN 1.585 arts. 76 e 77; Lei 13.043 art. 9º par. único e art. 10 §2º).
- **§§23-25 e conclusão** — a remuneração paga ao doador (e o JCP reembolsado, pergunta do consulente) é despesa só para PJ do lucro real (Lei 13.043 art. 6º §1º); para PF não reduz o ganho líquido nem o ajuste anual (legalidade estrita; CF art. 150 I e §6º; CTN arts. 97 e 99).

## Onde o COIR usa

- `docs/regras-fiscais.md` (Outros custos dedutíveis), `coir/ui/telas/outros_custos.py`, `docs/manual/telas/outros-custos.md`, `tests/fiscal_cases/test_outros_custos.py`, `CHANGELOG.md`: **confere**.
- `docs/regras-fiscais.md` linha "A conferir: o PDF oficial da SC 42/2020 ... não puderam ser lidos": **desatualizado** — a íntegra foi lida em 10/11/2026 e confirma o parecer.

## Pontos em aberto

- O consulente perguntou também sobre o reembolso de JCP pago pelo tomador; a SC responde junto (não dedutível por PF).
