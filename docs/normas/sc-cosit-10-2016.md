# Solução de Consulta Cosit nº 10, de 3 de fevereiro de 2016

| Campo | Valor |
|---|---|
| Ementa (resumo) | Só a capitalização de lucros ou de reservas constituídas com lucros aumenta o custo de aquisição da participação (ações bonificadas); capitalização de reserva de capital não aumenta |
| Fonte oficial | https://normas.receita.fazenda.gov.br/sijut2consulta/link.action?idAto=72020 |
| Texto lido | https://normasinternet2.receita.fazenda.gov.br/api/consulta-externa/ato/72020/visao/multivigente e anexo 39279 (íntegra) |
| Lido em | 2026-11-10 |
| Leitura | íntegra |
| Confiança | alta |
| Temas | eventos societários; bonificação; custo |

## Vigência e alterações

| Dispositivo | Redação | Vigente de | até | Alterado por |
|---|---|---|---|---|
| SC inteira | original | 2016 (DOU); SIJUT: vigente | — | não reformada |

## Dispositivos relevantes para o COIR

- **§§17-19** — Lei 7.713 art. 16 (§4º: custo zero de participações por incorporação de lucros e reservas, salvo §3º); a Lei 9.249 art. 10 par. único (hoje §1º, Lei 12.973/2014) dá às ações bonificadas o custo igual à parcela de lucro ou reserva **de lucros** capitalizada que cabe ao acionista.
- **§§21-22** — reserva de capital capitalizada não gera custo (redução de base de cálculo só por lei; CF art. 150 §6º; CTN arts. 97 e 111); a IN 84 art. 16 §2º e a IN 1.500 art. 8º, ao falar em "reservas", referem-se só às de lucros.

## Onde o COIR usa

- `docs/regras-fiscais.md` R8, `coir/servicos/eventos.py`, `tests/fiscal_cases/test_casos_fiscais.py` (Lei 9.249 art. 10 par. único): bonificação com o custo atribuído do aviso aos acionistas; bonificação sem custo (custo zero) aceita na tela de Eventos: **confere**.

## Pontos em aberto

- Bonificação feita com reserva de capital: custo zero, mesmo que o aviso informe um valor. Vale um aviso na tela quando a companhia informar a origem da reserva.
