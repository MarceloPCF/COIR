# Perguntas e Respostas — Tributação de Altas Rendas: lucros e dividendos (RFB, versão de 19/12/2025)

| Campo | Valor |
|---|---|
| Ementa (resumo) | Orientação da RFB sobre o IRRF de lucros e dividendos da Lei 15.270/2025 (residentes e não residentes) |
| Fonte oficial | https://www.gov.br/receitafederal/pt-br/centrais-de-conteudo/publicacoes/perguntas-e-respostas/dirf/manual_padrao_rfb_per_tributacao_cotin_v-19-12-2025.pdf |
| Texto lido | o mesmo PDF, convertido em texto |
| Lido em | 2026-11-10 |
| Leitura | íntegra (30 perguntas) |
| Confiança | alta quanto ao que a RFB diz; é orientação, não lei |
| Temas | proventos/JCP; eventos societários (bonificação); IRRF |

## Vigência e alterações

| Dispositivo | Redação | Vigente de | até | Alterado por |
|---|---|---|---|---|
| perguntas 1 a 30 | versão inicial | 2025-12-19 | — | o documento diz que pode ser atualizado |

## Dispositivos relevantes para o COIR

- **P3** — IRRF de dividendos desde janeiro de 2026; IR mínimo desde o ano-calendário 2026.
- **P4** — o IRRF de 10% é antecipação: deduzido no IR mínimo; se a renda total não passar de R$ 600 mil, é restituível.
- **P5 e P7** — 10% sobre o total quando passar de R$ 50 mil no mês, da mesma PJ para a mesma PF; até R$ 50 mil, sem retenção.
- **P8 e P9** — lucros até 2025 só escapam se a distribuição foi aprovada pelo órgão competente até 31/12/2025 (proposta da administração não basta; dividendo intermediário pela administração se o estatuto permitir, Lei 6.404 art. 204) e pagos nos termos aprovados até 2028.
- **P12 a P14** — capitalização de lucros (bonificação) é "emprego": desde 2026 entra no limite de R$ 50 mil e no IR mínimo; o valor capitalizado soma ao custo das ações na DAA; a isenção do art. 10 da Lei 9.249 continua no ajuste normal. Capitalização de lucros até 2025 aprovada até 31/12/2025 fica fora.
- **P14 (final)** — ganhos em bolsa entram na base do IR mínimo; ganhos de capital fora de bolsa não.

## Onde o COIR usa

- Não é citado no projeto. Base para o aviso da Lei 15.270 (`coir/engine/proventos.py`) e para a bonificação.

## Pontos em aberto

- O aviso dos R$ 50 mil do COIR não soma bonificações (P12). Ver `lei-15270-2025.md`.
- Erro material do próprio documento: a P14 cita "Lei nº 15.270, de 2024".
- A edição 2027 do P&R IRPF deve dizer em que ficha vai o IRRF de 10%; revisar `data/irpf/2027.yaml` quando sair.
