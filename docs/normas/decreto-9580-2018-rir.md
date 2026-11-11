# Decreto nº 9.580, de 22 de novembro de 2018 (Regulamento do Imposto sobre a Renda — RIR/2018)

| Campo | Valor |
|---|---|
| Ementa (resumo) | Regulamenta a tributação, a fiscalização, a arrecadação e a administração do imposto sobre a renda; consolida as leis, citando a lei-fonte de cada artigo |
| Fonte oficial | https://www.planalto.gov.br/ccivil_03/_ato2015-2018/2018/decreto/d9580.htm |
| Texto lido | https://www2.camara.leg.br/legin/fed/decret/2018/decreto-9580-22-novembro-2018-787360-normaatualizada-pe.html (anexo compilado) |
| Lido em | 2026-11-10 |
| Leitura | **parcial**: Livro III, Título III, Capítulos VIII a XI (arts. 837 a 852: swap, compensação, operações em bolsa, ouro, ganho líquido, à vista, custo, PME, opções, futuros, termo, IRRF 0,005%, day trade); Título IV (arts. 853 a 872: tratamento dos ganhos, perdas, dispensa, empréstimo de ativos); art. 35 IV e V (rendimentos isentos); arts. 128 a 141 percorridos (ganho de capital, isenção de pequeno valor art. 133, custo art. 136 a 141); art. 915 (prazo do DARF). O decreto tem mais de 1.000 artigos; os demais não tratam de renda variável de PF |
| Confiança | alta para o texto lido; um regulamento não cria regra, vale a lei que ele cita |
| Temas | ações; isenção 20 mil; day trade; opções; futuros; termo; BTC; proventos; IRRF; DARF |

DOU: 23/11/2018. Vigência na publicação; antes vigorava o RIR/1999 (Decreto 3.000/1999), com a mesma substância para renda variável.

## Vigência e alterações

| Dispositivo | Redação | Vigente de | até | Alterado por |
|---|---|---|---|---|
| arts. 839 a 851 (bolsa, ganho líquido, custo médio, PME, opções, futuros, termo, IRRF, day trade) | original | 23/11/2018 | — | sem alteração registrada no compilado |
| arts. 863 a 872 (empréstimo de ativos, reproduzindo a Lei 13.043 arts. 6º a 15) | original | 23/11/2018 | — | sem alteração no decreto, mas a base legal foi revogada em 01/01/2026 (Lei 15.265/2025 art. 40): desatualizados desde então |
| art. 915 (prazo do imposto mensal: último dia útil do mês seguinte) | original | 23/11/2018 | — | — |
| 2016 a 22/11/2018 | RIR/1999 (Decreto 3.000/1999), não lido nesta tarefa | — | 22/11/2018 | — |

## Dispositivos relevantes para o COIR

- **art. 35 V e e m** — isentos: ganhos líquidos de PF no à vista de ações até R$ 20 mil de alienações no mês (cita Lei 8.981 art. 72 §8º e Lei 11.033 art. 3º I); ganho de PME até 31/12/2023. **art. 35 V d** — rendimentos de FII (só a condição de negociação; os requisitos de cotistas estão na lei).
- **art. 133 I e art. 35 VI a** — pequeno valor fora de bolsa: R$ 20 mil (ações no balcão) e R$ 35 mil (demais).
- **art. 839** — 15% sobre ganhos líquidos em bolsas e assemelhadas; §6º manda pagar no prazo do art. 915.
- **art. 841** — ganho líquido é o resultado positivo das operações do mês, deduzidos custos e despesas necessários; §1º perdas compensam ganhos da mesma natureza nos meses seguintes (Lei 8.981 art. 72 §4º).
- **art. 842 e art. 843** — à vista: alienação menos custo; custo pela média ponderada; §1º bonificação por capitalização de lucros desde 1996 com custo igual à parcela capitalizada (Lei 9.249 art. 10 §1º); §2º custo na ausência de pagamento; §3º custo zero (partes beneficiárias gratuitas, ações acrescidas por desdobramento, bem sem valor determinável).
- **art. 845** — isenção de PME (Lei 13.043 arts. 16 e 17).
- **art. 846** — opções: ganho no encerramento da série; exercício pelos quatro casos (titular/lançador, compra/venda), com o prêmio somado ou deduzido; §1º sem venda no exercício, custo = preço de exercício ± prêmio; §2º média ponderada dos prêmios; §3º opção não encerrada nem exercida: prêmio é ganho do lançador e perda do titular no vencimento.
- **art. 847** — futuros: soma algébrica dos ajustes na liquidação, cessão ou encerramento (Lei 11.051 art. 32).
- **art. 848** — termo: comprador ganha a diferença entre a venda à vista na liquidação e o preço a termo; §1º sem venda na liquidação, custo = preço a termo.
- **art. 849** — IRRF de 0,005% e §6º dedução no mês, compensação nos meses seguintes, na declaração de ajuste e com ganho de capital em ações (Lei 11.033 art. 2º §§1º a 7º).
- **art. 851** — day trade: definição (mesma instituição), 1% na fonte, §4º dedução do imposto do mês, §5º restituição do saldo no fim do ano, §§6º e 7º compensação só entre day trades e alíquota de 20%, §8º II definitivo para PF.
- **art. 854 §2º II e art. 858 II** — para PF, rendimentos e ganhos líquidos são tributados de forma definitiva (não vão ao ajuste).
- **art. 855 §2º** — perdas de day trade só compensam ganhos de day trade.
- **arts. 863 a 872 (base legal revogada desde 2026)** — BTC: remuneração do doador como renda fixa; reembolso de proventos isento de IRRF; tomador apura na recompra, com corretagem e emolumentos como custo; devolução não é fato gerador.
- **art. 915** — o imposto apurado na forma do art. 839 (renda variável), do art. 122 (carnê-leão) e do art. 153 (ganho de capital) vence no último dia útil do mês seguinte ao da percepção.

## Onde o COIR usa

- `data/parametros_fiscais.csv` linha 2 (`aliquota_comum`: "RIR/2018 art. 839") — confere. Como o parâmetro começa em 2005, a fonte primária é a Lei 11.033; o RIR/2018 só vale de 23/11/2018 em diante.
- `docs/regras-fiscais.md` R12, `coir/engine/calendario.py:3`, `coir/engine/apuracao.py:14`, `coir/reports/fiscais.py:245`, `tests/fiscal_cases/test_casos_fiscais.py:96` ("RIR/2018 art. 915": último dia útil do mês seguinte) — confere. A lei citada pelo próprio art. 915 é a Lei 8.383/1991 art. 52 §§1º e 2º e a Lei 8.981 art. 21 §1º; para 2016 a 2018, valia o RIR/1999 com a mesma regra.
- `tests/fiscal_cases/test_casos_fiscais.py:408` ("RIR/2018; Perguntas e Respostas IRPF", sem artigo, para bonificação, desdobramento e redução de capital) — citação vaga: os artigos são o 843 §1º (bonificação) e o 843 §3º II (desdobramento); redução de capital com restituição não está nestes capítulos (IN RFB 1.585 art. 58 §8º).

## Pontos em aberto

- Os arts. 863 a 872 continuam no texto, mas a lei que reproduzem foi revogada em 01/01/2026; para BTC desde 2026, citar a Lei 15.265/2025 arts. 18 a 30, como o projeto já faz.
- O RIR/2018 cita a Lei 8.981 art. 72 §8º ao lado da Lei 11.033 art. 3º I para o limite de R$ 20 mil; isso não muda a crítica à citação do R5 (ver ficha da Lei 8.981): o §8º trata de isenção, não de compensação.
- Sugestão: trocar "RIR/2018" sem artigo no teste da linha 408 por "RIR/2018 art. 843 §§1º e 3º II".
