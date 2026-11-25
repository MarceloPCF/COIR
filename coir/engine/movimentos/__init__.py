# pyright: reportPrivateUsage=false, reportUnusedFunction=false
# (os nomes com _ são internos do pacote, usados entre os arquivos dele)
"""Livro de operações -> movimentos de posição, dia a dia (determinístico, sem relógio).

Ordem dentro do dia (por titular): eventos; liquidação de termo; day trade (casado pela ordem dos
negócios, art. 65 §3º; o excedente vira comum pelo preço do próprio negócio); compras;
exercícios; vendas; vencimento (pó).

Fontes:
- Day trade: IN RFB 1.585/2015 art. 65 (mesmo ativo, mesmo dia, mesma corretora; antes da
  MP 497/2010, a operação entre corretoras com liquidação física também não era day trade: IN SRF
  25/2001 art. 31 §6º); §13: exercício de opção não é day trade.
- Exercício (art. 60 II a-d e §§1º e 3º): titular de call soma o prêmio ao custo; lançador de
  call soma o prêmio à alienação; titular de put subtrai o prêmio da alienação; lançador de put
  subtrai do custo.
- Opção não exercida: o prêmio é perda (titular) ou ganho (lançador) no vencimento.
- Termo: as ações entram na carteira na liquidação, pelo preço a termo + custos da nota. Vendas
  à vista na mesma data consomem primeiro o lote do termo, pelo preço a termo (art. 62 I, linha do
  termo, sem isenção); o excedente sai da posição anterior pelo PM. O que sobra do lote fica no
  custo médio (art. 62 §1º e art. 58). Critérios do COIR de 20 e 21/10/2026.
- Data da liquidação do termo (compra e venda): a de `OperacaoMotor.liquidacao_termo`, que o
  serviço tira da movimentação da B3 quando ela mostra a liquidação antecipada (total ou parcial,
  uma operação por parte) e, sem isso, do vencimento do contrato, com aviso.
- Venda a termo: a alienação se dá na liquidação (art. 59 §2º IV), no vencimento ou antecipada. O
  ganho é o preço a termo, menos os custos da nota, menos o custo médio da posição na abertura
  desse dia (art. 62 II e III; art. 58): coberto e descoberto dão a mesma conta. Linha do termo,
  nunca isento; o valor bruto entra na soma do limite do mês da liquidação. Sem ações para
  entregar, o motor para. Parecer do especialista de 13/11/2026.
"""

from coir.engine.movimentos.percurso import carteira, movimentar
from coir.engine.movimentos.termo import termos_em_aberto

__all__ = ["carteira", "movimentar", "termos_em_aberto"]
