"""Aviso da Lei 15.270/2025 para bonificação em ações (decisão do usuário de 10/11/2026, leitura
conservadora): desde 2026, a bonificação feita com lucros é emprego de lucros e conta no limite
mensal de R$ 50 mil de dividendos da mesma empresa (P&R Altas Rendas, perguntas 12 a 14). O COIR
não sabe a origem da bonificação (lucros ou reservas de capital), então só avisa."""

from datetime import date

from coir.domain.fiscal import TipoEvento
from tests.fiscal_cases.apoio import Livro


def test_bonificacao_desde_2026_avisa_e_antes_nao() -> None:
    """C 100 AAAA3 a 10,00; bonificação de 20 ações com valor atribuído de 100,00 em 01/03/2025
    (sem aviso: antes da lei) e outra de 10 ações com 60.000,00 em 02/03/2026 (aviso)."""
    livro = Livro()
    livro.nota(date(2025, 1, 6), livro.op("C", "AAAA3", 100, "10"))
    livro.evento(date(2025, 3, 3), "AAAA3", TipoEvento.BONIFICACAO, "C", "20", "100")
    livro.evento(date(2026, 3, 2), "AAAA3", TipoEvento.BONIFICACAO, "C", "10", "60000")
    ap = livro.apurar(date(2026, 3, 31))
    [aviso] = [a for a in ap.avisos if "bonificação" in a]
    assert aviso.startswith("03/2026: bonificação de AAAA3 com valor atribuído de R$ 60.000,00")
    assert "Lei 15.270/2025" in aviso
