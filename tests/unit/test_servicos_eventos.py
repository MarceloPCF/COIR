"""F7a, tela 7: eventos lançados pela tela (conferidos pelo motor antes de gravar) e propostas
vindas da movimentação da B3, sobre o banco de demonstração (sem dados reais)."""

from dataclasses import replace
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path

import pytest

from coir import referencias
from coir.domain.fiscal import NaturezaEvento, TipoAtivo, TipoEvento, Visao
from coir.erros import ImportacaoError, MotorError
from coir.servicos import eventos
from coir.servicos.apuracao import carteira_banco
from coir.servicos.b3 import importar_planilhas
from coir.servicos.eventos import DadosEvento
from tests.unit.test_importador_b3 import CAB_MOV, planilha
from tools.banco_demo import criar

AGORA = datetime(2026, 10, 19, 12, 0, 0)
REFS, FISCAIS = referencias.carregar(), referencias.carregar_fiscais()
AGORA_B3 = "AGORA CORRETORA DE TITULOS E VALORES MOBILIARIOS S/A"
AGORA_CONTA = 1  # GGBR4: 300 ações (custo 18.687) de 2022 em diante
D = Decimal


@pytest.fixture
def demo(tmp_path: Path) -> Path:
    return criar(tmp_path / "demo")


def _carteira(banco: Path, dia: date) -> dict[str, tuple[Decimal, Decimal]]:
    pos, _ = carteira_banco(banco, REFS, FISCAIS, dia, Visao.TITULAR)
    return {p.ticker: (p.quantidade, p.custo.quantize(D(1))) for p in pos}


def _incorporacao(quantidade: str = "300") -> DadosEvento:
    return DadosEvento(
        tipo=TipoEvento.INCORPORACAO,
        data=date(2023, 3, 1),
        conta=AGORA_CONTA,
        ticker="ggbr4",
        quantidade=D(quantidade),
        ticker_destino="GOAU4",
        quantidade_destino=D(150),
        natureza=NaturezaEvento.INCORPORACAO_SOCIEDADE,
    )


def test_incorporacao_pela_tela_cadastra_o_ativo_recebido_e_leva_o_custo(demo: Path) -> None:
    """GOAU4 nunca foi negociado: o evento o cadastra (o motor recebe B com o custo de A)."""
    assert _carteira(demo, date(2023, 3, 1))["GGBR4"] == (300, D(18687))
    evento_id = eventos.salvar(demo, 1, _incorporacao(), REFS, FISCAIS)
    posicoes = _carteira(demo, date(2023, 3, 1))
    assert "GGBR4" not in posicoes and posicoes["GOAU4"] == (150, D(18687))
    [lancado] = [e for e in eventos.listar(demo, 1) if e.id == evento_id]
    assert (lancado.dados.ticker, lancado.dados.ticker_destino, lancado.conta_nome) == (
        "GGBR4",
        "GOAU4",
        "AGORA 0",
    )
    with pytest.raises(ImportacaoError, match="igual já lançado"):
        eventos.salvar(demo, 1, _incorporacao(), REFS, FISCAIS)


def test_evento_que_nao_fecha_com_as_operacoes_nao_e_gravado(demo: Path) -> None:
    antes = eventos.listar(demo, 1)
    with pytest.raises(ImportacaoError, match=r"não fecha com as operações.*posição 300"):
        eventos.salvar(demo, 1, _incorporacao("400"), REFS, FISCAIS)
    assert eventos.listar(demo, 1) == antes


def test_cisao_sem_percentual_nao_e_aceita(demo: Path) -> None:
    """Decisão do usuário (P4, 19/10/2026): sem o percentual divulgado, bloqueia."""
    d = DadosEvento(tipo=TipoEvento.CISAO, data=date(2023, 3, 1), conta=AGORA_CONTA,
                    ticker="GGBR4", quantidade=D(300), ticker_destino="GOAU4",
                    quantidade_destino=D(30))  # fmt: skip
    with pytest.raises(ImportacaoError, match="percentual do custo"):
        eventos.salvar(demo, 1, d, REFS, FISCAIS)
    com = eventos.salvar(demo, 1, DadosEvento(**{**d.__dict__, "percentual_custo": D("0.25")}),
                         REFS, FISCAIS)  # fmt: skip
    assert com > 0
    posicoes = _carteira(demo, date(2023, 3, 1))
    assert posicoes["GGBR4"] == (300, D(14015)) and posicoes["GOAU4"] == (30, D(4672))


def test_editar_e_excluir(demo: Path) -> None:
    evento_id = eventos.salvar(demo, 1, _incorporacao(), REFS, FISCAIS)
    eventos.salvar(demo, 1, DadosEvento(**{**_incorporacao().__dict__, "quantidade": D(100),
                                           "quantidade_destino": D(50)}),
                   REFS, FISCAIS, evento_id)  # fmt: skip
    assert _carteira(demo, date(2023, 3, 1))["GGBR4"] == (200, D(12458))
    eventos.excluir(demo, 1, evento_id, REFS, FISCAIS)
    assert _carteira(demo, date(2023, 3, 1))["GGBR4"] == (300, D(18687))
    with pytest.raises(ImportacaoError, match="não encontrado"):
        eventos.excluir(demo, 1, evento_id, REFS, FISCAIS)


def test_propostas_da_movimentacao_lancadas_pendentes_e_ignoradas(demo: Path) -> None:
    arquivo = planilha(demo.parent / "movimentacao-2023.xlsx", {"Movimentação": [
        CAB_MOV,
        ["Credito", "03/03/2023", "Desdobro", "GGBR4 - GERDAU S.A.", AGORA_B3, 300, "-", "-"],
        ["Credito", "10/04/2023", "Bonificação em Ativos", "GGBR4 - GERDAU S.A.", AGORA_B3, 30,
         "-", "-"],
        ["Credito", "10/04/2023", "Dividendo", "GGBR4 - GERDAU S.A.", AGORA_B3, 300, 0.5, 150],
    ]})  # fmt: skip
    importar_planilhas(demo, [arquivo], 1, AGORA)
    desdobro, bonificacao = sorted(eventos.propostas(demo, 1, REFS), key=lambda p: p.data)
    assert [p.situacao for p in (desdobro, bonificacao)] == ["pendente", "pendente"]
    assert (
        desdobro.dados.tipo,
        desdobro.dados.cv,
        desdobro.dados.quantidade,
        desdobro.dados.conta,
    ) == (TipoEvento.DESDOBRAMENTO, "C", 300, AGORA_CONTA)
    assert "data ex" in desdobro.nota and "custo atribuído" in bonificacao.nota
    # lançado com a data ex (dois dias antes do crédito): a proposta passa a "lançado"
    eventos.salvar(
        demo, 1, DadosEvento(**{**desdobro.dados.__dict__, "data": date(2023, 3, 1)}), REFS, FISCAIS
    )
    eventos.ignorar(demo, 1, bonificacao.chave, AGORA)
    situacoes = {p.movimentacao: p.situacao for p in eventos.propostas(demo, 1, REFS)}
    assert situacoes == {"Desdobro": "lançado", "Bonificação em Ativos": "ignorada"}
    eventos.ignorar(demo, 1, bonificacao.chave, AGORA, sim=False)
    assert {p.situacao for p in eventos.propostas(demo, 1, REFS)} == {"lançado", "pendente"}
    assert _carteira(demo, date(2023, 3, 1))["GGBR4"] == (600, D(18687))


def test_csv_com_incorporacao_e_codigo_antigo_do_ativo(demo: Path) -> None:
    """Colunas opcionais do CSV (ticker_destino, quantidade_destino, natureza); linhas antigas
    sem elas continuam com a mesma chave."""
    from coir.db.sessao import conectar

    linha = {"data": "2023-03-01", "corretora": "AGORA", "conta": "0", "ticker": "GGBR4",
             "tipo": "INCORPORACAO", "cv": "", "quantidade": "300", "valor": "",
             "corretora_destino": "", "conta_destino": "", "observacao": "", "origem": "",
             "ticker_destino": "GOAU4", "quantidade_destino": "150",
             "natureza": "incorporacao_sociedade", "percentual_custo": ""}  # fmt: skip
    sem_extras = {k: v for k, v in linha.items() if k not in eventos.EXTRAS}
    assert eventos._chave(sem_extras, 2) == eventos._chave(  # pyright: ignore[reportPrivateUsage]
        {**sem_extras, "natureza": ""}, 2
    )
    with conectar(demo) as fabrica, fabrica() as s:
        rel = eventos.importar_eventos(s, [linha])
    assert rel.novos == 1
    assert _carteira(demo, date(2023, 3, 1))["GOAU4"] == (150, D(18687))


def test_serie_convertida_e_um_ciclo_so_e_o_custo_levado_nao_e_premio() -> None:
    """Tela 8: XXXXA180 -> XXXXA17 (ajuste de série da B3) é o mesmo ciclo; prêmio = 800 - 502
    (sem contar o custo levado pelo evento como compra e venda de prêmio)."""
    from coir.servicos.derivativos import ciclos_opcoes
    from tests.fiscal_cases.apoio import Livro

    livro = Livro()
    venc = date(2025, 1, 17)
    livro.nota(date(2024, 12, 2), livro.op("C", "XXXXA180", 1000, "0.50", mercado="CALL",
                                           vencimento=venc), custos="2.00")  # fmt: skip
    livro.evento(date(2024, 12, 16), "XXXXA180", TipoEvento.INCORPORACAO, quantidade="1000",
                 para="XXXXA17", quantidade_destino="1000", natureza=NaturezaEvento.AJUSTE_SERIE,
                 tipo_ativo=TipoAtivo.OPCAO)  # fmt: skip
    livro.nota(date(2025, 1, 10), livro.op("V", "XXXXA17", 1000, "0.80", mercado="CALL",
                                           vencimento=venc), custos="1.50")  # fmt: skip
    fim = date(2025, 1, 31)
    [ciclo] = ciclos_opcoes(livro.movimentos(fim), livro.carteira(fim), 1, {}, REFS)
    assert (ciclo.ticker, ciclo.abertura, ciclo.situacao) == ("XXXXA17", date(2024, 12, 2),
                                                               "encerrada")  # fmt: skip
    assert (ciclo.premio, ciclo.resultado) == (D("296.50"), D("296.50"))


def _desdobro(**troca: object) -> DadosEvento:
    base = DadosEvento(tipo=TipoEvento.DESDOBRAMENTO, data=date(2023, 3, 1), conta=AGORA_CONTA,
                       ticker="GGBR4", cv="C", quantidade=D(300))  # fmt: skip
    return DadosEvento(**{**base.__dict__, **troca})


def test_cisao_em_duas_contas_so_grava_em_lote(demo: Path) -> None:
    """Revisão F7a A1: a cisão de uma posição em duas contas fecha só com os dois eventos."""
    transferencia = DadosEvento(tipo=TipoEvento.TRANSFERENCIA, data=date(2023, 2, 1),
                                conta=AGORA_CONTA, ticker="GGBR4", quantidade=D(100),
                                conta_destino=2)  # fmt: skip
    eventos.salvar(demo, 1, transferencia, REFS, FISCAIS)
    cisao = DadosEvento(tipo=TipoEvento.CISAO, data=date(2023, 3, 1), conta=AGORA_CONTA,
                        ticker="GGBR4", quantidade=D(200), ticker_destino="GOAU4",
                        quantidade_destino=D(20), percentual_custo=D("0.25"))  # fmt: skip
    with pytest.raises(ImportacaoError, match="lance o evento de cada conta"):
        eventos.salvar(demo, 1, cisao, REFS, FISCAIS)
    outra = DadosEvento(**{**cisao.__dict__, "conta": 2, "quantidade": D(100),
                           "quantidade_destino": D(10)})  # fmt: skip
    assert len(eventos.salvar_lote(demo, 1, [(cisao, None), (outra, None)], REFS, FISCAIS)) == 2
    assert _carteira(demo, date(2023, 3, 1))["GOAU4"] == (30, D(4672))


def test_excluir_evento_de_que_outro_depende_e_recusado(demo: Path) -> None:
    desdobro = eventos.salvar(demo, 1, _desdobro(), REFS, FISCAIS)
    eventos.salvar(demo, 1, _incorporacao("600"), REFS, FISCAIS)
    with pytest.raises(ImportacaoError, match="sem este evento o livro não fecha"):
        eventos.excluir(demo, 1, desdobro, REFS, FISCAIS)
    assert len(eventos.listar(demo, 1)) == 2


def test_mesmo_evento_na_tela_e_no_csv_nao_duplica(demo: Path) -> None:
    """Revisão F7a M5/M6: valor 0 na tela = valor vazio no CSV; --substituir não apaga o que
    foi lançado na tela."""
    from coir.db.sessao import conectar

    eventos.salvar(demo, 1, _desdobro(), REFS, FISCAIS)
    linha = {"data": "2023-03-01", "corretora": "AGORA", "conta": "0", "ticker": "GGBR4",
             "tipo": "DESDOBRAMENTO", "cv": "C", "quantidade": "300", "valor": "",
             "corretora_destino": "", "conta_destino": "", "observacao": "",
             "origem": ""}  # fmt: skip
    with conectar(demo) as fabrica, fabrica() as s:
        rel = eventos.importar_eventos(s, [linha])
    assert (rel.novos, rel.ja_no_banco) == (0, 1)
    with conectar(demo) as fabrica, fabrica() as s:
        rel = eventos.importar_eventos(s, [], substituir=True)
    assert rel.removidos == 0 and len(eventos.listar(demo, 1)) == 1


def _pacote_eventos(tmp: Path, linhas: list[str]) -> Path:
    """data/ com eventos e instrumentos de teste (ISINs fictícios) no lugar dos da B3."""
    import shutil

    from coir.recursos import pasta_data

    pasta = tmp / "pacote"
    shutil.copytree(pasta_data(), pasta)
    (pasta / "instrumentos_b3.csv").write_text(
        "ticker,isin,categoria,especificacao,emissor,nome,visto_desde,visto_ate\n"
        "GGBR4,BRGGBRTESTE1,SHARES,PN N1,GGBR,GERDAU,2026-10-20,2026-10-20\n"
        "GOAU4,BRGOAUTESTE1,SHARES,PN N1,GOAU,METALURGICA GERDAU,2026-10-20,2026-10-20\n"
        "IRBR3,BRIRBRTESTE1,SHARES,ON NM,IRBR,IRB,2026-10-20,2026-10-20\n",
        encoding="utf-8",
    )  # fmt: skip
    (pasta / "eventos_b3.csv").write_text(
        "emissor,isin,tipo,fator,data_com,aprovado_em,isin_emitido,preco,periodo,data_subscricao\n"
        + "".join(f"{x}\n" for x in linhas), encoding="utf-8")  # fmt: skip
    return pasta


def test_propostas_do_pacote_com_data_ex_e_quantidade_pela_posicao(demo: Path) -> None:
    """GGBR4: 300 na conta Ágora. Data com 28/02/2023 (terça) -> data ex 01/03/2023."""
    pasta = _pacote_eventos(demo.parent, [
        "GGBR,BRGGBRTESTE1,DESDOBRAMENTO,100,2023-02-28,2023-02-01,BRGGBRTESTE1,,,",
        "GGBR,BRGGBRTESTE1,BONIFICACAO,5,2023-05-31,2023-05-01,BRGGBRTESTE1,,,",
        "GGBR,BRGGBRTESTE1,GRUPAMENTO,0.33,2023-08-31,2023-08-01,BRGGBRTESTE1,,,",
        "GGBR,BRGGBRTESTE1,CIS RED CAP,28.29,2023-10-31,2023-10-01,BRGOAUTESTE1,,,",
        "GGBR,BRGGBRTESTE1,DESDOBRAMENTO,100,2008-01-31,2008-01-01,BRGGBRTESTE1,,,",  # sem posição
    ])  # fmt: skip
    lista = eventos.propostas(demo, 1, REFS, FISCAIS, pasta)
    do_pacote = {p.dados.tipo: p for p in lista if "eventos da B3" in p.movimentacao
                 and p.codigo == "GGBR4"}  # fmt: skip
    desdobro = do_pacote[TipoEvento.DESDOBRAMENTO]
    assert (desdobro.data, desdobro.dados.cv, desdobro.dados.quantidade, desdobro.dados.conta) == (
        date(2023, 3, 1),
        "C",
        300,
        AGORA_CONTA,
    )
    assert do_pacote[TipoEvento.BONIFICACAO].dados.quantidade == 15
    assert "custo atribuído" in do_pacote[TipoEvento.BONIFICACAO].nota
    grupamento = do_pacote[TipoEvento.GRUPAMENTO]  # 300 x 0,33 = 99: saem 201
    assert (grupamento.dados.cv, grupamento.dados.quantidade) == ("V", 201)
    cisao = do_pacote[TipoEvento.CISAO]  # 300 x 28,29% = 84,87: 84 GOAU4, sem percentual
    assert (cisao.dados.ticker_destino, cisao.dados.quantidade_destino) == ("GOAU4", 84)
    assert cisao.dados.percentual_custo is None and "percentual" in cisao.nota
    assert len([p for p in lista if p.codigo == "GGBR4"]) == 4  # o de 2008 não atinge a posição
    eventos.salvar(demo, 1, desdobro.dados, REFS, FISCAIS)
    [lancado] = [
        p
        for p in eventos.propostas(demo, 1, REFS, FISCAIS, pasta)
        if p.dados.tipo is TipoEvento.DESDOBRAMENTO and p.codigo == "GGBR4"
    ]
    assert lancado.situacao == "lançado"


def test_evento_do_objeto_propoe_ajuste_das_series_abertas(demo: Path) -> None:
    """Demonstração: IRBRE35 (600, vence em maio) aberta em fev/2023 na BTG; desdobramento 1:2 de
    IRBR3 -> a série recebe mais 600 (a B3 ajusta a quantidade)."""
    pasta = _pacote_eventos(demo.parent, [
        "IRBR,BRIRBRTESTE1,DESDOBRAMENTO,100,2023-02-28,2023-02-01,BRIRBRTESTE1,,,",
    ])  # fmt: skip
    series = {p.codigo: p for p in eventos.propostas(demo, 1, REFS, FISCAIS, pasta)
              if "eventos da B3" in p.movimentacao}  # fmt: skip
    assert "IRBRE35" in series
    ajuste = series["IRBRE35"]
    assert (ajuste.dados.tipo, ajuste.dados.cv, ajuste.dados.quantidade) == (
        TipoEvento.DESDOBRAMENTO,
        "C",
        600,
    )
    assert "código" in ajuste.nota


def test_ignorar_a_proposta_da_movimentacao_ignora_a_do_pacote(demo: Path) -> None:
    """Revisão F7b: o mesmo desdobramento vem da movimentação (crédito em 03/03/2023) e do pacote
    (data ex 01/03/2023); ignorado numa fonte, não volta pela outra."""
    arquivo = planilha(demo.parent / "movimentacao-2023.xlsx", {"Movimentação": [
        CAB_MOV,
        ["Credito", "03/03/2023", "Desdobro", "GGBR4 - GERDAU S.A.", AGORA_B3, 300, "-", "-"],
    ]})  # fmt: skip
    importar_planilhas(demo, [arquivo], 1, AGORA)
    pasta = _pacote_eventos(
        demo.parent, ["GGBR,BRGGBRTESTE1,DESDOBRAMENTO,100,2023-02-28,2023-02-01,BRGGBRTESTE1,,,"]
    )
    [da_movimentacao] = eventos.propostas(demo, 1, REFS)
    eventos.ignorar(demo, 1, da_movimentacao.chave, AGORA)
    [unica] = eventos.propostas(demo, 1, REFS, FISCAIS, pasta)
    assert ("eventos da B3" in unica.movimentacao, unica.situacao) == (True, "ignorada")


def test_custodia_inicial_de_ativo_sem_nota_cadastra_o_ativo(demo: Path) -> None:
    """Quem começa sem as notas antigas lança ativos que ainda não têm negócio no banco (rodada
    do beta, 26/10/2026); código fora do formato de negociação é recusado."""
    d = DadosEvento(tipo=TipoEvento.CUSTODIA_INICIAL, data=date(2023, 12, 31), conta=AGORA_CONTA,
                    ticker="itub4", cv="C", quantidade=D(100), valor=D(3000))  # fmt: skip
    eventos.salvar_lote(demo, 1, [(d, None)], REFS, FISCAIS)
    assert _carteira(demo, date(2024, 1, 31))["ITUB4"] == (D(100), D(3000))
    with pytest.raises(ImportacaoError, match="código de negociação inválido"):
        eventos.salvar(demo, 1, replace(d, ticker="ITAU"), REFS, FISCAIS)


# ----------------------------------------------- pendência no livro (teste com notas reais, 3.0.1)
def _pendencias(monkeypatch: pytest.MonkeyPatch, *erros: MotorError | None) -> None:
    """O motor para em `erros[0]` antes de gravar e em `erros[1]` depois."""
    fila = list(erros)
    monkeypatch.setattr(eventos, "_erro_do_livro", lambda *_: fila.pop(0))


ITSA = MotorError("exercício de 1050 ITSAL921 em 2023-12-15", date(2023, 12, 15))


@pytest.mark.parametrize(
    "depois",
    [
        MotorError(str(ITSA), ITSA.data),  # a pendência de antes continua (outro ativo)
        MotorError("exercício de 220 LRENL14 em 2024-12-20", date(2024, 12, 20)),  # resolveu ela
        None,
    ],
)
def test_pendencia_de_antes_nao_impede_gravar_outro_evento(
    demo: Path, monkeypatch: pytest.MonkeyPatch, depois: MotorError | None
) -> None:
    """Uma bonificação que falta trava o livro; o usuário lança os eventos um por vez."""
    _pendencias(monkeypatch, ITSA, depois)
    eventos.salvar(demo, 1, _desdobro(), REFS, FISCAIS)
    assert len(eventos.listar(demo, 1)) == 1


@pytest.mark.parametrize(
    ("antes", "depois"),
    [
        (None, ITSA),  # o livro fechava
        (ITSA, MotorError("outra pendência", date(2023, 3, 1))),  # mais cedo que a de antes
        (ITSA, MotorError("outra pendência", ITSA.data)),  # mesmo dia, outra pendência
    ],
)
def test_evento_que_cria_pendencia_nao_e_gravado(
    demo: Path, monkeypatch: pytest.MonkeyPatch, antes: MotorError | None, depois: MotorError
) -> None:
    _pendencias(monkeypatch, antes, depois)
    with pytest.raises(ImportacaoError, match="não fecha com as operações"):
        eventos.salvar(demo, 1, _desdobro(), REFS, FISCAIS)
    assert eventos.listar(demo, 1) == []


def test_propostas_do_pacote_com_o_livro_parado(
    demo: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """O livro para em 15/03/2023: os eventos até ali são propostos pela posição da véspera (é
    assim que se destrava o livro); os de depois esperam a pendência ser resolvida."""
    pasta = _pacote_eventos(demo.parent, [
        "GGBR,BRGGBRTESTE1,DESDOBRAMENTO,100,2023-02-28,2023-02-01,BRGGBRTESTE1,,,",
        "GGBR,BRGGBRTESTE1,BONIFICACAO,5,2023-05-31,2023-05-01,BRGGBRTESTE1,,,",
    ])  # fmt: skip
    real = eventos.movimentar

    def parado(*args: object, **kwargs: object) -> object:
        if args[3] == date.max:
            raise MotorError("exercício sem a opção", date(2023, 3, 15))
        return real(*args, **kwargs)  # type: ignore[arg-type]

    monkeypatch.setattr(eventos, "movimentar", parado)
    do_pacote = [p for p in eventos.propostas(demo, 1, REFS, FISCAIS, pasta)
                 if "eventos da B3" in p.movimentacao and p.codigo == "GGBR4"]  # fmt: skip
    assert [(p.dados.tipo, p.dados.quantidade) for p in do_pacote] == [
        (TipoEvento.DESDOBRAMENTO, 300)
    ]


def test_ativo_travado_nao_esconde_as_propostas_dos_outros(
    demo: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """Um ativo que o motor isola em 15/03/2023 só limita as propostas dele. Antes, o primeiro
    problema parava o livro, e nenhum evento posterior de nenhum ativo era proposto."""
    from coir.domain.fiscal import AtivoIsolado

    pasta = _pacote_eventos(demo.parent, [
        "GGBR,BRGGBRTESTE1,DESDOBRAMENTO,100,2023-02-28,2023-02-01,BRGGBRTESTE1,,,",
        "GGBR,BRGGBRTESTE1,BONIFICACAO,5,2023-05-31,2023-05-01,BRGGBRTESTE1,,,",
    ])  # fmt: skip
    real = eventos.movimentar

    def com_isolado(ativo: int) -> object:
        def movimentar(*args: object, isolados: list[AtivoIsolado] | None = None) -> object:
            if isolados is not None:
                isolados.append(AtivoIsolado(titular=1, ativo=ativo, ticker="TRAVA3",
                                             desde=date(2023, 3, 15), motivo="teste"))  # fmt: skip
            return real(*args, isolados=[])  # type: ignore[arg-type]

        return movimentar

    def propostas() -> list[TipoEvento]:
        return [p.dados.tipo for p in eventos.propostas(demo, 1, REFS, FISCAIS, pasta)
                if "eventos da B3" in p.movimentacao and p.codigo == "GGBR4"]  # fmt: skip

    monkeypatch.setattr(eventos, "movimentar", com_isolado(999999))  # outro ativo travou
    assert set(propostas()) == {TipoEvento.DESDOBRAMENTO, TipoEvento.BONIFICACAO}
    from sqlalchemy import select

    from coir.db.modelos import TickerHistorico
    from coir.db.sessao import conectar

    with conectar(demo) as fabrica, fabrica() as s:
        ggbr = s.scalars(select(TickerHistorico.ativo_id)
                         .where(TickerHistorico.ticker == "GGBR4")).first()  # fmt: skip
    assert ggbr is not None
    monkeypatch.setattr(eventos, "movimentar", com_isolado(ggbr))  # o próprio ativo travou
    assert propostas() == [TipoEvento.DESDOBRAMENTO]  # o de depois espera a pendência


def test_bonificacao_do_pacote_com_custo_atribuido(demo: Path) -> None:
    """O pacote traz o custo por ação quando o aviso aos acionistas informa (coluna
    custo_atribuido): a proposta sai com o custo total, senão o PM das vendas seguintes fica errado
    (IN RFB 1.585/2015 art. 58: a bonificação entra pelo custo atribuído pela companhia). GGBR4:
    300 ações, bonificação de 5% = 15 a R$ 11,54731631 = 173,2097... -> 173,21."""
    pasta = _pacote_eventos(demo.parent, [])
    (pasta / "eventos_b3.csv").write_text(
        "emissor,isin,tipo,fator,data_com,aprovado_em,isin_emitido,preco,periodo,data_subscricao,"
        "origem,referencia,custo_atribuido\n"
        "GGBR,BRGGBRTESTE1,BONIFICACAO,5,2023-05-31,2023-05-01,BRGGBRTESTE1,,,,cvm_ipe,,11.54731631\n",
        encoding="utf-8",
    )  # fmt: skip
    [bonus] = [p for p in eventos.propostas(demo, 1, REFS, FISCAIS, pasta)
               if "eventos da B3" in p.movimentacao and p.codigo == "GGBR4"]  # fmt: skip
    assert (bonus.dados.quantidade, bonus.dados.valor) == (15, Decimal("173.21"))
    assert "11.54731631 por ação" in bonus.nota and "informe" not in bonus.nota


def test_dinheiro_de_incorporacao_propoe_a_parte_do_custo_nao_o_dinheiro(demo: Path) -> None:
    """Molde de PETZ3 -> AUAU3 (parecer do especialista de 07/11/2026; SC Cosit 37/2020): o evento
    tira das ações recebidas a parte do custo da ação resgatada, em % do custo. GGBR4: 300 ações na
    data com; 10% do custo; o dinheiro é só informado, com o programa em que ele vai."""
    pasta = _pacote_eventos(demo.parent, [])
    (pasta / "eventos_b3.csv").write_text(
        "emissor,isin,tipo,fator,data_com,aprovado_em,isin_emitido,preco,periodo,data_subscricao,"
        "origem,referencia,custo_atribuido\n"
        "GGBR,BRGGBRTESTE1,REDUCAO_CAPITAL,10,2023-05-31,2023-05-01,BRGGBRTESTE1,0.50,,2023-06-20,"
        "cvm_ipe,fato relevante,\n",
        encoding="utf-8",
    )  # fmt: skip
    [resgate] = [p for p in eventos.propostas(demo, 1, REFS, FISCAIS, pasta)
                 if "eventos da B3" in p.movimentacao]  # fmt: skip
    assert resgate.dados.tipo is TipoEvento.REDUCAO_CAPITAL
    assert (resgate.dados.cv, resgate.data) == ("V", date(2023, 6, 1))
    assert resgate.dados.valor != Decimal("150.00")  # o dinheiro: 300 x 0,50
    assert "R$ 150.00" in resgate.nota and "GCAP" in resgate.nota and "20/06/2023" in resgate.nota
    assert resgate.situacao == "pendente"


def test_pendencia_na_visao_por_conta_nao_fica_escondida(
    demo: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """A visão por titular para no ITSA (dez/2023); a por conta, antes (evento na conta errada,
    mai/2021). Vale a mais antiga: senão o evento errado passa e o livro por conta, que propõe os
    eventos seguintes, trava sem o usuário saber por quê."""
    conta = MotorError("evento DESDOBRAMENTO de BIDI4 sem posição na conta", date(2021, 5, 27))

    def movimentar(*args: object) -> None:
        raise ITSA if args[-1] is eventos.Visao.TITULAR else conta

    from coir.db.sessao import conectar

    monkeypatch.setattr(eventos, "movimentar", movimentar)
    with conectar(demo) as fabrica, fabrica() as s:
        assert eventos._erro_do_livro(s, 1, REFS, FISCAIS) is conta
