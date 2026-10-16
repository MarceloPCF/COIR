# COIR v3 — Controle de Operações e Imposto de Renda (B3)

Sistema open source (AGPL v3), **local e offline**, que importa notas de corretagem SINACOR e planilhas da Área do
Investidor B3, calcula preço médio, apura o IR mensal de renda variável, gera DARF e prepara as fichas do IRPF.

> **Status:** em desenvolvimento. Prontos pela linha de comando: importação das notas, motor fiscal
> (PM, apuração, DARF) e exportação dos relatórios (fases F0 a F5). A interface gráfica vem na F6.
> A versão atual para uso é a **v2.1.1** (baseada em Excel), na tag [`v2.1.1`](../../tree/v2.1.1) e na
> [página de Releases](../../releases). O código dela está também em [`legacy/`](legacy/), como referência.

## Desenvolvimento

Requer [uv](https://docs.astral.sh/uv/) (ele instala o Python 3.13).

```bash
uv sync
uv run pre-commit install
uv run pytest
uv run coir --help
```

Os dados reais do usuário (notas em PDF, planilhas, banco) ficam **fora** do repositório, na pasta indicada pela
variável de ambiente `COIR_DADOS` (padrão: `~/COIR-dados`). Nunca versione PDFs, planilhas, `.db` ou CPFs —
o hook do pre-commit e o CI bloqueiam.

Regras fiscais: [`docs/regras-fiscais.md`](docs/regras-fiscais.md). Decisões de arquitetura: [`docs/adr/`](docs/adr/).
Notas de desenvolvimento com o Claude Code: [`docs/desenvolvimento/`](docs/desenvolvimento/).

## Uso pela linha de comando

```bash
uv run coir titular adicionar <CPF> --nome "Nome" # titular (as notas também o criam)
uv run coir importar                              # notas em PDF da pasta de dados
uv run coir eventos importar                      # eventos corporativos (eventos.csv)
uv run coir apurar 2025-03                        # mês com a memória de cálculo
uv run coir carteira 2025-12-31                   # posições e PM fiscal
uv run coir exportar --ano 2025                   # relatórios em xlsx, csv e pdf
```

`coir exportar` grava em `COIR_DADOS/relatorios/<ano>/` (ou em `--destino`, fora do repositório):
resumo fiscal, DARFs, renda variável na ordem do programa IRPF, IRRF e tributação exclusiva, operações
finalizadas, evolução, prejuízos a compensar, conferência do custo médio, carteira, bens e direitos e
livro de notas. Nada disso precisa do Excel instalado (ver [ADR 0003](docs/adr/0003-relatorios-xlsx-csv-pdf.md)).

## Licença

A v3 é distribuída sob a **GNU Affero General Public License v3.0 ou posterior** (AGPL-3.0-or-later) —
ver [LICENSE](LICENSE). Em resumo: qualquer pessoa pode usar, estudar, modificar e redistribuir o COIR de
graça; quem distribuir uma versão modificada, ou oferecê-la como serviço pela internet, precisa publicar o
código-fonte dessas modificações sob a mesma licença.

O código da v2 em [`legacy/`](legacy/) continua sob a licença MIT ([legacy/LICENSE](legacy/LICENSE)).
Decisão registrada em [docs/adr/0002-licenca-agpl.md](docs/adr/0002-licenca-agpl.md).
