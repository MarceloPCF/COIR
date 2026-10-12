# COIR v3 — Controle de Operações e Imposto de Renda (B3)

Sistema open source (MIT), **local e offline**, que importa notas de corretagem SINACOR e planilhas da Área do
Investidor B3, calcula preço médio, apura o IR mensal de renda variável, gera DARF e prepara as fichas do IRPF.

> **Status:** em desenvolvimento (fase F0 — fundação). Ainda não há versão utilizável da v3.
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

## Licença

MIT — ver [LICENSE](LICENSE).
