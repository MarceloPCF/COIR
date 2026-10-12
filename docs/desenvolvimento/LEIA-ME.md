# Contexto para o Claude Code — COIR v3

> `CLAUDE.md` e `docs/PLANO.md` são documentos de trabalho do mantenedor e ficam só no PC dele
> (estão no `.gitignore`). Os agentes em `.claude/agents/` são públicos.

Conteúdo desta pasta (copie tudo para a raiz do novo repositório):

| Arquivo | Para quê |
|---|---|
| `CLAUDE.md` | Lido automaticamente pelo Claude Code: stack, estrutura, regras inegociáveis, vocabulário |
| `agentes/*.md` | 4 subagentes do projeto (fiscal, parser, revisor, validador). **Copie para `.claude\agents\`** no novo repositório (a ferramenta remota não pode gravar em pastas `.claude`). |
| `docs/PLANO.md` | Plano completo: escopo, diagnóstico, decisões, arquitetura, dados, telas, roadmap, riscos |
| `docs/regras-fiscais.md` | Regras R1–R18 com fonte e grau de confiança |
| `docs/benchmark.md` | Concorrentes e posicionamento |
| `PROMPTS-INICIAIS.md` | Prompts prontos, fase a fase |
| `.gitignore` | Bloqueia PDFs, planilhas, banco e dados pessoais |

Pasta do projeto: `E:\Claude-IA\COIR-v3.0`. Os agentes já estão em `.claude\agents\`.

Dados pessoais (notas, relatórios B3, gabarito) ficam FORA do projeto, em `E:\Claude-IA\COIR-dados`.
No Claude Code: `/add-dir E:\Claude-IA\COIR-dados`. Variável de ambiente: `COIR_DADOS=E:\Claude-IA\COIR-dados`.
