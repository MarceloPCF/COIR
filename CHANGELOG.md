# Histórico de versões

Este projeto segue o [Versionamento Semântico](https://semver.org/lang/pt-BR/) (`MAIOR.MENOR.CORREÇÃO`):

* **MAIOR** – mudança que quebra a compatibilidade (ex.: novo formato da planilha de resultado);
* **MENOR** – funcionalidade nova que continua compatível;
* **CORREÇÃO** – correção de erro, sem funcionalidade nova.

O formato deste arquivo segue o [Keep a Changelog](https://keepachangelog.com/pt-BR/).

## [2.1.0] - 2026-10-08

Primeira versão do repositório reorganizado. O histórico anterior foi reiniciado (veja "Removido").

### Alterado
- Pastas reorganizadas: código em `coir/` (antes `Utils/`), planilha-modelo em `modelos/`, tabelas auxiliares em `dados/` (antes `Utils/Tickets/`) e capturas de tela em `docs/imagens/` (antes `Utils/Screenshots/`).
- `COIR.py` refatorado: despacho por corretora em uma tabela, validações em funções auxiliares e exigência explícita de Python 3.9.2 ou superior.
- A planilha de resultado passa a ser gravada uma única vez, ao final do processamento de todas as notas: os extratores das corretoras agora devolvem os dados em vez de gravar a cada nota.
- Ajustes nas áreas de leitura dos PDFs das corretoras BTG e XP/Rico/Clear e nas funções comuns (`coir/funcoes.py`).
- `requirements.txt` convertido para UTF-8 (as versões das bibliotecas não foram alteradas).

### Adicionado
- Leitura da conta do cliente por nota, quando um mesmo PDF reúne notas de contas diferentes.
- Correções manuais opcionais em `dados/ajustes_manuais.csv` (mudança de ticker, conta, corretora, quantidade ou preço); modelo em `dados/ajustes_manuais.modelo.csv`.
- Novos tickers em `dados/acoes.csv` (Lojas Renner, Moura Dubeux, PRIO, Porto Seguro).
- Número da versão exibido na abertura do programa (`VERSION` e `coir/__init__.py`).
- `.gitignore` e `.gitattributes`.
- Verificação contra dados pessoais (`tools/verificar_dados_pessoais.py`) e hook de pre-commit em `.githooks/`.
- Este `CHANGELOG.md`.

### Removido
- Executáveis e pacotes `.zip` das versões antigas, que existiam no histórico anterior do repositório.
- Link para um `style.css` inexistente em `index.html`.

## Versões anteriores (histórico reiniciado)

Até 2.0.4 o repositório usava tags fora do padrão e sem Releases. Elas foram removidas quando o histórico foi reiniciado. A tabela abaixo é apenas um registro das datas; os detalhes técnicos de cada uma não foram reconstruídos.

| Tag antiga | Data | Equivale a |
|---|---|---|
| v1.0.0 | 2024-01-05 | 1.0.0 |
| V1.0.1 | 2024-01-07 | 1.0.1 |
| v1.0.2 | 2024-02-08 | 1.0.2 |
| v1.0.3 | 2024-02-13 | 1.0.3 |
| v1.0.4 | 2024-09-19 | 1.0.4 |
| v.2.0 | 2024-11-13 | 2.0.0 (código dividido em pacote com um módulo por corretora) |
| v.2.01 | 2024-11-29 | 2.0.1 (página do projeto no GitHub Pages) |
| v.2.02 | 2024-12-25 | 2.0.2 |
| v2.03 | 2025-04-05 | 2.0.3 |
| v2.04 | 2025-04-05 | 2.0.4 |

Entre abril de 2025 e junho de 2026 houve atualizações sem tag (`xp_rico_clear.py`, `acoes.csv`, `opcoes.csv` e a planilha-modelo); elas estão incorporadas à 2.1.0.

[2.1.0]: https://github.com/MarceloPCF/COIR/releases/tag/v2.1.0
