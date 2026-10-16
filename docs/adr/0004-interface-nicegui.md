# ADR 0004 — Interface em NiceGUI com janela própria (pywebview/WebView2)

- Status: aceito (F6, 16/10/2026)

## Contexto
A F6 leva o ciclo do mês para fora do terminal. O stack já previa NiceGUI (plano, seção 3), com
estes requisitos:
- só Python: um mantenedor, sem build de Node;
- uso local e offline: nada da carteira sai do computador;
- abrir como aplicativo no Windows, sem instalar nada além do .zip (F9);
- tabelas densas e gráficos;
- licença compatível com AGPL-3.0;
- testes automáticos sem navegador.

## Decisão
- **NiceGUI 3** (MIT), sobre Quasar/Vue.
  - Tabelas: AG Grid Community (MIT).
  - Gráficos: ECharts (Apache-2.0).
  - Fontes e ícones servidos localmente.
- **Janela:**
  - `pywebview` (BSD) com o WebView2 do Windows (`ui.run(native=True)`);
  - sem WebView2 (verificado no registro), o COIR abre no navegador padrão;
  - o servidor escuta só em `127.0.0.1`, numa porta livre, sem recarga automática.
- **Camada:**
  - `coir/ui` só apresenta: importa `servicos`, `reports`, `domain`, `config` e `erros`, nunca
    `engine` nem `db`;
  - `servicos/estado.py` guarda o banco, as referências e o cache dos cálculos, que é refeito
    quando o arquivo do banco muda;
  - `nicegui` e `webview` são proibidos fora de `ui` (`tests/unit/test_camadas.py`).
- **Tema:**
  - um arquivo (`coir/ui/tema.py`) com as paletas clara e escura e as fontes do próprio Windows:
    Bahnschrift nos títulos e números, Segoe UI no texto, algarismos tabulares;
  - nenhuma fonte da web.
- **Testes:** `nicegui.testing.user_simulation`, sem navegador, com `pytest-asyncio`. O banco de
  demonstração (`tools/banco_demo.py`) usa as notas anonimizadas das fixtures.

## Alternativas
- **Flet** (Flutter): janela nativa sem WebView, mas sem uma grade de dados do porte do AG Grid e
  com um empacotamento menos maduro.
- **PySide6** (LGPL): nativo e maduro, mas com mais código de interface. Tabelas e gráficos
  ficariam por nossa conta.
- **Front-end em React/TypeScript:** duas stacks e build de Node para um mantenedor (plano,
  seção 3).

## Consequências
- Dependências novas: `nicegui` e `pywebview` (e, nos testes, `pytest-asyncio`). O pacote da F9
  fica maior, por causa dos arquivos de front-end do NiceGUI.
- O Windows 10/11 atual já traz o WebView2. Máquinas antigas sem ele caem no navegador, com o
  mesmo app.
