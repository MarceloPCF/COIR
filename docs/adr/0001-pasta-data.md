# ADR 0001 — Tabelas de referência em `data/` na raiz

- Status: aceito (F0, 05/10/2026)

## Contexto
O CLAUDE.md define `data/` (ativos, aliases, corretoras, contratos, feriados, parâmetros fiscais com vigência)
na raiz do repositório. Um wheel gerado pelo `uv_build` só leva o pacote `coir/`, então `data/` não iria junto.
A distribuição da v3 é o executável PyInstaller em modo pasta (F9), não um pacote no PyPI.

## Decisão
- `data/` continua na raiz: é público, versionado e revisado como dado, separado do código.
- O código encontra a pasta só por `coir.recursos.pasta_data()`:
  em desenvolvimento, `<repo>/data`; no executável, `sys._MEIPASS/data`.
- O spec do PyInstaller (F9) inclui `datas=[("data", "data")]`.
- Pasta ausente é erro explícito (`ConfiguracaoError`), nunca tabela vazia.

## Consequências
- Instalar via `pip install coir` não traz as tabelas. Se um dia houver publicação no PyPI, mover para
  `coir/data/` (recurso do pacote via `importlib.resources`) e atualizar esta ADR.
- O motor (`coir/engine`) não lê arquivos: recebe as tabelas já carregadas como parâmetro.
