# Histórico de versões

Este projeto segue o [Versionamento Semântico](https://semver.org/lang/pt-BR/) (`MAIOR.MENOR.CORREÇÃO`):

* **MAIOR** – mudança que quebra a compatibilidade (ex.: novo formato da planilha de resultado);
* **MENOR** – funcionalidade nova que continua compatível;
* **CORREÇÃO** – correção de erro, sem funcionalidade nova.

O formato deste arquivo segue o [Keep a Changelog](https://keepachangelog.com/pt-BR/).

## [2.2.0] - 2026-10-21

Atualização para o Python 3.15 e para as versões atuais das bibliotecas (pandas 3, numpy 2.5, tabula-py 2.10, xlwings 0.37, openpyxl 3.1). Compatibilidade com o pandas atual ([#11](https://github.com/MarceloPCF/COIR/issues/11)), com base na contribuição de [@tommmdl](https://github.com/tommmdl) ([#34](https://github.com/MarceloPCF/COIR/pull/34)).

Os resultados foram comparados com os da versão 2.1.1 em todas as notas de teste (XP, BTG e Rico), valor por valor, e no Excel em uma amostra de notas de todos os anos e corretoras.

### Corrigido
- Preenchimento de células vazias com `fillna(..., inplace=True)` em uma coluna trocado por atribuição. No pandas 3 a forma antiga não altera nada e os valores vazios das taxas continuariam vazios. (@tommmdl)
- Agrupamentos (`agrupar`, `agrupar_btg`, `agrupar_bmf`) usam `'sum'` em vez da função `sum` do Python, depreciada no pandas 2. (@tommmdl)
- `agrupar_operacoes` e `agrupar_operacoes_correcao`: no pandas 3, `groupby(Series)` seguido de `get_group` falha quando o grupo tem uma só linha, e o erro era engolido, descartando a operação sem aviso. Agora o agrupamento é feito pelo nome da coluna.
- Notas XP/Rico/Clear de várias páginas em que algumas linhas trazem só a quantidade (sem marca de observação) davam erro (`TypeError`) ou, no pandas 3, quantidade 0. Exemplo: XP 2024-11.
- Python 3.15 usa UTF-8 por padrão; a saída do `taskkill` (usado para fechar o Excel) é lida agora com tolerância a caracteres inválidos, em vez de interromper a gravação no Excel.

### Alterado
- Python mínimo: 3.12 (exigido pelo numpy 2.5). Pacote para Windows com Python 3.15.
- `requirements.txt` com as versões atuais testadas.
- `tools/montar_pacote.py` passou a calcular as dependências do pacote a partir das bibliotecas instaladas (arquivo `RECORD` de cada uma), em vez de uma lista fixa de arquivos, e aceita a versão do Python embutido informada pelo arquivo `.zip`.

## [2.1.1] - 2026-10-09

Correção de instalação: uma instalação nova da 2.1.0 falhava por incompatibilidade de bibliotecas. Nenhuma alteração no código do programa.

### Corrigido
- `requirements.txt` fixa `numpy==1.23.5` e `setuptools==69.5.1`. Sem isso o pip instalava o numpy 2.x, incompatível com o pandas 1.3.3, e o setuptools 81 ou superior, que não traz o `pkg_resources` usado pelo tabula-py 2.3.0.

### Alterado
- O pacote para Windows passa a levar o **Python 3.9.7 e o Java 17 (Temurin) embutidos**. Não é preciso instalar Python, bibliotecas nem Java: basta extrair o `.zip` e dar dois cliques em `COIR.bat`.
- Estrutura do pacote simplificada: na pasta principal ficam `COIR.bat`, `Entrada`, `Saida`, `Resultado`, `dados`, `modelos` e `sistema` (programa, Python e Java).
- A instalação manual (Python próprio, `pip install -r requirements.txt`) continua possível para quem usa o código-fonte.
- Novo `tools/montar_pacote.py`, que monta o pacote para Windows.
- Planilha-modelo `modelos/COIR.xlsb` revisada: uma linha que estava oculta foi corrigida.

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

[2.2.0]: https://github.com/MarceloPCF/COIR/releases/tag/v2.2.0
[2.1.1]: https://github.com/MarceloPCF/COIR/releases/tag/v2.1.1
[2.1.0]: https://github.com/MarceloPCF/COIR/releases/tag/v2.1.0
