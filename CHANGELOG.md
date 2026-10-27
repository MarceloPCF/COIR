# Histórico de versões

Este projeto segue o [Versionamento Semântico](https://semver.org/lang/pt-BR/) (`MAIOR.MENOR.CORREÇÃO`):

* **MAIOR** – mudança que quebra a compatibilidade (ex.: novo formato da planilha de resultado);
* **MENOR** – funcionalidade nova que continua compatível;
* **CORREÇÃO** – correção de erro, sem funcionalidade nova.

O formato deste arquivo segue o [Keep a Changelog](https://keepachangelog.com/pt-BR/).

## [3.0.1] - 2026-10-27

### Corrigido
- **A janela não abria** depois de extrair o `.zip` baixado: o Windows marca cada arquivo
  extraído como "baixado da internet", e com essa marca a DLL da janela não carregava. O
  programa ficava rodando sem janela e sem nenhum aviso. Agora o COIR tira essa marca das próprias
  DLLs ao abrir. Na 3.0.0, o contorno é desbloquear o `.zip` (Propriedades › Desbloquear) antes
  de extrair.

### Adicionado
- Página do manual **Se o COIR não abrir**: aviso do Windows, antivírus que põe o programa em
  quarentena e Controle Inteligente de Aplicativos. Link no LEIA-ME e no rodapé do site.

## [3.0.0] - 2026-10-26

Nova versão escrita do zero, sem dependência do Excel. É um programa para Windows que se baixa,
extrai e abre (`COIR.exe`), sem instalar Python nem Java. Os dados ficam numa pasta sua
(`COIR_DADOS`); nada sai do computador. O código da v2 (2.2.0) fica em `legacy/` como referência.

### Em resumo
- **Telas**: painel do ano, importação de notas e planilhas da B3, notas com o PDF ao lado,
  operações, apuração mensal com memória de cálculo, DARF (com multa e juros de mora, pagamento e
  comprovante), proventos, outros custos, declaração anual ficha a ficha, relatórios, carteira e
  preço médio, eventos corporativos, opções, futuros, conciliação com a B3, cadastros e
  configurações. Tema claro e escuro.
- **Leitura das notas**: um só extrator SINACOR (Bovespa e BM&F) para Ágora, BTG/Necton, Clear,
  Rico e XP, com conciliação ao centavo: nota que não fecha não entra.
- **Imposto**: preço médio por titular, day trade pela marcação da corretora, isenção de
  R$ 20 mil, compensação de prejuízo por categoria, FII, opções (exercício e pó), termo, futuros
  (inclusive em US$), IRRF por ficha, DARF com mínimo de R$ 10 e vencimento no último dia útil.
- **Declaração**: Bens e Direitos (um item por ativo, CNPJ e discriminação prontos), rendimentos
  isentos e exclusivos por fonte pagadora, renda variável mês a mês, conferência com a declaração
  entregue.
- **Tabelas públicas** (eventos, proventos, feriados, Selic) atualizadas pelo botão "Atualizar
  tabelas"; aviso de versão nova do programa (só avisa, não baixa).
- **Atualização segura**: o banco é migrado automaticamente com backup; banco criado por versão
  mais nova é recusado com explicação.
- O que ficou para depois está em [`docs/roteiro-pos-3.0.md`](docs/roteiro-pos-3.0.md).

### Por que meus números mudaram em relação à v2
A v3 refaz o cálculo a partir das notas, pela norma. Diferenças esperadas em relação à planilha
da v2 (todas são correções da v2):
- **Preço médio por titular**, somando todas as corretoras (a v2 usava o PM por conta).
- **Custos da nota inteiros**: a v3 usa o total que concilia com o resumo da nota; a v2 lançava
  parte dos custos.
- **IRRF de day trade (1%) deduzido** no mês; a v2 não o deduzia.
- **Saldo de IRRF transportado** para os meses seguintes do ano; a v2 o perdia no próprio mês.
- **Isenção de R$ 20 mil pelo valor das vendas de ações à vista do mês**; a v2 a inferia pelo IRRF,
  o que falhava com duas ou mais corretoras e podia isentar ETF.
- **Prejuízo de mês isento compensável**; a v2 o descartava.
- **FII como categoria própria** (20%, prejuízo separado); a v2 não tinha FII.
- **DARF**: mínimo de R$ 10 (abaixo disso acumula para o mês seguinte) e vencimento no último dia
  útil do mês seguinte; a v2 não aplicava o mínimo e usava o último dia corrido.
- **Eventos corporativos** (desdobramento, grupamento, bonificação, incorporação, cisão, troca de
  ticker) aplicados pelas tabelas públicas da B3, sem ajuste manual.
- **Exercício de opção** na linha do ativo-objeto, sem isenção; opção que virou pó apurada no
  vencimento.
- **IRRF até R$ 1,00** (dispensado, não retido) não vira crédito.

### Adicionado
- Licença da v3: AGPL-3.0-or-later (a v2 em `legacy/` continua MIT).
- Fase F0 (fundação): projeto uv com Python 3.13, estrutura `coir/`, configuração da pasta de dados (`COIR_DADOS`), CLI mínima, verificador de dados pessoais (pre-commit, pre-push e CI) e testes automatizados em Windows e Linux.
- Fase F8 (declaração anual): fichas do programa IRPF por exercício (`data/irpf/<exercício>.yaml`) com Bens e Direitos (um item por ativo, CNPJ do pacote público da B3, discriminação pronta), Dívidas (posição vendida), rendimentos isentos e exclusivos por fonte pagadora, renda variável mês a mês e imposto pago; tela "Declaração anual" com botão de copiar e marcação do que já foi lançado; leitura e conferência da declaração entregue (PDF do programa); data do crédito do JCP (a mais tarde entre a aprovação e a data com) pela B3, inclusive pelo histórico público de JCP desde 2016 (`data/proventos_historico_b3.csv`), pela declaração entregue ou informada; JCP pago mais de um ano depois do crédito a receber em cada 31/12; origem dos reembolsos de BTC; venda no dia da liquidação do termo pelo preço a termo, com o excedente saindo das ações antigas pelo custo médio; lista gerada das fontes fiscais (`docs/fontes-fiscais.md`).
- Fase F9 (versão 3.0): programa para Windows em pasta (`COIR.exe` e `coir-cli.exe`, PyInstaller) com `tools/empacotar.py` e workflow que cria o release como rascunho; identidade visual da logomarca (topo marinho, menu retrátil); aviso de versão nova; banco mais novo que o programa recusado com tela de erro; day trade pela marcação D da corretora (parâmetro `day_trade_criterio`); saldos iniciais em 31/12 (custódia e prejuízo por categoria); origem do reembolso de BTC informada pelo usuário; leitura da DIRPF 2024; registro de atividades legível; gerador de notas fictícias (`tools/notas_ficticias.py`) e banco de demonstração; capturas de tela automáticas (`tools/capturas.py`); manual do usuário em `docs/manual/`, guia de como adicionar corretora e site novo.

## [2.2.0] - 2026-10-21

Versão da v2 (planilha), em `legacy/`.

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

[3.0.1]: https://github.com/MarceloPCF/COIR/releases/tag/v3.0.1
[3.0.0]: https://github.com/MarceloPCF/COIR/releases/tag/v3.0.0
[2.2.0]: https://github.com/MarceloPCF/COIR/releases/tag/v2.2.0
[2.1.1]: https://github.com/MarceloPCF/COIR/releases/tag/v2.1.1
[2.1.0]: https://github.com/MarceloPCF/COIR/releases/tag/v2.1.0
