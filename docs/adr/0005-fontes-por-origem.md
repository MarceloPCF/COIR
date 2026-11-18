# ADR 0005 — Dados de referência separados por origem

- Status: aceito (18/11/2026)

## Contexto
As tabelas públicas de `data/` vêm de várias origens: dados abertos da CVM e do Banco Central
(ODbL 1.0, com atribuição), curadoria do projeto, normas e arquivos públicos da bolsa. Cada
origem tem a sua atribuição em `data/LICENCA_DADOS.md`.

Até a 3.1.2, os eventos de todas as origens ficavam numa tabela só (`eventos_b3.csv`, com a
coluna `origem`). Os serviços também liam as tabelas da bolsa pelo nome. Para trocar uma origem,
ou usar só algumas, seria preciso mexer no motor e nos serviços.

## Decisão
- **Conceitos, não arquivos:**
  - os serviços pedem o conceito a `coir.fontes` (eventos, proventos, emissores, instrumentos,
    tickers por ISIN, saídas), sem saber de que arquivo ele vem;
  - cada origem tem o seu provedor em `coir/fontes/provedores/` (`curadoria`, `cvm`, `bcb` e
    `b3`), e cada provedor lê só os arquivos da sua origem;
  - `coir/referencias.py` continua como fachada das tabelas fixas do motor.
- **Composição** (`coir/fontes/composicao.py`):
  - eventos: CVM > bolsa (a curadoria não tem tabela de eventos). O mesmo evento nas duas
    fontes entra uma vez, com as origens somadas; quando o conteúdo diverge, vale a CVM e a
    diferença fica registrada;
  - emissores e tipos de ativo: curadoria > bolsa > CVM. A bolsa reaproveita o código de emissor
    depois que a companhia ou o fundo sai, e a CVM guarda quem o usou antes: a bolsa diz quem
    tem o código hoje, e a CVM completa o que a bolsa não traz. O código que a CVM liga a mais de
    um CNPJ fica de fora;
  - exceção, no Bens e Direitos: o ticker antigo cujo código a bolsa deu depois a outra empresa
    (ou cujo CNPJ na bolsa é o de uma filial) usa o CNPJ da companhia que o declarou no
    formulário cadastral da CVM (`correcoes_por_ticker`);
  - ISIN de evento sem ticker em nenhuma tabela: o ticker sai da estrutura do ISIN (classe ->
    final do código), só quando o cadastro da CVM ou a curadoria confirma que ele existe.
- **Arquivos por origem (formato 3 do pacote):**
  - `eventos_cvm.csv` e `eventos_b3.csv`; o evento que as duas fontes trazem está nos dois;
  - `valores_mobiliarios_cvm.csv` (códigos das companhias) e `fundos_cvm.csv` (cotas de FII e
    Fiagro), dos dados abertos da CVM;
  - `saidas_b3.csv` passou a `saidas.csv`, porque é curadoria;
  - `pacote.json` lista os arquivos de cada fonte com a licença, tirados de `procedencia.csv`
    (o teste exige que todo arquivo de `data/` esteja lá);
  - o app aceita os formatos 2 e 3.
- **Chave `fontes.b3`** em `config.json`:
  - `pacote` (padrão): usa as tabelas da bolsa que vêm no pacote;
  - `desligada`: nenhum arquivo da bolsa é aberto. As telas Eventos e Proventos dizem o que
    fica sem cobertura.
- **Camadas:** `coir/fontes` importa só `domain`, `erros`, `parsing` e `recursos`. O motor não
  importa `fontes` (`tests/unit/test_camadas.py`).

## Consequências
- Trocar uma origem, ou usar só algumas, não mexe no motor nem nos serviços.
- A equivalência foi conferida a cada fase: com a fonte da bolsa ligada, o resultado é o mesmo,
  exceto onde a CVM preenche lacuna ou corrige o CNPJ do ticker antigo.
- Cobertura que só a bolsa dá:
  - nome de pregão de ativo fora das tabelas curadas;
  - eventos que só a bolsa lista;
  - proventos em dinheiro;
  - restituição de capital;
  - aviso de ativo que parou de negociar;
  - ETF, BDR, FIP, FIDC e FI-Infra, porque a CVM não publica dados abertos que liguem esses
    fundos ao código de negociação.
- Com a fonte da bolsa desligada, o código de um FII antigo que a bolsa deu depois a um ETF
  sai como FII. Ligada, vale o tipo da bolsa.

## Alternativas
- **Manter uma tabela só, com a coluna `origem`**: não permite desligar uma origem sem filtrar
  linha a linha em cada leitor.
- **CVM acima da bolsa em emissores e tipos**: troca o CNPJ e o tipo atuais pelos de quem usou
  o código antes.
