# Como adicionar uma corretora

O COIR lê notas de corretagem no padrão **SINACOR** (Bovespa e BM&F). Cada corretora tem um
**perfil** que diz como reconhecer a nota dela e qual **gramática** (variante do layout) usar.
Corretoras que imprimem a nota igual a outra já conhecida precisam só do perfil; um layout novo
precisa também de ajuste na gramática.

## Se você só usa o COIR

A nota não foi lida? Use o **Pedido de corretora** no site
(https://marcelopcf.github.io/COIR/#corretora) ou abra uma issue no GitHub, com:

- o nome da corretora, o segmento (Bovespa ou BM&F) e o ano da nota;
- a mensagem que o COIR mostrou na importação (tela Importar ou registro de atividades);
- nunca envie o PDF original, o nome, o CPF ou a conta. Se for preciso ver o layout, a resposta
  explica como mandar só o texto da nota com esses dados trocados por valores genéricos (veja
  "Fixture anonimizada" abaixo).

Um anonimizador dentro do programa está previsto para depois da 3.0
([roteiro](roteiro-pos-3.0.md)).

## Para quem desenvolve

### 1. Perfil (`coir/parsing/profiles/<corretora>.yaml`)

```yaml
# Nova Corretora: mesmo layout da XP nova
corretora: NOVA
prioridade: 10
ancoras: ["NOVA CORRETORA DE VALORES", "00.000.000/0001-00"]
gramaticas: {bovespa: xp, bmf: xp}
```

- `ancoras`: textos que aparecem só nas notas dessa corretora (razão social, CNPJ). A comparação
  ignora acentos, espaços e maiúsculas. O primeiro perfil (pela `prioridade`, menor primeiro) com
  alguma âncora na página vence: corretoras do mesmo grupo que imprimem o nome da controladora
  precisam de prioridade menor.
- `gramaticas`: a variante usada em cada segmento. Bovespa: `agora`, `xp` ou `btg`; BM&F: `xp` ou
  `btg` (as diferenças estão no início de `coir/parsing/bovespa.py` e `coir/parsing/bmf.py`).
  Segmento sem gramática faz a página ser rejeitada com mensagem clara.

### 2. Fixture anonimizada (`tests/fixtures/notas/`)

1. Extraia o texto da nota real com o mesmo leitor do COIR (pypdfium2):

   ```bash
   uv run python -c "from coir.parsing.texto_pdf import paginas; import sys; print('\f'.join(paginas(__import__('pathlib').Path(sys.argv[1]))))" nota.pdf > texto.txt
   ```

   Faça isso **fora do repositório** (na sua pasta de dados).
2. Anonimize o texto como descrito em [`tests/fixtures/notas/README.md`](../tests/fixtures/notas/README.md):
   nome do cliente por `FULANO DE TAL`, CPF por `000.000.000-00` (mantendo a pontuação), conta e
   código do cliente por `000000`, assessor por `0000`, endereço genérico e número da nota fictício.
   Mantenha ativos, datas, quantidades, preços, custos, IRRF e líquido: a nota precisa continuar
   fechando na conciliação.
3. Salve como `<corretora>_<ano>_<mês>_<segmento>.txt` e escreva o `<mesmo nome>.esperado.json`
   (operações, resumo, IRRF e líquido), no formato das fixtures existentes.
4. Rode `uv run python tools/verificar_dados_pessoais.py tests/fixtures/notas/*` antes de commitar.

### 3. Teste

`tests/unit/parsing/test_fixtures.py` extrai, concilia e compara cada fixture com o esperado
(atualize a contagem em `test_ha_dez_fixtures`). Rode:

```bash
uv run pytest tests/unit/parsing -q
```

### 4. Conciliação

A nota só entra se os números fecharem entre si, ao centavo (`coir/parsing/conciliar.py`):
negócios = resumo dos negócios, Total CBLC / Bolsa / Despesas = soma dos itens, líquido = CBLC −
Bolsa − Despesas (− IRRF quando a corretora desconta). Se a nova corretora fechar diferente (por
exemplo, um custo num grupo inesperado), ajuste a gramática, nunca a tolerância.

### 5. Nomes de ativos

Se a nota imprime o nome de pregão (ex.: `PETROBRAS PN`), o ticker vem de `data/ativos.csv` e da
lista pública da B3. Nome sem cadastro falha com sugestões; acrescente a linha em
`data/ativos.csv`.

### Exemplo sem dados reais

`tools/notas_ficticias.py` gera notas no layout da XP a partir de um roteiro
(`tools/roteiros/demonstracao.yaml`). Serve para ver o formato esperado do texto e para testar o
COIR sem notas reais.
