# Registro de auditoria

O objetivo: corrigir um problema de qualquer cliente só com o arquivo que ele manda, sem receber
nenhum dado pessoal dele. O código fica em `coir/auditoria/`, e os fatos do domínio viram eventos em
`coir/servicos/trilha.py`. Para o usuário, veja [a tela Suporte](manual/telas/suporte.md).

## Os dois registros

| | `logs/coir.log` | `logs/auditoria.jsonl` |
|---|---|---|
| Para quem | o usuário (Configurações > Registro de atividades) | o suporte |
| Formato | texto livre | JSON por linha |
| Tamanho | 1 MB × 6 | 2 MB × 6 |
| Conteúdo | mascara CPF; tem nomes de arquivo e a pasta de dados | sem nenhum dado pessoal, já na gravação |
| Sai do computador | nunca | só pela exportação, com ação do usuário |

## Uma linha

```json
{"t":"2026-11-05T14:37:44.663-03:00","s":"aa816ff7","n":34,"nv":"A","ev":"extracao.rejeicao","op":"74743f",
 "d":{"arquivo":"arq_ab2d34a66f","paginas":[1],"nota":"not_6406b5f257","data":"2025-01-10",
      "classe":"ConciliacaoError","campo":"líquido","diferenca":"-5.17",
      "texto":"nota 99999: líquido: esperado -99999.99, lido -99999.99"}}
```

`t` hora; `s` sessão (uma por abertura do programa); `n` sequência na sessão (buraco é linha perdida);
`nv` nível (I, A, E, C); `ev` evento; `op` a operação a que pertence, quando há; `d` os campos.

## Como ler um problema

- **"O programa travou ou fechou"**: última sessão (`s`), última linha. `importacao.arquivo` com
  `situacao: lendo` sem o `lido` depois aponta o arquivo; `log.registro` com nível C traz a pilha.
- **Erro numa tela**: `log.registro` nível E, com `classe`, `pilha` (`modulo:funcao:linha`) e `ui.tela`
  logo antes, dizendo onde o usuário estava.
- **Mensagem de erro na tela**: `erro.dominio`, com o ponto do código que criou o erro.
- **Nota rejeitada**: `extracao.rejeicao`. Na conciliação, `campo` e `diferenca`; ativo sem cadastro,
  `lido` traz o nome como a nota imprime.
- **Corretora ou layout novo**: `extracao.esqueleto` é a forma da página, linha a linha, sem dado:
  dá para escrever o perfil ou o ajuste da gramática a partir dela. Só um por forma, por sessão.
- **Resultado mudou**: `apuracao.mes` só aparece quando o mês muda na sessão; compare `assinatura` entre
  sessões, e grandezas e `incompletos`/`dependentes` para o motivo. `sessao.inicio` e `sessao.banco`
  dizem versão, revisão do banco e data das tabelas.
- **Linha que falta**: `portao.retidas` conta o que o portão segurou, por motivo.

`importacao.nota` com `situacao: gravada` sai dentro da transação: quem diz se ficou gravada é o
`importacao.fim` da mesma operação (`ok: false` desfaz o lote).

## Limite conhecido

Valores exatos não vêm. Erro que dependa do centavo de um cálculo ainda precisa de um caso que o
reproduza: peça a nota anonimizada (tela Suporte) ou o caso descrito pelo usuário. O registro mostra
que mês e que categoria olhar.

## Regra de privacidade: lista de permissão

Nada é "procurado e removido" na gravação. O que entra é só o que o catálogo permite:

- `catalogo.py`: cada evento lista seus campos; cada campo tem uma classe. Fora da lista não entra.
- `limpeza.py`: o que cada classe faz com o valor.
  - **pseudônimo**: arquivo, nota, conta e titular viram HMAC com segredo local
    (`_privado/segredo-registro.txt`, nunca enviado). Sem o segredo, não dá para testar CPFs um a um.
  - **grandeza**: sinal e potência de dez. **diferença**: exata até 100; acima, grandeza.
  - **texto**: só palavras de `vocabulario.py` (literais do código, sem docstrings, menos prenomes e
    sobrenomes comuns) e códigos de negociação; o resto vira forma (`«L7»`, `9.999,99`, `«caminho»`).
  - **impresso**: nome de ativo em maiúsculas, com a espécie (ON, PN...): é a única classe fora do
    vocabulário, porque "ativo sem cadastro" é o que se precisa ver. Sem a espécie, vira texto limpo.
    Se a leitura errar a coluna, o portão ainda procura os termos do titular nela.
- `portao.py`: na exportação, confere tudo de novo sem confiar no arquivo. Exige que a limpeza não
  mude nada no que leu, procura formatos pessoais e os termos do titular (cadastro, dados para ocultar,
  `termos-sensiveis.txt`). Linha com achado é retida. Depois confere o texto final inteiro: qualquer
  achado recusa a exportação.

## Para registrar algo novo

1. Acrescente o evento em `catalogo.py`, com a classe mais restrita que resolva o diagnóstico. Nunca
   crie campo para valor exato, nome ou caminho.
2. Registre em `trilha.py` (ou direto no serviço, se for uma linha). O motor (`coir/engine`) não
   registra nada: o serviço registra a partir do resultado.
3. `tests/unit/test_auditoria.py` confere que toda chamada do código está no catálogo.
4. Mudou muito texto de mensagem? `uv run python tools/gerar_vocabulario.py` (um teste avisa).
