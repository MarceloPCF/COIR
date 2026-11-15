# Como contribuir com o COIR

Obrigado pelo interesse. O COIR é mantido por uma pessoa, no tempo livre, e é distribuído de graça
sob a licença [AGPL-3.0-or-later](LICENSE). Este arquivo explica como ajudar e as condições para
que um código seu entre no projeto.

## Sem programar

- **Corretora que o COIR não lê, erro ou dúvida:** use a tela **Suporte** do programa. Ela gera
  uma cópia da nota com os seus dados trocados (nome, CPF, conta, endereço) e abre o e-mail para
  suporte.coir@gmail.com; você confere as imagens e anexa. Também dá para usar o formulário do
  site (https://coir.com.br) ou abrir uma issue no GitHub.
- **Nunca mande** o PDF original, planilhas da B3, o banco do COIR, nome, CPF ou número de conta,
  nem numa issue nem num pedido de contribuição: o que vai para o GitHub é público.

## Com código

Leia antes:

- o [README](README.md) (instalação com `uv`, testes e comandos);
- [`docs/regras-fiscais.md`](docs/regras-fiscais.md) e [`docs/fontes-fiscais.md`](docs/fontes-fiscais.md),
  se a mudança mexe em cálculo;
- [`docs/como-adicionar-corretora.md`](docs/como-adicionar-corretora.md), para perfil de corretora;
- [`docs/adr/`](docs/adr/), para as decisões de arquitetura.

Regras que o projeto não abre mão:

1. Dinheiro e quantidades com `decimal.Decimal`, nunca `float`.
2. Nenhum parâmetro fiscal no código (alíquota, limite, código de DARF...): tudo em `data/`, com
   vigência.
3. Regra fiscal nova ou alterada vem com teste de números conferidos à mão e com a fonte legal
   (lei, IN, Solução de Consulta, artigo e parágrafo) no docstring do teste.
4. Nenhum dado pessoal no repositório, nem em fixture: use notas fictícias
   (`tools/notas_ficticias.py`) ou texto anonimizado pela tela Suporte.
5. Erro explícito em vez de silêncio: nada de `except: pass`.

Antes de abrir o pedido de contribuição (pull request):

```bash
uv run ruff check . && uv run ruff format --check .
uv run pyright
uv run pytest -m "not slow"
```

Mensagens de commit no formato `tipo(área): descrição` (`feat`, `fix`, `test`, `docs`,
`refactor`), em português.

## Licença das contribuições

O COIR é publicado sob a AGPL, mas o mantenedor precisa poder, no futuro, publicá-lo também sob
outra licença (por exemplo, uma mais permissiva, ou uma licença comercial para módulos
complementares). Sem autorização expressa de cada pessoa que contribuiu, isso deixa de ser
possível. Por isso, só entram no projeto contribuições aceitas nos termos abaixo.

### Termo de Licença de Contribuição (versão 1.0, de 13/11/2026)

Ao marcar a caixa de aceite no pedido de contribuição, você (a pessoa que contribui) declara e
concorda que:

1. **Objeto.** Este termo vale para todo código, texto, imagem, tabela de dados ou outro material
   que você enviar ao repositório MarceloPCF/COIR por pedido de contribuição, commit, patch ou
   anexo de issue, com a intenção de que seja incorporado ao projeto (a "Contribuição").
2. **Licença.** Você concede ao mantenedor e titular do COIR (MarceloPCF no GitHub) e a
   quem ele vier a transferir a titularidade do projeto, uma **licença não exclusiva, gratuita,
   irrevogável, válida em todos os países e pelo prazo total de proteção dos direitos autorais**,
   para usar, reproduzir, modificar, adaptar, traduzir, incluir em outras obras, distribuir,
   publicar, exibir, oferecer como serviço pela internet e comercializar a Contribuição, em todas
   as modalidades de utilização existentes nesta data.
3. **Sublicença e outras licenças.** A licença inclui o direito de sublicenciar e de distribuir a
   Contribuição, sozinha ou dentro do COIR, sob a AGPL ou sob **qualquer outra licença**, livre ou
   não, gratuita ou paga.
4. **Gratuidade.** A licença é concedida sem nenhuma remuneração, agora ou no futuro.
5. **O que continua seu.** A licença não é exclusiva: você continua titular da Contribuição e
   pode usá-la como quiser, inclusive em outros projetos. Os seus direitos morais de autor
   continuam com você; o histórico do repositório guarda a autoria de cada mudança.
6. **Autoria e direito de licenciar.** A Contribuição é criação sua, ou você tem o direito de
   licenciá-la nestes termos. Se você a fez como empregado, estagiário, servidor ou prestador de
   serviço, e o vínculo atribui os direitos sobre programas ao contratante (Lei 9.609/1998, art.
   4º), você declara ter autorização dele para este termo. Se partes da Contribuição vierem de
   terceiros, você indica no pedido de contribuição a origem e a licença dessas partes.
7. **Dados pessoais.** A Contribuição não contém dados pessoais de ninguém (nome, CPF, número de
   conta, notas reais, planilhas da B3).
8. **Sem garantia.** A Contribuição é fornecida como está, sem garantia de nenhum tipo.
9. **Aceite.** A marcação da caixa de aceite, feita na sua conta do GitHub, com a versão deste
   termo, junto com a data e o identificador da conta que o GitHub registra no pedido de
   contribuição, vale como a sua concordância por escrito. Uma versão nova do termo só vale para
   as contribuições aceitas depois dela.

O mantenedor pode recusar qualquer contribuição, inclusive por não querer mantê-la.
