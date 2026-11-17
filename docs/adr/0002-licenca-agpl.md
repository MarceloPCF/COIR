# ADR 0002 — Licença AGPL-3.0-or-later para a v3

- Status: aceito (F0, 12/10/2026)

## Contexto
A v2 foi publicada sob MIT. Para a v3, o objetivo é que o COIR continue livre e gratuito para investidores
pessoa física. O alcance futuro do projeto é desconhecido. O risco mais concreto para uma ferramenta de IR é
alguém incorporar o código num produto fechado ou num serviço web pago sem devolver as melhorias.

## Decisão
- A v3 passa a ser **AGPL-3.0-or-later** (`LICENSE`, `pyproject.toml`).
- O código da v2 em `legacy/` continua MIT (`legacy/LICENSE`), como foi publicado.

## Por quê AGPL e não MIT ou GPL
- Trocar de uma licença restritiva para uma permissiva é sempre possível para o titular dos direitos; o
  contrário não: o que sai sob MIT fica MIT para sempre. Com o alcance incerto, a escolha reversível é a
  restritiva.
- A GPL não cobre o uso como serviço pela internet; a AGPL cobre (seção 13).
- Para quem usa o programa nada muda; o titular continua livre para relicenciar ou licenciar comercialmente.

## Consequências
- Todas as dependências previstas no CLAUDE.md (MIT, BSD, Apache-2.0; PyInstaller GPL com exceção) são
  compatíveis com AGPL v3. A proibição de PyMuPDF continua valendo por decisão de stack, não mais por licença;
  revê-la exige nova ADR.
- Contribuições de terceiros: antes de aceitar a primeira, criar `CONTRIBUTING.md` com cláusula em que o
  contribuidor autoriza o mantenedor a relicenciar a contribuição; sem ela, cada contribuição amarra a licença.
  Feito em 13/11/2026: `CONTRIBUTING.md` traz o Termo de Licença de Contribuição (versão 1.0), aceito
  por uma caixa no modelo de pull request (`.github/pull_request_template.md`). É **licença não
  exclusiva**, gratuita e irrevogável, com direito de sublicenciar e de distribuir sob outra licença,
  e não cessão: a Lei 9.610/1998 prevê o licenciamento como forma de transferência (art. 49), e a
  cessão pede mais (art. 50: sempre por escrito, presumida onerosa, com tempo, lugar e preço). O
  termo declara a gratuidade (contra a presunção do art. 50), cobre as modalidades de uso existentes
  na data (o art. 49, V não deixa ir além) e pede a autorização do empregador quando o vínculo der a
  ele os direitos sobre o programa (Lei 9.609/1998, art. 4º). Os direitos morais ficam com o autor
  (Lei 9.609, art. 2º §1º). Não passou por advogado.
- A interface (F6) deve mostrar a licença e o link do código-fonte (tela "Sobre"). Feito em
  11/11/2026 na seção "Licença e código-fonte" de Configurações.
- O pacote leva `AVISOS-DE-TERCEIROS.txt` com o texto integral das licenças de tudo o que vai no
  .zip, gerado por `tools/licencas.py` a partir dos executáveis e de
  `empacotamento/licencas/componentes.toml`; a montagem falha se algum arquivo ficar sem licença.
- Revisar esta decisão antes do lançamento da v3.0.
