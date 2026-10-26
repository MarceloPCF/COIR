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
- A interface (F6) deve mostrar a licença e o link do código-fonte (tela "Sobre").
- Revisar esta decisão antes do lançamento da v3.0.
