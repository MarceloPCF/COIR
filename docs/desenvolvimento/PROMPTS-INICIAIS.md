# Prompts para colar no Claude Code, em ordem

Pasta do projeto: `E:\Claude-IA\COIR-v3.0` (já tem CLAUDE.md, docs\, .claude\agents\ e .gitignore).
Dados reais: pasta EXTERNA `E:\Claude-IA\COIR-dados` (_notas-corretagens, _relatorios-b3, _gabarito).
Ao abrir o Claude Code na pasta do projeto, libere a pasta de dados com `/add-dir E:\Claude-IA\COIR-dados`
e defina a variável de ambiente `COIR_DADOS=E:\Claude-IA\COIR-dados`.

## F0 — Fundação
```
Leia CLAUDE.md e docs/PLANO.md. Estamos na fase F0, na pasta E:\Claude-IA\COIR-v3.0 (ainda sem git).
1) Faça git init, adicione o remote origin https://github.com/MarceloPCF/COIR, faça fetch, crie a tag v2.02
   apontando para origin/main e crie o branch v3 a partir de origin/main, preservando os arquivos que já estão
   na pasta (CLAUDE.md, docs/, .claude/, .gitignore, LEIA-ME.md, PROMPTS-INICIAIS.md).
2) Mova o código v2 versionado para legacy/ e atualize-o com a pasta local E:\Claude-IA\COIR-v.2.02, que é
   a mais recente (ignore PDFs, xlsx/xlsb/xlsm, Resultado/, Saida/, Entrada/, __pycache__ e pastas que começam com _).
   Mostre as diferenças de COIR.py, funcoes.py, btg.py e xp_rico_clear.py antes de sobrescrever.
3) Crie o projeto uv (Python 3.13) com a estrutura do CLAUDE.md e uma config que leia COIR_DADOS.
4) Garanta que nenhum PDF, planilha, .db ou CPF entre no Git (veja .gitignore) e configure ruff, pyright,
   pytest, pre-commit (com hook que bloqueia padrão de CPF) e GitHub Actions (Windows + Linux).
5) NÃO faça push. Use Plan mode antes de executar. Ao final, chame o revisor-codigo.
```

## F1 — Gabarito
```
Fase F1. Os dados reais estão em %COIR_DADOS% (fora do repo): _notas-corretagens, _relatorios-b3 e
_gabarito\gabarito.xlsb. NÃO copie nada para dentro do repositório. Leia a planilha com pyxlsb (sem Excel) e gere
um gabarito legível por máquina em %COIR_DADOS%\gabarito\: operações por nota, PM por ativo após cada
operação, eventos, apuração mensal por categoria e DARF. Mostre-me um resumo (contagens por ano/corretora,
totais por ano) para eu validar. Crie o comparador `coir validar-gabarito` com as tolerâncias do CLAUDE.md.
Peça ao parser-notas-sinacor 3 a 5 notas sintéticas anonimizadas para tests/fixtures/.
```

## F2 — Extrator único
```
Fase F2. Implemente coir/parsing conforme CLAUDE.md: pypdfium2/pdfplumber, parser SINACOR por regex,
perfis YAML por corretora, aliases com rapidfuzz e checagens de conciliação. Comece pelos testes
(fixtures sintéticas + gabarito privado). Use o parser-notas-sinacor para analisar os layouts e o
validador-gabarito ao final. Critério de saída: 100% das notas batem com o gabarito ou divergência explicada.
```

## F3 a F8
```
Fase F<n>. Siga docs/PLANO.md seção 8. Antes de codar regras fiscais, consulte o especialista-fiscal-b3 e
gere os casos de teste em tests/fiscal_cases/. Ao final: validador-gabarito e revisor-codigo.
```

## F9 — Release (só com minha autorização)
```
Fase F9. Gere o pacote PyInstaller (modo pasta) e o .zip. Antes de qualquer push: rode a varredura de dados
pessoais (CPF, contas, nomes, e-mails, PDFs, xlsx, .db) nos arquivos e no histórico do branch v3, mostre a lista
de arquivos que iriam para o GitHub e PARE. Só faça push, PR ou release depois que eu autorizar explicitamente.
```

## Dicas
- Uma fase por branch; PR com CI verde.
- Quando o especialista-fiscal-b3 devolver uma pergunta "decisão do usuário", responda antes de seguir.
- Use `/subtask` (fork) para tarefas paralelas curtas que precisam do contexto da conversa.
