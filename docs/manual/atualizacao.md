# Atualizar o COIR para uma versão nova

O programa e os seus dados ficam em pastas separadas. Trocar de versão é trocar só a pasta do
programa; os dados continuam onde estão.

## Onde fica cada coisa

| O quê | Onde |
|---|---|
| Programa (`COIR.exe`, `coir-cli.exe`, `_internal\`) | a pasta onde você extraiu o `.zip` |
| Banco de dados (`coir.db`) | a sua pasta de dados (`COIR_DADOS`), escolhida no primeiro uso |
| Backups automáticos | `COIR_DADOS\backups\` (os 10 mais recentes) |
| Tabelas públicas baixadas pelo "Atualizar tabelas" | `COIR_DADOS\referencias\atual\` |
| Relatórios exportados | `COIR_DADOS\relatorios\<ano>\` |
| Notas em PDF e planilhas da B3 | as pastas que você escolheu (por padrão, dentro de `COIR_DADOS`) |
| Qual é a pasta de dados e o tema escolhido | `%APPDATA%\COIR\config.json` |

## Passo a passo

1. Feche o COIR.
2. Baixe o `.zip` da versão nova na página de versões (o COIR avisa no Painel quando há uma, se a
   opção estiver ligada em Configurações).
3. Extraia numa pasta nova (ex.: `Documentos\COIR-3.1.0`) ou apague o conteúdo da pasta antiga e
   extraia por cima. **Não apague a pasta de dados.**
4. Abra o `COIR.exe` novo. Ele acha a pasta de dados pelo `config.json`.
5. Se o banco for de uma versão anterior, o COIR faz um **backup** em `COIR_DADOS\backups\` e
   **migra** o banco sozinho antes de abrir as telas.

Depois de conferir que tudo abriu bem, a pasta do programa antigo pode ser apagada.

## Voltar um backup

Em **Configurações › Backups**, escolha o backup e clique em restaurar. O COIR guarda uma cópia do
banco atual antes de restaurar. Use isso se algo der errado depois de uma importação ou de uma
atualização.

## Banco de uma versão mais nova

Se você abrir uma versão **mais antiga** do programa com um banco já migrado por uma versão mais
nova, o COIR não abre as telas e explica o motivo (o banco não é alterado). Abra com a versão mais
nova ou restaure, com ela, um backup anterior à atualização.

## Vindo da v2 (planilha)

A v3 não lê a planilha da v2. Importe as notas em PDF desde o início do seu histórico; se o
histórico das notas começar depois de você já ter ações, informe a custódia e os prejuízos em
**Cadastros › Saldos iniciais** (posição em 31/12 do ano anterior ao das primeiras notas). Os
números podem mudar em relação à v2: o [CHANGELOG](../../CHANGELOG.md) explica cada correção.
