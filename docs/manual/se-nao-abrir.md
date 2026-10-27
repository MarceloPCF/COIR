# Se o COIR não abrir

O COIR não tem assinatura digital. Por isso o Windows e alguns antivírus desconfiam de um programa
novo. Os casos abaixo cobrem quase
tudo. Os seus dados não correm risco em nenhum deles.

## 1. Aviso azul "O Windows protegeu o computador"

É o SmartScreen. Clique em **Mais informações** e depois em **Executar assim mesmo**. O aviso
aparece só na primeira vez.

## 2. Nada acontece depois dos dois cliques (versão 3.0.0)

Na 3.0.0, a janela não abria quando o `.zip` era extraído direto do download: o Windows marca os
arquivos como "baixados da internet" e essa marca impedia a janela de carregar. A partir da 3.0.1
o COIR resolve isso sozinho. Na 3.0.0:

1. Feche o COIR no Gerenciador de Tarefas (Ctrl+Shift+Esc), se ele aparecer lá.
2. Apague a pasta extraída.
3. Clique com o botão direito no `.zip` › **Propriedades** › marque **Desbloquear** › **OK**.
4. Extraia de novo e abra o `COIR.exe`.

## 3. "O Windows não pode acessar o dispositivo, caminho ou arquivo especificado"

Quase sempre é o antivírus, que pôs o `COIR.exe` em quarentena. Programas feitos em Python e sem
assinatura são confundidos com ameaça com alguma frequência (falso positivo).

1. Veja se o `COIR.exe` ainda está na pasta. Se sumiu, foi o antivírus.
2. Abra **Segurança do Windows** › **Proteção contra vírus e ameaças** › **Histórico de
   proteção**. Se houver um item com o COIR, abra-o e escolha **Permitir no dispositivo** (ou
   **Restaurar**).
3. Com outro antivírus instalado, procure a quarentena dele e restaure o arquivo, ou crie uma
   exceção para a pasta do COIR.
4. Extraia o `.zip` de novo, se o arquivo tiver sido apagado.

Antes de permitir, se quiser, confira que o `.zip` é o original: o arquivo `.sha256` publicado
na mesma página de versões tem o código do `.zip`. No PowerShell:

```
Get-FileHash .\COIR-windows.zip -Algorithm SHA256
```

## 4. "Controle Inteligente de Aplicativos bloqueou"

O Windows 11 pode vir com o **Controle Inteligente de Aplicativos** ligado. Ele bloqueia qualquer
programa sem assinatura e não tem a opção "executar assim mesmo". A configuração fica em
**Segurança do Windows** › **Controle de aplicativos e do navegador** › **Configurações do
Controle Inteligente de Aplicativos**. Desligá-lo vale para o computador inteiro e, depois de
desligado, o Windows não deixa ligar de novo sem reinstalar. Decida com calma: com ele ligado, o
COIR não abre.

## Ainda não abriu?

Abra o **Prompt de Comando** na pasta do COIR e rode `coir-cli.exe ui`. A janela preta mostra o
erro. Copie a mensagem (sem CPF, contas ou valores) e envie pelo
[pedido de ajuda](https://github.com/MarceloPCF/COIR/issues).

[Voltar ao índice do manual](README.md)
