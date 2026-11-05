# Suporte

![Suporte (banco de demonstração)](../imagens/suporte.png)

Para um problema no COIR: um erro numa tela, um resultado estranho, ou uma nota que ele não lê ou
lê errado. Você manda o registro ou a nota **sem os seus dados pessoais**, e quem desenvolve o COIR
corrige.

## Enviar o registro ao suporte

O COIR registra o que faz: importações, notas rejeitadas, propostas de eventos, cálculos e erros.
**Preparar o registro** grava, numa pasta nova dentro de `suporte` (na pasta de dados), os últimos 30
dias: um arquivo que você pode abrir para ver o que vai (`.jsonl`, texto) e o `.zip` para anexar.
**Escrever e-mail** abre uma mensagem para `suporte.coir@gmail.com`: anexe o `.zip`, conte o que
aconteceu e envie. Nada sai do seu computador sem essa ação sua.

O registro não tem CPF, nome, conta, endereço, telefone, e-mail, nomes de arquivos ou de pastas,
textos das notas nem valores exatos, nem no arquivo que fica no seu computador. Arquivo, nota, conta e
titular aparecem como um código que não dá para reverter (`arq_3fa9c17d20`), e valores só como ordem
de grandeza (`+1e3` é algo entre 1 mil e 10 mil). A exceção é a diferença de uma nota que não fecha,
exata até 100, porque é ela que explica o erro. Ele mostra os códigos dos ativos e as datas dos
problemas e das propostas de eventos. Antes de gravar, o COIR confere tudo de novo; uma linha que não
passa fica de fora, e a tela diz quantas foram.

Se o registro não bastar (uma corretora nova, um erro que dependa do centavo), o suporte pede a nota
anonimizada, abaixo.

## Dados para ocultar

Escolha o titular. Nome, CPF e contas vêm do cadastro. Complete com o que sai nas suas notas:
endereço, bairro, cidade, CEP, telefones, e-mail, assessor, banco/agência/conta corrente e
outros termos. Se mudou de endereço ou de telefone, ponha um por linha: as notas antigas
mostram os dados antigos.

O COIR procura esses dados em qualquer lugar da nota, em qualquer corretora: com ou sem acento,
maiúsculas ou minúsculas, com as palavras coladas (como em notas antigas) e números com ou sem
pontos e traços. Nome, CPF, contas, e-mail e números longos são trocados onde aparecerem. Uma
palavra solta do endereço (o bairro "Vale", por exemplo) só é trocada nas linhas que já têm
outro dado seu, porque a mesma palavra pode ser de outra coisa (a ação VALE ON).

Para quem tem vários titulares (um contador, por exemplo), cada um tem os seus dados. O COIR
sabe de quem é cada nota pelo CPF impresso nela, e procura os dados de todos os titulares em
todas as notas.

Esses dados ficam **só no seu computador**, na pasta de dados (`_privado`), e não vão no envio.

## Enviar notas para suporte

Escolha os PDFs das notas. O COIR grava, numa pasta nova dentro de `suporte` (na pasta de
dados), para cada nota:

- o **texto** que o COIR lê dela, com os dados pessoais trocados por fictícios do mesmo tamanho,
  no mesmo lugar (`nota-01.txt`);
- a **imagem** de cada página, com tarja preta sobre o que foi trocado (`nota-01-pagina-1.png`).

Os arquivos levam nomes neutros, porque o nome do PDF pode ter o seu nome. Os PDFs não são
guardados.

Uma nota só fica pronta se passar nas conferências:

- o COIR lê na cópia exatamente o que lê no original;
- nenhum dado pessoal do original sobra na cópia, e nenhum nome, CPF, conta ou e-mail que você
  informou;
- numa corretora que o COIR não conhece, o CPF impresso tem de ter dados para ocultar
  cadastrados, e o nome informado tem de aparecer na nota. Sem isso o COIR não sabe onde estão os
  seus dados e recusa a nota.

A lista mostra, para cada nota, de quem ela é, quantas páginas e quantos trechos foram
ocultados, ou por que foi recusada. **Ver** abre as imagens; **Abrir a pasta** mostra todas no
visualizador do Windows. Um contorno laranja marca um dado seu que ficou à mostra (como a palavra
"Vale" na ação): confira se é mesmo de outra coisa.

**Atenção:** mesmo sem os seus dados, as notas continuam mostrando as operações (ativos, datas,
quantidades e valores).

Depois de conferir, marque **Conferi as imagens** e clique em **Gravar o .zip**. **Escrever
e-mail** abre o seu programa de e-mail com uma mensagem para `suporte.coir@gmail.com`. Anexe o
`.zip`, descreva o problema e envie. Nada sai do seu computador sem essa ação sua.

Pela linha de comando: `coir notas anonimizar <PDFs ou pastas> [--imagens]`.

[Voltar ao índice do manual](../README.md)
