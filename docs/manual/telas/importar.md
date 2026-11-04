# Importar

![Importar (banco de demonstração)](../imagens/importar.png)

Duas abas:

- **Notas de corretagem (PDF)**: escolha os PDFs (ou arraste) ou importe a pasta das notas inteira. Cada nota é conferida ao centavo (negócios, custos e líquido) antes de ser gravada; a que não fecha não entra e o motivo aparece na lista. Importar o mesmo arquivo de novo não duplica nada. Um PDF pode ter várias notas e várias folhas.
- **PDF com senha** (notas enviadas por e-mail): o COIR tenta sozinho o começo do CPF dos titulares já cadastrados. Na primeira importação, ou se a senha for outra, digite-a no campo **Senha dos PDFs** antes de escolher os arquivos; ela vale para todo o lote e fica só na memória até fechar o programa (nunca é gravada). Pela linha de comando: `coir importar --senha 123`.
- **Planilhas da B3 (xlsx)**: as planilhas da Área do Investidor (movimentação, negociação, posição/consolidado mensal). A movimentação traz os proventos e os eventos; o consolidado serve à conciliação com a B3.

O titular sai do CPF impresso na nota; se ele não existir, é criado (confira o nome em Cadastros).

[Voltar ao índice do manual](../README.md)
