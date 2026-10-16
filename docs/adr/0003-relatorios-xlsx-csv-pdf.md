# ADR 0003 — Relatórios em xlsx (XlsxWriter), csv e PDF (reportlab)

- Status: aceito (F5, 16/10/2026)

## Contexto
A F5 exporta os relatórios fiscais sem depender do Excel. O stack já previa XlsxWriter para
xlsx; faltava uma biblioteca de PDF. Requisitos:
- Python puro ou com wheel para Windows e Linux;
- empacotável pelo PyInstaller (F9);
- licença compatível com AGPL-3.0;
- tabelas longas com cabeçalho repetido a cada página e "Página X de Y";
- acentos do português sem embutir fonte.

## Decisão
- **PDF:** reportlab (licença BSD). Usa o *platypus* (`SimpleDocTemplate`, `LongTable` com
  `repeatRows`) e as fontes padrão Helvetica (WinAnsi cobre os acentos do português).
- **xlsx:** XlsxWriter, só valores, sem fórmulas, compatível com o Excel 2016.
- **csv:** módulo `csv` da biblioteca padrão, no padrão brasileiro: `;`, vírgula decimal e UTF-8
  com BOM (escolha do usuário).
- **Camada:** as três saídas ficam em `coir/reports/` e partem de um modelo neutro (`Relatorio`).
  A interface (F6) usa o mesmo modelo. `reports` não importa banco, motor nem UI.
- **Decimal:** o xlsx guarda número como double. O escritor arredonda o Decimal à precisão da
  coluna (2 casas em dinheiro, 8 no PM) e só então converte, na gravação. É o único ponto com
  `float`. O csv e o PDF formatam o Decimal como texto.

## Alternativas
- **fpdf2** (LGPL-3.0): menor, mas com menos recursos para relatórios longos.
- **WeasyPrint:** depende de GTK/Pango, o que é ruim para o .zip portátil no Windows.

## Consequências
- Dependências novas: `xlsxwriter` e `reportlab`, com uns 4 MB a mais no pacote.
- Os testes leem o xlsx de volta com openpyxl e o PDF com pypdfium2, que já são dependências.
