====================================================
 TARJADOR FGTS — Instruções de instalação e uso
====================================================

INSTALAÇÃO (primeira vez):
  1. Instale o Python 3.9 ou superior: https://python.org
  2. Abra o terminal/cmd na pasta do programa
  3. Execute: pip install -r requirements.txt

EXECUÇÃO:
  python tarjador_fgts.py

====================================================
 COMO USAR
====================================================

1. PASTA DOS EXCELS
   Crie uma pasta (ex: C:\bases_fgts\) e coloque lá
   todos os arquivos Excel com os dados dos funcionários.
   Eles devem seguir o molde fornecido (mesmas colunas).
   Você pode ter quantos arquivos quiser — o sistema
   lê todos automaticamente.

2. SELECIONAR PDF
   Escolha o relatório RT (PDF) do mês que deseja processar.

3. CPF(s)
   Digite o(s) CPF(s) que NÃO devem ser tarjados.
   Separe múltiplos por vírgula ou Enter.
   Exemplo: 123.456.789-00, 987.654.321-00

4. NOME(s) — opcional
   Nome(s) como aparecem no documento.
   Use junto ou no lugar do CPF.

5. TOMADOR(es) — opcional
   CNPJ(s) dos Tomadores que devem aparecer no PDF.
   Se deixar em branco: TODOS os Tomadores são mantidos.
   O número do Tomador que não estiver nesta lista
   será tarjado, junto com todos os dados do seu bloco.

6. PROCESSAR
   O arquivo resultado será salvo na mesma pasta do PDF
   com o sufixo "_filtrado.pdf".

====================================================
 REGRAS DE TARJAMENTO
====================================================

• O sistema só age dentro dos blocos:
    "Tomador: XXXX" → "Total do Tomador"
  Cabeçalho, rodapé e totais nunca são tocados.

• Dentro de cada bloco:
  - Linha cujo CPF está na sua lista → MANTÉM visível
  - Linha cujo CPF NÃO está → TARJA (Nome até Total)

• Se um Tomador não estiver na sua lista de permitidos:
  - O CNPJ do Tomador é tarjado
  - Todos os dados do bloco são tarjados

====================================================
