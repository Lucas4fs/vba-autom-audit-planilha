<p align="center">
    <img src="Imagens\CapaProjetoVBA.png">
    <br>
    <h1 align="center">
    🟦🟧 PREENCHIMENTO/AUDITORIA DE PLANILHA NO EXCEL AUTOMATIZADO COM VBA 🟨🟪
    </h1>
</p>
<br>

<h2>
    📑 SUMÁRIO
</h2>

1. [INTRODUÇÃO](#1-introdução)<br>
   1.1 - [Intuito do Projeto](#11-intuito-do-projeto)<br>
   1.2 - [Situação Abordada](#12-situação-abordada)<br>
2. [DESENVOLVIMENTO](#2-desenvolvimento)<br>
3. [CONCLUSÃO](#3-conclusão)
4. [FERRAMENTAS UTILIZADAS](#4-ferramentas-utilizadas)
5. [COMO ADQUIRIR A AUTOMATIZAÇÃO](#5-como-adquirir-a-automatização)

## 1 INTRODUÇÃO

### 1.1 Intuito do Projeto

<p>
    O intuito do projeto é mostrar como usar programação para automatizar parte de uma tarefa que é realizada manualmente, além de economizar tempo é possível ter 100% de acertividade no resultado da tarefa dês de que todos os requisitos sejam cumpridos corretamente.
</p>

### 1.2 Situação Abordada

<p>
   Um supermercado tem uma planilha que registra a saída de todas as notas fiscais(vendas) contendo várias informações, entre elas estão os impostos de ICMS e PIS/COFINS que são pagos encima daqueles produtos que estão saindo do supermercado, para ter certeza de que o mercado está pagando a quantidade correta de imposto é feita uma auditoria comparando a planilha "Cliente.xlsm"(contém os dados inseridos pelo supermercado) com a planilha "Audit.xlsm" (contém a auditoria dos dados inseridos pelo cliente)
</p>

## 2 DESENVOLVIMENTO

- Planilha "Cliente.xlsm"

[DOWNLOAD](Cliente.xlsm)

- Planilha "Audit.xlsm"

[DOWNLOAD](Audit.xlsm)

<p>
  Dentro da planilha "Cliente.xlsm" existem várias colunas, cada uma tem sua função, abaixo será detalhado as funções das colunas que iremos usar:
</p>

 - Considerações ICMS

<p>
 Preenchida pelo auditor para informar se o ICMS inserido pelo cliente está correto ou incorreto, se estiver incorreto é necessário informar a porcentagem correta da alíquota.
</p>

- Considerações PIS/COFINS

<p>
 Preenchida pelo auditor para informar se o PIS/COFINS inserido pelo cliente está correto ou incorreto, se estiver incorreto é necessário informar o regime correto.
</p>

- codigo_produto

<p>
 Informa o código de identificação do produto.
</p>

- descricao_produto

<p>
 Informa o nome do produto.
</p>

- CST_ICMS

<p>
 Informa o CST do ICMS usado pelo supermercado.
</p>

- CST_ICMS_Audit

<p>
 Informa o CST do ICMS usado pelo auditor.
</p>

- Aliquota_ICMS

<p>
 Informa a alíquota não efetiva do produto inserida pelo supermercado.
</p>

- Aliquota_Efet_ICMS

<p>
 Informa a alíquota efetiva do produto inserida pelo supermercado (essa é a alíquota definitiva que é descontada em nota).
</p>

- Alíquota_ICMS_Audit

<p>
 Informa alíquota efetiva inserida pelo auditor.
</p>

- CST_PIS

<p>
 Informa o CST do PIS inserido pelo supermercado
</p>

- CST_PIS_Audit

<p>
 Informa o CST do PIS inserido pelo auditor.
</p>

- CST_Cofins

<p>
 Informa o CST do COFINS inserido pelo supermercado
</p>

- CST_Cofins_Audit

<p>
 Informa o CST do COFINS inserido pelo auditor
</p>

- Num Nf

<p>
 Informa o número da nota fiscal inserido pelo supermercado
</p>

<p>
  Dentro da planilha "Audit.xlsm" na aba "Exceções de ST Alíquota e ST" são cadastrados os mesmos produtos que o cliente possui na planilha "Cliente.xlsm", mas com a alíquota de ICMS correta
</p>

<img src="Imagens\colunasAudit.png">

<p>
  Dentro da planilha "Audit.xlsm" na aba "Exceções PIS Cofins Aliq 0" são cadastrados os mesmos produtos que o cliente possui na planilha "Cliente.xlsm", mas com o PIS e COFINS correto.
</p>

<img src="Imagens\colunasPISeCOFINS.png">

<p>
  Com as planilhas abertas apertamos "ALT + F8" para exibir as macros disponíveis, definimos um atalho para cada macro, no nosso exemplo iremos usar "CRTL + q" para preencher a coluna "Considerações ICMS" e "CRTL + m" para preencher a coluna "Considerações PIS/COFINS"
</p>

<img src="Imagens\definindoAtalhoMacro.png">
<img src="Imagens\definindoAtalhoPiseCofins.png">

<p>
Agora todas as vezes que selecionamos uma linha na planilha "Audit.xlsm" e apertamos "CRTL + q" a macro é executada, oque a macro em VBA faz é procurar o produto que foi selecionado na planilha "Audit.xlsm" dentro da planilha "Cliente.xlsm", ao encontrar esse produto compara os dados do mesmo produto em ambas as planilhas e traz um comentário referente a comparação na coluna "Considerações ICMS" dentro da planilha "Cliente.xlsm". Lembrando que todos os produtos que estejam dentro da seleção serão auditados, então quanto mais produtos selecionar, maior será o consumo de memória RAM no momento que a macro for executada.
</p>

<img src="Imagens\macroFuncionandoICMS.gif">

<p>
A mesma lógica serve para auditar o PIS e COFINS do produto, a diferença é que iremos usar o atalho "CRTL + m" e o comentário será inserido na coluna "Considerações PIS/COFINS".
</p>

<img src="Imagens\pisEcofinsAutom.gif">

<p>
A macro em VBA  que preenche a coluna "Considerações ICMS" segue uma lógica para poder inserir o comentário correto, abaixo mostramos a lista dos critérios que devem existir para que os comentários sejam inseridos:
</p>

- Cliente usando CST 00 ou 20 e auditor usando CST 40

Produto se enquadra em isenção

- Cliente usando CST 00 ou 20 e auditor usando CST 60

Produto se enquadra em ST

- Cliente usando CST 20 e auditor não usa nenhum CST

Produto não se enquadra em redução

- Cliente usando CST 40 e auditor usando CST 00, 20 ou 40

Produto não se enquadra em isenção

- Cliente usando CST 60 e auditor usando CST 00, 20 ou 40

Produto não se enquadra em ST

- Cliente usando CST 00 e auditor usando CST 20 com CBENEF de cesta básica

Produto se enquadra em redução para 7%

- Cliente usando CST 00 e auditor usando CST 20 com CBENEF de redução de 19% para 12%

Produto se enquadra em redução de 19% para 12%

- Cliente usando CST 00 e auditor usando CST 00 porém a alíquota do cliente está diferente da alíquota do auditor

Alíquota Incorreta

- Cliente usando alíquota diferente de 21 sem o adicional de protege e auditor usando alíquota 21

Alíquota Incorreta(ICMS 19% + 2% FCP)

- Cliente usando alíquota diferente de 27 sem o adicional de protege e auditor usando alíquota 27

Alíquota Incorreta(ICMS 25% + 2% FCP)

- Cliente usando CST 41 e auditor usando CBENEF que não se enquadra no CST 41

Produto Tributado

- Cliente usando CST 00, 20 ou 40 e auditor usando CBENEF que se enquadra no CST 41

Produto não Tributado

- Quando os dados de CST e alíquota do cliente são iguais aos dados do auditor

Ok Conferido

<p>
A macro em VBA  que preenche a coluna "Considerações PIS/COFINS" segue uma lógica para poder inserir o comentário correto, abaixo mostramos a lista dos critérios que devem existir para que os comentários sejam inseridos:
</p>

- Quando o auditor usa 1 no CST e o cliente usa algum CST diferente de 1

Produto Tributado

- Quando o auditor usa 4 no CST e o cliente usa algum CST diferente de 4

Produto Monofásico

- Quando o auditor usa 5 no CST e o cliente usa algum CST diferente de 5

Substituição Tributária

- Quando o auditor usa 6 no CST e o cliente usa algum CST diferente de 6

Produto Sujeito à Alíquota 0

- Quando o cliente usa um CST igual ao do auditor

Ok Conferido

## 3 CONCLUSÃO

<p>
No final do processo basta retornar a planilha "Cliente.xlsm" para o cliente, pois a mesma estará auditada.
</p>

## 4 FERRAMENTAS UTILIZADAS

- Excel

- Visual Basic for Applications

- Visual Studio Code

- Notepad++

- Git Hub


## 5 COMO ADQUIRIR A AUTOMATIZAÇÃO

<p>
Entrar em contato:

📞(62) 9 9677-8299<br>
📧lucasfonseca108.lf@gmail.com
</p>





