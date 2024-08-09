<p align="center">
    <h1 align="center">
        🟦🟧 PREENCHIMENTO/AUDITORIA DE PLANILHA NO EXCEL   AUTOMATIZADO COM VBA 🟨🟪
    </h1>
    <br>
    <img src="Imagens\CapaProjetoVBA.png">
</p>

<h2>
    📑 SUMÁRIO
</h2>

1. [INTRODUÇÃO](#1-introdução)<br>
   1.1 - [Intuito do Projeto](#11-intuito-do-projeto)<br>
   1.2 - [Situação Abordada](#12-situação-abordada)<br>
2. [DESENVOLVIMENTO](#2-desenvolvimento)<br>
3. [CONCLUSÃO](#3-conclusão)
4. [FERRAMENTAS UTILIZADAS](#4-ferramentas-utilizadas)

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

<p>
Abaixo serão disponibilizados os arquivos já prontos para uso, mas também será feita a explicação do zero para que todo o processo possa ser entendido:
</p>

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
  Com as planilhas abertas apertamos "ALT + F11" para abrir o ambiente de desenvolvimento VBA, expandimos as pastas do VBAProject (Cliente.xlsm) e inserimos um novo módulo:
</p>

<img src="Imagens\criandoPrimeiraModulo.png">

<p>
  Copie e cole o código abaixo no módulo criado:
</p>

```vba
Sub ConferirAliquotas()

' Desabilitar atualizações da tela para melhorar a performance
    Application.ScreenUpdating = False
    
    ' Definir variáveis para as planilhas e abas
    Dim wsCliente As Worksheet
    Dim wbAudit As Workbook
    Dim wsExcecoesST As Worksheet
    
    ' Definir variáveis para as colunas
    Dim colCodigoProduto As Long
    Dim colAliquotaEfetICMS As Long
    Dim colConsideracoes As Long
    Dim colCSTICMS As Long
    Dim colCodigoDeBarras As Long
    Dim colAliquota As Long
    Dim colIsencao As Long
    Dim colST As Long
    Dim colCBENEF As Long
    
    ' Definir planilhas e abas
    Set wsCliente = ThisWorkbook.Worksheets("Planilha1")
    Set wbAudit = Workbooks("Audit.xlsm")
    Set wsExcecoesST = wbAudit.Worksheets("Exceções de ST Alíquota e ST")
    
    ' Obter o índice das colunas na planilha "Cliente.xlsm"
    colCodigoProduto = ObterIndiceColunaPorNome(wsCliente, "codigo_produto")
    colAliquotaEfetICMS = ObterIndiceColunaPorNome(wsCliente, "Aliquota_Efet_ICMS")
    colConsideracoes = ObterIndiceColunaPorNome(wsCliente, "Considerações ICMS")
    colCSTICMS = ObterIndiceColunaPorNome(wsCliente, "CST_ICMS")
    
    ' Obter o índice das colunas na planilha "Audit.xlsm"
    colCodigoDeBarras = ObterIndiceColunaPorNome(wsExcecoesST, "Códigodebarras")
    colAliquota = ObterIndiceColunaPorNome(wsExcecoesST, "Alíquota")
    colIsencao = ObterIndiceColunaPorNome(wsExcecoesST, "Isenção")
    colST = ObterIndiceColunaPorNome(wsExcecoesST, "ST?")
    colCBENEF = ObterIndiceColunaPorNome(wsExcecoesST, "CBNEF")
    
    ' Verificar se todas as colunas foram encontradas
    If colCodigoProduto = -1 Or colAliquotaEfetICMS = -1 Or colConsideracoes = -1 Or colCodigoDeBarras = -1 Or colAliquota = -1 Or colCSTICMS = -1 Or colIsencao = -1 Or colST = -1 Or colCBENEF = -1 Then
        MsgBox "Uma ou mais colunas não foram encontradas. Verifique os nomes das colunas.", vbCritical
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    ' Verificar se há células selecionadas na planilha "Audit.xlsm"
    Dim selRange As Range
    Set selRange = Selection
    If selRange Is Nothing Then
        MsgBox "Por favor, selecione ao menos uma célula na planilha 'Audit.xlsm'.", vbExclamation
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    ' Percorrer cada linha selecionada na planilha "Audit.xlsm"
    Dim cel As Range
    For Each cel In selRange.Rows
        Dim linhaAudit As Long
        linhaAudit = cel.Row
        Dim codigoDeBarras As String
        codigoDeBarras = CStr(wsExcecoesST.Cells(linhaAudit, colCodigoDeBarras).Value)
        
        ' Percorrer todas as linhas da planilha "Cliente.xlsm"
        Dim lastRowCliente As Long
        lastRowCliente = wsCliente.Cells(wsCliente.Rows.Count, colCodigoProduto).End(xlUp).Row
        
        Dim i As Long
        For i = 2 To lastRowCliente
            Dim codigoProduto As String
            codigoProduto = CStr(wsCliente.Cells(i, colCodigoProduto).Value)
            
            ' Verificar se o código de barras da planilha "Audit.xlsm" está na planilha "Cliente.xlsm"
            If codigoProduto = codigoDeBarras Then
                ' Comparar alíquotas
                Dim aliquotaEfet As String
                aliquotaEfet = CStr(wsCliente.Cells(i, colAliquotaEfetICMS).Value)
                Dim aliquotaParam As String
                aliquotaParam = CStr(wsExcecoesST.Cells(linhaAudit, colAliquota).Value)
                
                ' Verificar as novas condições
                Dim cstICMS As String
                cstICMS = Trim(CStr(wsCliente.Cells(i, colCSTICMS).Value)) ' Ajuste para tratar espaços em branco
                
                Dim isencaoParam As String
                isencaoParam = LCase(CStr(wsExcecoesST.Cells(linhaAudit, colIsencao).Value))
                
                Dim stParam As String
                stParam = LCase(CStr(wsExcecoesST.Cells(linhaAudit, colST).Value))
                
                Dim cbenefParam As String
                cbenefParam = Trim(CStr(wsExcecoesST.Cells(linhaAudit, colCBENEF).Value))
                
                ' Condições adicionais
                If cbenefParam = "GO822019" Then
                    wsCliente.Cells(i, colConsideracoes).Value = "Produto de padaria, lanchonete ou confeitaria, se enquadra em redução para 7% de acordo com o Parecer."
                ElseIf (cstICMS = "0" Or cstICMS = "20" Or cstICMS = "40" Or cstICMS = "") And cbenefParam = "GO800004" Then
                    wsCliente.Cells(i, colConsideracoes).Value = "Produto não tributado"
                ElseIf (cstICMS = "0" Or cstICMS = "") And cbenefParam = "GO821022" Then
                    wsCliente.Cells(i, colConsideracoes).Value = "Produto se enquadra em redução de 19% para 12%"
                ElseIf (cstICMS = "0" Or cstICMS = "" Or cstICMS = "20") And stParam = "st" Then
                    wsCliente.Cells(i, colConsideracoes).Value = "Produto se enquadra em ST"
                ElseIf (cstICMS = "0" Or cstICMS = "" Or cstICMS = "20" Or cstICMS = "") And isencaoParam = "isenção" Then
                    wsCliente.Cells(i, colConsideracoes).Value = "Produto se enquadra em isenção"
                ElseIf cstICMS = "20" And stParam = "nãost" And isencaoParam = "sem isenção" And cbenefParam = "" Then
                    wsCliente.Cells(i, colConsideracoes).Value = "Produto não se enquadra em redução"
                ElseIf cstICMS = "40" And isencaoParam <> "isenção" Then
                    wsCliente.Cells(i, colConsideracoes).Value = "Produto não se enquadra em isenção"
                ElseIf cstICMS = "60" And stParam <> "st" Then
                    wsCliente.Cells(i, colConsideracoes).Value = "Produto não se enquadra em ST"
                ElseIf (cstICMS = "0" Or cstICMS = "") And stParam = "red" And (cbenefParam = "GO821019" Or cbenefParam = "GO821010") Then
                    wsCliente.Cells(i, colConsideracoes).Value = "Produto se enquadra em redução para 7%"
                ElseIf (cstICMS = "0" Or cstICMS = "") And stParam = "red" And cbenefParam = "GO821008" Then
                    wsCliente.Cells(i, colConsideracoes).Value = "Produto se enquadra em redução de 19% para 7%"
                ElseIf (cstICMS = "0" Or cstICMS = "") And stParam = "red" And cbenefParam = "GO821020" Then
                    wsCliente.Cells(i, colConsideracoes).Value = "Produto se enquadra em redução de 19% para 9%"
                ElseIf (cstICMS = "0" Or cstICMS = "") And cbenefParam = "GO821022" Then
                    wsCliente.Cells(i, colConsideracoes).Value = "Produto se enquadra em redução de 19% para 12%"
                ElseIf (cstICMS = "0" Or cstICMS = "") And (aliquotaEfet <> "21") And aliquotaParam = "21" Then
                    wsCliente.Cells(i, colConsideracoes).Value = "Alíquota Incorreta (ICMS 19% + 2% FCP)"
                ElseIf (cstICMS = "0" Or cstICMS = "") And (aliquotaEfet <> "27") And (aliquotaParam = "27") Then
                    wsCliente.Cells(i, colConsideracoes).Value = "Alíquota Incorreta (ICMS 25% + 2% FCP)"
                ElseIf cstICMS = "40" And aliquotaEfet <> "0" And isencaoParam = "isenção" Then
                    wsCliente.Cells(i, colConsideracoes).Value = "Produto se enquadra em isenção"
                ElseIf cstICMS = "60" And aliquotaEfet <> "0" And stParam = "st" Then
                    wsCliente.Cells(i, colConsideracoes).Value = "Produto se enquadra em ST"
                ElseIf aliquotaEfet = aliquotaParam Then
                    wsCliente.Cells(i, colConsideracoes).Value = "Ok Conferido"
                ElseIf cstICMS = "41" And (stParam = "red" Or stParam = "nãost") And isencaoParam <> "isenção" Then
                    wsCliente.Cells(i, colConsideracoes).Value = "Produto Tributado"
                Else
                    ' Se nenhuma condição específica for atendida, manter como "Ok Conferido" ou "Alíquota Incorreta"
                    If aliquotaEfet = "21" And aliquotaParam = "21" Then
                        wsCliente.Cells(i, colConsideracoes).Value = "Ok Conferido"
                    ElseIf aliquotaEfet = "27" And aliquotaParam = "27" Then
                        wsCliente.Cells(i, colConsideracoes).Value = "Ok Conferido"
                    Else
                        wsCliente.Cells(i, colConsideracoes).Value = "Alíquota Incorreta"
                    End If
                End If
            End If
        Next i
    Next cel
    
    ' Reativar atualizações da tela
    Application.ScreenUpdating = True

End Sub

' Função auxiliar para obter o índice da coluna por nome
Function ObterIndiceColunaPorNome(ws As Worksheet, nomeColuna As String) As Long
    Dim cel As Range
    For Each cel In ws.Rows(1).Cells
        If cel.Value = nomeColuna Then
            ObterIndiceColunaPorNome = cel.Column
            Exit Function
        End If
    Next cel
    ObterIndiceColunaPorNome = -1
End Function
```
<p>
  Crie outro módulo
</p>

<img src="Imagens\criandoSegundoModulo.png">

<p>
    Cole o código abaixo:
</p>

```vba
Sub ConferirPis()

    ' Desabilitar atualizações da tela para melhorar a performance
    Application.ScreenUpdating = False
    
    ' Definir variáveis para as planilhas e abas
    Dim wsCliente As Worksheet
    Dim wbAudit As Workbook
    Dim wsExcecoesPIS As Worksheet
    
    ' Definir variáveis para as colunas
    Dim colCodigoProduto As Long
    Dim colcstpis As Long ' CST_PIS
    Dim colConsideracoespis As Long ' Considerações PIS/COFINS
    
    Dim colCodigoDeBarras As Long ' Codigo de Barras
    Dim cstPisParam As Long
    
    ' Definir planilhas e abas
    Set wsCliente = ThisWorkbook.Worksheets("Planilha1")
    Set wbAudit = Workbooks("Audit.xlsm")
    Set wsExcecoesPIS = wbAudit.Worksheets("Exceções PIS Cofins Aliq 0")
    
    ' Obter o índice das colunas na planilha "Cliente.xlsm"
    colCodigoProduto = ObterIndiceColunaPorNome(wsCliente, "codigo_produto")
    colcstpis = ObterIndiceColunaPorNome(wsCliente, "CST_PIS")
    colConsideracoespis = ObterIndiceColunaPorNome(wsCliente, "Considerações PIS/COFINS")
    
    ' Obter o índice das colunas na planilha "Audit.xlsm"
    colCodigoDeBarras = ObterIndiceColunaPorNome(wsExcecoesPIS, "CodBarras2") ' CodBarras2 ao invês de Codigo de Barras
    cstPisParam = ObterIndiceColunaPorNome(wsExcecoesPIS, "CST")
    
    ' Verificar se todas as colunas foram encontradas
    If colCodigoProduto = -1 Or colcstpis = -1 Or colConsideracoespis = -1 Or colCodigoDeBarras = -1 Or cstPisParam = -1 Then
        MsgBox "Uma ou mais colunas não foram encontradas. Verifique os nomes das colunas.", vbCritical
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    ' Verificar se há células selecionadas na planilha "Audit.xlsm"
    Dim selRange As Range
    Set selRange = Selection
    If selRange Is Nothing Then
        MsgBox "Por favor, selecione ao menos uma célula na planilha 'Audit.xlsm'.", vbExclamation
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    ' Percorrer cada linha selecionada na planilha "Audit.xlsm"
    Dim cel As Range
    For Each cel In selRange.Rows
        Dim linhaAudit As Long
        linhaAudit = cel.Row
        Dim codigoDeBarras As String
        codigoDeBarras = CStr(wsExcecoesPIS.Cells(linhaAudit, colCodigoDeBarras).Value)
        
        ' Percorrer todas as linhas da planilha "Cliente.xlsm"
        Dim lastRowCliente As Long
        lastRowCliente = wsCliente.Cells(wsCliente.Rows.Count, colCodigoProduto).End(xlUp).Row
        
        Dim i As Long
        For i = 2 To lastRowCliente
            Dim codigoProduto As String
            codigoProduto = CStr(wsCliente.Cells(i, colCodigoProduto).Value)
            
            ' Verificar se o código de barras da planilha "Audit.xlsm" está na planilha "Cliente.xlsm"
        If codigoProduto = codigoDeBarras Then
                ' Comparar pis
                Dim pis As String
                pis = CStr(wsCliente.Cells(i, colcstpis).Value)
                Dim pisParam As String
                pisParam = CStr(wsExcecoesPIS.Cells(linhaAudit, cstPisParam).Value)
                
                ' Condição: se o pisParam for igual ao pis
            If pisParam = pis Then
                ' Inserir na coluna Considerações PIS/COFINS o valor "Ok Conferido"
                wsCliente.Cells(i, colConsideracoespis).Value = "Ok Conferido"
                
                ' Condição: se o pisParam for igual a 1 e o pis for diferente de 1
                ElseIf pisParam = "1" And pis <> "1" Then
                ' Inserir na coluna Considerações PIS/COFINS o valor "Produto Tributado"
                wsCliente.Cells(i, colConsideracoespis).Value = "Produto Tributado"
            
                ' Condição: se o pisParam for igual a 4 e o pis for diferente de 4
                ElseIf pisParam = "4" And pis <> "4" Then
                ' Inserir na coluna Considerações PIS/COFINS o valor "Produto Monofásico"
                wsCliente.Cells(i, colConsideracoespis).Value = "Produto Monofásico"
                
                ' Condição: se o pisParam for igual a 5 e o pis for diferente de 5
                ElseIf pisParam = "5" And pis <> "5" Then
                ' Inserir na coluna Considerações PIS/COFINS o valor "Substituição Tributária"
                wsCliente.Cells(i, colConsideracoespis).Value = "Substituição Tributária"
                
                ' Condição: se o pisParam for igual a 6 e o pis for diferente de 6
                ElseIf pisParam = "6" And pis <> "6" Then
                ' Inserir na coluna Considerações PIS/COFINS o valor "Produto Sujeito à Alíquota Zero"
                wsCliente.Cells(i, colConsideracoespis).Value = "Produto Sujeito à Alíquota Zero"
                
            End If
        End If
        
         Next i
    Next cel
    
    ' Reativar atualizações da tela
    Application.ScreenUpdating = True

End Sub

' Função auxiliar para obter o índice da coluna por nome
Function ObterIndiceColunaPorNome(ws As Worksheet, nomeColuna As String) As Long
    Dim cel As Range
    For Each cel In ws.Rows(1).Cells
        If cel.Value = nomeColuna Then
            ObterIndiceColunaPorNome = cel.Column
            Exit Function
        End If
    Next cel
    ObterIndiceColunaPorNome = -1
End Function
```

<p>
  Salve as macros criadas e feche o VBA
</p>

<img src="Imagens\salvandoMacros.png">
<img src="Imagens\fechandoVBA.png">

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
No final do processo basta retornar a planilha "Cliente.xlsm" auditada para o cliente.
</p>

## 4 FERRAMENTAS UTILIZADAS

- Excel

- Visual Basic for Applications

- Visual Studio Code

- Notepad++

- Git Hub






