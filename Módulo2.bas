Attribute VB_Name = "M�dulo2"
Sub ConferirPis()

    ' Desabilitar atualiza��es da tela para melhorar a performance
    Application.ScreenUpdating = False
    
    ' Definir vari�veis para as planilhas e abas
    Dim wsCliente As Worksheet
    Dim wbAudit As Workbook
    Dim wsExcecoesPIS As Worksheet
    
    ' Definir vari�veis para as colunas
    Dim colCodigoProduto As Long
    Dim colcstpis As Long ' CST_PIS
    Dim colConsideracoespis As Long ' Considera��es PIS/COFINS
    
    Dim colCodigoDeBarras As Long ' Codigo de Barras
    Dim cstPisParam As Long
    
    ' Definir planilhas e abas
    Set wsCliente = ThisWorkbook.Worksheets("Planilha1")
    Set wbAudit = Workbooks("Audit.xlsm")
    Set wsExcecoesPIS = wbAudit.Worksheets("Exce��es PIS Cofins Aliq 0")
    
    ' Obter o �ndice das colunas na planilha "Cliente.xlsm"
    colCodigoProduto = ObterIndiceColunaPorNome(wsCliente, "codigo_produto")
    colcstpis = ObterIndiceColunaPorNome(wsCliente, "CST_PIS")
    colConsideracoespis = ObterIndiceColunaPorNome(wsCliente, "Considera��es PIS/COFINS")
    
    ' Obter o �ndice das colunas na planilha "Audit.xlsm"
    colCodigoDeBarras = ObterIndiceColunaPorNome(wsExcecoesPIS, "CodBarras2") ' CodBarras2 ao inv�s de Codigo de Barras
    cstPisParam = ObterIndiceColunaPorNome(wsExcecoesPIS, "CST")
    
    ' Verificar se todas as colunas foram encontradas
    If colCodigoProduto = -1 Or colcstpis = -1 Or colConsideracoespis = -1 Or colCodigoDeBarras = -1 Or cstPisParam = -1 Then
        MsgBox "Uma ou mais colunas n�o foram encontradas. Verifique os nomes das colunas.", vbCritical
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    ' Verificar se h� c�lulas selecionadas na planilha "Audit.xlsm"
    Dim selRange As Range
    Set selRange = Selection
    If selRange Is Nothing Then
        MsgBox "Por favor, selecione ao menos uma c�lula na planilha 'Audit.xlsm'.", vbExclamation
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
            
            ' Verificar se o c�digo de barras da planilha "Audit.xlsm" est� na planilha "Cliente.xlsm"
        If codigoProduto = codigoDeBarras Then
                ' Comparar pis
                Dim pis As String
                pis = CStr(wsCliente.Cells(i, colcstpis).Value)
                Dim pisParam As String
                pisParam = CStr(wsExcecoesPIS.Cells(linhaAudit, cstPisParam).Value)
                
                ' Condi��o: se o pisParam for igual ao pis
            If pisParam = pis Then
                ' Inserir na coluna Considera��es PIS/COFINS o valor "Ok Conferido"
                wsCliente.Cells(i, colConsideracoespis).Value = "Ok Conferido"
                
                ' Condi��o: se o pisParam for igual a 1 e o pis for diferente de 1
                ElseIf pisParam = "1" And pis <> "1" Then
                ' Inserir na coluna Considera��es PIS/COFINS o valor "Produto Tributado"
                wsCliente.Cells(i, colConsideracoespis).Value = "Produto Tributado"
            
                ' Condi��o: se o pisParam for igual a 4 e o pis for diferente de 4
                ElseIf pisParam = "4" And pis <> "4" Then
                ' Inserir na coluna Considera��es PIS/COFINS o valor "Produto Monof�sico"
                wsCliente.Cells(i, colConsideracoespis).Value = "Produto Monof�sico"
                
                ' Condi��o: se o pisParam for igual a 5 e o pis for diferente de 5
                ElseIf pisParam = "5" And pis <> "5" Then
                ' Inserir na coluna Considera��es PIS/COFINS o valor "Substitui��o Tribut�ria"
                wsCliente.Cells(i, colConsideracoespis).Value = "Substitui��o Tribut�ria"
                
                ' Condi��o: se o pisParam for igual a 6 e o pis for diferente de 6
                ElseIf pisParam = "6" And pis <> "6" Then
                ' Inserir na coluna Considera��es PIS/COFINS o valor "Produto Sujeito � Al�quota Zero"
                wsCliente.Cells(i, colConsideracoespis).Value = "Produto Sujeito � Al�quota Zero"
                
            End If
        End If
        
         Next i
    Next cel
    
    ' Reativar atualiza��es da tela
    Application.ScreenUpdating = True

End Sub

' Fun��o auxiliar para obter o �ndice da coluna por nome
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
