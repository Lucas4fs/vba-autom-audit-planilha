Attribute VB_Name = "M�dulo1"
Sub ConferirAliquotas()

' Desabilitar atualiza��es da tela para melhorar a performance
    Application.ScreenUpdating = False
    
    ' Definir vari�veis para as planilhas e abas
    Dim wsCliente As Worksheet
    Dim wbAudit As Workbook
    Dim wsExcecoesST As Worksheet
    
    ' Definir vari�veis para as colunas
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
    Set wsExcecoesST = wbAudit.Worksheets("Exce��es de ST Al�quota e ST")
    
    ' Obter o �ndice das colunas na planilha "Cliente.xlsm"
    colCodigoProduto = ObterIndiceColunaPorNome(wsCliente, "codigo_produto")
    colAliquotaEfetICMS = ObterIndiceColunaPorNome(wsCliente, "Aliquota_Efet_ICMS")
    colConsideracoes = ObterIndiceColunaPorNome(wsCliente, "Considera��es ICMS")
    colCSTICMS = ObterIndiceColunaPorNome(wsCliente, "CST_ICMS")
    
    ' Obter o �ndice das colunas na planilha "Audit.xlsm"
    colCodigoDeBarras = ObterIndiceColunaPorNome(wsExcecoesST, "C�digodebarras")
    colAliquota = ObterIndiceColunaPorNome(wsExcecoesST, "Al�quota")
    colIsencao = ObterIndiceColunaPorNome(wsExcecoesST, "Isen��o")
    colST = ObterIndiceColunaPorNome(wsExcecoesST, "ST?")
    colCBENEF = ObterIndiceColunaPorNome(wsExcecoesST, "CBNEF")
    
    ' Verificar se todas as colunas foram encontradas
    If colCodigoProduto = -1 Or colAliquotaEfetICMS = -1 Or colConsideracoes = -1 Or colCodigoDeBarras = -1 Or colAliquota = -1 Or colCSTICMS = -1 Or colIsencao = -1 Or colST = -1 Or colCBENEF = -1 Then
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
        codigoDeBarras = CStr(wsExcecoesST.Cells(linhaAudit, colCodigoDeBarras).Value)
        
        ' Percorrer todas as linhas da planilha "Cliente.xlsm"
        Dim lastRowCliente As Long
        lastRowCliente = wsCliente.Cells(wsCliente.Rows.Count, colCodigoProduto).End(xlUp).Row
        
        Dim i As Long
        For i = 2 To lastRowCliente
            Dim codigoProduto As String
            codigoProduto = CStr(wsCliente.Cells(i, colCodigoProduto).Value)
            
            ' Verificar se o c�digo de barras da planilha "Audit.xlsm" est� na planilha "Cliente.xlsm"
            If codigoProduto = codigoDeBarras Then
                ' Comparar al�quotas
                Dim aliquotaEfet As String
                aliquotaEfet = CStr(wsCliente.Cells(i, colAliquotaEfetICMS).Value)
                Dim aliquotaParam As String
                aliquotaParam = CStr(wsExcecoesST.Cells(linhaAudit, colAliquota).Value)
                
                ' Verificar as novas condi��es
                Dim cstICMS As String
                cstICMS = Trim(CStr(wsCliente.Cells(i, colCSTICMS).Value)) ' Ajuste para tratar espa�os em branco
                
                Dim isencaoParam As String
                isencaoParam = LCase(CStr(wsExcecoesST.Cells(linhaAudit, colIsencao).Value))
                
                Dim stParam As String
                stParam = LCase(CStr(wsExcecoesST.Cells(linhaAudit, colST).Value))
                
                Dim cbenefParam As String
                cbenefParam = Trim(CStr(wsExcecoesST.Cells(linhaAudit, colCBENEF).Value))
                
                ' Condi��es adicionais
                If cbenefParam = "GO822019" Then
                    wsCliente.Cells(i, colConsideracoes).Value = "Produto de padaria, lanchonete ou confeitaria, se enquadra em redu��o para 7% de acordo com o Parecer."
                ElseIf (cstICMS = "0" Or cstICMS = "20" Or cstICMS = "40" Or cstICMS = "") And cbenefParam = "GO800004" Then
                    wsCliente.Cells(i, colConsideracoes).Value = "Produto n�o tributado"
                ElseIf (cstICMS = "0" Or cstICMS = "") And cbenefParam = "GO821022" Then
                    wsCliente.Cells(i, colConsideracoes).Value = "Produto se enquadra em redu��o de 19% para 12%"
                ElseIf (cstICMS = "0" Or cstICMS = "" Or cstICMS = "20") And stParam = "st" Then
                    wsCliente.Cells(i, colConsideracoes).Value = "Produto se enquadra em ST"
                ElseIf (cstICMS = "0" Or cstICMS = "" Or cstICMS = "20" Or cstICMS = "") And isencaoParam = "isen��o" Then
                    wsCliente.Cells(i, colConsideracoes).Value = "Produto se enquadra em isen��o"
                ElseIf cstICMS = "20" And stParam = "n�ost" And isencaoParam = "sem isen��o" And cbenefParam = "" Then
                    wsCliente.Cells(i, colConsideracoes).Value = "Produto n�o se enquadra em redu��o"
                ElseIf cstICMS = "40" And isencaoParam <> "isen��o" Then
                    wsCliente.Cells(i, colConsideracoes).Value = "Produto n�o se enquadra em isen��o"
                ElseIf cstICMS = "60" And stParam <> "st" Then
                    wsCliente.Cells(i, colConsideracoes).Value = "Produto n�o se enquadra em ST"
                ElseIf (cstICMS = "0" Or cstICMS = "") And stParam = "red" And (cbenefParam = "GO821019" Or cbenefParam = "GO821010") Then
                    wsCliente.Cells(i, colConsideracoes).Value = "Produto se enquadra em redu��o para 7%"
                ElseIf (cstICMS = "0" Or cstICMS = "") And stParam = "red" And cbenefParam = "GO821008" Then
                    wsCliente.Cells(i, colConsideracoes).Value = "Produto se enquadra em redu��o de 19% para 7%"
                ElseIf (cstICMS = "0" Or cstICMS = "") And stParam = "red" And cbenefParam = "GO821020" Then
                    wsCliente.Cells(i, colConsideracoes).Value = "Produto se enquadra em redu��o de 19% para 9%"
                ElseIf (cstICMS = "0" Or cstICMS = "") And cbenefParam = "GO821022" Then
                    wsCliente.Cells(i, colConsideracoes).Value = "Produto se enquadra em redu��o de 19% para 12%"
                ElseIf (cstICMS = "0" Or cstICMS = "") And (aliquotaEfet <> "21") And aliquotaParam = "21" Then
                    wsCliente.Cells(i, colConsideracoes).Value = "Al�quota Incorreta (ICMS 19% + 2% FCP)"
                ElseIf (cstICMS = "0" Or cstICMS = "") And (aliquotaEfet <> "27") And (aliquotaParam = "27") Then
                    wsCliente.Cells(i, colConsideracoes).Value = "Al�quota Incorreta (ICMS 25% + 2% FCP)"
                ElseIf cstICMS = "40" And aliquotaEfet <> "0" And isencaoParam = "isen��o" Then
                    wsCliente.Cells(i, colConsideracoes).Value = "Produto se enquadra em isen��o"
                ElseIf cstICMS = "60" And aliquotaEfet <> "0" And stParam = "st" Then
                    wsCliente.Cells(i, colConsideracoes).Value = "Produto se enquadra em ST"
                ElseIf aliquotaEfet = aliquotaParam Then
                    wsCliente.Cells(i, colConsideracoes).Value = "Ok Conferido"
                ElseIf cstICMS = "41" And (stParam = "red" Or stParam = "n�ost") And isencaoParam <> "isen��o" Then
                    wsCliente.Cells(i, colConsideracoes).Value = "Produto Tributado"
                Else
                    ' Se nenhuma condi��o espec�fica for atendida, manter como "Ok Conferido" ou "Al�quota Incorreta"
                    If aliquotaEfet = "21" And aliquotaParam = "21" Then
                        wsCliente.Cells(i, colConsideracoes).Value = "Ok Conferido"
                    ElseIf aliquotaEfet = "27" And aliquotaParam = "27" Then
                        wsCliente.Cells(i, colConsideracoes).Value = "Ok Conferido"
                    Else
                        wsCliente.Cells(i, colConsideracoes).Value = "Al�quota Incorreta"
                    End If
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
