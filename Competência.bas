Attribute VB_Name = "Competência"
Function FindUnidade(ByVal ws As Worksheet, ByVal unidade As String) As Range
    Dim unidadeCell As Range
    
    ws.Activate
    
    Set unidadeCell = ws.Cells.Find("APTO")
    Set unidadeCell = unidadeCell.Offset(1, 0)
    Set unidadeCell = unidadeCell.EntireRow.Find(unidade, LookIn:=xlFormulas)
    
    If Not unidadeCell Is Nothing Then
        Set FindUnidade = unidadeCell.Offset(1, 0)
    Else
        Set FindUnidade = Nothing
    End If
    
End Function

Function AssignRecebido(ByVal ws As Worksheet, ByVal unidade As String, ByVal valor As Double, ByVal dtPagamento As String)
    Dim competencia As Range
    Dim recebidoCell As Range
    Dim unidadeCell As Range
    Dim unidadeRecebido As Range
    
    ws.Activate
    
    Set competencia = FindCompetencia(ws, unidade, DateValue(dtPagamento))
    Set competencia = competencia.Offset(0, 1)
    
    Set recebidoCell = competencia.EntireColumn.Find(What:="Recebido", After:=competencia, SearchOrder:=xlByColumns)
    Set unidadeCell = FindUnidade(ws, unidade)
    
    Set unidadeRecebido = ws.Cells(recebidoCell.Row, unidadeCell.Column)
    
    unidadeRecebido.Value = unidadeRecebido.Value + valor
    
End Function

Function FindCompetencia(ByVal ws As Worksheet, ByVal unidade As String, ByVal dtPagamento As String) As Range
    Dim competenciaCell As Range
    Dim competencia As String

    Set competenciaCell = ws.Cells.Find(DateValue(dtPagamento), LookIn:=xlFormulas)
    
    If Not competenciaCell Is Nothing Then
        Set FindCompetencia = competenciaCell
    Else
        Set FindCompetencia = Nothing
    End If
    
    
    
End Function

Sub competencia()
    Dim planilhaPagamento, planilhaCompetencia As Workbook
    Dim abaPagamento, abaCompetencia, novaPlanilha As Worksheet
    Dim unidades, dtPagamentos, totais As Variant
    Dim hasUnidade, hasCompetencia, celulaUnidade, celulaDtPagamento, celulaValor As Range
    
    'filePath = "C:\Users\rafae\OneDrive\Documentos\Job\Ana Nunes (Contabilidade)\RECEBIMENTO OUTUBRO DE 2022.xlsx"
    filePath = Application.GetOpenFilename(FileFilter:="Excel Files (*.xls*, *.xlsx*, *.xlsm*, *.xlsb*),*.xls*,*.xlsx*,*.xlsm*,*.xlsb")
    Set planilhaPagamento = Workbooks.Open(filePath)
    Set planilhaCompetencia = ThisWorkbook
    
    Set abaPagamento = planilhaPagamento.Worksheets(1)
    Set abaCompetencia = planilhaCompetencia.Worksheets(LCase("Clientes"))
    
    unidades = Pagamentos.GetUnidades(abaPagamento)
    dtPagamentos = Pagamentos.GetDtPagamentos(abaPagamento)
    totais = Pagamentos.GetTotais(abaPagamento)
    
    i = 0
    j = 0
    Set novaPlanilha = planilhaCompetencia.Sheets.Add(After:=planilhaCompetencia.Worksheets(planilhaCompetencia.Worksheets.Count))
    novaPlanilha.Name = Format(DateTime.Now, "ddMMyyyy_hhmmss")
    For Each valor In totais
        Set hasUnidade = FindUnidade(abaCompetencia, unidades(i))
        Set hasCompetencia = FindCompetencia(abaCompetencia, unidades(i), "01/" & Format(dtPagamentos(i), "mm/yyyy"))
        
        Set celulaUnidade = novaPlanilha.Cells(i + 1, 1)
        Set celulaDtPagamento = novaPlanilha.Cells(i + 1, 2)
        Set celulaValor = novaPlanilha.Cells(i + 1, 3)
        Set resultado = novaPlanilha.Cells(i + 1, 4)
        
        celulaUnidade.NumberFormat = "@"
        celulaUnidade.Value = unidades(i)
        celulaDtPagamento.Value = Format(dtPagamentos(i), "dd/mm/yyyy")
        celulaValor.Value = valor
                
        If hasUnidade Is Nothing Then
                resultado.Value = "Unidade não encontrada"
                'Next valor
        ElseIf hasCompetencia Is Nothing Then
                resultado.Value = "Competência não encontrada"
                'Next valor
        Else
                Call AssignRecebido(abaCompetencia, unidades(i), valor, "01/" & Format(dtPagamentos(i), "mm/yyyy"))
                resultado.Value = "OK"
                
                j = j + 1
        End If
        i = i + 1
    Next valor
        

End Sub
