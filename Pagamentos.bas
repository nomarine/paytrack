Attribute VB_Name = "Pagamentos"
Function GetUnidades(ByVal ws As Worksheet)
    Dim titulo As String
    Dim searchRange As Range
    Dim valores As Variant
    Dim unidadeCell As Range
    Dim unidades() As Variant
    Dim ultimaLinha As Long
    
    Set searchRange = ws.UsedRange
    
    titulo = "Un."
    Set unidadeCell = searchRange.Find(titulo)
    
    With ws
        totalRows = .Rows.Count
        unidadeAddress = unidadeCell.Address(ColumnAbsolute = False, RowAbsolute = True)
        unidadeColumn = Left(unidadeAddress, 1)
        ultimaLinha = .Range(unidadeColumn & totalRows).End(xlUp).Row
        valores = ws.Range(unidadeAddress, unidadeColumn & ultimaLinha).Value2
    End With
    
    Dim i As Integer
    i = 0
    For Each valor In valores
        If Not IsEmpty(valor) And StrComp(valor, titulo) Then
            ReDim Preserve unidades(i)
            unidades(i) = valor
            i = i + 1
        End If
    Next valor
    
    GetUnidades = unidades
    
End Function

Function GetDtPagamentos(ByVal ws As Worksheet)
    Dim titulo As String
    Dim searchRange As Range
    Dim valores As Variant
    Dim dtPagamentoCell As Range
    Dim dtPagamentos() As Variant
    Dim ultimaLinha As Long
    
    Set searchRange = ws.UsedRange

    titulo = "Pagto."
    Set dtPagamentoCell = searchRange.Find(titulo)
    
    With ws
        totalRows = .Rows.Count
        dtPagamentoAddress = dtPagamentoCell.Address(ColumnAbsolute = False, RowAbsolute = True)
        dtPagamentoColumn = Left(dtPagamentoAddress, 1)
        ultimaLinha = .Range(dtPagamentoColumn & totalRows).End(xlUp).Row
        valores = ws.Range(dtPagamentoAddress, dtPagamentoColumn & ultimaLinha).Value2
    End With
    
    Dim i As Integer
    i = 0
    For Each valor In valores
        If Not IsEmpty(valor) And (valor Like "[0-9]*") Then
            ReDim Preserve dtPagamentos(i)
            dtPagamentos(i) = valor
            i = i + 1
        End If
    Next valor
    
    GetDtPagamentos = dtPagamentos
    
End Function

Function GetTotais(ByVal ws As Worksheet)
    Dim titulo As String
    Dim searchRange As Range
    Dim valores As Variant
    Dim dtPagamentoCell As Range
    Dim dtPagamentos() As Variant
    Dim totalCell As Range
    Dim totais() As Variant
    Dim ultimaLinha As Long
    
    Set searchRange = ws.UsedRange

    titulo = "Pagto."
    Set dtPagamentoCell = searchRange.Find(titulo)

    tituloTotal = "Total"
    Set totalCell = searchRange.Find(tituloTotal, , , xlWhole)
    
    With ws
        totalRows = .Rows.Count
        
        dtPagamentoAddress = dtPagamentoCell.Address(ColumnAbsolute = False, RowAbsolute = True)
        dtPagamentoColumn = Left(dtPagamentoAddress, 1)
        
        totalAddress = totalCell.Address(ColumnAbsolute = False, RowAbsolute = True)
        totalColumn = Left(totalAddress, 1)
        
        ultimaLinha = .Range(totalColumn & totalRows).End(xlUp).Row
        
        dtPagamentos = ws.Range(dtPagamentoAddress, dtPagamentoColumn & ultimaLinha).Value2
        valores = ws.Range(totalAddress, totalColumn & ultimaLinha).Value2
    End With
    
    Dim i, j As Integer
    i = 1
    j = 0
    For Each dtPagamento In dtPagamentos
        If Not IsEmpty(dtPagamento) And (dtPagamento Like "[0-9]*") Then
            ReDim Preserve totais(j)
            totais(j) = valores(i, 1)
            j = j + 1
        End If
        i = i + 1
    Next dtPagamento
    
    GetTotais = totais
    
End Function

Sub Pagamento()
    Dim filePath As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim unidades As Variant
    Dim totais As Variant
    Dim dtPagamentos As Variant
    
    filePath = "C:\Users\rafae\OneDrive\Documentos\Job\Ana Nunes (Contabilidade)\RECEBIMENTO OUTUBRO DE 2022.xlsx"
    Set wb = Workbooks.Open(filePath)
    Set ws = wb.Worksheets(1)
    ws.Activate
    
    unidades = GetUnidades(ws)
    dtPagamentos = GetDtPagamentos(ws)
    totais = GetTotais(ws)
    
End Sub


    
    
