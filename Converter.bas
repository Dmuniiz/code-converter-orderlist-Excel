Attribute VB_Name = "Módulo1"
Sub Converter()

    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim ultimaLinhaA As Long
    Dim ultimaLinhaE As Long
    Dim ultimaLinha As Long
    
    ultimaLinhaA = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    ultimaLinhaE = ws.Cells(ws.Rows.Count, "E").End(xlUp).Row
    ultimaLinhaF = ws.Cells(ws.Rows.Count, "F").End(xlUp).Row
    
    ' Definir a última linha a ser usada
    ultimaLinha = Application.WorksheetFunction.Max(ultimaLinhaA, ultimaLinhaE, ultimaLinhaF)
    
    ' Copiar a Coluna A para a Coluna B
    ws.Range("F2:F" & ultimaLinha).Copy Destination:=ws.Range("I2:I" & ultimaLinha)
    
    'Define as faixas de células
    Dim rngCodigos As Range
    Dim rngValores As Range
    Dim rngProcurar As Range
    Dim rngPreencher As Range
    Dim i As Long
    
    Set rngCodigos = ws.Range("A2:A" & ultimaLinha)
    Set rngValores = ws.Range("C2:C" & ultimaLinha)
    Set rngProcurar = ws.Range("E2:E" & ultimaLinha)
    Set rngPreencher = ws.Range("H2:H" & ultimaLinha)



' Itera através de cada célula na faixa rngProcurar
    For i = 1 To rngProcurar.Count

' Encontra a correspondência do valor em E na coluna A
    Dim matchIndex As Variant
        matchIndex = Application.Match(rngProcurar.Cells(i, 1).Value, rngCodigos, 0)

' Verifica se uma correspondência foi encontrada
    If Not IsError(matchIndex) Then
' Retorna o valor correspondente na coluna C e preenche na coluna desejada
        rngPreencher.Cells(i, 1).Value = rngValores.Cells(matchIndex, 1).Value
    Else:
        rngPreencher.Cells(i, 1).Value = "Não encontrado"
        End If
    Next i
End Sub

