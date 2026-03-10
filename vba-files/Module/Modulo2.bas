Attribute VB_Name = "Modulo2"
Function BuscarColumna(ws As Worksheet, nombreCol As String) As Long

    Dim lastCol As Long
    Dim i As Long
    
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    For i = 1 To lastCol
        If Trim(UCase(ws.Cells(1, i).Value)) = Trim(UCase(nombreCol)) Then
            BuscarColumna = i
            Exit Function
        End If
    Next i
    
    BuscarColumna = 0

End Function

