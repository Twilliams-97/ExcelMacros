Sub SortSheetsTabName()

    Dim wsCount As Integer
    Dim i As Integer
    Dim j As Integer
    
    wsCount = ActiveWorkbook.Sheets.Count
    
    For i = 1 To wsCount - 1
    
        For j = i + 1 To wsCount
    
            If Sheets(j).Name < Sheets(i).Name Then
    
                Sheets(j).Move before:=Sheets(i)
    
            End If
    
        Next j
    
    Next i

End Sub
