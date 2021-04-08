Function FindLastRow(ColumnLetter As String)
    
    'Gives you the last cell with data in the specified row
    'Will not work correctly if the last row is hidden
    
    'Enter into cell as =FindLastRow("Column Letter")
    
    Dim ColumnNumber As Long

    With ActiveSheet
    
        'Converts column letter to column number
        ColumnNumber = Range(ColumnLetter & 1).Column
        
        'Finds the last row in the column
        FindLastRow = .Cells(.Rows.Count, ColumnNumber).End(xlUp).row

    End With
    
End Function
