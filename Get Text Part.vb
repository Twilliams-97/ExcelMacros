Function GetText(CellRef As String)

    Dim StringLength As Integer
    
    StringLength = Len(CellRef)
    
    For i = 1 To StringLength
    
        If Not (IsNumeric(Mid(CellRef, i, 1))) Then Result = Result & Mid(CellRef, i, 1)
    
    Next i
    
    GetText = Result
    
End Function
