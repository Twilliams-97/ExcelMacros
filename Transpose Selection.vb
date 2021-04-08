Sub TransposeSelection()

    'Create variables
    Dim rng As Range
    Dim rngArray As Variant
    Dim i As Long
    Dim j As Long
    Dim overflowRng As Range
    Dim msgAns As Long
    
    'Record the selected range and it's contents
    Set rng = Selection
    rngArray = rng.Formula
    
    'Test the range and identify if any cells will be overwritten
    If rng.Rows.Count > rng.Columns.Count Then
    
        Set overflowRng = rng.Cells(1, 1). _
            Offset(0, rng.Columns.Count). _
            Resize(rng.Columns.Count, _
            rng.Rows.Count - rng.Columns.Count)
    
    ElseIf rng.Rows.Count < rng.Columns.Count Then
    
        Set overflowRng = rng.Cells(1, 1).Offset(rng.Rows.Count, 0). _
            Resize(rng.Columns.Count - rng.Rows.Count, rng.Rows.Count)
    
    End If
    
    If rng.Rows.Count <> rng.Columns.Count Then
    
        If Application.WorksheetFunction.CountA(overflowRng) > 0 Then
    
            msgAns = MsgBox("Worksheet data in " & overflowRng.Address & _
                " will be overwritten." & vbNewLine & _
                "Do you wish to continue?", vbYesNo)
    
        If msgAns = vbNo Then Exit Sub
    
        End If
    
    End If
    
    'Clear the rnage
    rng.Clear
    
    'Reapply the cells in transposted position
    For i = 1 To UBound(rngArray, 1)
    
        For j = 1 To UBound(rngArray, 2)
    
            rng.Cells(1, 1).Offset(j - 1, i - 1) = rngArray(i, j)
    
        Next
    
    Next

End Sub

