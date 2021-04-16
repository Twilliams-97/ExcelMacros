Sub MakeTwoColumns()

    Dim rng As Range
    Dim InputRng As Range, OutRng As Range
    
    xTitleId = "Every Other Rower"
    
    Set InputRng = Application.Selection
    Set InputRng = Application.InputBox("Select Range:", xTitleId, InputRng.Address, Type:=8)
    Set OutRng = Application.InputBox("Output to (single cell at top left of new range):", xTitleId, Type:=8)
    Set InputRng = InputRng.Columns(1)
    
    For i = 1 To InputRng.Rows.Count Step 2
        OutRng.Resize(1, 2).Value = Array(InputRng.Cells(i, 1).Value, InputRng.Cells(i + 1, 1).Value)
        Set OutRng = OutRng.Offset(1, 0)
    Next
    
End Sub
        
        Sub MakeThreeColumns()

    Dim rng As Range
    Dim InputRng As Range, OutRng As Range
    'Dim xrow As Integer
    
    xTitleId = "Turn 1 Column to Three"
    
    Set InputRng = Application.Selection
    Set InputRng = Application.InputBox("Select initial values :", xTitleId, InputRng.Address, Type:=8)
    Set OutRng = Application.InputBox("Output to (single cell at top left of new range):", xTitleId, Type:=8)
    Set InputRng = InputRng.Columns(1)
    
    For i = 1 To InputRng.Rows.Count Step 3
        OutRng.Resize(1, 3).Value = Array(InputRng.Cells(i, 1).Value, InputRng.Cells(i + 1, 1).Value, InputRng.Cells(i + 2, 1).Value)
        Set OutRng = OutRng.Offset(1, 0)
    Next
    
End Sub
