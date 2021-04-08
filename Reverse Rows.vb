Sub ReverseRows()
    
    'Create variables
    Dim rng As Range
    Dim rngArray As Variant
    Dim tempRng As Variant
    Dim i As Long
    Dim j As Long
    Dim k As Long
    
    'Record the selected range and it's contents
    Set rng = Selection
    rngArray = rng.Formula
    
    'Loop through all cells and create a temporary array
    For j = 1 To UBound(rngArray, 2)
        k = UBound(rngArray, 1)
        For i = 1 To UBound(rngArray, 1) / 2
            tempRng = rngArray(i, j)
            rngArray(i, j) = rngArray(k, j)
            rngArray(k, j) = tempRng
            k = k - 1
        Next
    Next
    
    'Apply the array
    rng.Formula = rngArray

End Sub
