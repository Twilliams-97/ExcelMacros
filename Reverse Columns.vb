Sub ReverseColumns()

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
    For i = 1 To UBound(rngArray, 1)
        k = UBound(rngArray, 2)
        For j = 1 To UBound(rngArray, 2) / 2
            tempRng = rngArray(i, j)
            rngArray(i, j) = rngArray(i, k)
            rngArray(i, k) = tempRng
            k = k - 1
        Next
    Next
    
    'Apply the array
    rng.Formula = rngArray

End Sub
