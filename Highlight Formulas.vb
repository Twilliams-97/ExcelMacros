Sub HighlightFormulas()

    Dim rng As Range
    
    For Each rng In Cells.SpecialCells(xlCellTypeFormulas)
    
        rng.Interior.ColorIndex = 36
    
    Next rng

End Sub
