Sub ClearSelectedData()

    'Clear all hardcoded values in the selected range
    'Ignores formulas etc
    Selection.SpecialCells(xlCellTypeConstants).ClearContents

End Sub
