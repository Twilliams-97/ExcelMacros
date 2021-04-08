Sub AddSuffix()

    Dim c As Range
    Dim suffixValue As Variant
    
    'Display inputbox to collect prefix text
    suffixValue = Application.InputBox(Prompt:="Enter Suffix:", _
        Title:="Suffix", Type:=2)
    
    'The User clicked Cancel
    If suffixValue = False Then Exit Sub
    
        'Loop through each cellin selection
        For Each c In Selection
    
            'Add Suffix where cell is not a formula or blank
            If Not c.HasFormula And c.Value <> "" Then
    
                c.Value = c.Value & suffixValue
    
            End If
    
    Next

End Sub

