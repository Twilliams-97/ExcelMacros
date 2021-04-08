Sub AddPrefix()

    Dim c As Range
    Dim prefixValue As Variant
    
    'Display inputbox to collect prefix text
    prefixValue = Application.InputBox(Prompt:="Enter prefix:", _
        Title:="Prefix", Type:=2)
    
    'The User clicked Cancel
    If prefixValue = False Then Exit Sub
    
    For Each c In Selection
    
        'Add prefix where cell is not a formula or blank
        If Not c.HasFormula And c.Value <> "" Then
    
            c.Value = prefixValue & c.Value
    
        End If
    
    Next

End Sub
