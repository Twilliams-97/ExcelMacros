Sub CapitaliseCells()

    Dim rng As Range
    Dim c As Range
    Dim letter As String
    Dim finalString As String
    Dim i As Integer
    
    Set rng = Selection
    
    'loop through each cell in selection
    For Each c In rng
    
    finalString = ""
    
    'loop through each letter in next string
    For i = 1 To Len(c)
    
        letter = Mid(c.Value, i, 1)
        letter = UCase(letter)

        'Add the letter onto the new text string
        finalString = finalString & letter
    
        Next i
    
        'Return the value back to cell
        c.Value = finalString
    
    Next c

End Sub
