Sub ApplySentenceCase()

    Dim rng As Range
    Dim c As Range
    Dim letter As String
    Dim capitalize As Boolean
    Dim finalString As String
    Dim i As Integer
    
    Set rng = Selection
    
    'loop through each cell in selection
    For Each c In rng
    
    finalString = ""
    
    capitalize = True
    
    'loop through each letter in next string
    For i = 1 To Len(c)
    
        letter = Mid(c.Value, i, 1)
    
        'If letter is a period, then turn on capitalize switch
        If letter = "." Then capitalize = True
    
            'If capitalize switch is on, then make upper case
            If capitalize = True Then
    
                letter = UCase(letter)
    
                'Turn off capitalize switch if capital found
                    If letter >= "A" And letter <= "Z" Then
                        capitalize = False
                    End If
    
                'If letter is not to be capitalized, then make lower case
            Else
                letter = LCase(letter)
            End If
    
            'Add the letter onto the new text string
            finalString = finalString & letter
    
        Next i
    
        'Return the value back to cell
        c.Value = finalString
    
    Next c

End Sub
