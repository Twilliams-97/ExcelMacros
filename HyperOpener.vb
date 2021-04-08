Sub HyperOpener()

    'Create variable to hold cells in the worksheet
    Dim c As Range
    
    'Loop through each cell in selection
    For Each c In Selection
        
        If c.Hyperlinks.Count > 0 Then
                    For i = 1 To c.Hyperlinks.Count
                        c.Hyperlinks(i).Follow
                    Next i
                    
        ElseIf InStr(1, c.Formula, "HYPERLINK", vbTextCompare) > 0 Then
                If InStr(1, c.Formula, ",") > 0 Then
                    Url = Evaluate(Left(c.Formula, InStr(1, c.Formula, ",") - 1) & ")")
                Else
                    Url = c.Value
                End If
                
            ActiveWorkbook.FollowHyperlink Address:=Url
        End If
            
    Next c
        

End Sub
