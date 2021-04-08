Sub InvertSheetSelection()

    'Create variable to hold list of selected worksheet
    Dim selectedList As String
    
    'Create variable to hold worksheets
    Dim ws As Worksheet
    
    'Create variable to switch after the first sheet selected
    Dim firstSheet As Boolean
    
    'Convert selected sheest to a text string
    For Each ws In ActiveWindow.SelectedSheets
        selectedList = selectedList & ws.Name & "[|]"
    Next ws
    
    'Set the toggle of first sheet
    firstSheet = True
    
    'Loop through each worksheet in the active workbook
    For Each ws In ActiveWorkbook.Sheets
    
        'Check if the worksheet was not previously selected
        If InStr(selectedList, ws.Name & "[|]") = 0 Then
    
            'Check the worksheet is visible
            If ws.Visible = xlSheetVisible Then
    
                'Select the sheet
                ws.Select firstSheet
    
                'First worksheet has been found, toggle to false
                firstSheet = False
    
            End If
    
        End If
    
    Next ws

End Sub
