Sub Select_Blank_Columns()
    'Select all entire blank columns in selected range
    
    Dim rCol As Range
    Dim rSelect As Range
    Dim rSelection As Range
     
      'Check that a range is selected
      If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a range first.", vbOKOnly, "Select Blank Columns Macro"
        Exit Sub
      End If
      
      'Check that multiple cells are selected
      If Selection.Cells.Count = 1 Then
        Set rSelection = ActiveSheet.UsedRange
      Else
        Set rSelection = Selection
      End If
     
      'Loop through each row and add blank rows to rSelect range
      For Each rCol In rSelection.Columns
        If WorksheetFunction.CountA(rCol) = 0 Then
          If rSelect Is Nothing Then
            Set rSelect = rCol
          Else
            Set rSelect = Union(rSelect, rCol)
          End If
        End If
      Next rCol
      
      'Select blank columns
      If rSelect Is Nothing Then
        MsgBox "No blank columns were found.", vbOKOnly, "Select Blank Columns Macro"
        Exit Sub
      Else
         rSelect.Select
      End If
  
End Sub

