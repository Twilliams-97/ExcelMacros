Sub Select_Entire_Full_Rows()
    'Select all entire blank rows in selected range
    
    Dim rRow As Range
    Dim rSelect As Range
    Dim rSelection As Range
     
      'Check that a range is selected
      If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a range first.", vbOKOnly, "Select Full Rows Macro"
        Exit Sub
      End If
      
      'Check that multiple cells are selected
      If Selection.Cells.Count = 1 Then
        Set rSelection = ActiveSheet.UsedRange
      Else
        Set rSelection = Selection
      End If
     
      'Loop through each row and add blank rows to rSelect range
      For Each rRow In rSelection.Rows
        If WorksheetFunction.CountA(rRow) <> 0 Then
          If rSelect Is Nothing Then
            Set rSelect = rRow
          Else
            Set rSelect = Union(rSelect, rRow)
          End If
        End If
      Next rRow
      
      'Select blank rows
      If rSelect Is Nothing Then
        MsgBox "No full rows were found.", vbOKOnly, "Select Full Rows Macro"
        Exit Sub
      Else
         rSelect.EntireRow.Select
      End If
  
End Sub
