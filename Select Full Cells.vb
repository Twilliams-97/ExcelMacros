Sub Select_Full_Cells()
    'Select all non-blank cells in selected range
    
    Dim fc As Range
    Dim rSelect As Range
    Dim rSelection As Range
     
      'Check that a range is selected
      If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a range first.", vbOKOnly, "Select Full Cells Macro"
        Exit Sub
      End If
      
      'Check that multiple cells are selected
      If Selection.Cells.Count = 1 Then
        Set rSelection = ActiveSheet.UsedRange
      Else
        Set rSelection = Selection
      End If
     
      'Loop through each row and add blank cells to rSelect range
      For Each fc In rSelection
        If WorksheetFunction.CountA(fc) <> 0 Then
          If rSelect Is Nothing Then
            Set rSelect = fc
          Else
            Set rSelect = Union(rSelect, fc)
          End If
        End If
      Next fc
      
      'Select blank cells
      If rSelect Is Nothing Then
        MsgBox "No full rows were found.", vbOKOnly, "Select Full Rows Macro"
        Exit Sub
      Else
         rSelect.Select
      End If
  
End Sub
