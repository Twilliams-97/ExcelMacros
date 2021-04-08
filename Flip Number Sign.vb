Sub FlipNumberSignage()

  'Create variable to hold cells in the worksheet
  Dim c As Range

  'Loop through each cell in selection
  For Each c In Selection

      'Test if the cell contents is a number
      If IsNumeric(c) Then

          'Convert signage for each cell
          c.Value = -c.Value

      End If

  Next c

End Sub
