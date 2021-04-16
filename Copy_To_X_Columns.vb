Sub Copy_to_X_Columns()

    Dim valperrow As Integer
    Dim outputrng As Range
    Dim inputrng As Range
    Dim bottomrow As Range
    
    xTitleId = "Turn 1 Column to X"

    'how many values to put per row
    
    valperrow = Application.InputBox("How many columns? :", xTitleId, Type:=1)
    interval = valperrow
    
    'First cell in the column
    
    Set inputrng = Application.Selection
    Set inputrng = Application.InputBox("Select First Cell in Column :", xTitleId, inputrng.Address, Type:=8)
    inputrng.Select
    first_row = ActiveCell.Row
    first_col = ActiveCell.Column
    
    
    'Get the last cell in the column
    Set bottomrow = Application.Selection
    Set bottomrow = Application.InputBox("Select Last Cell in Column:", xTitleId, bottomrow.Address, Type:=8)
    bottomrow.Select
    last_row = ActiveCell.Row
    
    'Uncomment the following (and comment above) if you want it to go all the way down the column
    
    'last_row = Cells(Rows.Count, first_col).End(xlUp).Row
    
    'first cell where you want the data to go
    Set outputrng = Application.Selection
    Set outputrng = Application.InputBox("Select top left of where you want it to go :", xTitleId, outputrng.Address, Type:=8)
    outputrng.Select
    
    dest_start_col = ActiveCell.Column '3
    dest_start_row = ActiveCell.Row '1

    dest_cur_col = dest_start_col
    dest_cur_row = dest_start_row
    

    For cur_row = first_row To last_row
  
        Cells(dest_cur_row, dest_cur_col).Value = Cells(cur_row, first_col).Value

        dest_cur_col = dest_cur_col + 1
    
        If (cur_row - (first_row - 1)) Mod interval = 0 Then

            dest_cur_col = dest_start_col
            dest_cur_row = dest_cur_row + 1

        End If
    
    Next

End Sub
