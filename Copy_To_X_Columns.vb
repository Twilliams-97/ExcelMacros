Function Redoer(ByVal textinput As String, ByRef Target As Range)

        Do While Not Valid
        
        Set newtarget = Application.InputBox(textinput, "Turn 1 Column to X", Target.Address, Type:=8)
        If newtarget.Count > 1 Then
            MsgBox "Only select one cell. Please try again"
            Valid = False
        Else
            Valid = True
        End If
    Loop
    
    newtarget.Select

End Function

Sub Copy_to_seXy_Columns()
    
    Dim outputrng As Range
    Dim inputrng As Range
    Dim bottomrow As Range    

    'how many values to put per row
    Dim valperrow As Integer
    valperrow = Application.InputBox("How many columns? :", xTitleId, , Type:=1)
    interval = valperrow
    
    'First cell in the column
    
    Set inputrng = Application.Selection
    firstcell = Redoer("Select First Cell in Column :", inputrng)
   
    first_row = ActiveCell.Row
    first_col = ActiveCell.Column 
    
    'Get the last cell in the column
    Set bottomrow = Application.Selection
    lastcell = Redoer("Select Last Row in Column:", bottomrow)
    
    last_row = ActiveCell.Row
    
    'Uncomment the following (and comment above) if you want it to go all the way down the column
    
    'last_row = Cells(Rows.Count, first_col).End(xlUp).Row
    
    'first cell where you want the data to go
    Set outputrng = Application.Selection
    outputcell = Redoer("Select top left of where you want it to go :", outputrng)
    
    dest_start_col = ActiveCell.Column
    dest_start_row = ActiveCell.Row

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
