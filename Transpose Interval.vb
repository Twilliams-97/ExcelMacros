Sub transpose_interval()

    Dim valperrow As Integer
    Dim InputRng As Range
    Dim OutputRng As Range
    
    xTitleId = "Turn 1 Column to X"

    'how many values to put per row
    
    valperrow = Application.InputBox("How many columns? :", xTitleId, Type:=1)
    
    interval = valperrow
    
     'data you want to transpose
    'Set InputRng = Application.Selection
    'Set InputtRng = Application.InputBox("Select top left of where you want it to go :", xTitleId, InputRng.Address, Type:=8)
    'InputRng.Select
    
    'first row of the data that you want to transpose
    
    first_row = 1 'ActiveCell.Row 'InputRng.Row
    
    'first column of the data that you want to transpose
    
    first_col = 1 'ActiveCell.Column  'InputRng.Column
    
     
    
    'first column where you want the data to go
    Set OutputRng = Application.Selection
    Set OutputRng = Application.InputBox("Select top left of where you want it to go :", xTitleId, OutputRng.Address, Type:=8)
    OutputRng.Select
    
    
    dest_start_col = ActiveCell.Column '3
    
    'first row where you want the data to go
    
    dest_start_row = ActiveCell.Row '1
    
     
    
    dest_cur_col = dest_start_col
    
    dest_cur_row = dest_start_row
    
     
    
    last_row = Cells(Rows.Count, first_col).End(xlUp).Row
    
     
    
    For cur_row = first_row To last_row
    
     
    
        Cells(dest_cur_row, dest_cur_col).Value = Cells(cur_row, first_col).Value
    
     
    
        dest_cur_col = dest_cur_col + 1
    
     
    
        If (cur_row - (first_row - 1)) Mod interval = 0 Then
    
       
    
            dest_cur_col = dest_start_col
    
       
    
            dest_cur_row = dest_cur_row + 1
    
       
    
        End If
    
     
    
    Next

 

End Sub
