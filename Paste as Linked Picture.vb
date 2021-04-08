Sub PasteCellsAsLinkedPicture()

    'Copy the selection
    Selection.Copy
    
    'Offset the picture to start in column next to selection
    ActiveCell.Offset(1, Selection.Columns.Count).Select
    
    'Paste the copy an a linked image
    ActiveSheet.Pictures.Paste Link:=True
    
    'Remove the marching ants
    Application.CutCopyMode = False

End Sub
