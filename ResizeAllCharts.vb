Sub ResizeAllCharts()

    'Create variables to hold chart dimensions
    Dim chtHeight As Long
    Dim chtWidth As Long
    
    'Create variable to loop through chart objects
    Dim chtObj As ChartObject
    
    'Get the size of the first selected chart
    chtHeight = ActiveChart.Parent.Height
    chtWidth = ActiveChart.Parent.Width
    
    For Each chtObj In ActiveSheet.ChartObjects
    
        chtObj.Height = chtHeight
        chtObj.Width = chtWidth
    
    Next chtObj

End Sub
