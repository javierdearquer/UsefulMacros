Sub GoToA1AndZoom()
    Dim ws As Worksheet 
    For Each ws In ActiveWorkbook.Sheets 
        ws.Activate 
        Cells(ActiveWindow.SplitRow + 1, ActiveWindow.SplitColumn + 1).Select 
        ActiveWindow.Zoom = 100 
    Next ws 
    ActiveWorkbook.Worksheets(1).Activate 
End Sub
