Sub GoToA1AndZoom()
    Dim ws As Worksheet
    Application.ScreenUpdating = False
    For Each ws In ThisWorkbook.Sheets
        ws.Activate
        ws.Range("A1").Select
        ActiveWindow.Zoom = 100
    Next ws
    Application.ScreenUpdating = True
End Sub
