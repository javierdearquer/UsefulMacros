Sub SaveEachTabAsSeparateFile()
    ' Macro to save each worksheet as a separate Excel file 
    ' in the same directory as the current workbook

    Dim ws As Worksheet
    Dim filePath As String

    ' Get the path of the current workbook
    filePath = ThisWorkbook.Path & "\"

    ' Disable screen updating and alerts for better performance
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Loop through each worksheet in the workbook
    For Each ws In ThisWorkbook.Worksheets
        ' Copy the worksheet to a new workbook
        ws.Copy

        ' Save the new workbook with the sheet name
        ActiveWorkbook.SaveAs fileName:=filePath & ws.Name & ".xlsx", _
                              FileFormat:=xlOpenXMLWorkbook
        
        ' Close the new workbook without saving further changes
        ActiveWorkbook.Close SaveChanges:=False
    Next ws

    ' Re-enable screen updating and alerts
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub

