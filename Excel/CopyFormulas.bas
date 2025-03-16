Sub Copy_Formulas()
    ' Macro to copy formulas from the "formula template" sheet 
    ' to all sheets in the workbook starting from the third one

    Dim ws As Worksheet
    Dim sourceWs As Worksheet
    Dim i As Integer

    ' Set the source worksheet
    Set sourceWs = ThisWorkbook.Worksheets("formula template")

    ' Copy formulas from the source sheet
    sourceWs.Cells.Copy

    ' Loop through all worksheets starting from the third one
    For i = 3 To ThisWorkbook.Worksheets.Count
        Set ws = ThisWorkbook.Worksheets(i)
        ws.Range("A1").PasteSpecial Paste:=xlPasteFormulas, _
                                    Operation:=xlNone, _
                                    SkipBlanks:=True, _
                                    Transpose:=False
    Next i

    ' Clear clipboard to free memory
    Application.CutCopyMode = False

End Sub

