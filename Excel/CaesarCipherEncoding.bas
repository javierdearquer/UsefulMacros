Sub CaesarCipherEncoding()
    Dim selectedRange As Range
    Dim shiftValue As Integer
    Dim cell As Range
    Dim originalText As String
    Dim encodedText As String
    Dim i As Integer
    Dim charCode As Integer
    Dim newCharCode As Integer
    
    ' Check if any cells are selected
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a range of cells.", vbExclamation
        Exit Sub
    End If
    
    ' Prompt user for the shift value
    shiftValue = InputBox("Enter the shift value (positive or negative):", "Caesar Cipher Encoding")
    
    ' Loop through each selected cell
    For Each selectedRange In Selection
        ' Check if cell contains text
        If IsNumeric(selectedRange.Value) = False Then
            originalText = selectedRange.Value
            encodedText = ""
            
            ' Loop through each character in the text
            For i = 1 To Len(originalText)
                charCode = Asc(Mid(originalText, i, 1))
                
                ' Apply shift to alphabetic characters
                If charCode >= 65 And charCode <= 90 Then ' Uppercase letters (ASCII range)
                    newCharCode = ((charCode - 65 + shiftValue) Mod 26 + 26) Mod 26 + 65
                ElseIf charCode >= 97 And charCode <= 122 Then ' Lowercase letters (ASCII range)
                    newCharCode = ((charCode - 97 + shiftValue) Mod 26 + 26) Mod 26 + 97
                Else
                    newCharCode = charCode ' Non-alphabetic characters remain unchanged
                End If
                
                encodedText = encodedText & Chr(newCharCode)
            Next i
            
            ' Replace the original text with the encoded text
            selectedRange.Value = encodedText
        End If
    Next selectedRange
    
    MsgBox "Text encoded successfully!", vbInformation
End Sub
