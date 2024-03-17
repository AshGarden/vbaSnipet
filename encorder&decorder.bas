Sub EnigmaEncryptCell()
    Dim key As Integer
    Dim cell As Range
    Set cell = Application.ActiveCell
    key = InputBox("暗号化のキーを入力してください（整数）") ' ユーザーにキーを入力させます
    cell.Value = EnigmaEncrypt(cell.Value, key)
End Sub

Sub EnigmaDecryptCell()
    Dim key As Integer
    Dim cell As Range
    Set cell = Application.ActiveCell
    key = InputBox("復号化のキーを入力してください（整数）") ' ユーザーにキーを入力させます
    cell.Value = EnigmaDecrypt(cell.Value, key)
End Sub

Function EnigmaEncrypt(inputString As String, key As Integer) As String
    Dim outputString As String
    Dim i As Integer
    For i = 1 To Len(inputString)
        outputString = outputString & Chr(Asc(Mid(inputString, i, 1)) + key)
    Next i
    EnigmaEncrypt = outputString
End Function

Function EnigmaDecrypt(inputString As String, key As Integer) As String
    Dim outputString As String
    Dim i As Integer
    For i = 1 To Len(inputString)
        outputString = outputString & Chr(Asc(Mid(inputString, i, 1)) - key)
    Next i
    EnigmaDecrypt = outputString
End Function
