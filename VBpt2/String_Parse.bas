Attribute VB_Name = "Module1"
Function sReplaceCharacters(strMainString As String, strOld As String, strNew As String) As String
Dim strNewString As String
Dim i As Integer
sReplaceCharacters = ""
For i = 1 To Len(strMainString)
    If Mid(strMainString, i, Len(strOld)) = strOld Then
        strNewString = strNewString & strNew
        i = i + Len(strOld) - 1
    Else
        strNewString = strNewString & Mid(strMainString, i, 1)
    End If
Next i
sReplaceCharacters = strNewString
End Function

