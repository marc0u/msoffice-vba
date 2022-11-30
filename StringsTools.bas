Attribute VB_Name = "Module1"
Public Function FileName(HideExt As Boolean, Optional CharLeft As Integer = 0)
    FileName = ThisWorkbook.Name
    If InStrRev(FileName, ".") > 0 And HideExt Then
        FileName = Left(FileName, InStrRev(FileName, ".") - 1)
    End If
    If CharLeft > 0 Then
        FileName = Left(FileName, CharLeft)
    End If
End Function

Public Function ReplaceText(Target As String, TxtToFind As String, TxtToReplace As String)
    ReplaceText = Replace(Target, TxtToFind, TxtToReplace)
End Function
