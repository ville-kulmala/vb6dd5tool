Attribute VB_Name = "modFile"
Option Explicit

Public Function FileToCol(ByVal Filename As String) As Collection
    Dim r As String
    Dim fn As Integer
    fn = FreeFile
    Set FileToCol = New Collection
    Open Filename For Input As fn
    Do While Not EOF(fn)
        Line Input #fn, r
        FileToCol.Add r
    Loop
    Close fn
End Function

Public Sub WriteColToFile(ByVal Filename As String, Col As Collection)
    Dim r As Variant
    Dim fn As Integer
    fn = FreeFile
    Open Filename For Output As fn
    For Each r In Col
        Print #fn, r
    Next
    Close fn
End Sub

Public Function FileExists(ByVal Filename As String) As Boolean
    If Filename = "" Then Exit Function
    If Filename = "\" Then Exit Function
    FileExists = (Dir(Filename, vbNormal) <> "")
End Function
