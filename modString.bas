Attribute VB_Name = "modString"
Option Explicit

Public Function GetLastIndexOf(ByVal Text As String, SubString As String) As Integer
    Dim i As Integer
    i = InStr(1, Text, SubString)
    Do While i > 0
        GetLastIndexOf = i
        i = InStr(i + 1, Text, SubString)
    Loop
    
End Function

Public Function GetTail(ByVal Text As String, Optional Delimiter As String = " ") As String
    Dim a As Variant
    If InStr(Text, Delimiter) > 0 Then
        a = Split(Text, Delimiter)
        GetTail = a(UBound(a))
    End If
End Function

Public Function GetHead(ByVal Text As String, Optional Delimeter As String = " ") As String
    Dim P As Integer
    P = InStr(Text, Delimeter)
    If P > 0 Then
        GetHead = Left(Text, P - 1)
    Else
        GetHead = Text
    End If
End Function

Public Function PopHead(ByRef Text As String, Optional Delimiter As String = " ") As String
    Dim P As Integer
    P = InStr(Text, Delimiter)
    If P > 0 Then
        PopHead = Left(Text, P - 1)
        Text = Mid(Text, P + 1)
    Else
        PopHead = Text
        Text = ""
    End If
End Function

Public Function ConcatCol(Col As Collection, Optional Delimeter As String = vbCrLf) As String
    Dim v As Variant
    For Each v In Col
        ConcatCol = ConcatCol & v & Delimeter
    Next
End Function

'Hoitaa rivinvaihdot yhteen pötköön yms.
Public Function EncodeText(ByVal Text As String) As String
    EncodeText = Replace(Text, vbCrLf, "\n;")
    EncodeText = Replace(EncodeText, vbTab, "\t;")
End Function

'Purkaa "encodingin"
Public Function DecodeText(ByVal Text As String) As String
    DecodeText = Replace(Text, "\n;", vbCrLf)
    DecodeText = Replace(DecodeText, "\t;", vbTab)
End Function

Public Function PathFromString(ByVal Filename As String) As String
    Dim i As Integer, n As Integer
    If Right(Filename, 1) = "\" Then
        PathFromString = Filename
    Else
        Do
            n = InStr(i + 1, Filename, "\")
            If n > 0 Then
                i = n
            Else
                Exit Do
            End If
        Loop
        PathFromString = Left(Filename, i)
    End If
End Function
