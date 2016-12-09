Attribute VB_Name = "modRelativePath"
Option Explicit

Public Function GetRelativePath(ByVal TargetPath As String, ByVal HomeDir As String)
    'TODO: bugaa viel‰ osittain
    Dim a As Variant
    Dim b As Variant
    Dim i As Integer
    Dim j As Integer
    Dim R As String
    a = Split(TargetPath, "\")
    b = Split(HomeDir, "\")
    
    Do
        If UBound(a) <= i Or UBound(b) <= i Then
            Exit Do
        End If
        If LCase(a(i)) <> LCase(b(i)) Then
            Exit Do
        End If
        i = i + 1
    Loop
    'i sis‰lt‰‰ nyt, mill‰ tasolla polku on sama.
    If i < UBound(b) Then
        For j = i To UBound(b) - 1
            R = R & "..\"
        Next
    End If
    
    
    If i < UBound(a) Then
        If R <> "" Then
            'Otetaan viimeinen \ pois
            R = Left(R, Len(R) - 1)
        Else
            R = "."
        End If
        For j = i To UBound(a)
            R = R & "\" & a(j)
        Next
    
    End If
    GetRelativePath = R
End Function

Public Function RelativeToAbsolutePath(ByVal RelativePath As String, ByVal CurrentPath As String)
    'TODO: bugaa viel‰ (aivan varmasti)
    Dim rel As Variant
    Dim Cur As Variant
    Dim i As Integer
    Dim j As Integer
    Dim cPos As Integer
    Dim R As String
    rel = Split(RelativePath, "\")
    Cur = Split(CurrentPath, "\")
    If rel(0) = "." Then
        'En ole en‰‰ varma tuosta p‰‰tteest‰ alla
        R = RemTail(CurrentPath, "\")
        For i = 1 To UBound(rel)
            R = R & "\" & rel(i)
        Next
    Else
        j = UBound(Cur)
        Do
            Select Case rel(i)
            Case ".."
                j = j - 1
            Case Else
                'K‰‰nnepointti. L‰hdet‰‰n toiseen suuntaan.
            End Select
            i = i + 1
        Loop
    End If
    
End Function

Public Function RemTail(ByVal Text As String, ByVal Tail As String)
    If Len(Tail) > Text Then
        Exit Function
    End If
    If Right(Text, Len(Tail)) <> Tail Then
    
        RemTail = Left(Text, Len(Text) - Tail)
    End If
    
End Function

Public Function AddTail(ByVal Text As String, ByVal Tail As String)
    If Len(Tail) > Len(Text) Then
        AddTail = Text & Tail
        Exit Function
    End If
    If Right(Text, Len(Tail)) <> Tail Then
    
        AddTail = Text & Tail
    Else
        AddTail = Text
    End If
    
End Function

Public Function ArrayConcat(a As Variant, ByVal Delim As String) As String
    Dim i As Integer
    For i = 0 To UBound(a)
        ArrayConcat = ArrayConcat & a & Delim
    Next
End Function

