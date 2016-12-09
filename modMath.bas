Attribute VB_Name = "modMath"
Option Explicit


Public Function Min(R1, R2)
    If R1 < R2 Then
        Min = R1
    Else
        Min = R2
    End If
End Function

Public Function Max(R1, R2)
    If R1 > R2 Then
        Max = R1
    Else
        Max = R2
    End If
End Function
