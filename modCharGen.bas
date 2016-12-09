Attribute VB_Name = "modCharGen"
Option Explicit

Public Function RollMod() As Integer
    Dim r As Single
    Do
        r = Rnd
        Select Case r
        Case Is < 0.33
            RollMod = RollMod - 1
        Case Is > 0.66
            RollMod = RollMod + 1
        Case Else
            Exit Do
        End Select
    Loop
End Function

Public Function GetStatMod(ByVal Stat As Integer) As Integer
    GetStatMod = Int((Stat - 10) / 2)
End Function

Public Function FormatStat(Stat As Integer) As String
    Dim m As Integer
    m = GetStatMod(Stat)
    FormatStat = Stat & "(" & IIf(m > 0, "+", "") & m & ")"
End Function

Public Function WriteCharacterStats(Optional Str As Integer = 10, Optional Dex As Integer = 10, Optional Con As Integer = 10, Optional It As Integer = 10, Optional Wis As Integer = 10, Optional Cha As Integer = 10) As String
    Dim s As String
    s = "Str" & vbTab & "Dex" & vbTab & "Con" & vbTab & "Int" & vbTab & "Wis" & vbTab & "Cha" & vbCrLf
    s = s & _
        FormatStat(RollMod + Str) & vbTab & _
        FormatStat(RollMod + Dex) & vbTab & _
        FormatStat(RollMod + Con) & vbTab & _
        FormatStat(RollMod + It) & vbTab & _
        FormatStat(RollMod + Wis) & vbTab & _
        FormatStat(RollMod + Cha) & vbCrLf

    WriteCharacterStats = s
End Function



