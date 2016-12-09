Attribute VB_Name = "modDD"
Option Explicit

Public Enum enAdvantageMode
    AMDisadvantage = -1
    AMNormal = 0
    AMAdvantage = 1
End Enum

Public Function AddPlus(ByVal Value As String) As String
    AddPlus = Value
    If IsNumeric(Value) Then
        If Value >= 0 Then
            AddPlus = "+" & Value
        End If
        
    End If
End Function

Public Function GetProfiencyBonus(ByVal CR As String) As String
    CR = GetHead(CR)    'Otetaan pois lopusta esim. 'PC'
    If IsNumeric(CR) Then
        Select Case Int(CR)
        Case 0 To 4: GetProfiencyBonus = 2
        Case 5 To 8: GetProfiencyBonus = 3
        Case 9 To 12: GetProfiencyBonus = 4
        Case 13 To 16: GetProfiencyBonus = 5
        Case 17 To 20: GetProfiencyBonus = 6
        Case 21 To 24: GetProfiencyBonus = 7
        Case 25 To 28: GetProfiencyBonus = 8
        Case 29 To 30: GetProfiencyBonus = 9
        End Select
    Else
        GetProfiencyBonus = 2
    End If
End Function

Public Function GetXPValue(CR As String) As String
    'Experience
    Select Case Trim(CR)
    Case "0":   GetXPValue = 10
    Case "1/8": GetXPValue = 25
    Case "1/4": GetXPValue = 50
    Case "1/2": GetXPValue = 100
    Case "1":   GetXPValue = 200
    Case "2":   GetXPValue = 450
    Case "3":   GetXPValue = 700
    Case "4":   GetXPValue = 1100
    Case "5":   GetXPValue = 1800
    Case "6":   GetXPValue = 2300
    Case "7":   GetXPValue = 2900
    Case "8":   GetXPValue = 3900
    Case "9":   GetXPValue = 5000
    Case "10":   GetXPValue = 7200
    Case "11":   GetXPValue = 8400
    Case "12":   GetXPValue = 10000
    Case "13":   GetXPValue = 11500
    Case "14":   GetXPValue = 13000
    Case "15":   GetXPValue = 15000
    Case "16":   GetXPValue = 18000
    Case Else
        GetXPValue = 0
    End Select
End Function

Public Function GetAbilityBonus(Score As Integer) As Integer
    GetAbilityBonus = Int((Score - 10) / 2)
End Function

Public Function GetAbilityFromStr(ByVal AbilityStr As String, ByVal Ability As String) As Integer
    'Esim: Str:14;Dex:12*;Con:14;Int:15*;Wis:12;Cha:8
        
        Dim s As String
        s = GetKeyValueLineValue(AbilityStr, Ability)
        s = Replace(s, "*", "")
        If IsNumeric(s) Then
            GetAbilityFromStr = s
        End If

End Function

Public Function GetKeyValueLineValue(ByVal ValueLine As String, ByVal Key As String) As String
    Dim a As Variant
    Dim i As Integer
    Dim s As String
    a = Split(ValueLine, ";")
    For i = 0 To UBound(a)
        s = a(i)
        If LCase(Trim(PopHead(s, ":"))) = LCase(Key) Then
            GetKeyValueLineValue = Trim(s)
            Exit Function
        End If
    Next
End Function

Public Function HasAbilitySaveBonus(ByVal AbilityStr As String, ByVal Ability As String) As Boolean
    Dim s As String
    s = GetKeyValueLineValue(AbilityStr, Ability)
    If InStr(s, "*") > 0 Then
        HasAbilitySaveBonus = True
    End If
End Function


'Voidaan hakea GetAttacks haulla keyvalue pairit...
'Voitaisiin lisätä popuppiin listassa vaikka "attack of opportunity" tyyppinen valinta
'TODO:
Public Function AttackWith(Attacker As clsCharacter, Target As clsCharacter, AttackKeyValue As clsKeyValuePair, InitiativeForm As frmInitiative, AdvantageMode As enAdvantageMode)
    
End Function
