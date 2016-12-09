Attribute VB_Name = "modDice"
Option Explicit

'Heitt‰‰ vain mahdollisuutta vastaan ja kertoo montako kertaa tarvi
'heitt‰‰ ennen kuin failas.
Public Function RollColChange(ByVal Change As Single) As Integer
    Dim Roll As Single
    RollColChange = -1
    Do
        RollColChange = RollColChange + 1
        Roll = Rnd
    Loop While Roll < Change
End Function

'Rikkoo rekursiivisesti ja heitt‰‰.
Public Function RollDice(ByVal Dice As String) As Integer
    Dim P As Integer
    Dice = Trim(Dice)
    P = InStr(2, Dice, "+")
    If P > 0 Then
        RollDice = RollDice(Left(Dice, P - 1)) + RollDice(Mid(Dice, P + 1))
    Else
        P = InStr(2, Dice, "-")
        If P > 0 Then
            RollDice = RollDice(Left(Dice, P - 1)) - RollDice(Mid(Dice, P + 1))
        Else
            Dice = Replace(Dice, "+", "")
            If Left(Dice, 1) = "-" Then
                RollDice = -RollDie(Mid(Dice, 2))
            Else
                RollDice = RollDie(Dice)
            End If
        End If
    End If
    
End Function

Public Function MakeDiceCrit(ByVal Dice As String) As String
    Dim P As Integer
    Dim D As String
    Dice = Trim(Dice)
    P = InStr(2, Dice, "+")
    If P > 0 Then
        MakeDiceCrit = MakeDiceCrit(Left(Dice, P - 1)) & "+" & MakeDiceCrit(Mid(Dice, P + 1))
    Else
        P = InStr(2, Dice, "-")
        If P > 0 Then
            'Miinus mukaan
            MakeDiceCrit = MakeDiceCrit(Left(Dice, P - 1)) & MakeDiceCrit(Mid(Dice, P))
        Else
            Dice = Replace(Dice, "+", "")
            If Left(Dice, 1) = "-" Then
                MakeDiceCrit = Dice
            Else
                P = InStr(LCase(Dice), "d")
                If P > 0 Then
                    D = Left(Dice, P - 1)
                    If IsNumeric(D) Then
                        MakeDiceCrit = CInt(D) * 2 & Mid(Dice, P)
                    ElseIf D = "" Then
                        MakeDiceCrit = 2 & Mid(Dice, P)
                    Else
                        MakeDiceCrit = Dice
                    End If
                Else
                    MakeDiceCrit = Dice
                End If
            End If
        End If
    End If
    
End Function

'Heitt‰‰ vaikka 2d6
Public Function RollDie(ByVal Die As String) As Integer
    Dim P As Integer
    Dim i As Integer
    Dim k As Integer
    Randomize Timer
    If IsNumeric(Die) And InStr(LCase(Die), "d") = 0 Then
        'ilm. esim 2d6 on numeerinen
        RollDie = Die
    Else
        P = InStr(LCase(Die), "d")
        If P > 1 Then
            k = CInt(Left(Die, P - 1))
        Else
            k = 1
        End If
        For i = 1 To k
            RollDie = RollDie + 1 + Int(Rnd * Mid(Die, P + 1))
        Next
    End If
End Function

Public Function RollD20(AdvantageMode As enAdvantageMode) As Integer
    Dim d1 As Integer
    Dim d2 As Integer
    If AdvantageMode = AMNormal Then
        RollD20 = Int(Rnd * 20) + 1
    Else
        d1 = Int(Rnd * 20) + 1
        d2 = Int(Rnd * 20) + 1
        If d1 < d2 Then
            If AdvantageMode = AMAdvantage Then
                RollD20 = d2
            Else
                RollD20 = d1
            End If
        Else
            If AdvantageMode = AMDisadvantage Then
                RollD20 = d2
            Else
                RollD20 = d1
            End If
        End If
        Debug.Print "Rolls: " & d1 & " & " & d2 & " with " & AdvantageMode & ":" & RollD20
    End If
    
End Function

