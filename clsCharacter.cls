VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCharacter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Name As String
Public Initiative As Single
Public InitiativeBase As Single
Public Hits As Integer
Public TempHits As Integer
Public MaxHits As Integer
Public Status As String
Public Size As Single
Public AC As String
Public Speed As String
Public CR As String
Public HD As String
Public Abilities As String

Public ReactionUsed As Boolean

Private pPictureFile As String
Public Location As String

Public SourceFile As String

Private WithEvents DragForm As frmDragForm
Attribute DragForm.VB_VarHelpID = -1
Private pFormScaling As Single
Private pFormZooming As Single

Public Event Resize(X As Single, Y As Single)
Public Event Move(X As Single, Y As Single, PathLen As Single, Pause As Boolean)
Public Event Clicked()

Public IsActive As Boolean  'Näytetäänkö listoissa, hypätäänkö yli?

Public InitiativeForm As frmInitiative

Public CharacterForm As frmCharacter

Private pEntries As New Collection    'of clsKeyValuePair     VANHAINFO-793 15-03-03: kollaasi kaikista arvoista

Public Opponent As clsCharacter

Private pAdvantageModes As Collection   'Tekstejä: mode + : + kohde (esim. hahmon nimi)

Public Sub SetAdvantageMode(Target As String, Mode As enAdvantageMode)
    Dim i As Variant
    Dim c As Collection
    Dim k As String
    Dim v As String
    If pAdvantageModes Is Nothing Then
        Set pAdvantageModes = New Collection
        pAdvantageModes.Add Mode & ":" & Target
    Else
        Set c = New Collection
        For Each i In pAdvantageModes
            k = i
            v = PopHead(k, ":")
            If k <> Target Then
                c.Add i
            End If
        Next
        If Mode <> AMNormal Then
            c.Add Mode & ":" & Target
        End If
        Set pAdvantageModes = c
    End If
End Sub

Public Function GetAdvantageMode(Target As String) As enAdvantageMode
    Dim i As Variant
    Dim k As String
    Dim v As String
    
    If pAdvantageModes Is Nothing Then
        GetAdvantageMode = AMNormal
        Exit Function
    End If
    For Each i In pAdvantageModes
        k = i
        v = PopHead(k, ":")
        If k = Target Then
            GetAdvantageMode = CInt(v)
            Exit Function
        End If
    Next
    
End Function

Public Property Get PictureFile() As String
    PictureFile = pPictureFile
End Property

Public Property Let PictureFile(Value As String)
    If FileExists(Value) Then
        pPictureFile = Value
    Else
        Dim P As String
        If Value <> "" Then
            P = GetTail(Value, "\")
            P = AddTail(GetSetting("CombatMapper", "CharacterLists", "PictureFolder", ""), "\") & P
            If FileExists(P) Then
                pPictureFile = P
            Else
                pPictureFile = Value
            End If
            If Err <> 0 Then
                Debug.Print
            End If
        End If
    End If
End Property

Public Property Let EntryValue(ByVal Key As String, ByVal Value As String)
    modDictionary.SetKeyValue pEntries, Key, Value
End Property

Public Property Get EntryValue(ByVal Key As String) As String
    EntryValue = modDictionary.GetKeyValue(pEntries, Key)
End Property

Public Function RemoveEntryValue(ByVal Key As String) As clsKeyValuePair
    Dim i As Integer
    i = IndexOfKeyValuePair(pEntries, Key)
    If i > 0 Then
        Set RemoveEntryValue = pEntries(i)
        pEntries.Remove i
    End If
End Function

Public Property Get EntryValueIndex(ByVal Key As String) As Integer
    EntryValueIndex = IndexOfKeyValuePair(pEntries, Key)
End Property


Public Function GetEntries() As Collection
    Set GetEntries = pEntries
End Function


Public Function GetPicture() As StdPicture
    If FileExists(PictureFile) Then
        Set GetPicture = LoadPicture(PictureFile)
    End If
End Function

Public Function GetDragForm(Create As Boolean) As frmDragForm
    If DragForm Is Nothing Then
        If Not Create Then
            Exit Function
        End If
        Set DragForm = New frmDragForm
        
        If Location <> "" And pFormScaling > 0 Then
            MoveToPosition
        End If
    End If
    
    Set GetDragForm = DragForm
    
End Function

Public Function MoveToPosition()
    On Error Resume Next
    With DragForm
        .AllowResizing = False
        .Move Me.Left * pFormScaling * pFormZooming, Me.Top * pFormScaling * pFormZooming, Me.Size * pFormScaling * pFormZooming, Me.Size * pFormScaling * pFormZooming
    End With
    Err.Clear
End Function

Public Property Get Left() As Single
    'Yläreuna. Sijainti x on keskellä...
    Dim l As String
    l = GetHead(Location, ":")
    If IsNumeric(l) Then
        Left = l
    Else
        If DragForm Is Nothing Then
            Left = 0
        Else
            Left = (DragForm.Left + DragForm.Width / 2) / pFormScaling / pFormZooming
        End If
        
    End If
End Property

Public Property Get Top() As Single
    Dim t As String
    t = GetTail(Location, ":")
    If IsNumeric(t) Then
        Top = t
    Else
        If DragForm Is Nothing Then
            Top = 0
        Else
            Top = (DragForm.Top + DragForm.Height / 2) / pFormScaling / pFormZooming
        End If
    End If
End Property

Private Sub Class_Initialize()
    IsActive = True
    Size = 5
    pFormScaling = 94
    pFormZooming = 1
    GetDragForm True    'Tehdään nyt alkajaisiksi aina
End Sub

Private Sub DragForm_Clicked()
    RaiseEvent Clicked
    If Not InitiativeForm Is Nothing Then
        Set InitiativeForm.SelectedCharacter = Me
    End If
End Sub

Private Sub DragForm_Move(X As Single, Y As Single, PathLen As Single, Pause As Boolean)
    RaiseEvent Move(X, Y, PathLen, Pause)
    Me.Location = X / pFormScaling / pFormZooming & ":" & Y / pFormScaling / pFormZooming
End Sub

Private Sub DragForm_Resize(X As Single, Y As Single)
    RaiseEvent Resize(X, Y)
End Sub

Public Property Let Scaling(ByVal Value As Single)
    pFormScaling = Value
    UpdateSize
End Property

Public Property Let Zooming(ByVal Value As Single)
    pFormZooming = Value
    UpdateSize
    'MoveToPosition
End Property

Public Sub UpdateSize()
    If Not DragForm Is Nothing Then
        With DragForm
            .Width = pFormScaling * Size * pFormZooming
            .Height = pFormScaling * Size * pFormZooming
            .AllowResizing = False
        End With
    End If
End Sub


Public Sub RefreshForm()
    If Not DragForm Is Nothing Then
        On Error Resume Next
        DragForm.BackColor = vbWhite
        DragForm.Picture = LoadPicture()
        DragForm.Caption = Me.Name  'Tooltippiä varten jotenkin...
        PaintWithAspectRatio DragForm, PictureFile, 0, 0, DragForm.ScaleWidth, DragForm.ScaleHeight
        'DragForm.PaintPicture Me.GetPicture, 0, 0, DragForm.ScaleWidth, DragForm.ScaleHeight
        Err.Clear
    End If
    If Not CharacterForm Is Nothing Then
        CharacterForm.ShowCharacter
    End If
End Sub

Public Function GetSpeed() As Integer
    Dim s As String
    s = GetHead(Speed)
    If IsNumeric(s) Then
        GetSpeed = s
    Else
        GetSpeed = 30
    End If
End Function

Public Function RollHitPoints() As Integer
    Dim sCon As String
    RollHitPoints = MaxHits
    If HD <> "" Then
        sCon = HD & GetHDBonus
        Debug.Print "Roll HP for '" & Name & "' : " & sCon
        RollHitPoints = RollDice(sCon)
    End If
    If RollHitPoints < 1 Then RollHitPoints = 1
    Hits = RollHitPoints
    MaxHits = RollHitPoints
End Function

Public Function GetHDBonus() As String
    Dim s As String
    Dim b As Integer
    Dim bs As String
    Dim crb As Integer
    s = HD
    s = PopHead(s, "d")
    If Not IsNumeric(s) Then
        s = 1
    End If
    bs = GetAbilityFromStr(Abilities, "con")
    If IsNumeric(bs) Then
        b = GetAbilityBonus(CInt(bs))
        If b < 0 Then
            GetHDBonus = Int(s) * b
        ElseIf b > 0 Then
            GetHDBonus = "+" & Int(s) * b
        End If
    End If
    
End Function

Public Function GetDistanceToPoint(ByVal X As Single, ByVal Y As Single, ByVal ByForm As Boolean) As Single
    If ByForm Then
        GetDistanceToPoint = Sqr(Abs(X - (DragForm.Left + DragForm.Width / 2)) ^ 2 + Abs(Y - (DragForm.Top + DragForm.Height / 2)) ^ 2)
    Else
        GetDistanceToPoint = Sqr(Abs(X - Left) ^ 2 + Abs(Y - Top) ^ 2)
    End If
End Function

'Palauttaa lokattavan tiedon
Public Function Hit(ByVal Damage As String, Optional ByVal IgnoreTempHits As Boolean) As String
    Dim d As Integer
    d = RollDice(Damage)
    If TempHits > 0 Then
        TempHits = TempHits - d
        If TempHits < 0 Then
            Hits = Hits + TempHits
            TempHits = 0
        End If
    Else
        Hits = Hits - d
    End If
    IsActive = IIf(Hits > 0, True, False)
    Hit = Name & " was hit for " & d & " points"
    If Not IsNumeric(Damage) Then
        Hit = Hit + " (" & Damage & ")"
    End If
    If Hits > 0 Then
        Hit = Hit & " and is at " & Hits & " (" & Format(Hits / MaxHits * 100, "0") & "%)"
    Else
        Hit = Hit & " and goes down at " & Hits & "."
    End If
End Function

Public Function GetSave(ByVal Ability As String)
    Dim sAbb As Integer
    sAbb = GetAbilityFromStr(Abilities, Ability)
    If sAbb > 0 Then
        sAbb = GetAbilityBonus(sAbb)
        If HasAbilitySaveBonus(Abilities, Ability) Then
            sAbb = sAbb + GetProfiencyBonus(CR)
        End If
        GetSave = sAbb
    End If
End Function

'Palauttaa KeyValuePairs
Public Function GetAttacks() As Collection
    Dim iKV As clsKeyValuePair
    'Esim.:
    'Attack-Crossbow: +4;1d8+2;50/100;pearcing damage
    Set GetAttacks = New Collection
    For Each iKV In pEntries
        If iKV.Key Like "Attack-*" And iKV.Value Like "*;*;*;*" Then
            GetAttacks.Add iKV
        End If
    Next
End Function

Public Function GetAttackNames() As Collection
    Dim c As Collection
    Dim iKV As clsKeyValuePair
    Set GetAttackNames = New Collection
    Set c = GetAttacks
    For Each iKV In c
        GetAttackNames.Add GetTail(iKV.Key, "-")
    Next
End Function

Public Sub Attack(ByVal AttackName As String, Target As clsCharacter, InitList As frmInitiative, Optional AdvantageMode As enAdvantageMode = AMNormal)
    Dim R As String
    Dim dr As Integer
    Dim d As Integer
    Dim AC As String
    Dim sAV As String
    Dim a As Variant
    Dim ams As String '
    Dim Dice As String
    Dim Desc As String
    'Attack-Crossbow: +4;1d8+2;50/100;pearcing damage
    sAV = Me.EntryValue("Attack-" & AttackName)
    If Not sAV Like "*;*;*;*" Then
        MsgBox "Attack '" & AttackName & "' is not an attack or could not be found"
        Exit Sub
    End If
    a = Split(sAV, ";", 4, vbTextCompare)
    dr = RollD20(AdvantageMode)
    R = RollDice(a(0) & "+" & dr)
    'cmdRoll(0).Caption = "Roll: " & R
    
    Select Case AdvantageMode
    Case enAdvantageMode.AMAdvantage: ams = " at advantage"
    Case enAdvantageMode.AMDisadvantage: ams = " at disvantage"
    End Select
    If Not Target Is Nothing Then
        AC = GetHead(Target.AC)
        If IsNumeric(AC) Then
            If CInt(AC) > CInt(R) Then
                InitList.LogEvent Name & " attacks " & Target.Name & " with " & AttackName & ams & " for " & R & " missing. Roll " & dr
            Else
                If dr = 20 Then
                    Dice = MakeDiceCrit(a(1))
                    d = RollDice(Dice)
                    Desc = "Attack hits CRITICAL (total " & R & " vs. AC" & AC & ") for " & d & " points (damage roll: " & Dice & ", attack roll " & dr & ams & ")"
                    If MsgBox(Desc, vbOKCancel, Name & " attacks " & Target.Name & " with " & AttackName) = vbOK Then
                        InitList.HitCharacter d, Me, Target, Name & " attacks " & Target.Name & ": " & Desc
                    End If
                
                Else
                    d = RollDice(a(1))
                    Desc = "Attack hits (total " & R & " vs. AC" & AC & ") for " & d & " points (roll " & dr & ams & ")"
                    If MsgBox(Desc, vbOKCancel, Name & " attacks " & Target.Name & " with " & AttackName) = vbOK Then
                        InitList.HitCharacter d, Me, Target, Name & " attacks " & Target.Name & ": " & Desc
                    End If
                End If
            End If
        End If
    End If

End Sub

'TODO: CharacterFormin asetus ja mahdollinen "kaappaus" jos vaikka CurrentCharacterilla on oma
' forminsa.
