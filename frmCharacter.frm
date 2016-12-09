VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCharacter 
   Caption         =   "Form1"
   ClientHeight    =   6000
   ClientLeft      =   3495
   ClientTop       =   3270
   ClientWidth     =   9150
   LinkTopic       =   "Form1"
   ScaleHeight     =   6000
   ScaleWidth      =   9150
   Begin VB.ComboBox cmbAdvantageMode 
      Height          =   315
      ItemData        =   "frmCharacter.frx":0000
      Left            =   4920
      List            =   "frmCharacter.frx":000D
      TabIndex        =   28
      Text            =   "No advantage"
      Top             =   4200
      Width           =   1575
   End
   Begin VB.TextBox txtNotes 
      Height          =   615
      Left            =   7800
      MultiLine       =   -1  'True
      TabIndex        =   27
      Text            =   "frmCharacter.frx":0038
      Top             =   3840
      Width           =   1335
   End
   Begin VB.CommandButton cmdArticleUp 
      Caption         =   "Move up"
      Height          =   252
      Left            =   6360
      TabIndex        =   26
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton cmdRenameArticle 
      Caption         =   "Rename..."
      Height          =   252
      Left            =   6360
      TabIndex        =   25
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton cmdTargetSelected 
      Caption         =   "Selected"
      Height          =   255
      Left            =   6720
      TabIndex        =   24
      Top             =   4200
      Width           =   1095
   End
   Begin VB.ComboBox cmbTarget 
      Height          =   315
      Left            =   1080
      TabIndex        =   22
      Top             =   4200
      Width           =   3735
   End
   Begin MSComctlLib.Toolbar tbrAttacks 
      Height          =   1380
      Left            =   6720
      TabIndex        =   21
      Top             =   4560
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   2434
      ButtonWidth     =   1746
      ButtonHeight    =   582
      Appearance      =   1
      TextAlignment   =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Attack"
            Key             =   "attack"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Add"
            Key             =   "add"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Remove"
            Key             =   "remove"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Clone"
            Key             =   "clone"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvEffects 
      Height          =   1332
      Left            =   4800
      TabIndex        =   20
      Top             =   2160
      Width           =   2412
      Visible         =   0   'False
      _ExtentX        =   4260
      _ExtentY        =   2355
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ListView lvAttacks 
      Height          =   1212
      Left            =   0
      TabIndex        =   19
      Top             =   4560
      Width           =   6612
      _ExtentX        =   11668
      _ExtentY        =   2143
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Weapon"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Attack"
         Object.Width           =   776
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Damage"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Range"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Description"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.CheckBox chkReactionUsed 
      Caption         =   "Reacted"
      Height          =   252
      Left            =   6360
      TabIndex        =   18
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton cmdRoll 
      Caption         =   "Attack"
      Height          =   252
      Index           =   0
      Left            =   6480
      TabIndex        =   17
      Top             =   3840
      Width           =   732
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   252
      Index           =   0
      Left            =   7320
      TabIndex        =   16
      Top             =   3600
      Width           =   492
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdHit 
      Caption         =   "Hit"
      Height          =   252
      Index           =   0
      Left            =   7200
      TabIndex        =   15
      Top             =   3840
      Width           =   612
   End
   Begin VB.TextBox txtType 
      Height          =   288
      Index           =   0
      Left            =   4320
      TabIndex        =   14
      Text            =   "piercing damage"
      Top             =   3840
      Width           =   2172
   End
   Begin VB.TextBox txtRange 
      Height          =   288
      Index           =   0
      Left            =   3480
      TabIndex        =   13
      Text            =   "range"
      Top             =   3840
      Width           =   852
   End
   Begin VB.TextBox txtDamage 
      Height          =   288
      Index           =   0
      Left            =   2640
      TabIndex        =   12
      Text            =   "1d8"
      Top             =   3840
      Width           =   852
   End
   Begin VB.TextBox txtAttack 
      Height          =   288
      Index           =   0
      Left            =   1800
      TabIndex        =   11
      Text            =   "+1"
      Top             =   3840
      Width           =   852
   End
   Begin VB.TextBox txtWeapon 
      Height          =   288
      Index           =   0
      Left            =   0
      TabIndex        =   10
      Text            =   "Weapon"
      Top             =   3840
      Width           =   1812
   End
   Begin VB.PictureBox picContainer 
      Height          =   1092
      Left            =   3240
      ScaleHeight     =   1035
      ScaleWidth      =   3075
      TabIndex        =   8
      Top             =   2280
      Width           =   3132
      Visible         =   0   'False
      Begin VB.PictureBox picFighter 
         AutoRedraw      =   -1  'True
         Height          =   972
         Left            =   0
         Picture         =   "frmCharacter.frx":003E
         ScaleHeight     =   915
         ScaleWidth      =   915
         TabIndex        =   9
         Top             =   0
         Width           =   972
         Visible         =   0   'False
      End
   End
   Begin VB.Timer tmrUpdate 
      Interval        =   100
      Left            =   600
      Top             =   2760
   End
   Begin VB.PictureBox picCharacter 
      AutoRedraw      =   -1  'True
      Height          =   972
      Left            =   3120
      ScaleHeight     =   915
      ScaleWidth      =   1035
      TabIndex        =   5
      Top             =   480
      Width           =   1092
   End
   Begin VB.CheckBox chkActive 
      Caption         =   "Active"
      Height          =   255
      Left            =   6360
      TabIndex        =   4
      Top             =   1440
      Value           =   1  'Checked
      Width           =   975
   End
   Begin VB.ListBox lstArticles 
      Height          =   1500
      IntegralHeight  =   0   'False
      Left            =   4320
      TabIndex        =   3
      Top             =   360
      Width           =   1935
   End
   Begin VB.CommandButton cmdAddArticle 
      Caption         =   "Add..."
      Height          =   252
      Left            =   6360
      TabIndex        =   2
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton cmdRemoveArticle 
      Caption         =   "Delete"
      Height          =   252
      Left            =   6360
      TabIndex        =   1
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox txtArticle 
      Height          =   1812
      Left            =   3120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1920
      Width           =   4095
   End
   Begin MSComctlLib.ListView lvCharacter 
      Height          =   3372
      Left            =   0
      TabIndex        =   6
      Top             =   360
      Width           =   3012
      _ExtentX        =   5318
      _ExtentY        =   5953
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Value"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Stat"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3480
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   9150
      _ExtentX        =   16140
      _ExtentY        =   741
      ButtonWidth     =   1667
      ButtonHeight    =   582
      Appearance      =   1
      TextAlignment   =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Update"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Abilities"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Target:"
      Height          =   375
      Left            =   120
      TabIndex        =   23
      Top             =   4200
      Width           =   975
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuEditValue 
         Caption         =   "Edit value"
         Shortcut        =   {F2}
      End
   End
End
Attribute VB_Name = "frmCharacter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pUpdate As Boolean  'Timer should do update.
Public pInitiative As frmInitiative
Private pCharacter As clsCharacter

Public Property Set InitiativeForm(Value As frmInitiative)
    Set pInitiative = Value
    ListTargets
End Property

Public Property Set Character(Value As clsCharacter)
    Set pCharacter = Value
    If Value Is Nothing Then
        ClearForm
    Else
        ShowCharacter
    End If
End Property

Public Property Get Character() As clsCharacter
    Set Character = pCharacter
End Property

Private Sub ClearForm()
    txtArticle.Text = ""
    lstArticles.Clear
End Sub

Private Sub chkActive_Click()
    If Not pCharacter Is Nothing Then
        pCharacter.IsActive = IIf(chkActive.Value = vbChecked, True, False)
        If Not pInitiative Is Nothing Then
            pInitiative.ListCharacters
        End If
    End If
End Sub


Private Sub chkReactionUsed_KeyUp(KeyCode As Integer, Shift As Integer)
    pUpdate = True
End Sub

Private Sub cmbAdvantageMode_Click()
    Dim SelChar As clsCharacter
    Dim AM As enAdvantageMode
    With cmbTarget
        If .ListIndex > -1 Then
            Set SelChar = pInitiative.Characters(.ItemData(.ListIndex))
            Select Case cmbAdvantageMode.ListIndex
            Case 0: AM = AMDisadvantage
            Case 1: AM = AMNormal
            Case 2: AM = AMAdvantage
            End Select
            pCharacter.SetAdvantageMode SelChar.Name, AM
        End If
    End With
End Sub

Private Sub cmbTarget_Click()
    Dim SelChar As clsCharacter
    Dim AM As Long
    With cmbTarget
        If .ListIndex > -1 Then
            Set SelChar = pInitiative.Characters(.ItemData(.ListIndex))
            Set pCharacter.Opponent = SelChar
            AM = CInt(pCharacter.GetAdvantageMode(SelChar.Name)) + 1
            cmbAdvantageMode.ListIndex = AM
            
        End If
    End With
End Sub

Private Sub cmdAdd_Click(Index As Integer)
    Dim s As String
    Dim t As String
    s = "Attack-" & Trim(txtWeapon(Index).Text)
    t = txtAttack(Index) & ";" & txtDamage(Index) & ";" & txtRange(Index) & "; " & txtType(Index)
    pCharacter.EntryValue(s) = t
    ShowCharacter
End Sub

Private Sub cmdAddArticle_Click()
    Dim s As String
    If pCharacter Is Nothing Then
        Exit Sub
    End If
    s = InputBox("Anna artikkelin nimi", "Lis‰‰ artikkeli")
    If s = "" Then
        Exit Sub
    End If
    pCharacter.EntryValue(s) = ShowTextEditor(pCharacter.EntryValue(s), Me, s)
    ShowCharacter
End Sub

Private Sub cmdArticleUp_Click()
    Dim oEntry As clsKeyValuePair
    Dim i As Integer
    If pCharacter Is Nothing Then
        Exit Sub
    End If
    If lstArticles.ListIndex < 1 Then
        Exit Sub
    End If
    i = lstArticles.ListIndex
    Set oEntry = pCharacter.GetEntries(i + 1)
    With pCharacter.GetEntries
        .Remove i + 1
        .Add oEntry, Before:=i
    End With
    lstArticles.ListIndex = i - 1   'Update valitsee sitten oikean...
    pUpdate = True
    
End Sub

Private Sub cmdHit_Click(Index As Integer)
    pInitiative.HitCharacter txtDamage(Index).Text, pCharacter, pInitiative.SelectedCharacter
End Sub

Private Sub cmdRemoveArticle_Click()
    If pCharacter Is Nothing Then
        Exit Sub
    End If
    If lstArticles.ListIndex = -1 Then
        Exit Sub
    End If
    pCharacter.RemoveEntryValue (lstArticles.List(lstArticles.ListIndex))
    pUpdate = True
End Sub

Private Sub cmdRenameArticle_Click()
    Dim oEntry As clsKeyValuePair
    Dim sNewName As String
    Dim sOldName As String
    If pCharacter Is Nothing Then
        Exit Sub
    End If
    If lstArticles.ListIndex = -1 Then
        Exit Sub
    End If
    sOldName = lstArticles.List(lstArticles.ListIndex)
    sNewName = InputBox("Please enter new name for article '" & sOldName & "'.", "Rename article", sOldName)
    If sNewName = "" Or sNewName = sOldName Then
        Exit Sub
    End If
    If pCharacter.EntryValueIndex(sNewName) > 0 Then
        MsgBox "Article already exists."
        Exit Sub
    End If
    With pCharacter
        Set oEntry = .GetEntries(.EntryValueIndex(sOldName))
        oEntry.Key = sNewName
    End With
    pUpdate = True

End Sub

Private Sub cmdRoll_Click(Index As Integer)
    Dim R As String
    Dim D As Integer
    Dim AC As String
    R = RollDice(txtAttack(Index).Text & "+d20")
    cmdRoll(Index).Caption = "Roll: " & R
    If Not pInitiative.SelectedCharacter Is Nothing Then
        AC = GetHead(pInitiative.SelectedCharacter.AC)
        If IsNumeric(AC) Then
            If CInt(AC) > CInt(R) Then
                pInitiative.LogEvent lvCharacter.ListItems("Name").Text & " attacks with " & txtWeapon(Index).Text & " for " & R & " missing"
            Else
                pInitiative.LogEvent lvCharacter.ListItems("Name").Text & " attacks with " & txtWeapon(Index).Text & "(vs" & AC & " for " & R & " hitting for:"
                MsgBox pInitiative.HitCharacter(txtDamage(Index).Text, pCharacter, pCharacter.Opponent)
                'Voisihan tuon kai canceloidakkin.
            End If
        End If
    End If
    
    
End Sub

Private Sub cmdTargetSelected_Click()
    Set pCharacter.Opponent = pInitiative.SelectedCharacter
    ListTargets
End Sub

Private Sub Form_Load()
    cmbAdvantageMode.ListIndex = 1
End Sub

Private Sub Form_Resize()
    With txtArticle
        .Width = Abs(Me.ScaleWidth - .Left)
'        .Height = Abs(Me.ScaleHeight - .Top)
    End With
    With lstArticles
        .Width = Abs(Me.ScaleWidth - .Left - cmdAddArticle.Width)
        cmdAddArticle.Left = .Width + .Left
        cmdRemoveArticle.Left = .Width + .Left
        cmdRenameArticle.Left = .Width + .Left
        cmdArticleUp.Left = .Width + .Left
        chkActive.Left = .Width + .Left
        chkReactionUsed.Left = .Width + .Left
    End With
    tbrAttacks.Left = Me.ScaleWidth - tbrAttacks.Width
    With lvAttacks
        .Move 0, .Top, Abs(Me.ScaleWidth - tbrAttacks.Width), Abs(Me.ScaleHeight - .Top)
    End With
    With txtNotes
        .Width = Abs(Me.ScaleWidth - .Left)
    End With
'    lvCharacter.Height = Me.ScaleHeight
End Sub

Private Sub lstArticles_Click()
    If lstArticles.ListIndex = -1 Or pCharacter Is Nothing Then
        txtArticle.Text = ""
    Else
        txtArticle.Text = pCharacter.EntryValue(lstArticles.List(lstArticles.ListIndex))
        'pUpdate = True
    End If
    
End Sub

Private Sub lstArticles_DblClick()
    Dim i As Integer
    If lstArticles.ListIndex > -1 Then
        If pCharacter Is Nothing Then
            lstArticles.Clear
        Else
            i = lstArticles.ListIndex
            With pCharacter
                .EntryValue(lstArticles.List(i)) = ShowTextEditor(.EntryValue(lstArticles.List(i)), Me)
            End With
        End If
    End If
End Sub

Private Sub AddIfWeapon(Key As String, Value As String)
    'Attack-Crossbow: +4;1d8+2;50/100;pearcing damage
    Dim a As Variant
    Dim i As Integer
    Dim li As ListItem
    Dim j As Integer
    If Key Like "Attack-*" And Value Like "*;*;*;*" Then
        a = Split(Value, ";", 4)
        If UBound(a) = 3 Then
'            i = txtWeapon.Count
'            Load txtWeapon(i)
'            With txtWeapon(i)
'                .Visible = True
'                .Move txtWeapon(0).Left, txtWeapon(0).Top + cmdAdd(0).Height * i
'                .Text = Trim(Mid(Key, 8))
'            End With
'            Load txtAttack(i)
'            With txtAttack(i)
'                .Visible = True
'                .Move txtAttack(0).Left, txtAttack(0).Top + cmdAdd(0).Height * i
'                .Text = Trim(a(0))
'            End With
'            Load txtDamage(i)
'            With txtDamage(i)
'                .Visible = True
'                .Move txtDamage(0).Left, txtDamage(0).Top + cmdAdd(0).Height * i
'                .Text = Trim(a(1))
'            End With
'            Load txtRange(i)
'            With txtRange(i)
'                .Visible = True
'                .Move txtRange(0).Left, txtRange(0).Top + cmdAdd(0).Height * i
'                .Text = Trim(a(2))
'            End With
'            Load txtType(i)
'            With txtType(i)
'                .Visible = True
'                .Move txtType(0).Left, txtType(0).Top + cmdAdd(0).Height * i
'                .Text = Trim(a(3))
'            End With
'            Load cmdHit(i)
'            With cmdHit(i)
'                .Visible = True
'                .Move cmdHit(0).Left, cmdHit(0).Top + cmdAdd(0).Height * i
'            End With
'            Load cmdAdd(i)
'            With cmdAdd(i)
'                .Visible = True
'                .Move cmdAdd(0).Left, cmdAdd(0).Top + cmdAdd(0).Height * i
'            End With
'            Load cmdRoll(i)
'            With cmdRoll(i)
'                .Visible = True
'                .Move cmdRoll(0).Left, cmdRoll(0).Top + cmdAdd(0).Height * i
'            End With
            'Listaan...
            Set li = lvAttacks.ListItems.Add(Key:=Key, Text:=Trim(Mid(Key, 8)))
            For j = 0 To 3
                li.SubItems(j + 1) = Trim(a(j))
            Next
            li.ToolTipText = Trim(a(3))
        End If
        
    End If
End Sub


Private Sub lvAttacks_AfterLabelEdit(Cancel As Integer, NewString As String)
    With lvAttacks.selectedItem
        pCharacter.EntryValue("Attack-" & NewString) = pCharacter.EntryValue(.Key)
        pCharacter.RemoveEntryValue .Key
        ShowCharacter
    End With
End Sub

Private Sub lvAttacks_DblClick()
    If lvAttacks.selectedItem Is Nothing Then
        Exit Sub
    End If
    EditAttack lvAttacks.selectedItem.Key
End Sub

Private Sub EditAttack(ByVal Key As String)
    Dim D As dlgAttack
    Dim a As String
    Dim pKey As String
    If Key = "" Then
        Exit Sub
    End If
    a = pCharacter.EntryValue(Key)
    Set D = New dlgAttack
    pKey = Key
    If D.ShowDialog(pCharacter, Key, a, Me) Then
        If pKey <> Key Then
            pCharacter.RemoveEntryValue pKey
        End If
        pCharacter.EntryValue(Key) = a
        ShowCharacter
    End If
End Sub

Private Sub lvAttacks_ItemClick(ByVal Item As MSComctlLib.ListItem)
    txtWeapon(0).Text = Item.Text
    txtAttack(0).Text = Item.SubItems(1)
    txtDamage(0).Text = Item.SubItems(2)
    txtRange(0).Text = Item.SubItems(3)
    txtType(0).Text = Item.SubItems(4)
End Sub

Private Sub lvCharacter_AfterLabelEdit(Cancel As Integer, NewString As String)
    pUpdate = True
End Sub

Private Sub mnuEditValue_Click()
    If Not lvCharacter.selectedItem Is Nothing Then
        lvCharacter.StartLabelEdit
    End If
End Sub

Private Sub picCharacter_Click()
    PicPicture
End Sub

Private Sub tbrAttacks_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "attack":  AttackWithSelected pCharacter.Opponent
    Case "add":     AddAttackDialog
    Case "remove":  DeleteSelectedAttack
    Case "clone":   CloneAttack
    End Select
End Sub

Private Sub AttackWithSelected(Optional Target As clsCharacter)
    Dim AM As enAdvantageMode
    If lvAttacks.selectedItem Is Nothing Then
        Exit Sub
    End If
    If Target Is Nothing Then
        Set Target = pInitiative.SelectedCharacter
    End If

    Select Case cmbAdvantageMode.ListIndex
    Case 0: AM = AMDisadvantage
    Case 1: AM = AMNormal
    Case 2: AM = AMAdvantage
    Case Else: AM = AMNormal
    End Select
    
    pCharacter.Attack lvAttacks.selectedItem.Text, Target, pInitiative, AM


End Sub

Private Sub AddAttackDialog()
    Dim D As dlgAttack
    Dim a As String
    Set D = New dlgAttack
    Dim Key As String
    Key = "Attack-New"
    a = GetDefaultAttackStats
    If D.ShowDialog(pCharacter, Key, a, Me) Then
        pCharacter.EntryValue(Key) = a
        ShowCharacter
    End If
End Sub

Private Sub DeleteSelectedAttack()
    If lvAttacks.selectedItem Is Nothing Then Exit Sub
    pCharacter.RemoveEntryValue lvAttacks.selectedItem.Key
    ShowCharacter
End Sub

Private Sub CloneAttack()

    If lvAttacks.selectedItem Is Nothing Then Exit Sub
    pCharacter.EntryValue(lvAttacks.selectedItem.Key & "+") = pCharacter.EntryValue(lvAttacks.selectedItem.Key)
    ShowCharacter
End Sub

Private Function GetDefaultAttackStats() As String
    Dim s As String
    Dim b As String
    Dim i As Integer
    Dim db As String
    Dim dbi As Integer
    i = GetAbilityFromStr(pCharacter.Abilities, "Str")
    If i <> 0 Then
        dbi = GetAbilityBonus(i)
        Select Case dbi
        Case Is < 0: db = dbi
        Case Is > 0: db = "+" & dbi
        End Select
        b = GetAbilityBonus(i) + GetProfiencyBonus(pCharacter.CR)
        
    Else
        b = GetProfiencyBonus(pCharacter.CR)
    End If
    s = IIf(b < 0, b, "+" & b)
    s = s & ";d6" & db & ";5;slashing damage"
    GetDefaultAttackStats = s
End Function

Private Sub tmrUpdate_Timer()
    If pUpdate Then
        If Not pCharacter Is Nothing Then
            UpdateCharacterListView pCharacter, lvCharacter, chkActive
            With pCharacter
                .ReactionUsed = IIf(chkReactionUsed.Value = vbChecked, True, False)
            End With
            If Not pInitiative Is Nothing Then
                pInitiative.ShowCharacter
            End If
            ShowCharacter
        End If
        pUpdate = False
    End If
End Sub

Public Sub PicPicture()
    If pCharacter Is Nothing Then
        Exit Sub
    End If
    On Error Resume Next
    Dim Filename As String
    Filename = pCharacter.PictureFile
    If Filename = "" Or Filename = "\" Then
        Filename = GetSetting("CombatMapper", "CharacterLists", "PictureFolder", "") & "*.*"
    End If
    With CommonDialog1
        .Filter = "Picture files|*.jpg;*.bmp;*.wmg;*.gif;*.ico"
        .Filename = Filename
        .CancelError = True
        .ShowOpen
        If Err = 0 Then
            pCharacter.PictureFile = .Filename
            pInitiative.ListCharacters
            ShowCharacter
            SaveSetting "CombatMapper", "CharacterLists", "PictureFolder", PathFromString(.Filename)
        End If
    End With
End Sub

'Copied from Initiative.
'Update changes to there.
Public Sub ShowCharacter()
    Dim nItm As ListItem
    Dim rCharacter As clsCharacter
    Dim iEntry As clsKeyValuePair
    
    If pCharacter Is Nothing Then
        Set rCharacter = CreateCharacter
    Else
        Set rCharacter = pCharacter
    End If
    
    With rCharacter
        lvCharacter.ListItems.Clear
        Set nItm = lvCharacter.ListItems.Add(, "Name", .Name)
        nItm.SubItems(1) = "Name"
        Set nItm = lvCharacter.ListItems.Add(, "Initiative", .Initiative)
        nItm.SubItems(1) = "Initiative"
        Set nItm = lvCharacter.ListItems.Add(, "InitiativeBase", .InitiativeBase)
        nItm.SubItems(1) = "Initiative base"
        Set nItm = lvCharacter.ListItems.Add(, "Hits", .Hits)
        nItm.SubItems(1) = "Hits"
        Set nItm = lvCharacter.ListItems.Add(, "TempHits", .TempHits)
        nItm.SubItems(1) = "TempHits"
        Set nItm = lvCharacter.ListItems.Add(, "MaxHits", .MaxHits)
        nItm.SubItems(1) = "MaxHits"
        Set nItm = lvCharacter.ListItems.Add(, "Status", .Status)
        nItm.SubItems(1) = "Status"
        Set nItm = lvCharacter.ListItems.Add(, "Size", .Size)
        nItm.SubItems(1) = "Size"
        Set nItm = lvCharacter.ListItems.Add(, "AC", .AC)
        nItm.SubItems(1) = "AC"
        Set nItm = lvCharacter.ListItems.Add(, "Speed", .Speed)
        nItm.SubItems(1) = "Speed"
        Set nItm = lvCharacter.ListItems.Add(, "HD", .HD)
        nItm.SubItems(1) = "HD"
        Set nItm = lvCharacter.ListItems.Add(, "CR", .CR)
        nItm.SubItems(1) = "CR"
        Set nItm = lvCharacter.ListItems.Add(, "Abilities", .Abilities)
        nItm.SubItems(1) = "Abilities"
        
        chkReactionUsed.Value = IIf(.ReactionUsed, vbChecked, vbUnchecked)
        If .PictureFile = "" Then
            Set picCharacter.Picture = picFighter.Picture
        Else
            picCharacter.Picture = LoadPicture()
            If FileExists(.PictureFile) Then
                PaintWithAspectRatio picCharacter, .PictureFile, 0, 0, picCharacter.ScaleWidth, picCharacter.ScaleHeight
                'picCharacter.PaintPicture .GetPicture, 0, 0, picCharacter.ScaleWidth, picCharacter.ScaleHeight
            Else
                picCharacter.PaintPicture picFighter.Picture, 0, 0, picCharacter.ScaleWidth, picCharacter.ScaleHeight
            End If
        End If
        chkActive.Value = IIf(.IsActive, vbChecked, vbUnchecked)
        Dim pOldArticleIndex As Integer
        pOldArticleIndex = lstArticles.ListIndex
        lstArticles.Clear
        txtArticle.Text = ""

        lvAttacks.ListItems.Clear
        For Each iEntry In .GetEntries
            lstArticles.AddItem iEntry.Key
            AddIfWeapon iEntry.Key, iEntry.Value
        Next
        If pOldArticleIndex > 0 And lstArticles.ListCount > pOldArticleIndex Then
            lstArticles.ListIndex = pOldArticleIndex
            lstArticles_Click
        ElseIf lstArticles.ListCount > 0 Then
            lstArticles.ListIndex = 0
            lstArticles_Click
        End If
        ListTargets
        txtNotes.Text = .EntryValue("notes")
    End With
    
End Sub

Private Sub ListTargets()
    Dim iChar As clsCharacter
    Dim i As Integer
    Dim sText As String
    cmbTarget.Clear
    On Error Resume Next
    If pInitiative Is Nothing Then Exit Sub
    For Each iChar In pInitiative.Characters
        i = i + 1
        If iChar.IsActive Then
            sText = iChar.Name
            Select Case pCharacter.GetAdvantageMode(iChar.Name)
            Case enAdvantageMode.AMAdvantage:       sText = sText & " (advantage)"
            Case enAdvantageMode.AMDisadvantage:    sText = sText & " (disadvantage)"
            End Select
            With cmbTarget
                .AddItem sText
                .ItemData(.NewIndex) = i
                If iChar Is pCharacter.Opponent Then
                    .ListIndex = i - 1
                    cmbTarget.Text = pCharacter.Opponent.Name
                    
                End If
            End With
        End If
    Next
    Debug.Print "jees"
End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Caption
    Case "Update"
        pUpdate = True
    Case "Abilities"
        ShowAbilitiesDialog
    End Select
End Sub

Private Sub ShowAbilitiesDialog()
    Dim D As dlgAbilities
    Set D = New dlgAbilities
    If D.ShowDialog(pCharacter) Then
        ShowCharacter
        pUpdate = True
    End If
    Unload D
End Sub

Private Sub txtNotes_Change()
    If Not pCharacter Is Nothing Then
        pCharacter.EntryValue("notes") = txtNotes.Text
    End If
End Sub
