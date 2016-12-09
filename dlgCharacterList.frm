VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form dlgCharacterList 
   Caption         =   "Selected characters"
   ClientHeight    =   4305
   ClientLeft      =   2430
   ClientTop       =   3915
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   ScaleHeight     =   4305
   ScaleWidth      =   5895
   Begin MSComctlLib.Toolbar tbrModifiers 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   635
      ButtonWidth     =   1905
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Modifier"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   4
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "advantage"
                  Text            =   "Advantage"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "disadvantage"
                  Text            =   "Disadvantage"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "vulnerability"
                  Text            =   "Vulnerability"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "resistance"
                  Text            =   "Resistance"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Remove"
            Key             =   "remove"
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtDC 
      Height          =   288
      Left            =   3480
      TabIndex        =   8
      Text            =   "12"
      Top             =   684
      Width           =   372
   End
   Begin VB.CommandButton cmdHitSelected 
      Caption         =   "Hit"
      Height          =   252
      Left            =   3960
      TabIndex        =   6
      Top             =   720
      Width           =   732
   End
   Begin VB.TextBox txtSave 
      Height          =   288
      Left            =   2520
      TabIndex        =   5
      Text            =   "Dex"
      Top             =   684
      Width           =   492
   End
   Begin VB.CheckBox chkAllowSave 
      Caption         =   "Save halves"
      Height          =   252
      Left            =   1320
      TabIndex        =   4
      Top             =   720
      Value           =   1  'Checked
      Width           =   1332
   End
   Begin VB.TextBox txtHitDamage 
      Height          =   288
      Left            =   600
      TabIndex        =   3
      Text            =   "8d6"
      Top             =   684
      Width           =   612
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   372
      Left            =   4560
      TabIndex        =   1
      Top             =   3840
      Width           =   1212
   End
   Begin MSComctlLib.ListView lvCharacters 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   5106
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Effects"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Character"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Hits"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Modifiers"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label2 
      Caption         =   "DC"
      Height          =   252
      Left            =   3120
      TabIndex        =   7
      Top             =   720
      Width           =   492
   End
   Begin VB.Label Label1 
      Caption         =   "Hit"
      Height          =   252
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1092
   End
End
Attribute VB_Name = "dlgCharacterList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pCharacters As Collection

Private OKPressed As Boolean
Public Sub ShowDialog(Characters As Collection, Optional OwnerForm As Form)
    Set pCharacters = Characters
    OKPressed = False
    ListCharacters True
    If OwnerForm Is Nothing Then
        Me.Show 1
    Else
        Me.Show 1, OwnerForm
    End If
    If OKPressed Then
        AffectCharacters
    End If
End Sub

Public Sub AddCharacter(Character As clsCharacter)
    Dim iCharacter As clsCharacter
    If pCharacters Is Nothing Then
        Set pCharacters = New Collection
    Else
        For Each iCharacter In pCharacters
            If iCharacter Is Character Then
                Exit Sub
            End If
        Next
    End If
    pCharacters.Add Character
    ListCharacters True
End Sub

Public Sub AffectCharacters()
    Dim iChar As clsCharacter
    Dim iItm As ListItem
    For Each iItm In lvCharacters.ListItems
        If iItm.Checked Then
            Set iChar = FindCharacter(iItm.Key, pCharacters)
            EffectCharacter iChar, iItm.Text
        End If
        
    Next
End Sub

Public Sub ListCharacters(SelectAll As Boolean)
    Dim iChar As clsCharacter
    Dim nItm As ListItem
    lvCharacters.ListItems.Clear
    For Each iChar In pCharacters
        Set nItm = lvCharacters.ListItems.Add(, iChar.Name, "")
        nItm.Ghosted = Not iChar.IsActive
        nItm.SubItems(1) = iChar.Name
        nItm.SubItems(2) = iChar.Hits & "/" & iChar.MaxHits
        nItm.Checked = True
    Next
End Sub

Private Sub cmdHitSelected_Click()

    Dim nItm As ListItem
    Dim iChar As clsCharacter
    Dim cSave As String
    Dim bHalve As Boolean
    Dim bSkip As Boolean
    Dim lSave As Integer
    Dim dDamage As String
    Dim iRoll As Integer
    Dim iDmg As Single
    dDamage = RollDice(txtHitDamage.Text)
    For Each nItm In lvCharacters.ListItems
        If nItm.Checked Then
            bHalve = False
            bSkip = False
            Set iChar = FindCharacter(nItm.Key, pCharacters)
            If chkAllowSave.Value = vbChecked Then
                cSave = iChar.GetSave(txtSave.Text)
                If Not IsNumeric(cSave) Then
                    cSave = InputBox("Save '" & txtSave.Text & "' defined for character '" & iChar.Name & "'. Enter save bonus", "No save defined", lSave)
                    If IsNumeric(cSave) Then
                        lSave = cSave
                    Else
                        bSkip = True
                    End If
                End If
                If Not bSkip Then
                    If InStr(nItm.SubItems(3), "A") Then
                        iRoll = Min(RollDie("d20"), RollDie("d20"))
                    ElseIf InStr(nItm.SubItems(3), "D") Then
                        iRoll = Max(RollDie("d20"), RollDie("d20"))
                    Else
                        iRoll = RollDie("d20")
                    End If
                    'Roll mod: a/d
                    If iRoll + CInt(cSave) >= CInt(txtDC.Text) Then
                        bHalve = True
                    End If
                End If
            Else
                bHalve = False
            End If
            If Not bSkip Then
                If bHalve Then
                    'Damage r/v
                    iDmg = Int((dDamage + 1) / 2)
                Else
                    iDmg = Int(dDamage)
                End If
                If InStr(nItm.SubItems(3), "R") > 0 Then
                    iDmg = Int((iDmg + 1) / 2)
                ElseIf InStr(nItm.SubItems(3), "V") > 0 Then
                    iDmg = iDmg * 2
                End If
                nItm.Text = nItm.Text + "Dmg:" & iDmg & ";"
                
            End If
        End If
    Next
    
End Sub

Private Sub cmdOK_Click()
    OKPressed = True
    Me.Hide
End Sub

Private Sub tbrModifiers_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "remove"
        RemoveSelectedCharacters
    End Select
End Sub

Private Sub RemoveSelectedCharacters()
    Dim nItem As ListItem
    Dim iChar As clsCharacter
    For Each nItem In lvCharacters.ListItems
        If nItem.Checked Then
            Set iChar = FindCharacter(nItem.Key, pCharacters)
            RemoveCharacter pCharacters, iChar
        End If
    Next
    ListCharacters True
End Sub

Private Sub tbrModifiers_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case ButtonMenu.Key
    Case "advantage":
        RemoveStatusFromSelected "D"
        ToggleStatusInSelected "A"
    Case "disadvantage"
        RemoveStatusFromSelected "A"
        ToggleStatusInSelected "D"
    Case "vulnerability"
        RemoveStatusFromSelected "R"
        ToggleStatusInSelected "V"
    Case "resistance"
        RemoveStatusFromSelected "V"
        ToggleStatusInSelected "R"
    End Select
End Sub

Private Sub RemoveStatusFromSelected(ByVal Status As String)
    Dim i As ListItem
    For Each i In lvCharacters.ListItems
        If i.Selected Then
            i.SubItems(3) = Replace(i.SubItems(3), Status, "")
        End If
    Next

End Sub

Private Sub ToggleStatusInSelected(ByVal Status As String)
    Dim i As ListItem
    For Each i In lvCharacters.ListItems
        If i.Selected Then
            If InStr(i.SubItems(3), Status) > 0 Then
                i.SubItems(3) = Replace(i.SubItems(3), Status, "")
            Else
                i.SubItems(3) = i.SubItems(3) & Status
            End If
        End If
    Next
End Sub
