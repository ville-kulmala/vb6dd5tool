VERSION 5.00
Begin VB.Form dlgAbilities 
   Caption         =   "Form1"
   ClientHeight    =   4080
   ClientLeft      =   3150
   ClientTop       =   2970
   ClientWidth     =   2970
   LinkTopic       =   "Form1"
   ScaleHeight     =   4080
   ScaleWidth      =   2970
   Begin VB.CommandButton cmdRoll 
      Caption         =   "Roll"
      Height          =   375
      Left            =   120
      TabIndex        =   34
      Top             =   3600
      Width           =   495
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   372
      Left            =   720
      TabIndex        =   33
      Top             =   3600
      Width           =   972
   End
   Begin VB.TextBox txtName 
      Height          =   288
      Left            =   720
      TabIndex        =   0
      Top             =   75
      Width           =   2172
   End
   Begin VB.TextBox txtAC 
      Height          =   288
      Left            =   720
      TabIndex        =   1
      Top             =   360
      Width           =   2172
   End
   Begin VB.CommandButton cmdRollHits 
      Caption         =   "Roll"
      Height          =   252
      Left            =   2280
      TabIndex        =   30
      Top             =   720
      Width           =   612
   End
   Begin VB.TextBox txtMaxHits 
      Height          =   288
      Left            =   1560
      TabIndex        =   3
      Top             =   720
      Width           =   612
   End
   Begin VB.TextBox txtHits 
      Height          =   288
      Left            =   720
      TabIndex        =   2
      Top             =   720
      Width           =   612
   End
   Begin VB.TextBox txtCR 
      Height          =   288
      Left            =   2280
      TabIndex        =   5
      Text            =   "3"
      Top             =   1080
      Width           =   492
   End
   Begin VB.TextBox txtHD 
      Alignment       =   1  'Right Justify
      Height          =   288
      Left            =   720
      TabIndex        =   4
      Text            =   "3d8"
      Top             =   1080
      Width           =   735
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   372
      Left            =   1800
      TabIndex        =   18
      Top             =   3600
      Width           =   1092
   End
   Begin VB.CheckBox chkSaveProf 
      Caption         =   "Charisma"
      Height          =   252
      Index           =   5
      Left            =   720
      TabIndex        =   17
      Tag             =   "Cha"
      Top             =   3240
      Width           =   1332
   End
   Begin VB.CheckBox chkSaveProf 
      Caption         =   "Wisdom"
      Height          =   252
      Index           =   4
      Left            =   720
      TabIndex        =   15
      Tag             =   "Wis"
      Top             =   2880
      Width           =   1332
   End
   Begin VB.CheckBox chkSaveProf 
      Caption         =   "Inteligence"
      Height          =   252
      Index           =   3
      Left            =   720
      TabIndex        =   13
      Tag             =   "Int"
      Top             =   2520
      Width           =   1332
   End
   Begin VB.CheckBox chkSaveProf 
      Caption         =   "Constitution"
      Height          =   252
      Index           =   2
      Left            =   720
      TabIndex        =   11
      Tag             =   "Con"
      Top             =   2160
      Width           =   1332
   End
   Begin VB.CheckBox chkSaveProf 
      Caption         =   "Dexterity"
      Height          =   252
      Index           =   1
      Left            =   720
      TabIndex        =   9
      Tag             =   "Dex"
      Top             =   1800
      Width           =   1332
   End
   Begin VB.CheckBox chkSaveProf 
      Caption         =   "Strength"
      Height          =   252
      Index           =   0
      Left            =   720
      TabIndex        =   7
      Tag             =   "Str"
      Top             =   1440
      Width           =   1332
   End
   Begin VB.TextBox txtAbility 
      Height          =   288
      Index           =   5
      Left            =   120
      TabIndex        =   16
      Tag             =   "Cha"
      Top             =   3240
      Width           =   492
   End
   Begin VB.TextBox txtAbility 
      Height          =   288
      Index           =   4
      Left            =   120
      TabIndex        =   14
      Tag             =   "Wis"
      Top             =   2880
      Width           =   492
   End
   Begin VB.TextBox txtAbility 
      Height          =   288
      Index           =   3
      Left            =   120
      TabIndex        =   12
      Tag             =   "Int"
      Top             =   2520
      Width           =   492
   End
   Begin VB.TextBox txtAbility 
      Height          =   288
      Index           =   2
      Left            =   120
      TabIndex        =   10
      Tag             =   "Con"
      Top             =   2160
      Width           =   492
   End
   Begin VB.TextBox txtAbility 
      Height          =   288
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Tag             =   "Dex"
      Top             =   1800
      Width           =   492
   End
   Begin VB.TextBox txtAbility 
      Height          =   288
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Tag             =   "Str"
      Top             =   1440
      Width           =   492
   End
   Begin VB.Label Label4 
      Caption         =   "Name:"
      Height          =   252
      Left            =   120
      TabIndex        =   32
      Top             =   120
      Width           =   1092
   End
   Begin VB.Label lblAC 
      Caption         =   "AC"
      Height          =   372
      Left            =   120
      TabIndex        =   31
      Top             =   390
      Width           =   492
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "/"
      Height          =   252
      Left            =   1320
      TabIndex        =   29
      Top             =   720
      Width           =   252
   End
   Begin VB.Label lblHits 
      Caption         =   "Hits"
      Height          =   252
      Left            =   120
      TabIndex        =   28
      Top             =   720
      Width           =   852
   End
   Begin VB.Label lblAbility 
      Caption         =   "+0"
      Height          =   252
      Index           =   5
      Left            =   2160
      TabIndex        =   27
      Top             =   3240
      Width           =   492
   End
   Begin VB.Label lblAbility 
      Caption         =   "+0"
      Height          =   252
      Index           =   4
      Left            =   2160
      TabIndex        =   26
      Top             =   2880
      Width           =   492
   End
   Begin VB.Label lblAbility 
      Caption         =   "+0"
      Height          =   252
      Index           =   3
      Left            =   2160
      TabIndex        =   25
      Top             =   2520
      Width           =   492
   End
   Begin VB.Label lblAbility 
      Caption         =   "+0"
      Height          =   252
      Index           =   2
      Left            =   2160
      TabIndex        =   24
      Top             =   2160
      Width           =   492
   End
   Begin VB.Label lblAbility 
      Caption         =   "+0"
      Height          =   252
      Index           =   1
      Left            =   2160
      TabIndex        =   23
      Top             =   1800
      Width           =   492
   End
   Begin VB.Label lblAbility 
      Caption         =   "+0"
      Height          =   252
      Index           =   0
      Left            =   2160
      TabIndex        =   22
      Top             =   1440
      Width           =   492
   End
   Begin VB.Label Label2 
      Caption         =   "CR"
      Height          =   252
      Left            =   1800
      TabIndex        =   21
      Top             =   1128
      Width           =   492
   End
   Begin VB.Label lblHDBonus 
      Caption         =   "+3"
      Height          =   255
      Left            =   1440
      TabIndex        =   20
      Top             =   1110
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "HD"
      Height          =   252
      Left            =   120
      TabIndex        =   19
      Top             =   1128
      Width           =   852
   End
End
Attribute VB_Name = "dlgAbilities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Character As clsCharacter
Private OKPressed As Boolean

Private Sub chkSaveProf_Click(Index As Integer)
    UpdateForm
End Sub

Private Sub cmdCancel_Click()
    OKPressed = False
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    OKPressed = True
    Me.Hide
End Sub

Public Function ShowDialog(c As clsCharacter, Optional Owner As Form) As Boolean
    Set Character = c
    If c Is Nothing Then Exit Function
    OKPressed = False
    SetAbilityString c.Abilities
    Me.Caption = c.Name
    txtName.Text = c.Name
    txtAC.Text = c.AC
    txtHits.Text = c.Hits
    txtMaxHits.Text = c.MaxHits
    txtHD.Text = c.HD
    txtCR.Text = c.CR
    If Owner Is Nothing Then
        Me.Show 1
    Else
        Me.Show 1, Owner
    End If
    If OKPressed Then
        If IsNumeric(txtAbility(1).Text) Then
            c.InitiativeBase = GetAbilityBonus(txtAbility(1).Text)
        End If
        c.Name = txtName.Text
        c.AC = txtAC.Text
        c.Hits = txtHits.Text
        c.MaxHits = txtMaxHits.Text
        c.HD = txtHD.Text
        c.CR = txtCR.Text
        c.Abilities = GetAbilityString
    End If
    ShowDialog = OKPressed
End Function

Public Function SetAbilityString(ByVal AbilityString As String)
    Dim i As Integer
    Dim a As String
    For i = txtAbility.LBound To txtAbility.UBound
        With txtAbility(i)
            a = GetAbilityFromStr(AbilityString, .Tag)
            .Text = IIf(a <> "0", a, "")
            chkSaveProf(i).Value = IIf(HasAbilitySaveBonus(AbilityString, .Tag), vbChecked, vbUnchecked)
        End With
    Next
End Function

Public Function GetAbilityString() As String
    Dim i As Integer
    For i = txtAbility.LBound To txtAbility.UBound
        GetAbilityString = GetAbilityString & txtAbility(i).Tag & ":" & txtAbility(i).Text & IIf(chkSaveProf(i).Value = vbChecked, "*", "") & ";"
    Next
End Function

Private Sub cmdRoll_Click()
    Dim i As Integer
    Randomize Timer
    For i = txtAbility.LBound To txtAbility.UBound
        If IsNumeric(txtAbility(i).Text) Then
            txtAbility(i).Text = txtAbility(i).Text + Int(Rnd * 6) - Int(Rnd * 6)
        End If
    Next
    cmdRollHits_Click
End Sub

Private Sub cmdRollHits_Click()
    If txtHD.Text <> "" Then
        txtHits.Text = RollDice(txtHD.Text + lblHDBonus.Caption)
    End If
    txtMaxHits.Text = txtHits.Text
End Sub

Private Sub txtAbility_Change(Index As Integer)
    UpdateForm
End Sub

Private Sub UpdateForm()
    Dim i As Integer
    lblHDBonus.Caption = GetHDBonus
    For i = 0 To 5
        If IsNumeric(txtAbility(i).Text) Then
            If chkSaveProf(i).Value = vbChecked Then
                lblAbility(i).Caption = AddPlus(GetAbilityBonus(txtAbility(i).Text) + GetProfiencyBonus(txtCR.Text))
            Else
                lblAbility(i).Caption = AddPlus(GetAbilityBonus(txtAbility(i).Text))
            End If
        Else
            lblAbility(i).Caption = "n/a"
        End If
    Next
End Sub

Private Sub txtHD_Change()
    UpdateForm
End Sub

Public Function GetHDBonus() As String
    Dim s As String
    Dim b As Integer
    Dim bs As String
    Dim crb As Integer
    s = txtHD.Text
    s = PopHead(s, "d")
    If Not IsNumeric(s) Then
        s = 1
    End If
    bs = txtAbility(2).Text
    If IsNumeric(bs) Then
        b = GetAbilityBonus(CInt(bs))
        If b < 0 Then
            GetHDBonus = Int(s) * b
        ElseIf b > 0 Then
            GetHDBonus = "+" & Int(s) * b
        End If
    End If
    
End Function
