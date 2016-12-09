VERSION 5.00
Begin VB.Form dlgAttack 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   1350
   ClientTop       =   1395
   ClientWidth     =   2970
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   2970
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   372
      Left            =   960
      TabIndex        =   12
      Top             =   2760
      Width           =   852
   End
   Begin VB.ComboBox cmbStat 
      Height          =   288
      ItemData        =   "dlgAttack.frx":0000
      Left            =   1800
      List            =   "dlgAttack.frx":0016
      TabIndex        =   6
      Text            =   "From stat"
      Top             =   480
      Width           =   1092
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   372
      Left            =   1920
      TabIndex        =   5
      Top             =   2760
      Width           =   972
   End
   Begin VB.TextBox txtWeapon 
      Height          =   288
      Left            =   840
      TabIndex        =   0
      Text            =   "Weapon"
      Top             =   120
      Width           =   2052
   End
   Begin VB.TextBox txtAttack 
      Height          =   288
      Left            =   840
      TabIndex        =   1
      Text            =   "+1"
      Top             =   480
      Width           =   852
   End
   Begin VB.TextBox txtDamage 
      Height          =   288
      Left            =   840
      TabIndex        =   2
      Text            =   "1d8"
      Top             =   840
      Width           =   2052
   End
   Begin VB.TextBox txtRange 
      Height          =   288
      Left            =   840
      TabIndex        =   3
      Text            =   "range"
      Top             =   1200
      Width           =   2052
   End
   Begin VB.TextBox txtType 
      Height          =   885
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   4
      Text            =   "dlgAttack.frx":0038
      Top             =   1800
      Width           =   2895
   End
   Begin VB.Label Label4 
      Caption         =   "Type"
      Height          =   252
      Left            =   0
      TabIndex        =   11
      Top             =   1560
      Width           =   732
   End
   Begin VB.Label Label3 
      Caption         =   "Range"
      Height          =   372
      Left            =   0
      TabIndex        =   10
      Top             =   1200
      Width           =   852
   End
   Begin VB.Label Label2 
      Caption         =   "Damage"
      Height          =   252
      Left            =   0
      TabIndex        =   9
      Top             =   840
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "Attack"
      Height          =   252
      Left            =   0
      TabIndex        =   8
      Top             =   480
      Width           =   612
   End
   Begin VB.Label lblWeapon 
      Caption         =   "Weapon"
      Height          =   372
      Left            =   0
      TabIndex        =   7
      Top             =   120
      Width           =   732
   End
End
Attribute VB_Name = "dlgAttack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private pKey As String
Private pValue As String

Private pCharacter As clsCharacter
Private pOKClicked As Boolean
Public Function ShowDialog(ByVal Character As clsCharacter, ByRef ArticleKey As String, ByRef ArticleContent As String, OwnerForm As Form) As Boolean
    Dim a As Variant
    txtWeapon.Text = GetTail(ArticleKey, "-")
    a = Split(ArticleContent, ";", 4)
    If UBound(a) < 3 Then
        Exit Function
    End If
    Set pCharacter = Character
    txtAttack.Text = Trim(a(0))
    txtDamage.Text = Trim(a(1))
    txtRange.Text = Trim(a(2))
    txtType.Text = Trim(a(3))
    Me.Caption = txtWeapon.Text
    Me.Show 1, OwnerForm
    If pOKClicked Then
        GetAttackString ArticleKey, ArticleContent
        ShowDialog = True
    Else
        ShowDialog = False
    End If
End Function

Private Sub cmbStat_Click()
    Dim b As String
    Dim s As Integer
    Dim d As Integer
    Dim i As Integer
    Dim lastMod As String
    s = GetAbilityFromStr(pCharacter.Abilities, cmbStat.Text)
    If s > 0 Then
        s = GetAbilityBonus(s)
        d = s
        s = s + CInt(GetProfiencyBonus(pCharacter.CR))
        lastMod = txtAttack.Text
        txtAttack.Text = AddPlus(s)
        
        txtDamage.Text = Replace(txtDamage.Text, lastMod, "")
        
        i = GetLastIndexOf(txtDamage.Text, "+")
        If i > 0 Then
            txtDamage.Text = Left(txtDamage.Text, i - 1)
        End If
        txtDamage.Text = txtDamage.Text & AddPlus(d)
    End If
    
End Sub

Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    If GetAttackString Then
        pOKClicked = True
        Me.Hide
    End If
End Sub

Private Function GetAttackString(Optional ArticleKey As String, Optional ArticleValue As String) As Boolean
    ArticleKey = "Attack-" & Trim(txtWeapon.Text)
    ArticleValue = txtAttack & ";" & txtDamage.Text & ";" & txtRange.Text & "; " & txtType.Text
    GetAttackString = True
End Function

Private Sub Combo1_Change()

End Sub
