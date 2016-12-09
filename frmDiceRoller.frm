VERSION 5.00
Begin VB.Form frmDiceRoller 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "DiceRoller"
   ClientHeight    =   4668
   ClientLeft      =   4416
   ClientTop       =   1356
   ClientWidth     =   2556
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4668
   ScaleWidth      =   2556
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtResults 
      Height          =   2052
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   14
      Top             =   2640
      Width           =   2532
   End
   Begin VB.CommandButton cmdRoll 
      Caption         =   "Roll"
      Height          =   312
      Index           =   6
      Left            =   1560
      TabIndex        =   13
      Top             =   2040
      Width           =   972
   End
   Begin VB.TextBox txtDice 
      Height          =   288
      Index           =   6
      Left            =   0
      TabIndex        =   12
      Text            =   "d100"
      Top             =   2040
      Width           =   1572
   End
   Begin VB.CommandButton cmdRoll 
      Caption         =   "Roll"
      Height          =   312
      Index           =   5
      Left            =   1560
      TabIndex        =   11
      Top             =   1800
      Width           =   972
   End
   Begin VB.TextBox txtDice 
      Height          =   288
      Index           =   5
      Left            =   0
      TabIndex        =   10
      Text            =   "d20"
      Top             =   1800
      Width           =   1572
   End
   Begin VB.CommandButton cmdRoll 
      Caption         =   "Roll"
      Height          =   312
      Index           =   4
      Left            =   1560
      TabIndex        =   9
      Top             =   1440
      Width           =   972
   End
   Begin VB.TextBox txtDice 
      Height          =   288
      Index           =   4
      Left            =   0
      TabIndex        =   8
      Text            =   "d12"
      Top             =   1440
      Width           =   1572
   End
   Begin VB.CommandButton cmdRoll 
      Caption         =   "Roll"
      Height          =   312
      Index           =   3
      Left            =   1560
      TabIndex        =   7
      Top             =   1080
      Width           =   972
   End
   Begin VB.TextBox txtDice 
      Height          =   288
      Index           =   3
      Left            =   0
      TabIndex        =   6
      Text            =   "d10"
      Top             =   1080
      Width           =   1572
   End
   Begin VB.CommandButton cmdRoll 
      Caption         =   "Roll"
      Height          =   312
      Index           =   2
      Left            =   1560
      TabIndex        =   5
      Top             =   720
      Width           =   972
   End
   Begin VB.TextBox txtDice 
      Height          =   288
      Index           =   2
      Left            =   0
      TabIndex        =   4
      Text            =   "d8"
      Top             =   720
      Width           =   1572
   End
   Begin VB.CommandButton cmdRoll 
      Caption         =   "Roll"
      Height          =   312
      Index           =   1
      Left            =   1560
      TabIndex        =   3
      Top             =   360
      Width           =   972
   End
   Begin VB.TextBox txtDice 
      Height          =   288
      Index           =   1
      Left            =   0
      TabIndex        =   2
      Text            =   "d6"
      Top             =   360
      Width           =   1572
   End
   Begin VB.CommandButton cmdRoll 
      Caption         =   "Roll"
      Height          =   312
      Index           =   0
      Left            =   1560
      TabIndex        =   1
      Top             =   0
      Width           =   972
   End
   Begin VB.TextBox txtDice 
      Height          =   288
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Text            =   "d4"
      Top             =   0
      Width           =   1572
   End
End
Attribute VB_Name = "frmDiceRoller"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdRoll_Click(Index As Integer)
    Dim s As String
    s = RollDice(txtDice(Index).Text)
    cmdRoll(Index).Caption = "Roll: " & s
    LogEvent "Roll: " & s & " (" & txtDice(Index).Text & ")"
End Sub

Public Sub LogEvent(ByVal sEvent As String)
    Dim pEnd As Boolean
    Dim pText As String
    pText = Format(Now, "hh:mm:ss") & vbTab & sEvent & vbCrLf
    pEnd = (txtResults.SelStart + txtResults.SelLength = Len(txtResults.Text))
    txtResults.Text = txtResults.Text & pText
    If pEnd Then
        txtResults.SelStart = Len(txtResults.Text) - Len(pText)
        txtResults.SelLength = Len(pText)
    End If
End Sub

Private Sub Form_Resize()
    Dim i As Integer
    Dim t As Single
    For i = 0 To cmdRoll.UBound
        txtDice(i).Move 0, txtDice(i).Height * i, Abs(Me.ScaleWidth - cmdRoll(i).Width)
        cmdRoll(i).Move Me.ScaleWidth - cmdRoll(i).Width, txtDice(i).Top
    Next
    t = txtDice(0).Height * txtDice.Count
    txtResults.Move 0, t, Me.ScaleWidth, Abs(Me.ScaleHeight - t)
End Sub
