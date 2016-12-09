VERSION 5.00
Begin VB.Form frmDragFormTester 
   Caption         =   "Form1"
   ClientHeight    =   3084
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3084
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   240
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Scaling:"
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   240
      Width           =   735
   End
   Begin VB.Label lblPos 
      Caption         =   "Label1"
      Height          =   255
      Left            =   1440
      TabIndex        =   3
      Top             =   2640
      Width           =   2655
   End
   Begin VB.Label lblSize 
      Caption         =   "Label1"
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   2040
      Width           =   2175
   End
End
Attribute VB_Name = "frmDragFormTester"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents pMyDragForm As frmDragForm
Attribute pMyDragForm.VB_VarHelpID = -1

Private pScaling As Single

Private Sub Form_Load()
    pScaling = 15
    Set pMyDragForm = New frmDragForm
    pMyDragForm.Show , Me

End Sub

Private Sub pMyDragForm_Move(X As Single, Y As Single, PathLen As Single, Pause As Boolean)
    lblPos.Caption = Format(X / pScaling, "0.00") & ":" & Format(Y / pScaling, "0.00") & " moved " & Format(PathLen / pScaling, "0.00")
End Sub

Private Sub pMyDragForm_Resize(X As Single, Y As Single)
    lblSize = Format(X / pScaling, "0.00") & "x" & Format(Y / pScaling, "0.00") & ", width scale for 10':" & X / 10
End Sub

Private Sub Text1_Change()
    If IsNumeric(Text1.Text) Then
        pScaling = Text1.Text
    End If
End Sub
