VERSION 5.00
Begin VB.Form frmImmediate 
   Caption         =   "Immediate"
   ClientHeight    =   7956
   ClientLeft      =   3048
   ClientTop       =   1452
   ClientWidth     =   5604
   LinkTopic       =   "Form1"
   ScaleHeight     =   7956
   ScaleWidth      =   5604
   Begin VB.TextBox txtImmediate 
      Height          =   2412
      Left            =   720
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   840
      Width           =   4092
   End
End
Attribute VB_Name = "frmImmediate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()
    txtImmediate.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Public Sub LogEvent(ByVal sEvent As String)
    Dim pEnd As Boolean
    Dim pText As String
    pText = Format(Now, "hh:mm:ss") & vbTab & sEvent & vbCrLf
    pEnd = (txtImmediate.SelStart + txtImmediate.SelLength = Len(txtImmediate.Text))
    txtImmediate.Text = txtImmediate.Text & pText
    If pEnd Then
        txtImmediate.SelStart = Len(txtImmediate.Text) - Len(pText)
        txtImmediate.SelLength = Len(pText)
    End If
End Sub

