VERSION 5.00
Begin VB.Form frmImmediate 
   Caption         =   "Immediate"
   ClientHeight    =   7950
   ClientLeft      =   3045
   ClientTop       =   1755
   ClientWidth     =   5610
   LinkTopic       =   "Form1"
   ScaleHeight     =   7950
   ScaleWidth      =   5610
   Begin VB.TextBox txtImmediate 
      Height          =   2412
      Left            =   720
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   840
      Width           =   4092
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuEditRollCharacter 
         Caption         =   "Roll Character stats"
      End
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

Private Sub mnuEditRollCharacter_Click()
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim w As Integer
    Dim t As Integer
    Dim d As Integer
    Dim tot(1 To 5) As Integer
    Randomize Timer
    txtImmediate.Text = txtImmediate.Text & Now & " Rolling character: " & vbCrLf
    For i = 1 To 6
        For j = 1 To 5
            t = 0
            w = 7
            For k = 1 To 4
                d = RollDie("d6")
                If d < w Then
                    w = d
                End If
                t = d + t
            Next
            t = t - w
            tot(j) = tot(j) + t
            txtImmediate.Text = txtImmediate.Text & vbTab & t
        Next
        txtImmediate.Text = txtImmediate.Text & vbCrLf
    Next
    txtImmediate.Text = txtImmediate.Text & "              ===================================" & vbCrLf
    For i = 1 To 5
        txtImmediate.Text = txtImmediate.Text & vbTab & tot(i)
    Next
    txtImmediate.Text = txtImmediate.Text & vbCrLf
End Sub
