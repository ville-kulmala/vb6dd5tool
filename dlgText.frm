VERSION 5.00
Begin VB.Form dlgText 
   Caption         =   "Form1"
   ClientHeight    =   3084
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3084
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   2640
      Width           =   975
   End
   Begin VB.TextBox txtText 
      Height          =   2295
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   3975
   End
End
Attribute VB_Name = "dlgText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event SaveClicked(ByVal Text As String)
Private pSaveClicked As Boolean

Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    RaiseEvent SaveClicked(txtText.Text)
    pSaveClicked = True
    Me.Hide
End Sub

Public Function ShowDialog(ByVal Text As String, Optional OwnerForm As Form, Optional Caption As String) As String
    txtText.Text = Text
    Me.Caption = Caption
    If OwnerForm Is Nothing Then
        Me.Show 1
    Else
        Me.Show 1, OwnerForm
    End If
    If pSaveClicked Then
        ShowDialog = txtText.Text
    Else
        ShowDialog = Text
    End If
End Function

Private Sub Form_Resize()
    Const MARGIN = 90
    cmdOK.Move Me.ScaleWidth - cmdOK.Width - MARGIN, Me.ScaleHeight - cmdOK.Height - MARGIN
    cmdCancel.Move Me.ScaleWidth - cmdOK.Width - cmdCancel.Width - MARGIN * 2, Me.ScaleHeight - MARGIN - cmdCancel.Height
    txtText.Move 0, 0, Me.ScaleWidth, cmdCancel.Top - MARGIN
    
End Sub
