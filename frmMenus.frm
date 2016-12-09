VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMenus 
   Caption         =   "Form1"
   ClientHeight    =   7956
   ClientLeft      =   1344
   ClientTop       =   1692
   ClientWidth     =   5604
   LinkTopic       =   "Form1"
   ScaleHeight     =   7956
   ScaleWidth      =   5604
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1560
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu pmnuDragForm 
      Caption         =   "DragFormPopupMenu"
      Begin VB.Menu pmnuDragFormCharacterList 
         Caption         =   "Characters within..."
      End
      Begin VB.Menu pmnuDragFormColor 
         Caption         =   "Color..."
      End
      Begin VB.Menu pmnuDragFormPicture 
         Caption         =   "Picture..."
      End
      Begin VB.Menu pmnuDragFormClose 
         Caption         =   "Close"
      End
   End
End
Attribute VB_Name = "frmMenus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Usage: set target form to this
' "global" form and call "popupmenu"

Public MapBackground As frmMapBackground
Public DragForm As frmDragForm
Public Characters As Collection

'SET: CharacterList!
Private Sub pmnuDragFormCharacterList_Click()
    Dim c As Collection
    Dim d As dlgCharacterList
    On Error GoTo ErrHandler
    With Me.DragForm

        Select Case .AreaType
        Case "ball"
            Set c = GetCharactersNearPoint(Characters, .Left + .Width / 2, .Top + .Width / 2, .Width / 2, False, True)
        Case Else
            Set c = GetCharactersInSquare(Characters, .Left, .Top, .Width, .Height, False, True)
        End Select
        
        If c.Count > 0 Then
            Set d = New dlgCharacterList
            d.ShowDialog c, DragForm
            MapBackground.ShowCharacters
        Else
            MsgBox ("No characters within area")
        End If
        
    End With
    Exit Sub
ErrHandler:
    Debug.Print "DragFormCharacterList", Err, Err.Description
End Sub

Private Sub pmnuDragFormClose_Click()
    Dim i As Integer
    DragForm.Hide
    For i = 1 To MapBackground.MyElements.Count
        If DragForm Is MapBackground.MyElements(i) Then
            MapBackground.MyElements.Remove i
            Exit For
        End If
    Next
    Unload DragForm
End Sub

Private Sub pmnuDragFormColor_Click()
    Dim Color As Long
    On Error Resume Next
    With CommonDialog1
        .Color = DragForm.BackColor
        .ShowColor
        Color = .Color
        Select Case DragForm.AreaType
        Case "ball"
            With DragForm
                .BackColor = Abs(Color - 255)
                .DrawFilledCircle Color, Color
                .SetTranslucent .BackColor, .Transparency, LWA_BOTH
            End With
        Case Else
            DragForm.BackColor = .Color
        End Select
    End With
End Sub

Private Sub pmnuDragFormPicture_Click()
    On Error Resume Next
    With CommonDialog1
        .Filter = "Picture files|*.jpg;*.bmp;*.wmg;*.gif;*.ico"
        .Filename = DragForm.PictureFile
        .CancelError = True
        .ShowOpen
        If Err = 0 Then
            DragForm.PictureFile = .Filename
        End If
    End With
End Sub
