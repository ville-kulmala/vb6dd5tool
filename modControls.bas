Attribute VB_Name = "modControls"
Option Explicit

Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long
Public Const HWND_BOTTOM = 1
Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOP = 0
Public Const HWND_TOPMOST = -1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOREDRAW = &H8
Public Const SWP_NOSIZE = &H1
Public Const SWP_SHOWWINDOW = &H40
Public Const SWP_NOACTIVATE = &H10



'Palauttaa saman jos peruttiin, muuten muutetun tekstin
Public Function ShowTextEditor(Text As String, Optional OwnerForm As Form, Optional Caption As String) As String
    Dim d As dlgText
    Set d = New dlgText
    ShowTextEditor = d.ShowDialog(Text, OwnerForm, Caption)
End Function
