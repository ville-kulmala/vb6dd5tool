Attribute VB_Name = "modTransparency"
Option Explicit

Public Const LWA_COLORKEY = 1
Public Const LWA_ALPHA = 2
Public Const LWA_BOTH = 3
Public Const WS_EX_LAYERED = &H80000
Public Const GWL_EXSTYLE = -20
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal Color As Long, ByVal X As Byte, ByVal alpha As Long) As Boolean
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long

Public Sub SetTransparent(Target As Form, Color As Long)
    SetTranslucent Target.hwnd, Color, 255, LWA_COLORKEY
End Sub

Public Sub SetTranslucent(ThehWnd As Long, Color As Long, nTrans As Integer, Flag As Byte)
    On Error GoTo ErrorRtn
    'SetWindowLong and SetLayeredWindowAttributes are API functions, see MSDN for details
    Dim attrib As Long
    attrib = GetWindowLong(ThehWnd, GWL_EXSTYLE)
    SetWindowLong ThehWnd, GWL_EXSTYLE, attrib Or WS_EX_LAYERED
    'anything with color value color will completely disappear if flag = 1 or flag = 3
    SetLayeredWindowAttributes ThehWnd, Color, nTrans, Flag
    Exit Sub
ErrorRtn:
    Debug.Print "SetTranslucent:", Err.Description & " Source : " & Err.Source
End Sub

