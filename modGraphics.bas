Attribute VB_Name = "modGraphics"
Option Explicit


Public Function PaintWithAspectRatio(ByVal Target As Object, ByVal Filename As String, X As Single, Y As Single, Width As Single, Height As Single)
    Dim P As StdPicture
    Dim sTargetAspectRatio As Single
    Dim sOrigAspectRatio As Single
    Dim sNewWidth As Single, sNewHeight As Single
    If Not FileExists(Filename) Then Exit Function
    Set P = LoadPicture(Filename)
    sOrigAspectRatio = P.Width / P.Height
    sTargetAspectRatio = Width / Height
    If sOrigAspectRatio < sTargetAspectRatio Then
        'Debug.Print "Kuva on korkeampi kuin ala => korkeus säilyy, kapeempi kuva"
        sNewWidth = Height * sOrigAspectRatio
        sNewHeight = Height
        Target.PaintPicture P, X + (Width - sNewWidth) / 2, Y, sNewWidth, sNewHeight
    Else
        'Debug.Print "Kuva on leveämpi kuin ala"
        sNewWidth = Width
        sNewHeight = Width / sOrigAspectRatio
        Target.PaintPicture P, X, Y + (Height - sNewHeight) / 2, sNewWidth, sNewHeight
    End If
End Function
