VERSION 5.00
Begin VB.Form frmDragForm 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   3012
   ClientLeft      =   1488
   ClientTop       =   2100
   ClientWidth     =   3432
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3012
   ScaleWidth      =   3432
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picResize 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   2055
      Index           =   3
      Left            =   0
      ScaleHeight     =   2052
      ScaleWidth      =   432
      TabIndex        =   3
      Top             =   480
      Width           =   435
   End
   Begin VB.PictureBox picResize 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   2
      Left            =   360
      ScaleHeight     =   372
      ScaleWidth      =   2832
      TabIndex        =   2
      Top             =   2640
      Width           =   2835
   End
   Begin VB.PictureBox picResize 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   1935
      Index           =   1
      Left            =   2880
      ScaleHeight     =   1932
      ScaleWidth      =   432
      TabIndex        =   1
      Top             =   720
      Width           =   435
   End
   Begin VB.PictureBox picResize 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   0
      Left            =   360
      ScaleHeight     =   372
      ScaleWidth      =   2832
      TabIndex        =   0
      Top             =   0
      Width           =   2835
   End
   Begin VB.Shape shpHighlight 
      Height          =   972
      Left            =   840
      Top             =   1080
      Width           =   1332
      Visible         =   0   'False
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      Caption         =   "Label1"
      ForeColor       =   &H8000000E&
      Height          =   192
      Left            =   90
      TabIndex        =   4
      Top             =   90
      Width           =   492
      Visible         =   0   'False
   End
End
Attribute VB_Name = "frmDragForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pX As Single
Private pY As Single
Private pMDown As Boolean
Private pPath As Single
Private pStartX As Single
Private pStartY As Single

Private cMoves As New Collection    'x;y muodossa
Private pMoveRedo As New Collection

Public Event Resize(X As Single, Y As Single)
Public Event Move(X As Single, Y As Single, PathLen As Single, Pause As Boolean)
Public Event Clicked()
Public Event StepTo(ByVal Direction As String)
Public AreaType As String   'Square (default), ball etc.

Private pTransparency As Integer    '255 fully transparent
Private pAlphaColor As Long
Private pPictureFile As String
Private pHighlighted As Boolean
Public Function GetElementRows(Optional AddToCol As Collection) As Collection
    If AddToCol Is Nothing Then
        Set GetElementRows = New Collection
    Else
        Set GetElementRows = AddToCol
    End If
    With GetElementRows
        .Add "[element]"
        .Add "transparency : " & pTransparency
        .Add "alphacolor : " & pAlphaColor
        .Add "areatype : " & AreaType
        .Add "location : " & Me.Left & ";" & Me.Top & ";" & Me.Width & ";" & Me.Height
        .Add "backcolor : " & Me.BackColor
        .Add "fillcolor : " & Me.FillColor
        .Add "picturefile : " & Me.PictureFile
    End With
End Function

Public Function PopElementRows(Rows As Collection)
    Dim sKey As String
    Dim sValue As String
    Do While Rows.Count > 0
        If LCase(Rows(1)) = "[element]" Then
            'Ihan vaan titteli...
            If PopElementRows Then
                'Exitoidaan ennen kuin mennään seuraavaan
                Exit Function
            Else
                PopElementRows = True
            End If
        ElseIf Left(Rows(1), 1) = "[" Then
            'Chracterit...
            Exit Do
        ElseIf Rows(1) = "" Then
            'Tyhjät unohdetaan
        Else
            sValue = Rows(1)
            sKey = Trim(PopHead(sValue, ":"))
            Select Case LCase(sKey)
            Case "transparency":    pTransparency = Trim(sValue)
            Case "alphacolor":      pAlphaColor = Trim(sValue)
            Case "areatype":        AreaType = Trim(sValue)
            Case "location":        SetLocation Trim(sValue)
            Case "backcolor":       Me.BackColor = Trim(sValue)
            Case "fillcolor":       Me.FillColor = Trim(sValue)
            Case "picturefile":     Me.PictureFile = Trim(sValue)
            Case Else
                Debug.Print "Can't read row:" & Rows(1)
            End Select
        End If
        Rows.Remove 1
    Loop
    Select Case AreaType
    Case "ball"
        DrawFilledCircle Me.FillColor, Me.FillColor
        AllowResizing = False
    Case Else
    End Select
    If pTransparency <> 0 Then
        SetTranslucent pAlphaColor, pTransparency, LWA_BOTH
    End If
    
End Function

Public Sub SetLocation(Location As String)
    Dim a As Variant
    a = Split(Location, ";")
    If UBound(a) = 3 Then
        Me.Move Trim(a(0)), Trim(a(1)), Trim(a(2)), Trim(a(3))
    End If
End Sub


Public Property Get AllowResizing() As Boolean
    AllowResizing = picResize(0).Visible
End Property

Public Property Let AllowResizing(Value As Boolean)
    Dim i As Integer
    For i = 0 To 3
        picResize(i).Visible = Value
    Next
End Property

Public Property Get CenterX() As Single
    CenterX = Me.Left + Me.Width / 2
End Property

Public Property Get CenterY() As Single
    CenterY = Me.Top + Me.Height / 2
End Property

Private Sub Form_Click()
    RaiseEvent Clicked
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    'Ei tule viestiä => ei toimi
    Dim rPause As Boolean
    Dim pStep As Single
    
    If KeyCode = vbKeyEscape Then
        If pMDown Then
            pMDown = False
            Me.Move pStartX, pStartY
    
        End If
    ElseIf KeyCode = vbKeyZ And Shift = vbCtrlMask Then
        UndoMove
    ElseIf KeyCode = vbKeyY And Shift = vbCtrlMask Then
        RedoMove
    ElseIf KeyCode = vbKeyA Then
        'TODO: implementaatio siirroille. Myös väli-ilmansuunnat
        ' voisi laskea koko kierroksen liikettä: OSITTAIN TEHTY
        'TODO: Slots (Spell slots yms.)
        'TODO: linkkiosuus (web-sivu) hahmoon
        'TODO: palautuvat powerit (kuten slotit)
        'TODO: picture folder, character folder yms.: PICTUREFOLDER TEHTY...
        'TODO: Altitude
        RaiseEvent StepTo("w")
        Me.Move Me.Left - Me.Width, Me.Top
        pPath = pPath + Me.Width
        RaiseEvent Move(Me.Left + Me.Width / 2, Me.Top + Me.Height / 2, pPath, rPause)
        
    ElseIf KeyCode = vbKeyS Then
        RaiseEvent StepTo("e")
        Me.Move Me.Left + Me.Width
        pPath = pPath + Me.Width
        RaiseEvent Move(Me.Left + Me.Width / 2, Me.Top + Me.Height / 2, pPath, rPause)

    ElseIf KeyCode = vbKeyW Then
        RaiseEvent StepTo("n")
        Me.Move Me.Left, Me.Top - Me.Height
        pPath = pPath + Me.Height
        RaiseEvent Move(Me.Left + Me.Width / 2, Me.Top + Me.Height / 2, pPath, rPause)
    
    ElseIf KeyCode = vbKeyZ Then
        RaiseEvent StepTo("s")
        Me.Move Me.Left, Me.Top + Me.Height
        pPath = pPath + Me.Height
        RaiseEvent Move(Me.Left + Me.Width / 2, Me.Top + Me.Height / 2, pPath, rPause)
    
    End If
End Sub

Private Sub Form_Load()
    AllowResizing = True
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        pMDown = True
        pX = X
        pY = Y
        Set pMoveRedo = New Collection
        pStartX = Me.Left
        pStartY = Me.Top
        pPath = 0
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim dX As Single, dY As Single
    Dim rPause As Boolean
    If pMDown Then
        dX = pX - X
        dY = pY - Y
        pPath = pPath + Sqr(Abs(dX ^ 2) + Abs(dY ^ 2))
        Me.Move Me.Left - dX, Me.Top - dY
        RaiseEvent Move(Me.Left + Me.Width / 2, Me.Top + Me.Height / 2, pPath, rPause)
        If rPause Then
            pPath = pPath - Sqr(Abs(dX ^ 2) + Abs(dY ^ 2))
            Me.Move Me.Left + dX, Me.Top + dY
            RaiseEvent Move(Me.Left + Me.Width / 2, Me.Top + Me.Height / 2, pPath, rPause)
            pMDown = False
        End If
    Else
        If Me.Caption <> "" Then
            'Jotan tuultip juttua ois ehkä kiva
        End If
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    pMDown = False
    If Button = vbRightButton Then
        Set frmMenus.DragForm = Me
        
        Me.PopupMenu frmMenus.pmnuDragForm
        Set frmMenus.DragForm = Nothing
    Else
        If pStartX <> Me.Left And pStartY <> Me.Top Then
            cMoves.Add pStartX & ";" & pStartY
        End If
    End If
End Sub

Private Sub Form_Resize()
    Const BAR_WIDTH = 90
    picResize(0).Move 0, 0, Me.ScaleWidth, BAR_WIDTH
    picResize(1).Move Me.ScaleWidth - BAR_WIDTH, 0, BAR_WIDTH, Me.ScaleHeight
    picResize(2).Move 0, Me.ScaleHeight - BAR_WIDTH, Me.ScaleWidth, BAR_WIDTH
    picResize(3).Move 0, 0, BAR_WIDTH, Me.ScaleHeight
    RaiseEvent Resize(Me.Width, Me.Height)
    
End Sub

Private Sub lblStatus_Click()
    Form_Click
End Sub

Private Sub lblStatus_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_MouseDown Button, Shift, X + lblStatus.Left, Y + lblStatus.Top
End Sub

Private Sub lblStatus_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_MouseMove Button, Shift, X + lblStatus.Left, Y + lblStatus.Top
End Sub

Private Sub lblStatus_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_MouseUp Button, Shift, X + lblStatus.Left, Y + lblStatus.Top
End Sub

Private Sub picResize_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    pX = X
    pY = Y
    pMDown = True
End Sub

Private Sub picResize_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim dX As Single, dY As Single
    If pMDown Then
        Select Case Index
        Case 0  'Top
            dY = pY - Y
            If Me.Height + dY > 90 Then
                Me.Height = Me.Height + dY
                Me.Top = Me.Top - dY
            End If
        Case 1  'Right
            dX = pX - X
            If Me.Width + dX > 90 Then
                Me.Width = Me.Width - dX
            End If
        Case 2  'Bottom
            dY = pY - Y
            If Me.Height - dY > 90 Then
                Me.Height = Me.Height - dY
            End If
        Case 3  'Left
            dX = pX - X
            If Me.Width + dX > 90 Then
                Me.Width = Me.Width + dX
                Me.Left = Me.Left - dX
            End If
        End Select
    End If
End Sub

Private Sub picResize_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    pMDown = False
End Sub

Private Sub UndoMove()
    Dim P As Variant
    If cMoves.Count > 0 Then
        P = cMoves(cMoves.Count)
        cMoves.Remove cMoves.Count
        pMoveRedo.Add Me.Left & ";" & Me.Top
        Me.Move Left(P, InStr(P, ";") - 1), Mid(P, InStr(P, ";") + 1)
        
    End If
End Sub

Private Sub RedoMove()
    Dim P As Variant
    If pMoveRedo.Count > 0 Then
        P = pMoveRedo(pMoveRedo.Count)
        pMoveRedo.Remove pMoveRedo.Count
        Me.Move Left(P, InStr(P, ";") - 1), Mid(P, InStr(P, ";") + 1)
        cMoves.Add P
    End If
End Sub

Public Sub DrawFilledCircle(Color As Long, FillColor As Long)
    Me.FillColor = FillColor
    Me.FillStyle = vbSolid
    Me.Circle (Me.ScaleWidth / 2, Me.ScaleHeight / 2), Min(Me.ScaleWidth / 2, Me.ScaleHeight / 2), Color
End Sub

Public Sub SetTransparency(Color As Long)
    modTransparency.SetTransparent Me, Color
    
End Sub

Public Sub SetTranslucent(Color As Long, nTrans As Integer, LWA_Flag As Byte)
    pAlphaColor = Color
    pTransparency = nTrans
    modTransparency.SetTranslucent Me.hwnd, Color, nTrans, LWA_Flag
End Sub

Public Property Get Transparency() As Integer
    Transparency = pTransparency
End Property

Public Property Get PictureFile() As String
    PictureFile = pPictureFile
End Property

Public Property Let PictureFile(ByVal Value As String)
    pPictureFile = Value
    On Error Resume Next
    If pPictureFile = "" Then Exit Property
    Me.PaintPicture LoadPicture(Value), 0, 0, Me.ScaleWidth, Me.ScaleHeight
    
    If Err <> 0 Then
        MsgBox "Kuvaa ei voitu ladata. Kuva: " & Value & vbCrLf & Err.Description
    End If
    Err.Clear
End Property

Public Property Let Highlighted(ByVal Value As Boolean)
    pHighlighted = Value
    If pHighlighted Then
        With shpHighlight
            .BorderWidth = 5
            .BorderColor = vbHighlight
            .Visible = True
            .Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
        End With
        'Helpompi tarttua / tämä hoidetaan muualla.
        'modTransparency.SetTranslucent Me.hwnd, pAlphaColor, 1, LWA_BOTH
    Else
        With shpHighlight
            .Visible = False
            .Move -2, -2, 2, 2
        End With
        'modTransparency.SetTranslucent Me.hwnd, pAlphaColor, pTransparency, LWA_BOTH
    End If
End Property

Public Property Get Highlighted() As Boolean
    Highlighted = pHighlighted
End Property
