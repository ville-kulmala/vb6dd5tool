VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMapBackground 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Kartta"
   ClientHeight    =   5730
   ClientLeft      =   4620
   ClientTop       =   3030
   ClientWidth     =   10935
   LinkTopic       =   "Form1"
   ScaleHeight     =   5730
   ScaleWidth      =   10935
   Begin VB.PictureBox Picture1 
      Height          =   492
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   1
      Top             =   120
      Width           =   492
   End
   Begin VB.PictureBox picHolder 
      Height          =   4215
      Left            =   0
      ScaleHeight     =   4155
      ScaleWidth      =   5955
      TabIndex        =   0
      Top             =   0
      Width           =   6015
      Begin VB.Shape shpRange 
         BorderColor     =   &H000000FF&
         Height          =   2052
         Left            =   1200
         Shape           =   3  'Circle
         Top             =   1080
         Width           =   2052
         Visible         =   0   'False
      End
      Begin VB.Line lneCone 
         BorderColor     =   &H000000FF&
         Index           =   3
         Visible         =   0   'False
         X1              =   2880
         X2              =   4680
         Y1              =   840
         Y2              =   1680
      End
      Begin VB.Line lneCone 
         BorderColor     =   &H000000FF&
         Index           =   2
         Visible         =   0   'False
         X1              =   2280
         X2              =   4080
         Y1              =   1320
         Y2              =   2160
      End
      Begin VB.Line lneCone 
         BorderColor     =   &H000000FF&
         Index           =   1
         Visible         =   0   'False
         X1              =   1800
         X2              =   3600
         Y1              =   1800
         Y2              =   2640
      End
      Begin VB.Line lneCone 
         BorderColor     =   &H000000FF&
         Index           =   0
         Visible         =   0   'False
         X1              =   1200
         X2              =   3000
         Y1              =   2640
         Y2              =   3480
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8280
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuHidden 
      Caption         =   "mnuHidden"
      Visible         =   0   'False
      Begin VB.Menu mnuHiddenLoadMap 
         Caption         =   "Load map..."
      End
      Begin VB.Menu mnuHiddenSaveMap 
         Caption         =   "Save map..."
      End
      Begin VB.Menu mnuHideenSetPicture 
         Caption         =   "Set picture..."
      End
      Begin VB.Menu mnuHiddenHome 
         Caption         =   "Home characters"
      End
      Begin VB.Menu mnuHiddenHomeMap 
         Caption         =   "Home map"
      End
      Begin VB.Menu mnuHiddenPlaceCharacter 
         Caption         =   "Place character"
      End
      Begin VB.Menu mnuHiddenZoom 
         Caption         =   "Zoom"
         Begin VB.Menu mnuHiddenZoomTo 
            Caption         =   "50%"
            Index           =   0
         End
         Begin VB.Menu mnuHiddenZoomTo 
            Caption         =   "100%"
            Index           =   1
         End
         Begin VB.Menu mnuHiddenZoomTo 
            Caption         =   "150%"
            Index           =   2
         End
         Begin VB.Menu mnuHiddenZoomTo 
            Caption         =   "200%"
            Index           =   3
         End
         Begin VB.Menu mnuHiddenZoomTo 
            Caption         =   "250%"
            Index           =   4
         End
      End
      Begin VB.Menu mnuHiddenAdd 
         Caption         =   "Add"
         Begin VB.Menu mnuHiddenAddSquare 
            Caption         =   "Square"
         End
         Begin VB.Menu mnuHiddenAddBalls 
            Caption         =   "Ball area"
            Begin VB.Menu mnuHiddenAddBall 
               Caption         =   "R 5'"
               Index           =   1
            End
            Begin VB.Menu mnuHiddenAddBall 
               Caption         =   "R 10'"
               Index           =   2
            End
            Begin VB.Menu mnuHiddenAddBall 
               Caption         =   "R 15'"
               Index           =   3
            End
            Begin VB.Menu mnuHiddenAddBall 
               Caption         =   "R 20'"
               Index           =   4
            End
            Begin VB.Menu mnuHiddenAddBall 
               Caption         =   "R 25'"
               Index           =   5
            End
            Begin VB.Menu mnuHiddenAddBall 
               Caption         =   "R 30'"
               Index           =   6
            End
            Begin VB.Menu mnuHiddenAddBall 
               Caption         =   "R 35'"
               Index           =   7
            End
            Begin VB.Menu mnuHiddenAddBall 
               Caption         =   "R 40'"
               Index           =   8
            End
         End
      End
      Begin VB.Menu mnuHiddenCone 
         Caption         =   "Cone"
      End
   End
End
Attribute VB_Name = "frmMapBackground"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pCharacters As Collection
Private pScaling As Single

Private pZooming As Single

Private pX As Single
Private pY As Single
Private pMDown As Boolean

Public Event Move(X As Single, Y As Single, dX As Single, dY As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Private pMode As String

Private pMyElements As New Collection

Public Initiative As frmInitiative

Private pMapPicture As String

Private pMapFile As String  'Avattu karttatiedosto

Public Property Get MyElements() As Collection
    Set MyElements = pMyElements
End Property

Public Function MapToString() As String
    MapToString = ConcatCol(GetMapRows)

End Function

Public Function GetMapRows(Optional AddToCol As Collection) As Collection
    Dim iElement As frmDragForm
    If AddToCol Is Nothing Then
        Set GetMapRows = New Collection
    Else
        Set GetMapRows = AddToCol
    End If
    With GetMapRows
        .Add "[map]"
        .Add "main picture : " & pMapPicture
        .Add "form location : " & Me.Left & ";" & Me.Top & ";" & Me.Width & ";" & Me.Height
        .Add "picture location : " & picHolder.Left & ";" & picHolder.Top & ";" & picHolder.Width & ";" & picHolder.Height
        .Add "scaling : " & Scaling
        For Each iElement In pMyElements
            iElement.GetElementRows GetMapRows
        Next
        GetCharacterStrings Initiative.Characters, GetMapRows
    End With
End Function

Public Sub PopMapRows(Rows As Collection)
    'TODO!!!
    'PopRows tyyppinen juttu tms.
    Dim sKey As String
    Dim sValue As String
    Dim e As frmDragForm
    Dim iChar As clsCharacter
    Dim c As Collection
    Do While Rows.Count > 0
        If LCase(Rows(1)) = "[map]" Then
            'Ihan vaan titteli...
        ElseIf LCase(Rows(1)) = "[element]" Then
            Set e = New frmDragForm
            If e.PopElementRows(Rows) Then
                pMyElements.Add e
                e.Show , Me
            Else
                Unload e
            End If
            If Rows.Count > 0 Then
                Rows.Add "", Before:=1
            End If
        ElseIf Left(Rows(1), 1) = "[" Then
            'Chracterit...
            Set c = ReadCharacters(Rows)
            For Each iChar In c
                Initiative.AddCharacter iChar
            Next
            Exit Do
        ElseIf Rows(1) = "" Then
        
        Else
            sValue = Rows(1)
            sKey = Trim(PopHead(sValue, ":"))
            Select Case LCase(sKey)
            Case "main picture":        SetMainPicture (Trim(sValue))
            Case "form location":       SetLocation Trim(sValue)
            Case "picture location":    SetBackgroundLocation Trim(sValue)
            Case "scaling":             Scaling = Trim(sValue)
                Initiative.txtScaling.Text = Scaling
            Case Else
                Debug.Print "Can't read row:" & Rows(1)
            End Select
        End If
        If Rows.Count > 0 Then
            Rows.Remove 1
        End If
    Loop
End Sub

Public Sub SetLocation(Location As String)
    Dim a As Variant
    a = Split(Location, ";")
    If UBound(a) = 3 Then
        Me.Move Trim(a(0)), Trim(a(1)), Trim(a(2)), Trim(a(3))
    End If
End Sub

Public Sub SetBackgroundLocation(Location As String)
    Dim a As Variant
    a = Split(Location, ";")
    If UBound(a) = 3 Then
        picHolder.Move Trim(a(0)), Trim(a(1)), Trim(a(2)), Trim(a(3))
    End If
End Sub

Public Function GetScaling() As Single
    GetScaling = pScaling
End Function

Public Sub SetCharacters(Characters As Collection)
    Set pCharacters = Characters
    ShowCharacters
End Sub

Public Sub ShowCharacters()
    Dim i As clsCharacter
    Dim f As frmDragForm
    If pCharacters Is Nothing Then Exit Sub
    For Each i In pCharacters
        Set f = i.GetDragForm(True)
        f.Hide
        f.Show , Me
        i.RefreshForm
    Next
End Sub

Public Property Let Scaling(ByVal Value As Single)
    Dim i As clsCharacter
    pScaling = Value
    For Each i In pCharacters
        i.Scaling = Value
        i.RefreshForm
    Next
End Property

Public Property Get Scaling() As Single
    Scaling = pScaling
End Property

Public Property Let Zooming(ByVal Value As Single)
    Dim i As clsCharacter
    Dim pX As Single, pY As Single, sX As Single, sY As Single
    'V‰hennet‰‰n t‰m‰n formin yl‰reuna.
    'TODO: Voi olla v‰h‰n karkea viel‰
    pX = Me.Left + (Me.Width - Me.ScaleWidth) / 2
    pY = Me.Top + (Me.Height - Me.ScaleHeight - (Me.Width - Me.ScaleWidth) / 2)
    
    For Each i In pCharacters
        With i.GetDragForm(True)
            
            .Move (.Left - pX) / pZooming * Value + pX, (.Top - pY) / pZooming * Value + pY
        End With
    Next
    With picHolder
        .Move .Left / pZooming * Value, .Top / pZooming * Value
    End With
    
    pZooming = Value
    For Each i In pCharacters
        i.Zooming = Value
        i.RefreshForm
    Next
    SetMainPicture pMapPicture
End Property

Public Property Get Zooming() As Single
    Zooming = pZooming
End Property

Public Sub HomeObjects()
    Dim i As clsCharacter
    Dim p As Integer
    For Each i In pCharacters
        p = p + 1
        With i.GetDragForm(True)
            .Width = pScaling * i.Size
            .Height = pScaling * i.Size
            .AllowResizing = False
            If .Left - .Width < Me.Left Or .Left > Me.Width + Me.Left Then
                .Left = Me.Left
                .Top = Me.Top + .Height * p
            End If
            If .Top - .Height < Me.Top Or .Top > Me.Top + Me.Height Then
                .Left = Me.Left
                .Top = Me.Top + .Height * p
            End If
        End With
    Next
End Sub

Private Sub Form_GotFocus()
    Set frmMenus.Characters = pCharacters
End Sub

Private Sub Form_Initialize()
    pZooming = 1
End Sub

Private Sub mnuHiddenAddBall_Click(Index As Integer)
    pMode = "ball:" & Index * 5
End Sub

Private Sub mnuHiddenAddSquare_Click()
    pMode = "square"
End Sub

Private Sub mnuHiddenCone_Click()
    pMode = "cone"
End Sub

Private Sub mnuHiddenHome_Click()
    HomeObjects
End Sub

Private Sub mnuHiddenHomeMap_Click()
    MoveHolder picHolder.Left, picHolder.Top
End Sub

Public Function GetOffsetX() As Single
    Dim bs As Single
    bs = (Me.ScaleWidth - Me.Width) / 2
    GetOffsetX = Me.Left + picHolder.Left - bs
End Function

Public Function GetOffsetY() As Single
    Dim bU As Single
    bU = Me.ScaleHeight - Me.Height - (Me.ScaleWidth - Me.Width)
    
    GetOffsetY = Me.Top + picHolder.Top - bU
End Function

Private Sub mnuHiddenLoadMap_Click()
    On Error Resume Next
    With CommonDialog1
        .CancelError = True
        .Filter = "Map files|*.map"
        .ShowOpen
        If Err = 0 Then
            LoadMap .Filename
        End If
    End With
End Sub

Public Sub LoadMap(ByVal Filename As String)
    Dim Rows As Collection
    Set Rows = FileToCol(Filename)
    If Rows.Count > 1 Then
        Me.PopMapRows Rows
        'TODO: hahmojen luku j‰ljelle j‰‰neist‰
        
        pMapFile = Filename
    End If

End Sub

Private Sub mnuHiddenPlaceCharacter_Click()
    On Error Resume Next
    With Initiative.SelectedCharacter.GetDragForm(True)
        .Move GetOffsetX + pX, GetOffsetY + pY
        '.Location = Int(pX / pScaling) & ":" & Int(pY / pScaling)
        '.MoveToPosition
        .Visible = True ' Jos oli ulkopuolella, ei n‰kynyt
    End With
End Sub

Private Sub mnuHiddenSaveMap_Click()
    On Error Resume Next
    Dim Rows As Collection
    With CommonDialog1
        .CancelError = True
        .Filter = "Map files|*.map"
        .Filename = pMapFile
        .ShowSave
        If Err = 0 Then
            Set Rows = GetMapRows
            'TODO: characterit kans
            WriteColToFile .Filename, Rows
            pMapFile = .Filename
        End If
    End With
End Sub

Private Sub mnuHiddenZoomTo_Click(Index As Integer)
    Me.Zooming = 0.5 + (Index * 0.5)
End Sub

Private Sub mnuHideenSetPicture_Click()
    On Error Resume Next
    With CommonDialog1
        .CancelError = True
        .Filter = "Picture files|*.bmp;*.wmf;*.jpg;*.gif"
        .ShowOpen
        If Err = 0 Then
            SetMainPicture .Filename
        End If
    End With
End Sub

Public Sub SetMainPicture(ByVal Filename As String)
    Dim p As StdPicture
    If Filename = "" Then
        Exit Sub
    End If
    Set p = LoadPicture(Filename)
    With picHolder
        On Error GoTo 0
        pMapPicture = Filename
        .AutoRedraw = True
        .Width = .ScaleX(p.Width, vbTwips, vbHimetric) * pZooming
        .Height = .ScaleY(p.Height, vbTwips, vbHimetric) * pZooming
        picHolder.PaintPicture p, 0, 0, .ScaleWidth, .ScaleHeight
        
        '.Width = ScaleX(picHolder.Width, vbTwips, vbHimetric) * .Picture.Width
        '.Height = ScaleY(picHolder.Height, vbTwips, vbHimetric) * .Picture.Height
    End With
    
End Sub

Private Sub picHolder_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If pMode <> "" Then
        DoAddElement X, Y
    Else
        If Button = vbLeftButton Then
            pMDown = True
            pX = X
            pY = Y
        End If
    End If
End Sub

Private Sub picHolder_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim dX As Single, dY As Single
    Select Case Button
    Case vbLeftButton
        'TODO: t‰m‰ aika turhaa. Voitaisiin tehd‰ temput kun tiedet‰‰n napin olevan alhaalla
        If pMDown Then
            dX = pX - X
            dY = pY - Y
            MoveHolder dX, dY
        Else
            RaiseEvent MouseMove(Button, Shift, X, Y)
        End If
    Case Else
        Select Case pMode
        Case "cone"
            ConeMouseMove Button, Shift, X, Y
        End Select
        RaiseEvent MouseMove(Button, Shift, X, Y)
    End Select
End Sub

Public Sub ConeMouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim pSelPos As Point
    Dim i As Integer
    If Initiative.CurInitiative Is Nothing Then
        For i = 0 To lneCone.UBound
            lneCone(i).Visible = False
        Next
        pMode = ""
        Exit Sub
    End If
    Dim d As frmDragForm
    Set d = Initiative.CurInitiative.GetDragForm(True)
    With d
        pSelPos.X = GetFormOnMapX(d) + .Width / 2
        pSelPos.Y = GetFormOnMapY(d) + .Height / 2
    End With
    
    With lneCone(0)
        .X1 = pSelPos.X
        .Y1 = pSelPos.Y
        .X2 = X
        .Y2 = Y
    End With
    With lneCone(1)
        .X1 = X + (Y - pSelPos.Y) / 2
        .Y2 = Y + (X - pSelPos.X) / 2
        .X2 = X - (Y - pSelPos.Y) / 2
        .Y1 = Y - (X - pSelPos.X) / 2
    End With
    With lneCone(2)
        .X1 = pSelPos.X
        .Y1 = pSelPos.Y
        .X2 = lneCone(1).X2
        .Y2 = lneCone(1).Y2
    End With
    With lneCone(3)
        .X1 = pSelPos.X
        .Y1 = pSelPos.Y
        .X2 = lneCone(1).X1
        .Y2 = lneCone(1).Y1
    End With
    For i = 0 To lneCone.UBound
        lneCone(i).Visible = True
    Next
    Me.Caption = "Cone length: " & Format(Sqr(Abs(lneCone(0).X1 - lneCone(0).X2) ^ 2 + Abs(lneCone(0).Y1 - lneCone(0).Y2) ^ 2) / pScaling - 2.5, "0") & " (" & GetConeCharacters.Count & " characters)"
End Sub

Public Sub MoveHolder(ByVal dX As Single, dY As Single)
    With picHolder
        picHolder.Move .Left - dX, .Top - dY
    End With
    MoveElements dX, dY
    RaiseEvent Move(Me.Left + Me.Width / 2, Me.Top + Me.Height / 2, dX, dY)

End Sub

Private Sub picHolder_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'TODO: pMDown on aika turhaa. Button kertoo kyll‰
    If Not pMDown Then
        Select Case Button
        Case vbLeftButton

        Case vbRightButton
            pX = X  'Placement info
            pY = Y
            'Jos tehd‰‰n area of effect, tarvitaan lista hahmoista
            Set frmMenus.Characters = pCharacters
            Set frmMenus.MapBackground = Me
            Me.PopupMenu mnuHidden
        End Select
    Else
        Select Case Button
        Case vbLeftButton
            Select Case pMode
            End Select
        End Select
    End If
    
    pMDown = False
    
End Sub
Private Function GetConeCharacters() As Collection
    Dim pChar As Point, T1 As Point, T2 As Point, T0 As Point
    Dim Character As clsCharacter
    Dim pSel As Collection
    Set pSel = New Collection

    
    T0.X = GetOffsetX + lneCone(0).X1
    T0.Y = GetOffsetY + lneCone(0).Y1
    T1.X = GetOffsetX + lneCone(1).X1
    T1.Y = GetOffsetY + lneCone(1).Y1
    T2.X = GetOffsetX + lneCone(1).X2
    T2.Y = GetOffsetY + lneCone(1).Y2
    For Each Character In Initiative.Characters
        If Character.IsActive Then
            With Character.GetDragForm(True)
                pChar.X = .Left + .Width / 2
                pChar.Y = .Top + .Height / 2
            End With
            If IsPointInTriangle(pChar, T0, T1, T2) Then
                pSel.Add Character
            End If
        End If
    Next
    Set GetConeCharacters = pSel
End Function

Private Sub ConeMouseUp(X As Single, Y As Single)
    Dim pSel As Collection
    Dim i As Integer
    pMode = ""
    Set pSel = GetConeCharacters
    If pSel.Count > 0 Then
        Dim d As dlgCharacterList
        Set d = New dlgCharacterList
        d.ShowDialog pSel, Me
        ShowCharacters
    End If
    For i = 0 To lneCone.UBound
        lneCone(i).Visible = False
    Next
End Sub

Public Sub EnsureVisible(Character As clsCharacter)
    If Character Is Nothing Then Exit Sub
    Dim dX As Single, dY As Single
    With Character.GetDragForm(True)
        If .Left + .Width < Me.Left Then
            dX = .Left - Me.Left
        ElseIf .Left > Me.Left + Me.Width Then
            dX = ((.Left + .Width / 2) - (Me.Left + Me.Width / 2))
        End If
        If Me.Top > .Top + .Height Then
            dY = .Top - Me.Top
        ElseIf .Top > Me.Height + Me.Top Then
            dY = ((.Top + .Width / 2) - (Me.Top + Me.Width / 2))
        End If
        MoveHolder dX, dY
        'Ympyr‰ hahmon alle!
        PlaceCircle Character
    End With
End Sub

Public Sub PlaceCircle(Character As clsCharacter)
    Dim X As Single
    Dim Y As Single
    If Character Is Nothing Then
        shpRange.Visible = False
    Else
        X = GetFormOnMapX(Character.GetDragForm(True)) + Character.GetDragForm(True).Width / 2
        Y = GetFormOnMapY(Character.GetDragForm(True)) + Character.GetDragForm(True).Height / 2
        shpRange.Move X - pScaling * 30, Y - pScaling * 30, pScaling * 60, pScaling * 60
        shpRange.Visible = True
    End If
End Sub


Private Sub Picture1_Click()
    'Mouseupissa
    'Me.PopupMenu mnuHidden
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        pMDown = True
        pX = X
        pY = Y
    End If
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim dX As Single, dY As Single
    If pMDown Then
        dX = pX - X
        dY = pY - Y
        Me.Move Me.Left - dX, Me.Top - dY
        RaiseEvent Move(Me.Left + Me.Width / 2, Me.Top + Me.Height / 2, dX, dY)
        MoveElements dX, dY
    End If
End Sub

Private Sub MoveElements(ByVal X As Single, ByVal Y As Single)
    Dim i As frmDragForm
    For Each i In pMyElements
        i.Move i.Left - X, i.Top - Y
        i.Visible = IsWithinArea(i, Me)
    Next
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not pMDown Then
        If Button = vbRightButton Then
            'Jos tehd‰‰n pallo, tarvitaan lista hahmoista
            Set frmMenus.Characters = pCharacters
            Set frmMenus.MapBackground = Me
            Me.PopupMenu mnuHidden
        End If
    End If
    pMDown = False
End Sub

Private Sub DoAddElement(X As Single, Y As Single)
    Dim h As String
    Dim t As String
    Dim df As frmDragForm
    h = PopHead(pMode, ":")
    t = pMode
    Select Case LCase(h)
    Case "ball"
        pMyElements.Add CreateBall(X + GetOffsetX, Y + GetOffsetY, CSng(t) * pScaling, vbRed, Me)
    Case "square"
        If t = "" Then t = 20
        Set df = CreateSquare(X + GetOffsetX, Y + GetOffsetY, CSng(t) * pScaling, vbBlack, Me)
        pMyElements.Add df
        'Ei peit‰ hahmoja alleen jos nekin eiv‰t ole setparentoitu
        ' kaikki tai ei mit‰‰n
        'SetParent df.hWnd, me.hwnd
    Case "cone"
        ConeMouseUp X, Y
    End Select
    pMode = ""
End Sub

Private Function GetFormOnMapX(f As Form) As Single
    If f Is Nothing Then
        GetFormOnMapX = 0
        Exit Function
    End If
    GetFormOnMapX = f.Left - Me.Left + (Me.ScaleWidth - Me.Width) / 2 - picHolder.Left
End Function

Private Function GetFormOnMapY(f As Form) As Single
    If f Is Nothing Then
        GetFormOnMapY = 0
        Exit Function
    End If
    GetFormOnMapY = f.Top - Me.Top - (Me.Height - Me.ScaleHeight + (Me.ScaleWidth - Me.Width) / 2) - picHolder.Top
End Function

