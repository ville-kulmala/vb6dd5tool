Attribute VB_Name = "modMap"
Option Explicit

Public Const MAP_PICTUREFILE As String = "picturefile"
Public Const MAP_SCALING As String = "scaling"
Public Const MAP_CANVASLOCATION As String = "canvaslocation"
Public Const MAP_FORMLOCATION As String = "formlocation"

Public Function CreateBall(X As Single, Y As Single, R As Single, Color As Long, Parent As Form) As frmDragForm
    Dim f As frmDragForm
    Set f = New frmDragForm
    If Parent Is Nothing Then
        f.Show
    Else
        f.Show , Parent
    End If
    f.Move X - R, Y - R, R * 2, R * 2
    f.BackColor = Abs(Color - 255)
    f.DrawFilledCircle Color, Color
    f.SetTranslucent f.BackColor, 168, LWA_BOTH
    f.AllowResizing = False
    f.AreaType = "ball"
    Set CreateBall = f
End Function

Public Function CreateSquare(X As Single, Y As Single, R As Single, Color As Long, Parent As Form) As frmDragForm
    Dim f As frmDragForm
    Set f = New frmDragForm
    If Parent Is Nothing Then
        f.Show
    Else
        f.Show , Parent
    End If
    f.Move X - R, Y - R, R * 2, R * 2
    f.BackColor = Color
    f.AllowResizing = True
    Set CreateSquare = f
End Function

Public Function GetCharactersNearPoint(AllCharacters As Collection, ByVal X As Single, ByVal Y As Single, ByVal MaxDistance As Single, ByVal ActiveOnly As Boolean, ByVal ByForm As Boolean) As Collection
    Dim iCharacter As clsCharacter
    Set GetCharactersNearPoint = New Collection
    For Each iCharacter In AllCharacters
        If iCharacter.IsActive Or Not ActiveOnly Then
            If iCharacter.GetDistanceToPoint(X, Y, ByForm) <= MaxDistance Then
                GetCharactersNearPoint.Add iCharacter
            End If
        End If
    Next
End Function

Public Function GetCharactersInSquare(AllCharacters As Collection, ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single, ByVal ActiveOnly, ByVal ByForm As Boolean) As Collection
    Dim iCharacter As clsCharacter
    Dim cX As Single, cY As Single
    Set GetCharactersInSquare = New Collection
    For Each iCharacter In AllCharacters
        If iCharacter.IsActive Or Not ActiveOnly Then
            If ByForm Then
                With iCharacter.GetDragForm(True)
                    cX = .Left + .Width / 2
                    cY = .Top + .Height / 2
                End With
                If cX >= X And cY <= X + Width Then
                    If cY >= Y And cY <= Y + Height Then
                        GetCharactersInSquare.Add iCharacter
                    End If
                End If
            Else
                If iCharacter.Left >= X And iCharacter.Left <= X + Width Then
                    If iCharacter.Top >= Y And iCharacter.Top <= Y + Height Then
                        GetCharactersInSquare.Add iCharacter
                    End If
                End If
            End If
        End If
    Next
End Function

Public Function IsWithinArea(R As Form, Area As Form) As Boolean
    If R.Left + R.Width < Area.Left Or R.Left > Area.Left + Area.Width Then
        IsWithinArea = False
        Exit Function
    End If
    If R.Top + R.Height < Area.Top Or R.Top > Area.Top + Area.Height Then
        IsWithinArea = False
        Exit Function
    End If
    IsWithinArea = True
End Function



