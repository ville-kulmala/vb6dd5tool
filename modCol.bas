Attribute VB_Name = "modCol"
Option Explicit

Public Function AddToCol(Item As Variant, Optional Col As Collection) As Collection
    If Col Is Nothing Then
        Set Col = New Collection
    End If
    Col.Add Item
    Set AddToCol = Col
End Function

Public Function FindObjFromCol(Item As Variant, Col As Collection) As Integer
    Dim i As Integer
    Dim v As Variant
    For Each v In Col
        i = i + 1
        If Item Is v Then
            FindObjFromCol = i
            Exit For
        End If
    Next
End Function

Public Function CombineCol(LHC As Collection, RHC As Collection) As Collection
    Dim iVar As Variant
    Set CombineCol = New Collection
    For Each iVar In LHC
        CombineCol.Add iVar
    Next
    For Each iVar In RHC
        CombineCol.Add iVar
    Next
End Function

Public Function FindStringFromCol(Col As Collection, Value As String) As Integer
    Dim i As Integer
    For i = 1 To Col.Count
        If LCase(Col(i)) = LCase(Value) Then
            FindStringFromCol = i
            Exit Function
        End If
    Next
End Function

Public Function RemoveStrFromCol(Col As Collection, Value As String) As Boolean
    Dim i As Integer
    i = FindStringFromCol(Col, Value)
    If i > 0 Then
        Col.Remove i
        RemoveStrFromCol = True
    End If
End Function

Public Function ColAddFirst(Col As Collection, Value As String, AsUnique As Boolean)
    If AsUnique Then
        RemoveStrFromCol Col, Value
    End If
    If Col.Count > 0 Then
        Col.Add Value, Before:=1
    Else
        Col.Add Value
    End If
End Function

Public Function ColFromString(ByVal Expression As String, Optional ByVal Delimiter As String, Optional ByVal Limit As Integer) As Collection
    Dim a As Variant
    Dim i As Integer
    If Limit > 0 Then
        a = Split(Expression, Delimiter, Limit)
    Else
        a = Split(Expression, Delimiter)
    End If
    Set ColFromString = New Collection
    For i = 0 To UBound(a)
        ColFromString.Add a(i)
    Next
End Function

