Attribute VB_Name = "modDictionary"
Option Explicit

Public Sub SetKeyValue(Dictionary As Collection, Key As String, Value As String)
    Dim KeyValuePair As clsKeyValuePair
    For Each KeyValuePair In Dictionary
        If KeyValuePair.Key = Key Then
            KeyValuePair.Value = Value
            Exit Sub
        End If
    Next
    Set KeyValuePair = New clsKeyValuePair
    With KeyValuePair
        .Key = Key
        .Value = Value
    End With
    Dictionary.Add KeyValuePair
End Sub

Public Function GetKeyValue(Dictionary As Collection, Key As String, Optional DefValue As String) As String
    Dim KeyValuePair As clsKeyValuePair
    For Each KeyValuePair In Dictionary
        If KeyValuePair.Key = Key Then
            GetKeyValue = KeyValuePair.Value
            Exit Function
        End If
    Next
    GetKeyValue = DefValue
End Function

Public Function FindKeyValuePair(Dictionary As Collection, Key As String) As clsKeyValuePair
    Dim KeyValuePair As clsKeyValuePair
    For Each KeyValuePair In Dictionary
        If KeyValuePair.Key = Key Then
            Set FindKeyValuePair = KeyValuePair
            Exit Function
        End If
    Next
End Function

Public Function IndexOfKeyValuePair(Dictionary As Collection, Key As String) As Integer
    Dim i As Integer
    Dim KeyValuePair As clsKeyValuePair
    For Each KeyValuePair In Dictionary
        i = i + 1
        If KeyValuePair.Key = Key Then
            IndexOfKeyValuePair = i
            Exit Function
        End If
    Next
End Function
