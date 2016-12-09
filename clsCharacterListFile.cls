VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCharacterListFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Filename As String

Public Characters As Collection

Public Sub LoadCharacters(Optional ByVal File As String)
    If File = "" Then
        File = Filename
    Else
        Filename = File
    End If
    Set Characters = modCharacters.LoadCharacters(File)
End Sub

Public Function GetCharacter(ByVal Name As String) As clsCharacter
    Dim iCharacter As clsCharacter
    If Characters Is Nothing Then Exit Function
    For Each iCharacter In Characters
        If iCharacter.Name = Name Then
            Set GetCharacter = iCharacter
            Exit Function
        End If
    Next
End Function
