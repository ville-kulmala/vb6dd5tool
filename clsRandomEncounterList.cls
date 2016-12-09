VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRandomEncounterList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Random Encounter lista: *.enc
' Voitaisiin tallentaa samaan kansion *.lst ja *.map kanssa
' ja ladata samaan tapaan kuin "CharacterLibrary"

'Malli:
' Title<tab>forest encounter, day
'   (tämä on listattu, jonka mukaan taulu luetaan)
' Propability<tab>15%
'   (kertoo, tuliko encounteria)
' <emphasis><tab><creature>:<number>;
'Priority: painoarvo. Lista tehdään painoarvoista niin, että painoarvot lopulta lasketaan yhteen ja heitetään d[painoarvo] ja mihin osuu...
' Esim.:
' 1   Orc:2d6;Troll:d3-1
' 2   Goblin:2d6;Warg:d3;Hobgoblin:d4
' 1   Owlbear:1
' 1   Troll:d3
' 1   %double%
' Emphasis yhteensä: 6 => heitto d5, 1 => örkkejä, 2-3 => goblineita, 4 => pöllökarhu, 5 => peikko, 6 => "heitä kahdesti" ja yhdistä
'Tarkoitus olisi sitten tehdä näistä uusi taistelunäkymä valituilla vaihtoehdoilla.
'Muuta: voitaisiin tehdä viite toiseen tauluun

Public Title As String
Public Propability As Single
Public Items As Collection

Public Filename As String

Public Function ReadFile(sFilename As String)
    Dim cRows As Collection
    Dim iRow As Variant
    Dim k As String
    Dim s As String
    Filename = sFilename
    Set cRows = FileToCol(sFilename)
    Set Items = New Collection
    For Each iRow In cRows
        s = iRow
        k = PopHead(s, vbTab)
        If LCase(k) = "title" Then
            Title = Trim(s)
        ElseIf LCase(k) = "propability" Then
            s = Trim(Replace(s, "%", ""))
            If IsNumeric(s) Then
                Propability = s
            Else
                Debug.Print "Encounter File: " & sFilename & ": propability not numeric"
            End If
        ElseIf IsNumeric(k) Then
            CreateEncounterItem iRow, Items
        ElseIf Trim(k) <> "" Then
            Debug.Print "Encounter list: '" & sFilename & "': unknown line '" & iRow
        End If
    Next
End Function

Public Function RollEncounter(Optional AddToCol As Collection, Optional ForceRoll As Boolean) As Collection
    Dim d As Integer
    Dim i As Integer
    Dim e As clsEncounterItem
    If AddToCol Is Nothing Then
        Set AddToCol = New Collection
    End If
    If Not ForceRoll And RollDie("d100") < Propability Then
        AddToCol.Add "No encounter during time"
    Else
        d = RollDie("d" & GetUpperLimit)
        For Each e In Items
            i = i + e.Emphasis
            If i >= d Then
                Set AddToCol = RollSpecials(e.GetCreatureList(AddToCol))
                Exit For
            End If
        Next
    End If
    Set RollEncounter = AddToCol
End Function

Private Function RollSpecials(CreatureList As Collection) As Collection
    Dim ToCol As Collection
    Dim i As Variant
    Dim bReroll As Boolean
    Set ToCol = New Collection
    For Each i In CreatureList
        Select Case LCase(i)
        Case "%double%"
            RollEncounter ToCol, True
            RollEncounter ToCol, True
            bReroll = True
        Case Else
            ToCol.Add i
        End Select
    Next
    
    If bReroll Then
        Set RollSpecials = RollSpecials(ToCol)
    Else
        Set RollSpecials = ToCol
    End If

End Function

Private Function GetUpperLimit() As Integer
    Dim i As clsEncounterItem
    For Each i In Items
        GetUpperLimit = GetUpperLimit + i.Emphasis
    Next
End Function
