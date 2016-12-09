Attribute VB_Name = "modEncounter"
Option Explicit

Public Function CreateEncounterItem(ByVal Line As String, Optional AddToCol As Collection) As clsEncounterItem
    Dim a As Variant
    Dim i As Integer
    Set CreateEncounterItem = New clsEncounterItem
    With CreateEncounterItem
        Set .Items = New Collection
        .Line = Line
        .Emphasis = PopHead(Line, vbTab)
        If Not IsNumeric(.Emphasis) Then
            Debug.Print "EncounterItem.ReadLine exists: not numeric emphasis"
            Set CreateEncounterItem = Nothing
            Exit Function
        End If
        a = Split(Line, ";")
        For i = 0 To UBound(a)
            .Items.Add a(i)
        Next
    End With
    If Not AddToCol Is Nothing Then
        AddToCol.Add CreateEncounterItem
    End If
End Function

'TEST: Debug.Print ConcatCol(GetCreatureList(CreateEncounterItem("1" & vbtaB & "Orc:d6;Goblin:d6")))
Public Function GetCreatureList(EncounterItem As clsEncounterItem, Optional ToCol As Collection) As Collection
    Dim v As Variant
    Dim s As String
    Dim k As String
    Dim i As Integer
    Randomize Timer
    If ToCol Is Nothing Then
        Set ToCol = New Collection
    End If
    With EncounterItem
        For Each v In .Items
            If InStr(v, ":") Then
                s = v   'varianttista ei ehkä voi popata... stringiksi
                k = PopHead(s, ":")
                For i = 1 To RollDice(s)
                    ToCol.Add k
                Next
            Else
                ToCol.Add v
            End If
        Next
    End With
    Set GetCreatureList = ToCol
End Function

'Test: Debug.Print concatcol(ReadEncounterList("C:\Documents and Settings\Ville\My Documents\Dropbox\DD5\CombatMapper\InuList\CharacterLists\Wilderness-day.enc").RollEncounter)
Public Function ReadEncounterList(ByVal Filename As String, Optional AddToCol As Collection) As clsRandomEncounterList
    Set ReadEncounterList = New clsRandomEncounterList
    ReadEncounterList.ReadFile Filename
    If Not AddToCol Is Nothing Then
        AddToCol.Add ReadEncounterList
    End If
End Function
