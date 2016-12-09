Attribute VB_Name = "modCharacters"
Option Explicit

Public Const CHAR_INITIATIVE As String = "initiative"
Public Const CHAR_INITIATIVEBASE As String = "initiativebase"
Public Const CHAR_HITS As String = "hits"
Public Const CHAR_TEMPHITS As String = "temphits"
Public Const CHAR_MAXHITS As String = "maxhits"
Public Const CHAR_STATUS As String = "status"
Public Const CHAR_PICTUREFILE As String = "picturefile"
Public Const CHAR_LOCATION As String = "location"
Public Const CHAR_SIZE As String = "size"
Public Const CHAR_AC As String = "ac"
Public Const CHAR_SPEED As String = "speed"
Public Const CHAR_HD As String = "hd"
Public Const CHAR_CR As String = "cr"
Public Const CHAR_ABILITIES As String = "abilities"


Public Function GetFirstInitiative(Characters As Collection) As clsCharacter
    Dim i As clsCharacter
    Dim Init As Single
    Init = -100 ' Eiköhän tämä ole alempi kuin huonoimmalla dexillä...
    For Each i In Characters
        If i.Initiative > Init And i.IsActive Then
            Set GetFirstInitiative = i
            Init = i.Initiative
        End If
    Next
End Function

Public Function GetNextInitiative(Characters As Collection, Current As clsCharacter, Optional ByRef NewRound As Boolean) As clsCharacter
    Dim i As clsCharacter
    Dim nh As Single
    For Each i In Characters
        If i.Initiative < Current.Initiative And i.IsActive Then
            If nh < i.Initiative Then
                nh = i.Initiative
                Set GetNextInitiative = i
            End If
        End If
    Next
    If GetNextInitiative Is Nothing Then
        NewRound = True
        Set GetNextInitiative = GetFirstInitiative(Characters)
    End If
End Function

Public Function SortCharacters(Characters As Collection) As Collection
    Dim c As Collection
    Dim i As clsCharacter
    Dim R As clsCharacter
    Dim rC As Collection
    Dim lC As Collection
    
    Set SortCharacters = New Collection
    If Characters Is Nothing Then
        Exit Function
    ElseIf Characters.Count = 0 Then
        Exit Function
    End If
    Set rC = New Collection
    Set lC = New Collection
    For Each i In Characters
        If R Is Nothing Then
            Set R = i
        ElseIf R.Initiative > i.Initiative Then
            rC.Add i
        ElseIf R.Initiative < i.Initiative Then
            lC.Add i
        Else
            i.Initiative = i.Initiative - 0.0001
            rC.Add i
        End If
    Next
    rC.Add R
    If lC.Count = 0 Then
        Set SortCharacters = rC
    ElseIf rC.Count = 0 Then
        Set SortCharacters = lC
    Else
        Set SortCharacters = CombineCol(SortCharacters(lC), SortCharacters(rC))
    End If
    
End Function

Public Function GetPrevInitiative(Characters As Collection, Current As clsCharacter, Optional ByRef LastRound As Boolean) As clsCharacter
    Dim c As Collection
    Dim i As clsCharacter
    Set c = SortCharacters(Characters)
    For Each i In c
        If i Is Current Then
            If GetPrevInitiative Is Nothing Then
                Set GetPrevInitiative = GetLastInitiative(Characters)
                'LastRound = True
            Else
                Exit Function
            End If
        Else
            If i.IsActive And i.Hits > 0 Then
                Set GetPrevInitiative = i
            End If
        End If
    Next
End Function

Public Function GetLastInitiative(Characters As Collection) As clsCharacter
    Dim c As Collection
    Dim i As clsCharacter
    'Set c = SortCharacters(Characters)
    For Each i In Characters
        If i.IsActive Or i.Hits > 0 Then
            If GetLastInitiative Is Nothing Then
                Set GetLastInitiative = i
            ElseIf GetLastInitiative.Initiative > i.Initiative Then
                Set GetLastInitiative = i
            End If
        End If
    Next
End Function

Public Function FindCharacter(Name As String, Characters As Collection) As clsCharacter
    Dim iChar As clsCharacter
    For Each iChar In Characters
        If iChar.Name = Name Then
            Set FindCharacter = iChar
            Exit For
        End If
    Next
End Function

Public Function RemoveCharacter(Characters As Collection, Character As clsCharacter) As Boolean
    Dim i As Integer
    i = FindObjFromCol(Character, Characters)
    If i > 0 Then
        Characters.Remove i
        RemoveCharacter = True
    End If
End Function

Public Function GetNextFreeName(Name As String, Characters As Collection) As String
    Dim iChar As clsCharacter
    Dim i As Integer
    Dim sBase As String
    Dim sNum As String
    Set iChar = FindCharacter(Name, Characters)
    If iChar Is Nothing Then
        GetNextFreeName = Name
    Else
        sNum = GetTail(Name)
        If IsNumeric(sNum) Then
            sBase = Left(Name, Len(Name) - Len(sNum))
        Else
            sBase = Name & " "
        End If
            
        For i = 1 To 1000
            If FindCharacter(sBase & i, Characters) Is Nothing Then
                GetNextFreeName = sBase & i
                Exit For
            End If
        Next
    End If
End Function

Public Function CreateCharacter(Optional Name As String = "Name", Optional Initiative As Single = 0, Optional InitiativeBase As Single = 0, Optional Hits As Integer = 10, Optional MaxHits As Integer = 10, Optional AC As Integer = 10) As clsCharacter
    Set CreateCharacter = New clsCharacter
    With CreateCharacter
        .Name = Name
        .Initiative = Initiative
        .Hits = Hits
        .MaxHits = MaxHits
        .InitiativeBase = InitiativeBase
        .AC = AC
    End With
End Function

Public Function SaveCharacterToFile(Character As clsCharacter, Optional ByVal Filename As String) As Boolean
    Dim c As Collection
    Dim n As Collection
    Dim i As clsCharacter
    If Filename = "" Then Filename = Character.SourceFile
    If Filename = "" Then
        Debug.Print "SaveCharacterToFile: no filename"
        Exit Function
    End If
    Set c = LoadCharacters(Filename)
    Set n = New Collection
    n.Add Character
    For Each i In c
        If i.Name = Character.Name Then
            Debug.Print "SaveCharacterToFile: replacing character by name"
        Else
            n.Add i
        End If
    Next
    SaveCharacters Filename, n
    SaveCharacterToFile = True
    
End Function

'Lataa tiedostosta, lukee kollaasiin
Public Function LoadCharacters(ByVal Filename As String) As Collection
    Dim iCharacter As clsCharacter
    Set LoadCharacters = ReadCharacters(FileToCol(Filename))
    For Each iCharacter In LoadCharacters
        iCharacter.SourceFile = Filename
    Next
End Function

Public Function ReadCharacters(c As Collection) As Collection
    Dim R As Variant
    Dim cStrings As Collection
    Dim sHead As String, sTail As String
    Dim nCharacter As clsCharacter
    Set ReadCharacters = New Collection
    For Each R In c
        If Left(R, 1) = "[" Then
            If Not cStrings Is Nothing And Not nCharacter Is Nothing Then
                GetCharacterFromStrings cStrings, nCharacter
            End If
            Set cStrings = New Collection
            Set nCharacter = New clsCharacter
            nCharacter.Name = Replace(Mid(R, 2), "]", "")
            ReadCharacters.Add nCharacter
        Else
            If Trim(R) <> "" Then
                If cStrings Is Nothing Then
                    Debug.Print "LoadCharacters: now cStrings initiated. Ommiting: " & R
                Else
                    cStrings.Add R
                End If
            End If
        End If
    Next
    If Not cStrings Is Nothing And Not nCharacter Is Nothing Then
        GetCharacterFromStrings cStrings, nCharacter
    End If

End Function

Public Function GetCharacterFromStrings(ByVal Strings As Collection, Optional nCharacter As clsCharacter) As clsCharacter
    Dim R As Variant
    Dim sTail As String
    Dim sHead As String
    If nCharacter Is Nothing Then
        Set nCharacter = New clsCharacter
    End If
    For Each R In Strings
        If Left(R, 1) = "[" Then
            nCharacter.Name = Replace(Mid(R, 2), "]", "")
        Else
            On Error Resume Next
            sTail = Trim(R)
            sHead = Trim(PopHead(sTail, vbTab))
            Select Case LCase(sHead)
            Case CHAR_HITS:
                nCharacter.Hits = sTail
            Case CHAR_MAXHITS:
                nCharacter.MaxHits = sTail
            Case CHAR_TEMPHITS:
                nCharacter.TempHits = sTail
            Case CHAR_INITIATIVE
                nCharacter.Initiative = sTail
            Case CHAR_INITIATIVEBASE
                nCharacter.InitiativeBase = sTail
            Case CHAR_STATUS
                nCharacter.Status = sTail
            Case CHAR_PICTUREFILE
                nCharacter.PictureFile = sTail
            Case CHAR_LOCATION
                nCharacter.Location = sTail
                nCharacter.MoveToPosition
            Case CHAR_SIZE
                nCharacter.Size = sTail
            Case CHAR_AC
                nCharacter.AC = sTail
            Case CHAR_SPEED
                nCharacter.Speed = sTail
            Case CHAR_HD
                nCharacter.HD = sTail
            Case CHAR_CR
                nCharacter.CR = sTail
            Case CHAR_ABILITIES
                nCharacter.Abilities = sTail
            Case Else
                nCharacter.EntryValue(sHead) = DecodeText(sTail)
            End Select
            If Err <> 0 Then
                Debug.Print "CharFromString: error:" & Err.Number & " " & Err.Description, R
                Err.Clear
            End If
        End If
    Next
    Set GetCharacterFromStrings = nCharacter
End Function

Public Sub SaveCharacters(Filename As String, Characters As Collection)
    WriteColToFile Filename, GetCharacterStrings(Characters)
End Sub

Public Function GetCharacterStrings(Characters As Collection, Optional AppendTo As Collection) As Collection
    Dim iChar As clsCharacter
    Dim iEntry As clsKeyValuePair
    Dim c As Collection
    If AppendTo Is Nothing Then
        Set c = New Collection
    Else
        Set c = AppendTo
    End If
    For Each iChar In Characters
        With iChar
            c.Add "[" & .Name & "]"
            c.Add CHAR_INITIATIVEBASE & vbTab & .InitiativeBase
            c.Add CHAR_INITIATIVE & vbTab & .Initiative
            c.Add CHAR_HITS & vbTab & .Hits
            c.Add CHAR_TEMPHITS & vbTab & .TempHits
            c.Add CHAR_MAXHITS & vbTab & .MaxHits
            c.Add CHAR_STATUS & vbTab & .Status
            c.Add CHAR_PICTUREFILE & vbTab & .PictureFile
            c.Add CHAR_LOCATION & vbTab & .Location
            c.Add CHAR_SIZE & vbTab & .Size
            c.Add CHAR_AC & vbTab & .AC
            c.Add CHAR_SPEED & vbTab & .Speed
            c.Add CHAR_HD & vbTab & .HD
            c.Add CHAR_CR & vbTab & .CR
            c.Add CHAR_ABILITIES & vbTab & .Abilities
            For Each iEntry In .GetEntries
                c.Add iEntry.Key & vbTab & EncodeText(iEntry.Value)
            Next
        End With
    Next
    Set GetCharacterStrings = c
End Function

Public Function CloneCharacter(Character As clsCharacter) As clsCharacter
    Set CloneCharacter = GetCharacterFromStrings(GetCharacterStrings(AddToCol(Character)))
    CloneCharacter.SourceFile = Character.SourceFile
End Function

Public Sub ListCharacters(Characters As Collection, lvInitiative As ListView, Optional SelectedCharacter As clsCharacter, Optional CurrentInitiative As clsCharacter, Optional ShowInactive As Boolean = True)
    Dim Sel As ListItem
    Dim Cur As ListItem
    Dim nItem As ListItem
    Dim iChar As clsCharacter
    Dim SmallIcons As ImageList
    lvInitiative.ListItems.Clear
    If Not lvInitiative.SmallIcons Is Nothing Then
        Set SmallIcons = lvInitiative.SmallIcons
        Set lvInitiative.SmallIcons = Nothing
        On Error Resume Next
        
        SmallIcons.ListImages.Clear
        SmallIcons.ImageWidth = 64
        SmallIcons.ImageHeight = 64
        
        For Each iChar In Characters
            If FileExists(iChar.PictureFile) Then
                SmallIcons.ListImages.Add , iChar.PictureFile, iChar.GetPicture
            End If
        Next
        Set lvInitiative.SmallIcons = SmallIcons
        
    End If
    For Each iChar In Characters
        If iChar.IsActive Or ShowInactive Then
            With iChar
                Set nItem = lvInitiative.ListItems.Add(, , Format(iChar.Initiative, "00.00"))
                nItem.SubItems(1) = .Name
                nItem.SubItems(2) = .Hits & "/" & .MaxHits & " (" & .TempHits & ")"
                nItem.SubItems(3) = .Status
                If lvInitiative.ColumnHeaders.Count > 4 Then
                    nItem.SubItems(4) = .EntryValue("notes")
                End If
                If lvInitiative.ColumnHeaders.Count > 9 Then
                    nItem.SubItems(5) = .GetSave("Str")
                    nItem.SubItems(6) = .GetSave("Dex")
                    nItem.SubItems(7) = .GetSave("Con")
                    nItem.SubItems(8) = .GetSave("Int")
                    nItem.SubItems(9) = .GetSave("Wis")
                    nItem.SubItems(10) = .GetSave("Cha")
                End If
                If .Hits <= 0 Then
                    nItem.ForeColor = vbGrayText
                ElseIf .Hits < .MaxHits / 2 Then
                    nItem.ForeColor = vbRed
                Else
                    nItem.ForeColor = vbBlack
                End If
                If Not SmallIcons Is Nothing And .PictureFile <> "" Then
                    nItem.SmallIcon = .PictureFile
                End If
            End With
            If iChar Is SelectedCharacter Then
                Set Sel = nItem
            End If
            If iChar Is CurrentInitiative Then
                nItem.Bold = True
                Set Cur = nItem
            Else
                nItem.Bold = False
            End If
            If Not iChar.IsActive Then
                nItem.Ghosted = True
            End If
        End If
    Next
    If Not Cur Is Nothing Then
        Cur.Selected = True
        Cur.EnsureVisible
    End If
    If Not Sel Is Nothing Then
        Sel.Selected = True
        Sel.EnsureVisible
    End If
End Sub

Public Sub RollInitiatives(Characters As Collection)
    Dim Character As clsCharacter
    For Each Character In Characters
        Character.Initiative = Character.InitiativeBase + CSng(Format(Rnd * 20 + 1, "0.00"))
    Next
    Set Characters = SortCharacters(Characters)
End Sub

Public Function UpdateCharacterListView(Character As clsCharacter, lvCharacter As ListView, chkActive As CheckBox) As Boolean
    On Error Resume Next
    With Character
        .Name = lvCharacter.ListItems("Name")
        .Hits = lvCharacter.ListItems("Hits")
        .TempHits = lvCharacter.ListItems("TempHits")
        .MaxHits = lvCharacter.ListItems("MaxHits")
        .Status = lvCharacter.ListItems("Status")
        .Initiative = lvCharacter.ListItems("Initiative")
        .InitiativeBase = lvCharacter.ListItems("InitiativeBase")
        .Size = lvCharacter.ListItems("Size")
        .AC = lvCharacter.ListItems("AC")
        .Speed = lvCharacter.ListItems("Speed")
        .HD = lvCharacter.ListItems("HD")
        .CR = lvCharacter.ListItems("CR")
        .Abilities = lvCharacter.ListItems("Abilities")
        .IsActive = IIf(chkActive.Value = vbChecked, True, False)
    End With
    If Err <> 0 Then
        MsgBox "Error in character information: " & Err.Description
    Else
        UpdateCharacterListView = True
    End If
End Function

Public Sub EffectCharacter(Character As clsCharacter, Effects As String)
    'Effects: "Dmg:34;Status:Sleep;Active:true"
    Dim a As Variant
    Dim i As Integer
    Dim e As String
    If Effects = "" Then
        Exit Sub
    End If
    a = Split(Effects, ";")
    For i = 0 To UBound(a)
        e = a(i)
        Select Case LCase(PopHead(e, ":"))
        Case "dmg": Character.Hit e
        Case "status": Character.Status = e & " " & Character.Status
        Case "active": Character.IsActive = IIf(e = "false", False, True)
        End Select
    Next
    
End Sub
