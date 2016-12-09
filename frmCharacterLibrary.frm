VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCharacterLibrary 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   5415
   ClientLeft      =   2700
   ClientTop       =   2895
   ClientWidth     =   4860
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   4860
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtFilter 
      Height          =   288
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4332
   End
   Begin MSComctlLib.ListView lvCharacters 
      Height          =   2172
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   3732
      _ExtentX        =   6588
      _ExtentY        =   3836
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "character"
         Text            =   "Character"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "cr"
         Text            =   "CR"
         Object.Width           =   706
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "file"
         Text            =   "File"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "date"
         Text            =   "Date"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3000
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFileOpenFolder 
         Caption         =   "Open folder..."
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "View"
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "Refresh"
         Shortcut        =   {F5}
      End
   End
End
Attribute VB_Name = "frmCharacterLibrary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pFolder As String

Private pCharacterFiles As Collection

Public Event AddCharacter(Character As clsCharacter)

Private Sub Form_Load()
    pFolder = GetSetting("CombatMapper", "CharacterLists", "CharacterListFolder", App.Path)
    LoadFolder
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SaveSetting "CombatMapper", "CharacterLists", "CharacterListFolder", pFolder
End Sub

Private Sub Form_Resize()
    txtFilter.Width = Me.ScaleWidth
    lvCharacters.Move 0, txtFilter.Height, Me.ScaleWidth, Me.ScaleHeight - txtFilter.Height
End Sub

Private Sub lvCharacters_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Dim Key As String
    Dim Order As String
    With lvCharacters
        If .Tag <> "" Then
            Key = Mid(.Tag, 2)
            Order = Left(.Tag, 1)
        End If
        If Key = ColumnHeader.Key Then
            Select Case Order
            Case "a":
                .SortOrder = lvwDescending
                .Tag = "d" & ColumnHeader.Key
            Case "d":
                .SortOrder = lvwAscending
                .Tag = "a" & ColumnHeader.Key
            End Select
        Else
            .SortOrder = lvwAscending
            .Tag = "a" & ColumnHeader.Key
        End If
        .SortKey = ColumnHeader.Index - 1
    End With
End Sub

Private Sub lvCharacters_DblClick()
    If Not lvCharacters.selectedItem Is Nothing Then
        RaiseEvent AddCharacter(GetCharacter(lvCharacters.selectedItem.Key))
    End If
End Sub

Private Function GetCharacter(ByVal Key As String) As clsCharacter
    Dim f As String
    Dim c As String
    Dim cf As clsCharacterListFile
    c = PopHead(Key, "@")
    f = Key
    Set cf = GetCharacterFile(f)
    If Not cf Is Nothing Then
        Set GetCharacter = cf.GetCharacter(c)
    End If
End Function

'Hakee ensimm‰isen hahmon, jolla on m‰‰ritelty nimi
Public Function GetCharacterByName(ByVal Name As String) As clsCharacter
    Dim cf As clsCharacterListFile
    For Each cf In pCharacterFiles
        Set GetCharacterByName = cf.GetCharacter(Name)
        If Not GetCharacterByName Is Nothing Then
            Exit Function
        End If
    Next
End Function

Private Function GetCharacterFile(ByVal Filename As String) As clsCharacterListFile
    Dim iCharacterFile As clsCharacterListFile
    
    For Each iCharacterFile In pCharacterFiles
        If iCharacterFile.Filename = Filename Then
            Set GetCharacterFile = iCharacterFile
            Exit Function
        End If
    Next
End Function

Private Sub mnuFileOpenFolder_Click()
    On Error Resume Next
    With CommonDialog1
        .Filename = pFolder & "*.lst"
        .Filter = "Character lists|*.lst"
        .ShowOpen
        If Err = 0 Then
            pFolder = PathFromString(.Filename)
            LoadFolder
        End If
    End With
End Sub

Public Sub LoadFolder()
    Dim sFile As String
    Dim cFiles As Collection
    Dim iFile As Variant
    Dim nCharacterFile As clsCharacterListFile
    Set pCharacterFiles = New Collection
    Me.Caption = "Characters from " & pFolder
    sFile = Dir(pFolder & "*.lst", vbNormal)
    'Hahmon lis‰ys k‰ytt‰‰ picturefiless‰ FileExistsi‰ => ker‰t‰‰n
    'tiedostot ensin
    Set cFiles = New Collection
    Do While sFile <> ""
        If sFile <> "." And sFile <> ".." Then
            cFiles.Add pFolder & sFile
        End If
        sFile = Dir
    Loop
    For Each iFile In cFiles
        Set nCharacterFile = New clsCharacterListFile
        nCharacterFile.LoadCharacters iFile
        pCharacterFiles.Add nCharacterFile

    Next
    ListCharacters
End Sub

Public Sub ListCharacters()
    Dim iCharacterFile As clsCharacterListFile
    Dim iCharacter As clsCharacter
    Dim nItm As ListItem
    Dim R As RegExp
    Dim bMatch As Boolean
    If txtFilter.Text <> "" Then
        Set R = New RegExp
        R.Pattern = txtFilter.Text
        R.IgnoreCase = True
    End If

    lvCharacters.ListItems.Clear
    For Each iCharacterFile In pCharacterFiles
        For Each iCharacter In iCharacterFile.Characters
            If R Is Nothing Then
                bMatch = True
            Else
                bMatch = R.Test(iCharacter.Name)
            End If
            If bMatch Then
                Set nItm = lvCharacters.ListItems.Add(, iCharacter.Name & "@" & iCharacterFile.Filename, iCharacter.Name)
                nItm.SubItems(1) = iCharacter.CR
                nItm.SubItems(2) = iCharacterFile.Filename
                nItm.SubItems(3) = Format(FileDateTime(iCharacterFile.Filename), "yyyy-mm-dd hh:nn:ss")
            End If
        Next
    Next
End Sub

Private Sub mnuViewRefresh_Click()
    LoadFolder
End Sub

Private Sub txtFilter_Change()
    ListCharacters
End Sub

