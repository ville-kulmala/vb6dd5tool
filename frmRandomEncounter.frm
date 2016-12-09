VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmRandomEncounter 
   Caption         =   "Form1"
   ClientHeight    =   4395
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   4395
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPushEncounter 
      Caption         =   "Add creatures"
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   3840
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2160
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtEncounter 
      Height          =   1455
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   1800
      Width           =   4215
   End
   Begin MSComctlLib.ListView lvEncounters 
      Height          =   1815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   3201
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Title"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "%"
         Object.Width           =   794
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "File"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFileOpenFolder 
         Caption         =   "Open folder..."
      End
   End
End
Attribute VB_Name = "frmRandomEncounter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event EncounterListPushed(EncounterList As Collection)

Private pEncounterLists As Collection
Private pFolder As String

Private Sub cmdPushEncounter_Click()
    RaiseEvent EncounterListPushed(ColFromString(txtEncounter, vbCrLf))
End Sub

Private Sub Form_Load()
    pFolder = GetSetting("CombatMapper", "CharacterLists", "EncounterListFolder", "")
    If pFolder = "" Then
        pFolder = GetSetting("CombatMapper", "CharacterLists", "CharacterListFolder", App.Path)
    End If
    LoadFolder
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SaveSetting "CombatMapper", "CharacterLists", "EncounterListFolder", pFolder
End Sub


Public Sub LoadFolder()
    Dim sFile As String
    Dim cFiles As Collection
    Dim iFile As Variant
    Dim nEnc As clsRandomEncounterList
    
    Set pEncounterLists = New Collection
    Me.Caption = "Encounter lists from " & pFolder
    sFile = Dir(pFolder & "*.enc", vbNormal)
    'Ker‰t‰‰n tiedot ensin, vaikka t‰‰ll‰ ei sis‰ll‰ dirri‰ k‰ytet‰kk‰‰n
    Set cFiles = New Collection
    Do While sFile <> ""
        If sFile <> "." And sFile <> ".." Then
            cFiles.Add pFolder & sFile
        End If
        sFile = Dir
    Loop
    For Each iFile In cFiles
        ReadEncounterList iFile, pEncounterLists
    Next
    ListEncounters
End Sub

Private Sub ListEncounters()
    Dim iList As clsRandomEncounterList
    Dim nItem As ListItem
    With lvEncounters.ListItems
        .Clear
        For Each iList In pEncounterLists
            Set nItem = .Add(, iList.Filename, iList.Title)
            nItem.SubItems(1) = iList.Propability
            nItem.SubItems(2) = iList.Filename
        Next
    End With
End Sub

Private Function GetListByFile(Filename As String)
    Dim iList As clsRandomEncounterList
    For Each iList In pEncounterLists
        If iList.Filename = Filename Then
            Set GetListByFile = iList
        End If
    Next

End Function

Private Sub Form_Resize()
    Const PADDING = 90
    lvEncounters.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight / 2
    txtEncounter.Move 0, Me.ScaleHeight / 2, Me.ScaleWidth, Abs(Me.ScaleHeight / 2 - PADDING * 2 - cmdPushEncounter.Height)
    cmdPushEncounter.Move Me.ScaleWidth - PADDING - cmdPushEncounter.Width, Me.ScaleHeight - PADDING - cmdPushEncounter.Height
End Sub

Private Sub lvEncounters_DblClick()
    Dim iList As clsRandomEncounterList
    If lvEncounters.SelectedItem Is Nothing Then
        Exit Sub
    End If
    Set iList = GetListByFile(lvEncounters.SelectedItem.Key)
    If iList Is Nothing Then
        Exit Sub
    End If
    txtEncounter.Text = ConcatCol(iList.RollEncounter)
End Sub

Private Sub mnuFileOpenFolder_Click()
    On Error Resume Next
    With CommonDialog1
        .Filename = pFolder
        .Filter = "Encounter lists|*.enc"
        .ShowOpen
        If Err = 0 Then
            pFolder = PathFromString(.Filename)
            LoadFolder
        End If
    End With
End Sub
