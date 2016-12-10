VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmInitiative 
   Caption         =   "Initiatives"
   ClientHeight    =   7635
   ClientLeft      =   2310
   ClientTop       =   2070
   ClientWidth     =   7980
   LinkTopic       =   "Form1"
   ScaleHeight     =   7635
   ScaleWidth      =   7980
   Begin VB.CommandButton cmdRoll 
      Caption         =   "Roll"
      Height          =   372
      Left            =   6240
      TabIndex        =   21
      Top             =   3120
      Width           =   852
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   20
      Top             =   7380
      Width           =   7980
      _ExtentX        =   14076
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.CheckBox chkReactionUsed 
      Caption         =   "Reacted"
      Height          =   252
      Left            =   6960
      TabIndex        =   19
      Top             =   1320
      Width           =   975
   End
   Begin MSComDlg.CommonDialog PictureCommonDialog 
      Left            =   3360
      Top             =   3600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer tmrUpdate 
      Left            =   3120
      Top             =   3600
   End
   Begin VB.PictureBox picVarious 
      Height          =   2055
      Left            =   960
      ScaleHeight     =   1995
      ScaleWidth      =   4275
      TabIndex        =   16
      Top             =   4440
      Width           =   4335
      Visible         =   0   'False
      Begin VB.Timer tmrDeclaration 
         Left            =   0
         Top             =   0
      End
      Begin VB.PictureBox picFighter 
         AutoRedraw      =   -1  'True
         Height          =   972
         Left            =   0
         Picture         =   "frmInitiative.frx":0000
         ScaleHeight     =   915
         ScaleWidth      =   915
         TabIndex        =   17
         Top             =   0
         Width           =   972
         Visible         =   0   'False
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   0
         Top             =   0
         _ExtentX        =   979
         _ExtentY        =   979
         BackColor       =   -2147483643
         ImageWidth      =   40
         ImageHeight     =   40
         MaskColor       =   12632256
         _Version        =   393216
      End
      Begin WMPLibCtl.WindowsMediaPlayer MP 
         Height          =   495
         Left            =   480
         TabIndex        =   18
         Top             =   1080
         Width           =   735
         Visible         =   0   'False
         URL             =   ""
         rate            =   1
         balance         =   0
         currentPosition =   0
         defaultFrame    =   ""
         playCount       =   1
         autoStart       =   -1  'True
         currentMarker   =   0
         invokeURLs      =   -1  'True
         baseURL         =   ""
         volume          =   50
         mute            =   0   'False
         uiMode          =   "full"
         stretchToFit    =   0   'False
         windowlessVideo =   0   'False
         enabled         =   -1  'True
         enableContextMenu=   -1  'True
         fullScreen      =   0   'False
         SAMIStyle       =   ""
         SAMILang        =   ""
         SAMIFilename    =   ""
         captioningID    =   ""
         enableErrorDialogs=   0   'False
         _cx             =   1296
         _cy             =   873
      End
   End
   Begin VB.TextBox txtArticle 
      Height          =   1455
      Left            =   3120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   15
      Top             =   1560
      Width           =   4815
   End
   Begin VB.CommandButton cmdRemoveArticle 
      Caption         =   "Delete"
      Height          =   375
      Left            =   6960
      TabIndex        =   14
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton cmdAddArticle 
      Caption         =   "Add..."
      Height          =   375
      Left            =   6960
      TabIndex        =   13
      Top             =   360
      Width           =   855
   End
   Begin VB.ListBox lstArticles 
      Height          =   1020
      IntegralHeight  =   0   'False
      Left            =   4320
      TabIndex        =   12
      Top             =   480
      Width           =   2652
   End
   Begin VB.CommandButton cmdSetScaling 
      Caption         =   "Set"
      Height          =   255
      Left            =   3720
      TabIndex        =   10
      Top             =   3120
      Width           =   615
   End
   Begin VB.TextBox txtScaling 
      Height          =   285
      Left            =   3000
      TabIndex        =   8
      Text            =   "96"
      Top             =   3120
      Width           =   735
   End
   Begin VB.PictureBox picCharacter 
      AutoRedraw      =   -1  'True
      Height          =   972
      Left            =   3120
      ScaleHeight     =   915
      ScaleWidth      =   1035
      TabIndex        =   7
      Top             =   480
      Width           =   1092
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2880
      Top             =   3600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtDeclarationTime 
      Height          =   285
      Left            =   1200
      TabIndex        =   6
      Text            =   "10"
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton cmdDeclarationTimer 
      Caption         =   "Start! (F4)"
      Height          =   372
      Left            =   0
      TabIndex        =   5
      Top             =   3120
      Width           =   1212
   End
   Begin VB.TextBox txtHit 
      Height          =   285
      Left            =   4440
      TabIndex        =   4
      Text            =   "d10"
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton cmdHit 
      Caption         =   "Hit"
      Height          =   375
      Left            =   5280
      TabIndex        =   3
      Top             =   3120
      Width           =   972
   End
   Begin MSComctlLib.ListView lvCharacter 
      Height          =   2655
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   4683
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Value"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Stat"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7980
      _ExtentX        =   14076
      _ExtentY        =   741
      ButtonWidth     =   1640
      ButtonHeight    =   582
      Appearance      =   1
      TextAlignment   =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Clone"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Delete"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Add"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Update"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Prev"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Next"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Stop"
            Key             =   "stop"
            Object.ToolTipText     =   "Stop movement at the end of Speed"
            Style           =   1
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Center"
            Key             =   "center"
            Object.ToolTipText     =   "Center map on character"
            Style           =   1
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvInitiative 
      Height          =   3852
      Left            =   0
      TabIndex        =   0
      Top             =   3480
      Width           =   6852
      _ExtentX        =   12091
      _ExtentY        =   6800
      View            =   3
      LabelEdit       =   1
      SortOrder       =   -1  'True
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   11
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Initiative"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Hits"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Status"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Notes"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Str"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Dex"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Con"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Int"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Wis"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Cha"
         Object.Width           =   882
      EndProperty
   End
   Begin VB.CheckBox chkActive 
      Caption         =   "Active"
      Height          =   255
      Left            =   6960
      TabIndex        =   11
      Top             =   1080
      Value           =   1  'Checked
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Scaling:"
      Height          =   255
      Left            =   2280
      TabIndex        =   9
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "New..."
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "Open..."
      End
      Begin VB.Menu mnuFileOpenCharacterLists 
         Caption         =   "Recent Character Lists"
         Begin VB.Menu mnuFileOpenCharacterList 
            Caption         =   "(No history)"
            Index           =   0
         End
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "Save"
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save as..."
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuEditRollInitiatives 
         Caption         =   "Roll Initiatives"
      End
      Begin VB.Menu mnuEditValue 
         Caption         =   "Edit value"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuEditStartCounter 
         Caption         =   "Start!"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuEditPickPicture 
         Caption         =   "Pick picture..."
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "View"
      Begin VB.Menu mnuViewPublicList 
         Caption         =   "Public list"
      End
      Begin VB.Menu mnuViewDragFormTester 
         Caption         =   "Drag form tester"
      End
      Begin VB.Menu mnuViewShowMap 
         Caption         =   "Show Map"
      End
      Begin VB.Menu mnuViewImmediate 
         Caption         =   "Immediate Window"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuViewDiceRoller 
         Caption         =   "Dice roller"
      End
      Begin VB.Menu mnuViewCurrentCharacter 
         Caption         =   "Current character"
      End
      Begin VB.Menu mnuViewCharacterLibrary 
         Caption         =   "Character library..."
      End
      Begin VB.Menu mnuViewRandomEncounter 
         Caption         =   "Random Encounters..."
      End
      Begin VB.Menu mnuViewSep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewXpValue 
         Caption         =   "Calculate XP value"
      End
   End
   Begin VB.Menu pmnuInuList 
      Caption         =   "pmnuInuList"
      Visible         =   0   'False
      Begin VB.Menu pmnuInuListSetInitiative 
         Caption         =   "Set initiative"
      End
      Begin VB.Menu pmnuInuListRollInitiative 
         Caption         =   "Roll initiative"
      End
      Begin VB.Menu pmnuInuListRollHP 
         Caption         =   "Roll hitpoints"
      End
      Begin VB.Menu pmnuInuListSaveCharacter 
         Caption         =   "Save character"
      End
      Begin VB.Menu pmnuInuListAddToList 
         Caption         =   "Add to list"
      End
      Begin VB.Menu pmnuInuListTarget 
         Caption         =   "Target"
         Begin VB.Menu pmnuInuListAttackTarget 
            Caption         =   "pmnuInuListAttackTarget"
            Index           =   0
         End
      End
      Begin VB.Menu pmnuInuListAttack 
         Caption         =   "Attack"
         Begin VB.Menu pmnuInuListAttackWith 
            Caption         =   "pmnuInuListAttackWith"
            Index           =   0
         End
      End
      Begin VB.Menu mnuInuListAttackMode 
         Caption         =   "Attack Mode"
         Begin VB.Menu mnuInuListAttackModeItem 
            Caption         =   "Disadvantage"
            Index           =   0
         End
         Begin VB.Menu mnuInuListAttackModeItem 
            Caption         =   "Normal"
            Checked         =   -1  'True
            Index           =   1
         End
         Begin VB.Menu mnuInuListAttackModeItem 
            Caption         =   "Advantage"
            Index           =   2
         End
      End
      Begin VB.Menu mnuInuListCurInit 
         Caption         =   "Against Current Initiative"
         Begin VB.Menu mnuInuListCurInitAMMode 
            Caption         =   "Disadvantage"
            Index           =   0
         End
         Begin VB.Menu mnuInuListCurInitAMMode 
            Caption         =   "Normal"
            Index           =   1
         End
         Begin VB.Menu mnuInuListCurInitAMMode 
            Caption         =   "Advantage"
            Index           =   2
         End
      End
   End
End
Attribute VB_Name = "frmInitiative"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Ongelmia Windows 10:ss‰: MSCOMCTL.OCX ei pelit‰.
'Vaihda PROJEKTITIEDOSTOON t‰m‰:
'Object={831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0; MSCOMCTL.OCX
Public Characters As Collection
Private pSelectedCharacter As clsCharacter
Private WithEvents pCurInitiative As clsCharacter
Attribute pCurInitiative.VB_VarHelpID = -1

Private WithEvents pPublicList As frmPublicList
Attribute pPublicList.VB_VarHelpID = -1

Private WithEvents MyMap As frmMapBackground
Attribute MyMap.VB_VarHelpID = -1

Private pImmediate As frmImmediate

Private pRound As Integer

Private pListFile As String

Private pCurCharacterForm As frmCharacter

Private RecentCharacterFiles As New Collection

Private WithEvents pCharacterLibrary As frmCharacterLibrary
Attribute pCharacterLibrary.VB_VarHelpID = -1

Private WithEvents pRandomEncounters As frmRandomEncounter
Attribute pRandomEncounters.VB_VarHelpID = -1

Private pCharacterList As dlgCharacterList

'Current characters form
Public Function GetCurCharacterForm(Optional CreateNeeded As Boolean = True) As frmCharacter
    If pCurCharacterForm Is Nothing And CreateNeeded Then
        Set pCurCharacterForm = New frmCharacter
    End If
    If Not pCurCharacterForm Is Nothing Then
        With pCurCharacterForm
            Set .Character = pCurInitiative
            .Show , Me
            .Caption = "Current initiative"
            Set .InitiativeForm = Me
        End With

    End If
    Set GetCurCharacterForm = pCurCharacterForm
End Function

Public Sub LogEvent(ByVal sEvent As String)
    If pImmediate Is Nothing Then Exit Sub
    pImmediate.LogEvent sEvent
End Sub

Private Sub chkActive_Click()
    If Not pSelectedCharacter Is Nothing Then
        pSelectedCharacter.IsActive = IIf(chkActive.Value = vbChecked, True, False)
        ListCharacters
    End If
End Sub

Private Sub chkReactionUsed_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UpdateCharacter
End Sub

Private Sub cmdAddArticle_Click()
    Dim s As String
    If pSelectedCharacter Is Nothing Then
        Exit Sub
    End If
    s = InputBox("Anna artikkelin nimi", "Lis‰‰ artikkeli")
    If s = "" Then
        Exit Sub
    End If
    pSelectedCharacter.EntryValue(s) = ShowTextEditor(pSelectedCharacter.EntryValue(s), Me, s)
    ShowCharacter
End Sub

Private Sub cmdDeclarationTimer_Click()
    'Starts timer clock.
    With tmrDeclaration
        If .Interval = 0 Then
            .Interval = 1000
            .Enabled = True
            tmrDeclaration_Timer
        Else
            .Enabled = False
            .Interval = 0
        End If
    End With
End Sub

Private Sub cmdHit_Click()
    Dim d As Integer
    d = RollDice(txtHit.Text)
    cmdHit.Caption = "Hit " & d
    HitCharacter d, pCurInitiative, pSelectedCharacter
End Sub

Public Function HitCharacter(ByVal Damage As String, Attacker As clsCharacter, Target As clsCharacter, Optional Desc As String) As String
    Dim pText As String

    If Attacker Is Nothing Then
        pText = "[Unknown]"
    Else
        pText = Attacker.Name
    End If
    If Target Is Nothing Then
        MsgBox "No selected character"
        Exit Function
    End If
    With Target
        If Desc <> "" Then
            pText = Desc
            .Hit Damage 'T‰ytyy se damagekin tehd‰...
        Else
            pText = pText + " hits: " + .Hit(Damage)
        End If
        LogEvent pText
    
        'UpdateCharacter 'Kirjoittaisko vauriot uudestaan pois...
        ListCharacters
        ShowCharacter
        'UpdateCharacterTimed
    End With
    HitCharacter = pText
End Function

Private Sub cmdRemoveArticle_Click()
    If pSelectedCharacter Is Nothing Then
        Exit Sub
    End If
    If lstArticles.ListIndex = -1 Then
        Exit Sub
    End If
    pSelectedCharacter.RemoveEntryValue (lstArticles.List(lstArticles.ListIndex))
End Sub

Private Sub cmdRoll_Click()
    cmdRoll.Caption = "Roll: " & RollDice(txtHit.Text)
End Sub

Private Sub cmdSetScaling_Click()
    If MyMap Is Nothing Or Not IsNumeric(txtScaling.Text) Then
        Exit Sub
    End If
    MyMap.Scaling = txtScaling.Text
End Sub

Private Sub Form_Load()
    Set Characters = New Collection
    ShowCharacter
    mnuViewImmediate_Click
    LoadRecentFileList
    ListRecentFileLists
End Sub

Private Sub ListRecentFileLists()
    Dim i As Integer
    mnuFileOpenCharacterList(0).Visible = True
    'Ei voi...
    'For i = 1 To mnuFileOpenCharacterList.UBound
    '    Unload mnuFileOpenCharacterList
    'Next
    If RecentCharacterFiles.Count > 0 Then
        For i = 1 To RecentCharacterFiles.Count
            If mnuFileOpenCharacterList.UBound < i Then
                Load mnuFileOpenCharacterList(i)
            End If
            With mnuFileOpenCharacterList(i)
                .Caption = "&" & i & ": " & RecentCharacterFiles(i)
                .Visible = True
            End With
        Next
        mnuFileOpenCharacterList(0).Visible = False
    End If
End Sub

Private Sub LoadRecentFileList()
    Dim i As Integer
    Dim sFile As String
    Set RecentCharacterFiles = New Collection
    For i = 1 To 10
        sFile = GetSetting("CombatMapper", "CharacterLists", "ListFile_" & i, "")
        If sFile <> "" Then
            RecentCharacterFiles.Add sFile
        Else
            Exit For
        End If
    Next
End Sub

Private Sub SaveRecentFileLists()
    Dim v As Variant
    Dim i As Integer
    i = 1
    For Each v In RecentCharacterFiles
        SaveSetting "CombatMapper", "CharacterLists", "ListFile_" & i, v
        i = i + 1
    Next
    SaveSetting "CombatMapper", "CharacterLists", "ListFile_" & i, ""
End Sub

Private Sub PlayGong()
    MP.URL = App.Path & "\notify.wav"
    MP.Controls.play
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SaveRecentFileLists
End Sub

Private Sub Form_Resize()
    With lvInitiative
        .Height = Abs(Me.ScaleHeight - .Top)
        .Width = Me.ScaleWidth
    End With
    With txtArticle
        .Width = Abs(Me.ScaleWidth - .Left)
    End With
    With lstArticles
        .Width = Abs(Me.ScaleWidth - cmdAddArticle.Width - .Left)
        cmdAddArticle.Left = .Left + .Width
        cmdRemoveArticle.Left = .Left + .Width
        chkActive.Left = .Left + .Width
        chkReactionUsed.Left = .Left + .Width
    End With
End Sub

Private Sub lstArticles_Click()
    If lstArticles.ListIndex = -1 Or pSelectedCharacter Is Nothing Then
        txtArticle.Text = ""
    Else
        txtArticle.Text = pSelectedCharacter.EntryValue(lstArticles.List(lstArticles.ListIndex))
    End If
    
End Sub

Private Sub lstArticles_DblClick()
    Dim i As Integer
    If lstArticles.ListIndex > -1 Then
        If pSelectedCharacter Is Nothing Then
            lstArticles.Clear
        Else
            i = lstArticles.ListIndex
            With pSelectedCharacter
                .EntryValue(lstArticles.List(i)) = ShowTextEditor(.EntryValue(lstArticles.List(i)), Me)
            End With
        End If
    End If
End Sub

Private Sub lvCharacter_AfterLabelEdit(Cancel As Integer, NewString As String)
    UpdateCharacterTimed
End Sub

Private Sub lvInitiative_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set SelectedCharacter = FindCharacter(Item.SubItems(1), Characters)
End Sub

Private Sub lvInitiative_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        PopulateInuListMenu
        Me.PopupMenu pmnuInuList
    End If
End Sub

Private Sub PopulateInuListMenu()
    Dim SelChar As clsCharacter
    Dim cChars As Collection
    Dim iChar As clsCharacter
    Dim cAttacks As Collection
    Dim i As Integer
    Dim v As Variant
    On Error Resume Next
    If lvInitiative.selectedItem Is Nothing Then
        pmnuInuListTarget.Enabled = False
        pmnuInuListAttack.Enabled = False
        mnuInuListCurInit.Enabled = False
    Else
        pmnuInuListTarget.Enabled = True
        pmnuInuListAttack.Enabled = True
        mnuInuListCurInit.Enabled = True
        pmnuInuListAttackTarget(0).Visible = True
        pmnuInuListAttackWith(0).Visible = True
        For i = 1 To pmnuInuListAttackTarget.UBound
            Unload pmnuInuListAttackTarget(i)
        Next
        For i = 1 To pmnuInuListAttackWith.UBound
            Unload pmnuInuListAttackWith(i)
        Next
        i = 1
        For Each iChar In Characters
            Load pmnuInuListAttackTarget(i)
            With pmnuInuListAttackTarget(i)
                .Caption = iChar.Name
                If SelectedCharacter.Opponent Is iChar Then
                    .Checked = True
                End If
                If SelectedCharacter Is iChar Then
                    .Caption = .Caption & " (self)"
                End If
                If Not iChar.IsActive Then
                    .Caption = .Caption & " (inactive)"
                    .Enabled = False
                End If
                Select Case SelectedCharacter.GetAdvantageMode(iChar.Name)
                Case enAdvantageMode.AMAdvantage:
                    .Caption = .Caption & " (advantage)"
                Case enAdvantageMode.AMDisadvantage:
                    .Caption = .Caption & " (disadvantage)"
                End Select
                .Tag = iChar.Name
            End With
            i = i + 1
        Next
        Set cAttacks = SelectedCharacter.GetAttackNames
        i = 1
        For Each v In cAttacks
            Load pmnuInuListAttackWith(i)
            With pmnuInuListAttackWith(i)
                .Caption = v
            End With
            i = i + 1
        Next
        pmnuInuListAttackTarget(0).Visible = False
        pmnuInuListAttackWith(0).Visible = False
        Dim AM As Long
        If Not pCurInitiative Is Nothing Then
            mnuInuListCurInit.Caption = "Has against " & pCurInitiative.Name
            AM = SelectedCharacter.GetAdvantageMode(pCurInitiative.Name)
            For i = 0 To 2
                If i = AM + 1 Then
                    mnuInuListCurInitAMMode(i).Checked = True
                Else
                    mnuInuListCurInitAMMode(i).Checked = False
                End If
            Next
        End If
        
    End If
    If Err <> 0 Then
        Debug.Print "PopulateInuListMenu:", Err.Description
        Err.Clear
    End If
End Sub

Private Sub mnuEditPickPicture_Click()
    PicPicture
End Sub

Public Sub PicPicture()
    If pSelectedCharacter Is Nothing Then
        Exit Sub
    End If
    On Error Resume Next
    With PictureCommonDialog
        .Filter = "Picture files|*.jpg;*.bmp;*.wmg;*.gif;*.ico"
        .Filename = pSelectedCharacter.PictureFile
        .CancelError = True
        .ShowOpen
        If Err = 0 Then
            pSelectedCharacter.PictureFile = .Filename
            ListCharacters
            ShowCharacter
        End If
    End With
End Sub

Private Sub mnuEditRollInitiatives_Click()
    pRound = 1
    RollInitiatives Characters
    ListCharacters
End Sub

Private Sub mnuEditStartCounter_Click()
    cmdDeclarationTimer_Click
End Sub

Private Sub mnuEditValue_Click()
    If Not lvCharacter.selectedItem Is Nothing Then
        lvCharacter.StartLabelEdit
    End If
End Sub

Private Sub mnuFileNew_Click()
    Dim d As New frmInitiative
    d.Show
End Sub

Private Sub mnuFileOpen_Click()

    On Error Resume Next
    With CommonDialog1
        .CancelError = True
        .Filter = "Map|*.map|Character lists|*.lst;*.txt;*.tab"
        .Filename = pListFile
        .ShowOpen
        If Err <> 0 Then
            Err.Clear
            Exit Sub
        End If
        Select Case .FilterIndex
        Case 1
            LoadMap .Filename
        Case 2
            pListFile = .Filename
            LoadCharactersFromFile (.Filename)
        End Select
    End With
    
    Debug.Print Err.Description
    Err.Clear
End Sub

Private Sub LoadMap(Filename As String)
    If MyMap Is Nothing Then
        mnuViewShowMap_Click
    End If
    MyMap.LoadMap Filename
End Sub


Private Sub LoadCharactersFromFile(Filename As String)
    Dim c As Collection
    Dim iChar As clsCharacter
    Set c = LoadCharacters(Filename)
    For Each iChar In c
        AddCharacter iChar
    Next
    ListCharacters
    pListFile = Filename
    ColAddFirst RecentCharacterFiles, Filename, True
    ListRecentFileLists
End Sub

Private Sub mnuFileOpenCharacterList_Click(Index As Integer)
    Dim sFilename As String
    If Index > 0 Then
        sFilename = mnuFileOpenCharacterList(Index).Caption
        sFilename = Mid(sFilename, InStr(sFilename, " ") + 1)
        LoadCharactersFromFile sFilename
    End If
End Sub

Private Sub mnuFileSave_Click()
    If pListFile = "" Then
        SaveAs
    Else
        SaveCharacters pListFile, Characters
    End If
End Sub

Public Sub SaveAs(Optional FilterIndex As Integer = 1, Optional Filename As String)
    'FilterIndex:
    '1: koko lista
    '2: vain valittu (curInitiative)
    If Filename = "" Then
        Filename = pListFile
    End If
    On Error Resume Next
    With CommonDialog1
        .CancelError = True
        .Filter = "Character lists|*.lst;*.txt;*.tab|Current character|*.lst"
        .FilterIndex = FilterIndex
        .Filename = Filename
        .ShowSave
        If Err <> 0 Then
            Err.Clear
            Exit Sub
        End If
        Select Case .FilterIndex
        Case 1: SaveCharacters .Filename, Characters
        Case 2:
            'Tallennetaan selected character. Sen tallennus voidaan tehd‰ listasta oikealla clikill‰.
            If SelectedCharacter Is Nothing Then
                MsgBox "No current character"
            Else
                SaveCharacters .Filename, AddToCol(SelectedCharacter)
            End If
        End Select
        pListFile = .Filename
    End With
    
    Debug.Print Err.Description
    Err.Clear

End Sub


Private Sub mnuFileSaveAs_Click()
    SaveAs
End Sub

Private Sub mnuInuListCurInitAdvantage_Click()
    If Not pCurInitiative Is Nothing Then
        SelectedCharacter.SetAdvantageMode pCurInitiative.Name, AMAdvantage
    End If
End Sub

Private Sub mnuInuListCurInitAMMode_Click(Index As Integer)
    Dim AM As Long
    If pCurInitiative Is Nothing Then
        Debug.Print "No initiative!"
    Else
        SelectedCharacter.SetAdvantageMode pCurInitiative.Name, Index - 1
    End If
    
End Sub

Private Sub mnuView_Click()
    If pPublicList Is Nothing Then
        mnuViewPublicList.Checked = False
    Else
        mnuViewPublicList.Checked = True
    End If
End Sub

Private Sub mnuViewCharacterLibrary_Click()
    If pCharacterLibrary Is Nothing Then
        Set pCharacterLibrary = New frmCharacterLibrary
    End If
    pCharacterLibrary.Show , Me
End Sub

Private Sub mnuViewCurrentCharacter_Click()
    GetCurCharacterForm True
End Sub

Private Sub mnuViewDiceRoller_Click()
    Dim f As frmDiceRoller
    Set f = New frmDiceRoller
    f.Show , Me
End Sub

Private Sub mnuViewDragFormTester_Click()
    Dim f As frmDragFormTester
    Set f = New frmDragFormTester
    f.Show , Me
End Sub

Private Sub mnuViewImmediate_Click()
    If pImmediate Is Nothing Then
        Set pImmediate = New frmImmediate
    End If
    pImmediate.Show , Me
End Sub

Private Sub mnuViewPublicList_Click()
    Dim pCol As ColumnHeader
    If pPublicList Is Nothing Then
        Set pPublicList = New frmPublicList
        pPublicList.Show , Me
        Set pPublicList.ListView.SmallIcons = ImageList1
    Else
        Unload pPublicList
        ListCharacters
    End If
    Debug.Print
    
End Sub

Private Sub mnuViewRandomEncounter_Click()
    If pRandomEncounters Is Nothing Then
        Set pRandomEncounters = New frmRandomEncounter
    End If
    pRandomEncounters.Show , Me
    
End Sub

Private Sub mnuViewShowMap_Click()
    If MyMap Is Nothing Then
        Set MyMap = New frmMapBackground
    End If
    MyMap.Show , Me
    MyMap.SetCharacters Characters
    Set MyMap.Initiative = Me
    If IsNumeric(txtScaling.Text) Then
        MyMap.Scaling = txtScaling.Text
    End If
End Sub


Private Sub mnuViewXpValue_Click()
    Dim i As clsCharacter
    Dim XP As Long
    For Each i In Characters
        XP = XP + GetXPValue(i.CR)
    Next
    MsgBox "M‰‰ritettyjen CR tietojen perusteella XP: " & XP
End Sub

Private Sub MyMap_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not pSelectedCharacter Is Nothing Then
        If Not pSelectedCharacter.GetDragForm(False) Is Nothing Then
            With pSelectedCharacter.GetDragForm(False)
                Me.Caption = "Range:" & Format(Sqr(Abs((X - .Left - .Width / 2 + MyMap.Left + MyMap.picHolder.Left) / MyMap.Scaling) ^ 2 + _
                    Abs((Y - .Top - .Height / 2 + MyMap.Top + MyMap.picHolder.Top) / MyMap.Scaling) ^ 2), "0.0")
            End With
        End If
    End If
End Sub

Private Sub MyMap_Move(X As Single, Y As Single, dX As Single, dY As Single)
    Dim iChar As clsCharacter
    Debug.Print X
    For Each iChar In Characters
        If Not iChar.GetDragForm(False) Is Nothing Then
            With iChar.GetDragForm(False)
                .Left = .Left - dX
                .Top = .Top - dY
                .Visible = IsWithinArea(iChar.GetDragForm(False), MyMap)
            End With
        End If
    Next
End Sub

Private Sub pCharacterLibrary_AddCharacter(Character As clsCharacter)
    If Character Is Nothing Then Exit Sub
    CloneCharacter Character
End Sub

Private Sub pCurInitiative_Move(X As Single, Y As Single, PathLen As Single, Pause As Boolean)
    Me.Caption = "Position:" & Format(X / GetScaling, "0.0") & ":" & Format(Y / GetScaling, "0.0") & " Move: " & PathLen / GetScaling
    If Not MyMap Is Nothing Then
        MyMap.Caption = Me.Caption
    End If
    If Toolbar1.Buttons("stop").Value = tbrPressed Then
        If PathLen / GetScaling > pCurInitiative.GetSpeed Then
            Pause = True
        End If
    End If
    
End Sub

Private Sub pCurInitiative_Resize(X As Single, Y As Single)
    Me.Caption = "Size:" & Format(X / GetScaling, "0.0") & "x" & Format(Y / GetScaling, "0.0")
End Sub

Private Function GetScaling() As Single
    If IsNumeric(txtScaling.Text) Then
        If CSng(txtScaling.Text) > 0 Then
            GetScaling = txtScaling.Text
        End If
    End If
    GetScaling = 96
End Function

Private Sub picCharacter_DblClick()
    PicPicture
End Sub

Private Sub pmnuInuListAddToList_Click()
    If Not SelectedCharacter Is Nothing Then
        AddToCharacterList SelectedCharacter
    End If
End Sub

Private Function AddToCharacterList(Character As clsCharacter)
    If pCharacterList Is Nothing Then
        Set pCharacterList = New dlgCharacterList
    End If
    pCharacterList.AddCharacter Character
    pCharacterList.Show , Me
    
End Function

Private Sub pmnuInuListAttackTarget_Click(Index As Integer)
    If Index = 0 Then
        Exit Sub
    End If
    Set SelectedCharacter.Opponent = Characters(Index)
End Sub

Private Sub pmnuInuListAttackWith_Click(Index As Integer)
    If Index = 0 Then
        Exit Sub
    End If
    If SelectedCharacter.Opponent Is Nothing Then
        MsgBox "No target selected. Select opponent first."
        Exit Sub
    End If
    With pmnuInuListAttackWith(Index)
        
        SelectedCharacter.Attack .Caption, SelectedCharacter.Opponent, Me, SelectedCharacter.GetAdvantageMode(SelectedCharacter.Opponent.Name)
        'TODO: menu itemit attack modelle
        'Pit‰isi s‰ilˆ‰ hahmon tiedoissa.
    End With
    
End Sub

Private Sub pmnuInuListRollHP_Click()
    If Not SelectedCharacter Is Nothing Then
        SelectedCharacter.RollHitPoints
        ShowCharacter
        ListCharacters
    End If
End Sub

Private Sub pmnuInuListRollInitiative_Click()
    If Not SelectedCharacter Is Nothing Then
        RollInitiatives AddToCol(SelectedCharacter)
        ShowCharacter
        ListCharacters
    End If
End Sub

Private Sub pmnuInuListSaveCharacter_Click()
    If Not SelectedCharacter Is Nothing Then
        If SelectedCharacter.SourceFile <> "" Then
            SaveCharacterToFile SelectedCharacter
        Else
            SaveAs 2
        End If
    End If
End Sub

Private Sub pmnuInuListSetInitiative_Click()
    Dim P As clsCharacter
    If lvInitiative.selectedItem Is Nothing Then
        Exit Sub
    End If
    Set CurInitiative = FindCharacter(lvInitiative.selectedItem.SubItems(1), Characters)
    
    Set SelectedCharacter = pCurInitiative
    ListCharacters
End Sub

Private Sub pPublicList_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set pPublicList = Nothing
End Sub

Private Sub pRandomEncounters_EncounterListPushed(EncounterList As Collection)
    Dim i As Variant
    Dim nChar As clsCharacter
    mnuViewCharacterLibrary_Click   'Ladataan varmuuden vuoksi
    For Each i In EncounterList
        If Trim(i) <> "" Then
            Set nChar = pCharacterLibrary.GetCharacterByName(i)
            If Not nChar Is Nothing Then
                CloneCharacter nChar
                nChar.RollHitPoints
                RollInitiatives AddToCol(nChar)
            End If
        End If
    Next
    ListCharacters
End Sub

Private Sub tmrDeclaration_Timer()
    With cmdDeclarationTimer
        If IsNumeric(.Caption) Then
            If .Caption > 0 Then
                .Caption = .Caption - 1
            Else
                .Caption = "Time up!"
                PlayGong
                tmrDeclaration.Interval = 0
                tmrDeclaration.Enabled = False
            End If
        Else
            .Caption = 10
        End If
        If Not pPublicList Is Nothing Then
            If pCurInitiative Is Nothing Then
                pPublicList.Caption = .Caption
            Else
                pPublicList.Caption = pCurInitiative.Name & ": " & .Caption
            End If
        End If
    End With
End Sub

Private Sub tmrUpdate_Timer()
    'P‰ivitt‰‰ heti hetken p‰‰st‰
    UpdateCharacter
    With tmrUpdate
        .Interval = 0
        .Enabled = False
    End With
End Sub

'K‰ynnist‰‰ ajastimen, joka p‰ivitt‰‰ hahmon ja lopettaa (yll‰)
Private Sub UpdateCharacterTimed()
    With tmrUpdate
        .Enabled = True
        .Interval = 100
    End With
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Caption
    Case "Add"
        Set SelectedCharacter = AddCharacter
        ShowAbilitiesDialog
    Case "Delete"
        DeleteCharacter
    Case "Clone"
        Set SelectedCharacter = CloneCharacter
        ShowAbilitiesDialog
    Case "Update"
        UpdateCharacter
    Case "Prev"
        PrevInitiative
    Case "Next"
        NextInitiative
    End Select
End Sub

Private Function CloneCharacter(Optional Character As clsCharacter) As clsCharacter

    Dim nCharacter As clsCharacter
    If Character Is Nothing Then
        Set Character = pSelectedCharacter
    End If
    If Not Character Is Nothing Then
        Set nCharacter = modCharacters.CloneCharacter(Character)
        Set CloneCharacter = AddCharacter(nCharacter)
    End If
End Function

Private Sub DeleteCharacter(Optional Character As clsCharacter)
    Dim i As Integer
    If Character Is Nothing Then
        Set Character = pSelectedCharacter
    End If
    RemoveCharacter Characters, Character
    ListCharacters
End Sub


Private Sub NextInitiative()
    SelectNextCharacter
    
    cmdDeclarationTimer.Caption = "Start!"
    
End Sub

Private Sub PrevInitiative()
    SelectPrevCharacter
    cmdDeclarationTimer.Caption = "Start!"
End Sub

Private Sub SelectPrevCharacter()
    Dim bPrevRound As Boolean
    If Characters Is Nothing Then
        Debug.Print "No characters"
        Exit Sub
    End If
    If Characters.Count = 0 Then
        Debug.Print "0 characters"
        Exit Sub
    End If
    If pCurInitiative Is Nothing Then
        pRound = 1
        Set CurInitiative = GetFirstInitiative(Characters)
    Else
        Set CurInitiative = GetPrevInitiative(Characters, pCurInitiative, bPrevRound)
    End If
    If bPrevRound Then
        PreviousRound
    End If
    Set SelectedCharacter = pCurInitiative
    ListCharacters
End Sub

Private Sub SelectNextCharacter()
    Dim bNextRound As Boolean
    If Characters Is Nothing Then
        Debug.Print "No characters"
        Exit Sub
    End If
    If Characters.Count = 0 Then
        Debug.Print "0 characters"
        Exit Sub
    End If
    If pCurInitiative Is Nothing Then
        Set CurInitiative = GetFirstInitiative(Characters)
    Else
        Set CurInitiative = GetNextInitiative(Characters, pCurInitiative, bNextRound)
    End If
    If Not pCurInitiative Is Nothing Then
        pCurInitiative.ReactionUsed = False
    End If
    If bNextRound Then
        NextRound
    End If
    Set SelectedCharacter = pCurInitiative
    ListCharacters
End Sub

Public Property Get CurInitiative() As clsCharacter
    Set CurInitiative = pCurInitiative
End Property

Public Property Set CurInitiative(Value As clsCharacter)
    Dim pText As String
    If pCurInitiative Is Nothing Then
        If pRound = 0 Then
            pRound = 1
        End If
    End If
    Set pCurInitiative = Value
    Set GetCurCharacterForm.Character = Value
    If Value Is Nothing Then
        GetCurCharacterForm.Caption = "No active characters"
    Else
        GetCurCharacterForm.Caption = "Current initiative: " & Value.Name
    End If
    
    If Value Is Nothing Then
        pText = "No initiative sequence"
    Else
        pText = "Round " & pRound & ":" & Value.Initiative & " " & Value.Name
    End If
    If Not pImmediate Is Nothing Then
        pImmediate.Caption = pText
    End If
    If Not pPublicList Is Nothing Then
        pPublicList.Caption = pText
    End If
        
    Me.Caption = pText
End Property

Public Sub NextRound()
    MsgBox "Seuraava kierros!"
    pRound = pRound + 1
    LogEvent "Round " & pRound
End Sub

Public Sub PreviousRound()
    MsgBox "Edellinen kierros!"
    pRound = pRound - 1
    LogEvent "Resume round " & pRound
End Sub

Public Function AddCharacter(Optional Character As clsCharacter) As clsCharacter
    Dim nChar As clsCharacter
    If Character Is Nothing Then
        Set nChar = New clsCharacter
        If Not UpdateToCharacter(nChar) Then
            Exit Function
        End If
        Randomize Timer
        nChar.Initiative = nChar.InitiativeBase + CSng(Format(Rnd * 20, "0.00"))
    Else
        Set nChar = Character
    End If
    Set nChar.InitiativeForm = Me
    nChar.Name = GetNextFreeName(nChar.Name, Characters)
    nChar.Initiative = nChar.Initiative - Format(Rnd, "0.00")
    Characters.Add nChar
    Set SelectedCharacter = nChar
    ListCharacters
    Set AddCharacter = nChar
    If Not MyMap Is Nothing Then
        MyMap.ShowCharacters
    End If
End Function

Public Sub ShowAbilitiesDialog()
    Dim d As dlgAbilities
    Set d = New dlgAbilities
    If d.ShowDialog(pSelectedCharacter) Then
        ShowCharacter
        'pUpdate = True
        ListCharacters
    End If
    Unload d

End Sub


Public Sub UpdateCharacter()
    If Not pSelectedCharacter Is Nothing Then
        UpdateToCharacter pSelectedCharacter
    End If
    ListCharacters
End Sub

Public Sub ListCharacters()
    modCharacters.ListCharacters Characters, lvInitiative, pSelectedCharacter, pCurInitiative
    If Not pPublicList Is Nothing Then
        modCharacters.ListCharacters Characters, pPublicList.ListView, pSelectedCharacter, pCurInitiative, False
    End If
End Sub

Public Property Set SelectedCharacter(Character As clsCharacter)
    Dim iChar As clsCharacter
    Set pSelectedCharacter = Character
    ShowCharacter
    
    If Not MyMap Is Nothing Then
        If Not pCurInitiative Is Nothing Then
            Me.Caption = "Range from current initiative: " & Format(Sqr(Abs(pCurInitiative.GetDragForm(True).CenterX - pSelectedCharacter.GetDragForm(True).CenterX) ^ 2 + Abs(pCurInitiative.GetDragForm(True).CenterY - pSelectedCharacter.GetDragForm(True).CenterY) ^ 2) / MyMap.Scaling, "0.0")
            MyMap.Caption = Me.Caption
        End If
        'Hoidetaan jonkinlainen n‰kyvyys!
        For Each iChar In Characters
            With iChar.GetDragForm(True)
                .Highlighted = (Character Is iChar)
                If iChar.IsActive Then
                    .SetTranslucent vbWhite, 0, LWA_COLORKEY
                Else
                    .SetTranslucent vbWhite, 120, LWA_BOTH
                End If
                '.Visible = iChar.IsActive
                .lblStatus.Caption = iChar.Status
                If iChar.Hits <= 0 Then
                     .lblStatus.Caption = "U"
                End If
                If .lblStatus.Caption <> "" Then
                    .lblStatus.Visible = True
                Else
                    .lblStatus.Visible = False
                End If
            End With
        Next
        If Not pSelectedCharacter Is Nothing Then
            pSelectedCharacter.GetDragForm(True).ZOrder
            If Toolbar1.Buttons("center").Value = tbrPressed Then
                MyMap.EnsureVisible pSelectedCharacter
            End If
        End If
        
    End If
End Property

Public Property Get SelectedCharacter() As clsCharacter
    Set SelectedCharacter = pSelectedCharacter
End Property

Private Function UpdateToCharacter(Character As clsCharacter) As Boolean
    UpdateToCharacter = UpdateCharacterListView(Character, lvCharacter, chkActive)
    With Character
        .ReactionUsed = IIf(chkReactionUsed.Value = vbChecked, True, False)
    End With
End Function

'Same sub is found in frmCharacter. Dirty, but copy changes to that form
' when made here with following:
' pSelectedCharacter => pCharacter
Public Sub ShowCharacter()
    Dim nItm As ListItem
    Dim rCharacter As clsCharacter
    Dim iEntry As clsKeyValuePair
    
    If pSelectedCharacter Is Nothing Then
        Set rCharacter = CreateCharacter
    Else
        Set rCharacter = pSelectedCharacter
    End If
    
    With rCharacter
        lvCharacter.ListItems.Clear
        Set nItm = lvCharacter.ListItems.Add(, "Name", .Name)
        nItm.SubItems(1) = "Name"
        Set nItm = lvCharacter.ListItems.Add(, "Initiative", .Initiative)
        nItm.SubItems(1) = "Initiative"
        Set nItm = lvCharacter.ListItems.Add(, "InitiativeBase", .InitiativeBase)
        nItm.SubItems(1) = "Initiative base"
        Set nItm = lvCharacter.ListItems.Add(, "Hits", .Hits)
        nItm.SubItems(1) = "Hits"
        Set nItm = lvCharacter.ListItems.Add(, "TempHits", .TempHits)
        nItm.SubItems(1) = "TempHits"
        Set nItm = lvCharacter.ListItems.Add(, "MaxHits", .MaxHits)
        nItm.SubItems(1) = "MaxHits"
        Set nItm = lvCharacter.ListItems.Add(, "Status", .Status)
        nItm.SubItems(1) = "Status"
        Set nItm = lvCharacter.ListItems.Add(, "Size", .Size)
        nItm.SubItems(1) = "Size"
        Set nItm = lvCharacter.ListItems.Add(, "AC", .AC)
        nItm.SubItems(1) = "AC"
        Set nItm = lvCharacter.ListItems.Add(, "Speed", .Speed)
        nItm.SubItems(1) = "Speed"
        Set nItm = lvCharacter.ListItems.Add(, "HD", .HD)
        nItm.SubItems(1) = "HD"
        Set nItm = lvCharacter.ListItems.Add(, "CR", .CR)
        nItm.SubItems(1) = "CR"
        Set nItm = lvCharacter.ListItems.Add(, "Abilities", .Abilities)
        nItm.SubItems(1) = "Abilities"
        chkReactionUsed.Value = IIf(.ReactionUsed, vbChecked, vbUnchecked)
        If .PictureFile = "" Then
            Set picCharacter.Picture = picFighter.Picture
        Else
            picCharacter.Picture = LoadPicture()
            If FileExists(.PictureFile) Then
                PaintWithAspectRatio picCharacter, .PictureFile, 0, 0, picCharacter.ScaleWidth, picCharacter.ScaleHeight
                'picCharacter.PaintPicture .GetPicture, 0, 0, picCharacter.ScaleWidth, picCharacter.ScaleHeight
                .RefreshForm
            Else
                picCharacter.PaintPicture picFighter.Picture, 0, 0, picCharacter.ScaleWidth, picCharacter.ScaleHeight
            End If
        End If
        chkActive.Value = IIf(.IsActive, vbChecked, vbUnchecked)
        lstArticles.Clear
        txtArticle.Text = ""
        For Each iEntry In .GetEntries
            lstArticles.AddItem iEntry.Key
        Next
        If lstArticles.ListCount > 0 Then
            lstArticles.ListIndex = 0
            lstArticles_Click
        End If
    End With
    
End Sub

