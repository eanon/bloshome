VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{E01CB74C-0B6E-4933-899F-96A702EAA873}#4.0#0"; "DiFtpCli6_FFhMOD.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BlosHome"
   ClientHeight    =   7725
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   11430
   Icon            =   "main.frx":0000
   LinkTopic       =   "main"
   MaxButton       =   0   'False
   ScaleHeight     =   7725
   ScaleWidth      =   11430
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDownloadArt 
      Caption         =   "< &Move"
      Enabled         =   0   'False
      Height          =   510
      Index           =   1
      Left            =   5220
      TabIndex        =   13
      Top             =   4905
      Width           =   975
   End
   Begin VB.CommandButton cmdCleanupBlog 
      Caption         =   "Cle&an"
      Enabled         =   0   'False
      Height          =   615
      Left            =   5220
      TabIndex        =   15
      Top             =   6690
      Width           =   975
   End
   Begin VB.CommandButton cmdUploadArt 
      Caption         =   "Previe&w >"
      Enabled         =   0   'False
      Height          =   510
      Index           =   0
      Left            =   5220
      TabIndex        =   10
      Top             =   3030
      Width           =   975
   End
   Begin MSComctlLib.ImageList lstImg 
      Left            =   5430
      Top             =   2205
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":19B2
            Key             =   "Category"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":1E0A
            Key             =   "Article"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":2262
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":232B
            Key             =   "Up"
         EndProperty
      EndProperty
   End
   Begin DiFtpCli6_FFhMOD.FtpCli FtpCli 
      Left            =   5220
      Top             =   2475
      _ExtentX        =   873
      _ExtentY        =   873
      UserName        =   ""
      Password        =   ""
   End
   Begin VB.Timer timerCnnx 
      Enabled         =   0   'False
      Left            =   5760
      Top             =   2550
   End
   Begin VB.CommandButton cmdBackupBlog 
      Caption         =   "BAC&KUP BLOG   <<<<<"
      Enabled         =   0   'False
      Height          =   945
      Left            =   5220
      TabIndex        =   14
      Top             =   5730
      Width           =   975
   End
   Begin VB.CommandButton cmdDownloadArt 
      Caption         =   "< &Copy"
      Enabled         =   0   'False
      Height          =   510
      Index           =   0
      Left            =   5220
      TabIndex        =   12
      Top             =   4380
      Width           =   975
   End
   Begin VB.CommandButton cmdUploadArt 
      Caption         =   "Pu&blish >"
      Enabled         =   0   'False
      Height          =   510
      Index           =   1
      Left            =   5220
      TabIndex        =   11
      Top             =   3555
      Width           =   975
   End
   Begin VB.CommandButton cmdProj 
      Caption         =   "Parameter&s"
      Enabled         =   0   'False
      Height          =   675
      Left            =   120
      TabIndex        =   0
      Top             =   345
      Width           =   1590
   End
   Begin VB.Frame framRemote 
      Caption         =   "Remote blog"
      Enabled         =   0   'False
      Height          =   4935
      Left            =   6240
      TabIndex        =   16
      Top             =   2385
      Width           =   5070
      Begin VB.OptionButton optStatusArt 
         Caption         =   "Public"
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   4110
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   4380
         Value           =   -1  'True
         Width           =   750
      End
      Begin VB.OptionButton optStatusArt 
         Caption         =   "Private"
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   4110
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   4080
         Width           =   750
      End
      Begin VB.CommandButton cmdSeeRemoteArt 
         Caption         =   "&See online"
         Enabled         =   0   'False
         Height          =   255
         Left            =   945
         TabIndex        =   26
         Top             =   4500
         Width           =   1110
      End
      Begin VB.CommandButton cmdDelRemoteArt 
         Caption         =   "Dele&te"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2070
         TabIndex        =   27
         Top             =   4500
         Width           =   900
      End
      Begin VB.CommandButton cmdRenRemoteArt 
         Caption         =   "Rena&me"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2985
         TabIndex        =   28
         Top             =   4500
         Width           =   990
      End
      Begin VB.CommandButton cmdRenRemoteCat 
         Caption         =   "Re&name"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2610
         TabIndex        =   23
         Top             =   4170
         Width           =   990
      End
      Begin VB.CommandButton cmdRmRemoteCat 
         Caption         =   "&Delete"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1680
         TabIndex        =   22
         Top             =   4170
         Width           =   900
      End
      Begin VB.CommandButton cmdMkRemoteCat 
         Caption         =   "&Create"
         Enabled         =   0   'False
         Height          =   255
         Left            =   945
         TabIndex        =   21
         Top             =   4170
         Width           =   705
      End
      Begin MSComctlLib.ListView lstRemote 
         Height          =   3465
         Left            =   105
         TabIndex        =   18
         Top             =   570
         Width           =   4845
         _ExtentX        =   8546
         _ExtentY        =   6112
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "lstImg"
         SmallIcons      =   "lstImg"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         Enabled         =   0   'False
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Date"
            Object.Width           =   1852
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Time"
            Object.Width           =   1094
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Size"
            Object.Width           =   838
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Joint"
            Object.Width           =   917
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Enc."
            Object.Width           =   1058
         EndProperty
      End
      Begin VB.Shape shapStatus 
         BorderColor     =   &H00404040&
         BorderStyle     =   3  'Dot
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   765
         Left            =   4035
         Top             =   3990
         Width           =   900
      End
      Begin VB.Label lblRemoteArts 
         AutoSize        =   -1  'True
         Caption         =   "Article :"
         Enabled         =   0   'False
         Height          =   195
         Left            =   375
         TabIndex        =   25
         Top             =   4515
         Width           =   525
      End
      Begin VB.Label lblRemoteCats 
         AutoSize        =   -1  'True
         Caption         =   "Category :"
         Enabled         =   0   'False
         Height          =   195
         Left            =   135
         TabIndex        =   19
         Top             =   4185
         Width           =   720
      End
      Begin VB.Label lblRemotePath 
         Height          =   300
         Left            =   210
         TabIndex        =   17
         Top             =   270
         Width           =   4515
      End
   End
   Begin VB.Frame framLocal 
      Caption         =   "Local workspace"
      Enabled         =   0   'False
      Height          =   4935
      Left            =   105
      TabIndex        =   2
      Top             =   2385
      Width           =   5070
      Begin VB.CommandButton cmdCopyLocalArt 
         Caption         =   "Cop&y"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1440
         TabIndex        =   6
         Top             =   4500
         Width           =   645
      End
      Begin VB.CommandButton cmdEditLocalArt 
         Caption         =   "E&dit"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2115
         TabIndex        =   7
         Top             =   4500
         Width           =   645
      End
      Begin VB.CommandButton cmdRenLocalArt 
         Caption         =   "R&ename"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3720
         TabIndex        =   9
         Top             =   4500
         Width           =   990
      End
      Begin VB.CommandButton cmdDelLocalArt 
         Caption         =   "De&lete"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2790
         TabIndex        =   8
         Top             =   4500
         Width           =   900
      End
      Begin VB.CommandButton cmdMkLocalArt 
         Caption         =   "C&reate"
         Enabled         =   0   'False
         Height          =   255
         Left            =   780
         TabIndex        =   5
         Top             =   4500
         Width           =   630
      End
      Begin MSComctlLib.ListView lstLocal 
         Height          =   4095
         Left            =   105
         TabIndex        =   3
         Top             =   285
         Width           =   4845
         _ExtentX        =   8546
         _ExtentY        =   7223
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "lstImg"
         SmallIcons      =   "lstImg"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         Enabled         =   0   'False
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Date"
            Object.Width           =   1852
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Time"
            Object.Width           =   1094
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Size"
            Object.Width           =   838
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Joint"
            Object.Width           =   917
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Enc."
            Object.Width           =   1058
         EndProperty
      End
      Begin VB.Label lblLocalArts 
         AutoSize        =   -1  'True
         Caption         =   "Article :"
         Enabled         =   0   'False
         Height          =   195
         Left            =   210
         TabIndex        =   4
         Top             =   4515
         Width           =   525
      End
   End
   Begin MSComctlLib.StatusBar barStatus 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      Negotiate       =   -1  'True
      TabIndex        =   29
      Top             =   7440
      Width           =   11430
      _ExtentX        =   20161
      _ExtentY        =   503
      ShowTips        =   0   'False
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
            Text            =   "Offline"
            TextSave        =   "Offline"
            Key             =   "cnnx"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   15875
            MinWidth        =   15875
            Picture         =   "main.frx":2647
            Key             =   "url"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdCnnx 
      Caption         =   "Co&nnection"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   1065
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   1590
   End
   Begin VB.Image imgDeco 
      Height          =   2100
      Left            =   1800
      Picture         =   "main.frx":29E3
      Top             =   45
      Width           =   9525
   End
   Begin VB.Line lineMenu 
      X1              =   -30
      X2              =   11430
      Y1              =   30
      Y2              =   30
   End
   Begin VB.Menu mnuProj 
      Caption         =   "&Project"
      Begin VB.Menu mnuProj_New 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuProj_Open 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuProj_Del 
         Caption         =   "&Delete"
      End
   End
   Begin VB.Menu mnuLang 
      Caption         =   "&Language"
      Begin VB.Menu mnuLang_EN 
         Caption         =   "&EN"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuLang_FR 
         Caption         =   "&FR"
      End
   End
   Begin VB.Menu mnuInfo 
      Caption         =   "&?"
      Begin VB.Menu mnuInfo_Help 
         Caption         =   "&Help"
      End
      Begin VB.Menu mnuInfo_About 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'BlosHome (c) FFh Lab / Eric Lequien, 2009-2013 - http://ffh-lab.com
'This frame represents the main interface

Option Explicit
Option Base 1

Private Type SECUFILE   'describe useful elements for checking file integrity (after transmission)
    file As String          'filename, with or without path according to needing (without path if current remote dir)
    goodsize As Long        'expected file size for the given file (will be compared with observed one after upload)
    integrity As Integer    'result of size checking (-1:corrupted, 0:undefined, 1: OK, matches the good size)
End Type                    '***LATER : we could improve integrity checking using CRC32 or better (but slower)

Private bMouseoverStatusOpt As Boolean 'allows to differenciate manual and auto-click on optStatusArt

Private Sub barStatus_PanelDblClick(ByVal Panel As MSComctlLib.Panel)
    If Panel.Key = "url" Then
        'Attempts to reach the blog using the default Web browser
        If Left$(barStatus.Panels.item("url").Text, 7) = "http://" Then
            OpenMIME barStatus.Panels.item("url").Text
        End If
    End If
End Sub

Private Sub cmdBackupBlog_Click()
    'Backups the entirety of remote blog toward a local diectory
    '***TODO : write the function with these elements in mind :
    '- local directory could be a subdir of project workspace, or beside it, or an app subdir like "backups"
    '- backup could be compressed (using ZIP format, since already used in BlosHome), taking care of paths and timestamp
    '- name of compressed archive could be something like "<project_name>_<date>_<hour>.zip"
    '- backup process should save cgi-bin (blosxom-related) content too, including eventual timestamps-index plugin data
    '- to preserve articles timestamps (if no timestamp-index plugin), we'll build a timestamps.txt for future reference
    '***LATER : a restore feature should be added ; it should do this : upload, chemod of what required, touch every file
    '           from our referential "timestamps.txt" if not any "timestamp-index" plugin data has been found in backup.
    MsgBox arMsg(70) & ".", vbInformation
End Sub

Private Sub cmdCleanupBlog_Click()
    'Cleans the remote blog and local data
    '***TODO : write this function with these elements in mind :
    '- deletes eventual temporary files, remote orphan images (not binded with an article),
    '- erase obsolets infos in local and remote meta.cache
    '... [extendable list, according to BlosHome working and eventual subsequent waste it could imply)
    MsgBox arMsg(70) & ".", vbInformation
End Sub

Private Sub cmdCnnx_Click()
    'Attempts to connect/disconnect to FTP server
    If Left(cmdCnnx.Caption, 1) = "C" Then
        CnnxDcnnx True  'Connexion
    Else
        CnnxDcnnx False 'Déconnexion
    End If
End Sub

Private Sub cmdCopyLocalArt_Click()
    'Duplicates a local article (for the puirpose to create a new one on this base)
    Dim strKey As String
    Dim strArt As String
    
    Dim strOrgArt As String
    Dim strNewArt As String
    
    Dim strTitle As String
    Dim strMsg As String
    Dim strDefault  As String
    
    Dim bOK As Boolean
    
    'checking
    If lstLocal.ListItems.Count = 0 Then Exit Sub
    
    If lstLocal.SelectedItem.Selected = False Then
        MsgBox arMsg(71) & " !", vbExclamation
        Exit Sub
    End If
    
    strKey = lstLocal.SelectedItem.Key
    strOrgArt = GetWorkspace() & "\" & strKey
    
    If Dir(strCurrArt) = "" Then
        MsgBox arMsg(72) & " !", vbExclamation
        PopulateClearLocalList True
        Exit Sub
    End If
    
    'user input
    strTitle = arMsg(5)
    strMsg = arMsg(6) & " '" & strKey & "'" & vbNewLine _
                & vbNewLine _
                & arMsg(7) & " : " & vbNewLine _
                & "- " & arMsg(8) & " /*\<>|?" & Chr(34) & ":" & vbNewLine _
                & "- " & arMsg(9) & vbNewLine _
                & "- " & arMsg(10) & vbNewLine _
                & "- " & arMsg(11) & vbNewLine _
                & "- " & arMsg(12) & vbNewLine _
                & "- " & arMsg(13) & " '" & structProj.art_ext & "' " & arMsg(14) & vbNewLine _
                & vbNewLine _
                & arMsg(15) & structProj.art_ext & "'"

    strDefault = ""
    bOK = False
    
    Do
        strArt = InputBox(strMsg, strTitle, strDefault)
        If strArt = "" Then Exit Sub
        strDefault = strArt 'save for eventual next loop
        strArt = CheckAndMkFilename(UnAccent(LCase(strArt)), False, True)
        
        If strArt <> "" Then
            If IsNumeric(Left$(strArt, 1)) = False Then
                bOK = True
            Else
                MsgBox arMsg(73) & " !", vbExclamation
                bOK = False
            End If
        Else
            MsgBox arMsg(74) & " !", vbExclamation
            bOK = False
        End If
        
        If Right$(strArt, 4) <> structProj.art_ext Then
            strArt = strArt & structProj.art_ext
        End If
        
        strNewArt = GetWorkspace() & "\" & strArt
        
        If Dir(strNewArt) <> "" Then
            MsgBox arMsg(75) & " !", vbExclamation
            bOK = False
        End If
    Loop Until bOK = True
    
    'effective application
    Screen.MousePointer = vbHourglass
    On Error GoTo cmdCopyLocalArt_Click_Error
    FileCopy strOrgArt, strNewArt
    If Dir(strNewArt) = "" Then Err.Raise 513
    Screen.MousePointer = vbDefault
    
    PopulateClearLocalList True

cmdCopyLocalArt_Click_End:
   On Error GoTo 0
   Exit Sub

cmdCopyLocalArt_Click_Error:
    MsgBox arMsg(76) & " '" & strKey & "' !", vbExclamation
    Resume cmdCopyLocalArt_Click_End
End Sub

Private Sub cmdDelLocalArt_Click()
    'Deletes the selected local article
    'WARNING : if in a circumstance associated images was moved beside article in the local workspace, it would be
    '          useful to delete-it too (at this time, images stay in their original location on disk and relative path
    '          are modified on fly during upload (direct publish or publish as preview cmds), and in zipped archive).
    Dim strKey As String
    Dim strArt As String
    
    If lstLocal.ListItems.Count = 0 Then Exit Sub
    
    If lstLocal.SelectedItem.Selected = False Then
        MsgBox arMsg(71) & " !", vbExclamation
        Exit Sub
    End If
    
    strKey = lstLocal.SelectedItem.Key
    strArt = GetWorkspace() & "\" & strKey
    
    If Dir(strArt) = "" Then
        MsgBox arMsg(72) & " !", vbExclamation
        PopulateClearLocalList True
    Else
        Dim nRet As Integer
        
        nRet = MsgBox(arMsg(77) & " '" & strKey & "' ?", vbQuestion + vbYesNo)
        If nRet = vbYes Then
            Kill strArt
            PopulateClearLocalList True
        End If
    End If
End Sub

Private Sub cmdDelRemoteArt_Click()
    'Deletes the selected remote article
    '***LATER : réindexes (or manage concerned cache/index data files directly) if a cache/index data file(s),
    '           autogenerated by cache/index plugins, are not automatically aware of deletion. For example,
    '           in my blogs, "/cgi-bin/blog/002entries_timestamp" generates "state/entries_timestamp.index".
    Dim strKey As String
    Dim strType As String
    
    Dim strArt As String         'remote path
    Dim strFilename As String    'filmename only
    Dim strDisplayName As String 'different than filename in case of a private (preview) article ; because of prefix
    
    Dim strTmpFile As String
    Dim arNotDeleted() As String
    
    Dim nIdx As Integer
    Dim strMsg As String
    Dim nRet As Integer
    Dim bRet As Boolean
    
    'checking
    If lstRemote.SelectedItem.Selected = False Then
        MsgBox arMsg(71) & " !", vbExclamation
        Exit Sub
    End If
    
    strType = lstRemote.SelectedItem.SmallIcon
    If strType <> "Article" And strType <> "Preview" Then
        MsgBox arMsg(78) & " !", vbInformation
        Exit Sub
    End If
    
    'init
    strKey = lstRemote.SelectedItem.Key
    strDisplayName = lstRemote.SelectedItem.Text
    strFilename = FtpCli.RemoteFiles(strKey).FileName
    strArt = FtpCli.RemoteDir & "/" & strFilename
    
    strMsg = arMsg(16) & " '" & strDisplayName & "' ?"
                              
    nRet = MsgBox(strMsg, vbQuestion + vbYesNo)
    If nRet = vbNo Then Exit Sub
    
    On Error GoTo cmdDelRemoteArt_Click_Error
    Screen.MousePointer = vbHourglass
    
    'deletes the article and its eventual attached images
    bRet = DelCurrRemoteArt(arNotDeleted)
    If bRet = False Then
        If arNotDeleted(1) = "failed" Then
            Err.Raise 513
        Else
            Err.Raise 514
        End If
    End If
    
    'user info
    strMsg = arMsg(17) & " '" & strDisplayName & "' " & arMsg(18) & " !"
    MsgBox strMsg, vbInformation

cmdDelRemoteArt_Click_End:
    If strTmpFile <> "" And Dir(strTmpFile) <> "" Then Kill strTmpFile
    Screen.MousePointer = vbDefault
    On Error GoTo 0
   Exit Sub

cmdDelRemoteArt_Click_Error:
    Select Case Err.Number
        Case 513
            strMsg = arMsg(19) & " '" & strDisplayName & "'" & vbNewLine & arMsg(20) & " !"
        Case 514
            If arNotDeleted(1) = strFilename Then
                strMsg = arMsg(21) & " '" & strDisplayName & "'" & vbNewLine & arMsg(20) & " !"
            Else
                strMsg = arMsg(22) & " '" & strDisplayName & "' " & arMsg(23) & "," & vbNewLine & _
                         arMsg(24) & " :"
                For nIdx = LBound(arNotDeleted) To UBound(arNotDeleted)
                    strMsg = strMsg & vbNewLine & "- '" & arNotDeleted(nIdx) & "'"
                Next
                strMsg = strMsg & vbNewLine & vbNewLine & arMsg(25) & vbNewLine & _
                         arMsg(26) & " '" & arMsg(27) & "'"
            End If
        Case Else
            MsgBox "Error #" & Err.Number & "@ bloshome/frmMain/cmdDelRemoteArt_Click/#" & Erl & " : " & Err.Description, vbExclamation
    End Select
    
    MsgBox strMsg, vbExclamation
    Resume cmdDelRemoteArt_Click_End
End Sub

Private Sub cmdDownloadArt_Click(index As Integer)
    'Rapatriates a copie of a published article (for the purpose to review-it or to use-it as base for a new article)
    '***TODO : implement case which need real download if alternative fast-way appears to be impossible because of
    '          absence of this article in zipped archives ; we'll branch to future fct from cmdUploadArt_Click/***LATER
    Dim strKey As String
    Dim strDisplayName As String
    Dim strFilename As String
    Dim strArt As String
    Dim strWorkspace As String
    
    Dim strIdArchiv As String
    Dim bRet As Boolean
    
    Dim strTempFile As String    'useful to tests for checking if zipped article in archive is the right one :
    Dim nPreparedArtSize As Long
    Dim nCurrPreparedArtSize As Long
    Dim nOrgAddedNameLen As Integer
    Dim nCurrAddedNameLen As Integer
    Dim arImg() As TRANSFILE
    Dim nImgs As Integer
    Dim strEnc As String
    Dim strIgnored As String
    
    Dim bSizeOK As Boolean
    Dim bImgOK As Boolean
    Dim bEncOK As Boolean
    Dim bDone As Boolean
    
    Dim strServerAttachs As String 'list using "," as separator

    Dim strMsg As String
    Dim nIdx As Integer

    'init and security checking
    On Error Resume Next           'avoid err #91 if lstRemote is empty
    If lstRemote.SelectedItem.Selected = False Then
        MsgBox arMsg(71) & " !", vbExclamation
        Exit Sub
    End If
    On Error GoTo 0
    
    strKey = lstRemote.SelectedItem.Key
    strDisplayName = lstRemote.SelectedItem.Text
    strFilename = FtpCli.RemoteFiles(strKey).FileName
    strArt = FtpCli.RemoteDir & "/" & strFilename
    strWorkspace = GetWorkspace() & "\"
    
    If Dir(strWorkspace & strFilename) <> "" Then
        MsgBox arMsg(79) & " ! ", vbInformation
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    'attempts to avoid real download using local zipped archive
    strIdArchiv = BuildIdArchive(strFilename)
    bRet = SetGetArchive(strDisplayName, strIdArchiv, False, nOrgAddedNameLen, strServerAttachs)
    If bRet = True Then
        'checks if zipped article coming from archive is the same one as the one requested for download
        'IMPORTANT : take cares to do that the tests itselves don't need any download !
        
        'test if size of the article in format as ready-to-be-published = size of article currenly on blog
        'NB : calculation of nPreparedArtSize takes the prepared size coming from archive without consideration
        '     of any eventual modifications in image names, then we add the true value of these modifications
        '     as it was observed during real publication. Finally, a size adjustement is done deducting the
        '     characters which has been eliminated during transfer because of OS difference ; Win (CRLF) -> *nix (LF)
        '***LATER : extends the test to article's CRC/MD5 computation compared to CRC/MD5 info stored in meta.cache
        strTempFile = GetTmpFile("FB")
        FileCopy strWorkspace & strDisplayName, strTempFile
        arImg = PrepArtForPublishing(strTempFile, nCurrPreparedArtSize, nCurrAddedNameLen, strIgnored, True)
        Kill strTempFile
        
        nPreparedArtSize = (nCurrPreparedArtSize - nCurrAddedNameLen) + nOrgAddedNameLen _
                           - CountOccurr(LoadText(strWorkspace & strDisplayName), vbNewLine)
                           
        If arImg(1).local = "failed" Or nPreparedArtSize <> FtpCli.RemoteFiles(strKey).FileSize Then
            bSizeOK = False
        Else
            bSizeOK = True
        End If
                
        'tests if attached elements (ie. images, but could be expanded to others types) are identical
        If arImg(1).local <> "" Then
            nImgs = UBound(arImg)
        Else
            nImgs = 0
        End If
        
        If nImgs = Int(MetaCache(strArt, "Joint", "GET")) Then
            bImgOK = True
        Else
            bImgOK = False
        End If
        
        'tests the encoding
        If IsUTF8(LoadText(strWorkspace & strDisplayName)) = True Then
            strEnc = "UTF-8"
        Else
            strEnc = "ANSI"
        End If
        
        If strEnc = UCase(MetaCache(strArt, "Encode", "GET")) Then
            bEncOK = True
        Else
            bEncOK = False
        End If
        
        'results
        If bSizeOK = True And bImgOK = True And bEncOK = True Then
            bDone = True
        Else
            'deletes the article coming from archive, temporary stored in local workspace
            'NB : the article coming from archive is not the right one => we deletes the extracted one in workspace,
            '     and its eventual attached images, before to proceed with a real download of the remote article !
            'WARNING : takes care to not delete a file which would exist in workspace before archive extraction ;
            '          maybe that SetGetArchive should return arExtracted() in case of 'get'
            '***LATER : deletes the eventual extracted images too
            ')
            Kill strWorkspace & strDisplayName
            bDone = False
        End If
    End If
    
    'goes with real download (since alternative fast-way above failed)
    '***TODO : This part has to be implemented ! We could creates a funtion which would combine download and upload
    '          operation, reusing code coming from cmdUploadArt_Click() ; this fct could be prototyped DownUpLoad(bOp)
    If bDone = False Then
        MsgBoxEx arMsg(80) & vbNewLine & arMsg(81) & vbNewLine & arMsg(82) & " !", vbInformation
        Exit Sub
    End If

    'deletes the remote article and its eventual attached images if user asked for "move"
    If index = 1 Then
        Dim arNotDeleted() As String
        bRet = DelCurrRemoteArt(arNotDeleted, strWorkspace & strDisplayName, strServerAttachs)
    
        If bRet = False Then
            If arNotDeleted(1) = "failed" Then
                strMsg = arMsg(19) & " '" & strDisplayName & "'" & vbNewLine & arMsg(20) & " !"
            Else
                If arNotDeleted(1) = strFilename Then
                    strMsg = arMsg(21) & " '" & strDisplayName & "'" & vbNewLine & arMsg(20) & " !"
                Else
                    strMsg = arMsg(22) & " '" & strDisplayName & "' " & arMsg(23) & "," & vbNewLine & _
                             arMsg(24) & " :"
                    For nIdx = LBound(arNotDeleted) To UBound(arNotDeleted)
                        strMsg = strMsg & vbNewLine & "- '" & arNotDeleted(nIdx) & "'"
                    Next
                    strMsg = strMsg & vbNewLine & vbNewLine & arMsg(25) & vbNewLine & _
                             arMsg(26) & " '" & arMsg(27) & "'"
                End If
            End If
            
            MsgBox strMsg, vbExclamation
        End If
        PopulateRemoteList
    End If

    PopulateClearLocalList True
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdEditLocalArt_Click()
    'Loads the selected local article in the editor
    Dim strKey As String
    
    If lstLocal.ListItems.Count = 0 Then Exit Sub
    
    If lstLocal.SelectedItem.Selected = False Then
        MsgBox arMsg(71) & " !", vbExclamation
        Exit Sub
    End If
    
    strKey = lstLocal.SelectedItem.Key
    strCurrArt = GetWorkspace() & "\" & strKey
    
    If Dir(strCurrArt) = "" Then
        MsgBox arMsg(72) & " !", vbExclamation
        PopulateClearLocalList True
        strCurrArt = ""
        Exit Sub
    End If
    
    If Dir(App.Path & "\" & EDITOR_DIR & "\" & EDITOR_TEMPLATE) = "" Then
        '***LATER : extend the tests to others crucial files in the TinyMCE tree
        MsgBox arMsg(83) & " !", vbExclamation
        PopulateClearLocalList True
        strCurrArt = ""
        Exit Sub
    End If
    
    frmEdit.DoShow Me
End Sub

Private Sub cmdMkLocalArt_Click()
    'Displays the editor for writing a new article
    Dim strTitle As String
    Dim strMsg As String
    Dim strDefault  As String
    
    Dim strArt As String
    Dim bOK As Boolean
    Dim bRet As Boolean
    
    'user input
    strTitle = arMsg(28)
    strMsg = arMsg(29) & "." & vbNewLine _
                & vbNewLine _
                & arMsg(7) & " : " & vbNewLine _
                & "- " & arMsg(8) & " /*\<>|?" & Chr(34) & ":" & vbNewLine _
                & "- " & arMsg(9) & vbNewLine _
                & "- " & arMsg(10) & vbNewLine _
                & "- " & arMsg(11) & vbNewLine _
                & "- " & arMsg(12) & vbNewLine _
                & "- " & arMsg(13) & " '" & structProj.art_ext & "' " & arMsg(14) & vbNewLine _
                & vbNewLine & arMsg(15) & structProj.art_ext & "'"

    strDefault = ""
    bOK = False
    
    Do
        strArt = InputBox(strMsg, strTitle, strDefault)
        If strArt = "" Then Exit Sub
        strDefault = strArt 'saves for eventual next loop
        strArt = CheckAndMkFilename(UnAccent(LCase(strArt)), False, True)
        
        If strArt <> "" Then
            If IsNumeric(Left$(strArt, 1)) = False Then
                bOK = True
            Else
                MsgBox arMsg(73) & " !", vbExclamation
                bOK = False
            End If
        Else
            MsgBox arMsg(74) & " !", vbExclamation
            bOK = False
        End If
        
        If Right$(strArt, 4) <> structProj.art_ext Then
            strArt = strArt & structProj.art_ext
        End If
        
        strCurrArt = GetWorkspace() & "\" & strArt
        
        If Dir(strCurrArt) <> "" Then
            MsgBox arMsg(75) & " !", vbExclamation
            strCurrArt = ""
            bOK = False
        End If
    Loop Until bOK = True
    
    'effective application
    Screen.MousePointer = vbHourglass
    bRet = SaveText(strCurrArt, "")
    Screen.MousePointer = vbDefault
    
    If bRet = False Then
        MsgBox arMsg(84) & " '" & strCurrArt & "' !", vbExclamation
        strCurrArt = ""
        Exit Sub
    End If
    
    PopulateClearLocalList True
    frmEdit.DoShow Me
End Sub

Private Sub cmdMkRemoteCat_Click()
    'Creates a new category (i.e. remote directory) under the current level
    Dim strTitle As String
    Dim strMsg As String
    Dim strDefault As String
    
    Dim strCat As String
    Dim bOK As Boolean
    
    strTitle = arMsg(30)
    strMsg = arMsg(31) & "." & vbNewLine _
                & vbNewLine _
                & arMsg(7) & " : " & vbNewLine _
                & "- " & arMsg(8) & " /*\<>|?" & Chr(34) & ":" & vbNewLine _
                & "- " & arMsg(9) & vbNewLine _
                & "- " & arMsg(10) & vbNewLine _
                & "- " & arMsg(11) & vbNewLine _
                & "- " & arMsg(12) & vbNewLine _
                & vbNewLine _
                & arMsg(32) & "." & vbNewLine _
                & vbNewLine _
                & arMsg(33) & "." & vbNewLine

    strDefault = ""
    bOK = False
    
    Do
        strCat = InputBox(strMsg, strTitle, strDefault)
        If strCat = "" Then Exit Sub
        strDefault = strCat 'saves for eventual next loop
        strCat = CheckAndMkFilename(UnAccent(LCase(strCat)), False, True)
        If strCat <> "" Then
            If IsNumeric(Left$(strCat, 1)) = False Then
                bOK = True
            Else
                MsgBox arMsg(73) & " !", vbExclamation
            End If
        Else
            MsgBox arMsg(74) & " !", vbExclamation
        End If
    Loop Until bOK = True
    
    Screen.MousePointer = vbHourglass
    If FtpCli.Ftp_MKD(strCat) = True Then PopulateRemoteList  'si false, FtpCli affiche l'erreur ad hoc
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdProj_Click()
    'Edits the current project settings
    frmProject.bNoCancel = False
    frmProject.Show 1, Me
End Sub

Private Sub cmdUploadArt_Click(index As Integer)
          'Manages the upload of current article toward current remote category
          'IN : an index indicating the operation to proceed with (according to collection command)
          '     - 0 : publishes a private preview of given article (will be protected by a prefix and password)
          '     - 1 : directly publishes the given article to be reachable by all visitors
          'REQ : custom structure types called TRANSFILE and SECUFILE
          '***LATER : put real-transfer part in a function which could be called from here and cmdDownloadArt_Click
          '           (of course, it requires to implement case 2 and 3 of the Select Case below)
          Dim strCat As String        'category of the article
          Dim strKey As String        'index name in source or target list
          Dim strArt As String        'full path toward original article
          
          Dim strTempName As String   'temporary name alone
          Dim strTempFile As String   'full source path
          Dim strFinalName As String  'name after target renaming
          
          Dim strVerb As String       'type of action
          Dim strObj As String        'type of concerned object
          Dim strFrom As String       'nature of source context
          Dim strTo As String         'nature of target context
          
          Dim nRet As Integer
          Dim bRet As Boolean
          Dim bTransRenaming As Boolean
          
          Dim strMsg As String
          Dim strIdArchiv As String
          
          Dim nPreparedArtSize As Long   'size of final source article (after preparation) in bytes
          Dim nAddedNameLen As Integer   'added length due to eventual image renamings in PrepArtForPublishing
          Dim strServerAttachs As String 'list of name of final attached files ("," separated)
          
          Dim nIdx As Integer
          Dim nIdx2 As Integer
          Dim strErr As String
          
          Dim arImg() As TRANSFILE      'transfer characteristics of images to be transferred
          Dim arUploaded() As SECUFILE  'already uploaded files (for deletion) or error message (in case of failure)
          
          'expresses the operation
1         Select Case index
              Case 0:
2                 strVerb = arMsg(34)
3                 strObj = arMsg(35)
4                 strFrom = arMsg(36)
5                 strTo = arMsg(37)
6             Case 1:
7                 strVerb = arMsg(34)
8                 strObj = arMsg(38)
9                 strFrom = arMsg(36)
10                strTo = arMsg(37)
11            Case 2:
                  '***WARNING : cannot happen ; prepared for future function as expressed in ***LATER ahead
12                strVerb = arMsg(39)
13                strObj = arMsg(40)
14                strFrom = arMsg(37)
15                strTo = arMsg(36)
16            Case 3:
                  '***WARNING : cannot happen ; prepared for future function as expressed in ***LATER ahead
17                strVerb = arMsg(39)
18                strObj = arMsg(41)
19                strFrom = arMsg(37)
20                strTo = arMsg(36)
21            Case Else
22                Exit Sub
23        End Select
          
          'init and security checking
24        On Error Resume Next  'avoid error #91 if lstLocal is empty
25        If lstLocal.SelectedItem.Selected = False Then
26            MsgBox arMsg(71) & " !", vbExclamation
27            Exit Sub
28        End If
29        On Error GoTo 0
          
30        strKey = lstLocal.SelectedItem.Key
31        strArt = GetWorkspace() & "\" & strKey
          
32        If Dir(strArt) = "" Then
33            MsgBox arMsg(85) & " !", vbExclamation
34            PopulateClearLocalList True
35            Exit Sub
36        End If
          
37        If FtpCli.Ftp_SIZE(strKey) = True _
                      Or FtpCli.Ftp_SIZE(structProj.preview_prefix & "-" & strKey) = True Then
38            nRet = MsgBox(arMsg(86) & " ! " & arMsg(87) & " ?" & _
                      vbNewLine & vbNewLine & arMsg(88) & ".", vbQuestion + vbYesNo)
39            If nRet = vbYes Then
                  '***TODO
40                MsgBox arMsg(89) & "...", vbInformation
41                Exit Sub
42            Else
43                Exit Sub
44            End If
45        End If
          
46        strCat = UCase(GetCurrRemoteCat(":", True))
          
          'prepares the article
47        Screen.MousePointer = vbHourglass
          
48        strTempName = strKey & ".tmp"
49        strTempFile = GetWorkspace() & "\" & strTempName
          
50        FileCopy strArt, strTempFile
51        arImg = PrepArtForPublishing(strTempFile, nPreparedArtSize, nAddedNameLen, strServerAttachs, True)
          
52        If arImg(1).local = "failed" Then
53            Kill strTempFile
54            Screen.MousePointer = vbDefault
55            Exit Sub
56        End If
          
          'prepares eventual images
          '***LATER : gathers redimensionned copies of the images in the workspace
          
          'summary before to go
57        strMsg = arMsg(42) & " " & strVerb & " " & strObj & " '" & strKey & "'"
          
58        If arImg(1).local <> "" Then
59            If UBound(arImg) > 1 Then
60                strMsg = strMsg & vbNewLine & vbNewLine & "et ses images associées :"
61                For nIdx = 1 To UBound(arImg)
62                    strMsg = strMsg & vbNewLine & "- " & arImg(nIdx).local
                  
63                    If arImg(nIdx).remote <> Right(arImg(nIdx).local, Len(arImg(nIdx).local) _
                                               - RevInStr(arImg(nIdx).local, "/", False)) Then
64                        strMsg = strMsg & " -> " & arImg(nIdx).remote
65                        bTransRenaming = True
66                    End If
67                Next
68            Else
69                strMsg = strMsg & vbNewLine & vbNewLine & "et son image associée : " & arImg(1).local
70                If arImg(1).remote <> Right$(arImg(1).local, Len(arImg(1).local) _
                                           - RevInStr(arImg(1).local, "/", False)) Then
71                    strMsg = strMsg & " -> " & arImg(1).remote
72                    bTransRenaming = True
73                End If
74            End If
75        End If
          
76        strMsg = strMsg & vbNewLine & vbNewLine & "dans la catégorie '" & strCat & "' ?"
          
77        If bTransRenaming = True Then
78            strMsg = strMsg & vbNewLine & vbNewLine & "__" & vbNewLine & _
                      arMsg(90) & "," & vbNewLine & arMsg(91) & " '->'"
79        End If
          
80        nRet = MsgBoxEx(strMsg, vbQuestion + vbYesNo)
              
          'effective transfers
81        On Error GoTo UploadArtErr
82        If nRet = vbYes Then
              'determine the progression using equal weight for uploads and finalisations
              Dim nProgressMax As Integer   'max width of the gauge (in twips)
              Dim nProgressSteps As Integer 'number of steps to provide
              Dim nProgressStep As Integer  'value of one step (in twips)
              
83            nProgressMax = frmProgress.shapFrame.Width
              
84            If arImg(1).local = "" Then
85                nProgressSteps = 2                        'upload article -> finalisation
86            Else
87                nProgressSteps = 1 + UBound(arImg) + 1    'upload article -> upload images -> finalisation
88            End If
              
89            nProgressStep = nProgressMax / nProgressSteps
90            frmProgress.Show 0, Me
              
              'ASCII upload of temporary copie of the article
91            frmProgress.lblStage.Caption = arMsg(43)
92            frmProgress.lblDetail.Caption = strTempName
              
93            bRet = FtpCli.Met_PUT(strTempFile, , "A")     'using the "A" parameter of ELN_MODded Met_PUT
94            Kill strTempFile
95            If bRet = False Then
96                strMsg = arMsg(44) & " ! " & arMsg(45) & "..."
97                Screen.MousePointer = vbDefault
98                MsgBox strMsg, vbExclamation
99                Exit Sub
100           Else
101               ReDim arUploaded(1 To 1)
102               arUploaded(1).file = strTempName          'goodsize takes care it was an ASCII transfer
103               arUploaded(1).goodsize = nPreparedArtSize - CountOccurr(LoadText(strArt), vbNewLine)
104           End If
              
105           frmProgress.shapBar.Width = frmProgress.shapBar.Width + nProgressStep
              
              'binary upload of attached images
106           If arImg(1).local <> "" Then
107               For nIdx = 1 To UBound(arImg)
108                   frmProgress.lblDetail.Caption = arImg(nIdx).local & " -> " & arImg(nIdx).remote

109                   bRet = FtpCli.Met_PUT(arImg(nIdx).local, arImg(nIdx).remote, "I")
110                   If bRet = False Then
                          'FAILED (publication aborted)
111                       strErr = arMsg(46) & " '" & arImg(nIdx).local & "' !"
112                       Err.Raise 513
113                   Else
                          'update of the already uploaded files list (to delete in case of error)
114                       ReDim Preserve arUploaded(1 To UBound(arUploaded) + 1)
115                       arUploaded(UBound(arUploaded)).file = arImg(nIdx).remote 'binary transfered
116                       arUploaded(UBound(arUploaded)).goodsize = FileLen(arImg(nIdx).local)
117                   End If
                      
118                   frmProgress.shapBar.Width = frmProgress.shapBar.Width + nProgressStep
119               Next
120           End If
              
              'transfer done, we now apply the required visibility (public or private preview)
121           frmProgress.lblStage.Caption = arMsg(47) & " " & strTo
122           frmProgress.lblDetail.Caption = strTempName & " -> " & strKey
              
123           Select Case index
                  Case 0: 'private preview
124                   strFinalName = structProj.preview_prefix & "-" & strKey
125                   bRet = FtpCli.Met_RENAME(strTempName, strFinalName)
126               Case 1: 'public article
127                   strFinalName = strKey
128                   bRet = FtpCli.Met_RENAME(strTempName, strFinalName)
129               Case 2: 'rapatriate article copy
                      '***TODO : see ***LATER ahead
130               Case 3: 'rapatriate full backup
                      '***TODO : see ***LATER ahead
131               Case Else
132                   MsgBox "Bad 'index/1' in cmdUploadArt/Click call", vbCritical, "***DEBUG"
133           End Select
              
134           frmProgress.shapBar.Width = frmProgress.shapBar.Width + nProgressStep
135           frmProgress.timerClose.Enabled = True
              
136           If bRet = False Then
                  'FAILED (publication aborted)
137               strErr = arMsg(48) & " " & strTo & " !"
138               Err.Raise 513
139           Else
                  'valides the remote article/images, then deletes the local one if required
                  Dim nRequiredRemoteSize As Long   'ASCII transfer => decrease of one byte at every CRLF
                  Dim nCurrRemoteSize As Long
                  Dim bCorruption As Boolean
                  
                  'validation
140               arUploaded(1).file = strFinalName  'updates the already uploaded files list
141               FtpCli.Met_DIR
                  
142               bCorruption = False
143               For nIdx2 = 1 To UBound(arUploaded)
144                   nRequiredRemoteSize = arUploaded(nIdx2).goodsize
145                   nCurrRemoteSize = FtpCli.RemoteFiles("K" & arUploaded(nIdx2).file).FileSize
146                   If nCurrRemoteSize = nRequiredRemoteSize Then
147                       arUploaded(nIdx2).integrity = 1
148                   Else
149                       arUploaded(nIdx2).integrity = -1
150                       bCorruption = True
151                   End If
152               Next
                  
153               If bCorruption = False Then
                      '(no academic, but allows user to see the 100% on gauge)
154                   Do While IsLoaded("frmProgress") = True
155                       DoEvents
156                   Loop
                      
                      'SUCCES (publication correctly done)
                      '***TODO : maybe gather cases 0 and 1 (and/or think about eventual differences)
157                   strMsg = arMsg(49) & " ! "
158                   Select Case index
                          Case 0: 'preview OK : archives w/ deletion in workspace, then tell to user how to see online
159                           strIdArchiv = BuildIdArchive(strFinalName)
160                           bRet = SetGetArchive(strKey, strIdArchiv, True, nAddedNameLen, strServerAttachs, False)
161                           If bRet = False Then
162                               strErr = arMsg(50) & " !"
163                               Err.Raise 513     '***LATER : see if error 513 is the most pertinent
164                           End If
                              
165                           PopulateClearLocalList True
166                           PopulateRemoteList
                              
167                           strMsg = arMsg(51) & " '" & strKey & "'" _
                                          & " " & arMsg(52) & "." & vbNewLine _
                                          & vbNewLine _
                                          & "- " & arMsg(53) & "." & vbNewLine _
                                          & "- " & arMsg(54) & "." & vbNewLine _
                                          & "- " & arMsg(55) & "."
                          
168                           MsgBoxEx strMsg, vbInformation
169                       Case 1: 'public OK : archives w/ deletion in workspace, then tell to user how to see online
170                           strIdArchiv = BuildIdArchive(strFinalName)
171                           bRet = SetGetArchive(strKey, strIdArchiv, True, nAddedNameLen, strServerAttachs, False)
172                           If bRet = False Then
173                               strErr = arMsg(50) & " !"
174                               Err.Raise 513     '***LATER : see if error 513 is the most pertinent
175                           End If
                              
176                           PopulateClearLocalList True
177                           PopulateRemoteList
                              
178                           strMsg = arMsg(56) & " '" & strKey & "' " & arMsg(57) & "." & vbNewLine _
                                          & vbNewLine _
                                          & arMsg(53) & vbNewLine _
                                          & arMsg(58) & "."
                          
179                           MsgBoxEx strMsg, vbInformation
180                       Case 2: 'copy of the article rapatriated locally
                              '***TODO : see ***LATER ahead
181                           MsgBox arMsg(92)
182                       Case 3: 'full backup done
                              '***TODO : see ***LATER ahead
183                           MsgBox arMsg(92)
184                       Case Else
185                           MsgBox "Bad 'index/2' in cmdUploadArt/Click call", vbCritical, "***DEBUG"
186                   End Select
187               Else
                      'FAILED (publication aborted)
188                   strErr = arMsg(59) & " !"
189                   Err.Raise 513
190               End If
191           End If
192       End If
          
UploadArtEnd:
193       If IsLoaded("frmProgress") = True Then Unload frmProgress
194       Screen.MousePointer = vbDefault
195       On Error GoTo 0
196       Exit Sub
          
UploadArtErr:
          'error handling
          '(all errors being fatal for the publication, we try to delete the already transfered files before all)
          
          'WARNING : Met_DELETE does not always generate a return 250 as awaited by FtpCli, so I prefer
          '          to check the good deletion using Ftp_SIZE (which *must* fail)
197       For nIdx = UBound(arUploaded) To 1 Step -1
198           FtpCli.Met_DELETE arUploaded(nIdx).file
199           If FtpCli.Ftp_SIZE(arUploaded(nIdx).file) = False Then
                  'deletes the element in the structures array
200               For nIdx2 = nIdx + 1 To UBound(arUploaded)
201                   arUploaded(nIdx2 - 1) = arUploaded(nIdx2)
202               Next
203               If UBound(arUploaded) > 1 Then
204                   ReDim Preserve arUploaded(LBound(arUploaded) To UBound(arUploaded) - 1)
205               End If
206           End If
207       Next
          
208       Screen.MousePointer = vbDefault
          
209       Select Case Err.Number
              Case 513
210               If strErr <> "" Then strMsg = strErr & vbNewLine & vbNewLine
              
                  'tests if the array is not empty
211               If CBool(Not arUploaded) = False Then
212                   If UBound(arUploaded) > 1 Then
213                       strMsg = strMsg & "Des fichiers orphelins n'ont pu être supprimés du serveur :"
214                       For nIdx = 1 To UBound(arUploaded)
215                           strMsg = strMsg & vbNewLine & "- " & arUploaded(nIdx).file & " ("
216                           Select Case arUploaded(nIdx).integrity
                                  Case -1: strMsg = strMsg & "-"
217                               Case 1: strMsg = strMsg & "+"
218                               Case Else: strMsg = strMsg & "?"  '0
219                           End Select
220                       Next
221                       strMsg = strMsg & ")" & vbNewLine & vbNewLine & _
                                      "Vous devrez effacer ces fichiers par vous-même (ou 'Nettoyer') "
222                   Else ' UBound(arUploaded) = 1
223                       strMsg = strMsg & "Un fichier orphelin n'a pu être supprimé du serveur :"
224                       strMsg = strMsg & vbNewLine & "- " & arUploaded(1).file & " ("
225                       Select Case arUploaded(nIdx).integrity
                              Case -1: strMsg = strMsg & "-"
226                           Case 1: strMsg = strMsg & "+"
227                           Case Else: strMsg = strMsg & "?"  '0
228                       End Select
229                       strMsg = strMsg & ")" & vbNewLine & vbNewLine & _
                                      "Vous devrez effacer ce fichier par vous-même (ou 'Nettoyer') "
230                   End If
          
231                   strMsg = strMsg & vbNewLine & "ou réessayer de publier ce même article" & _
                               vbNewLine & "pour tenter de résoudre le problème."
                               
232                   strMsg = strMsg & vbNewLine & vbNewLine & "__" & vbNewLine & _
                              "NB : l'état de chaque fichier sur serveur est indiqué entre parenthèse :" & vbNewLine & _
                              "     '+' signifie que le fichier semble intègre (arrivé entier)" & vbNewLine & _
                              "     '-' signifie que le fichier semble corrompu (incomplet)" & vbNewLine & _
                              "     '?' signifie que son état est indéterminé (non vérifié)"
233               Else
234                   strMsg = strMsg & "Fermez BlosHome et rouvrez-le pour réessayer..."
235               End If
                  
236               Screen.MousePointer = vbDefault
237               MsgBox strMsg, vbExclamation
238           Case Else
239               strMsg = "Error #" & Err.Number & "@ bloshome/frmMain/cmdUploadArt/#" & Erl & " : " & Err.Description
240               strMsg = strMsg & vbNewLine & "LastDLLError : " & Err.LastDllError
241               strMsg = strMsg & " / GetLastError : " & GetLastError()
242               MsgBox strMsg, vbExclamation
243       End Select
          
244       PopulateRemoteList
245       PopulateClearLocalList True
          
246       Resume UploadArtEnd
End Sub

Private Sub cmdRenLocalArt_Click()
    'Renames a local article
    Dim strKey As String
    Dim strArt As String
    
    Dim strOldArt As String
    Dim strNewArt As String
    
    Dim strTitle As String
    Dim strMsg As String
    Dim strDefault  As String
    
    Dim bOK As Boolean
    
    'basic checking
    If lstLocal.ListItems.Count = 0 Then Exit Sub
    
    If lstLocal.SelectedItem.Selected = False Then
        MsgBox arMsg(71) & " !", vbExclamation
        Exit Sub
    End If
    
    strKey = lstLocal.SelectedItem.Key
    strOldArt = GetWorkspace() & "\" & strKey
    
    If Dir(strCurrArt) = "" Then
        MsgBox arMsg(72) & " !", vbExclamation
        PopulateClearLocalList True
        Exit Sub
    End If
    
    'user input
    strTitle = arMsg(60)
    strMsg = arMsg(61) & " '" & strKey & "'" & vbNewLine _
                & vbNewLine _
                & arMsg(7) & " : " & vbNewLine _
                & "- " & arMsg(8) & " /*\<>|?" & Chr(34) & ":" & vbNewLine _
                & "- " & arMsg(9) & vbNewLine _
                & "- " & arMsg(10) & vbNewLine _
                & "- " & arMsg(11) & vbNewLine _
                & "- " & arMsg(12) & vbNewLine _
                & "- " & arMsg(13) & " '" & structProj.art_ext & "' " & arMsg(14) & vbNewLine _
                & vbNewLine & arMsg(15) & structProj.art_ext & "'."

    strDefault = GetFilenamePrefix(strKey)
    bOK = False
    
    Do
        strArt = InputBox(strMsg, strTitle, strDefault)
        If strArt = "" Then Exit Sub
        strDefault = strArt 'saves for eventual next loop
        strArt = CheckAndMkFilename(UnAccent(LCase(strArt)), False, True)
        
        If strArt <> "" Then
            If IsNumeric(Left$(strArt, 1)) = False Then
                bOK = True
            Else
                MsgBox arMsg(73) & " !", vbExclamation
                bOK = False
            End If
        Else
            MsgBox arMsg(74) & " !", vbExclamation
            bOK = False
        End If
        
        If Right$(strArt, 4) <> structProj.art_ext Then
            strArt = strArt & structProj.art_ext
        End If
        
        strNewArt = GetWorkspace() & "\" & strArt
        
        If Dir(strNewArt) <> "" Then
            MsgBox arMsg(75) & " !", vbExclamation
            bOK = False
        End If
    Loop Until bOK = True
    
    'effective application
    Screen.MousePointer = vbHourglass
    On Error GoTo cmdRenLocalArt_Click_Error
    Name strOldArt As strNewArt
    Screen.MousePointer = vbDefault
    
    PopulateClearLocalList True

cmdRenLocalArt_Click_End:
   On Error GoTo 0
   Exit Sub

cmdRenLocalArt_Click_Error:
    MsgBox arMsg(93) & " '" & strKey & "' !", vbExclamation
    Resume cmdRenLocalArt_Click_End
End Sub

Private Sub cmdRenRemoteArt_Click()
    'Renames a remote article (takes care of eventual remote/local timestamp cache/index files)
    '***TODO : function to impemented considering these points :
    '- echo the update in cache file (local and distant, according to projet settings)
    '- updates name in local archive too
    '- see OptStatus_Click
    MsgBox arMsg(70) & "...", vbInformation
End Sub

Private Sub cmdRenRemoteCat_Click()
    'Renames a category (takes care of eventual remote/local timestamp cache/index files)
    '***TODO : function to impemented considering these points :
    '- echo the update in cache file (local and distant, according to projet settings)
    '- updates category/path name in local archive too
    MsgBox arMsg(70) & "...", vbInformation
End Sub

Private Sub cmdRmRemoteCat_Click()
    'Deletes the selected category (if empty only)
    Dim strKey As String
    Dim nRet As Integer
    Dim nCount As Integer
    
    If lstRemote.SelectedItem.Selected = False Then
        MsgBox arMsg(94) & " !", vbExclamation
        Exit Sub
    End If
    
    If lstRemote.SelectedItem.SmallIcon <> "Category" Then
        MsgBox arMsg(95) & " !", vbInformation
        Exit Sub
    End If
    
    strKey = lstRemote.SelectedItem.Key
    
    'vérification de sécurité
    Screen.MousePointer = vbHourglass
    FtpCli.Met_CD FtpCli.RemoteFiles(strKey).FileName
    FtpCli.Met_DIR
    nCount = FtpCli.RemoteFiles.Count
    FtpCli.Met_CDUP
    FtpCli.Met_DIR
    Screen.MousePointer = vbDefault
    If nCount > 0 Then
        MsgBox arMsg(96) & " !" & vbNewLine & arMsg(97) & ".", vbInformation
        Exit Sub
    End If
    
    'effective deletion
    If strKey <> ".." Then
        FtpCli.Met_DIR
        If FtpCli.RemoteFiles(strKey).FileType = 0 Then 'directory
            nRet = MsgBox(arMsg(98) & " '" & lstRemote.SelectedItem.Text & "' ?", vbQuestion + vbYesNo)
            If nRet = vbYes Then
                Screen.MousePointer = vbHourglass
                FtpCli.Ftp_RMD FtpCli.RemoteFiles(strKey).FileName
                PopulateRemoteList
                Screen.MousePointer = vbDefault
            End If
            Exit Sub
        End If
    End If
    
    MsgBox arMsg(94) & " !", vbExclamation
End Sub

Private Sub cmdSeeRemoteArt_Click()
    'Navigates to the selected remote article using the default Web browser in Windows
    Dim bInvalid As Boolean
    Dim strMsg As String
    Dim nRet As Integer
    
    Dim strType As String
    Dim strArt As String
    Dim strURL As String
    
    If lstRemote.SelectedItem.Selected = False Then
        MsgBox arMsg(71) & " !", vbExclamation
        Exit Sub
    End If
    
    strType = lstRemote.SelectedItem.SmallIcon
    
    If strType <> "Article" And strType <> "Preview" Then
        MsgBox arMsg(78) & " !", vbInformation
        Exit Sub
    End If
    
    If structProj.blosxom_url = "" Then bInvalid = True
    If Left$(structProj.blosxom_url, 7) <> "http://" _
        Or LCase(Right$(structProj.blosxom_url, 11)) <> "blosxom.cgi" Then bInvalid = True
        
    If bInvalid = True Then
        strMsg = arMsg(62)
        nRet = MsgBox(strMsg, vbQuestion + vbYesNo)
        If nRet = vbYes Then
            frmProject.bNoCancel = True
            frmProject.Show 1
            ApplyCleanProject True
        Else
            Exit Sub
        End If
    End If
    
    strArt = FtpCli.RemoteFiles(lstRemote.SelectedItem.Key).FileName
    strArt = Left$(strArt, Len(strArt) - Len(structProj.art_ext)) & structProj.flav_ext
    
    strURL = structProj.blosxom_url & "/" & GetCurrRemoteCat()
    If Right$(strURL, 1) <> "/" Then strURL = strURL & "/" 'fixes bug when blog root
    strURL = strURL & strArt
    
    If strType = "Preview" Then
        strURL = strURL & "?preview=" & structProj.preview_pass
    End If
    
    OpenMIME strURL
End Sub

Private Sub Form_Load()
    'Allows one instance of BlosHome only
    If HowManyRunning("bloshome") > 1 Then
        MsgBox arMsg(99) & " !", vbExclamation
        bRightLoading = False
        Exit Sub
    End If
    
    'Assocates the multi-resolution icon resource to the application
    SetIcon Me.hwnd, "AAA", True
    
    'Loads the configuration of the application and starts a log file for the session
    CheckMandatory
    bRightLoading = LoadAllSettings()
    BeginLog
    
    'Adds an hidden column on lstRemote to be able to order with categories ahead at PopulateRemoteList() time
    '***LATER : find better than width=0 to hide the column
    frmMain.lstRemote.ColumnHeaders.Add 7, , "Sort", 0 'one based
    
    'Preloads the editor (the TinyMCE initilization being slow, it will be faster to show/hide rather than load/unload)
    Load frmEdit
    
    'Some UI-strings adjustements afterward
    cmdUploadArt(0).Caption = cmdUploadArt(0).Caption & " >"
    cmdUploadArt(1).Caption = cmdUploadArt(1).Caption & " >"

    cmdDownloadArt(0).Caption = "< " & cmdDownloadArt(0).Caption
    cmdDownloadArt(1).Caption = "< " & cmdDownloadArt(1).Caption
    
    cmdBackupBlog.Caption = cmdBackupBlog.Caption & vbNewLine & "<<<<<"
    
    lblLocalArts.Caption = lblLocalArts.Caption & " :"
    lblRemoteCats.Caption = lblRemoteCats.Caption & " :"
    lblRemoteArts.Caption = lblRemoteArts.Caption & " :"
End Sub

Private Sub Form_Activate()
    'Will terminate the application if something wrong during form load (e.g. problem w/ required lang file)
    If bRightLoading = False Then
        Unload Me
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Forces a clean disconnection if necessary
    If Left$(UCase(cmdCnnx.Caption), 1) = "D" Then FtpCli.Met_BYE
    
    'Saves the configuration and closes the log (one per session) before to exit
    SaveAllSettings
    EndLog
    
    'Unloads the editor (still in RAM by design)
    Unload frmEdit
End Sub

Private Sub FtpCli_ProtocolArrival(data As String)
    'Logs the incoming FTP infos
    DoLog "<-" & data
End Sub

Private Sub FtpCli_ProtocolSend(data As String)
    'Logs the outcoming FTP infos
    DoLog "->" & data
End Sub

Private Sub lstLocal_DblClick()
    'Same as cmdEditLocalArt
    cmdEditLocalArt_Click
End Sub

Private Sub lstRemote_Click()
    'Shows the status and pertinente commands for the selected entry in lstRemote
    UpdateRemoteContextCmds
    ShowSelectedArtStatus
End Sub

Private Sub lstRemote_DblClick()
    'Navigates to the selected category (i.e. directory change) or reachs the selected article online
    Dim strKey As String
    
    Screen.MousePointer = vbHourglass
    If lstRemote.SelectedItem.Selected = False Then Exit Sub
    
    strKey = lstRemote.SelectedItem.Key
    
    If strKey = ".." Then
        FtpCli.Met_CDUP
        PopulateRemoteList
    Else
        If FtpCli.RemoteFiles(strKey).FileType = 0 Then 'répertoire
            FtpCli.Met_CD FtpCli.RemoteFiles(strKey).FileName
            PopulateRemoteList
        Else
            cmdSeeRemoteArt_Click
        End If
    End If

    Screen.MousePointer = vbDefault
End Sub

Private Sub lstRemote_KeyUp(KeyCode As Integer, Shift As Integer)
    'Same as lstRemote_Click
    UpdateRemoteContextCmds
    ShowSelectedArtStatus
End Sub

Private Sub mnuInfo_About_Click()
    'Displays the about-box (in modal mode)
    frmAbout.Show 1, Me
End Sub

Private Sub mnuInfo_Help_Click()
    'Try to open the PDF help matching the current language using the default Windows software (e.g. Adobe Reader)
    Dim strHelpFile As String
    
    strHelpFile = App.Path & "\bloshome_" & LCase(GetLang()) & ".pdf"
    If Dir(strHelpFile) = "" Then
        MsgBox arMsg(100) & " : '" & strHelpFile & "'", vbOKOnly + vbExclamation
        Exit Sub
    End If
    OpenMIME strHelpFile
End Sub

Private Sub mnuLang_EN_Click()
    'Activates English language file
    '***LATER : implement a way to manage menu dynamically from list of languages files
    '          (i.e. improve from current bilingual state to a multilingual mode)
    SetLang "EN", "main"
End Sub

Private Sub mnuLang_FR_Click()
    'Activates French language file
    '***LATER : see mnuLang_EN_Click
    SetLang "FR", "main"
End Sub

Private Sub mnuProj_Del_Click()
    'Displays the dialog-box for project deletion
    '(will manage to ask for closing of current project if the one to delete is the current one)
    bOpenDelProj = False
    frmSelect.Show 1, Me
End Sub

Private Sub mnuProj_New_Click()
    'Displays the dialog-box for settings of a new project
    '(will ask for closing of current project before to define a new one)
    If IsFreeSDI("créer un nouveau projet.") = True Then
        frmProject.bNoCancel = False
        frmProject.Show 1, Me
    End If
End Sub

Private Sub mnuProj_Open_Click()
    'Displays the dialog-box for existing project opening
    '(will ask for closing of current project before to open a new one ; this since BlosHome is SDI designed)
    bOpenDelProj = True
    frmSelect.Show 1, Me
End Sub

Private Sub optStatusArt_Click(index As Integer)
    'Switches the status of the selected article (private preview or public article)
    Dim strCurrName As String
    Dim strNewName As String
    Dim bRet As Boolean
    Dim strMsg As String
    Dim nIcon As Integer
    
    If structProj.with_preview = False Then Exit Sub 'preview feature must be enabled in current project settings
    If bMouseoverStatusOpt = False Then Exit Sub
    If lstRemote.SelectedItem.Selected = False Then Exit Sub
    If optStatusArt(index).Value = False Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    strCurrName = FtpCli.RemoteFiles(lstRemote.SelectedItem.Key).FileName
    nIcon = vbInformation
    
    If index = 0 Then
        'public -> private
        strNewName = structProj.preview_prefix & "-" & strCurrName
        strMsg = "'" & strCurrName & "' " & arMsg(63)
    Else
        'private -> public
        strNewName = Replace(strCurrName, structProj.preview_prefix & "-", "", , 1, vbTextCompare)
        strMsg = "'" & strNewName & "' " & arMsg(64)
        strMsg = strMsg & vbNewLine & vbNewLine & arMsg(65)
    End If
    
    bRet = FtpCli.Met_RENAME(strCurrName, strNewName)
    PopulateRemoteList
    
    Screen.MousePointer = vbDefault
    If bRet = False Then
        strMsg = arMsg(66) & " !"
        nIcon = vbExclamation
    End If
    
    MsgBox strMsg
End Sub

Private Sub optStatusArt_GotFocus(index As Integer)
    bMouseoverStatusOpt = True
End Sub

Private Sub optStatusArt_LostFocus(index As Integer)
    bMouseoverStatusOpt = False
End Sub

Private Sub timerCnnx_Timer()
    'Periodical process about connection status
    '***LATER : if BlosHome should become multilingual (not bilingual only), this fucntion should be reviewed to detect
    '           the connection status whetever be the word in cmdCnnx button (currently OK for FR|EN only)
    
    nTimerPass = nTimerPass + 1
    
    If UCase(Left$(cmdCnnx.Caption, 1)) = "C" Then
        'DISCONNECTED (button displays "Conne[x[ct]ion" ; x|ct being the difference between FR|EN)
        'NB : timer started on cmdCnnx_Click(), then modified if connection succeeded or stopped if failure
        
        'Informs user about countdown before timeout during attempt of connection
        '(from 1st pass, even before the end of 1st interval)
        Dim nRest As Integer 'tps restant avant timeout en secondes
        
        nRest = FtpCli.TimeOutDelay - (nTimerPass * timerCnnx.Interval / 1000)
        barStatus.Panels.item("cnnx").Text = arMsg(67) & "... (" & Trim$(Str(nRest)) & ")"
    Else ' "D"
        'CONNECTED (button displays other thing)
        'NB : timer started on proved connection and stopped on disconnection
        
        'Sends a command adapted to maintain the connection (unless rejected or ignored by server)
        'and informs user about time of connection in status-bar (from 2nd pass, at the end of 1st interval)
        'NB : Ftp_STAT has been MODded for a good/better support of server return
        Dim nAlea As Integer
        Dim nMn As Integer
        Dim strUnit As String
        Dim bStillCnnx As Boolean
        
        If nTimerPass = 1 Then
            Randomize Timer
            Exit Sub
        End If
        
        nAlea = Int(Rnd * 3) + 1 'in [1, 6]
        
        Select Case nAlea
            Case 1
                bStillCnnx = FtpCli.Ftp_NOOP()
            Case 2
                bStillCnnx = FtpCli.FTP_STAT()
            Case Else '3
                bStillCnnx = FtpCli.Met_DIR()
        End Select
        
        If bStillCnnx = True Then
            'tells about connection duration
            nMn = (nTimerPass - 1) * Int((timerCnnx.Interval / 1000 / 60) + 0.5) 'rounds to lower minute
            strUnit = arMsg(68)
            If nMn > 1 Then strUnit = strUnit & "s"
            barStatus.Panels.item("cnnx").Text = arMsg(69) & " " & nMn & " " & strUnit
        Else
            'tells about observed disconnection
            ApplyCnnx False
        End If
    End If
    
    DoEvents
End Sub
