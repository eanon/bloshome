VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{E9D07A90-4BDE-46EA-BFB2-8FC4989DA45B}#2.2#0"; "DiFtpCli6.ocx"
Begin VB.Form frmTest 
   Caption         =   "Test DiFtpClient"
   ClientHeight    =   5415
   ClientLeft      =   1605
   ClientTop       =   2835
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   5415
   ScaleWidth      =   11880
   Begin VB.PictureBox pctCommand 
      Height          =   495
      Left            =   1920
      ScaleHeight     =   435
      ScaleWidth      =   8595
      TabIndex        =   17
      Top             =   4440
      Visible         =   0   'False
      Width           =   8655
      Begin VB.CommandButton cmdCommand 
         Caption         =   "Send"
         Height          =   255
         Left            =   6840
         TabIndex        =   20
         Top             =   120
         Width           =   735
      End
      Begin VB.TextBox txtCommand 
         Height          =   285
         Left            =   1560
         TabIndex        =   19
         Top             =   120
         Width           =   5175
      End
      Begin VB.Label Label2 
         Caption         =   "Commande libre"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   120
         Width           =   1335
      End
   End
   Begin DiFtpCli6.FtpCli FtpCli1 
      Left            =   480
      Top             =   4560
      _extentx        =   873
      _extenty        =   873
      remotehost      =   "Localhost"
      username        =   "anonymous"
      password        =   "jean-luc@"
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1080
      Top             =   4440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":0000
            Key             =   "Folder"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":0458
            Key             =   "File"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":08B0
            Key             =   "Link"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":0D04
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":1158
            Key             =   "Up"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdRecv 
      Caption         =   "<=="
      Height          =   375
      Left            =   4200
      TabIndex        =   10
      Top             =   2880
      Width           =   735
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "==>"
      Height          =   375
      Left            =   4200
      TabIndex        =   9
      Top             =   2280
      Width           =   735
   End
   Begin VB.Frame Frame2 
      Caption         =   "Remote system"
      Height          =   4095
      Left            =   5040
      TabIndex        =   5
      Top             =   240
      Width           =   6375
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Refresh"
         Height          =   375
         Left            =   5400
         TabIndex        =   15
         Top             =   3600
         Width           =   855
      End
      Begin VB.CommandButton cmdDele 
         Caption         =   "Delete"
         Height          =   375
         Left            =   5400
         TabIndex        =   14
         Top             =   2040
         Width           =   855
      End
      Begin VB.CommandButton cmdRen 
         Caption         =   "Rename"
         Height          =   375
         Left            =   5400
         TabIndex        =   13
         Top             =   1560
         Width           =   855
      End
      Begin VB.CommandButton cmdMkDir 
         Caption         =   "Mk Dir"
         Height          =   375
         Left            =   5400
         TabIndex        =   12
         Top             =   1080
         Width           =   855
      End
      Begin VB.CommandButton cmdChDir 
         Caption         =   "Chg Dir"
         Height          =   375
         Left            =   5400
         TabIndex        =   11
         Top             =   600
         Width           =   855
      End
      Begin MSComctlLib.ListView lstDir 
         Height          =   3375
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   5953
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "FileName"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Time"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Size"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Description"
            Object.Width           =   5292
         EndProperty
      End
      Begin VB.Label lblRemoteDir 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   960
         TabIndex        =   7
         Top             =   240
         Width           =   5175
      End
      Begin VB.Label Label1 
         Caption         =   "current dir"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Local system"
      Height          =   4215
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3975
      Begin VB.FileListBox File1 
         Height          =   2040
         Left            =   120
         MultiSelect     =   2  'Extended
         TabIndex        =   4
         Top             =   1920
         Width           =   3735
      End
      Begin VB.DirListBox Dir1 
         Height          =   1215
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   3735
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   3735
      End
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   5040
      Width           =   1455
   End
   Begin VB.Label lblRapport 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   1920
      TabIndex        =   16
      Top             =   5040
      Width           =   8175
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdChDir_Click()
  Dim SelKey As String
  Dim FileName As String
  If lstDir.SelectedItem Is Nothing Then
    Exit Sub
  End If
  SelKey = lstDir.SelectedItem.Key
  If SelKey = "." Then
    FtpCli1.Met_CDUP
    FillDirList
    Exit Sub
  End If
  Select Case FtpCli1.RemoteFiles(SelKey).FileType
  Case 0
    'directory
    FtpCli1.Met_CD FtpCli1.RemoteFiles(SelKey).FileName
    FtpCli1.Met_DIR
    FillDirList
    lblRapport = ""
  Case 1
    'file
    'nop
  Case Else
    'nop
  End Select
End Sub

Private Sub cmdCommand_Click()
  FtpCli1.Met_QUOTE txtCommand
End Sub

Private Sub cmdConnect_Click()
  If cmdConnect.Caption = "Connect" Then
    frmConnect.Show 1
    FtpCli1.RemoteHost = frmConnect.txtHost
    FtpCli1.RemotePort = 21
    FtpCli1.UserName = frmConnect.txtUser
    FtpCli1.Password = frmConnect.txtPass
    If FtpCli1.Met_CONNECT Then
      cmdSend.Enabled = True
      cmdRecv.Enabled = True
      cmdMkDir.Enabled = True
      cmdChDir.Enabled = True
      cmdRen.Enabled = True
      cmdDele.Enabled = True
      cmdRefresh.Enabled = True
      pctCommand.Visible = True
      lstDir.Enabled = True
      cmdConnect.Caption = "Disconnect"
      lblRemoteDir = FtpCli1.RemoteDir
      FillDirList
    End If
  Else
    If FtpCli1.Met_BYE Then
      cmdSend.Enabled = False
      cmdRecv.Enabled = False
      cmdMkDir.Enabled = False
      cmdChDir.Enabled = False
      cmdRen.Enabled = False
      cmdDele.Enabled = False
      cmdRefresh.Enabled = False
      cmdConnect.Caption = "Connect"
      pctCommand.Visible = False
      lstDir.ListItems.Clear
    End If
  End If
End Sub

Private Sub cmdDele_Click()
  Dim iPnt As Integer
  Dim SelKey As String
  Dim FileName As String
  For iPnt = 1 To lstDir.ListItems.Count
    If lstDir.ListItems(iPnt).Selected Then
      SelKey = lstDir.ListItems(iPnt).Key
      If SelKey <> "." Then 'the up
        Select Case FtpCli1.RemoteFiles(SelKey).FileType
        Case 0
          'directory
          FtpCli1.Ftp_RMD FtpCli1.RemoteFiles(SelKey).FileName
        Case 1
          'file
          FtpCli1.Met_DELETE FtpCli1.RemoteFiles(SelKey).FileName
        Case Else
          'nop
        End Select
      End If
    End If
  Next
  FillDirList
End Sub

Private Sub cmdMkDir_Click()
  Dim RemoteDir As String
  RemoteDir = InputBox("Name for the new directory", "MkDir")
  If RemoteDir <> "" Then
    If FtpCli1.Ftp_MKD(RemoteDir) <> True Then
      MsgBox "Make directory fail"
    Else
      FillDirList
    End If
  End If
End Sub

Private Sub cmdRecv_Click()
  Dim iPnt As Integer
  Dim FileName As String
  Dim FileType As Integer
  Dim SelKey As String
  For iPnt = 1 To lstDir.ListItems.Count
    If lstDir.ListItems(iPnt).Selected Then
      SelKey = lstDir.SelectedItem.Key
      FileName = FtpCli1.RemoteFiles(SelKey).FileName
      FileType = FtpCli1.RemoteFiles(SelKey).FileType
      If FileType = FtpFile Then
        FtpCli1.Met_GET FileName, Dir1.Path & "\" & FileName
      'to do : if filetype=ftpdir
      End If
    End If
  Next
  File1.Refresh
  FillDirList
End Sub

Private Sub cmdRefresh_Click()
  FillDirList
End Sub

Private Sub cmdRen_Click()
  Dim FrFileName As String
  Dim ToFileName As String
  Dim SelKey As String
  If lstDir.SelectedItem Is Nothing Then
    Exit Sub
  End If
  SelKey = lstDir.SelectedItem.Key
  FrFileName = FtpCli1.RemoteFiles(SelKey).FileName
  ToFileName = InputBox("New name for " & FrFileName, "Rename file")
  If ToFileName <> "" Then
    FtpCli1.Met_RENAME FrFileName, ToFileName
    FillDirList
  End If
End Sub

Private Sub cmdSend_Click()
  Dim iPnt As Integer
  For iPnt = 0 To File1.ListCount - 1
    If File1.Selected(iPnt) Then
      FtpCli1.Met_PUT Dir1.Path & "\" & File1.List(iPnt)
    End If
  Next
  FillDirList
End Sub

Private Sub Drive1_Change()
' Définit le chemin d'accès du répertoire.
  Dir1.Path = Drive1.Drive
End Sub

Private Sub Dir1_Change()
' Définit le chemin d'accès du fichier.
  File1.Path = Dir1.Path
End Sub




Private Sub File1_DblClick()
  Dim iPnt As Integer
  iPnt = File1.ListIndex
  If iPnt < 0 Then
    Exit Sub
  End If
  FtpCli1.Met_PUT Dir1.Path & "\" & File1.List(iPnt)
  FillDirList
End Sub

Private Sub Form_Load()
  cmdSend.Enabled = False
  cmdRecv.Enabled = False
  cmdMkDir.Enabled = False
  cmdChDir.Enabled = False
  cmdRen.Enabled = False
  cmdDele.Enabled = False
  cmdRefresh.Enabled = False
  cmdConnect.Caption = "Connect"
  On Error Resume Next
  Dir1.Path = App.Path & "\Temp"
  frmDialog.Show , Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Unload frmDialog
End Sub

Private Sub FtpCli1_ProtocolArrival(Data As String)
  frmDialog.List1.AddItem "<=" & Data
  frmDialog.List1.ListIndex = frmDialog.List1.ListCount - 1
End Sub

Private Sub FtpCli1_ProtocolSend(Data As String)
  frmDialog.List1.AddItem "=>" & Data
End Sub





Private Sub FtpCli1_RecieveProgress(Size As Long)
  lblRapport = CStr(Size)
End Sub

Private Sub FtpCli1_SendProgress(Pcent As Integer, Size As Long)
   lblRapport = CStr(Size) & "  " & CStr(Pcent) & "%"
End Sub

Private Sub lstDir_DblClick()
  Dim SelKey As String
  Dim FileName As String
  If lstDir.SelectedItem Is Nothing Then
    Exit Sub
  End If
  SelKey = lstDir.SelectedItem.Key
  If SelKey = "." Then
    FtpCli1.Met_CDUP
    FillDirList
    Exit Sub
  End If
  Select Case FtpCli1.RemoteFiles(SelKey).FileType
  Case 0
    'directory
    FtpCli1.Met_CD FtpCli1.RemoteFiles(SelKey).FileName
    FillDirList
  Case 1
    'file
    FtpCli1.Met_GET FtpCli1.RemoteFiles(SelKey).FileName, Dir1.Path & "\" & FtpCli1.RemoteFiles(SelKey).FileName
    FillDirList
  Case Else
    'nop
  End Select
End Sub

Private Sub FillDirList()
  Dim iPnt As Integer
  Dim jPnt As Integer
  Dim FileItems As Variant
  Dim iTem As ListItem
  Dim ImgName As String
  Dim FileName As String
  Dim RemoteFile As RemoteFile
  lstDir.ListItems.Clear
  Set iTem = lstDir.ListItems.Add(, ".", ".", "Up", "Up")
  If FtpCli1.Met_DIR Then
    For Each RemoteFile In FtpCli1.RemoteFiles
      Select Case RemoteFile.FileType
      Case 0
        ImgName = "Folder"
      Case 1
        ImgName = "File"
      Case 2
        ImgName = "Link"
      Case Else
        ImgName = "Unknwon"
      End Select
      On Error Resume Next
      Set iTem = lstDir.ListItems.Add(, RemoteFile.Key, RemoteFile.FileName, ImgName, ImgName)
      iTem.SubItems(1) = Format(RemoteFile.FileDate, "dd/mm/yyyy")
      iTem.SubItems(2) = Format(RemoteFile.FileDate, "hh:nn")
      iTem.SubItems(3) = RemoteFile.FileSize
      iTem.SubItems(4) = RemoteFile.Description
    Next
  End If
  lblRemoteDir = FtpCli1.RemoteDir
  lblRapport = ""
End Sub

