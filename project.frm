VERSION 5.00
Begin VB.Form frmProject 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Parameters"
   ClientHeight    =   8400
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   10185
   ControlBox      =   0   'False
   LinkTopic       =   "project"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8400
   ScaleWidth      =   10185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picIcon 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   8880
      Negotiate       =   -1  'True
      Picture         =   "project.frx":0000
      ScaleHeight     =   337.12
      ScaleMode       =   0  'User
      ScaleWidth      =   337.12
      TabIndex        =   48
      Top             =   7200
      Width           =   510
   End
   Begin VB.Frame framChapo 
      Caption         =   "Lede"
      Height          =   1695
      Left            =   4305
      TabIndex        =   41
      Top             =   6570
      Width           =   3915
      Begin VB.CheckBox chkChapo 
         Caption         =   "possibility to indicate a lede limit"
         Height          =   240
         Left            =   300
         TabIndex        =   42
         Top             =   345
         Width           =   3435
      End
      Begin VB.TextBox txtChapo 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1410
         TabIndex        =   45
         Text            =   "more"
         Top             =   675
         Width           =   720
      End
      Begin VB.Label lblChapoTagEnding 
         Caption         =   " -->"
         Height          =   195
         Left            =   2115
         TabIndex        =   46
         Top             =   705
         Width           =   240
      End
      Begin VB.Label lblChapoTagStart 
         Caption         =   "<!-- "
         Height          =   195
         Left            =   1140
         TabIndex        =   44
         Top             =   705
         Width           =   255
      End
      Begin VB.Label lblChapoPlugin 
         Caption         =   "Requires the SeeMore plugin by Todd Larason (or equivalent) on server-side."
         ForeColor       =   &H00808080&
         Height          =   420
         Left            =   270
         TabIndex        =   47
         Top             =   1095
         Width           =   3435
      End
      Begin VB.Label lblChapo 
         AutoSize        =   -1  'True
         Caption         =   "Separator :"
         Height          =   195
         Left            =   300
         TabIndex        =   43
         Top             =   705
         Width           =   780
      End
   End
   Begin VB.Frame framStatus 
      Caption         =   "Status"
      Height          =   2025
      Left            =   135
      TabIndex        =   34
      Top             =   6240
      Width           =   3915
      Begin VB.TextBox txtPrevPass 
         Height          =   285
         Left            =   945
         TabIndex        =   39
         Top             =   1005
         Width           =   2100
      End
      Begin VB.TextBox txtPrevPrefix 
         Height          =   285
         Left            =   945
         TabIndex        =   37
         Top             =   675
         Width           =   2100
      End
      Begin VB.CheckBox chkPrev 
         Caption         =   "possibility of preview on publication"
         Height          =   240
         Left            =   300
         TabIndex        =   35
         Top             =   345
         Width           =   3435
      End
      Begin VB.Label lblPrevPass 
         AutoSize        =   -1  'True
         Caption         =   "Pass :"
         Height          =   195
         Left            =   300
         TabIndex        =   38
         Top             =   1035
         Width           =   435
      End
      Begin VB.Label lblPrevPrefix 
         AutoSize        =   -1  'True
         Caption         =   "Prefix :"
         Height          =   195
         Left            =   300
         TabIndex        =   36
         Top             =   705
         Width           =   480
      End
      Begin VB.Label lblPrevPlugin 
         Caption         =   "Requires the Preview plugin by Jason Thaxter (or equivalent) on server-side."
         ForeColor       =   &H00808080&
         Height          =   420
         Left            =   270
         TabIndex        =   40
         Top             =   1410
         Width           =   3435
      End
   End
   Begin VB.Frame framArt 
      Caption         =   "Content"
      Height          =   1245
      Left            =   135
      TabIndex        =   24
      Top             =   4875
      Width           =   3915
      Begin VB.OptionButton optArtEncode 
         Caption         =   "UTF-8"
         Height          =   195
         Index           =   1
         Left            =   2745
         TabIndex        =   29
         Top             =   780
         Width           =   825
      End
      Begin VB.OptionButton optArtEncode 
         Caption         =   "ANSI"
         Height          =   195
         Index           =   0
         Left            =   1920
         TabIndex        =   28
         Top             =   780
         Value           =   -1  'True
         Width           =   720
      End
      Begin VB.TextBox txtImgMax 
         Height          =   285
         Left            =   1905
         TabIndex        =   25
         Top             =   375
         Width           =   1770
      End
      Begin VB.Label lblArtEncode 
         Alignment       =   1  'Right Justify
         Caption         =   "Encoding :"
         Height          =   285
         Left            =   30
         TabIndex        =   27
         Top             =   765
         Width           =   1800
      End
      Begin VB.Label lblImgMax 
         Alignment       =   1  'Right Justify
         Caption         =   "Max image size (KB) :"
         Height          =   285
         Left            =   30
         TabIndex        =   26
         Top             =   405
         Width           =   1800
      End
   End
   Begin VB.Frame framTree 
      Caption         =   "Tree"
      Height          =   2010
      Left            =   135
      TabIndex        =   15
      Top             =   2730
      Width           =   3915
      Begin VB.TextBox txtFlavExt 
         Height          =   285
         Left            =   1905
         TabIndex        =   22
         Top             =   1485
         Width           =   1770
      End
      Begin VB.TextBox txtExcludedPaths 
         Height          =   285
         Left            =   1905
         TabIndex        =   18
         Top             =   735
         Width           =   1770
      End
      Begin VB.TextBox txtArtExt 
         Height          =   285
         Left            =   1905
         TabIndex        =   20
         Top             =   1110
         Width           =   1770
      End
      Begin VB.TextBox txtArtRoot 
         Height          =   285
         Left            =   1905
         TabIndex        =   16
         Top             =   375
         Width           =   1770
      End
      Begin VB.Label lblFlavExt 
         Alignment       =   1  'Right Justify
         Caption         =   "Flavour extension :"
         Height          =   285
         Left            =   30
         TabIndex        =   23
         Top             =   1500
         Width           =   1800
      End
      Begin VB.Label lblExcluded 
         Alignment       =   1  'Right Justify
         Caption         =   "Excluded paths :"
         Height          =   285
         Left            =   30
         TabIndex        =   19
         Top             =   765
         Width           =   1800
      End
      Begin VB.Label lblArtExt 
         Alignment       =   1  'Right Justify
         Caption         =   "Articles extension :"
         Height          =   285
         Left            =   30
         TabIndex        =   21
         Top             =   1125
         Width           =   1800
      End
      Begin VB.Label lblArtRoot 
         Alignment       =   1  'Right Justify
         Caption         =   "Articles root :"
         Height          =   285
         Left            =   30
         TabIndex        =   17
         Top             =   405
         Width           =   1800
      End
   End
   Begin VB.Frame framCnnx 
      Caption         =   "Connection"
      Height          =   1635
      Left            =   135
      TabIndex        =   6
      Top             =   960
      Width           =   3915
      Begin VB.TextBox txtPort 
         Height          =   285
         Left            =   3165
         TabIndex        =   8
         Top             =   375
         Width           =   495
      End
      Begin VB.TextBox txtHost 
         Height          =   285
         Left            =   1590
         TabIndex        =   7
         Top             =   375
         Width           =   1515
      End
      Begin VB.CommandButton cmdShowPass 
         Caption         =   "show"
         Height          =   285
         Left            =   2940
         TabIndex        =   13
         Top             =   1095
         Width           =   720
      End
      Begin VB.TextBox txtUser 
         Height          =   285
         Left            =   1590
         TabIndex        =   10
         Top             =   735
         Width           =   2070
      End
      Begin VB.TextBox txtPass 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1590
         PasswordChar    =   "*"
         TabIndex        =   12
         Top             =   1095
         Width           =   1335
      End
      Begin VB.Label lblUser 
         Alignment       =   1  'Right Justify
         Caption         =   "Username :"
         Height          =   285
         Left            =   135
         TabIndex        =   11
         Top             =   765
         Width           =   1380
      End
      Begin VB.Label lblPass 
         Alignment       =   1  'Right Justify
         Caption         =   "Password :"
         Height          =   285
         Left            =   135
         TabIndex        =   14
         Top             =   1110
         Width           =   1380
      End
      Begin VB.Label lblHostPort 
         Alignment       =   1  'Right Justify
         Caption         =   "Server - Port :"
         Height          =   285
         Left            =   135
         TabIndex        =   9
         Top             =   405
         Width           =   1380
      End
   End
   Begin VB.Frame framCSS 
      Caption         =   "CSS"
      Height          =   5520
      Left            =   4320
      TabIndex        =   30
      Top             =   960
      Width           =   5715
      Begin VB.TextBox txtCSS 
         Height          =   3675
         Left            =   135
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   33
         Top             =   1650
         Width           =   5430
      End
      Begin VB.Label lblCSS 
         AutoSize        =   -1  'True
         Caption         =   $"project.frx":19B2
         Height          =   390
         Index           =   1
         Left            =   135
         TabIndex        =   32
         Top             =   1065
         Width           =   5505
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblCSS 
         Caption         =   $"project.frx":1A55
         Height          =   675
         Index           =   0
         Left            =   135
         TabIndex        =   31
         Top             =   300
         Width           =   5445
         WordWrap        =   -1  'True
      End
   End
   Begin VB.TextBox txtBlosxomURL 
      Height          =   285
      Left            =   1500
      TabIndex        =   3
      Top             =   480
      Width           =   7140
   End
   Begin VB.TextBox txtTitle 
      Height          =   285
      Left            =   1500
      TabIndex        =   0
      Top             =   98
      Width           =   7140
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   300
      Left            =   8805
      TabIndex        =   5
      Top             =   495
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   345
      Left            =   8805
      TabIndex        =   2
      Top             =   90
      Width           =   1215
   End
   Begin VB.Label lblBlosxomURL 
      Alignment       =   1  'Right Justify
      Caption         =   "Blosxom URL :"
      Height          =   285
      Left            =   45
      TabIndex        =   4
      Top             =   495
      Width           =   1380
   End
   Begin VB.Label lblTitle 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Project name :"
      Height          =   285
      Left            =   180
      TabIndex        =   1
      Top             =   120
      Width           =   1245
   End
End
Attribute VB_Name = "frmProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'BlosHome (c) FFh Lab / Eric Lequien, 2009-2013 - http://ffh-lab.com
'This dialog-box handles project settings

Option Explicit
Option Base 1

Const PASSCHAR = "*"

Public bNoCancel As Boolean 'allows to inhibate the Cancel command for an incomplete project

Private Sub cmdCancel_Click()
    'Closes the dialog-box without any saving
    Unload Me
End Sub

Private Sub cmdShowPass_Click()
    'Shows/hides the password (hidden by default at dialog-box loading)
    If txtPass.PasswordChar = PASSCHAR Then
        txtPass.PasswordChar = ""
        cmdShowPass.Caption = arMsg(114)
    Else
        txtPass.PasswordChar = PASSCHAR
        cmdShowPass.Caption = arMsg(115)
    End If
End Sub

Private Sub Form_Load()
    'Informs the fields according to the global structure strucProj
    Dim strLang As String
    
    If structProj.FileName = "" Then
        Me.Caption = Me.Caption & " - nouveau projet"
    Else
        Me.Caption = Me.Caption & " - " & structProj.FileName
    
        txtTitle.Text = structProj.title
        txtBlosxomURL.Text = structProj.blosxom_url
        
        txtHost.Text = structProj.host
        txtPort.Text = structProj.port
        txtUser.Text = structProj.user
        txtPass.Text = structProj.pass
        txtPass.PasswordChar = PASSCHAR
        
        txtArtRoot.Text = structProj.art_root
        txtArtExt.Text = structProj.art_ext
        txtFlavExt.Text = structProj.flav_ext
        txtExcludedPaths.Text = structProj.excluded_paths
        
        If IsNumeric(structProj.img_max) = True Then
            txtImgMax.Text = structProj.img_max
        End If
        
        If structProj.art_encode = "UTF-8" Then
            optArtEncode(1).Value = True
        Else 'ANSI by default
            optArtEncode(0).Value = True
        End If
        
        If structProj.with_preview <> "" Then chkPrev.Value = structProj.with_preview
        txtPrevPrefix.Text = structProj.preview_prefix
        txtPrevPass.Text = structProj.preview_pass
        
        If structProj.with_chapo <> "" Then chkChapo.Value = structProj.with_chapo
        txtChapo.Text = structProj.chapo_limit
        
        txtCSS.Text = LoadText(structProj.css)
    End If
    
    'adjusts the interface according to global variables (tuned by caller before dialog-box loading)
    '***LATER : here, just one global variable is considered/useful, but it could be extended to several
    frmProject.cmdCancel.Enabled = Not (bNoCancel)
    
    'initialize the elements which depend of selected language (if none, keep design-time state ; ie. EN)
    strLang = GetLang()
    If strLang <> "(unknown)" Then SetLang strLang, "project"
    
    'some UI-strings adjustements afterward
    lblTitle.Caption = lblTitle.Caption & " :"
    lblBlosxomURL.Caption = lblBlosxomURL.Caption & " :"
    lblHostPort.Caption = lblHostPort.Caption & " :"
    lblUser.Caption = lblUser.Caption & " :"
    lblPass.Caption = lblPass.Caption & " :"
    lblArtRoot.Caption = lblArtRoot.Caption & " :"
    lblExcluded.Caption = lblExcluded.Caption & " :"
    lblArtExt.Caption = lblArtExt.Caption & " :"
    lblFlavExt.Caption = lblFlavExt.Caption & " :"
    lblImgMax.Caption = lblImgMax.Caption & " :"
    lblArtEncode.Caption = lblArtEncode.Caption & " :"
    lblPrevPrefix.Caption = lblPrevPrefix.Caption & " :"
    lblPrevPass.Caption = lblPrevPass.Caption & " :"
    lblChapo.Caption = lblChapo.Caption & " :"
End Sub

Private Sub cmdOK_Click()
    '(Re)save the project file, then clode the dialog-box
    'NB : we validate/filter values of fields in the same time
    Dim bNewProj As Boolean
        
    If structProj.FileName = "" Then
        bNewProj = True
    Else
        bNewProj = False
    End If
    
    structProj.title = txtTitle.Text
    structProj.blosxom_url = txtBlosxomURL.Text
    
    structProj.host = txtHost.Text
    
    If IsNumeric(txtPort.Text) = True Then
        structProj.port = Val(txtPort.Text)
    Else
        structProj.port = ""
    End If
    
    structProj.user = txtUser.Text
    structProj.pass = txtPass.Text
    
    structProj.art_root = txtArtRoot.Text
    structProj.art_ext = txtArtExt.Text
    structProj.flav_ext = txtFlavExt.Text
    structProj.excluded_paths = txtExcludedPaths.Text
    
    If IsNumeric(txtImgMax.Text) = True Then
        structProj.img_max = Val(txtImgMax.Text)
    Else
        structProj.img_max = ""
    End If
    
    If optArtEncode(1).Value = True Then
        structProj.art_encode = "UTF-8"
    Else 'ANSI by default
        structProj.art_encode = "ANSI"
    End If
    
    structProj.with_preview = chkPrev.Value
    structProj.preview_prefix = txtPrevPrefix.Text
    structProj.preview_pass = txtPrevPass.Text
    
    structProj.with_chapo = chkChapo.Value
    structProj.chapo_limit = txtChapo.Text
    
    structProj.css = txtCSS.Text 'contains data rather than file path until we update CSS file below
    
    SaveProject
    ApplyCleanProject True
    
    If bNewProj = True Then EnableLocalWork True
    Unload Me
End Sub
