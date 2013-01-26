VERSION 5.00
Begin VB.Form frmAbout 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About BlosHome"
   ClientHeight    =   6015
   ClientLeft      =   1305
   ClientTop       =   2010
   ClientWidth     =   5790
   ClipControls    =   0   'False
   Icon            =   "about.frx":0000
   LinkTopic       =   "about"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4151.657
   ScaleMode       =   0  'User
   ScaleWidth      =   5437.11
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtLegal 
      CausesValidation=   0   'False
      ForeColor       =   &H00404040&
      Height          =   1335
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   17
      Text            =   "about.frx":19B2
      Top             =   4560
      Width           =   5535
   End
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
      Left            =   5100
      Negotiate       =   -1  'True
      Picture         =   "about.frx":1E09
      ScaleHeight     =   337.12
      ScaleMode       =   0  'User
      ScaleWidth      =   337.12
      TabIndex        =   1
      Top             =   90
      Width           =   510
   End
   Begin VB.Line Line2 
      Index           =   1
      X1              =   0
      X2              =   5408.938
      Y1              =   3064.565
      Y2              =   3064.565
   End
   Begin VB.Label lblZip 
      Caption         =   "- The Zlib / Zip interface class © yan35 <- Jack <- Andrew McMillan"
      Height          =   195
      Left            =   210
      TabIndex        =   10
      Top             =   2415
      Width           =   5280
   End
   Begin VB.Label lblZlib 
      Caption         =   "- The (de)compression library Zlib © Jean-loup Gailly and Mark Adler"
      Height          =   195
      Left            =   225
      TabIndex        =   9
      Top             =   2175
      Width           =   5280
   End
   Begin VB.Label lblEZImage 
      Caption         =   "- The EasyImage plugin (for TinyMCE) © FFh Lab / Eric Lequien"
      Height          =   195
      Left            =   210
      TabIndex        =   7
      Top             =   1710
      Width           =   5280
   End
   Begin VB.Line Line2 
      Index           =   2
      X1              =   0
      X2              =   5408.938
      Y1              =   1905
      Y2              =   1905
   End
   Begin VB.Label lblBlosxom 
      AutoSize        =   -1  'True
      Caption         =   "- The active community around Emanuel Sprosec : "
      Height          =   195
      Index           =   3
      Left            =   210
      TabIndex        =   15
      Top             =   4095
      Width           =   3615
   End
   Begin VB.Label lblBlosxomSite 
      AutoSize        =   -1  'True
      Caption         =   "http://muli.cc"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   1
      Left            =   3855
      MousePointer    =   10  'Up Arrow
      TabIndex        =   16
      Top             =   4095
      Width           =   960
   End
   Begin VB.Label lblBlosxomSite 
      AutoSize        =   -1  'True
      Caption         =   "http://www.blosxom.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   0
      Left            =   3495
      MousePointer    =   10  'Up Arrow
      TabIndex        =   14
      Top             =   3870
      Width           =   1785
   End
   Begin VB.Label lblBlosxom 
      AutoSize        =   -1  'True
      Caption         =   $"about.frx":37BB
      Height          =   585
      Index           =   0
      Left            =   150
      TabIndex        =   11
      Top             =   2880
      Width           =   5475
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblBlosxom 
      AutoSize        =   -1  'True
      Caption         =   "To know more about the Blosxom blog engine :"
      Height          =   195
      Index           =   1
      Left            =   150
      TabIndex        =   12
      Top             =   3615
      Width           =   3330
   End
   Begin VB.Label lblBlosxom 
      AutoSize        =   -1  'True
      Caption         =   "- The original site by its author, Raël Dornfest :"
      Height          =   195
      Index           =   2
      Left            =   210
      TabIndex        =   13
      Top             =   3870
      Width           =   3255
   End
   Begin VB.Label lblDiFtpCli 
      Caption         =   "- A modified version of the FTP client DiFtpCli6 © Jean-Luc Delbeke"
      Height          =   195
      Left            =   210
      TabIndex        =   8
      Top             =   1950
      Width           =   5280
   End
   Begin VB.Label lblTinyMCE 
      Caption         =   "- The WYSIWYG HTML editor TinyMCE © Moxiecode Systems AB"
      Height          =   195
      Left            =   210
      TabIndex        =   6
      Top             =   1485
      Width           =   5280
   End
   Begin VB.Label lblTiers 
      AutoSize        =   -1  'True
      Caption         =   "This software embeds these third-party components :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   150
      TabIndex        =   5
      Top             =   1215
      Width           =   4500
   End
   Begin VB.Line Line2 
      Index           =   0
      X1              =   0
      X2              =   5408.938
      Y1              =   755.788
      Y2              =   755.788
   End
   Begin VB.Label lblWebsite 
      Caption         =   "Website"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   3720
      MousePointer    =   10  'Up Arrow
      TabIndex        =   4
      Top             =   675
      Width           =   1860
   End
   Begin VB.Label lblCopyright 
      Caption         =   "Copyright"
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   150
      TabIndex        =   3
      Top             =   675
      Width           =   3555
   End
   Begin VB.Label lblDescription 
      Caption         =   "Description"
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   150
      TabIndex        =   2
      Top             =   420
      Width           =   4275
   End
   Begin VB.Label lblName 
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   150
      TabIndex        =   0
      Top             =   165
      Width           =   4275
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'BlosHome (c) FFh Lab / Eric Lequien, 2009-2013 - http://ffh-lab.com
'This about-box talks about BlosHome

Option Explicit
Option Base 1

Private Sub Form_Load()
    Dim strLang As String

    'Initializes the different fields from the VERSION_INFO resource
    lblName.Caption = App.title & " " & App.Major & "." & App.Minor & " rev." & App.Revision & " alpha EN/FR"
    lblCopyright.Caption = "Copyright " & App.LegalCopyright & "  -"
    lblWebsite.Caption = App.CompanyName
    lblDescription.Caption = arMsg(4)
    
    'init elts which depend of selected language (if none, keep design-time state ; i.e. English)
    strLang = GetLang()
    If strLang <> "(unknown)" Then SetLang strLang, "about"
End Sub

Private Sub lblBlosxomSite_Click(index As Integer)
    'Attempts to reach the Blosxom-related site (one per label index) using the default web browser
    If Left$(lblBlosxomSite(index).Caption, 7) <> "http://" Then
        MsgBox lblBlosxomSite(index).Caption & " " & arMsg(116) & " !", vbCritical, "DEBUG"
    End If
    OpenMIME (lblBlosxomSite(index).Caption)
End Sub

Private Sub lblWebsite_Click()
    'Attempts to reach the FFh Lab site using the default Web browser
    If Left$(lblWebsite.Caption, 7) <> "http://" Then
        MsgBox App.CompanyName & " " & arMsg(116) & " !", vbCritical, "DEBUG"
    End If
    OpenMIME (lblWebsite.Caption)
End Sub
