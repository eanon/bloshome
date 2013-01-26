VERSION 5.00
Begin VB.Form frmProgress 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1245
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5190
   ControlBox      =   0   'False
   FillColor       =   &H00404040&
   Icon            =   "progress.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1245
   ScaleWidth      =   5190
   ShowInTaskbar   =   0   'False
   Begin VB.Timer timerClose 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   0
      Top             =   0
   End
   Begin VB.Label lblDetail 
      Alignment       =   2  'Center
      Caption         =   "[detail]"
      Height          =   480
      Left            =   120
      TabIndex        =   1
      Top             =   735
      Width           =   4995
   End
   Begin VB.Shape shapBar 
      BorderColor     =   &H00404040&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   240
      Left            =   105
      Top             =   420
      Width           =   495
   End
   Begin VB.Shape shapFrame 
      BorderColor     =   &H00404040&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   105
      Top             =   405
      Width           =   5000
   End
   Begin VB.Label lblStage 
      Alignment       =   2  'Center
      Caption         =   "[stage]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   0
      Top             =   90
      Width           =   4995
   End
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'BlosHome (c) FFh Lab / Eric Lequien, 2009-2013 - http://ffh-lab.com
' This window handles a progress-bar mechanism

Option Explicit

Private Sub Form_Activate()
    'Centers on the main interface
    '***LATER : to be more generic, we could center on any owner/parent (rather than frmMain only)
    Left = frmMain.Left + ((frmMain.Width - frmProgress.Width) / 2)
    Top = frmMain.Top + ((frmMain.Height - frmProgress.Height) / 2)
End Sub

Private Sub Form_Load()
    Dim strLang As String
    
    'Resets the progression
    lblStage.Caption = ""
    lblDetail.Caption = ""
    shapBar.Width = 0
    
    'Initializes elements which depend of selected language (if none, keep design-time state ; ie. EN)
    strLang = GetLang()
    If strLang <> "(unknown)" Then SetLang strLang, "progress"
End Sub

Private Sub timerClose_Timer()
    'This timer allows to differ closing from request-for-closing, so that user has time to see 100%
    'USE : caller will do "frmProgress.timerClose.enabled = true" rather than "Unload frmProgress"
    Static nPass As Integer
    
    nPass = nPass + 1
    If nPass > 1 Then Unload Me
End Sub
