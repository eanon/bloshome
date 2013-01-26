VERSION 5.00
Begin VB.Form frmSelect 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Projets"
   ClientHeight    =   2835
   ClientLeft      =   2985
   ClientTop       =   2430
   ClientWidth     =   4470
   ControlBox      =   0   'False
   LinkTopic       =   "select"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Annuler"
      Default         =   -1  'True
      Height          =   510
      Left            =   3525
      TabIndex        =   1
      Top             =   75
      Width           =   840
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   510
      Left            =   2595
      TabIndex        =   0
      Top             =   75
      Width           =   840
   End
   Begin VB.ListBox lstProjects 
      Height          =   1815
      Left            =   90
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   705
      Width           =   4290
   End
   Begin VB.Label lblStatus 
      Caption         =   "not any project found"
      ForeColor       =   &H00404040&
      Height          =   240
      Left            =   120
      TabIndex        =   4
      Top             =   2580
      Width           =   4860
   End
   Begin VB.Label lblGuide 
      Caption         =   "common box to select project for opening or deletion."
      Height          =   465
      Left            =   150
      TabIndex        =   2
      Top             =   120
      Width           =   2370
   End
End
Attribute VB_Name = "frmSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'BlosHome (c) FFh Lab / Eric Lequien, 2009-2013 - http://ffh-lab.com
'This dialog-box manages project files selection for opening or deletion (related to "Project" menu)

Option Explicit
Option Base 1

Private Sub cmdCancel_Click()
    'Quit the dialog-box, nothing else being done
    Unload Me
End Sub

Private Sub cmdOK_Click()
    'Opens or deletes the selected project according to the bOpenDelProj flag
    Dim strProj As String
    Dim bDoIt As Boolean
    
    If lstProjects.ListIndex = -1 Then Exit Sub
    
    strProj = lstProjects.List(lstProjects.ListIndex)
    bDoIt = False
    
    If bOpenDelProj = True Then
        'Opens the project taking care to verify that the way is clear
        If strProj = structProj.FileName Then
            bDoIt = IsFreeSDI(arMsg(137) & ".")
        Else
            bDoIt = IsFreeSDI(arMsg(138) & ".")
        End If
        
        If bDoIt = True Then
            structProj.FileName = lstProjects.List(lstProjects.ListIndex)
            LoadUnloadProject True
            Unload Me
        End If
    Else
        'Deletes the project taking care to ask for confirmation if it's the current loaded one
        If strProj = structProj.FileName Then
            bDoIt = IsFreeSDI(arMsg(139) & ".")
        Else
            bDoIt = True
        End If
        
        If bDoIt = True Then
            DelProject App.Path & "\" & PROJECTS_DIR & "\" & strProj
            PopulateProjList
        End If
    End If
End Sub

Private Sub Form_Load()
    'Initializes the dialog-box content according to the operation required by bOpenDelProj
    Dim strMsg As String
    Dim strLang As String
    
    strMsg = arMsg(134) & " "
    
    If bOpenDelProj = True Then
        lblGuide = strMsg & arMsg(135) & "."
    Else
        lblGuide = strMsg & arMsg(136) & "."
    End If
    
    PopulateProjList

    'init elts which depend of selected language (if none, keep design-time state ; i.e. English)
    strLang = GetLang()
    If strLang <> "(unknown)" Then SetLang strLang, "select"
End Sub

Private Sub lstProjects_DblClick()
    'Same as clicking OK to open selected project
    If bOpenDelProj = True Then
        cmdOK_Click
    End If
End Sub

Sub PopulateProjList()
    'Populates the list with the existing projects file
    Dim nProjs As Integer
    Dim strProj As String
    
    strProj = Dir(App.Path & "\" & PROJECTS_DIR & "\*" & PROJFILE_EXT, vbNormal)
    
    lstProjects.Clear
    While (strProj <> "")
        lstProjects.AddItem strProj
        strProj = Dir
    Wend
    
    nProjs = lstProjects.ListCount
    If nProjs > 1 Then
        lblStatus.Caption = nProjs & " " & arMsg(140)
    Else
        lblStatus.Caption = nProjs & " " & arMsg(141)
    End If
End Sub
