VERSION 5.00
Begin VB.Form frmConnect 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Connection"
   ClientHeight    =   2265
   ClientLeft      =   3840
   ClientTop       =   4305
   ClientWidth     =   5295
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   5295
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtHost 
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   0
      Width           =   3495
   End
   Begin VB.TextBox txtUser 
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   360
      Width           =   3495
   End
   Begin VB.TextBox txtPass 
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   720
      Width           =   3495
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Host Name"
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "User Name"
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "PassWord"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   840
      Width           =   1455
   End
End
Attribute VB_Name = "frmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOk_Click()
  SaveSetting App.Title, "Connection", "Host", txtHost
  SaveSetting App.Title, "Connection", "User", txtUser
  SaveSetting App.Title, "Connection", "Pass", txtPass
  Me.Hide
End Sub

Private Sub Form_Load()
  txtHost = GetSetting(App.Title, "Connection", "Host", "")
  txtUser = GetSetting(App.Title, "Connection", "User", "")
  txtPass = GetSetting(App.Title, "Connection", "Pass", "")
End Sub
