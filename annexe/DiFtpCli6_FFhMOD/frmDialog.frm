VERSION 5.00
Begin VB.Form frmDialog 
   Caption         =   "Ftp dialog"
   ClientHeight    =   2640
   ClientLeft      =   2220
   ClientTop       =   7725
   ClientWidth     =   11880
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   2640
   ScaleWidth      =   11880
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   7455
   End
End
Attribute VB_Name = "frmDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Resize()
  List1.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
  If List1.ListCount > 32000 Then
    List1.RemoveItem 0
  End If
  
End Sub
