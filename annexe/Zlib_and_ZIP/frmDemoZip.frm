VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDemoClsZip 
   Caption         =   "Demo clsZip avec zLib"
   ClientHeight    =   6780
   ClientLeft      =   330
   ClientTop       =   4245
   ClientWidth     =   4485
   LinkTopic       =   "Form1"
   ScaleHeight     =   6780
   ScaleWidth      =   4485
   Begin VB.CommandButton cmdSu 
      Caption         =   "Su&pprimer un Fichier du fichier zipp�  TestVersGros..."
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   5160
      Width           =   4215
   End
   Begin VB.CommandButton CmdAjZip 
      Caption         =   "A&jouter 1 fichier au fichier zipp� TestVersGrosZip.Zip"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   4680
      Width           =   4215
   End
   Begin VB.CommandButton CmdConceptYB 
      Caption         =   "Autre &m�thode Cr�er 1 fichier Zip    (fichiers *.* avec chemin complet, pour l'exemple)    M�thode compatible gros Zip final"
      Height          =   735
      Left            =   120
      TabIndex        =   10
      Top             =   3840
      Width           =   4215
   End
   Begin VB.CommandButton CdeUZip 
      Caption         =   "Affichage et D�zippage d'1 fichier Zip"
      Height          =   375
      Left            =   720
      TabIndex        =   9
      Top             =   5880
      Width           =   3015
   End
   Begin VB.CommandButton CdeSupp 
      Caption         =   "&Supprimer un Fichier de la s�lection avant Zippage"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   1920
      Width           =   4215
   End
   Begin VB.CommandButton cdeAjouterAvecR�pertoiresRelatifs 
      Caption         =   "Ajouter &fichiers *.TXT (avec chemin relatif)"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   4215
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   6525
      Width           =   4485
      _ExtentX        =   7911
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   7382
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cdeListeFichiers 
      Caption         =   "&Liste des fichiers ins�r�s"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2760
      Width           =   4215
   End
   Begin VB.CommandButton cdeAjouterAvecR�pertoires 
      Caption         =   "Ajouter &fichiers *.CLS (avec chemin complet)"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   4215
   End
   Begin VB.CommandButton cdeCr��Zip 
      Caption         =   "&Cr�� le fichier Zip et g�rer son commentaire ""Test.Zip"""
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3240
      Width           =   4215
   End
   Begin VB.CommandButton cdeCombien 
      Caption         =   "Renvoie le &nombre de fichiers d�j� ajout�s"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   2400
      Width           =   4215
   End
   Begin VB.CommandButton cdeAjouterSansR�pertoire 
      Caption         =   "&Ajouter fichiers *.TXT (pas de chemin)"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   4215
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0080FFFF&
      X1              =   120
      X2              =   4320
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Label Label1 
      Caption         =   "Cr�ation du fichier ""Test.Zip"" avec certains fichiers pr�sents sur ce r�pertoire"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "frmDemoClsZip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim MonZip  As clsZip

Private Sub CdeSupp_Click()
    Dim r
    MsgBox "Regarder dans la liste des fichiers s�lectionn�s � zipper, le n� de celui � supprimer"
    cdeListeFichiers_Click
    r = InputBox("Indiquer le N� du fichier � effacer de la liste � zipper", "Entrer le param�tre")
    If Not IsNumeric(r) Then
        MsgBox "C'est un n� qui est demand� !", vbExclamation, "Echec dans la saisie"
        Exit Sub
    End If
    'suppression
    If MonZip.FileDel(r) Then
        StatusBar.Panels(1).Text = "Fichier supprim� des �l�ments � zipper"
    Else
        StatusBar.Panels(1).Text = "Erreur suppression ! Les �l�ments � zipper sont inchang�s."
    End If

End Sub

Private Sub CdeUZip_Click()
    frmInfoUnZipClsZip.Show
    Unload Me
End Sub

Private Sub CmdAjZip_Click()
    If MonZip.FileAddIntoZip(App.Path & "\TestVersGrosZip.zip", _
    InputBox("Indiquer le nom du fichier � ajouter avec son chemin", _
    "Ajout d'1 fichier � 1 zip existant"), WithCompletePath) Then
        StatusBar.Panels(1).Text = "fichier saisi ajout� au fichier Zip existant"
    Else
        StatusBar.Panels(1).Text = "Echec ajout au zip, Zip existant inchang�"
    End If
End Sub

Private Sub CmdConceptYB_Click()
    Dim r, n As Long
    r = MonZip.CreateZip(App.Path & "\TestVersGrosZip.zip", True)
    n = MonZip.ybFileAdd(App.Path & "\*.*", WithCompletePath, CLng(r), True)
    If MonZip.WriteEndZip(CLng(r)) Then
        StatusBar.Panels(1).Text = "fichier zip cr�� avec " & CStr(n) & " fichiers"
    Else
        Kill App.Path & "\TestVersGrosZip.zip"
        StatusBar.Panels(1).Text = "Echec de cr�ation du fichier zip"
    End If
End Sub

Private Sub cmdSu_Click()
    If MonZip.FileRemoveFromZip(App.Path & "\TestVersGrosZip.zip", _
    InputBox("Indiquer le nom du fichier � ajouter avec son chemin", _
    "Ajout d'1 fichier � 1 zip existant")) Then
        StatusBar.Panels(1).Text = "fichier saisi supprim� du fichier Zip existant"
    Else
        StatusBar.Panels(1).Text = "Echec de suppression du zip, Zip existant inchang�"
    End If
End Sub

Private Sub Form_Load()

    ' Pr�pare un nouveau Zip
    Set MonZip = New clsZip

End Sub

Private Sub cdeAjouterSansR�pertoire_Click()

    Dim r As Long
    r = MonZip.FileAdd(App.Path & "\*.txt", WithoutPath, True)
    StatusBar.Panels(1).Text = CStr(r) & " fichiers ajout�s � zipper"
    
End Sub

Private Sub cdeAjouterAvecR�pertoires_Click()

    Dim r As Long
    r = MonZip.FileAdd("*.cls", WithCompletePath, True)
    StatusBar.Panels(1).Text = CStr(r) & " fichiers ajout�s � zipper"
    
End Sub

Private Sub cdeAjouterAvecR�pertoiresRelatifs_Click()

    Dim r As Long

    ' Volontairement, on choisit un r�pertoire en dessous du notre
    ' pour montrer que si le chemin de l'application n'est pas celui
    ' des fichiers cibles, on ins�re le chemin complet
    r = MonZip.FileAdd(App.Path & "\..\*.txt", WithRelativePath, True)
    StatusBar.Panels(1).Text = CStr(r) & " fichiers ajout�s a zipper"

End Sub

Private Sub cdeCombien_Click()

    StatusBar.Panels(1).Text = "Il y a " & CStr(MonZip.FileCount) & " fichiers s�lectionn�s pour l'instant."

End Sub

Private Sub cdeListeFichiers_Click()

    Dim Qt� As Long, Temp As String, r As Long
    
    ' R�cup�re le nombre de fichiers
    Qt� = MonZip.FileCount
    ' Pour chaque fichier, on va r�cup�rer le nom
    With MonZip
        For r = 1 To Qt�
            Temp = Temp & r & " - " & CStr(.FileName(r)) & " - compression � : " & .SizeCompressed(r) & _
            " - taille orig. : " & .SizeUncompressed(r) & vbCrLf
        Next r
        MsgBox Temp
    End With
End Sub

Private Sub cdeCr��Zip_Click()

    Dim Ret As Boolean
    
    If MonZip.FileCount = 0 Then Exit Sub
    
    ' Avant de zipper, on ajoute un commentaire pour faire joli
    '   (Vous le verrez en ouvrant le zip avec Winzip ou autre)
    MonZip.Comment = "Source initiale de Jack, disponible sur http://www.vbfrance.com/code.aspx?ID=24072" & vbCrLf & _
    "Modifi�e par Yan35"
    MonZip.Comment = InputBox("Si vous le souhaitez, modifiez le commentaire du Zip � cr�er :", "Gestion commentaire", _
                    MonZip.Comment)
    ' Cr�� le zip
    Ret = MonZip.WriteZip(App.Path & "\Test.zip", True)
    ' Teste l'info revoy�e
    If Ret Then
        StatusBar.Panels(1).Text = """Test.Zip"" a �t� cr��."
    Else
        StatusBar.Panels(1).Text = "Erreur de cr�ation du Zip"
    End If
    
End Sub

