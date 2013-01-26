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
      Caption         =   "Su&pprimer un Fichier du fichier zippé  TestVersGros..."
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   5160
      Width           =   4215
   End
   Begin VB.CommandButton CmdAjZip 
      Caption         =   "A&jouter 1 fichier au fichier zippé TestVersGrosZip.Zip"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   4680
      Width           =   4215
   End
   Begin VB.CommandButton CmdConceptYB 
      Caption         =   "Autre &méthode Créer 1 fichier Zip    (fichiers *.* avec chemin complet, pour l'exemple)    Méthode compatible gros Zip final"
      Height          =   735
      Left            =   120
      TabIndex        =   10
      Top             =   3840
      Width           =   4215
   End
   Begin VB.CommandButton CdeUZip 
      Caption         =   "Affichage et Dézippage d'1 fichier Zip"
      Height          =   375
      Left            =   720
      TabIndex        =   9
      Top             =   5880
      Width           =   3015
   End
   Begin VB.CommandButton CdeSupp 
      Caption         =   "&Supprimer un Fichier de la sélection avant Zippage"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   1920
      Width           =   4215
   End
   Begin VB.CommandButton cdeAjouterAvecRépertoiresRelatifs 
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
      Caption         =   "&Liste des fichiers insérés"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2760
      Width           =   4215
   End
   Begin VB.CommandButton cdeAjouterAvecRépertoires 
      Caption         =   "Ajouter &fichiers *.CLS (avec chemin complet)"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   4215
   End
   Begin VB.CommandButton cdeCrééZip 
      Caption         =   "&Créé le fichier Zip et gérer son commentaire ""Test.Zip"""
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3240
      Width           =   4215
   End
   Begin VB.CommandButton cdeCombien 
      Caption         =   "Renvoie le &nombre de fichiers déjà ajoutés"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   2400
      Width           =   4215
   End
   Begin VB.CommandButton cdeAjouterSansRépertoire 
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
      Caption         =   "Création du fichier ""Test.Zip"" avec certains fichiers présents sur ce répertoire"
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
    MsgBox "Regarder dans la liste des fichiers sélectionnés à zipper, le n° de celui à supprimer"
    cdeListeFichiers_Click
    r = InputBox("Indiquer le N° du fichier à effacer de la liste à zipper", "Entrer le paramètre")
    If Not IsNumeric(r) Then
        MsgBox "C'est un n° qui est demandé !", vbExclamation, "Echec dans la saisie"
        Exit Sub
    End If
    'suppression
    If MonZip.FileDel(r) Then
        StatusBar.Panels(1).Text = "Fichier supprimé des éléments à zipper"
    Else
        StatusBar.Panels(1).Text = "Erreur suppression ! Les éléments à zipper sont inchangés."
    End If

End Sub

Private Sub CdeUZip_Click()
    frmInfoUnZipClsZip.Show
    Unload Me
End Sub

Private Sub CmdAjZip_Click()
    If MonZip.FileAddIntoZip(App.Path & "\TestVersGrosZip.zip", _
    InputBox("Indiquer le nom du fichier à ajouter avec son chemin", _
    "Ajout d'1 fichier à 1 zip existant"), WithCompletePath) Then
        StatusBar.Panels(1).Text = "fichier saisi ajouté au fichier Zip existant"
    Else
        StatusBar.Panels(1).Text = "Echec ajout au zip, Zip existant inchangé"
    End If
End Sub

Private Sub CmdConceptYB_Click()
    Dim r, n As Long
    r = MonZip.CreateZip(App.Path & "\TestVersGrosZip.zip", True)
    n = MonZip.ybFileAdd(App.Path & "\*.*", WithCompletePath, CLng(r), True)
    If MonZip.WriteEndZip(CLng(r)) Then
        StatusBar.Panels(1).Text = "fichier zip créé avec " & CStr(n) & " fichiers"
    Else
        Kill App.Path & "\TestVersGrosZip.zip"
        StatusBar.Panels(1).Text = "Echec de création du fichier zip"
    End If
End Sub

Private Sub cmdSu_Click()
    If MonZip.FileRemoveFromZip(App.Path & "\TestVersGrosZip.zip", _
    InputBox("Indiquer le nom du fichier à ajouter avec son chemin", _
    "Ajout d'1 fichier à 1 zip existant")) Then
        StatusBar.Panels(1).Text = "fichier saisi supprimé du fichier Zip existant"
    Else
        StatusBar.Panels(1).Text = "Echec de suppression du zip, Zip existant inchangé"
    End If
End Sub

Private Sub Form_Load()

    ' Prépare un nouveau Zip
    Set MonZip = New clsZip

End Sub

Private Sub cdeAjouterSansRépertoire_Click()

    Dim r As Long
    r = MonZip.FileAdd(App.Path & "\*.txt", WithoutPath, True)
    StatusBar.Panels(1).Text = CStr(r) & " fichiers ajoutés à zipper"
    
End Sub

Private Sub cdeAjouterAvecRépertoires_Click()

    Dim r As Long
    r = MonZip.FileAdd("*.cls", WithCompletePath, True)
    StatusBar.Panels(1).Text = CStr(r) & " fichiers ajoutés à zipper"
    
End Sub

Private Sub cdeAjouterAvecRépertoiresRelatifs_Click()

    Dim r As Long

    ' Volontairement, on choisit un répertoire en dessous du notre
    ' pour montrer que si le chemin de l'application n'est pas celui
    ' des fichiers cibles, on insère le chemin complet
    r = MonZip.FileAdd(App.Path & "\..\*.txt", WithRelativePath, True)
    StatusBar.Panels(1).Text = CStr(r) & " fichiers ajoutés a zipper"

End Sub

Private Sub cdeCombien_Click()

    StatusBar.Panels(1).Text = "Il y a " & CStr(MonZip.FileCount) & " fichiers sélectionnés pour l'instant."

End Sub

Private Sub cdeListeFichiers_Click()

    Dim Qté As Long, Temp As String, r As Long
    
    ' Récupère le nombre de fichiers
    Qté = MonZip.FileCount
    ' Pour chaque fichier, on va récupérer le nom
    With MonZip
        For r = 1 To Qté
            Temp = Temp & r & " - " & CStr(.FileName(r)) & " - compression à : " & .SizeCompressed(r) & _
            " - taille orig. : " & .SizeUncompressed(r) & vbCrLf
        Next r
        MsgBox Temp
    End With
End Sub

Private Sub cdeCrééZip_Click()

    Dim Ret As Boolean
    
    If MonZip.FileCount = 0 Then Exit Sub
    
    ' Avant de zipper, on ajoute un commentaire pour faire joli
    '   (Vous le verrez en ouvrant le zip avec Winzip ou autre)
    MonZip.Comment = "Source initiale de Jack, disponible sur http://www.vbfrance.com/code.aspx?ID=24072" & vbCrLf & _
    "Modifiée par Yan35"
    MonZip.Comment = InputBox("Si vous le souhaitez, modifiez le commentaire du Zip à créer :", "Gestion commentaire", _
                    MonZip.Comment)
    ' Créé le zip
    Ret = MonZip.WriteZip(App.Path & "\Test.zip", True)
    ' Teste l'info revoyée
    If Ret Then
        StatusBar.Panels(1).Text = """Test.Zip"" a été créé."
    Else
        StatusBar.Panels(1).Text = "Erreur de création du Zip"
    End If
    
End Sub

