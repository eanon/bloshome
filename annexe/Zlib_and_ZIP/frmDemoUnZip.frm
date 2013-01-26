VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInfoUnZipClsZip 
   Caption         =   "Infos et Dézippage d'un fichier Zip avec zLib"
   ClientHeight    =   4110
   ClientLeft      =   240
   ClientTop       =   4065
   ClientWidth     =   4185
   LinkTopic       =   "Form1"
   ScaleHeight     =   4110
   ScaleWidth      =   4185
   Begin VB.CommandButton cmdLoadFileZ 
      Caption         =   "&Ouvrir fichier zippé sans le décompresser"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   3000
      Width           =   3855
   End
   Begin VB.CommandButton cdeExtraireUnFichier 
      Caption         =   "&Extraire un seul fichier"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   2040
      Width           =   3855
   End
   Begin VB.CommandButton cdeVoirContenuZip 
      Caption         =   "Voir le contenu du Zip (fenêtre de debug)"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   3855
   End
   Begin VB.CommandButton cdeVoirComment 
      Caption         =   "&Voir le commentaire du Zip"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   3855
   End
   Begin MSComctlLib.ProgressBar ProgressBar 
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   3600
      Visible         =   0   'False
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   3855
      Width           =   4185
      _ExtentX        =   7382
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   6853
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cdeExtraireTout 
      Caption         =   "Extrait &tous les fichiers"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   2520
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "Tous les fichiers extraits seront placés dans le sous-répertoire ""Fichiers extraits"""
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   3855
   End
   Begin VB.Label lblNbFichiers 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   3240
      TabIndex        =   4
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Nombre de fichiers dans le Zip "
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "frmInfoUnZipClsZip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Déclare un objet Zip conforme à la classe
Dim WithEvents MonZip As clsZip
Attribute MonZip.VB_VarHelpID = -1

Private NomFichierZip As String     ' Nom unique (plus facile pour faire vos essais)

Private Sub cmdLoadFileZ_Click()
    ' même chose que extraction (optimiser les lignes de code !)
    Dim NoFichier As Long
    Dim sTemp As String
    
    ' Choix du n° du fichier ) extraire
    sTemp = InputBox("Entrer le numéro du fichier à extraire :" & vbCrLf & _
                     "Numéro compris entre 1 et " & CStr(MonZip.inFileCount) & vbCrLf & _
                     "Ordre des fichiers : Voir fenêtre de debug après avoir cliqué sur " & _
                     "le bouton ""Voir le contenu du Zip (...)""", App.Title, 1)
    If sTemp = vbNullString Then Exit Sub
    NoFichier = Val(sTemp)
    
    ' Vérification du n°
    If NoFichier < 1 Or NoFichier > MonZip.inFileCount Then
        Beep
        StatusBar.Panels(1).Text = "Il n'existe pas de fichier n° " & CStr(NoFichier) & " dans le fichier Zip"
        Exit Sub
    End If
    
    ' Ouvre le Zip (s'il ne l'ai pas encore)
    If Not MonZip.ZipIsOpen Then MonZip.ZipOpen (NomFichierZip)
    '----------------------------------------------------------------
    
    'extraction vers Rép Temp de windows sans intervention de l'utilisateur ce qui donne l'impression qu'il n'y a pas de décompression
    sTemp = MonZip.ExtractSgFileToTmp(NoFichier)
    'ouvrir le fichier décompressé en Temp avec gestion de son autodestruction lors du prochain démarrage de windows
    'inscription d'une clé RunOnce dans la Bdr qui lance un fichier destrucFic.bat (voir commentaires dans clsZip)
    Call MonZip.LoadZippedFile(sTemp)
    
End Sub

Private Sub Form_Load()

    ' Défini l'objet Zip
    Set MonZip = New clsZip

    ' Défini ici le nom du Zip (utilisé plusieurs fois dans cette feuille,
    '   sera plus facile pour faire vos propres essais)
    NomFichierZip = App.Path & "\Test.zip"
    
    ' Ouvre le Zip
    MonZip.ZipOpen (NomFichierZip)
    
    ' Affiche le nombre de fichiers
    lblNbFichiers.Caption = CStr(MonZip.inFileCount)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    If Not MonZip Is Nothing Then
        ' On détruit l'objet Zip avant de sortir (plus propre)
        '   et surtout permet de fermer le Zip !
        Set MonZip = Nothing
    End If

End Sub

Private Sub cdeExtraireTout_Click()

    ' Ouvre le Zip (s'il ne l'ai pas encore)
    If Not MonZip.ZipIsOpen Then MonZip.ZipOpen (NomFichierZip)
    
    ' Lance l'extraction
    Call MonZip.ExtractAllFiles(App.Path & "\Fichiers extraits", True, False, True) 'option conserve les dates
    
End Sub

Private Sub cdeExtraireUnFichier_Click()

    Dim NoFichier As Long
    Dim sTemp As String
    
    ' Choix du n° du fichier ) extraire
    sTemp = InputBox("Entrer le numéro du fichier à extraire :" & vbCrLf & _
                     "Numéro compris entre 1 et " & CStr(MonZip.inFileCount) & vbCrLf & _
                     "Ordre des fichiers : Voir fenêtre de debug après avoir cliqué sur " & _
                     "le bouton ""Voir le contenu du Zip (...)""", App.Title, 1)
    If sTemp = vbNullString Then Exit Sub
    NoFichier = Val(sTemp)
    
    ' Vérification du n°
    If NoFichier < 1 Or NoFichier > MonZip.inFileCount Then
        Beep
        StatusBar.Panels(1).Text = "Il n'existe pas de fichier n° " & CStr(NoFichier) & " dans le fichier Zip"
        Exit Sub
    End If
    
    ' Ouvre le Zip (s'il ne l'ai pas encore)
    If Not MonZip.ZipIsOpen Then MonZip.ZipOpen (NomFichierZip)
    
    ' Extraction du fichier sur le répertoire de l'application
    Call MonZip.ExtractSingleFile(NoFichier, App.Path & "\Fichiers extraits", False, True)

End Sub

Private Sub cdeVoirComment_Click()

    MsgBox MonZip.ZipComment

End Sub

Private Sub cdeVoirContenuZip_Click()

    ' Pour lire les noms des fichiers, leurs dates/heure et taille avant/après
    '   il suffit de faire une boucle :
    Dim r As Long
    With MonZip
        For r = 1 To .inFileCount
            Debug.Print r, _
                        .inFileName(r), _
                        .FileDateAndTime(r), _
                        .FileCompressedSize(r) & " octets (compressé)", _
                        .FileUncompressedSize(r) & " octets (décompressé)", _
                        " Taux = "; CStr(.FileUncompressedSize(r) / .FileCompressedSize(r))
        Next r
    End With
    
End Sub

' Données fournies par la Classe
Private Sub MonZip_Progress(ByVal Percent As Long, _
                            ByRef Cancel As Boolean)
    ProgressBar.Value = Percent
    ' En faisant passer Cancel à True, on interrompt le dézippage

End Sub

' Données fournies par la Classe
Private Sub MonZip_Status(ByVal Text As String)
    
    Me.StatusBar.Panels(1).Text = Text
    Select Case Text
        Case "Début d'extraction"
            ProgressBar.Value = 0
            ProgressBar.Visible = True
        Case "Extraction terminée"
            ProgressBar.Visible = False
    End Select
    
End Sub

' Données fournies par la Classe
Private Sub MonZip_ZipError(ByVal Number As eZipError, ByVal Description As String, _
                            Cancel As Boolean)
    
    MsgBox "Erreur lors de l'ouverture ou l'extraction du Zip" & vbCrLf & _
           "Erreur " & CStr(Number) & " - " & Description, _
           vbCritical + vbOKOnly, App.Title
           
End Sub
