VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsZip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Regroupement dans une classe pour faciliter la mise en oeuvre des fonctions pour Zipper et déZipper de ZLIB.DLL, libre de droits. Facilite le partage de cette classe commune à plusieurs applications."
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'BlosHome (c) FFh Lab / Eric Lequien, 2009-2013 - http://ffh-lab.com
'Class binding zlib and ZIP format ; adaptation by Jack from a Andrew Mc Millan code (see below)
'NB : to optimize code, I've just commented all unused constants, variables and functions
'***TODO : translate original comments from Jack below

'--------------------------------------------------------------------------------------------
'Auteur : Jack (Jack@developers-association.org)
'         qui a lui-même adapté le Code original : Andrew McMillan (andrew@paradoxes.info)
'Ajout du 11 juin 2006 par Jack :
'- Conversion des chaines des noms de fichiers de ISO en ASCII étendu (et réciproquement)
'     afin d'être compatible avec WinZip (clsZip_File et clsZip_UnZip)
'- Déclaration des Functions de DLL avec les types corrects (pour VB 2005)
'- Correction : Le chemin dans un Zip ne doit jamais comporter le nom d'un disque??
'Dépendances : Zlib.dll - Free (libre de droits)
'Objectifs et modif faites :
'-Regroupement dans une classe les 2 classes de Jack, (Zip et Unzip) pour faciliter l'intégration
'et la mise en oeuvre des fonctions de la librairie ZLIB.DLL, libre de droits, dans un projet.
'-J'ai ajouté un évènement ProgressLoad car l'opération de chargement d'un fichier peut être longue
'si sa taille est importante et l'utilisateur peut croire que le système est planté en ajoutant
'1 barre de progression lié à ce nouvel évènement on solutionne cet inconvénient.
'De ce fait, j'ai rebâptisé l'évènement Progress -> ProgressExtract
'En fait cet évènement n'est plus très utile depuis la fonction ybGetCentralDirEndPos qui rend le chargement instantané !
'-Le 10/07/2006 : ajout de 3 méthodes et modification de l'organisation du zippage pour permettre
'un zip de très grande taille (j'ai testé 1 zip de 1,900 Go avec succès) et améliorer les performances.
'Les méthodes ajoutées sont: createZip, ybFileAdd, WriteEndZip. Elles doivent être employées ensemble.
'Il est toujours possible d'utiliser l'autre concept, (FileAdd qui crée 1 collection de fichiers à zipper et
'WriteZip qui crée le fichier zip et écrit la collection dans le fichier zip), mais pas simultanément.
'Néanmoins il faut se souvenir que ce concept est gourmand en ressources (conso mémoire, lenteur)
'car il maintient des tableaux d'octets des fichiers à zipper pendant toute l'élaboration de la collection,
'ce qui limite la taille du fichier zip final à la saturation de la mémoire. Utilisable plutôt ponctuellement.
'Mon concept est de créer et ouvrir un fichier Zip par CreateZip dans lequel on ira ajouter
'au fur et à mesure le LocalFileHeader de chaque fichier à zipper, ce qui permet de vider la
'mémoire au fur et à mesure des compressions réalisées et ne limite donc plus la taille du fichier
'final et améliore considérablement les performances. Ceci est réalisé par ybFileAdd. A la fin de
'l'élaboration de la collection des fichiers à zipper, on écrit le CentralFileHeader, le commentaire
'la signature et colCentralDirEnd avec WriteEndZip.
'-Le 14/06/2006 : gestion compression et décompression de fichiers de taille 0 octet.
'-Le 11/07/2006 : transformation de CompressBytes sub en fonction pour qu'elle renvoie 0 ou 1 code erreur
'pour gestion possible dans la méthode appelante ou créer 1 évènement afin que le programme utilisant
'cette classe puisse gérer l'erreur. Test de l'erreur Dépassement de mémoire pour les gros fichiers
'à zipper et éviter 1 plantage. Dans ce cas, pour l'instant empacter le gros fichier dans le zip final
'avec le calcul de son crc32 hors Zlib,tout en renvoyant 1 message d'info adapté à l'utilisateur.
'Par contre lenteur liée à taille du fichier ! et double lecture du fichier source (déjà gros), 1 calcul
'du crc, l'autre empactage. Il faudra réduire à 1 seule lecture + tard lorsqu'on compressera ces gros fichiers !
'-Le 30/07/2006: ajout d'une méthode FileAddIntoZip qui permet d'ajouter 1 fichier à zipper dans 1 zip
'déjà existant.
'-Le 01/08/2006: ajout d'une méthode FileRemoveFromZip qui permet de supprimer 1 fichier d'un zip existant.
'-du 08/08/2006 au 12/11/2006: transformation de l'objet ZipFile et sa collection en 1 tableau typé
'pour fusionner les 2 classes clsZip et Zipfile (clarté du code, faciliter intégration dans 1 projet)
'et optimiser l'utilisation mémoire, et donc la vitesse. Regroupement également des champs communs
'de localFileHeader et CentralFileHeader dans 1 type commun pour économiser mémoire.
'-Le 30/10/2006 : adaptation de la mise à jour de Jack (version du 18/09/2006) sur la gestion des accents.
'-du 13/11/2006 au 15/11/2006: Remplacement des appels FSO par les API FindFirstFile et FindNextFile
'pour éliminer dépendances et faciliter l'intégration de la classe dans un projet, améliorer vitesse.
'-Le 23/11/2006 : _ajout de possibilités d'interface avec l'utilisateur sur l'évènement ZipError.
'                 _ajout de possibilités de neutraliser le Crc32 dans les méthodes de zippage pour
'                  améliorer la vitesse, mais cette option peut conduire le zip à être incompatible
'                  avec certains logiciels de décompression, les méthodes concernées sont donc paramétrées
'                  par défaut avec calcul du CRC32. Pour éliminer le calcul du crc il faut placer excludeCRC32 à True.
'-Le 05/12/2006 : gestion du niveau de compression, initialisé au taux par défaut, (niv 6) (ajout propriété LevelCompression)
'-Semaine 03/2007: -ajout d'1 méthode ExtractSgFileToTmp pour décompresser au + vite dans le répertoire Temp de Windows
'                  afin d'obtenir rapidemment des infos sur le fichier compressé (genre version, dates, commentaires ...)
'                  et préparation à un chargement "à la volée" du fichier (sans décompression apparente !...)
'                  -à cette occasion mise en place d'1 option "KeepModDate" sur les méthodes de décompression
'                  pour choisir de décompresser le ou les fichiers en conservant leur date de dernière modif.
'                  (peut-être important lorsqu'il y a des tests de dates par des applis tiers sur les fichiers !)
'-Semaine 06/2007 : -ajout méthode pour lancer directement une fichier dont le type est connu sans le dézipper
'                   (dans la réalité il y a quand-même décompression mais dans le rép Temp et la suppression est gérée
'                    c'est don transparent pour l'ulisateur.) La méthode sous-entend d'avoir exécuté ExtractSgFileToTemp au préalable!
'--------------------------------------------------------------------------------------------
'Inconvénients à l'origine : cette classe ne peut pas gèrer 1 fichier zip final de grande taille
'(en zippage ou dézippage)**, ainsi que zipper-dézipper des gros fichiers (*W)
'** dézippage de grandes tailles devenue possible par ybGetCentralDirEndPos, (précédemment limité à #200Mo)
'** zippage de grande taille éliminé le 10/07/2006 par nouveau concept YB.
'(*W) J'ai utilisé 1 parade pour le zippage d'1 gros fichier, en fait j'empacte ce fichier mais sans le
'zipper, forcément c'est long, d'autant que je contrôle quand même son Crc32 pour compatibilité avec
'définition d'un fichier Zip et donc accès du Zip pour logiciels externes, mais au moins c'est empacté
'et il n'y a donc plus de limite dans aucun sens ni fichier zip final de grande taille, ni fichier de grande
'taille dans 1 fichier zip. La seule contrainte qui reste c'est 1 contrainte système, la FAT32 !
'Ceci reste à améliorer, il faudrait pouvoir compresser et contrôler le crc32 rapidement !!!!
'
'---------------------------------------------------------------------------------------------
'/////////////Méthodes :
'             ++++zippage
'FileAdd(ByVal FilePath As String, _
        ByVal PathChoice As enumPathChoice, _
        Optional ByVal RecurseDirectories As Boolean = False, _
        Optional excludeCRC32 As Boolean = False) As Long  :  ajout de fichier au tableau de fichiers à ZIPPER
'FileAddSingle(ByVal FilePath As String, _
        ByVal PathChoice As enumPathChoice, _
        Optional excludeCRC32 As Boolean = False) As Long  :  Ajout d'un seul fichier à la collection à Zipper ?
'FileDel(ByVal Index As Long) As Boolean    : Supprime un des fichiers de la collection (avant création du zip)
'WriteZip(FilePath As String, _
        Optional Overwrite As Boolean) As Boolean   : Créé le fichier ZIP lui même à partir de la collection
'                 qui en fait utilise la classe ZipFile de A. McMillan qui réalise la compression
' (A) ces méthodes doivent s'employer ensemble mais pas avec les anciennes qui précèdent
' (A) CreateZip(FilePath As String, Optional Overwrite As Boolean) As Long  :  Crée le fichier Zip avec gestion overwrite
' (A) ybFileAdd(ByVal FilePath As String, ByVal PathChoice As enumPathChoice, lgNoFileZip As Long, _
        Optional ByVal RecurseDirectories As Boolean = False, Optional excludeCRC32 As Boolean = False) As Long : ajoute le
'        fichier à zipper au fichier zip ouvert dont le canal est indiqué en lgNoFileZip.
' (A) WriteEndZip(lgNoFileZip As Long) As Boolean : termine l'écriture du fichier zip ouvert dont le canal est indiqué en lgNoFileZip.
'FileAddIntoZip(strZipPath As String, strFilePath As String, PathChoice As enumPathChoice, _
        Optional excludeCRC32 As Boolean = False, Optional OverWrite As Boolean = False) As Boolean : permet d'ajouter 1 fichier
'               à zipper dans 1 zip déjà existant avec Option OverWrite, par défaut à False, si OverWrite par de gestion de date, overwrite brut.
'FileRemoveFromZip(strZipPath As String, strFilePath As String) As Boolean : permet de supprimer 1 zip d'un zip existant.
'               ++++informations/dézippage
'ZipOpen(ByVal ZipPath As String) As Boolean        : ouvre le fichier zip (le charge pour analyse en lecture)
'ExtractAllFiles(ByVal FolderPath As String, _
        Optional ByVal PreservePath As Boolean, _
        Optional ByVal Overwrite As Boolean, _
        Optional ByVal KeepModDate As Boolean) As Boolean  : extrait tous les fichiers du zip
'ExtractSingleFile(ByVal FileNumber As Long, _
        ByVal FolderPath As String, _
        Optional ByVal PreservePath As Boolean, _
        Optional ByVal Overwrite As Boolean, _
        Optional ByVal KeepModDate As Boolean) As Boolean  : extrait un fichier d'après son n° dans le catalogue du zip
'ExtractSgFileToTmp(lgNuFile As Long) As String            : extrait un fichier dans le Temp de Windows en conservant sa date de modif
'                                                            (voir commentaires dans la fonction)
'LoadZippedFile(strFilePath As String) as Boolean          : lance l'exécutable attaché au fichier apparemment sans le dézipper.
'                                                            (voir commentaires dans la fonction)
'
'////////////Propriétés :
'FileCount() As Long                        : nombre de fichiers inclus dans le tableau-collection à zipper
'FileName(ByVal Index As Long) As String    : nom du fichier correspondant à l'index dans la collection à zipper
'Comment() As String (lecture/écriture)   : commentaire du Zip en lecture/écriture avant création du fichier
'SizeCompressed(ByVal Index As Long) As Long: renvoie taille du fichier compressé dans la collection
'SizeUncompressed(ByVal Index As Long) As Long :Renvoie la taille décompressée (originale) du fichier
'inFileCount() As Long    : Renvoie le nombre de fichiers contenus dans le fichier Zip en excluant les répertoires
'                           pas la récursivité mais le nombre de répertoires
'ZipIsOpen() As Boolean   : Renvoie True si un Zip est déjà ouvert
'ZipComment() As String   : Renvoie le commentaire du Zip
'inFileName(ByVal Index As Long) As String : Renvoie le nom du fichier désigné par Index
'FileCompressedSize(ByVal Index As Long) As Long : Renvoie la taille du fichier compressé désigné par Index
'FileUncompressedSize(ByVal Index As Long) As Long :Renvoie la taille originale du fichier désigné par Index
'FileDateAndTime(ByVal Index As Long) As Date : Renvoie la date et l'heure du fichier original désigné par Index
'ZLibVersion() As String  : Renvoie la version de la Zlib.DLL
'LevelCompression() as eLevelCompres : fixe ou renvoie le niveau de compression qui est initialisé au taux par défaut
''
'///////////Evènements :
'ProgressExtract(ByVal Percent As Long, ByRef Cancel As Boolean) : renvoie % de progression (décompress)
'Status(ByVal Text As String)                             : renvoie info déroulement
'ZipError(ByVal Number As eZipError, ByVal Description As String, Cancel as boolean): renvoie une erreur
'                                       dont la gestion peut être contrôlée par Cancel. (information/dézippage)
'FileZipErr(strFilePath as string, lgErr as long, strDescription as string) : erreur pendant zippage
'
'Enumérations :
'enumPathChoice     : options sur Path
'eZipError          : erreurs possibles gérées (information/dézippage)
'eLevelCompres      : les taux de compression
'
'------------------------------------------------------------------------------------------------
'Fonctionnement : ++Zippage: création d'un tableau (précédemment collection) de fichiers à compresser et intégrer
'                 au fichier Zip définitif. Le tableau peut être manipulé (ajouter ou supprimer ou ... des fichiers)
'                 tant que le fichier zip définitif n'est pas créé.
'                 +Autre méthode (ajout YB 07/2006): Création du fichier zip dès le départ, ajout des données compressées
'                 au fur et à mesure dans le fichier zip, écriture à la fin du process de la structure et fermeture du fichier zip.
'                 ++DéZippage : charge le fichier en mémoire par openFile, recherche le catalogue des fichiers
'                 inclus dans le zip (structure du fichier), puis manipule les fichiers d'après leur index
'                 auquel correspond 1 offset dans le zip.
'-----------------------------------------------------------------------------------------------
'Améliorations envisagées :
'  -Gérer compression et calcul crc32 de gros fichiers à zipper (actuellement empactés mais non compressés)
'   (pour y parvenir, il faudrait changer l'entrée-srotie de Compress et Uncompress de Zlib qui sont
'    actuellement des tableaux d'octets ce qui conduit à des dépassements de mémoire lors de gros fichiers,
'    en étudiant + les possibilités de Zlib, il y a peut-être la réponse à ce problème, ou alors
'    gérer la compression hors de zlib, mais alors le crc des gros fichiers étant déjà externe à zlib
'    quid de passer par zlib !
'  -Utiliser les API ReadFile et WriteFile pour optimiser performances et réduire les limites de tailles !
'  -Crypter/décrypter le fichier ZIP  (voir définition fichier zip, éventuelles fonctions dans Zlib version + récente)
'--------------------------------------------------------------------------------------------------

Option Explicit

'API lecture/écriture Fichiers  ReadFile et WriteFile   ---------------------------------------------
Const FILE_BEGIN = 0
Const FILE_SHARE_READ = &H1
Const FILE_SHARE_WRITE = &H2
'Const CREATE_NEW = 1
'Const CREATE_ALWAYS = 2
Const OPEN_EXISTING = 3
Const GENERIC_READ = &H80000000
Const GENERIC_WRITE = &H40000000
'Const FILE_END = 2
'Const FILE_CURRENT = 1
Const INVALID_HANDLE_VALUE As Long = -1
Const ERROR As Long = &HFFFFFFFF

Private Declare Function ReadFile Lib "kernel32" _
(ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Any) As Long
'Private Declare Function WriteFile Lib "kernel32" _
(ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Any) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" _
(ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function SetFilePointer Lib "kernel32" _
(ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Function GetFileSize Lib "kernel32" _
(ByVal hFile As Long, lpFileSizeHigh As Long) As Long

' déclarations autres API   ------------------------------------------------------------------------------
'Private Declare Function CopyMemory_Long_Byte Lib "kernel32" Alias "RtlMoveMemory" (dest As Long, Src As Byte, ByVal Length As Long) As Long
Private Declare Function CopyMemory_Byte_Byte Lib "kernel32" Alias "RtlMoveMemory" (dest As Byte, Src As Byte, ByVal Length As Long) As Long
Private Declare Sub CopyMemory_Int_Byte Lib "kernel32" Alias "RtlMoveMemory" (dest As Integer, Source As Byte, ByVal Length As Long)
Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpBuffer As String, ByVal lpString As Long) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Long) As Long
' API pour Récupérer le chemin du répertoire temporaire ---------------------------------------------------
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
'API pour gérer les dates du type FileTime
'FileTimeToSystemTime avec Api Infos et recherche Fichiers
'Private Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long                                                   '
Private Declare Function LocalFileTimeToFileTime Lib "kernel32" (lpLocalFileTime As FILETIME, lpFileTime As FILETIME) As Long
Private Declare Function SetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As Any, lpLastAccessTime As Any, lpLastWriteTime As Any) As Long
Private Declare Function SystemTimeToFileTime Lib "kernel32" (lpSystemTime As SYSTEMTIME, lpFileTime As FILETIME) As Long
' API Noms de fichiers Longs vers Petits ---------------------------------------------------------------------------------------
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, _
    ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
' déclarations API Infos et recherche Fichiers -------------------------------------------------------------
Private Declare Function DosDateTimeToFileTime Lib "kernel32.dll" (ByVal wFatDate As Long, ByVal wFatTime As Long, lpFileTime As FILETIME) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Declare Function GetFullPathName Lib "kernel32" Alias "GetFullPathNameA" (ByVal lpFileName As String, ByVal nBufferLength As Long, _
            ByVal lpBuffer As String, ByVal lpFilePart As String) As Long
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

' API et constantes pour lancer appli à partir d'un fichier ------------------------------------------------
Private Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, _
    ByVal lpResult As String) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'Const NORMAL_PRIORITY_CLASS = &H20&
'Const REALTIME_PRIORITY_CLASS = &H100
'Const INFINITE = -1&
'Const SW_SHOWMAXIMIZED = 3
'Const SW_SHOWMINIMIZED = 2
'Const SW_SHOWNORMAL = 1
'Const STARTF_USESHOWWINDOW = &H1

' API, Type Infos mémoire -------------------------------------------------------------------------------------------
Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)
Private Type MEMORYSTATUS
    dwLength As Long            'la taille de la structure.
    dwMemoryLoad As Long        'le % de mémoire actuellement utilisé.
    dwTotalPhys As Long         'la taille totale de la mémoire physique (en octets).
    dwAvailPhys As Long         'la taille de la mémoire physique disponible (en octets).
    dwTotalPageFile As Long     'la taille totale que peut atteindre le fichier d'échange (en octets).
    dwAvailPageFile As Long     'l'espace disponible dans le fichier d'échange (en octets).
    dwTotalVirtual As Long      'la mémoire totale pouvant être utilisée par l'application courante (en octets).
    dwAvailVirtual As Long      'l'espace libre totale (en octets).
End Type

'API mise à jour Bdr -------------------------------------------------------------------------------------------------
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" _
    (ByVal hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" _
    (ByVal hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, _
     lpData As Any, ByVal cbData As Long) As Long

' déclarations Zlib  ----------------------------------------------------------------------------------------
Private Declare Function Compress_Byte_Long_Byte Lib "zlib.dll" Alias "compress" (dest As Any, destLen As Long, Src As Byte, ByVal srcLen As Long) As Long
Private Declare Function Crc32_Byte Lib "zlib.dll" Alias "crc32" (ByVal crc As Long, buffer As Byte, ByVal Length As Long) As Long
Private Declare Function ZLibVer Lib "zlib" Alias "zlibVersion" () As Long
Private Declare Function UnCompress_Byte_Long_Byte Lib "zlib.dll" Alias "uncompress" (dest As Byte, destLen As Long, Src As Byte, ByVal srcLen As Long) As Long
Private Declare Function lCRC32_Byte Lib "zlib.dll" Alias "crc32" (ByVal crc As Long, buffer As Byte, ByVal Length As Long) As Long
' autres déclarations pour Zlib  (pour zStream)
Private Declare Function deflate Lib "zlib.dll" (vStream As zStream, ByVal flush As Long) As Long
Private Declare Function deflateEnd Lib "zlib.dll" (vStream As zStream) As Long
Private Declare Function deflateInit2 Lib "zlib.dll" Alias "deflateInit2_" (vStream As zStream, ByVal Level As Long, _
        ByVal method As Long, ByVal windowBits As Long, ByVal memLevel As Long, _
        ByVal strategy As Long, ByVal Version As String, ByVal stream_size As Long) As Long
'Private Declare Function inflate Lib "zlib.dll" (vStream As zStream, ByVal vflush As Long) As Long
'Private Declare Function inflateEnd Lib "zlib.dll" (vStream As zStream) As Long
'Private Declare Function inflateInit2 Lib "zlib.dll" Alias "inflateInit2_" (vStream As zStream, _
        ByVal vWindowBits As Long, ByVal vVersion As String, ByVal vLen As Long) As Long
''compress2 ??

' Types et variables    ------------------------------------------------------------------
Private Type zStream
    next_in     As Long
    avail_in    As Long
    total_in    As Long
    next_out    As Long
    avail_out   As Long
    total_out   As Long
    Msg         As Long
    state       As Long
    zalloc      As Long
    zfree       As Long
    opaque      As Long
    data_type   As Long
    adler       As Long
    Reserved    As Long
End Type

Private Type typCenteralDirEnd
    EndOFCentralDirSignature            As Long
    NumberOfThisDisk                    As Integer
    NumberOfDiskWithCentralDir          As Integer
    EntriesInTheCentralDirThisOnDisk    As Integer
    EntriesInTheCentralDir              As Integer
    SizeOfCentralDir                    As Long
    OffSetOfCentralDir                  As Long
    ZipFileCommentLength                As Integer
End Type

Private Type typComFileHeader
    GeneralPurposeBitFlag           As Integer
    CompressionMethod               As Integer
    LastModFileTime                 As Integer
    LastModFileDate                 As Integer
    CRC32                           As Long
    CompressedSize                  As Long
    UncompressedSize                As Long
    FileNameLength                  As Integer
    ExtraFieldLength                As Integer
End Type

Private Type typCentralFileHeaderA
    CentralFileHeaderSigniature     As Long
    VersionMadeBy                   As Integer
    VersionNeededToExtract          As Integer
End Type
Private Type typCentralFileHeaderC
    FileCommentLength               As Integer
    DiskNumberStart                 As Integer
    InternalFileAttributes          As Integer
    ExternalFileAttributes          As Long
    RelativeOffsetOfLocalHeader     As Long
End Type

Private Type typLocalFileHeader
    LocalFileHeaderSignature      As Long
    VersionNeededToExtract        As Integer
End Type

Private Type typFileInfos
    FileName            As String
    CompressedSize      As Long
    UncompressedSize    As Long
    LastModFileDate     As Long
    LastModFileTime     As Long
End Type

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private Const MAX_PATH = 260
'Private Const FILE_ATTRIBUTE_READONLY = &H1
'Private Const FILE_ATTRIBUTE_HIDDEN = &H2
'Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
'Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
'Private Const FILE_ATTRIBUTE_NORMAL = &H80
'Private Const FILE_ATTRIBUTE_TEMPORARY = &H100
'Private Const FILE_ATTRIBUTE_COMPRESSED = &H800
'Private Const MAXDWORD = &HFFFF

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type

Private Const EndOFCentralDirSignature      As Long = &H6054B50
Private Const CentralFileHeaderSigniature   As Long = &H2014B50
Private Const LocalFileHeaderSignature      As Long = &H4034B50
'Private Const ExtLocSig                     As Long = &H8074B50

'Autres constantes Zlib pour Inflate/deflate---
Private Const ZLIB_VERSION           As String = "1.1.4"
'----------------------------------------------
'Private Const Z_NO_FLUSH             As Long = 0
'Private Const Z_PARTIAL_FLUSH        As Long = 1
'Private Const Z_SYNC_FLUSH           As Long = 2
'Private Const Z_FULL_FLUSH           As Long = 3
Private Const Z_FINISH               As Long = 4
'----------------------------------------------
Private Const Z_OK                   As Long = 0
'Private Const Z_STREAM_END           As Long = 1
'Private Const Z_NEED_DICT            As Long = 2
'Private Const Z_ERRNO                As Long = (-1)
'Private Const Z_STREAM_ERROR         As Long = (-2)
'Private Const Z_DATA_ERROR           As Long = (-3)
'Private Const Z_MEM_ERROR            As Long = (-4)
'Private Const Z_BUF_ERROR            As Long = (-5)
'Private Const Z_VERSION_ERROR        As Long = (-6)
'----------------------------------------------
Private Const Z_DEFAULT_STRATEGY     As Long = 0
'Private Const Z_FILTERED             As Long = 1
'Private Const Z_HUFFMAN_ONLY         As Long = 2
'----------------------------------------------
Private Const Z_BINARY               As Long = 0
'Private Const Z_ASCII                As Long = 1
'Private Const Z_UNKNOWN              As Long = 2
'----------------------------------------------
Private Const Z_DEFLATED             As Long = 8
'Private Const Z_NULL                 As Long = 0
Private Const MAX_MEM_LEVEL          As Long = 9
Private Const MAX_WBITS              As Long = 15
'----------------------------------------------

Public Enum enumPathChoice
    WithoutPath = 0         ' Pas de mémorisation du chemin du fichier
    WithCompletePath = 1    ' Avec mémorisation du chemin du fichier
    WithRelativePath = 2    ' Avec mémorisation du chemin du fichier par
                            '   rapport au répertoire actuel
End Enum

Public Enum eZipError
    zeZLibNotInstalled = 1
    zeNotZipFile = 2
    zeNoZipOpenFile = 3
    zeUnsupportedCompressionMethod = 4
    zeChecksumError = 5
    zeFileNotFound = 10
    zeFileAlreadyExists = 11
    zeCantRemoveFile = 12
    zeCantCreateFolder = 13
    zeStructureError = 14
    zeMemory = 15
    zeUnKnownError = 16
    zeDeflate = 17
    zeNoexecutable = 18
End Enum

Public Enum eLevelCompres
    Z_NO_COMPRESSION = 0            'pas de compression
    Z_BEST_SPEED = 1                'meilleure vitesse, mais compression faible
    Z_LOW_LEVEL2 = 2
    Z_LOW_LEVEL3 = 3
    Z_LOW_LEVEL4 = 4
    Z_MIDDLE_LEVEL5 = 5
    Z_DEFAULT_COMPRESSION = (-1)    'compression standard optimisée, équivalent niveau 6.
    Z_MIDDLE_LEVEL7 = 7
    Z_HIGH_LEVEL8 = 8
    Z_BEST_COMPRESSION = 9          'meilleure compression, mais vitesse lente
End Enum

' gestion des accents du 31/10/2006
Private Enum eOEMStringTypeConversion
    [ANSI to ASCII]
    [ASCII to ANSI]
End Enum
Private sASCI As String     ' Voir leur remplissage dans Class_Initialize
Private sANSI As String
'
'variables tableau dynamique typé pour remplacer collection d'objets ZipFile, annoncé + rapide
Private Type cls_ZipFile
    commun            As typComFileHeader
    localFileHeaderZF As typLocalFileHeader
    centralFileHeaderZFA As typCentralFileHeaderA
    centralFileHeaderZFC As typCentralFileHeaderC
    FileNameZF As String
    extraFieldZF As String
    commentZF As String
    ASCIIFNZF As String
    filedataZF() As Byte
    bigFileZF  As Boolean
End Type
Private varCommunZF As typComFileHeader
Private tabFiles() As cls_ZipFile       'tableau collection Fichiers à zipper
Private cptTab As Long      'Index et compteur tableau collection Fichiers à zipper
'table pour calcul du crc32
Private arrCRC32(255) As Long
'
' variables internes
Private Fh As Long                                  ' n° d'ouverture du fichier Zip (dézippage)
Private CentralFileHeaderA  As typCentralFileHeaderA
Private CentralFileHeaderC  As typCentralFileHeaderC
Private CentralDirEnd       As typCenteralDirEnd
Private CentralDirEndPos    As Long
Private LocalFileHeader     As typLocalFileHeader
Private colCentralDirEnd    As typCenteralDirEnd    ' pour la structure des données
Private ZipFileComment      As String               ' pour le commentaire du Zip
Private mFileInfos()        As typFileInfos
Private mFileCountTemp      As Long                 ' Nb fichiers (y compris les sous-rép)
Private mFileCount          As Long                 ' Nb fichiers corrigé (valeur public)
Private mZipComment         As String
Private sFileName           As String
Private sASCII_FileName     As String
Private ExtraField          As String
'Private FileComment         As String
Private m_Nombre            As Long                 ' interne à fonction CheckInsert
Private m_LevelComp         As Long                 ' taux de compression
Private arrDelFile()       As String               ' tableau pour renseigner les fichiers déprov dans Temp à détruire au redémarrage de Windows

' Conversion ASCII étendu <--> ISO 8859-1 (de l'OS)
' Voir Conv_ISO_ASCII et Conv_ASCII_ISO
Private Table_Conversion_ASCII_ISO As Variant
' Conversion ASCII étendu <--> ISO 8859-1 (de VB)
' Voir Conv_ISO_ASCII et Conv_ASCII_ISO
Private Table_Conversion_ISO_ASCII As Variant

Public Event ProgressExtract(ByVal Percent As Long, ByRef Cancel As Boolean) 'en passant Cancel à True on stoppe l'opération
Public Event ProgressLoad(ByVal Percent As Long, ByRef Cancel As Boolean)  'en passant Cancel à True on stoppe l'opération
Public Event Status(ByVal Text As String)
Public Event ZipError(ByVal Number As eZipError, ByVal Description As String, Cancel As Boolean)  ' en passant Cancel à True on intervient sur l'opération
Public Event FileZipErr(strFilePath As String, lgErr As Long, strDescription As String)
'

Public Property Get FileCount() As Long
    
    ' Renvoie le nombre de fichiers déjà ajoutés au tableau des éléments à zipper
    On Error GoTo Erreur
    FileCount = UBound(tabFiles)
    Exit Property
Erreur:
FileCount = 0   'il y erreur s'il n'y a pas de données dans le tableau

End Property

Public Property Get Comment() As String
    ' Relit le commentaire du Zip
    Comment = ZipFileComment
End Property
Public Property Let Comment(s As String)
    ' Dépose un Commentaire au Zip
    ZipFileComment = s
    colCentralDirEnd.ZipFileCommentLength = Len(s)
End Property

Public Property Get FileName(ByVal index As Long) As String
    ' Renvoie le nom du fichier dont l'index est Index (judicieux, non ?)
    If index > UBound(tabFiles) Then Exit Property
    FileName = tabFiles(index).FileNameZF
End Property

Public Property Get SizeCompressed(ByVal index As Long) As Long
    ' Renvoie la taille compressée du fichier
    If index > UBound(tabFiles) Then Exit Property
    SizeCompressed = tabFiles(index).commun.CompressedSize
End Property

Public Property Get SizeUncompressed(ByVal index As Long) As Long
    ' Renvoie la taille décompressée (originale) du fichier
    If index > UBound(tabFiles) Then Exit Property
    SizeUncompressed = tabFiles(index).commun.UncompressedSize
End Property

Public Function FileDel(ByVal index As Long) As Boolean
    ' Supprime un des fichiers à zipper du tableau (avant création du zip)
    On Error GoTo Erreur
    Dim I As Long
    If index > UBound(tabFiles) Or index < 1 Then Exit Function

    If index < UBound(tabFiles) Then   ' si la suppression n'est pas le dernier enregistrement
        For I = index + 1 To UBound(tabFiles)
            'transfert tous les éléments qui suivent la suppression à partir de sa place dans le tableau
            tabFiles(I - 1).commun = tabFiles(I).commun
            tabFiles(I - 1).localFileHeaderZF = tabFiles(I).localFileHeaderZF
            tabFiles(I - 1).centralFileHeaderZFA = tabFiles(I).centralFileHeaderZFA
            tabFiles(I - 1).centralFileHeaderZFC = tabFiles(I).centralFileHeaderZFC
            tabFiles(I - 1).FileNameZF = tabFiles(I).FileNameZF
            tabFiles(I - 1).extraFieldZF = tabFiles(I).extraFieldZF
            tabFiles(I - 1).commentZF = tabFiles(I).commentZF
            tabFiles(I - 1).ASCIIFNZF = tabFiles(I).ASCIIFNZF
            tabFiles(I - 1).filedataZF = tabFiles(I).filedataZF
            tabFiles(I - 1).bigFileZF = tabFiles(I).bigFileZF
        Next
    End If
    'supprime le dernier enregistrement du tableau (dans tous les cas)
    If UBound(tabFiles) > 1 Then    'si on a encore + d'1 enregistremt ds le tableau
        ReDim Preserve tabFiles(1 To UBound(tabFiles) - 1)
    Else    'si on est le dernier enregist du tableau
        Erase tabFiles  'on l'efface directement
    End If
    cptTab = cptTab - 1 'maintenance du compteur de tabFiles
    FileDel = True
    Exit Function
Erreur:
    FileDel = False
    
End Function

Public Function FileAdd(ByVal FilePath As String, _
                        ByVal PathChoice As enumPathChoice, _
                        Optional ByVal RecurseDirectories As Boolean = False, _
                        Optional excludeCRC32 As Boolean = False) As Long
    ' Ajoute un ou des fichiers à la collection/tableau
    
    ' FilePath           : Fichier avec son chemin complet
    ' PathChoice         : Définit si on veut garder le chemin du fichier
    ' RecurseDirectories : Définit s'il faut chercher les fichiers dans les sous-répertoires
    
    Dim NbFichiers As Long
    Dim StrFilter As String         'variable filtre de FindFirstFile ...
    ' Par défaut, pas d'ajout de fichier
    FileAdd = 0
    
    ' Récupère le nom réel complet(au cas où on utilise des \..\
    '                           ou bien si on ne précise pas de chemin
    'par api
    Dim buffer As String, Ret As Long
    buffer = Space(IIf(Len(FilePath) < 255, 255, Len(FilePath) + 1)) 'parfois des noms de fichiers provenance Internet sont > 255 de windows
    Ret = GetFullPathName(FilePath, Len(buffer), buffer, "")
    FilePath = Left(buffer, Ret)
    
    If InStr(1, FilePath, "*") + InStr(1, FilePath, "?") = 0 Then ' il s'agit d'1 seul fichier à zipper
        'si le fichier existe bien
        If FichierExiste(FilePath) Then
            ' On l'insère parmi les fichiers à zipper
            NbFichiers = NbFichiers + FileAddSingle(FilePath, PathChoice, excludeCRC32)
        Else
            Exit Function
        End If
    Else    ' détection des fichiers concernés
        'fonction checkInsert : recherche les fichiers correspondants au filtre, avec gestion
        'option récursif ou pas, et ajoute les fichiers trouvés au tableau-collection à zipper.
        StrFilter = GetFileName(FilePath) 'extrait le nom du fichier et son extension avec ou pas des caractères génériques
                                          'pour servir de filtre aux Api FindFirstFile ...
        NbFichiers = NbFichiers + CheckInsert(Replace(FilePath, StrFilter, ""), StrFilter, 0, PathChoice, RecurseDirectories, excludeCRC32)
    End If  'traitement orienté 1 fichier ou plusieurs avec caractères génériques
    
    ' Retourne le nombre de fichiers ajoutés
    FileAdd = NbFichiers
    
End Function

Private Function FileAddSingle(ByVal FilePath As String, _
                               ByVal PathChoice As enumPathChoice, _
                               Optional excludeCRC32 As Boolean = False) As Long

    ' C'était la fonction originale (ou presque de l'auteur)
    ' On l'utilise pour ajouter un (seul) fichier au tableau
    ' (Voir FileAdd)
    
    FileAddSingle = 0
' test éliminé pour éliminer redondance test existance, en effet, sauf si FileAddSingle est appelée
' seule, l'existance du fichier a déjà été testée dans les procédures ou fonctions qui la précédent,
' d'autre part si cette fonction est appelée en autonome, l'absence éventuelle du fichier sera traitée dans
' la gestion d'erreurs. Evite donc éventuel test redondant !
'''    ' On ressort si le fichier n'existe pas
'''    If FichierExiste(FilePath) = False Then
'''        Exit Function
'''    End If
    
    Dim lgTab As Long, cZipError As Long, d As Date
    On Error GoTo errFAS
    d = FileDateTime(FilePath)
    'pour économiser des ressources ajoute 1 enregistrement au tableau pour stocker le tableau
    'd'octets comprimés directement dans le tableau typé. Si lors du contrôle du compress ou d'1
    'erreur, cet enregistrement devait être supprimé il suffira de faire 1 redim - 1 du tableau
    'et du compteur pour le supprimer, mais on évite de transférer les données compressées dans
    '1 tableau d'octets qu'il faudra copier ensuite dans ce tableau typé et donc doubler l'utilisation
    'mémoire avant de supprimer les données de ce tableau intermédiaire et faire 1 copie de plus !
    lgTab = UBound(tabFiles) + 1
    cptTab = lgTab  'place le cpteur sur 1 nouvel enregistrement
    ReDim Preserve tabFiles(1 To lgTab)
    
    ' Par défaut, le nom du fichier est zippé sans le chemin
    '   mais il faut (ici) fournir le nom avec le chemin
    ' Par défaut, on ne garde que le nom du fichier (pas le path)
    sFileName = GetFileName(FilePath)
    sASCII_FileName = OEMConvert(sFileName, [ANSI to ASCII])
    With varCommunZF
        .FileNameLength = Len(sFileName)
        .LastModFileDate = GetDOSDate(d)
        .LastModFileTime = GetDOSTime(d)
        
            'lance la compression et stocke éventuelles erreurs qui seront gérées en amont
            cZipError = CompressBytes(FilePath, tabFiles(cptTab).filedataZF, .CompressedSize, .UncompressedSize, excludeCRC32)
    ' Si on veut garder le chemin, on le dit ici
    Select Case PathChoice
        Case WithoutPath        ' Sans le chemin
            ' rien à faire, c'est comme ça par défaut
        Case WithCompletePath   ' Avec le chemin complet
            sFileName = FilePath
            sASCII_FileName = OEMConvert(sFileName, [ANSI to ASCII])
            .FileNameLength = Len(sFileName)
        Case WithRelativePath   ' Avec un chemin relatif
            ' ... relatif au chemin actuel
            If Left(UCase(FilePath), Len(CurDir)) = UCase(CurDir) Then
                ' Si le chemin actuel est dans le nom du fichier, Ok
                sFileName = Mid(FilePath, Len(CurDir) + 2)
                sASCII_FileName = OEMConvert(sFileName, [ANSI to ASCII])
                .FileNameLength = Len(sFileName)
            Else
                ' Sinon, on met le chemin complet
                sFileName = FilePath
                sASCII_FileName = OEMConvert(sFileName, [ANSI to ASCII])
                .FileNameLength = Len(sFileName)
            End If
    End Select
    End With

    'controle qu 'il n'y a pas d'erreur et complète l'ajout de l'enregistrement du fichier au tableau
    'ou le supprimer en cas d'erreur
    Select Case cZipError
        Case 0, 7     ' c'est OK on complète les données au tableau
            With tabFiles(cptTab)
                .commun = varCommunZF
                .localFileHeaderZF = LocalFileHeader
                .centralFileHeaderZFA = CentralFileHeaderA
                .centralFileHeaderZFC = CentralFileHeaderC
                .FileNameZF = sFileName
                .ASCIIFNZF = sASCII_FileName
                If cZipError = 7 Then
                    .bigFileZF = True
                    RaiseEvent FileZipErr(FilePath, 7, "Fichier trop important pour être compressé, le fichier a été empacté dans le fichier final mais sans compression !")
                End If
            End With
            ' Signale que l'ajout c'est bien passé (tout est dans le tableau)
            FileAddSingle = 1
        Case Else   'erreur qui entraîne retrait de l'enregistrement ajouté au tableau
            'on n'ajoute pas l'enregistrement au tableau, on replace l'index à la bonne valeur
            cptTab = cptTab - 1
            'supprime l'enregistrement ajouté au tableau
            ReDim Preserve tabFiles(1 To cptTab)
            'par défaut FileAddSingle est à 0 donc on le laisse
            'on gère le message d'erreur en fonction de son n°
            If cZipError = 70 Then 'permission refusée (violation de partage)
                RaiseEvent FileZipErr(FilePath, 70, "Violation de Partage, le fichier a été exclus du traitement")
            Else 'autre
                RaiseEvent FileZipErr(FilePath, cZipError, "Erreur non testée, le fichier a été exclus du traitement")
            End If
    End Select
    On Error GoTo 0
    
Exit Function
errFAS:
If Err.Number = 9 Then 'au cas où le tableau ne serait pas encore initialisé on aurait indice hors plage
    lgTab = 1
    Resume Next
Else
    'on n'ajoute pas l'enregistrement au tableau, on replace l'index à la bonne valeur
    cptTab = cptTab - 1
    'supprime l'enregistrement ajouté au tableau
    If cptTab > 0 Then
        ReDim Preserve tabFiles(1 To cptTab)
    Else
        ReDim tabFiles(0)
    End If
    'par défaut FileAddSingle est à 0 donc on le laisse
   RaiseEvent FileZipErr(FilePath, Err.Number, "Erreur non testée, le fichier a été exclus du traitement")
   Exit Function
End If

End Function

Public Function WriteZip(FilePath As String, _
                         Optional OverWrite As Boolean) As Boolean

    ' Crée le fichier ZIP lui même à partir du tableau élaboré
    ' FilePath  : Nom du fichier ZIP avec son chemin complet
    ' Overwrite : On l'écrase s'il existe déjà ?

    Dim l As Long
    Dim Fh As Long
   Dim I As Long
    ' Si le Zip existe déjà, on le détruit (si Overwrite est à True)
    If FichierExiste(FilePath, False) Then
        If OverWrite Then
            Kill FilePath
        Else
            RaiseEvent FileZipErr(FilePath, 11, "Création -No Overwrite- du Zip abandonnée car le fichier existe déjà")
            Erase tabFiles
            cptTab = 0
            WriteZip = False
            Exit Function
        End If
    End If
    
    ' Ouvre le Zip et enregistre chacune des structures des fichiers du tableau
    Fh = FreeFile
    Open FilePath For Binary As #Fh
    For I = 1 To UBound(tabFiles)
        WriteLocalFileHeader Fh, I
    Next
    l = Seek(Fh)
    For I = 1 To UBound(tabFiles)
        WriteCentralFileHeader Fh, I
    Next
    With colCentralDirEnd
        .EndOFCentralDirSignature = &H6054B50
        .EntriesInTheCentralDirThisOnDisk = UBound(tabFiles)
        .EntriesInTheCentralDir = .EntriesInTheCentralDirThisOnDisk
        .SizeOfCentralDir = Seek(Fh) - l
        .OffSetOfCentralDir = l - 1
    End With
    Put #Fh, , colCentralDirEnd    ' Ajoute la description de structure
    Put #Fh, , ZipFileComment   ' Ajoute le commentaire général
    Close #Fh
    Erase tabFiles
    cptTab = 0

    ' Signale que la création c'est bien passée
    WriteZip = True
   
End Function

Public Function ZipOpen(ByVal ZipPath As String) As Boolean
    Dim Cancel As Boolean

    RaiseEvent Status("Ouverture du Zip")
    ZipClose

    If FichierExiste(ZipPath) = False Then
        RaiseEvent ZipError(zeFileNotFound, "Le fichier " & ZipPath & " n'existe pas", False)
        Exit Function
    End If
    
    ' Ouverture du fichier Zip en binaire pour lecture indexée
    Fh = FreeFile
    Open ZipPath For Binary As #Fh
    DoEvents
    
    ' Récupère l'offset de la structure descriptive du Zip
'''    CentralDirEndPos = zGetCentralDirEndPos(fh) 'cette opération peut être longue selon la taille
    CentralDirEndPos = ybGetCentralDirEndPos(ZipPath)
    'j'y ai ajouté renvoi de la progression 100%=95% de ZipOpen
    
    If CentralDirEndPos > 0 Then
        ZipOpen = True
        RaiseEvent Status("Zip ouvert")
    ElseIf CentralDirEndPos < 0 Then  'teste qu'on a annulé le chargement lors de zGetCentralDirEndPos
        RaiseEvent Status("Chargement annulé par l'utilisateur")
        Exit Function
    Else
        RaiseEvent ZipError(zeNotZipFile, "Le fichier " & ZipPath & " n'est pas un fichier Zip", False)
        Exit Function
    End If
    
    '---------- Récupère les infos concernant les fichiers inclus dans le Zip
    ' NB : Quand on interroge le fichier ZIP, il renvoie un nombre de fichiers
    ' Or, parmi ces fichiers, il y a une ligne pour chaque sous répertoire créé
    '   et ne faut pas les comptabiliser dans le nombre (UncompressedSize = 0).
    ' L'appel de zReadCentralDirEnd va renvoyer un nombre erroné que nous allons
    '   corriger en scrutant tout le contenu du zip.
    ' --> Le maintient de cette partie de code est donc primordiale pour garder
    '     un chiffre correct !

    ' Vérifie que la DLL est bien là (et répond)
    If Len(ZLibVersion) = 0 Then
        RaiseEvent ZipError(zeZLibNotInstalled, "La librairie Zlib n'est pas installée", False)
        Exit Function
    End If
    Dim r As Long, Temp As String, Ret As Boolean
    ' Récupère l'emplacement de la structure descriptive du Zip
    '   (dont le nombre de fichiers et le commentaire du Zip)
    Call zReadCentralDirEnd(CentralDirEndPos)
    ' Prépare le tableau d'accueil des données
    ReDim mFileInfos(mFileCountTemp) As typFileInfos
    ' Se positionne au début de la structure d'infos du contenu
    Seek #Fh, CentralDirEnd.OffSetOfCentralDir + 1
    ' Va querir les données (y cause bien, hein ?)
    r = 1
    Do While r <= mFileCountTemp
        ' Cherche les infos les une après les autres (important)
        Ret = zReadCentralFileHeader(Temp)
        ' On ne tient pas compte des sous-répertoires identifiés par taille à 0
'''        Do While Ret And (varCommunZF.UncompressedSize = 0)  ' FAUX ==> un fichier peut être à 0
        Do While Ret And (CentralFileHeaderC.InternalFileAttributes = FILE_ATTRIBUTE_DIRECTORY)
            ' Si on vient de tomber sur un sous-répertoire, on
            '  recommence la lecture
            mFileCountTemp = mFileCountTemp - 1
            Ret = zReadCentralFileHeader(Temp)
            DoEvents
        Loop
        If Ret Then ' si on n'est pas arrivé à la fin des données
            ' et les mémorise
            mFileInfos(r).FileName = Temp
            mFileInfos(r).CompressedSize = varCommunZF.CompressedSize
            mFileInfos(r).UncompressedSize = varCommunZF.UncompressedSize
            mFileInfos(r).LastModFileDate = varCommunZF.LastModFileDate
            mFileInfos(r).LastModFileTime = varCommunZF.LastModFileTime
        End If
        r = r + 1
        DoEvents
    Loop
    ' Rafraichit le nombre réel de fichiers
    mFileCount = mFileCountTemp
    'renvoie l'évènement
    RaiseEvent ProgressLoad(100, Cancel) 'ici cancel n'est pas géré car il ne sert + !
    DoEvents
End Function

Public Sub ZipClose()
    If Fh <> 0 Then
        Close #Fh
        Fh = 0
        RaiseEvent Status("Zip fermé")
    End If
    CentralDirEndPos = 0
    
End Sub

Public Function ExtractSingleFile(ByVal FileNumber As Long, _
                                  ByVal FolderPath As String, _
                                  Optional ByVal PreservePath As Boolean, _
                                  Optional ByVal OverWrite As Boolean, _
                                  Optional ByVal KeepModDate As Boolean) As Boolean

    Dim l As Long, Ret As Boolean
    Dim sFileName As String
    Dim FilePos As Long
    Dim Cancel As Boolean
    
    ExtractSingleFile = False 'par défaut placé à faux, car si sortie prématurée => erreur, sinon placé à true à la fin
    If Len(ZLibVersion) = 0 Then
        Exit Function
    End If
    
    RaiseEvent Status("Début d'extraction")
    
    If CentralDirEndPos = 0 Then
        RaiseEvent ZipError(zeNoZipOpenFile, "Il n'y a pas de fichier Zip ouvert", False)
        Exit Function
    End If
    
    If FolderExists(FolderPath) = False Then
        If Not zCreateFolder(FolderPath) Then
            RaiseEvent ZipError(zeCantCreateFolder, "Impossible de créer le répertoire " & FolderPath, False)
            Exit Function
        End If
    End If

    ' Pour lire un fichier en particulier, il faut relire le fichier dans l'ordre.
    ' Il faut donc lire les entêtes de chaque fichier avant de trouver le bon
    If zReadCentralDirEnd(CentralDirEndPos) Then
        Seek #Fh, CentralDirEnd.OffSetOfCentralDir + 1
        For l = 1 To mFileCount
            Ret = zReadCentralFileHeader(sFileName)
            ' Les sous-répertoires inclus dans le Zip ont une taille compressée
            '   et décompressée de zéro octets. On recherche l'entrée suivante
    '''        Do While Ret And (varCommunZF.UncompressedSize = 0)  ' FAUX ==> un fichier peut être à 0
            Do While Ret And (CentralFileHeaderC.InternalFileAttributes = FILE_ATTRIBUTE_DIRECTORY)
                Ret = zReadCentralFileHeader(sFileName)
                DoEvents
            Loop
            If Ret Then ' si on n'est pas arrivé à la fin des données
                ' Précaution : Quand le chemin comporte le nom du disque, on a un ':'
                sFileName = Replace(sFileName, ":", "_")    ' ???? si lettre non gérée lors du zip
                If l = FileNumber Then
                    If PreservePath Then
                        ' Ajoute le répertoire du Zip au répertoire courant
                        zCheckFolder FolderPath, zGetFilePath(sFileName)
                    Else
                        ' Ou alors ne garde que le nom du fichier
                        sFileName = zGetFileName(sFileName)
                    End If
                    ' Mémo position dans le fichier car l'extraction déplace l'offset
                    FilePos = Seek(Fh)
                    RaiseEvent Status("Extraction ...\" & sFileName)
                    If FichierExiste(FolderPath & "\" & sFileName) Then
                        If OverWrite Then
                            If zRemoveFile(FolderPath & "\" & sFileName) Then
                               If zExtractFile(FolderPath & "\" & sFileName) = False Then
                                    ' Repositionne l'offset
                                    Seek Fh, FilePos
                                    Exit Function
                                End If
                            Else
                                RaiseEvent ZipError(zeCantRemoveFile, "Ne peut pas supprimer le fichier " & FolderPath & "\" & sFileName, False)
                                Exit Function
                            End If
                        Else
                            RaiseEvent ZipError(zeFileAlreadyExists, "Le fichier " & FolderPath & "\" & sFileName & " existe déjà", False)
                            Exit Function
                        End If
                    Else
                        If zExtractFile(FolderPath & "\" & sFileName) = False Then
                            ' Repositionne l'offset
                            Seek Fh, FilePos
                            Exit Function
                        End If
                    End If
                    ' Repositionne l'offset
                    Seek Fh, FilePos
                End If
            End If
            'Remise en place des dates si KeepModDate est à True
            '(on place la date de dernière modif dans date création et dernière modif, et la date en cours dans dernier accès)
            '(la date de création n'étant pas stockée dans la structure du zip on ne peut que reprendre à sa place la date de
            'la dernière màj, qui de toutes façons est la plus pertinente.
            If KeepModDate Then
                    'Récupération de la dernière date de modif pour la restituer au fichier
                    Dim maDateFT As FILETIME, maDateLT As FILETIME
                    ' Obtient la date et l'heure Dos stockée lors de la compression et la convertit en format FileTime pour conversion
                    Call DosDateTimeToFileTime(mFileInfos(l).LastModFileDate, mFileInfos(l).LastModFileTime, maDateLT)
                    ' convertit ensuite cette date locale à la date Filetime
                    Call LocalFileTimeToFileTime(maDateLT, maDateFT)
                    ' Ouvre le fichier en écriture
                    Dim hFile As Long, DateJ As FILETIME, sysDateJ As SYSTEMTIME
                    hFile = CreateFile(FolderPath & "\" & sFileName, GENERIC_WRITE, FILE_SHARE_WRITE Or FILE_SHARE_READ, ByVal CLng(0), OPEN_EXISTING, 0, 0)
                    If hFile <> INVALID_HANDLE_VALUE Then ' s'il n'y a pas d'erreur d'accès au fichier
                        ' On change les dates du fichier, et voilà c'est mieux maintenant !
                        
                        'retaitement date du jour pour date écriture
                        sysDateJ.wDay = Day(Now)
                        sysDateJ.wMonth = Month(Now)
                        sysDateJ.wYear = Year(Now)
                        sysDateJ.wHour = Hour(Now)
                        sysDateJ.wMinute = Minute(Now)
                        sysDateJ.wSecond = Second(Now)
                        Call SystemTimeToFileTime(sysDateJ, DateJ)
                        Call LocalFileTimeToFileTime(DateJ, DateJ)
                        '
                        Call SetFileTime(hFile, maDateFT, DateJ, maDateFT) 'création,dernier accès, modification
                        ' on ferme le fichier ouvert
                        CloseHandle (hFile)
                    End If  'sinon les dates ne seront pas corrigées
            End If  'fin reprise des dates
            RaiseEvent ProgressExtract(100, Cancel)
            DoEvents
        Next
        ExtractSingleFile = True
    End If

    RaiseEvent Status("Extraction terminée")

End Function

Public Function ExtractAllFiles(ByVal FolderPath As String, _
                                Optional ByVal PreservePath As Boolean, _
                                Optional ByVal OverWrite As Boolean, _
                                Optional ByVal KeepModDate As Boolean) As Boolean
    Dim l As Long, Ret As Boolean
    Dim sFileName As String
    Dim FilePos As Long
    Dim Cancel As Boolean
    
    If Len(ZLibVersion) = 0 Then
        Exit Function
    End If
    
    RaiseEvent Status("Début d'extraction")
    
    If CentralDirEndPos = 0 Then
        RaiseEvent ZipError(zeNoZipOpenFile, "Il n'y a pas de fichier Zip ouvert", False)
        Exit Function
    End If
    
    If FolderExists(FolderPath) = False Then
        If Not zCreateFolder(FolderPath) Then
            RaiseEvent ZipError(zeCantCreateFolder, "Impossible de créer le répertoire " & FolderPath, False)
            Exit Function
        End If
    End If

    If zReadCentralDirEnd(CentralDirEndPos) Then
        Seek #Fh, CentralDirEnd.OffSetOfCentralDir + 1
        For l = 1 To mFileCount
            Ret = zReadCentralFileHeader(sFileName)
            ' Les sous-répertoires inclus dans le Zip ont une taille compressée
            '   et décompressée de zéro octets. On recherche l'entrée suivante
            ' C'est vrai mais il faut analyser l'attribut plutôt que la taille car il pourrait y
            ' avoir 1 ou des fichiers de taille 0 dans le zip, c'est vrai c'est un peu bête de comprimer
            ' un fichier de taille 0 qui une fois comprimé fera 12 octets !!!! mais ça peut arriver ....
'''            Do While Ret And (varCommunZF.UncompressedSize = 0)
            Do While Ret And (CentralFileHeaderC.InternalFileAttributes = FILE_ATTRIBUTE_DIRECTORY)
                Ret = zReadCentralFileHeader(sFileName)
                DoEvents
            Loop
            If Ret Then ' si on n'est pas arrivé à la fin des données
                ' Précaution : Quand le chemin comporte le nom du disque, on a un ':'
                sFileName = Replace(sFileName, ":", "_")
                If PreservePath Then
                    ' Ajoute le répertoire du Zip au répertoire courant
                    zCheckFolder FolderPath, zGetFilePath(sFileName)
                Else
                    ' Ou alors ne garde que le nom du fichier
                    sFileName = zGetFileName(sFileName)
                End If
                RaiseEvent Status("Extraction ...\" & sFileName)
                ' Mémo position dans le fichier car l'extraction déplace l'offset
                FilePos = Seek(Fh)
                If FichierExiste(FolderPath & "\" & sFileName) Then
                    If OverWrite Then
                        If zRemoveFile(FolderPath & "\" & sFileName) Then
                            zExtractFile FolderPath & "\" & sFileName
                        Else
                            RaiseEvent ZipError(zeCantRemoveFile, "Ne peut pas supprimer le fichier " & FolderPath & "\" & sFileName, False)
                        End If
                    Else
                        RaiseEvent ZipError(zeFileAlreadyExists, "Le fichier " & FolderPath & "\" & sFileName & " existe déjà", False)
                    End If
                Else
                    zExtractFile FolderPath & "\" & sFileName
                End If
                ' Repositionne l'offset
                Seek Fh, FilePos
                'Remise en place des dates si KeepModDate est à True
                '(on place la date de dernière modif dans date création et dernière modif, et la date en cours dans dernier accès)
                '(la date de création n'étant pas stockée dans la structure du zip on ne peut que reprendre à sa place la date de
                'la dernière màj, qui de toutes façons est la plus pertinente.
                If KeepModDate Then
                        'Récupération de la dernière date de modif pour la restituer au fichier
                        Dim maDateFT As FILETIME, maDateLT As FILETIME
                        ' Obtient la date et l'heure Dos stockée lors de la compression et la convertit en format FileTime pour conversion
                        Call DosDateTimeToFileTime(mFileInfos(l).LastModFileDate, mFileInfos(l).LastModFileTime, maDateLT)
                        ' convertit ensuite cette date locale à la date Filetime
                        Call LocalFileTimeToFileTime(maDateLT, maDateFT)
                        ' Ouvre le fichier en écriture
                        Dim hFile As Long, DateJ As FILETIME, sysDateJ As SYSTEMTIME
                        hFile = CreateFile(FolderPath & "\" & sFileName, GENERIC_WRITE, FILE_SHARE_WRITE Or FILE_SHARE_READ, ByVal CLng(0), OPEN_EXISTING, 0, 0)
                        If hFile <> INVALID_HANDLE_VALUE Then ' s'il n'y a pas d'erreur d'accès au fichier
                            ' On change les dates du fichier, et voilà c'est mieux maintenant !
                            
                            'retaitement date du jour pour date écriture
                            sysDateJ.wDay = Day(Now)
                            sysDateJ.wMonth = Month(Now)
                            sysDateJ.wYear = Year(Now)
                            sysDateJ.wHour = Hour(Now)
                            sysDateJ.wMinute = Minute(Now)
                            sysDateJ.wSecond = Second(Now)
                            Call SystemTimeToFileTime(sysDateJ, DateJ)
                            Call LocalFileTimeToFileTime(DateJ, DateJ)
                            '
                            Call SetFileTime(hFile, maDateFT, DateJ, maDateFT) 'création,dernier accès, modification
                            ' on ferme le fichier ouvert
                            CloseHandle (hFile)
                        End If  'sinon les dates ne seront pas corrigées
                End If  'fin reprise des dates
                '
                RaiseEvent ProgressExtract((l / CentralDirEnd.EntriesInTheCentralDir) * 100, Cancel)
                If Cancel Then
                    Exit Function
                End If
            End If
            DoEvents
        Next
        ExtractAllFiles = True
    End If

    RaiseEvent Status("Extraction terminée")

End Function

Public Property Get ZipIsOpen() As Boolean
    ' Renvoie True si un Zip est déjà ouvert
    ZipIsOpen = Not (CentralDirEndPos = 0)
End Property

Public Property Get ZipComment() As String
    ' Renvoie le commentaire du fichier Zip existant
    ZipComment = mZipComment
End Property

Public Property Get inFileCount() As Long
    ' Renvoie le nombre de fichiers contenus dans le Zip
    inFileCount = mFileCount
End Property
    
Public Property Get inFileName(ByVal index As Long) As String
    ' Renvoie le nom du fichier désigné par Index
    inFileName = mFileInfos(index).FileName
End Property

Public Property Get FileCompressedSize(ByVal index As Long) As Long
    ' Renvoie la taille du fichier compressé désigné par Index
    FileCompressedSize = mFileInfos(index).CompressedSize
End Property

Public Property Get FileUncompressedSize(ByVal index As Long) As Long
    ' Renvoie la taille originale du fichier désigné par Index
    FileUncompressedSize = mFileInfos(index).UncompressedSize
End Property

Public Property Get FileDateAndTime(ByVal index As Long) As Date
    ' Renvoie la date et l'heure du fichier original désigné par Index
    ' Le format fourni est du type DOS. On le converti en date/heure pour fichier
    '   puis au format standard, puis on le triture pour qu'il apparaisse selon
    '   le choix de l'utilisateur (dépend de la langue)
    Dim maDate1 As FILETIME, maDate2 As SYSTEMTIME
    Call DosDateTimeToFileTime(mFileInfos(index).LastModFileDate, _
                               mFileInfos(index).LastModFileTime, _
                               maDate1)
    Call FileTimeToSystemTime(maDate1, maDate2)
    With maDate2
        FileDateAndTime = DateSerial(.wYear, .wMonth, .wDay) & " " & _
                          TimeSerial(.wHour, .wMinute, .wSecond)
    End With
End Property

Public Property Get ZLibVersion() As String
    ' Renvoie la version de la DLL (permet aussi de savoir si elle existe bien)
    On Error GoTo eh
    ZLibVersion = zPointerToString(ZLibVer)
    Exit Property
eh:
    RaiseEvent ZipError(zeZLibNotInstalled, "La librairie Zlib n'est pas installée", False)
End Property

Public Property Get LevelCompression() As eLevelCompres
    ' Renvoie le niveau de compression en cours
    LevelCompression = m_LevelComp
End Property
Public Property Let LevelCompression(lgLevel As eLevelCompres)
    ' Fixe le niveau de compression
    m_LevelComp = lgLevel
End Property

Private Sub Class_Initialize()
    Erase tabFiles      'efface le tableau à chaque initialisation
    ReDim arrDelFile(0) ' crée le 1° enreg à vide au tableau pour renseigner fichier à détruire
    
''    ' Si la librairie n'existe pas, on la recréé
''    If fichierexiste(App.Path & "\ZLib.dll")=false Then Call ExtraitDLL
    
    With LocalFileHeader
        .LocalFileHeaderSignature = &H4034B50
        .VersionNeededToExtract = 20
    End With
    With varCommunZF
        .GeneralPurposeBitFlag = 0 '2       'ne concerne pas la gestion des disques
        .CompressionMethod = 8              'ne concerne pas le taux
    End With
    With CentralFileHeaderA
        .CentralFileHeaderSigniature = &H2014B50
        .VersionMadeBy = 20
        .VersionNeededToExtract = 20
    End With
    CentralFileHeaderC.InternalFileAttributes = 32
    '
    m_Nombre = 0 'ràz
    'initialiser constante taux compression standard, équivalent niveau 6
    m_LevelComp = Z_DEFAULT_COMPRESSION
    
    ' Constantes de conversion ASCII <> ANSI
    sASCI = Chr$(132) & Chr$(148) & Chr$(129) & Chr$(225) & Chr$(142) & Chr$(153) & Chr$(154) & Chr$(128) & Chr$(130) & Chr$(131) & Chr$(133) & Chr$(134) & Chr$(135) & Chr$(136) & Chr$(137) & Chr$(138) & Chr$(139) & Chr$(140) & Chr$(143) & Chr$(144) & Chr$(147) & Chr$(150) & Chr$(151) & Chr$(20) & Chr$(145) & Chr$(146) & Chr$(155) & Chr$(156) & Chr$(166) & Chr$(167) & Chr$(171) & Chr$(172) & Chr$(237) & Chr$(241) & Chr$(246) & Chr$(253) & Chr$(21) & Chr$(157)
    sANSI = Chr$(228) & Chr$(246) & Chr$(252) & Chr$(223) & Chr$(196) & Chr$(214) & Chr$(220) & Chr$(199) & Chr$(233) & Chr$(226) & Chr$(224) & Chr$(229) & Chr$(231) & Chr$(234) & Chr$(235) & Chr$(232) & Chr$(239) & Chr$(238) & Chr$(197) & Chr$(201) & Chr$(244) & Chr$(251) & Chr$(249) & Chr$(182) & Chr$(230) & Chr$(198) & Chr$(162) & Chr$(163) & Chr$(170) & Chr$(186) & Chr$(189) & Chr$(188) & Chr$(216) & Chr$(177) & Chr$(247) & Chr$(178) & Chr$(167) & Chr$(165)
                
    '  initialise une table pour calculer Crc32
    Dim vntA As Variant
    Dim I As Integer, iOffset As Integer
    Dim nLen As Integer

    iOffset = 0
    nLen = 32
    vntA = Array( _
        &H0, &H77073096, &HEE0E612C, &H990951BA, _
        &H76DC419, &H706AF48F, &HE963A535, &H9E6495A3, _
        &HEDB8832, &H79DCB8A4, &HE0D5E91E, &H97D2D988, _
        &H9B64C2B, &H7EB17CBD, &HE7B82D07, &H90BF1D91, _
        &H1DB71064, &H6AB020F2, &HF3B97148, &H84BE41DE, _
        &H1ADAD47D, &H6DDDE4EB, &HF4D4B551, &H83D385C7, _
        &H136C9856, &H646BA8C0, &HFD62F97A, &H8A65C9EC, _
        &H14015C4F, &H63066CD9, &HFA0F3D63, &H8D080DF5)

    For I = iOffset To iOffset + nLen - 1
        arrCRC32(I) = vntA(I - iOffset)
    Next
    iOffset = iOffset + nLen
    
    vntA = Array( _
        &H3B6E20C8, &H4C69105E, &HD56041E4, &HA2677172, _
        &H3C03E4D1, &H4B04D447, &HD20D85FD, &HA50AB56B, _
        &H35B5A8FA, &H42B2986C, &HDBBBC9D6, &HACBCF940, _
        &H32D86CE3, &H45DF5C75, &HDCD60DCF, &HABD13D59, _
        &H26D930AC, &H51DE003A, &HC8D75180, &HBFD06116, _
        &H21B4F4B5, &H56B3C423, &HCFBA9599, &HB8BDA50F, _
        &H2802B89E, &H5F058808, &HC60CD9B2, &HB10BE924, _
        &H2F6F7C87, &H58684C11, &HC1611DAB, &HB6662D3D)

    For I = iOffset To iOffset + nLen - 1
        arrCRC32(I) = vntA(I - iOffset)
    Next
    iOffset = iOffset + nLen
    
    vntA = Array( _
        &H76DC4190, &H1DB7106, &H98D220BC, &HEFD5102A, _
        &H71B18589, &H6B6B51F, &H9FBFE4A5, &HE8B8D433, _
        &H7807C9A2, &HF00F934, &H9609A88E, &HE10E9818, _
        &H7F6A0DBB, &H86D3D2D, &H91646C97, &HE6635C01, _
        &H6B6B51F4, &H1C6C6162, &H856530D8, &HF262004E, _
        &H6C0695ED, &H1B01A57B, &H8208F4C1, &HF50FC457, _
        &H65B0D9C6, &H12B7E950, &H8BBEB8EA, &HFCB9887C, _
        &H62DD1DDF, &H15DA2D49, &H8CD37CF3, &HFBD44C65)

    For I = iOffset To iOffset + nLen - 1
        arrCRC32(I) = vntA(I - iOffset)
    Next
    iOffset = iOffset + nLen
    
    vntA = Array( _
        &H4DB26158, &H3AB551CE, &HA3BC0074, &HD4BB30E2, _
        &H4ADFA541, &H3DD895D7, &HA4D1C46D, &HD3D6F4FB, _
        &H4369E96A, &H346ED9FC, &HAD678846, &HDA60B8D0, _
        &H44042D73, &H33031DE5, &HAA0A4C5F, &HDD0D7CC9, _
        &H5005713C, &H270241AA, &HBE0B1010, &HC90C2086, _
        &H5768B525, &H206F85B3, &HB966D409, &HCE61E49F, _
        &H5EDEF90E, &H29D9C998, &HB0D09822, &HC7D7A8B4, _
        &H59B33D17, &H2EB40D81, &HB7BD5C3B, &HC0BA6CAD)
        
    For I = iOffset To iOffset + nLen - 1
        arrCRC32(I) = vntA(I - iOffset)
    Next
    iOffset = iOffset + nLen

    vntA = Array( _
        &HEDB88320, &H9ABFB3B6, &H3B6E20C, &H74B1D29A, _
        &HEAD54739, &H9DD277AF, &H4DB2615, &H73DC1683, _
        &HE3630B12, &H94643B84, &HD6D6A3E, &H7A6A5AA8, _
        &HE40ECF0B, &H9309FF9D, &HA00AE27, &H7D079EB1, _
        &HF00F9344, &H8708A3D2, &H1E01F268, &H6906C2FE, _
        &HF762575D, &H806567CB, &H196C3671, &H6E6B06E7, _
        &HFED41B76, &H89D32BE0, &H10DA7A5A, &H67DD4ACC, _
        &HF9B9DF6F, &H8EBEEFF9, &H17B7BE43, &H60B08ED5)

    For I = iOffset To iOffset + nLen - 1
        arrCRC32(I) = vntA(I - iOffset)
    Next
    iOffset = iOffset + nLen

    vntA = Array( _
        &HD6D6A3E8, &HA1D1937E, &H38D8C2C4, &H4FDFF252, _
        &HD1BB67F1, &HA6BC5767, &H3FB506DD, &H48B2364B, _
        &HD80D2BDA, &HAF0A1B4C, &H36034AF6, &H41047A60, _
        &HDF60EFC3, &HA867DF55, &H316E8EEF, &H4669BE79, _
        &HCB61B38C, &HBC66831A, &H256FD2A0, &H5268E236, _
        &HCC0C7795, &HBB0B4703, &H220216B9, &H5505262F, _
        &HC5BA3BBE, &HB2BD0B28, &H2BB45A92, &H5CB36A04, _
        &HC2D7FFA7, &HB5D0CF31, &H2CD99E8B, &H5BDEAE1D)

    For I = iOffset To iOffset + nLen - 1
        arrCRC32(I) = vntA(I - iOffset)
    Next
    iOffset = iOffset + nLen

    vntA = Array( _
        &H9B64C2B0, &HEC63F226, &H756AA39C, &H26D930A, _
        &H9C0906A9, &HEB0E363F, &H72076785, &H5005713, _
        &H95BF4A82, &HE2B87A14, &H7BB12BAE, &HCB61B38, _
        &H92D28E9B, &HE5D5BE0D, &H7CDCEFB7, &HBDBDF21, _
        &H86D3D2D4, &HF1D4E242, &H68DDB3F8, &H1FDA836E, _
        &H81BE16CD, &HF6B9265B, &H6FB077E1, &H18B74777, _
        &H88085AE6, &HFF0F6A70, &H66063BCA, &H11010B5C, _
        &H8F659EFF, &HF862AE69, &H616BFFD3, &H166CCF45)

    For I = iOffset To iOffset + nLen - 1
        arrCRC32(I) = vntA(I - iOffset)
    Next
    iOffset = iOffset + nLen

    vntA = Array( _
        &HA00AE278, &HD70DD2EE, &H4E048354, &H3903B3C2, _
        &HA7672661, &HD06016F7, &H4969474D, &H3E6E77DB, _
        &HAED16A4A, &HD9D65ADC, &H40DF0B66, &H37D83BF0, _
        &HA9BCAE53, &HDEBB9EC5, &H47B2CF7F, &H30B5FFE9, _
        &HBDBDF21C, &HCABAC28A, &H53B39330, &H24B4A3A6, _
        &HBAD03605, &HCDD70693, &H54DE5729, &H23D967BF, _
        &HB3667A2E, &HC4614AB8, &H5D681B02, &H2A6F2B94, _
        &HB40BBE37, &HC30C8EA1, &H5A05DF1B, &H2D02EF8D)
        
    For I = iOffset To iOffset + nLen - 1
        arrCRC32(I) = vntA(I - iOffset)
    Next
    iOffset = iOffset + nLen

End Sub

Private Sub Class_Terminate()
    ZipClose
    Erase tabFiles   'efface le tableau à chaque fermeture de la classe
    cptTab = 0      'index du tableau
    ' si le tableau des fichiers à détruire contient des enregistrements (ubound >0, c'est 1 tableau base 0 et on met à jour à partie de 1)
    ' générer la destruction des fichiers provisoire au redémarrage de Windows par un fichier .bat déclenché par la clé RunOnce de la Bdr.
    If UBound(arrDelFile) > 0 Then
        Dim Fnu As Byte, I As Long, strBuff As String, lgRet As Long, strFicBat As String
        Fnu = FreeFile
        strFicBat = zGetFilePath(arrDelFile(1)) & "\destrucFic.bat"
        'crée le batch Dos
        If FichierExiste(strFicBat, False) Then
            'si le fichier existe déjà, le complète et supprime la 1° instruction del %0
            Open strFicBat For Binary As #Fnu
            Seek #Fnu, LOF(Fnu) - 7
            For I = 1 To UBound(arrDelFile)
                strBuff = Space$(255)
                lgRet = GetShortPathName(arrDelFile(I), strBuff, 255) ' passer en nom court
                Put #Fnu, , "del " & Left$(strBuff, lgRet) & vbNewLine 'pour détruire les fichiers
            Next
            Put #Fnu, , "del %0" & vbNewLine 'pour que ce fichier s'autodétruise
        Else
            Open strFicBat For Output As #Fnu
            Print #Fnu, "@echo off"
            For I = 1 To UBound(arrDelFile)
                strBuff = Space$(255)
                lgRet = GetShortPathName(arrDelFile(I), strBuff, 255) ' passer en nom court
                Print #Fnu, "del " & Left$(strBuff, lgRet)  'pour détruire les fichiers
            Next
            Print #Fnu, "del %0" 'pour que ce fichier s'autodétruise
        End If
        Close #Fnu
        'générer la bdr (HKEY_LOCAL_MACHINE)
        Dim Ident As Long
        'Dim Donnee As String
        'Dim TailleBuffer As Long
        lgRet = RegCreateKey(&H80000002, "SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce", Ident)
        lgRet = RegSetValueEx(Ident, "DelFile_clsZip", 0&, 1, ByVal strFicBat, Len(strFicBat) + 1)
    End If
    Erase arrDelFile
End Sub

'########### Les outils internes de la classe
Private Function zGetFileName(ByVal Path As String) As String
    
    ' On récupère ce qu'il y a derrière le dernier \
    Dim l As Long
    l = InStrRev(Path, "\")
    If l > 0 Then
        zGetFileName = Right$(Path, Len(Path) - l)
    Else
        zGetFileName = Path
    End If
    
End Function

Private Function zGetFilePath(ByVal Path As String) As String
    
    Dim l As Long
    
    ' On récupère ce qu'il y a avant le dernier \
    l = InStrRev(Path, "\")
    If l > 0 Then
        zGetFilePath = Left$(Path, l - 1)
    End If
    
End Function

Private Sub zCheckFolder(ByVal FolderPath As String, ByVal CheckPath As String)
     
    ' Créé le chemin CheckPath derrière FolderPath
    Dim s() As String
    Dim v As Variant

    s = Split(CheckPath, "\")
    For Each v In s
        If Right(FolderPath, 1) <> "\" Then
            FolderPath = FolderPath & "\" & v
        Else
            FolderPath = FolderPath & v
        End If
        If FolderExists(FolderPath) = False Then MkDir FolderPath
    Next
    
End Sub

Private Function zReadCentralFileHeader(ByRef sFileName As String) As Boolean
'récupère infos sur fichier en lisant File header dans Central directory structure
'et les interpête

    Dim ExtraField As String
    Dim Comment As String
    
    ' On va récupérer les infos concernant ce fichier.
    ' CentralFileHeader contient :
    ' - soit les données si ce sont des chiffres,
    ' - soit les offsets (décalage) de chacune des variables chaine qui le compose :
    '    il suffit de récolter les données en respectant les tailles fournies
    zReadCentralFileHeader = False
    Get #Fh, , CentralFileHeaderA    ' Lit la structure centralFileHeader
    Get #Fh, , varCommunZF           ' Lit la structure centralFileHeader
    Get #Fh, , CentralFileHeaderC    ' Lit la structure centralFileHeader
    If CentralFileHeaderA.CentralFileHeaderSigniature = CentralFileHeaderSigniature Then
        ' puis lit les chaines en fonction des tailles
        sFileName = Space(varCommunZF.FileNameLength)
        Get #Fh, , sFileName
        ' Retransforme le nom du fichier ASCII étendu en ISO
        sFileName = OEMConvert(sFileName, [ASCII to ANSI])
        sFileName = Replace(sFileName, "/", "\")
        ExtraField = Space(varCommunZF.ExtraFieldLength)
        Get #Fh, , ExtraField
        Comment = Space(CentralFileHeaderC.FileCommentLength)
        Get #Fh, , Comment ' du fichier, pas du zip (ne sert à rien sous windows)
    Else
        ' Dans le cas où on interroge après la dernière structure
        ' Normalement, ça ne doit jamais arriver (si mFileCount est bien géré !)
        Exit Function
    End If
    zReadCentralFileHeader = True
    ' CentralFileHeader est commun a toute la classe
    ' Il gardera donc ces données jusqu'au prochain appel de cette Sub
    ' -> Important pour la mémo des paramètres dans ZipOpen
    
End Function

Private Function zReadCentralDirEnd(ByVal Position As Long) As Boolean

    Dim Temp As String
    
    Get #Fh, Position, CentralDirEnd
    Temp = Space(CentralDirEnd.ZipFileCommentLength)
    Get #Fh, , Temp
    mZipComment = Temp
    mFileCountTemp = CentralDirEnd.EntriesInTheCentralDir
    
    zReadCentralDirEnd = CentralDirEnd.NumberOfThisDisk = CentralDirEnd.NumberOfDiskWithCentralDir

End Function

Private Function zExtractFile(ByVal Path As String) As Boolean

    ' Extrait le fichier en cours : dépend de l'offset qui est défini
    '   dans les fonctions appelantes.
    Dim b() As Byte, lgMax As Long, lgSize As Long, lgmem As MEMORYSTATUS
    Dim Cancel As Boolean 'pour gérer retour erreur par utilisateur

    On Error GoTo testErr
    zExtractFile = True 'par défaut supposé correct
    Get #Fh, CentralFileHeaderC.RelativeOffsetOfLocalHeader + 1, LocalFileHeader
    Get #Fh, , varCommunZF
    If LocalFileHeader.LocalFileHeaderSignature = LocalFileHeaderSignature Then
        ' On lit les données sFileName et ExtraField mais on ne s'en sert pas ici
        ' C'est juste pour décaler l'offset correctement
        ' On aurait pu faire un Seek, mais bon ...
        sFileName = Space(varCommunZF.FileNameLength)
        Get #Fh, , sFileName
        ' Retransforme le nom du fichier ASCII étendu en ISO
        sFileName = OEMConvert(sFileName, [ASCII to ANSI])
        ExtraField = Space(varCommunZF.ExtraFieldLength)
        Get #Fh, , ExtraField
        ReDim b(varCommunZF.CompressedSize - 1) 'bug taille gros fichiers en gestion d'erreurs
        Get #Fh, , b
        If varCommunZF.CompressionMethod = 0 Then 'No Compression
            zSaveFile Path, b, varCommunZF.UncompressedSize  ' écrit le fichier sur disque
        ElseIf varCommunZF.CompressionMethod = 8 Then 'Deflate Method
            If zUnCompressBytes(b, varCommunZF.CRC32) Then  'si crc identique, on poursuit
                zSaveFile Path, b, varCommunZF.UncompressedSize   ' écrit le fichier sur disque
            Else        'si erreur de crc => message erreur utilisateur et gestion choix utilisateur
                RaiseEvent ZipError(zeChecksumError, "Erreur de checksum (CRC) dans " & Path, Cancel)
                DoEvents
                If Cancel = True Then   'l'utilsateur a choisi de décompresser quand-même
                    zSaveFile Path, b, varCommunZF.UncompressedSize   ' écrit le fichier sur disque
                Else        'l'utilisateur arrête le décompression
                    zExtractFile = False
                End If
            End If
        Else        'Autres méthodes de compression Non Gérées, pour l'instant
            RaiseEvent ZipError(zeUnsupportedCompressionMethod, _
                                "La méthode de compression pour " & sFileName & _
                                " n'est pas supportée", False)
            zExtractFile = False
        End If
    Else        ' la signature du zip n'est pas correcte
        RaiseEvent ZipError(zeStructureError, "Structure corrompue dans le fichier " & Path, False)
        zExtractFile = False
    End If
    On Error GoTo 0
    Exit Function
    
testErr:
    If Err.Number = 7 Then  'capa mémoire
        If varCommunZF.CompressionMethod = 0 Then 'si le gros fichier est sans compression (élaboré avec cette classe)
            ' passer la décompression ici en paquets pour écrire le fichier décompressé
                'identification de l'utilisation raisonnable de la mémoire phisyque par la tableau
                GlobalMemoryStatus lgmem
                lgMax = lgmem.dwAvailPhys / 4   'mémoire physique dispo/4 pour sécu sur le reste appli et les autres applis
                'écriture par paquets
                lgSize = varCommunZF.UncompressedSize
                Dim lfh As Long
                lfh = FreeFile
                Open Path For Binary As #lfh
                Do While lgSize > lgMax
                    DoEvents
                    ReDim b(lgMax - 1)
                    lgSize = lgSize - lgMax
                    Get #Fh, , b
                    Put #lfh, , b       ' écrit un morceau du fichier sur disque
                Loop
                'préparation et écriture du dernier paquet
                ReDim b(lgSize - 1)
                Get #Fh, , b
                Put #lfh, , b
                Close #lfh
                On Error GoTo 0
                Exit Function
        Else    'tout est perdu on ne peut pas faire
            RaiseEvent ZipError(zeMemory, "La taille du fichier compressé " & sFileName & _
            " n'est pas supportée par le logiciel.", False)
            zExtractFile = False
            On Error GoTo 0
            Exit Function
        End If
    Else        'autre
        RaiseEvent ZipError(zeUnKnownError, "Erreur non identifiée lors de la décompression du fichier " & _
        sFileName, False)
        zExtractFile = False
        On Error GoTo 0
        Exit Function
    End If
End Function

Private Function zCreateFolder(ByVal Path As String) As Boolean

     ' On construit l'arborescence en partant du début
     ' (permet de recontruire un arbre entier)
     Dim s() As String, Temp As String, r As Long
     On Error GoTo Erreur
     s = Split(Path, "\")
     For r = 0 To UBound(s)
          Temp = Temp & s(r)    '& "\"
          If FolderExists(Temp) = False Then MkDir Temp    'la fonction folderExists ajoutera le \ qui manquait
     Next r
     zCreateFolder = True
     Exit Function
     
'''     On Error GoTo Erreur
'''          If FolderExists(Path) = False Then MkDir Path
'''     zCreateFolder = True
'''     Exit Function
     
Erreur:
On Error GoTo 0
End Function

Private Function zRemoveFile(Path As String) As Boolean

    ' Supprime un fichier
    On Error GoTo eh
    Kill Path
    zRemoveFile = True
eh:
End Function

'Private Function zGetCentralDirEndPos(ByVal Fh As Long) As Long
'    Dim data() As Byte      'tableau d'octets pour recevoir le fichier
'    Dim l As Long
'    Dim m As Long
'    Dim Cancel As Boolean
'    ReDim data(LOF(fh) - 1)     'attention peut dépendre mémoire du matos
'        Get #fh, 1, data
'    For l = UBound(data) - 3 To LBound(data) Step -1
'        CopyMemory_Long_Byte m, data(l), 4
'        If m = EndOFCentralDirSignature Then
'            zGetCentralDirEndPos = l + 1
'            Exit Function
'        End If
'    Next
'End Function

Private Function zUnCompressBytes(ByRef buffer() As Byte, ByVal CRC32 As Long) As Boolean
    
    Dim b() As Byte
    Dim BufferSize As Long
    Dim FileSize As Long
    Dim crc As Long
    Dim r As Long

    ReDim b(UBound(buffer) + 2)
    
    'Zlib's Uncompress method expects the 2 byte head that the Compress method adds
    'so we put that on first. Luckily it's always the same value.
    b(0) = 120
    b(1) = 156
    CopyMemory_Byte_Byte b(2), buffer(0), UBound(buffer) + 1

    FileSize = UBound(buffer) + 3
    BufferSize = varCommunZF.UncompressedSize * 1.01 + 12       ' OK si compress avec CompressBytes et même avec compressDeflate, donc maintenu pour compatibilité
'    BufferSize = varCommunZF.UncompressedSize * 1.5 + 12       ' si compress avec CompressDeflate
    ReDim buffer(BufferSize - 1) As Byte
    
    r = UnCompress_Byte_Long_Byte(buffer(0), BufferSize, b(0), FileSize)
    If varCommunZF.UncompressedSize > 0 Then    'pour gérer décompression des fichiers à 0 ou pas
        ReDim Preserve buffer(varCommunZF.UncompressedSize - 1)
        crc = lCRC32_Byte(0&, buffer(0), UBound(buffer) + 1)
        If crc = CRC32 Then
            zUnCompressBytes = True
        Else
            zUnCompressBytes = False
        End If
    Else
        ReDim Preserve buffer(0)
        zUnCompressBytes = True     'forcément ok
    End If

End Function

Private Sub zSaveFile(ByVal Path As String, ByRef data() As Byte, lgSizeFile As Long)

    ' Sauvegarde les données dans le fichier
    Dim lfh As Long
    lfh = FreeFile
    Open Path For Binary As #lfh
       If lgSizeFile > 0 Then     'pour gérer décompression des fichiers à 0 ou pas
          Put #lfh, , data
        End If
    Close #lfh

End Sub

Private Function zPointerToString(Pointer As Long) As String

    ' Utilisé par ZLibVersion (héritage de C)
    Dim l As Long
    Dim s As String
    
    l = lstrlen(Pointer)
    s = Space(l)
    l = lstrcpy(s, Pointer)
    If l > 0 Then
        zPointerToString = s
    End If
    
End Function

Public Function Conv_ASCII_ISO(ByVal sChaine As String) As String
    
    ' Convertit les codes des caractères ASCII étendu en code de caractères ISO 8859-1
    ' (codes ascii > 127 : les caractères accentués ne sont pas aux mêmes places)
    ' Voir les tableaux Table_Conversion_ASCII_ISO() et Table_Conversion_ISO_ASCII()
    ' sChaine est passée ByVal pour ne pas être modifiée ici.
    
    ' La fonction inverse se trouve dans la clsZip_File : Conv_ISO_ASCII
    
    Dim r As Long
    Dim t As Byte
    
    For r = 1 To Len(sChaine)
        t = Asc(Mid$(sChaine, r, 1))
        Mid$(sChaine, r, 1) = Chr$(Table_Conversion_ASCII_ISO(t))
    Next r
    Conv_ASCII_ISO = sChaine
    
End Function

Private Function ybGetCentralDirEndPos(ZipPath As String) As Long
' Plus rapide pour les gros fichiers et moins limitée par la taille des fichiers (2,147 Go) que la version d'origine
' zGetCentralDirEndPos qui passait pr des tableaux, + gestion de la progression !
    Dim hOrgFile As Long
    Dim nSize As Long
    Dim Ret As Long
    Dim l As Long
    Dim m As Long
    Dim Cancel As Boolean
    Dim lu As Long
    Const nbAlire As Long = 4   'nbre d'octets à lire
    On Error GoTo ErrorHandler

    'Récupère un Handle pour manipuler le fichier
    hOrgFile = CreateFile(ZipPath, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, _
                            ByVal 0&, OPEN_EXISTING, 0, 0)
    'Renvoi d'une erreur -> Problème d'accès au fichier
    If hOrgFile = INVALID_HANDLE_VALUE Then
        ybGetCentralDirEndPos = 0
        RaiseEvent ZipError(zeNotZipFile, "Le fichier " & ZipPath & " n'est pas accessible ou a été déplacé !", False)
        'on pourrait éventuellement gérer un choix de l'utilisateur
        DoEvents
        Exit Function
    End If
    
    'Taille du fichier
    nSize = GetFileSize(hOrgFile, 0)
    'Attention, GetFileSize peut renvoyer une taille de fichier maximum de : 2,147,483,647 bytes.
    If nSize = ERROR Then
        ybGetCentralDirEndPos = 0
        RaiseEvent ZipError(zeNotZipFile, "La Taille du fichier " & ZipPath & " est trop importante !" & _
        vbNewLine & "Ce logiciel ne gère pas des tailles de fichiers supérieurs à la FAT32.", False)
        DoEvents
        Exit Function
    End If
    DoEvents
        
    For l = nSize - 3 To 1 Step -1
        'déplace la position du pointeur
        Ret = SetFilePointer(hOrgFile, l, 0, FILE_BEGIN)
        'Charge le contenu du nbre d'octets lus (définis par nbAlire) depuis la position l du pointeur
        'dans le fichier ouvert dans la variable m. Il est indispensable de placer la variable lu de type
        'long dont on ne se sert pas pour récupérer les valeurs en m, si on remplaçait cette variable
        'par 0, comme on le voit dans beaucoup d'exemples, on ne récupèrerait que 0 dans m.
        'Il y aurait aussi la possibilité de récupérer en remplaçant m par ByVal et 1 variable string,
        '1 chaine ASCII et de la transformer mais ça serait + long !
        ReadFile hOrgFile, m, nbAlire, lu, ByVal 0&
        If m = CLng(EndOFCentralDirSignature) Then
            ybGetCentralDirEndPos = l + 1
            'Ferme le fichier
            CloseHandle hOrgFile
            'barre de progression      (plus trés utile c'est tellement rapide !)
            RaiseEvent ProgressLoad(95, Cancel)
            Exit Function
        End If
        DoEvents
    Next
    'Ferme le fichier
    CloseHandle hOrgFile
    On Error GoTo 0
    Exit Function
ErrorHandler:
    ybGetCentralDirEndPos = 0
    CloseHandle hOrgFile

End Function

Public Function CreateZip(FilePath As String, Optional OverWrite As Boolean) As Long

    ' Créé le fichier ZIP lui même dans lequel on va ajouter les fichiers
    ' au fur et à mesure par ybFileAdd
    ' FilePath  : Nom du fichier ZIP avec son chemin complet
    ' Overwrite : On l'écrase s'il existe déjà ?
    ' Retourne le N° d'ouverture de fichier ou 0 si non créé (existant et non overwrite)
    ' Remet le tableau (ou collection à blanc) pour être sûr si ça n'avait pas été fait

    'Dim l As Long
    Dim Fh As Long
   
    CreateZip = 0 'par défaut
    ' Si le Zip existe déjà, on le détruit (si Overwrite est à True)
    If FichierExiste(FilePath, False) Then
        If OverWrite Then
            On Error GoTo Fin
            Kill FilePath
        Else
            RaiseEvent FileZipErr(FilePath, 11, "Création -No Overwrite- du Zip abandonnée car le fichier existe déjà")
            Exit Function
        End If
    End If
    
    ' Ouvre/Crée le Zip pour enregistrer chacune des structures des fichiers
    Fh = FreeFile
    Open FilePath For Binary As #Fh
    CreateZip = Fh
    Erase tabFiles          'par sécurité
Fin:
On Error GoTo 0

End Function

Public Function ybFileAdd(ByVal FilePath As String, _
                        ByVal PathChoice As enumPathChoice, lgNoFileZip As Long, _
                        Optional ByVal RecurseDirectories As Boolean = False, _
                        Optional excludeCRC32 As Boolean = False) As Long
    ' Ajoute un ou des fichiers du tableau (dont LocalFileHeader) au Zip ouvert
    
    ' FilePath           : Fichier avec son chemin complet
    ' PathChoice         : Définit si on veut garder le chemin du fichier
    ' lgNoFileZip        : N° de canal du fichier Zip ouvert par CreateZip
    ' RecurseDirectories : Définit s'il faut chercher les fichiers dans les sous-répertoires
    
    Dim NbFichiers As Long
    ' Par défaut, pas d'ajout de fichier
    ybFileAdd = 0
    
    ' Récupère par api le nom réel (au cas où on utilise des \..\
    '           ou bien si on ne précise pas de chemin)
    Dim buffer As String, Ret As Long
    buffer = Space(IIf(Len(FilePath) < 255, 255, Len(FilePath) + 1)) 'parfois des noms de fichiers provenance Internet sont > 255 de windows
    Ret = GetFullPathName(FilePath, Len(buffer), buffer, "")
    FilePath = Left(buffer, Ret)
    
    Dim StrFilter As String         'variable filtre de FindFirstFile ...
    If InStr(1, FilePath, "*") + InStr(1, FilePath, "?") = 0 Then ' il s'agit d'1 seul fichier à zipper
        'si le fichier existe bien
        If FichierExiste(FilePath) Then
            ' On l'insère parmi les fichiers à zipper
            NbFichiers = NbFichiers + FileAddSingle(FilePath, PathChoice, excludeCRC32)
        Else
            Exit Function
        End If
    Else    ' détection des fichiers concernés
        'fonction checkInsert : recherche les fichiers correspondants au filtre, avec gestion
        'option récursif ou pas, et ajoute les fichiers trouvés au tableau-collection à zipper.
        StrFilter = GetFileName(FilePath) 'extrait le nom du fichier et son extension avec ou pas des caractères génériques
                                          'pour servir de filtre aux Api FindFirstFile ...
        NbFichiers = NbFichiers + CheckInsert(Replace(FilePath, StrFilter, ""), StrFilter, 0, PathChoice, RecurseDirectories, excludeCRC32)
    End If  'traitement orienté 1 fichier ou plusieurs avec caractères génériques
    
    ' Retourne le nombre de fichiers ajoutés
    ybFileAdd = NbFichiers
    
    'Ajoute les éléments au fichier zip pour vider la mémoire au fur et à mesure car en cas de
    'zip volumineux il y aurait des problèmes mémoire !....
    Dim I As Long
    For I = UBound(tabFiles) - NbFichiers + 1 To UBound(tabFiles) 'peut aussi bien géré 1 fichier que des fichiers récursifs
        WriteLocalFileHeader lgNoFileZip, I
        Erase tabFiles(I).filedataZF    'efface les données comprimées du tableau de fichiers de la mémoire
                                        'mais pas le tableau à ce stade
    Next
    
End Function

Public Function WriteEndZip(lgNoFileZip As Long) As Boolean

    ' Termine le fichier ZIP lui même après ajouts des fichiers par ybFileAdd
    ' lgNoFileZip : N° de canal du fichier Zip ouvert par CreateZip

    Dim l As Long
    'Dim Fh As Long
    Dim I As Long
   
    l = Seek(lgNoFileZip)
    For I = 1 To UBound(tabFiles)
       WriteCentralFileHeader lgNoFileZip, I
    Next
    With colCentralDirEnd
        .EndOFCentralDirSignature = &H6054B50
        .EntriesInTheCentralDirThisOnDisk = UBound(tabFiles)
        .EntriesInTheCentralDir = .EntriesInTheCentralDirThisOnDisk
        .SizeOfCentralDir = Seek(lgNoFileZip) - l
        .OffSetOfCentralDir = l - 1
    End With
    Put #lgNoFileZip, , colCentralDirEnd    ' Ajoute la description de structure
    Put #lgNoFileZip, , ZipFileComment   ' Ajoute le commentaire général
    Close #lgNoFileZip
    Erase tabFiles
    cptTab = 0
    
    ' Signale que la création c'est bien passée
    WriteEndZip = True
   
End Function

Public Function FileAddIntoZip(strZipPath As String, strFilePath As String, _
                               PathChoice As enumPathChoice, _
                               Optional excludeCRC32 As Boolean = False, _
                               Optional OverWrite As Boolean = False) As Boolean
'Le 30/07/2006: FileAddIntoZip qui permet d'ajouter 1 fichier à zipper dans 1 zip déjà existant.
'à ne pas confondre avec FileAdd ou FileAddSingle qui ajoute à la collection ou au tableau qui
'vont être zippés.  Ce qui permet d'effacer la collection ou le tableau dès le départ pour ne pas
'traîner des données précédentes si la procédure appelante ne pas faire 1 set nothing.
'strZipPath = le fichier Zip auquel sera ajouté le fichier à zipper
'strFilePath = le fichier à zipper
'choix relatif au chemin selon énumération enumPathChoice
'Option OverWrite, par défaut à False, si OverWrite par de gestion de date, overwrite brut.

Dim CentralDirEndPos As Long    ', f As ZipFile
Dim hFile As Long           ' pour le handle Api read et writefile
Erase tabFiles      ' par sécurité !
     FileAddIntoZip = False 'initialise la fonction à non réussie pour les sorties avant fin de fonction
    'test existance du fichier zip par API createFile et ouverture en même tps pour API readfile et WriTeFile
    hFile = CreateFile(strZipPath, &H80000000, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)
    'Renvoi d'une erreur -> fichier inexistant puisque Flag OPEN_EXISTING
    If hFile = INVALID_HANDLE_VALUE Then
        CloseHandle hFile
        RaiseEvent FileZipErr(strZipPath, 10, "Ce fichier n'existe pas")
        Exit Function
    End If
    'fermeture fichier Zip ouvert par API (closeHandle)
    CloseHandle hFile
    ' Récupère l'offset de la structure descriptive du Zip : EndofCentralDirectoryRecord
    CentralDirEndPos = ybGetCentralDirEndPos(strZipPath)
    If CentralDirEndPos <= 0 Then
        RaiseEvent FileZipErr(strZipPath, 2, "Le fichier n'est pas un fichier au format Zip ou Sv")
        Exit Function
    End If
    ' C'est OK, on Ouvre le fichier Zip en binaire pour lecture indexée
    Fh = FreeFile
    Open strZipPath For Binary As #Fh
    Dim oldCentralDirectory As String, lgW As Long, lgAjt As Long
    ' Lit la structure descriptive du Zip : EndofCentralDirectoryRecord, pour récupérer offset de CentralDirectory
    ' et toutes les infos de typCenteralDirEnd dont le nombre de fichiers et le commentaire du Zip
    Call zReadCentralDirEnd(CentralDirEndPos)
    ' stocke le centralDirectory dans 1 variable pour le décaller ensuite lors de l'insertion du ou des nouveaux fichiers
    oldCentralDirectory = Space$(CentralDirEnd.SizeOfCentralDir)
    ' Se positionne au début de centralDirectory
    Seek #Fh, CentralDirEnd.OffSetOfCentralDir + 1
    Get #Fh, , oldCentralDirectory 'structure stockée
    '
    'Ajout 16/12/2006 : scan contenu du Zip pour gérer Overwrite ou pas si existence de fichier dans le zip
        Dim l As Long, strFile As String, strtest As String, Ret As Boolean
        'prépa pour test en fonction option PathChoice
        If PathChoice = WithoutPath Then ' sans le path
            strtest = GetFileName(strFilePath)
        Else                    'relatif ou absolu le path doit être comparé
            strtest = strFilePath
        End If
        Seek #Fh, CentralDirEnd.OffSetOfCentralDir + 1
        For l = 1 To CentralDirEnd.EntriesInTheCentralDir
            Ret = zReadCentralFileHeader(strFile)
            If UCase(strtest) = UCase(strFile) Then
                'fichier existe dans zip avec même gestion du chemin, on teste OverWrite
                If OverWrite Then   ' option écrasement
                    ' il faut effacer ce fichier existant (comparaison des dates négligée, à effectuer, boff !!!, on peut aussi vouloir écraser par + ancien !)
                    Close #Fh   'fermer ce fichier pour utiliser la procdure existante d'effacement
                    If FileRemoveFromZip(strZipPath, strtest) Then ' on efface
                        ' on réouvre le zip à traiter
                        Fh = FreeFile
                        Open strZipPath For Binary As #Fh
                        ' on reprend les nouveaux offsets
                        Call zReadCentralDirEnd(ybGetCentralDirEndPos(strZipPath))
                        oldCentralDirectory = Space$(CentralDirEnd.SizeOfCentralDir)
                        Seek #Fh, CentralDirEnd.OffSetOfCentralDir + 1
                        Get #Fh, , oldCentralDirectory 'structure stockée
                        Exit For ' on sort du test puisqu'on a écrasé le fichier, on va réécrire le nouveau dans la suite normale)
                    Else    ' si effacement a échoué, on gère comme option non overwrite
                        'déjà fermé #fh
                            'Info utilis (pas d'évènement, à voir à l'utilisation, il y a l'info pas Fait !)
                        Exit Function 'on sort
                    End If
                Else  ' si on écrase pas, on sort
                    Close #Fh
                    Exit Function
                End If
            End If
        Next
    '------------------------------------------------------------------
    '
    ' Se Repositionne au début de centralDirectory pour écrire le nouveau fichier à la place
    Seek #Fh, CentralDirEnd.OffSetOfCentralDir + 1
    'Ajoute un ou des fichiers à la collection et LocalFileHeader au Zip ouvert et recalcule le nouveau nbre
    lgAjt = ybFileAdd(strFilePath, PathChoice, Fh, , excludeCRC32)
    CentralDirEnd.EntriesInTheCentralDir = CentralDirEnd.EntriesInTheCentralDir + lgAjt
    'Réécrit CentralDirectory et EndofCentralDirectoryRecord rectifiés à 1 nouvelle fin du fichier Zip
    '- d'abord l'ancien centralDirectory (à la suite)
    lgW = Seek(Fh)
    Put #Fh, , oldCentralDirectory
    '- puis les nouveaux éléments du ou des nouveaux fichiers à CentralDirectory et EndofCentralDirectoryRecord
    '  après rectification de EndofCentralDirectoryRecord
    If lgAjt > 0 Then   'A ne réaliser que s'il y a eu ajout lors du passage ybfileadd (pourrait être à 0 =>Fichier inexistant, ....!)
        For cptTab = 1 To UBound(tabFiles)
            WriteCentralFileHeader Fh, cptTab
            DoEvents
        Next
    End If
    With CentralDirEnd
        .EndOFCentralDirSignature = &H6054B50
        .EntriesInTheCentralDirThisOnDisk = .EntriesInTheCentralDir
        .OffSetOfCentralDir = lgW - 1
        .SizeOfCentralDir = Seek(Fh) - lgW
    End With
    Put #Fh, , CentralDirEnd    ' Ajoute la description de structure
    Put #Fh, , mZipComment   ' Ajoute le commentaire général
    Close #Fh
    'efface le tableau qui ne servira + dans le traitement et évitera un cumul des données
    'si la classe n'était pas fermée dans la procédure appelante
    Erase tabFiles
    
    FileAddIntoZip = True
    If lgAjt < 1 Then FileAddIntoZip = False 'on force à faux puisqu'il n'y a pas eu d'ajout, bien que la fonction se soit déroulée normalement
End Function

Public Function FileRemoveFromZip(strZipPath As String, strFilePath As String) As Boolean
'Le 01/08/2006: FileRemoveFromZip qui permet de supprimer 1 fichier zippé d'1 zip déjà existant.
'à ne pas confondre avec FileDel qui supprime de la collection ou au tableau 1 fichier qui
'va être zippé.  Ce qui permet d'effacer la collection ou le tableau dès le départ pour ne pas
'traîner des données précédentes si la procédure appelante ne pas faire 1 set nothing.
'strZipPath = le fichier Zip duquel sera supprimé le fichier zippé
'strFilePath = le fichier fichier zippé à retirer

Dim CentralDirEndPos As Long
Dim hFile As Long           ' pour le handle Api read et writefile
Erase tabFiles          ' par sécurité !
     FileRemoveFromZip = False 'initialise la fonction à non réussie pour les sorties avant fin de fonction
    'test existance du fichier zip par API createFile et ouverture en même tps pour API readfile et WriTeFile
    hFile = CreateFile(strZipPath, &H80000000, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)
    'Renvoi d'une erreur -> fichier inexistant puisque Flag OPEN_EXISTING
    If hFile = INVALID_HANDLE_VALUE Then
        CloseHandle hFile
        RaiseEvent FileZipErr(strZipPath, 10, "Ce fichier n'existe pas")
        Exit Function
    End If
    'fermeture fichier Zip ouvert par API (closeHandle)
    CloseHandle hFile
    ' Récupère l'offset de la structure descriptive du Zip : EndofCentralDirectoryRecord
    CentralDirEndPos = ybGetCentralDirEndPos(strZipPath)
    If CentralDirEndPos <= 0 Then
        RaiseEvent FileZipErr(strZipPath, 2, "Le fichier n'est pas un fichier au format Zip ou Sv")
        Exit Function
    End If
    ' Ouverture du fichier Zip en binaire pour lecture indexée
    RaiseEvent Status("Suppression élément interne à fichier comprimé démarrée.")
    Fh = FreeFile
    Open strZipPath For Binary As #Fh
    Dim lgW As Long, I As Long, Ret As Boolean, sFileName As String, lgQ As Long
    Dim lgOffData As Long, lgSizeData As Long, lgSizeCdirect As Long, lgOffCdirect As Long
    Dim strData As String, lgOffsetCD As Long, Fn As Integer
    ' Lit la structure descriptive du Zip : EndofCentralDirectoryRecord, pour récupérer offset de CentralDirectory
    ' et toutes les infos de typCenteralDirEnd dont le nombre de fichiers et le commentaire du Zip
    Call zReadCentralDirEnd(CentralDirEndPos)
    On Error GoTo ErFRFZ
        ' Se Repositionne au début de centralDirectory pour le lire
        Seek #Fh, CentralDirEnd.OffSetOfCentralDir + 1
        For I = 1 To CentralDirEnd.EntriesInTheCentralDir
            Ret = zReadCentralFileHeader(sFileName)
            ' Les sous-répertoires inclus dans le Zip ont une taille compressée
            '   et décompressée de zéro octets. On recherche l'entrée suivante
            Do While Ret And (varCommunZF.UncompressedSize = 0)
                Ret = zReadCentralFileHeader(sFileName)
                DoEvents
            Loop
            'analyse si on est le fichier qui nous interesse
            If UCase(sFileName) = UCase(strFilePath) Then 'si ok
                lgQ = I 'stocke le quantième du fichier supprimé pour rectif + loin du CD  (ajout 17/12/06 correct. bug)
                'il nous faut récupérer son offset et sa longueur avec son localFileHeader sans DataDescriptor
                'pour aller réécrire ce qui le suit à sa place
                With varCommunZF
                    lgOffData = CentralFileHeaderC.RelativeOffsetOfLocalHeader 'son offset Data
                    lgSizeData = (30 + .FileNameLength + .ExtraFieldLength) + .CompressedSize  'longueur localFileHeader+data
                    'il faut aussi récupérer son offset dans centralDirectory pour même opération sur centralDirectory
                    lgSizeCdirect = Len(CentralFileHeaderA) + Len(varCommunZF) + Len(CentralFileHeaderC) + .FileNameLength + .ExtraFieldLength + CentralFileHeaderC.FileCommentLength 'sa longueur centralDirectory
                    lgOffCdirect = Seek(Fh) - 1 - lgSizeCdirect               'son offset centralDirectory
                End With
                'Réécrit les morceaux => attention gros fichiers, les gérer par erreur, en tranches
                Fn = FreeFile
                Open strZipPath & "~" For Binary As #Fn
                    'du début des données au fichier supprimé
                Dim lgTemp As Long  ' var prov pour gérer erreurs gros fichier
                If lgOffData > 0 Then   'si le fichier supprimé n'est pas le 1°
                lgTemp = lgOffData
                    strData = Space$(lgTemp) 'dimensionne la variable de récup
                    Seek #Fh, 1     'se positionne au début du fichier
                    Get #Fh, , strData                'stocke les données
                    Put #Fn, , strData          'réécrit les 1° données
                ElseIf lgOffData < 0 Then   'erreur
                    Close #Fn
                    Kill strZipPath & "~"
                    Close #Fh
                    Exit Function
                End If
                    'du fichier suivant le fichier supprimé à la fin des données
                Seek #Fh, lgOffData + lgSizeData + 1   'se positionne à la suite du fichier à effacer
                lgTemp = CentralDirEnd.OffSetOfCentralDir - (lgOffData + lgSizeData)
                strData = Space$(lgTemp) 'dimensionne la variable de récup
                Get #Fh, , strData                'stocke lesdonnées qui suivent fichier effacé au centralDirectory
                Put #Fn, , strData          'réécrit les données suivantes sur le fichier à effacer
                        'ça optimise et économise des ressources d'écrire sur le même fichier plutôt que nouveau fichier et effacer l'ancien !
                lgOffsetCD = Seek(Fn)                  'récupère le nouvel offset de centralDirectory
                    'du début du centralDiretory à l'offset du fichier supprimé dans centralDir..
                strData = Space$(lgOffCdirect - CentralDirEnd.OffSetOfCentralDir) 'dimensionne la variable de récup
                Seek #Fh, CentralDirEnd.OffSetOfCentralDir + 1    'se positionne au début du centralDir...
                Get #Fh, , strData                'stocke les infos du centralDir... qui précédent celles du fichier
                Put #Fn, , strData          'réécrit début du centralDirec....
                lgW = Seek(Fh)                  'récupère où écrire la fin de centralDirectory
                    'de la suite du fichier supprimé dans centralDiretory à la fin de centralDir..
                strData = Space$((CentralDirEndPos - 1) - (lgOffCdirect + lgSizeCdirect)) 'dimensionne la variable de récup     -1 !!!!!!!!!!!!!!!!!
                Seek #Fh, (lgOffCdirect + lgSizeCdirect) + 1    'se positionne à la suite du fichier supprimé dans centralDir...
                Get #Fh, , strData                'stocke les infos du centralDir... qui précédent celles du fichier
                Put #Fn, , strData          'réécrit la suite du centralDirec....
                lgW = Seek(Fh)                  'récupère la position de la fin de centralDirectory
                'il faut rectifier rectifier centralDirEnd du nombre de fichiers et ....
                With CentralDirEnd
                    .EndOFCentralDirSignature = &H6054B50
                    .EntriesInTheCentralDir = .EntriesInTheCentralDir - 1 '- le fichier supprimé
                    .EntriesInTheCentralDirThisOnDisk = .EntriesInTheCentralDir
                    .OffSetOfCentralDir = lgOffsetCD - 1
                    .SizeOfCentralDir = Seek(Fn) - lgOffsetCD
                End With
                    'la fin du fichier : CentralDirEnd rectifié et commentaire
                Close #Fh
                Put #Fn, , CentralDirEnd    ' Ajoute la description de structure
                Put #Fn, , mZipComment   ' Ajoute le commentaire général
                'avant de fermer le nouveau fichier, il faut rectifier tous les offsetlocalHeader du
                'centralDirectory en les réduisant de la longueur totale occupée par le fichier supprimé
                Seek #Fn, lgOffsetCD
                For lgW = 1 To CentralDirEnd.EntriesInTheCentralDir
                    ' CentralFileHeader contient :
                    ' - soit les données si ce sont des chiffres,
                    ' - soit les offsets (décalage) de chacune des variables chaine qui le compose :
                    lgOffsetCD = Seek(Fn)
                    Get #Fn, , CentralFileHeaderA    ' Lit la structure
                    Get #Fn, , varCommunZF           ' Lit la structure
                    Get #Fn, , CentralFileHeaderC    ' Lit la structure
                    With varCommunZF
                        If CentralFileHeaderA.CentralFileHeaderSigniature = CentralFileHeaderSigniature Then
                            'rectifie le nouvel offset du fichier
                            'ATTENTION UNIQUEMENT de ceux qui suivent le fichier supprimé (ajout 17/12/06 correct. bug)
                            If lgW > (lgQ - 1) Then
                                CentralFileHeaderC.RelativeOffsetOfLocalHeader = CentralFileHeaderC.RelativeOffsetOfLocalHeader - lgSizeData
                            End If
                            'se repositionne au début de la structure pour la réécrire
                            Seek #Fn, lgOffsetCD
                            Put #Fn, , CentralFileHeaderA
                            Put #Fn, , varCommunZF
                            Put #Fn, , CentralFileHeaderC
                            ' puis lit les chaines variables en fonction de leurs tailles pour avancer dans centralDirect..
                            strData = Space(.FileNameLength)
                            Get #Fn, , strData
                            strData = Space(.ExtraFieldLength)
                            Get #Fn, , strData
                            strData = Space(CentralFileHeaderC.FileCommentLength)
                            Get #Fn, , strData
                        Else
                            Close #Fn
                            Kill strZipPath & "~"
                            Exit Function
                        End If
                    End With
                Next
                'fermer le nouveau fichier
                Close #Fn
                'tout est fini correct, remplacer le fichier
                ZipClose    'car on a fermé le fichier zip ouvert dans la procédure, mais il peut aussi être ouvert pour gestion en amont !
                Kill strZipPath
                Name strZipPath & "~" As strZipPath
                FileRemoveFromZip = True    'signale màj finie correcte
                RaiseEvent Status("Suppression élément interne à fichier comprimé terminée.")
                Exit For    'sort de la boucle
            End If
        Next
Erase tabFiles
'pour la gestion d'erreur, au minimum traitement des gros fichiers par lots
Exit Function
ErFRFZ:
Dim lgTranche As Long, a As Byte
Select Case Err.Number
    Case 7, 14 'problème mémoire, chaine trop grande => trop gros fichier
        RaiseEvent Status("Suppression élément interne à fichier comprimé de grande taille : traitement long")
        'l'évènement permettra d'enclancher 1 souris occupé ou 1 info d'attente dans la proc appelante !
        DoEvents
        lgTranche = lgTemp \ 10
        strData = Space$(lgTranche)    'ah, la mémoire va mieux
        For a = 1 To 9
            Get #Fh, , strData                'stocke lesdonnées qui suivent fichier effacé au centralDirectory
            DoEvents
            Put #Fn, , strData          'réécrit les données suivantes sur le fichier à effacer
            DoEvents
        Next
        strData = Space$(lgTemp - (lgTranche * 9)) ' on redimensionne le dernièr paquet avec le solde des paquets arrondis
        If Err.Number = 7 Then Resume      ' on repart là où on était bloqué par la taille mémoire
        If Err.Number = 14 Then Resume Next ' on repart à l'instruction suivante où on était bloqué puisqu'on a réalisé les instructions
    Case Else
        'cas non traité on annule l'opération proprement
        Close #Fn
        Kill strZipPath & "~"
        Close #Fh
        Exit Function
End Select
End Function

Public Function ExtractSgFileToTmp(lgNuFile As Long) As String
' décompresse au plus vite, controles allégés,un fichier dont le n° dans le zip est
' passé en argument, dans le répertoire temporaire de Windows dans le but d'obtenir
' rapidemment des infos sur ce fichier ou de le charger en fonction de son extension.
' le fichier étant destiné à être détruit la fonction renvoie son chemin et son nom
' ou 1 chaine vide si la fonction a échouée.
' la fonction sous-entend que ZipOpen a été exécutée et ne le teste pas (objectif rapidité)
' et CentralDirEndPos identifiée. (de toutes façon si on connaît le n° du fichier !...)
' L'extraction étant temporaire elle sera overwrite si 1 même fichier existe
' et le répertoire d'origine ne sera pas préservé.

Dim strWinTmp As String, lgLenWinTmp As Long
Dim Buf As String * 128
Screen.MousePointer = 11
'récupère le répertoire Temp de windows
lgLenWinTmp = GetTempPath(128, Buf)
If lgLenWinTmp = 0 Then
    ExtractSgFileToTmp = ""
    Screen.MousePointer = 0
    Exit Function
Else
   strWinTmp = Left(Buf, lgLenWinTmp)
End If

Dim l As Long, Ret As Boolean, sFileName As String
' Lecture des entêtes de chaque fichier pour trouver le bon
If zReadCentralDirEnd(CentralDirEndPos) Then
    Seek #Fh, CentralDirEnd.OffSetOfCentralDir + 1
    For l = 1 To mFileCount
        Ret = zReadCentralFileHeader(sFileName)
        ' Rejet des sous-répertoires inclus dans le Zip
        Do While Ret And (CentralFileHeaderC.InternalFileAttributes = FILE_ATTRIBUTE_DIRECTORY)
            Ret = zReadCentralFileHeader(sFileName)
            DoEvents
        Loop
        If Ret Then ' si on n'est pas arrivé à la fin des données
            If l = lgNuFile Then
                ' On ne garde que le nom du fichier
                sFileName = zGetFileName(sFileName)
                    'Récupération de la dernière date de modif pour la restituer au fichier
                    Dim maDateFT As FILETIME, maDateLT As FILETIME
                    ' Obtient la date et l'heure Dos stockée lors de la compression et la convertit en format FileTime pour conversion
                    Call DosDateTimeToFileTime(mFileInfos(l).LastModFileDate, mFileInfos(l).LastModFileTime, maDateLT)
                    ' convertit ensuite cette date locale à la date Filetime
                    Call LocalFileTimeToFileTime(maDateLT, maDateFT)
                    '
                If FichierExiste(strWinTmp & sFileName) Then Kill strWinTmp & sFileName
                If zExtractFile(strWinTmp & sFileName) = False Then
                    ExtractSgFileToTmp = ""
                    Screen.MousePointer = 0
                    Exit Function
                Else
                    'remise en place de la date de dernière modif du fichier récupérée au plus haut
                    ' Ouvre le fichier en écriture
                    Dim hFile As Long, DateJ As FILETIME, sysDateJ As SYSTEMTIME
                    hFile = CreateFile(strWinTmp & sFileName, GENERIC_WRITE, FILE_SHARE_WRITE Or FILE_SHARE_READ, ByVal CLng(0), OPEN_EXISTING, 0, 0)
                    If hFile <> INVALID_HANDLE_VALUE Then ' s'il n'y a pas d'erreur d'accès au fichier
                        ' On change les dates du fichier, et voilà c'est mieux maintenant !
                        
                        'retaitement date du jour pour date écriture
                        sysDateJ.wDay = Day(Now)
                        sysDateJ.wMonth = Month(Now)
                        sysDateJ.wYear = Year(Now)
                        sysDateJ.wHour = Hour(Now)
                        sysDateJ.wMinute = Minute(Now)
                        sysDateJ.wSecond = Second(Now)
                        Call SystemTimeToFileTime(sysDateJ, DateJ)
                        Call LocalFileTimeToFileTime(DateJ, DateJ)
                        '
                        Call SetFileTime(hFile, maDateFT, DateJ, maDateFT) 'création,dernier accès, modification
                        ' on ferme le fichier ouvert
                        CloseHandle (hFile)
                    End If  'sinon les dates ne seront pas corrigées
                End If
                Exit For    'inutile de continuer même pour le callage des offsets pour les autres opé, elles relisent depuis centraldirend
            End If
        End If
    Next
End If

ExtractSgFileToTmp = strWinTmp & sFileName
Screen.MousePointer = 0
End Function

Public Function LoadZippedFile(strFilePath As String) As Boolean
'Lance l'exécutable attaché au fichier zippé dans 1 fenêtre externe sans décompression apparente,
'et donc sans modification du zip. En fait cette méthode s'emploie en complément de ExtractSgFileToTmp
'qui doit être réalisée avant et qui décompresse rapidement le ficher dans le répertoire Temp de Windows.
'la fonction renvoie vrai ou faux si elle a réussie + 1 évènement en cas d'échec.

LoadZippedFile = False 'par défaut au départ
Dim strEXE As String, lgRep As Long, strPath As String
strEXE = Space$(250)  ' Initialisation du buffer pour récupérer l'exe
strPath = zGetFilePath(strFilePath)
' Appel à la fonction
lgRep = FindExecutable(Replace(strFilePath, strPath & "\", ""), strPath, strEXE)
' Traitement de la valeur de retour
lgRep = InStr(1, strEXE, vbNullChar)
If lgRep = 0 Then   ' si non trouvé
    'est déjà par défaut à faux donc on renvoie juste 1 évènements erreur et on sort !
    RaiseEvent ZipError(zeNoexecutable, "Il n'y a pas d'exécutable attaché à ce type de fichier dans le système d'exploitation !", False)
    Exit Function
Else
    'retraite la chaine reçue
    strEXE = Left$(strEXE, InStr(1, strEXE, vbNullChar) - 1)
    If Len(strEXE) < 5 Then 'l'appli à lancer doit au minimum avoir 1 nom + le point et l'extension, soit au moins 5 caractères
        RaiseEvent ZipError(zeNoexecutable, "Il n'y a pas d'exécutable attaché à ce type de fichier dans le système d'exploitation !", False)
        Exit Function
    End If
'''    'lance l'exe par shell de vb
'''    Dim dblR As Double
'''    dblR = Shell(strEXE & " " & Chr(34) & strFilePath & Chr(34), vbNormalFocus)
    'lance l'exe plutôt avec l'API shellExecute
    lgRep = ShellExecute(Screen.ActiveForm.hwnd, "Open", strFilePath, vbNullString, vbNullString, vbNormalFocus)
    'Pour la destruction du fichier provisoire décompressé strfilepath qui ne peut être réalisé tant que l'exe lancé
    'tourne et dont on ne peut pas connaître l'arrêt ou alors il aurait fallu le lancer par 1 autre API mais comme on souhaite
    'éventuellement pouvoir sortir de la classe zip et du programme l'utilisant avant que l'exe lancé ne cesse, ce n'était pas
    'possible, il faudra donc détruire le fichier décompressé en différé. J'utiliserais la façon suivante :
    'je place dans 1 tableau niveau module le ou les chemin(s) fichier(s) à détruire et lors de l'évènement Terminate de
    'la classe je crée un batch DOS .bat qui détruira le ou les fichiers lors du prochain redémarrage de l'ordi. J'utilise
    'un fichier DOS plutôt qu'un VBS que certains antivirus rejettent alors que le bon vieux DOS fonctionnera tout seul et
    'pourra s'autodétruire au démarrage de l'ordi. Pour ce faire, je lance le .Bat par la clé RunOnce dans la Bdr.
    '(voir class_Terminate)
    ReDim Preserve arrDelFile(UBound(arrDelFile) + 1)
    arrDelFile(UBound(arrDelFile)) = strFilePath
End If



End Function

'################### anciennes fonctions ZipFile #################################################

Private Function GetDOSDate(ByVal ModDate As Date) As Integer

    Dim Day As Long
    Dim Month As Long
    Dim Year As Long
    
    Dim b(0 To 1) As Byte
    
    'There's a Windows API Function FileTimeToDosDateTime
    'but I couldn't get it to work so I did this the hard way

    'Bits    Contents
    '0–4     Day of the month (1–31)
    '5–8     Month (1 = January, 2 = February, etc.)
    '9–15    Year offset from 1980 (add 1980 to get actual year)
    
    Day = DatePart("d", ModDate)
    Month = DatePart("m", ModDate)
    Year = DatePart("yyyy", ModDate) - 1980
    
    If GetBit(Day, 0) Then SetBit b(0), 0
    If GetBit(Day, 1) Then SetBit b(0), 1
    If GetBit(Day, 2) Then SetBit b(0), 2
    If GetBit(Day, 3) Then SetBit b(0), 3
    If GetBit(Day, 4) Then SetBit b(0), 4
    If GetBit(Month, 0) Then SetBit b(0), 5
    If GetBit(Month, 1) Then SetBit b(0), 6
    If GetBit(Month, 2) Then SetBit b(0), 7
    If GetBit(Month, 3) Then SetBit b(1), 0
    If GetBit(Year, 0) Then SetBit b(1), 1
    If GetBit(Year, 1) Then SetBit b(1), 2
    If GetBit(Year, 2) Then SetBit b(1), 3
    If GetBit(Year, 3) Then SetBit b(1), 4
    If GetBit(Year, 4) Then SetBit b(1), 5
    If GetBit(Year, 5) Then SetBit b(1), 6
    If GetBit(Year, 6) Then SetBit b(1), 7
    
    CopyMemory_Int_Byte GetDOSDate, b(0), 2
    
End Function

Private Function GetDOSTime(ByVal ModDate As Date) As Integer

    Dim Second As Long
    Dim Minute As Long
    Dim Hour As Long
    
    Dim b(0 To 1) As Byte
    
    'Bits    Contents
    '0–4     Second divided by 2
    '5–10    Minute (0–59)
    '11–15   Hour (0–23 on a 24-hour clock)
    
    Second = DatePart("s", ModDate) \ 2
    Minute = DatePart("n", ModDate)
    Hour = DatePart("h", ModDate)
    
    If GetBit(Second, 0) Then SetBit b(0), 0
    If GetBit(Second, 1) Then SetBit b(0), 1
    If GetBit(Second, 2) Then SetBit b(0), 2
    If GetBit(Second, 3) Then SetBit b(0), 3
    If GetBit(Second, 4) Then SetBit b(0), 4
    If GetBit(Minute, 0) Then SetBit b(0), 5
    If GetBit(Minute, 1) Then SetBit b(0), 6
    If GetBit(Minute, 2) Then SetBit b(0), 7
    If GetBit(Minute, 3) Then SetBit b(1), 0
    If GetBit(Minute, 4) Then SetBit b(1), 1
    If GetBit(Minute, 5) Then SetBit b(1), 2
    If GetBit(Hour, 0) Then SetBit b(1), 3
    If GetBit(Hour, 1) Then SetBit b(1), 4
    If GetBit(Hour, 2) Then SetBit b(1), 5
    If GetBit(Hour, 3) Then SetBit b(1), 6
    If GetBit(Hour, 4) Then SetBit b(1), 7
    
    CopyMemory_Int_Byte GetDOSTime, b(0), 2
    
End Function

Private Function GetFileName(ByVal FilePath As String) As String

    Dim a() As String
    a = Split(FilePath, "\")
    GetFileName = a(UBound(a))

End Function

Public Function Conv_ISO_ASCII(ByVal sChaine As String) As String
    
    ' Convertit les codes des caractères ISO 8859-1 en code de caractères ASCII étendu
    ' (codes ascii > 127 : les caractères accentués ne sont pas aux mêmes places)
    ' Voir les tableaux Table_Conversion_ASCII_ISO() et Table_Conversion_ISO_ASCII()
    ' sChaine est passée ByVal pour ne pas être modifiée ici.
    
    ' La fonction inverse se trouve dans la clsZip_UnZip : Conv_ASCII_ISO
    
    Dim r As Long
    Dim t As Byte
    For r = 1 To Len(sChaine)
        t = Asc(Mid$(sChaine, r, 1))
        Mid$(sChaine, r, 1) = Chr$(Table_Conversion_ISO_ASCII(t))
    Next r
    Conv_ISO_ASCII = sChaine
    
End Function

Private Function CompressBytes(ByVal FilePath As String, _
                          ByRef FileData() As Byte, _
                          ByRef CompressedSize As Long, _
                          ByRef UncompressedSize As Long, _
                          Optional excludeCRC32 As Boolean = False) As Long
' retourne 0 ou 1 code erreur
' rajout le 23/11/2006: option exclusion du calcul CRC pour gain de vitesse au détriment
'                       de la fiabilité puisqu'il n'y a + de controle CRC, mais n'en faisont pas tout 1 plat !
                          
'Fonctionne mais pas pour les gros fichiers  => Ajout d'une gestion d'erreur

    Dim buffer() As Byte
    Dim BufferSize As Long
    Dim FileSize As Long
    Dim crc As Long
    Dim Fh As Long
    Dim r As Long

On Error GoTo Koinc

    CompressBytes = 0 'par défaut
    Fh = FreeFile
    Open FilePath For Binary As #Fh
        FileSize = LOF(Fh)
    
    If FileSize > 0 Then
        ReDim FileData(FileSize - 1)  ' c'est ici qu'on part en gestion gros fichiers si Erreur capa Mémoire
    Else
        ReDim FileData(0)
    End If
    
    Get #Fh, , FileData
    Close #Fh

    If excludeCRC32 = False Then    'cas avec calcul du CRC
        crc = Crc32_Byte(0&, FileData(0), UBound(FileData) + 1)
    Else
        crc = 0
    End If
    
'''''''' Comme il n'y a aucun intérêt à essayer de rezipper un fichier déjà zippé, on en élimine ici
'''''''' le traitement de zippage, on empactera directement le fichier zip dans le nouveau zip.
'''''''' on gagne ainsi des ressources et du temps machine. (rezipper un zip n'apporte pas de gain
'''''''' de place, excepté si on compresse avec un niveau supérieur à celui du 1° zippage.

'''''''' Vrai en théorie, mais après tests, il s'avère que le temps est légèrement + long si on empacte
'''''''' direct. En fait l'éco de temps réalisée est bouffée par le test sur tous les fichiers pour
'''''''' savoir si on est un zip. Je laisse ici ce commentaire pour s'en souvenir et ne pas retenter
'''''''' cette opération + tard !

'''''''If InStrRev(sFileName, ".zip", , vbBinaryCompare) > 0 Then 'si le nom de fichier comporte .Zip
'''''''     BufferSize = FileSize
'''''''Else    'tout autre fichier que zip, on passe par la compression
    BufferSize = FileSize * 1.01 + 12
    ReDim buffer(BufferSize)
    
'----------------- Utilisation Deflate de zLib -----------------
buffer = CompressDeflate(FileData, FileSize, , r)
If r = Z_OK Then  'si compression OK
    BufferSize = UBound(buffer) + 1
    ReDim FileData(BufferSize - 1)
    FileData = buffer
    GoTo Ftest
Else
    CompressBytes = r
    Exit Function
End If
'---------------------------------------------------------------
    
    r = Compress_Byte_Long_Byte(buffer(0), BufferSize, FileData(0), FileSize)
    BufferSize = BufferSize - 6
    
    'When using the Compress method, ZLib adds a 2 byte head
    'and a 4 byte tail. The head must be removed for zip
    'compatability and the tail is not necessary.
    
    ReDim FileData(BufferSize - 1)
    CopyMemory_Byte_Byte FileData(0), buffer(2), BufferSize

Ftest:      ' fin commune à Deflate ou compress_Byte_Long_Byte
'''''''''End If      'test sur zippage d'un fichier déjà zippé
    Erase buffer

    varCommunZF.CRC32 = crc
    
    UncompressedSize = FileSize
    CompressedSize = BufferSize
    
    Exit Function
Koinc:
    'puisqu'il y a 1 erreur, on considère (dans 1 premier temps) que le fichier ne sera pas traité ici
    'mais qu'on l'empactera peut-être, on positionne donc le paramètre UncompressedSize et on
    'compressedSize à la taille du fichier afin que ces élèments soient servis si l'erreur était
    'mettra à jour les paramètres manquants commun dans fonction et proc bigFiles,WriteLocalFileHeader.
    CompressBytes = Err.Number
    If Err.Number <> 70 Then    'si l'erreur est 70 violation de données, on écrira rien en sortie donc inutile de perdre du temps
        ' mais sinon renseigner les éléments du header du zip
        UncompressedSize = FileLen(FilePath) 'je sais on l'a par Lof(), mais l'erreur est peut-être d'ouverture
        ' La gestion du gros fichier et son empactage se terminera directement dans WriteLocalFileHeader
    End If
    'on ferme éventuellement le fichier ouvert et on efface éventuel buffer créé
    Close #Fh
    Erase buffer
    Exit Function

End Function

Public Function CompressDeflate(ByRef wStream() As Byte, lgSizeSource As Long, _
                        Optional Block As Boolean = False, Optional codeError As Long = 0) As Variant
'wStream            = tableau des données à compresser
'lgSizeSource       = longueur des données à compresser
'Block (optionnel)  = paramètre optional par défaut à False, signifie qu'on est 1 partie d'1 fichier :
'                     ne pas écrire deflateEnd dans ce cas.
'codeError          = résultat code erreur Deflate correspondant à eZipError
'La fonction compresse et renvoie 1 tableau des données compressées.

    Dim vInfo   As zStream
    Dim arrTarget()  As Byte  'arrTarget      = tableau des données compressées

'    Dim lgSizeSource   As Long
'    lgSizeSource = UBound(wStream)        'version exemple d'origine mais c'est faux il manque +1, mais fait bugguer
On Error GoTo errDeflate
    If lgSizeSource Then
'        ReDim arrTarget(lgSizeSource * 1.5)
        ReDim arrTarget(lgSizeSource * 1.5 + 12)
        With vInfo
            .next_in = VarPtr(wStream(LBound(wStream)))   'obtient adresse du tableau
            .avail_in = lgSizeSource '+ 1      'version exemple origine avec +1, pourquoi ?
            .next_out = VarPtr(arrTarget(0))     'obtient adresse du tableau
'            .avail_out = lgSizeSource + 1.5        '???+1.5 pourquoi pas *1.5
            .avail_out = lgSizeSource * 1.5 + 12
            .data_type = Z_BINARY
            deflateInit2 vInfo, m_LevelComp, Z_DEFLATED, -MAX_WBITS, MAX_MEM_LEVEL, Z_DEFAULT_STRATEGY, ZLIB_VERSION, Len(vInfo)
            If Block = False Then   'orientation par blocks ou pas
                codeError = deflate(vInfo, Z_FINISH)
                If codeError < 0 Then     'tester erreur <0
                    Erase arrTarget 'efface le tableau de sortie au cas où il y aurait quelquechose dedans !
                    codeError = zeDeflate
               Else    'erreur
                    codeError = deflateEnd(vInfo)   'renvoie z_OK si correct
                End If
            Else            ' **************** PARTIE A REVOIR (blocks pour gros fichiers)*******
                codeError = deflate(vInfo, Z_FINISH)
                If codeError < 0 Then     'tester erreur < 0
                    Erase arrTarget 'efface le tableau de sortie au cas où il y aurait quelquechose dedans !
                    codeError = zeDeflate
                End If
                            ' **************** FIN PARTIE A REVOIR ******************************
            End If
            If .total_out Then  's'il y a 1 donnée, redimensionne le tableau
                ReDim Preserve arrTarget(.total_out - 1)
                CompressDeflate = arrTarget
            End If
        End With
    Else    'si lgSizeSource =0 'pas de compress mais reprendre wStream() en sortie (cas où le tableau = 0 octet, ubound du tableau = 0, il y aurait 1 bug sans ce traitement)
        CompressDeflate = wStream
        codeError = 0
    End If
    Exit Function
    
errDeflate:
Debug.Print Err.Description & Err.Number & " " & lgSizeSource
Erase arrTarget 'efface le tableau de sortie au cas où il y aurait quelquechose dedans !
codeError = 17 'code erreur
End Function

Private Sub SetBit(ByRef b As Byte, ByVal Bit As Long)

    b = b Or (2 ^ Bit)

End Sub

Private Function GetBit(ByVal l As Long, ByVal Bit As Long) As Boolean

    GetBit = ((l And 2 ^ Bit) > 0)

End Function

Private Function bigFiles(ByVal FilePath As String, lgIndex As Long, Optional lgNoFile As Long = 0) As Long
'Calcule (octet par octet) le Crc32 à partir du nom d'1 fichier au lieu d'1 tableau d'octets en entrée dans Zlib
'Eventuellement écrit, par blocks, dans 1 fichier dont le canal est communiqué en paramètre optionnel,
'les données lues durant cette opération, pour simple Empactage. Dans ce cas, on économise 1 lecture,
'et on optimise la vitesse d'écriture par les blocks.
'lgIndex est la valeur de l'index tableau du fichier misà jour.

Dim hOrgFile As Long
Dim nSize As Long
Dim Ret As Long
Dim I As Long   'long pour capacité importante
Dim lgCRC As Long
Dim bytT As Byte
Dim bytC As Byte
Dim lgTmp As Long
Dim arrByte() As Byte   'tableau d'octets si on écrit les données dans 1 autre fichier
Dim lgmem As MEMORYSTATUS
Dim lgMax As Long       'valeur maxi à prendre en mémoire
Dim intFois As Integer  'nbre de passages
    bigFiles = 0    'par défaut au départ
    tabFiles(lgIndex).commun.CompressionMethod = 0  'pas de compression
    tabFiles(lgIndex).commun.CompressedSize = tabFiles(lgIndex).commun.UncompressedSize
    'Récupère un Handle pour manipuler le fichier
    hOrgFile = CreateFile(FilePath, &H80000000, FILE_SHARE_READ Or FILE_SHARE_WRITE, _
                            ByVal 0&, OPEN_EXISTING, 0, 0)
    'Renvoi d'une erreur -> Problème d'accès au fichier
    If hOrgFile = INVALID_HANDLE_VALUE Then
        RaiseEvent ZipError(zeUnKnownError, "Le fichier " & FilePath & " n'est pas ou plus accessible !", False)
        DoEvents
        Exit Function
    End If
    'Taille du fichier
    nSize = GetFileSize(hOrgFile, 0)
    'Attention, GetFileSize peut renvoyer une taille de fichier maximum de : 2,147,483,647 bytes.
    If nSize = ERROR Then
        RaiseEvent ZipError(zeUnKnownError, "La Taille du fichier " & FilePath & _
        " dépasse celle prévue pour la FAT32 !", False)
        DoEvents
        Exit Function
    End If
   'Initialise le pointeur sur le fichier
    SetFilePointer hOrgFile, 0, 0, FILE_BEGIN
    
    'si on doit réécrire les données lues du fichier dans 1 autre fichier
    'via 1 tableau d'octets pour aller + vite
    If lgNoFile > 0 Then
        'recherche le potentiel de mémoire dispo pour redimensionner au mieux 1 tableau(compris taille fichier/capa mémoire)
        GlobalMemoryStatus lgmem
        lgMax = lgmem.dwAvailPhys / 4   'mémoire physique dispo/4 pour sécu sur le reste appli et les autres applis
        If nSize > lgMax Then   'dimension des tableaux d'octets, si taille fichier > à mémoire raisonnable à utiliser
            ReDim arrByte(1 To lgMax)
        Else    ' si taille fichier <= utilisation raisonnable mémoire
            ReDim arrByte(1 To nSize)
            lgMax = nSize
        End If
        intFois = 0 'initialise la variable à 0
    End If
    '
    On Error GoTo BigSize
    ' Une table de correspondance des 255 caractères ASCII pour le calcul du CRC a été initialisée
    ' lors de l'initialisation de la classe (donc 1 seule fois) dans la variable tableau arrCRC32
    ' initialisation de  lgCRC = 0xffffffff
    ' pour chaque octet du fichier, on fait :
    ' calcule lgCRC = (lgCRC >> 8) ^ Table[(lgCRC & 0xFF) ^ byte]
    ' on renvoie le CRC32 au dernier octet => lgCRC ^ 0xffffffff
    
    lgCRC = &HFFFFFFFF
    For I = 1 To nSize
        'lit le contenu du fichier octet par octet avec la variable
        ReadFile hOrgFile, bytC, 1, Ret, ByVal 0&
        bytT = (lgCRC And &HFF) Xor bytC
        lgTmp = (lgCRC And &H7FFFFFFF) \ 256
        If (lgCRC And &H80000000) <> 0 Then
            lgTmp = lgTmp Or &H800000
        End If
        lgCRC = lgTmp Xor arrCRC32(bytT)
        If I Mod 5000 = 0 Then DoEvents 'fait 1 doevents lorsque i est divisible par 5000
        If lgNoFile > 0 Then
            'insertion dans le tableau -> (pour éviter 1 test à chq octet on va gérer quand le tableau
            'sera plein et donc l'écriture dans le fichier, par la gestion d'erreur "dépasst d'indice"
            'ce qui ne fera plus qu'1 test lors de l'erreur !
            arrByte(I - (intFois * lgMax)) = bytC
        End If 'si màj fichier
    Next
    bigFiles = lgCRC Xor &HFFFFFFFF
    'màj fichier dernier block (tableau)
    If lgNoFile > 0 Then
        ReDim Preserve arrByte(1 To nSize - (intFois * lgMax))  'redimensionne le tableau à la taille du dernier block
                                                                ' elle était peut-être juste si capa mémoire importante
        Put #lgNoFile, , arrByte    'écrit le dernier block d'octet
        Erase arrByte               'efface le tableau
    End If
    'Ferme le fichier
    CloseHandle hOrgFile
    Exit Function
BigSize:
    If Err.Number = 9 Then  'dépassement d'indice
        Put #lgNoFile, , arrByte    'écrit le block d'octet
        ReDim arrByte(1 To lgMax)   'redimensionne en effaçant le tableau
        intFois = intFois + 1       'incrémente 1 cpteur de block pour gérer les indices tableau
        DoEvents
        Resume Next
    Else
'        Debug.Print Err.Description & " " & Err.Number
        RaiseEvent ZipError(zeUnKnownError, "Erreur inconnue sur le fichier " & FilePath, False)
        DoEvents
        Exit Function
    End If

End Function


'Private Function bigFiles_Deflate(ByVal FilePath As String, lgIndex As Long, Optional lgNoFile As Long = 0) As Long
'Calcule (octet par octet) le Crc32 à partir du nom d'1 fichier au lieu d'1 tableau d'octets en entrée dans Zlib
'Eventuellement écrit, par blocks, dans 1 fichier dont le canal est communiqué en paramètre optionnel,
'les données lues durant cette opération, pour Compression. Dans ce cas, on économise 1 lecture,
'et on optimise la vitesse d'écriture par les blocks.
'lgIndex est la valeur de l'index tableau du fichier misà jour.

'Dim hOrgFile As Long
'Dim nSize As Long
'Dim Ret As Long
'Dim I As Long   'long pour capacité importante
'Dim lgCRC As Long
'Dim bytT As Byte
'Dim bytC As Byte
'Dim lgTmp As Long
'Dim arrByte() As Byte   'tableau d'octets si on écrit les données dans 1 autre fichier
'Dim buffer() As Byte    'tableau d'octets compressés
'Dim lgmem As MEMORYSTATUS
'Dim lgMax As Long       'valeur maxi à prendre en mémoire
'Dim intFois As Integer  'nbre de passages
'    bigFiles_Deflate = 0    'par défaut au départ
'    tabFiles(lgIndex).commun.CompressionMethod = 8
'    'Récupère un Handle pour manipuler le fichier
'    hOrgFile = CreateFile(FilePath, &H80000000, FILE_SHARE_READ Or FILE_SHARE_WRITE, _
'                            ByVal 0&, OPEN_EXISTING, 0, 0)
'    'Renvoi d'une erreur -> Problème d'accès au fichier
'    If hOrgFile = INVALID_HANDLE_VALUE Then
'        RaiseEvent ZipError(zeUnKnownError, "Le fichier " & FilePath & " n'est pas ou plus accessible !", False)
'        DoEvents
'        Exit Function
'    End If
'    'Taille du fichier
'    nSize = GetFileSize(hOrgFile, 0)
'    'Attention, GetFileSize peut renvoyer une taille de fichier maximum de : 2,147,483,647 bytes.
'    If nSize = ERROR Then
'        RaiseEvent ZipError(zeUnKnownError, "La Taille du fichier " & FilePath & _
'        " dépasse celle prévue pour la FAT32 !", False)
'        DoEvents
'        Exit Function
'    End If
'   'Initialise le pointeur sur le fichier
'    SetFilePointer hOrgFile, 0, 0, FILE_BEGIN
'
'    'si on doit réécrire les données lues du fichier dans 1 autre fichier
'    'via 1 tableau d'octets pour aller + vite
'    If lgNoFile > 0 Then
'        'recherche le potentiel de mémoire dispo pour redimensionner au mieux 1 tableau(compris taille fichier/capa mémoire)
'        GlobalMemoryStatus lgmem
'        lgMax = lgmem.dwAvailPhys / 4   'mémoire physique dispo/4 pour sécu sur le reste appli et les autres applis
'        lgMax = 5000      'pour test donnée bidon à supprimer 5Ko
'        If nSize > lgMax Then   'dimension des tableaux d'octets, si taille fichier > à mémoire raisonnable à utiliser
'            ReDim arrByte(1 To lgMax)
'        Else    ' si taille fichier <= utilisation raisonnable mémoire
'            ReDim arrByte(1 To nSize)
'            lgMax = nSize
'        End If
'        intFois = 0 'initialise la variable à 0
'    End If
'    '
'    On Error GoTo BigSize
'    ' Une table de correspondance des 255 caractères ASCII pour le calcul du CRC a été initialisée
'    ' lors de l'initialisation de la classe (donc 1 seule fois) dans la variable tableau arrCRC32
'    ' initialisation de  lgCRC = 0xffffffff
'    ' pour chaque octet du fichier, on fait :
'    ' calcule lgCRC = (lgCRC >> 8) ^ Table[(lgCRC & 0xFF) ^ byte]
'    ' on renvoie le CRC32 au dernier octet => lgCRC ^ 0xffffffff
'
'    lgCRC = &HFFFFFFFF
'    For I = 1 To nSize
'        'lit le contenu du fichier octet par octet avec la variable
'        ReadFile hOrgFile, bytC, 1, Ret, ByVal 0&
'        bytT = (lgCRC And &HFF) Xor bytC
'        lgTmp = (lgCRC And &H7FFFFFFF) \ 256
'        If (lgCRC And &H80000000) <> 0 Then
'            lgTmp = lgTmp Or &H800000
'        End If
'        lgCRC = lgTmp Xor arrCRC32(bytT)
'        If I Mod 5000 = 0 Then DoEvents 'fait 1 doevents lorsque i est divisible par 5000
'        If lgNoFile > 0 Then
'            'insertion dans le tableau -> (pour éviter 1 test à chq octet on va gérer quand le tableau
'            'sera plein et donc l'écriture dans le fichier, par la gestion d'erreur "dépasst d'indice"
'            'ce qui ne fera plus qu'1 test lors de l'erreur !
'            arrByte(I - (intFois * lgMax)) = bytC
'        End If 'si màj fichier
'    Next
'    bigFiles_Deflate = lgCRC Xor &HFFFFFFFF
'    'màj fichier dernier block (tableau)
'    If lgNoFile > 0 Then
'        ReDim Preserve arrByte(1 To nSize - (intFois * lgMax))  'redimensionne le tableau à la taille du dernier block
'                                                                ' elle était peut-être juste si capa mémoire importante
'        buffer = CompressDeflate(arrByte, nSize - (intFois * lgMax), , Ret) 'écrit le dernier block d'octet compressé
'        If Ret = 0 Then 'si compression OK
'            Put #lgNoFile, , buffer
'            tabFiles(lgIndex).commun.CompressedSize = tabFiles(lgIndex).commun.CompressedSize + UBound(buffer) + 1
'        Else    'si erreur de compression
'            RaiseEvent ZipError(zeDeflate, "Erreur durant compression. Fichier exclus traitement !" & FilePath, False)
'            DoEvents
'        End If
'        Erase arrByte               'efface le tableau
'        Erase buffer                'efface le tableau
'    End If
'    'Ferme le fichier
'    CloseHandle hOrgFile
'    Exit Function
'BigSize:
'    If Err.Number = 9 Then  'dépassement d'indice
'        buffer = CompressDeflate(arrByte, lgMax, True, Ret) 'écrit les blocks d'octets compressés
'        If Ret = 0 Then 'si compression OK
'            Put #lgNoFile, , buffer
'            tabFiles(lgIndex).commun.CompressedSize = tabFiles(lgIndex).commun.CompressedSize + UBound(buffer) + 1
'            Erase buffer
'            ReDim arrByte(1 To lgMax)   'redimensionne en effaçant le tableau
'            intFois = intFois + 1       'incrémente 1 cpteur de block pour gérer les indices tableau
'            DoEvents
'            Resume Next
'        Else    'si erreur de compression
'            RaiseEvent ZipError(zeDeflate, "Erreur durant compression. Fichier exclus traitement !" & FilePath, False)
'            DoEvents
'            Erase arrByte               'efface le tableau
'            Erase buffer                'efface le tableau
'            CloseHandle hOrgFile
'            Exit Function
'        End If
'    Else
''        Debug.Print Err.Description & " " & Err.Number
'        RaiseEvent ZipError(zeUnKnownError, "Erreur inconnue sur le fichier " & FilePath, False)
'        DoEvents
'        Exit Function
'    End If
'End Function

Private Sub WriteLocalFileHeader(ByVal Fh As Long, index As Long)
'ajouter évènement progression !!!! (pour alimenter 2 barres)

Dim FilePos As Long 'pour récupérer la position
With tabFiles(index)
    .centralFileHeaderZFC.RelativeOffsetOfLocalHeader = Loc(Fh)
    'ajouter test taille fichier si + grand que FAT32 => erreur
    
    Put #Fh, , .localFileHeaderZF
    FilePos = Seek(Fh)      ' on en aura besoin si on est 1 gros fichier
    Put #Fh, , .commun
    Put #Fh, , OEMConvert(.ASCIIFNZF, [ANSI to ASCII])
    Put #Fh, , .extraFieldZF
    If .bigFileZF = False Then    'si on n'est pas 1 gros fichier ayant provoqué manque mémoire
        Put #Fh, , .filedataZF
    Else    'manque mémoire => on ne compresse pas
        Dim Ret As Long
        ' calcul du Crc32 à la main sans Zlib (+long, mais pas de limite de taille, excepté FAT32)
        ' et écrit les données lues à l'occasion du calcul dans le fichier de sortie (zip)
        .commun.CRC32 = bigFiles(OEMConvert(.ASCIIFNZF, [ASCII to ANSI]), index, Fh) 'le passage de fh déclenche l'écriture en fichier zip
' en test bigFiles_Deflate, même chose que bigFiles mais compress (se fait par compressDeflate , Ne marche pas encore
'        .commun.CRC32 = bigFiles_Deflate(OEMConvert(.ASCIIFNZF, [ASCII to ANSI]), index, fh) 'le passage de fh déclenche l'écriture en fichier zip
        ' mais attention à l'offset en réécrivant le crc => commun
        Ret = Seek(Fh)  'repère où on est dans le fichier
        ' réécrit .commun renseigné du crc à sa bonne position
                    'si deflate à remettre à jour  dans big...._deflate
        Put #Fh, FilePos, .commun
        'se repositionne où on avait fini d'écrire les données
        Seek #Fh, Ret
    End If  'orientation fichier standard ou gros
    Erase .filedataZF  'Ouf, ça décharge la mémoire !....
End With
End Sub

Private Sub WriteCentralFileHeader(ByVal Fh As Long, index As Long)
With tabFiles(index)
    Put #Fh, , .centralFileHeaderZFA
    Put #Fh, , .commun
    Put #Fh, , .centralFileHeaderZFC
    Put #Fh, , OEMConvert(.ASCIIFNZF, [ANSI to ASCII])
    Put #Fh, , .extraFieldZF
    Put #Fh, , .commentZF
End With
End Sub

Private Function OEMConvert(ByRef Texte As String, _
                            ByVal ConvertMode As eOEMStringTypeConversion) As String
    
    ' Vient de source de EBartSoft : http://www.vbfrance.com/code.aspx?ID=39592
    
    Dim Src  As String
    Dim dest As String
    Dim r    As Long
    Dim Pos  As Long
    
    OEMConvert = Texte
    
    ' sASCI et sANSI sont renseignés dans Class_Initialize
    If (ConvertMode = [ANSI to ASCII]) Then
        Src = sANSI
        dest = sASCI
    Else
        Src = sASCI
        dest = sANSI
    End If
    
    For r = 1 To Len(Src)
        Do
            Pos = InStr(OEMConvert, Mid$(Src, r, 1))
            If (Pos = 0) Then Exit Do
            Mid$(OEMConvert, Pos, 1) = Mid$(dest, r, 1)
        Loop
    Next

End Function

Private Function FichierExiste(strFilePath As String, Optional RenvoiErreur As Boolean = True) As Boolean
    'Teste que le fichier passé en paramètre existe bien
    'RenvoiErreur, par défaut à Vrai, peut-être shuntée, utile lorsqu'on
    'se sert de cette fonction pour chercher l'inexistance du fichier où un renvoi d'erreur est alors injustifié
    'c'est le cas, par exemple de la méthode CreateZip.

Dim lpFD As WIN32_FIND_DATA     'pointeur vers structure de type WIN32_FIND_DATA
Dim hFdFile As Long             'handle de la fonction API FindFirstFile
'initialisations
FichierExiste = False 'par défaut avant test
'recherche fichier donné en argument
hFdFile = FindFirstFile(strFilePath, lpFD)
If hFdFile = INVALID_HANDLE_VALUE Then  'inexistant
    CloseHandle hFdFile
    If RenvoiErreur Then RaiseEvent FileZipErr(strFilePath, 10, "Ce fichier n'existe pas (n° erreur zip)")
    Exit Function
End If
CloseHandle hFdFile
FichierExiste = True
End Function

Private Function CheckInsert(strFindPath As String, StrFilter As String, lgRecurs As Long, _
                            PathChoice As Long, opRecurs As Boolean, _
                            Optional excludeCRC32 As Boolean = False) As Long
'--------------------------------------------------------
'Cette fonction recherche tous les fichiers dans le chemin spécifié et ses sous-repertoires
'S'il y a 1 filtre de sélection, il ne s'appliquera que pour les fichiers, pas sur les sous-répertoires
'qui font tous l'objet de l'analyse, néanmoins le résultat n'affichera que les s-répert contenant les
'fichiers répondant au filtre.
'Elle retourne le nombre d'occurences trouvées:le nbre de fichiers.
'Elle retourne les fichiers et répertoires tels qu'ils viennent sans tri !

'strFindpath = le chemin des fichiers à chercher correspondant au filtre
'strFilter   = filtre/définition des fichiers à chercher dans le répertoire strFindpath
'lgRecurs    = variable interne à la fonction (cpte le nbre de récursivité), doit être passée à 0
'PathChoice  = choix sur le chemin transmis et à transmettre
'opRecurs    = Vrai/Faux si la fonction doit être récursive. (sous-répertoires de strFindPath)

'variables internes
Dim lpFD As WIN32_FIND_DATA     'pointeur vers structure de type WIN32_FIND_DATA
Dim hFdFile As Long             'handle de la fonction API FindFirstFile
Dim lgSRep As Long              'nbre de sous-répertoires
Dim strTmp As String            'variable de travail des string
Dim subPath()                   'tableau pour récupérer les sous-répertoires et accélérer la recherche en éliminant un passage
Dim Id As Long                  'index
Dim iLen As Integer             'longueur du nom
    
'initialisations
If lgRecurs = 0 Then m_Nombre = 0 'initialisation à 0 au démarrage seulement
If Right(strFindPath, 1) <> "\" Then 'le Path fourni doit se terminer par \ => contrôle
    strFindPath = strFindPath & "\" ' on ajoute le \ manquant
End If  'controle syntaxe du path
CheckInsert = 0   'par défaut retour à 0
' Mises à blanc lpFD (Variable texte uniquement)
lpFD.cAlternate = String$(14, 0)
lpFD.cFileName = String$(MAX_PATH, 0)
'récup premier fichier du chemin donné en argument
hFdFile = FindFirstFile(strFindPath & StrFilter, lpFD)

'Recherche
If hFdFile <> INVALID_HANDLE_VALUE Then
    Do
       'Récupère le nom de l'élèment du répertoire
        iLen = InStr(lpFD.cFileName, Chr$(0)) - 1
        strTmp = Mid$(lpFD.cFileName, 1, iLen)
        'Orientation du traitement selon nature de l'élèment du répertoire
        If strTmp <> ".." And strTmp <> "." And (lpFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = 0 Then
            'on est un fichier, c'est le moment de l'ajouter au tableau des fichiers à zipper
            m_Nombre = m_Nombre + FileAddSingle(strFindPath & strTmp, PathChoice, excludeCRC32)
        ElseIf strTmp <> ".." And strTmp <> "." And StrFilter = "*.*" Then
            'on est 1 sous-répertoires mais ATTENTION qui répond au filtre => on ne le
            'récupère que si le filtre est *.* car si le filtre est + sélectif on aurait que
            'les sous-répertoires correspondants, donc dans ce cas on fera 1 autre scan pour avoir
            'tous les sous-répertoires et ces éléments seront récupérés à ce moment là !
            lgSRep = lgSRep + 1
            ReDim Preserve subPath(1 To lgSRep)
                subPath(lgSRep) = strFindPath & strTmp & "\" 'chemin et nom du sous-répertoire
        Else
            'on est ".." ou "." et on ne fait rien  (FindFirstFile et FindNextFile renvoient ces éléments lors du scan)
        End If 'nature de l'élément récupéré
       ' Remises à blanc lpFD (Variable texte uniquement)
       lpFD.cAlternate = String$(14, 0)
       lpFD.cFileName = String$(MAX_PATH, 0)
       iLen = 0
       DoEvents 'pour donner la main au système
    Loop Until FindNextFile(hFdFile, lpFD) = 0 'on récupère le prochain fichier
Else            ' Recherche 1° fichier sans éléments
    ' si le filtre est <> *.* on re-scan pour analyser tous les sous-répertoires du chemin
    ' et non que ceux répondant au filtre
    If StrFilter <> "*.*" Then
        ' va extraire les sous-répertoires plus loin
    Else    'sinon
        'l'association chemin+filtre n'existe pas
        'renvoyer status inaccesible
        RaiseEvent Status("Association Chemin, filtre inaccessible")
        CheckInsert = m_Nombre
        Exit Function
    End If  'analyse du filtre
End If  'handle Recherche
'on stoppe le scan du répertoire passé en argument
FindClose hFdFile
' si le filtre est <> *.* on re-scan pour extraire uniquement tous les sous-répertoires du chemin
' et non que ceux répondant au filtre
If StrFilter <> "*.*" Then
    strTmp = ""
    hFdFile = FindFirstFile(strFindPath & "*.*", lpFD)
    Do
       'Récupère le nom de l'élèment du répertoire
        iLen = InStr(lpFD.cFileName, Chr$(0)) - 1
        strTmp = Mid$(lpFD.cFileName, 1, iLen)
        'gère tous les sous-répertoires du répertoire indépendament du filtre d'origine
        If strTmp <> ".." And strTmp <> "." And (lpFD.dwFileAttributes = FILE_ATTRIBUTE_DIRECTORY) Then
            'on est 1 sous-répertoire
            lgSRep = lgSRep + 1
            ReDim Preserve subPath(1 To lgSRep)
                subPath(lgSRep) = strFindPath & strTmp & "\" 'chemin et nom du sous-répertoire
        End If
       ' Remises à blanc lpFD (Variable texte uniquement)
       lpFD.cAlternate = String$(14, 0)
       lpFD.cFileName = String$(MAX_PATH, 0)
       iLen = 0
       DoEvents 'pour donner la main au système
    Loop Until FindNextFile(hFdFile, lpFD) = 0 'on récupère le prochain élément
    'on stoppe le scan du répertoire passé en argument
    FindClose hFdFile
End If  'passage suppl si filtre <> *.* pour les sous-répertoires
'
'Scan récursif des sous-répertoires, s'il y en a ! et si on l'a demandé
If lgSRep > 0 And opRecurs = True Then
    For Id = 1 To UBound(subPath)
           CheckInsert CStr(subPath(Id)), StrFilter, lgSRep, PathChoice, opRecurs, excludeCRC32 ' ici lgSRep permet de gérer le niveau de récursivité
    Next    'boucle récursive sous-répertoires
End If  'S'il y a des sous-répertoires
' c'est fini si on est en 1 appel de la fonction
If lgRecurs = 0 Then     ' seulement lors du 1° appel de la fonction
    '---on retourne le nombre de fichiers trouvées et mise à jour évènement status
    CheckInsert = m_Nombre
End If
    
End Function

Private Function FolderExists(strFindPath As String) As Boolean
'Teste si le dossier passé en argument existe
'Renvoie Vrai /Faux

'variables internes
Dim lpFD As WIN32_FIND_DATA     'pointeur vers structure de type WIN32_FIND_DATA
Dim hFdFile As Long             'handle de la fonction API FindFirstFile
'Dim iLen As Integer             'longueur du nom
    
'initialisations
FolderExists = False
If Right(strFindPath, 1) <> "\" Then 'le Path fourni doit se terminer par \ => contrôle
    strFindPath = strFindPath & "\" ' on ajoute le \ manquant
End If  'controle syntaxe du path
' Mises à blanc lpFD (Variable texte uniquement)
lpFD.cAlternate = String$(14, 0)
lpFD.cFileName = String$(MAX_PATH, 0)

'Scan pour test
hFdFile = FindFirstFile(strFindPath & "*.*", lpFD)
If hFdFile <> INVALID_HANDLE_VALUE Then
    ' on a au moins récupéré 1 élément dans le répertoire, donc il existe
    ' en effet, même si le répertoire est vide, il y aura toujours . ou ..
    FolderExists = True
Else    ' erreur donc n'existe pas car même si vide il y aurait toujours . ou ..
    ' on fait rien puisque la fonction est initialisée à false, on finit et on sort
End If
'on stoppe le scan du répertoire passé en argument
FindClose hFdFile

End Function

'----------------------------------------------------------------------------------------------------
''''Private Sub ExtraitDLL()
''''
''''     ' J'ai mis la DLL dans une ressource, comme ça, il elle n'existe pas, on l'extrait
''''     ' On n'est pas obligé de la mettre dans le répertoire de Windows\System32
''''
''''     Dim Temp() As Byte, ff As Integer
''''     ' -------------- Fichier ZLib.DLL
''''     Temp = LoadResData(101, "LIBRAIRIE_ZLIB")
''''     ff = FreeFile
''''     Open App.Path & "\ZLib.DLL" For Binary Access Write As #ff
''''          Put #ff, , Temp
''''     Close #ff
''''
''''End Sub
