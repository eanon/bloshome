Attribute VB_Name = "basGlobal"
'BlosHome (c) FFh Lab / Eric Lequien, 2009-2013 - http://ffh-lab.com
'This module contains the global functions (i.e. eventually portable toward other projects)

Option Explicit
Option Base 1

Public Const EDITOR_DIR = "tiny_mce"
Public Const EDITOR_CONTENT_CSS = "~content.css"    'css copy of the current project

Public Const EDITOR_TEMPLATE = "editor.html"        'editor template
Public Const EDITOR_CURRENT = "~editor.html"        'customized editor in use (built from editor template)

Public Const EDITOR_TITLE_CSS_HOLDER = "CSS-TITLE_HOLDER"
Public Const EDITOR_TITLE_HOLDER = "<!-- TITLE_HOLDER -->"

Public Const EDITOR_CSS_HOLDER = "CSS_HOLDER"
Public Const EDITOR_CHARSET_HOLDER = "iso-8859-1"
Public Const EDITOR_LANG_HOLDER = "language : 'en',"
Public Const EDITOR_CONTENT_HOLDER = "<!-- CONTENT_HOLDER -->"

Public Const PAGEBREAK_HOLDER = ",PAGEBREAK_HOLDER"
Public Const PAGEBREAK_TAG_HOLDER = "PAGEBREAK_TAG_HOLDER"
Public Const PAGEBREAK_BTN_HOLDER = ",PAGEBREAK_BTN_HOLDER"

Public Const EDITOR_GENERATOR_SIGHT = "<!-- Written in BlosHome (c) ffh-lab.com, using TinyMCE (c) moxiecode.com -->"
Public Const EDITOR_ENCODING_WITNESS = "<!-- é -->"    'to differenciate ANSI/UTF-8 whatever be the file content
Public Const EDITOR_ENCODING_WITNESS_AS_ENTITE = "<!-- &eacute; -->" 'TinyMCE converts *_ENCODING_WITNESS to HTML entity

Public bRightLoading As Boolean 'ask to terminate appli ASAP (i.e. on first form activation) if false
Public bOpenDelProj As Boolean  'determine the operation to do with the frmSelect dialog

Public strCurrArt As String     'indicates the article currently loaded in the editor

Public Type TRANSFILE           'describes useful elements to tranfer (upload or download) a file
    local As String
    remote As String
End Type

Public nTimerPass As Long       'counts the number of pass through timer event (resetted on StartStopTimer)

Public Sub PopulateRemoteList()
    'Populates frmMain.lstRemote with remote articles and categories (i.e. directories) at current tree level
    'NB : global since it acts on frmMain but callable from frmProject
    Dim item As ListItem
    Dim fileRemote As RemoteFile
    
    Dim strImg As String
    Dim strDisplayName As String
    Dim nColor As Long
    
    Dim strPathFromRoot As String
    Dim arExcluded() As String
    Dim bExcluded As Boolean
    Dim nIdx As Integer
    
    frmMain.lstRemote.ListItems.Clear
    
    If frmMain.FtpCli.Met_DIR Then
        'user info
        frmMain.lblRemotePath = frmMain.FtpCli.RemoteDir
        
        'add an access to the upper level (unless we're already at root one)
        If Len(frmMain.FtpCli.RemoteDir) > Len(structProj.art_root) Then
            strImg = "Up"
            Set item = frmMain.lstRemote.ListItems.Add(, "..", "..", strImg, strImg)
            
            For nIdx = 1 To 6
                item.SubItems(nIdx) = " "  'évite les colonnes non sélectionnable
            Next                           'et classe l'entrée devant les catégories du fait ' '>'A'>'Z'
        End If

        'enumerates the remote directories and files
        For Each fileRemote In frmMain.FtpCli.RemoteFiles
            'détermine the appropriated icon according to the type returned by DiFtpCli and blosxom status
            '(DiFtpCli can return these types : FtpDirectory=0 | FtpFile=1 | FtpLink=2 | FtpUnkwnon=4)
            Select Case fileRemote.FileType
                Case 0:
                    strImg = "Category"
                Case 1:
                    If Left$(fileRemote.FileName, Len(structProj.preview_prefix) + 1) _
                                                    = structProj.preview_prefix & "-" Then
                        strImg = "Preview"
                    Else
                        strImg = "Article"
                    End If
                Case Else:
                    strImg = "[ignored]" 'will be not displayed
            End Select
            
            'excludes the directories "." et ".." returned by server
            '(".." being already forced above, since some servers don't return-it)
            bExcluded = False
            If fileRemote.FileName = "." Or fileRemote.FileName = ".." Then
                bExcluded = True
            Else
                'decides if this path must be ignored according to the project settings
                arExcluded() = Split(structProj.excluded_paths, ",")
                For nIdx = LBound(arExcluded) To UBound(arExcluded)
                    strPathFromRoot = frmMain.FtpCli.RemoteDir & "/" & fileRemote.FileName
                    If InStr(strPathFromRoot, arExcluded(nIdx)) <> 0 Then bExcluded = True
                Next
            End If
            
            'ajusts the appearance of the item (name to display and color)
            strDisplayName = fileRemote.FileName
            nColor = RGB(0, 0, 0)
            If strImg = "Preview" Then
                'private previews will appear without prefix and greyed
                strDisplayName = Right$(strDisplayName, Len(strDisplayName) - Len(structProj.preview_prefix) - 1)
                nColor = RGB(129, 129, 129)
            End If
                        
            'effective addition
            If bExcluded = False _
                And (((strImg = "Article" Or strImg = "Preview") _
                        And Right$(fileRemote.FileName, 4) = structProj.art_ext) _
                            Or strImg = "Category") Then
                On Error Resume Next
                Set item = frmMain.lstRemote.ListItems.Add(, fileRemote.Key, strDisplayName, strImg, strImg)
                
                item.SubItems(1) = Format(fileRemote.FileDate, "dd/mm/yyyy")
                item.SubItems(2) = Format(fileRemote.FileDate, "hh:nn")
                item.SubItems(3) = Int((fileRemote.FileSize / 1024) + 1)
                
                'differenciated treatment for directories and files
                '(including hidden key column to sort things with directories above files)
                If fileRemote.FileType = 0 Then
                    item.SubItems(4) = " "
                    item.SubItems(5) = " "
                    item.SubItems(6) = "A" 'folder will be sorted at top
                Else
                    item.SubItems(4) = MetaCache(frmMain.FtpCli.RemoteDir & "/" & fileRemote.FileName, "Joint", "GET")
                    item.SubItems(5) = LCase(MetaCache(frmMain.FtpCli.RemoteDir & "/" & fileRemote.FileName, "Encode", "GET"))
                    item.SubItems(6) = "Z" 'and the rest (i.e. files) below
                End If
                
                item.ForeColor = nColor
                For nIdx = 1 To item.ListSubItems.Count
                    item.ListSubItems(nIdx).ForeColor = nColor
                Next
            End If
        Next
    End If
    
    'ensures a double sorting : categories at top, then files by alphabetic order below
    frmMain.lstRemote.SortKey = 0  'column "Name" (zero based)
    frmMain.lstRemote.SortOrder = lvwAscending
    frmMain.lstRemote.Sorted = True
    
    frmMain.lstRemote.SortKey = 6  'column "Sort" (zero based)
    frmMain.lstRemote.SortOrder = lvwAscending
    frmMain.lstRemote.Sorted = True
    
    'ensures not any item remains selected
    If frmMain.lstRemote.ListItems.Count > 0 Then frmMain.lstRemote.SelectedItem.Selected = False
End Sub

Function MetaCache(strArt As String, strInfo As String, strOp As String, _
                    Optional strQuickRef As String = "") As String
    'Maintains local cache for "slow infos" (too slow ones to acquire remotely every time) about remote articles
    'IN :
    '- remote path of the article file ; if file absent -> info deleted (not the section)
    '- name of the info to obtain :
    '   . "Joint" : counts the number of attachments to the article
    '   . "Encode" : indicates the article encoding (ANSI or UTF-8)
    '- requested operation :
    '   . "GET" : obtains the info from cache or from real remote file if nothing in cache
    '   . "UPDATE" : forces the update (or first writing) of the info in cache
    '   . "REMOVE" : deletes all infos about the given article in cache (ignores strInfo)
    '   . "RENAME" : move the infos toward a new article name (strInfo indicates the new name)
    '- path to a referential local copy, in order to avoid any download on UPDATE or GET->UPDATE
    '  (of course, this reference takes sense from the point several MetaCache() calls are planned)
    'OUT : value of the obtained info (obtained directly from cache or after access to the remote article)
    '      or "" in case of failure during process
    'NB : there is one "meta.cache" INI-file by project,
    '     this file is maintained every time the blog tree is modified ; ***TODO : check all cases are covered
    'WARNING : caller must take care to connect prior to call (to avoid error if remote access is necessary)
    'SECU : this fct takes care to process articles only (neither directories nor filename w/o valid art_ext)
    '***LATER : provide a call type to acquire all remote file infos from a single call (i.e one download only)
    Dim strCache As String
    Dim strValue As String
    Dim bRet As Boolean
    
    Dim strTmpFile As String
    Dim strContent As String
    
    MetaCache = ""
    
    If Right$(strArt, Len(structProj.art_ext)) <> structProj.art_ext Then Exit Function
    strCache = GetWorkspace() & "\meta.cache"
    
MetaCache_Operation:
    Select Case strOp
        Case "GET"
            'Try to obtain the info from local cache file
            strValue = LoadIni(strCache, strArt, strInfo)
            If strValue = "" Then
                'nothing in cache => need to acquire the info from remote article file itself
                strOp = "UPDATE"
                GoTo MetaCache_Operation
            End If
        Case "UPDATE"
            'Forces the info in cache from the reality of remote article (creates the info in cache if absent)
            If strQuickRef = "" Then
                strTmpFile = GetTmpFile("BH")
                bRet = frmMain.FtpCli.Met_GET(strArt, strTmpFile, "A") 'with MODded Met_GET to support "A"
                If bRet = False Then Exit Function
            Else
                If Dir(strQuickRef) = "" Then Exit Function
                strTmpFile = strQuickRef
            End If
            
            Select Case strInfo
                Case "Joint"
                    strValue = CountRemoteArtAnnexes(strTmpFile)
                Case "Encode"
                    If IsUTF8(LoadText(strTmpFile)) = True Then
                        strValue = "UTF-8"
                    Else
                        strValue = "ANSI" 'default
                    End If
                Case Else
                    MsgBox "Bad 'info' in MetaCache() call", vbCritical, "DEBUG"
                    Exit Function
            End Select
            
            'saves in cache to accelerate next accesses
            SaveIni strCache, strArt, strInfo, strValue
            
            'final clean-up
            If strTmpFile <> "" And Dir(strTmpFile) <> "" Then Kill strTmpFile
        Case "REMOVE"
            'Deletes the section about the given article
            RemoveIniSection strCache, strArt
        Case "RENAME"
            'Rename the section about the given article
            '***TODO : check effective working !
            strContent = LoadText(strCache)
            strContent = Replace(strContent, strArt, strInfo, , 1, vbTextCompare)
            SaveText strCache, strContent
        Case Else
            MsgBox "Bad 'op.' in MetaCache() call", vbCritical, "DEBUG"
            Exit Function
    End Select
    
    ArrangeINISections strCache
    MetaCache = strValue
End Function

Public Sub PopulateClearLocalList(bOp)
    'Populates frmMain.lstLocal with the local articles in stand by, or empties-it
    'IN : flag indicating the sense of the operation
    'NB : global since acts on frmMain but callable from frmProject
    Dim strWorkspace As String
    Dim strFilename As String
    Dim strFullPath As String
    
    Dim item As ListItem
    Dim bFirstLoop As Boolean
    Dim strImg As String
    
    frmMain.lstLocal.ListItems.Clear
    If bOp = False Then Exit Sub
    
    strWorkspace = GetWorkspace() & "\"
    
    strImg = "Article"  'we display articles only
    
    If Dir(strWorkspace & "*.*") <> "" Then
        'enumeration
        bFirstLoop = True
        Do
            'file by file
            If bFirstLoop = True Then
                strFilename = Dir(strWorkspace & "*.*", vbNormal)
                bFirstLoop = False
            Else
                strFilename = Dir
            End If
            
            'if article type (extension based)
            If Right$(strFilename, 4) = structProj.art_ext Then
                'we diplay-it in workspace
                strFullPath = strWorkspace & strFilename
                
                On Error Resume Next
                Set item = frmMain.lstLocal.ListItems.Add(, strFilename, strFilename, strImg, strImg)
                item.SubItems(1) = Format(FileDateTime(strFullPath), "dd/mm/yyyy")
                item.SubItems(2) = Format(FileDateTime(strFullPath), "hh:nn")
                item.SubItems(3) = Int((FileLen(strFullPath) / 1024) + 1)
                item.SubItems(4) = CountRemoteArtAnnexes(strFullPath)
                
                If IsUTF8(LoadText(strFullPath)) = True Then
                    item.SubItems(5) = "utf-8"
                Else
                    item.SubItems(5) = "ansi" 'default
                End If
            End If
        Loop While strFilename <> ""
    End If
    
    If frmMain.lstLocal.ListItems.Count > 0 Then
        frmMain.lstLocal.SelectedItem.Selected = False 'ensures not any item remains selected
    End If
End Sub

Public Function CountRemoteArtAnnexes(strArt As String) As Integer
    'Returns the number of files associated with a given article
    '***LATER :
    '- actually, this function search for the "src=["|']" used to indicate the URL in <IMG> tags,
    '  but later see if this way allow to detect all type of attachment (i.e. does all associated
    '  files are referenced with a "src=" attribute in an HTML tag ; I mean all kind of file types
    '  using all kind of tag types - video, PDF, object, iframe, etc)
    '- maybe use this function in the framework of cmdUploadArt()_Click
    Dim strData As String
    Dim arMark(1 To 2) As String
    Dim nIdx As Integer
    
    arMark(1) = "src='"
    arMark(2) = "src=" & Chr(34)
    
    CountRemoteArtAnnexes = 0
    
    strData = LoadText(strArt)
    If strData = "" Then Exit Function
        
    For nIdx = 1 To UBound(arMark)
        CountRemoteArtAnnexes = CountRemoteArtAnnexes + CountOccurr(strData, arMark(nIdx))
    Next
End Function

Public Function EnumLocalArtAnnexes(strArt As String, Optional bWithPath As Boolean = True) As String()
    'Returns the local path of images attached to the given article
    'IN :
    '- path to the local article to analyze
    '- flag indicating if the path should be included in the returned string or not (yes by default)
    'OUT :
    '- array containing local paths to every attached image (found in <img src=>)
    '- array of one empty element ("") if there is no attached image
    '- array of one element with "FAILED" value if the process encountered a problem
    'WARNING : returned array is one-based
    '***LATER :
    '- extend the notion of attached image to the one of attached file (video, audio, flash, pdf, document, etc)
    '- maybe use this function in the framework of cmdUploadArt()_Click
    Dim nImgPos As Long
    Dim strContent As String
    Dim arAttach() As String 'local paths of the images
    
    Const FAILED = "failed"
    
    'init
    ReDim arAttach(1 To 1)
    arAttach(1) = ""
    
    'checking
    If Dir(strArt) = "" Then
        arAttach(1) = FAILED
        EnumLocalArtAnnexes = arAttach()
        Exit Function
    End If
    strContent = LoadText(strArt)

    'collects the image paths (or name only) from 'src' attribute of <img> tags (case insensitive)
    '***LATER : add case where src attribute use single-quote (') rather than double ones (")
    Do
        nImgPos = InStr(nImgPos + 1, LCase(strContent), "<img", vbTextCompare)
        If nImgPos <> 0 Then
            Dim nSrcStart As Long
            Dim nSrcEnd As Long
            Dim strSRC As String
            
            nSrcStart = InStr(nImgPos, strContent, "src=" & Chr(34), vbTextCompare) + 5
            nSrcEnd = InStr(nSrcStart, strContent, Chr(34), vbTextCompare) - 1
            strSRC = Mid$(strContent, nSrcStart, nSrcEnd - nSrcStart + 1)
            
            If Left$(strSRC, 8) = "file:///" Then strSRC = Right$(strSRC, Len(strSRC) - 8)
            
            If bWithPath = False Then
                strSRC = GetFileName(strSRC)
            End If
            
            strSRC = CheckAndMkFilename(strSRC, False, False, True, True)
            
            If arAttach(1) <> "" Then ReDim Preserve arAttach(1 To UBound(arAttach()) + 1)
            arAttach(UBound(arAttach())) = strSRC
        End If
    Loop While nImgPos <> 0
    
    EnumLocalArtAnnexes = arAttach()
End Function

Public Function PrepArtForPublishing(strArt As String, ByRef nFinalArtSize As Long, _
                                    ByRef nAddedNameLen As Integer, ByRef strFinalAttachList As String, _
                                    Optional bWarn As Boolean = False) As TRANSFILE()
    'Prepares an article for publication
    'IN :
    '- path to the local article (or temporary copy since the file will be modified)
    '- buffer variable which will receive the number of characters of the final article
    '  (mayb be used to check if uploaded article is complete - not corrupted)
    '- buffer variable which will receive the number of characters added by the images renaming
    '  (renaming happens when duplicate exist on server in same blog category ; will be noted in archive info)
    '- buffer variable which will receive the list of attachment - their names - on server)
    '  (this to keep local trace of renaming in archive, in "info" file ; see SetGetArchive)
    '- flag indicating if the function should display a message in case of problem
    'OUT : size of the article on server (outside of any CRLF conversion) in nFinalArtSize
    '      and a TRANSFILE structures array containing :
    '      - local path of any image to transfer
    '      - final names of these images on server (the ones indicated in "src" attributes of <img> tags)
    '      or
    '      - "" as first element of the array if the article contains no image
    '      - "failed" as first element if failure
    'WARNING : returned array is one-based
    'NB : nFinalArtSize takes into account the path modifications (i.e. final <img> URLs)
    '     as the modification about image names (when duplicate still exist in category).
    'REQ :
    '- MsgBoxEx() for clever management of the hourglass cursor
    '- declaration of the FROM_TO structure type as :
    '   Public Type TRANSFILE
    '       local As String
    '       remote As String
    '   End Type
    '***LATER :
    '- extend the notion of associated image to the one of associated file (of any type)
    '- maybe see to use GetArtAttachments() at beiginning of treatment ; to condense code.
    Dim bStop As Boolean
    Dim nRet As Integer
    
    Dim nImgPos As Long
    Dim strContent As String
    Dim strMsg As String
    
    Dim arAttach() As TRANSFILE 'vecteur [chemins locaux des images, noms de destination]
    
    Dim nOrgNameLen As Integer 'pour calcul argIN nAddedNameLen
    Dim nDeltaNameLen As Integer
    
    Const FAILED = "failed"
    Const BLOG_IMG_PATH = "<$url /><$path />" '***LATER : permettre choix via params projet
    
    If Dir(strArt) = "" Then Exit Function
    strContent = LoadText(strArt)
        
    ReDim arAttach(1 To 1)
    arAttach(1).local = ""
    
    'basic checking
    bStop = False
    
    If bWarn = True And strContent = "" Then
        strMsg = arMsg(101) & " ! " & arMsg(102) & " ?"
        nRet = MsgBoxEx(strMsg, vbQuestion + vbYesNo)
        If nRet = vbNo Then bStop = True
    End If
    
    If bWarn And InStr(1, strContent, "<p>", vbTextCompare) = 0 Then
        strMsg = arMsg(103) & " ! " & arMsg(102) & " ?"
        nRet = MsgBoxEx(strMsg, vbQuestion + vbYesNo)
        If nRet = vbNo Then bStop = True
    End If
    
    If bStop = True Then
        arAttach(1).local = FAILED
        PrepArtForPublishing = arAttach
    End If

    'saves local image paths, then adjust (imply modification or not) them for the blog context
    Do
        nImgPos = InStr(nImgPos + 1, LCase(strContent), "<img", vbTextCompare)
        If nImgPos <> 0 Then
            Dim nRemoteIdx As Integer
            Dim fileRemote As RemoteFile
            Dim nFileIndice As Integer
            
            Dim nPos As Long
            Dim strCurr As String
            Dim strNew As String
            
            Dim nSrcStart As Long
            Dim nSrcEnd As Long
            Dim strSRC As String
            Dim nOrgLenSRC As Long
            Dim strSrcName As String
            
            Dim strInsert As String
            Dim nDecay As Integer
        
            'searches "src" attribute of <img> tag (case insensitive)
            '***LATER : considers case where path is surounded by single-quotes (') rather than double ones (")
            nSrcStart = InStr(nImgPos, strContent, "src=" & Chr(34), vbTextCompare) + 5
            nSrcEnd = InStr(nSrcStart, strContent, Chr(34), vbTextCompare) - 1
            
            strSRC = Mid$(strContent, nSrcStart, nSrcEnd - nSrcStart + 1)
            nOrgLenSRC = Len(strSRC)
            nPos = RevInStr(strSRC, "/", False)
            strSrcName = Right$(strSRC, Len(strSRC) - nPos)
            
            'normalizes the final filename to be compatible with all kind of servers (*nix/win)
            '(i.e. lower cases, without accent, without space nor "%20")
            strSrcName = UnAccent(LCase(CheckAndMkFilename(strSrcName, True, True, True)))
            
            'ensures not any duplicate image name still exists in the (remote) category ; rename if necessary
            nOrgNameLen = Len(strSrcName)
PrepArtForPublishing_CheckName:
            For Each fileRemote In frmMain.FtpCli.RemoteFiles
                nRemoteIdx = nRemoteIdx + 1
                If fileRemote.FileName = strSrcName Then
                    'need to rename the image on server and its reference in the final article
                    Dim strOrgSrcName As String
                    
                    If nFileIndice = 0 Then
                        'stores original name (to be able to be back on it before any index attempt)
                        strOrgSrcName = strSrcName
                    Else
                        'reset to original name before to insert an index
                        '(starting from a stable state avoids any "_x" accumulation before the extension)
                        strSrcName = strOrgSrcName
                    End If
                    
                    nPos = RevInStr(strSrcName, ".", False)
                    nFileIndice = nFileIndice + 1
                    strSrcName = Left$(strSrcName, nPos - 1) & "_" & nFileIndice & _
                                    Right$(strSrcName, Len(strSrcName) - nPos + 1)
                    GoTo PrepArtForPublishing_CheckName
                End If
            Next
            
            nDeltaNameLen = Len(strSrcName) - nOrgNameLen
            nAddedNameLen = nAddedNameLen + nDeltaNameLen
            
            nFileIndice = 0
            
            'searches if the name is not used in the framework of a pop-up call
            'which would be generated by the "Easy Image" (ezimage) plugin for TinyMCE
            strCurr = "window.open('" & strSRC & "'"
            If InStr(1, strContent, strCurr, vbTextCompare) <> 0 Then
                strInsert = BLOG_IMG_PATH & "/" & strSrcName
                strNew = Replace(strCurr, strSRC, strInsert, , , vbTextCompare)
                strContent = Replace(strContent, strCurr, strNew, , 1, vbTextCompare)
                
                nDecay = Len(strInsert) - nOrgLenSRC  '(màj des positions référentielles utiles):
                nSrcStart = nSrcStart + nDecay
                nSrcEnd = nSrcEnd + nDecay
            
                nAddedNameLen = nAddedNameLen + nDeltaNameLen '(prise en compte influence 2e renommage)
            End If
            
            'inserts an HTML comments which will store local path of the original local image
            strInsert = "<!-- " & strSRC & " -->"
            strContent = Left$(strContent, nImgPos - 1) & strInsert & _
                            Right$(strContent, Len(strContent) - nImgPos + 1)
                                    
            nDecay = Len(strInsert)         '(updates referential useful positions)
            nImgPos = nImgPos + nDecay
            nSrcStart = nSrcStart + nDecay
            nSrcEnd = nSrcEnd + nDecay
           
            'modifies source of <img> tag for the blog context
            '(no subsequent nImgPos update, because the modfication is done after "<IMG")
            strInsert = BLOG_IMG_PATH & "/" & strSrcName
            strContent = Left$(strContent, nSrcStart - 1) & strInsert & _
                            Right$(strContent, Len(strContent) - nSrcEnd)
                            
            'notes the characteristics of the image to transfer
            '(including modification of strSRC to allow access to the copy)
            If Left$(strSRC, 8) = "file:///" Then strSRC = Right$(strSRC, Len(strSRC) - 8)
            strSRC = CheckAndMkFilename(strSRC, True, False, True, True)
            
            If arAttach(1).local <> "" Then ReDim Preserve arAttach(1 To UBound(arAttach) + 1)
            arAttach(UBound(arAttach)).local = strSRC      'full local path
            arAttach(UBound(arAttach)).remote = strSrcName 'future filename on server
            
            If Len(strFinalAttachList) > 0 Then strFinalAttachList = strFinalAttachList & ","
            strFinalAttachList = strFinalAttachList & strSrcName
        End If
    Loop While nImgPos <> 0
    
    'updates the article with modified content
    nFinalArtSize = Len(strContent)
    SaveText strArt, strContent
    
    'returns the array of local image paths (those will be uploaded and, eventually, renammed during transfer)
    PrepArtForPublishing = arAttach
End Function

Function GetCurrRemoteCat(Optional strSubstSepar As String = "", _
                            Optional bRootAsString As Boolean = False) As String
    'Determines current remote category (path from blog root)
    'IN :
    '- string we want to use to replace each "/" separator in final returned path (remains "/" if absent)
    '- flag indicating if caller wants the function changes root as "<ROOT>" rather than an empty string ("")
    'OUT : useful part of path toward current category (directory)
    '      or "" if nothing found (i.e. not connected)
    'NB : of course, this function can return path to a subcategory (including parent ones in path)
    'USE : GetCurrRemoteCat("::") to return something like "reflexion::health::sport"
    Dim strCat As String
    
    strCat = frmMain.FtpCli.RemoteDir
    If strCat = "" Then
        GetCurrRemoteCat = ""
        Exit Function
    End If
    
    If Left$(UCase(strCat), Len(structProj.art_root)) = UCase(structProj.art_root) Then
        strCat = Right$(strCat, Len(strCat) - Len(structProj.art_root))
    End If
    
    If Left$(strCat, 1) = "/" Then
        strCat = Right$(strCat, Len(strCat) - 1)
    End If
    
    If strCat = "" And bRootAsString = True Then
        strCat = "<ROOT>"
    Else
        If strSubstSepar <> "" Then
            strCat = Replace(strCat, "/", strSubstSepar, , , vbTextCompare)
        End If
    End If
    
    GetCurrRemoteCat = strCat
End Function

Sub CnnxDcnnx(bOp)
    'Connects to or disconnects from remote blog
    'IN : flag indicating the sense of the operation
    'NB : includes management of hourglass cursor and user interface
    Dim strMsg As String
    Dim nRet As Integer
    
    frmMain.cmdCnnx.Enabled = False
    Screen.MousePointer = vbHourglass
    
    If bOp = True Then
        'connects (FtpCli takes care to display eventual error message)
        frmMain.barStatus.Panels.item("cnnx").Text = arMsg(67) & "... " & _
                                                "(" & frmMain.FtpCli.TimeOutDelay & ")"
        If structProj.host <> "" And structProj.port <> "" _
            And structProj.user <> "" And structProj.pass <> "" Then
            '(connection parameters)
            frmMain.FtpCli.RemoteHost = structProj.host
            frmMain.FtpCli.RemotePort = structProj.port
            frmMain.FtpCli.UserName = structProj.user
            frmMain.FtpCli.Password = structProj.pass
            
            StartStopTimer frmMain.timerCnnx, True, 1000, nTimerPass 'every seconde
            
            If frmMain.FtpCli.Met_CONNECT Then
                ApplyCnnx True
            Else
                ApplyCnnx False
                MsgBoxEx arMsg(108) & " '" & frmMain.FtpCli.RemoteHost & "'" & _
                        " " & arMsg(109) & " " & frmMain.FtpCli.RemotePort & vbNewLine & _
                        " " & arMsg(110) & " '" & frmMain.FtpCli.UserName & "'" & _
                        " " & arMsg(111) & " '" & frmMain.FtpCli.Password & "'", vbExclamation
            End If
        Else
            '(missing parameters)
            Screen.MousePointer = vbDefault
            strMsg = arMsg(104) & " !" & _
                      vbNewLine & vbNewLine & arMsg(112) & " ?"
                            
            nRet = MsgBox(strMsg, vbYesNo)
            
            If nRet = vbYes Then
                frmProject.bNoCancel = True
                frmProject.Show 1
            End If
        End If
    Else
        'disconnects (FtpCli takes care to display eventual error message)
        If frmMain.FtpCli.Met_BYE Then ApplyCnnx False
    End If
    
    Screen.MousePointer = vbDefault
    frmMain.cmdCnnx.Enabled = True
End Sub

Function IsFreeSDI(strAction As String) As Boolean
    'Checks no project is loaded in the Single Document Interface (asks user what to do if there is one)
    'IN : string completing "You are going to", indicating the reason why caller asks for release of the SDI
    'OUT : true if the way is clear (no project loaded or current one unloaded)
    '      false if there's a current project in use that user don't want to close
    'USE : prior to File/[New|Open|Del] operations
    'NB : note that structProj.filename is the first informed when loading a project
    If structProj.FileName <> "" Then
        Dim strMsg As String
        Dim nRet As Integer
        
        strMsg = arMsg(105) & " " & strAction & vbNewLine & arMsg(106) & "." & _
                 vbNewLine & vbNewLine & arMsg(107) & " ?"
        
        nRet = MsgBox(strMsg, vbQuestion + vbYesNo)
        If nRet = vbNo Then
            IsFreeSDI = False
            Exit Function
        Else
            LoadUnloadProject False
        End If
    End If
    IsFreeSDI = True
End Function

Function GetArtStatus(strFile As String) As Boolean
    'Determines the current status (public/private) of given article
    'IN : filename of the article in current remote directory or in local workspace
    'OUT : true if public, false if private
    'NB : all articles will be considered public if structProj.with_preview is false (no preview plugin used)
    '***LATER : we could add a frmProject/'Check' cmd, checking for and reading preview plugin value on server
    GetArtStatus = True
    If structProj.with_preview = False Then Exit Function
    
    If Left$(strFile, Len(structProj.preview_prefix) + 1) = structProj.preview_prefix & "-" Then
        GetArtStatus = False
    End If
End Function

Sub UpdateRemoteContextCmds()
    'Adjusts access to the remote commands according to current selection in remote blog list
    'NB : unless frmMain.optStatusArt(0|1) managed via ShowSelectedArtStatus()
    Dim strKey As String
    Dim bArt As Boolean
    Dim bCat As Boolean
    
    Sleep 500 'avoids conflict with eventual ongoing double-click
    
    If frmMain.lstRemote.ListItems.Count = 0 Then
        bCat = False
        bArt = False
    Else
        If frmMain.lstRemote.SelectedItem.Selected = True Then
            strKey = frmMain.lstRemote.SelectedItem.Key
            
            If strKey = ".." Then
                bCat = False
                bArt = False
            Else
                If frmMain.FtpCli.RemoteFiles(strKey).FileType = 1 Then
                    'article
                    bCat = False
                    bArt = True
                Else
                    'category
                    bCat = True
                    bArt = False
                End If
            End If
        Else
            bCat = False
            bArt = False
        End If
    End If

    'frmMain.lblRemoteCats.Enabled = bCat
    frmMain.cmdRmRemoteCat.Enabled = bCat
    frmMain.cmdRenRemoteCat.Enabled = bCat
    
    'frmMain.lblRemoteArts.Enabled = bArt
    frmMain.cmdSeeRemoteArt.Enabled = bArt
    frmMain.cmdDelRemoteArt.Enabled = bArt
    frmMain.cmdRenRemoteArt.Enabled = bArt
End Sub

Sub ShowSelectedArtStatus()
    'Checks status (public/private) of current selected article in remote blog list
    'NB : information is shown in the user interface with the optStatusArt(0|1) switch
    Dim strKey As String
    Dim strFilename As String
    
    If structProj.with_preview = False Then Exit Sub 'only if required by project settings
    
    Sleep 500 'avoids conflict with eventual ongoing double-click
    
    If frmMain.lstRemote.SelectedItem.Selected = True Then
        strKey = frmMain.lstRemote.SelectedItem.Key
        
        If strKey = ".." Then
            ShowStatusCmd False
            Exit Sub
        End If
        
        If frmMain.FtpCli.RemoteFiles(strKey).FileType = 1 Then
            'only if this is a file
            ShowStatusCmd True
            strFilename = frmMain.FtpCli.RemoteFiles(strKey).FileName
            frmMain.optStatusArt(1).Value = GetArtStatus(strFilename)
            frmMain.optStatusArt(0).Value = Not (frmMain.optStatusArt(1).Value)
        Else
            ShowStatusCmd False
        End If
    Else
        'if no selection or selection lost, we hide command(s)
        ShowStatusCmd False
    End If
End Sub

Sub ShowStatusCmd(bOp As Boolean)
    'Shows/hides the switch command which shows and acts on status of remote articles
    'IN : sense of the operation (true:shows / false:hides)
    Dim nIdx As Integer
    
    For nIdx = 0 To frmMain.optStatusArt.UBound
        frmMain.optStatusArt(nIdx).Visible = bOp
        frmMain.shapStatus.Visible = bOp
    Next
End Sub

Sub EnableStatusCmd(bOp As Boolean)
    'Enables/disables the switch command which shows and acts on status of remote articles
    'IN : sense of the operation (true:enables / false:disables)
    Dim nIdx As Integer
    
    For nIdx = 0 To frmMain.optStatusArt.UBound
        frmMain.optStatusArt(nIdx).Enabled = bOp
    Next
End Sub

Sub ResetStatusCmd()
    'Resets optStatusArt(0|1) switch command to default state
    
    'project-independent properties
    frmMain.optStatusArt(0).Value = False
    frmMain.optStatusArt(1).Value = True 'show 'public' by default
    
    EnableStatusCmd False
    
    'project-dependent properties
    ShowStatusCmd Int(structProj.with_preview)
End Sub

Function SetGetArchive(strArt As String, strId As String, bOp As Boolean, nAddedNameLen As Integer, _
                        strServerAttachs As String, Optional bKeepSrc As Boolean = True) As Boolean
    'Stores/retrieves a given article to/from zip archive
    'IN :
    '- filename of the article to manipulate (without path since between archive and workspace only)
    '- identication string as "[date_heure_distante]_[catégorie]" to add beside article name in archive name
    '  (this is possible only because thisfucntion is called after an upload or download, or assimilated)
    '- flag indicating the sense of the operation (true:archiving / false:unpacking)
    '- number of characters which will diferenciate the article in archive and the one on server
    '  (this difference being due to eventual images renaming done during upload to avoid duplicates)
    '  (this value has no relationship with difference of size due to CRLF->LF conversion when upload to *nix)
    '  . if bOp=true (set), this info will be stored in the archive
    '  . if bOp=false (get), this will act as buffer variable which will receive the info read from archive
    '- list of final attached image names on server (after upload) ; list format is "name1,name2,name3,...,nameN"
    '  . if bOp=true (set), this infos list will be stored in archive
    '  . if bOp=false (get), this will act asbuffer variable which will receive the infos list from archive
    '- optional flag indicating if source file must be preserved (true, default) or deleted (false)
    'OUT : flag saying true:successfull or false:failure
    'NB : in the context of zip archive, original local paths of the attached images are kept in the local
    '     article itself (in "src" attribute of <img> tags as it is for article in local workspace context),
    '     while in server context (after upload) original local image paths are kept in an added HTML comment).
    '     (so, images in zip archive are not directly referenced by the article, but saved in case original
    '     local ones would be deleted by user, outside of BlosHome scope)
    Dim strWorkspace As String
    Dim strArchive As String
    Dim strZip As String
    
    Dim oZip As clsZip  'interface between zlib.dll and ZIP format
    
    Dim arAttach() As String
    Dim arExtracted() As String
    Dim nAddedFiles As Integer
    
    Dim strTmpPath As String
    Dim strTmpFile As String
    
    Dim nRet As Long
    Dim bRet As Boolean
    Dim nIdx As Integer
    Dim nIdx2 As Integer
    
    Dim nPos As Integer
    Dim nMax As Integer
    
    Const INFO = "info" 'INI-file containing extra infos to store in the zip archive
    
    'init
    SetGetArchive = False
    If InStr(strArt, "\") <> 0 Or InStr(strArt, "/") <> 0 Then Exit Function
    
    strWorkspace = GetWorkspace() & "\"
    strArchive = strWorkspace & "archive\"
    strZip = strArchive & strId & "_" & GetFilenamePrefix(strArt) & ".zip"
    
    If Dir(strArchive, vbDirectory) = "" Then MkDir strArchive
    
    'treats the article
    Set oZip = New clsZip
    On Error GoTo SetGetArchive_Error
    
    If bOp = True Then
        'create the ZIP archive (failure will trigger error #513)
        
        'pushes the article to queue
        nRet = oZip.FileAdd(strWorkspace & strArt, WithoutPath)
        If nRet <> 1 Then Err.Raise 513
        nAddedFiles = 1
        
        'pushes the infos file to queue
        strTmpFile = GetTmpPath(True)
        If Dir(strTmpPath, vbDirectory) = "" Then Err.Raise 513
        strTmpFile = strTmpFile & INFO
        If Dir(strTmpFile) <> "" Then Kill strTmpFile 'supprime eventuel fichier faisant obstruction ;)
        
        SaveIni strTmpFile, "Online", "Added Name Len", Trim$(Str(nAddedNameLen))
        SaveIni strTmpFile, "Online", "Attachs List", strServerAttachs
        If Dir(strTmpFile) = "" Then Err.Raise 513
        
        nRet = oZip.FileAdd(strTmpFile, WithoutPath)
        If nRet <> 1 Then Err.Raise 513
        nAddedFiles = nAddedFiles + 1
        
        'pushes the attached images to queue
        arAttach() = EnumLocalArtAnnexes(strWorkspace & strArt)
        If arAttach(1) = "failed" Then Err.Raise 513
        
        If arAttach(1) <> "" Then
            For nIdx = 1 To UBound(arAttach())
                nRet = oZip.FileAdd(arAttach(nIdx), WithoutPath) 'mise en file d'attente d'une image
                If nRet <> 1 Then Err.Raise 513                  '(soit, 1 fichier)
                nAddedFiles = nAddedFiles + 1
            Next
        End If
        
        'effective zipping (to .zip file prefixed with an ID allowing to retrieve upload date and category)
        bRet = oZip.WriteZip(strZip, True) 'default compression level (level 6)
        If bRet = False Then Err.Raise 513
        oZip.ZipClose
        
        'tests the zip and deletes the article in workspace if required
        '***LATER : go further, testing the integrity of zip file
        oZip.ZipOpen strZip
        If Dir(strZip) <> "" And oZip.inFileCount = nAddedFiles Then
            oZip.ZipClose
            If bKeepSrc = False Then Kill strWorkspace & strArt
        Else
            Err.Raise 513
        End If
    Else
        'extracts the zip archive in workspace (failure will trigger appropriated error according to progress)
    
        'opens the zip
        oZip.ZipOpen strZip
        If Not oZip.ZipIsOpen Then Err.Raise 514
                
        'locates the article file in the zip
        nMax = oZip.inFileCount
        If nMax < 1 Then Err.Raise 514
        
        nPos = 0
        For nIdx = 1 To nMax
            If oZip.inFileName(nIdx) = strArt Then
                nPos = nIdx
                Exit For
            End If
        Next
        If nPos = 0 Then Err.Raise 514
        
        'effective extraction of the article toward workspace
        If Dir(strWorkspace & strArt) <> "" Then Err.Raise 514
        bRet = oZip.ExtractSingleFile(nPos, strWorkspace, True, False, False)
        If bRet = False Then Err.Raise 514
        
        'retrieves the informations from "info" file (i.e. locates it -> extracts it -> reads it)
        '(from this point, a failure will imply error #515, since requires article deletion)
        nPos = 0
        For nIdx = 1 To nMax
            If oZip.inFileName(nIdx) = INFO Then
                nPos = nIdx
                Exit For
            End If
        Next
        If nPos = 0 Then Err.Raise 515
        
        strTmpPath = GetTmpPath(True)
        If Dir(strTmpPath, vbDirectory) = "" Then Err.Raise 515
        strTmpFile = strTmpPath & INFO
        If Dir(strTmpFile) <> "" Then Kill strTmpFile 'deletes eventual obstructing file
        
        bRet = oZip.ExtractSingleFile(nPos, strTmpPath, True, True, False)
        If bRet = False Then Err.Raise 515
        
        nAddedNameLen = LoadIni(strTmpFile, "Online", "Added Name Len")
        strServerAttachs = LoadIni(strTmpFile, "Online", "Attachs List")
        
        'determines if necessary to restore the images in their original location
        '(from this point, a failure will imply error #516, since requires article and extracted images deletion)
        arAttach() = EnumLocalArtAnnexes(strWorkspace & strArt)
        If arAttach(1) = "failed" Then Err.Raise 516 'will delete article in workspace
        If arAttach(1) <> "" Then
            nMax = oZip.inFileCount
            If nMax < 1 Then Err.Raise 516
            
            For nIdx = 1 To UBound(arAttach)
                If Dir(arAttach(nIdx)) = "" Then
                    'localises the image file inside zip
                    nPos = 0
                    For nIdx2 = 1 To nMax
                        If oZip.inFileName(nIdx) = arAttach(nIdx) Then
                            nPos = nIdx
                            Exit For
                        End If
                    Next
                    If nPos = 0 Then Err.Raise 516
                    
                    'extracts the image toward its original location
                    bRet = oZip.ExtractSingleFile(nPos, strWorkspace, True, False, False)
                    If bRet = False Then Err.Raise 516
                    
                    'informs the extracted images list for eventual undo on error
                    ReDim Preserve arExtracted(UBound(arExtracted) + 1)
                    arExtracted(UBound(arExtracted)) = arAttach(nIdx)
                End If
            Next
        End If
               
        'tests the article in workspace
        If Dir(strWorkspace & strArt) = "" Then Err.Raise 516
        
        'deletes the zip archive if required
        oZip.ZipClose
        If bKeepSrc = False Then Kill strZip
    End If
    
    SetGetArchive = True

SetGetArchive_End:
   On Error GoTo 0
   Exit Function

SetGetArchive_Error:
    If oZip.ZipIsOpen = True Then oZip.ZipClose
    
    Select Case Err.Number
        Case 513
            'failed to archive => deletes eventual incomplete zip
            If Dir(strZip) <> "" Then Kill strZip
        Case 515
            'failed to extract article => deletes the article in workspace
            If Dir(strWorkspace & strArt) <> "" Then Kill strWorkspace & strArt
        Case 516
            'failed to extract image => deletes the article and eventual extracted image
            If Dir(strWorkspace & strArt) <> "" Then Kill strWorkspace & strArt
            If UBound(arExtracted) > 0 Then
                For nIdx = 1 To UBound(arExtracted)
                    If Dir(arExtracted(nIdx)) <> "" Then Kill arExtracted(nIdx)
                Next
            End If
        Case Else 'including 514
            '(does nothing)
    End Select
    
    If strTmpFile <> "" And Dir(strTmpFile) <> "" Then Kill strTmpFile
    
    Set oZip = Nothing
    Resume SetGetArchive_End
End Function

Function BuildIdArchive(strArt As String) As String
    'Builds the identication string to pass as strId to SetGetArchive()
    'IN : article filename for which generates this ID (article must be reachable on current remote directory)
    'OUT : identifier as "jjmmaaaa-hhmm_cat[-subcat[-subsubcat[...]]]" or "" in case of failure
    '***LATER :
    '- check if connected and remote article file is in current category
    '- adds the secondes in the ID part about hour (i.e. "hhmmss") ; requires to modify FtpCli
    Dim strId As String
    
    BuildIdArchive = ""
    
    On Error Resume Next
    strId = frmMain.FtpCli.RemoteFiles("K" & strArt).FileDate 'date/time on server
    If Right$(strId, 3) = ":00" Then strId = Left$(strId, Len(strId) - 3) 'sans secondes (cf. ***LATER)
    On Error GoTo 0
    If strId = "" Then Exit Function
    
    strId = Replace(strId, "/", "")                   'removes the signs
    strId = Replace(strId, ":", "")
    strId = Replace(strId, " ", "-")
    
    On Error Resume Next
    strId = strId & "_" & GetCurrRemoteCat("-", True) 'current remote category
    On Error GoTo 0
    If InStr(strId, "_") = 0 Then Exit Function
    
    strId = Replace(strId, "<", "")                   'removes the signs if <ROOT>
    strId = Replace(strId, ">", "")
    
    BuildIdArchive = strId
End Function

Function DelCurrRemoteArt(arNotDeleted() As String, _
                            Optional strLocalCopy As String = "", _
                            Optional strAttachList As String) As Boolean
    'Deletes the selected remote article and its eventual attached images
    'IN :
    '- buffer array which will receive the list of the files which were not deleted in spite of upload failure
    '  . if there is one "failed" element, it means the failure happened during analysis, before any deletion)
    '  . if all has been well deleted, the array contains a single empty element ("")
    '- path to a local copy avoiding all real download to be able to analyze the article ("" = donwload required)
    '- list of attached-images-to-delete to provide with strLocalCopy to avoid all download, to be able to
    '  analyze the remote article (list will be formatted as "name1,name2,name3,...,nameN")
    'OUT : true if successfull, false otherwise
    'WARNING : caller is responsible to check if there is a valid selection in lstRemote
    Dim strKey As String
    Dim strFilename As String
    Dim strRemoteFile As String
    
    Dim strTmpFile As String
    Dim strArt As String
    
    Dim bRet As Boolean
    Dim nIdx As Integer
    
    On Error GoTo DelCurrRemoteArt_Error

    'init (no checking about the existence of a selection or not ; see WARNING ahead)
    strKey = frmMain.lstRemote.SelectedItem.Key
    strFilename = frmMain.FtpCli.RemoteFiles(strKey).FileName
    strRemoteFile = frmMain.FtpCli.RemoteDir & "/" & strFilename

    'analyzes the attached images (on local copy provided or after download)
    If strLocalCopy = "" Then
        'starts from remote file
        strTmpFile = GetTmpFile("BH")
        If strTmpFile = "" Then Err.Raise 513
        bRet = frmMain.FtpCli.Met_GET(strFilename, strTmpFile, "A")
        If bRet = False Then Err.Raise 513
        strArt = strTmpFile
    
        arNotDeleted = EnumLocalArtAnnexes(strArt, False)
        If arNotDeleted(1) = "failed" Then Err.Raise 513
    Else
        'starts from local copy
        strArt = strLocalCopy
        arNotDeleted = Split(strAttachList, ",")
        arNotDeleted = SetOptionBase(arNotDeleted(), 1)
    End If
    
    'adds the article filename to arNotDeleted(), in order to return an exhaustive info if error
    ReDim Preserve arNotDeleted(UBound(arNotDeleted) + 1)
    arNotDeleted(UBound(arNotDeleted)) = strFilename
    
    'deletes the article (with update of local cache)
    bRet = frmMain.FtpCli.Met_DELETE(strFilename)
    If bRet = False Then Err.Raise 513
    ReDim Preserve arNotDeleted(UBound(arNotDeleted) - 1)
    MetaCache strRemoteFile, "", "REMOVE"
    
    'deletes the eventual attached images
    If arNotDeleted(1) <> "" Then
        Dim nMax As Integer
        nMax = UBound(arNotDeleted)
        For nIdx = nMax To 1 Step -1
            bRet = frmMain.FtpCli.Met_DELETE(arNotDeleted(UBound(arNotDeleted)))
            If bRet = False Then Err.Raise 513
            If UBound(arNotDeleted) > 1 Then
                ReDim Preserve arNotDeleted(UBound(arNotDeleted) - 1)
            Else
                arNotDeleted(1) = ""
            End If
        Next
    End If
    
    DelCurrRemoteArt = True

DelCurrRemoteArt_End:
    PopulateRemoteList
    On Error GoTo 0
    Exit Function

DelCurrRemoteArt_Error:
    If strTmpFile <> "" Then Kill strTmpFile 'below, reverses to obtain a more logical presentation
    If arNotDeleted(UBound(arNotDeleted)) = strFilename Then arNotDeleted = RevStrArray(arNotDeleted)
    DelCurrRemoteArt = False
    Resume DelCurrRemoteArt_End
End Function

Function FixObjectTag(strIn As String) As String
    '******************************************************************************************************
    'This function is experimental and doesn't seems to be enough to allow video integration in an article.
    'TinyMCE seems to corrupt any <object>, but not only at closing-time, and this fct has been designed
    'to act after editor closing... So, it's not enough since <object> can be already corrupted, or can be
    'corrupted again at next opening (in the BlosHome editor). So, a better way would be to find a solution
    'from within TinyMCE (maybe reducing its cleaning action or using a IE-compatible video plugin)
    '******************************************************************************************************
    'Ensures any <object> tag contains the required param about URL to work in the editor (IE activeX)
    'IN : HTML content to parse
    'OUT : resulting HTML content
    'FIX : this function tries to fix a known bug in IE which corrupts the integrety of <object> tags
    '***TODO : this function and its call in wbEditor_StatusTextChange() ; also, maybe, extend to <iframe>
    '***LATER : this function is subject to modification accoring to evolution of IE-bug and TinyMCE
    Dim strSearch1 As String
    Dim strSearch2 As String
    
    Dim nTagStart As Integer
    Dim nTagStop As Integer
    Dim strTag As String
    
    Dim nURLStart As Integer
    Dim nURLStop As Integer
    Dim strURL As String
    
    Dim strSrcParam As String
    Dim strOUT As String
    
    strOUT = strIn
    
    Do
        'isolates every <object>...</object> tag pairs
        strSearch1 = "<object "
        nTagStart = InStr(nTagStart + 1, LCase(strOUT), strSearch1, vbTextCompare)
        If nTagStart <> 0 Then
            strSearch1 = "</object>"
            nTagStop = InStr(nTagStart + 1, LCase(strOUT), strSearch1, vbTextCompare) + Len(strSearch1) - 1
            strTag = Mid$(strOUT, nTagStart, nTagStop - nTagStart + 1)
            
            'determines the media URL in "data" attribute
            strSearch2 = "data=" & Chr(34)
            nURLStart = InStr(nTagStart + 1, LCase(strTag), strSearch2, vbTextCompare) + Len(strSearch2)
            strSearch2 = Chr(34)
            nURLStop = InStr(nURLStart, LCase(strTag), strSearch2, vbTextCompare) - 1
            strURL = Mid$(strTag, nURLStart, nURLStop - nURLStart + 1)
            
            'adds the "movie" parameter (***LATER : maybe check if effectively missing before)
            strSrcParam = "<param name='movie' value='" & strURL & "'>"
            strTag = Left$(strTag, Len(strTag) - Len(strSearch1)) & strSrcParam & strSearch1
            
            'replaces the tag in resulting HTML context
            strOUT = Left$(strOUT, nTagStart - 1) & strTag & Right$(strOUT, Len(strOUT) - nTagStop)
        End If
    Loop While nTagStart <> 0
    
    FixObjectTag = strOUT
End Function

