Attribute VB_Name = "basProj"
'BlosHome (c) FFh Lab / Eric Lequien, 2009-2013 - http://ffh-lab.com
'This module contains the mechanism to manage project file (these operations will be logged)

Option Explicit

Public Type PROJECT     'projet settings
    FileName As String  'autodefined from title on SaveProject() the 1st time
    
    blosxom_url As String
    title As String
    
    host As String
    port As String
    user As String
    pass As String
    
    art_root As String
    art_ext As String
    flav_ext As String
    excluded_paths As String
    
    img_max As String
    art_encode As String 'ANSI / UTF-8
    
    with_preview As String
    preview_prefix As String
    preview_pass As String
    
    with_chapo As String
    chapo_limit As String
        
    css As String       'contains CSS file path, unless when used as data buffer during frmProject saving
End Type

Public structProj As PROJECT 'populated from a project file at project opening

Public Const PROJFILE_EXT = ".bhp"
Public Const PROJECTS_DIR = "data"
Public Const ARTICLE_CSS = "article.css"

Function GetWorkspace() As String
    'Determines current local project directory from project filename (just removing final file extension)
    If structProj.FileName = "" Then Exit Function
    GetWorkspace = App.Path & "\" & PROJECTS_DIR & "\" & Left$(structProj.FileName, _
                                                Len(structProj.FileName) - Len(PROJFILE_EXT))
End Function

Sub SaveProject()
    'Saves/creates the different components of a project from current structProj structure
    'NB : a project is : an INI project file, a directory acting as local workspace, a CSS file for the editor
    'USE : this fct is called after any (re)definition of a project (i.e. when closing frmProject)
    Dim strProjDir As String
    Dim strCSSFile As String
    
    Dim strProjFile As String
    Dim strSection As String
    Dim strKey As String
    Dim strBuffer As String
    
    'autodefines a project filename from its first time title
    If structProj.FileName = "" Then
        structProj.FileName = CheckAndMkFilename(LCase(structProj.title), True, True) & PROJFILE_EXT
    End If
    
    'defines project file and workspace paths
    strProjFile = App.Path & "\" & PROJECTS_DIR & "\" & structProj.FileName
    strProjDir = GetWorkspace() 'same name as proj. file without extension
    
    'creates workspace and its subdir(s) if doesn't yet exist
    If Dir(strProjDir, vbDirectory) = "" Then MkDir strProjDir
    If Dir(strProjDir & "\archive", vbDirectory) = "" Then MkDir strProjDir & "\archive"
    
    'creates or updates the CSS file
    strCSSFile = strProjDir & "\" & ARTICLE_CSS
    SaveText strCSSFile, structProj.css
    structProj.css = strCSSFile 'replace previous data with file path (which now contains updated data)
    
    'ensures to save a normalized project
    NormalizeProject
    
    'creates or updates project file with current parameters
    'The 'Main' section
    strSection = "Main"
    
        'Project Title
        strKey = "Title"
        strBuffer = structProj.title
        SaveIni strProjFile, strSection, strKey, strBuffer
        
    'The 'Server' section
    strSection = "Server"
    
        'Host
        strKey = "Host"
        strBuffer = structProj.host
        SaveIni strProjFile, strSection, strKey, strBuffer

        'Port
        strKey = "Port"
        strBuffer = structProj.port
        SaveIni strProjFile, strSection, strKey, strBuffer

        'User
        strKey = "User"
        strBuffer = structProj.user
        SaveIni strProjFile, strSection, strKey, strBuffer
        
        'Pass
        strKey = "Pass"
        strBuffer = structProj.pass
        SaveIni strProjFile, strSection, strKey, strBuffer
        
    'The 'Blog' section
    strSection = "Tree"
    
        'Articles Root
        strKey = "Art. Root"
        strBuffer = structProj.art_root
        SaveIni strProjFile, strSection, strKey, strBuffer

        'Articles Extension
        strKey = "Art. Ext."
        strBuffer = structProj.art_ext
        SaveIni strProjFile, strSection, strKey, strBuffer
        
        'Flavour Extension
        strKey = "Flavour Ext."
        strBuffer = structProj.flav_ext
        SaveIni strProjFile, strSection, strKey, strBuffer

        'Excluded Paths
        strKey = "Excluded Paths"
        strBuffer = structProj.excluded_paths
        SaveIni strProjFile, strSection, strKey, strBuffer

        'Blosxom URL
        strKey = "Blosxom URL"
        strBuffer = structProj.blosxom_url
        SaveIni strProjFile, strSection, strKey, strBuffer
        
    'The 'Article' section
    strSection = "Article"
    
        'Image Max.
        strKey = "Image Max."
        strBuffer = structProj.img_max
        SaveIni strProjFile, strSection, strKey, strBuffer

        'Articles Encoding
        strKey = "Art. Encoding"
        strBuffer = structProj.art_encode
        SaveIni strProjFile, strSection, strKey, strBuffer
        
    'The 'Status' section (private preview or public)
    strSection = "Status"

        strKey = "w/Preview"
        strBuffer = structProj.with_preview
        SaveIni strProjFile, strSection, strKey, strBuffer
        
        strKey = "Preview Prefix"
        strBuffer = structProj.preview_prefix
        SaveIni strProjFile, strSection, strKey, strBuffer
        
        strKey = "Preview Pass"
        strBuffer = structProj.preview_pass
        SaveIni strProjFile, strSection, strKey, strBuffer
    
    'The 'Chapo' section
    strSection = "Chapo"
    
        strKey = "w/Chapo"
        strBuffer = structProj.with_chapo
        SaveIni strProjFile, strSection, strKey, strBuffer
        
        strKey = "Chapo Limit"
        strBuffer = structProj.chapo_limit
        SaveIni strProjFile, strSection, strKey, strBuffer
    
    ArrangeINISections strProjFile
    
    'manages log file
    DoLog "Save project '" & strProjFile & "' and its workspace '" & strProjDir & "'"
End Sub

Sub NormalizeProject()
    'Normalizes some project settings which need a strict format
    '***LATER : just a starting point, this list being an extensible place holder to manage any normalization
    
    'article extension must start with a dot (.)
    If Left$(structProj.art_ext, 1) <> "." Then structProj.art_ext = "." & structProj.art_ext
End Sub

Sub LoadUnloadProject(bOp As Boolean)
    'Loads or unloads the different project settings
    'IN : a flag indicating the sense of the operation :
    ' - If bOp is true, we load the projet as this :
    '   . project file (current structProj.filename being the starting point to complete the other members)
    '   . css file (we just put its full path in structProj.css)
    '   (this kind of call is made on frmSelect exit after a frmMain/Project/Open menu cmd)
    ' - If bOp is false, we unload the project as this :
    '   . we reset entire structProj structure
    '   . we erase eventual temporary files leaved behind
    'NB : options which will be modified during normalization process (ex. art_ext w/o '.'), will
    '     be saved in their new format at next SaveProject()
    Dim strProjFile As String
    
    strProjFile = App.Path & "\" & PROJECTS_DIR & "\" & structProj.FileName
    
    If bOp = True Then
        'Loads project previously indicated in structProj.filename
        '---------------------------------------------------------
        Dim strSection As String
        Dim strKey As String
        Dim strBuffer As String
        Dim bIncomplet As Boolean
        
        bIncomplet = False
        EnableLocalWork False
        
        'store CSS file path
        '(effective loading of CSS data in TinyMCE's arbo will be done every time editor will be loaded)
        structProj.css = GetWorkspace() & "\" & ARTICLE_CSS
        
        'The 'Main' section
        strSection = "Main"
        
            'Project Title
            strKey = "Title"
            strBuffer = LoadIni(strProjFile, strSection, strKey)
            If strBuffer = "" Then bIncomplet = True
            structProj.title = strBuffer
            
        'The 'Server' section
        strSection = "Server"
        
            'Host
            strKey = "Host"
            strBuffer = LoadIni(strProjFile, strSection, strKey)
            If strBuffer = "" Then bIncomplet = True
            structProj.host = strBuffer
    
            'Port
            strKey = "Port"
            strBuffer = LoadIni(strProjFile, strSection, strKey)
            If strBuffer = "" Then bIncomplet = True
            structProj.port = strBuffer
    
            'User
            strKey = "User"
            strBuffer = LoadIni(strProjFile, strSection, strKey)
            If strBuffer = "" Then bIncomplet = True
            structProj.user = strBuffer
            
            'Pass
            strKey = "Pass"
            strBuffer = LoadIni(strProjFile, strSection, strKey)
            If strBuffer = "" Then bIncomplet = True
            structProj.pass = strBuffer
            
        'The 'Blog' section
        strSection = "Tree"
        
            'Articles Root
            strKey = "Art. Root"
            strBuffer = LoadIni(strProjFile, strSection, strKey)
            If strBuffer = "" Then bIncomplet = True
            structProj.art_root = strBuffer
    
            'Articles Extension (w/ normalisation)
            strKey = "Art. Ext."
            strBuffer = LoadIni(strProjFile, strSection, strKey)
            If strBuffer = "" Then bIncomplet = True
            structProj.art_ext = strBuffer
            
            'Flavour Extension
            strKey = "Flavour Ext."
            strBuffer = LoadIni(strProjFile, strSection, strKey)
            If strBuffer = "" Then bIncomplet = True
            structProj.flav_ext = strBuffer
            
            'Excluded Paths (facultatif)
            strKey = "Excluded Paths"
            strBuffer = LoadIni(strProjFile, strSection, strKey)
            structProj.excluded_paths = strBuffer
            
            'Blosxom URL
            strKey = "Blosxom URL"
            strBuffer = LoadIni(strProjFile, strSection, strKey)
            If strBuffer = "" Then bIncomplet = True
            structProj.blosxom_url = strBuffer
        
        'The 'Article' section
        strSection = "Article"
        
            'Image Max. (facultatif)
            strKey = "Image Max."
            strBuffer = LoadIni(strProjFile, strSection, strKey)
            structProj.img_max = strBuffer
    
            'Articles Extension
            strKey = "Art. Encoding"
            strBuffer = LoadIni(strProjFile, strSection, strKey)
            If strBuffer = "" Then bIncomplet = True
            structProj.art_encode = strBuffer
        
        'The 'Status' section
        strSection = "Status"
    
            strKey = "w/Preview"
            strBuffer = LoadIni(strProjFile, strSection, strKey)
            If strBuffer = "" Then bIncomplet = True
            structProj.with_preview = strBuffer
            
            strKey = "Preview Prefix" '(facultatif)
            strBuffer = LoadIni(strProjFile, strSection, strKey)
            structProj.preview_prefix = strBuffer
            
            strKey = "Preview Pass" '(facultatif)
            strBuffer = LoadIni(strProjFile, strSection, strKey)
            structProj.preview_pass = strBuffer
        
        'The 'Chapo' section
        strSection = "Chapo"
        
            strKey = "w/Chapo"
            strBuffer = LoadIni(strProjFile, strSection, strKey)
            If strBuffer = "" Then bIncomplet = True
            structProj.with_chapo = strBuffer
            
            strKey = "Chapo Limit" '(forced to factory value, rather than blocking, if absent)
            strBuffer = LoadIni(strProjFile, strSection, strKey)
            If strBuffer = "" Then strBuffer = "more"
            structProj.chapo_limit = strBuffer
        
        If bIncomplet = True Then
            MsgBox arMsg(129) & " : " & arMsg(130) & " !"
            frmProject.bNoCancel = True
            frmProject.Show 1
        End If
    
        NormalizeProject
        ApplyCleanProject True
        EnableLocalWork True
        DoLog "Load project '" & strProjFile & "'" 'manage log
    Else
        'Close eventual connection, clean-up disk environment and unload project
        '-----------------------------------------------------------------------
        Dim strEditorCurrent As String
        Dim strContentCSS As String
        
        If Left$(frmMain.cmdCnnx.Caption, 1) <> "C" Then CnnxDcnnx False
        EnableLocalWork False
        PopulateClearLocalList False
        ApplyCleanProject False
        
        strContentCSS = App.Path & "\" & EDITOR_DIR & "\" & EDITOR_CONTENT_CSS
        strEditorCurrent = App.Path & "\" & EDITOR_DIR & "\" & EDITOR_CURRENT
        
        If Dir(strContentCSS) <> "" Then Kill strContentCSS
        If Dir(strEditorCurrent) <> "" Then Kill strEditorCurrent
        
        structProj.FileName = ""
        structProj.title = ""
        structProj.blosxom_url = ""
        
        structProj.host = ""
        structProj.port = ""
        structProj.user = ""
        structProj.pass = ""
        
        structProj.art_root = ""
        structProj.art_ext = ""
        structProj.flav_ext = ""
        structProj.excluded_paths = ""
        
        structProj.img_max = ""
        structProj.art_encode = ""
        
        structProj.with_preview = ""
        structProj.preview_prefix = ""
        structProj.preview_pass = ""
        
        structProj.with_chapo = ""
        structProj.chapo_limit = ""
        
        structProj.css = ""
        
        'manage log file
        DoLog "Unload project '" & strProjFile & "'"
    End If
End Sub

Sub ApplyCleanProject(bOp)
    'Apply project settings to depending UI elements
    'IN : a flag indicating the sense of the operation
    '     - True to apply freshly (re)loaded project
    '       (called just after loading or modification of a project, to refresh the interface)
    '     - False to remove any trace of project in UI elements that are not binded with workspace/remote-list
    If bOp = True Then
        frmMain.barStatus.Panels.item("url").Text = structProj.blosxom_url
    Else
        frmMain.barStatus.Panels.item("url").Text = ""
    End If
    
    PopulateClearLocalList bOp
    EnableTransCmds False, True
    ResetStatusCmd
End Sub

Sub ApplyCnnx(bCnnx As Boolean)
    'Adjusts the interface to the observed connection status (called at every connection and deconnection)
    If bCnnx = True Then
        'Connected
        frmMain.FtpCli.Met_CD structProj.art_root
        frmMain.FtpCli.Met_DIR
        PopulateRemoteList
        
        frmMain.cmdCnnx.Caption = arMsg(117)
        frmMain.barStatus.Panels.item("cnnx").Text = arMsg(118)
        
        EnableRemoteWork True
        StartStopTimer frmMain.timerCnnx, True, 60000, nTimerPass 'chaque minute
    Else
        'Deconnected
        frmMain.lblRemotePath.Caption = ""
        frmMain.lstRemote.ListItems.Clear
        frmMain.cmdCnnx.Caption = arMsg(119)
        frmMain.barStatus.Panels.item("cnnx").Text = arMsg(120)
        
        EnableRemoteWork False
        StartStopTimer frmMain.timerCnnx, False
    End If
End Sub

Sub DelProject(strProjFile As String)
    'Deletes a given project file and its workspace if empty
    'IN : full path of the project file
    Dim strProjDir As String
    Dim bKeepDir As Boolean
    Dim strMsg As String
    
    strProjDir = Left$(strProjFile, Len(strProjFile) - Len(PROJFILE_EXT))
    
    Kill strProjFile
    
    If Dir(strProjDir, vbDirectory) <> "" Then
        If Dir(strProjDir & "\*.*") = "" Then
            bKeepDir = False
            RmDir strProjDir
        Else
            bKeepDir = True
            MsgBox arMsg(131) & " '" & strProjDir & "' " & arMsg(132) & "." & vbNewLine & arMsg(133) & " !"
        End If
    End If

    'manage log
    strMsg = arMsg(121) & " '" & strProjFile & "'"
    If bKeepDir = False Then
        strMsg = strMsg & " " & arMsg(122) & " '" & strProjDir & "'"
    Else
        strMsg = strMsg & " " & arMsg(123) & " '" & strProjDir & "'"
    End If
    DoLog strMsg
End Sub

Sub EnableLocalWork(bOnOff As Boolean)
    'Manages the UI elements to be able to work or not on a project locally
    'IN : a flag indicationg the sense of the operation
    '     - called with True after LoadProject for existing project, and frmProject/OK for new project
    '     - called with False in process of closing a loaded project, before to do a File/New|Open|Del
    Dim strMsg As String
    
    If bOnOff = True Then
        frmMain.Caption = App.EXEName & " - " & structProj.title
    Else
        frmMain.Caption = App.EXEName
    End If
    
    frmMain.cmdProj.Enabled = bOnOff
    frmMain.cmdCnnx.Enabled = bOnOff
    
    frmMain.framLocal.Enabled = bOnOff
    frmMain.lstLocal.Enabled = bOnOff
    
    frmMain.lblLocalArts.Enabled = bOnOff
    frmMain.cmdMkLocalArt.Enabled = bOnOff
    frmMain.cmdCopyLocalArt.Enabled = bOnOff
    frmMain.cmdEditLocalArt.Enabled = bOnOff
    frmMain.cmdDelLocalArt.Enabled = bOnOff
    frmMain.cmdRenLocalArt.Enabled = bOnOff
    
    If bOnOff = True Then
        strMsg = arMsg(124)
    Else
        strMsg = arMsg(125)
    End If
    DoLog strMsg
End Sub

Sub EnableRemoteWork(bOnOff As Boolean)
    'Manages the UI elements to be able to work or not on a project remotely
    '(called after cmdCnnx succeeds ; i.e. we're connected to the FTP server)
    Dim strMsg As String
    
    frmMain.framRemote.Enabled = bOnOff
    frmMain.lstRemote.Enabled = bOnOff
    
    frmMain.lblRemoteCats.Enabled = bOnOff
    frmMain.cmdMkRemoteCat.Enabled = bOnOff
    frmMain.cmdRmRemoteCat.Enabled = bOnOff
    frmMain.cmdRenRemoteCat.Enabled = bOnOff
    
    frmMain.lblRemoteArts.Enabled = bOnOff
    frmMain.cmdSeeRemoteArt.Enabled = bOnOff
    frmMain.cmdDelRemoteArt.Enabled = bOnOff
    frmMain.cmdRenRemoteArt.Enabled = bOnOff
    
    EnableTransCmds bOnOff
    
    EnableStatusCmd bOnOff
    If bOnOff = True Then
        ShowStatusCmd False     'enabled but hidden until effective selection of an article in the list
        UpdateRemoteContextCmds
    Else
        ResetStatusCmd          'we restore status as it was before connection (depending of project settings)
    End If
    
    If bOnOff = True Then
        strMsg = arMsg(126) & " '" & structProj.host & ":" & structProj.port & "' as '" & _
                            structProj.user & "' " & arMsg(127)
    Else
        strMsg = App.title & " " & arMsg(128)
    End If
    DoLog strMsg
End Sub

Sub EnableTransCmds(bOnOff As Boolean, Optional bProjSynchro As Boolean = False)
    'Manages transfer commands between local workspace and remote blog (to allow upload/download or not)
    'IN : a flag indicationg the sense of the operation
    '     another flag indicating if we have to manage visibility of project-depending commands
    'NB : use of second argIN flag only when project or parameters change, so at ApplyCleanProject()
    Dim nIdx As Integer
    
    'shows/hides project-dependent control-commands
    If bProjSynchro = True Then
        frmMain.cmdUploadArt(0).Visible = structProj.with_preview 'w/ or w/o private-preview
    End If
    
    'enables/disables control-commands
    For nIdx = 0 To 1
        frmMain.cmdUploadArt(nIdx).Enabled = bOnOff
        frmMain.cmdDownloadArt(nIdx).Enabled = bOnOff
    Next
        
    'frmMain.cmdGetArtCopy.Enabled = bOnOff
    'frmMain.cmdBackupBlog.Enabled = bOnOff
    'frmMain.cmdCleanupBlog.Enabled = bOnOff
End Sub

