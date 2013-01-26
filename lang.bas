Attribute VB_Name = "basLang"
'BlosHome (c) FFh Lab / Eric Lequien, 2009-2013 - http://ffh-lab.com
'This module contains the mecanism to manage the localized .lng files
'NB : language will have an impact on UI, messages displayed, help file used and editor charset.

Option Explicit
Option Base 0

Const MSG_MAX = 142 'must match highest index in "[Messages]" section of .lng file
Public arMsg(0 To MSG_MAX) As String 'will store all messages which may be presented to user

Function LoadLangFile(strFile As String, strPart As String) As Boolean
    'Loads the given language file
    'IN :
    ' - the language file to load
    ' - the part to load in "main", "about", "edit", "progress", "project", "select"
    '   (ie. frame names, unless "main" meaning "all about main frame and app messages")
    'OUT : true if well loaded, false otherwise (problem of integrity)
    'WARNING : every section will be collected at loading-time of concerned frame,
    '          unless frmEdit, managed w/ frmMain/cmd[MkLocalArt|EditLocalArt], since stays loaded all session long
    'NB: we use fixed messages rather than the ones from the ".lng" file until the .lng's [Messages] section be loaded
    '***LATER : optimize, removing the redondant code block ; maybe a LoadItem() for every element
    '           and going through do/while loop rather than fixed list (every entry having name of associated ctrl)
    Dim nIdx As Integer
    Dim strSection As String
    Dim strKey As String
    Dim strBuffer As String
    Dim strMsg As String
       
    On Error GoTo LoadLangFile_Err
    If Dir(strFile) = "" Then Error 1
        
    Select Case strPart
        Case "main"
            'The 'Messages' section
            '(we treat this ahead to have error messages available for the rest of the fct)
            strSection = "Messages"
            For nIdx = 0 To MSG_MAX
                strKey = nIdx
                strBuffer = LoadIni(strFile, strSection, strKey)
                If strBuffer = "" Then Error 2
                arMsg(nIdx) = strBuffer
                If arMsg(nIdx) <> strBuffer Then Error 3
            Next
            
            'The 'Main Interface' section
            strSection = "Main Interface"
                
            'mnuProj
            strKey = "mnuProj"
            strBuffer = LoadIni(strFile, strSection, strKey)
            If strBuffer = "" Then Error 2
            frmMain.mnuProj.Caption = strBuffer
            If frmMain.mnuProj.Caption <> strBuffer Then Error 3
            
            'mnuProj_New
            strKey = "mnuProj_New"
            strBuffer = LoadIni(strFile, strSection, strKey)
            If strBuffer = "" Then Error 2
            frmMain.mnuProj_New.Caption = strBuffer
            If frmMain.mnuProj_New.Caption <> strBuffer Then Error 3
            
            'mnuProj_Open
            strKey = "mnuProj_Open"
            strBuffer = LoadIni(strFile, strSection, strKey)
            If strBuffer = "" Then Error 2
            frmMain.mnuProj_Open.Caption = strBuffer
            If frmMain.mnuProj_Open.Caption <> strBuffer Then Error 3
            
            'mnuProj_Del
            strKey = "mnuProj_Del"
            strBuffer = LoadIni(strFile, strSection, strKey)
            If strBuffer = "" Then Error 2
            frmMain.mnuProj_Del.Caption = strBuffer
            If frmMain.mnuProj_Del.Caption <> strBuffer Then Error 3
            
            'mnuLang
            strKey = "mnuLang"
            strBuffer = LoadIni(strFile, strSection, strKey)
            If strBuffer = "" Then Error 2
            frmMain.mnuLang.Caption = strBuffer
            If frmMain.mnuLang.Caption <> strBuffer Then Error 3
            
            'mnuInfo_Help
            strKey = "mnuInfo_Help"
            strBuffer = LoadIni(strFile, strSection, strKey)
            If strBuffer = "" Then Error 2
            frmMain.mnuInfo_Help.Caption = strBuffer
            If frmMain.mnuInfo_Help.Caption <> strBuffer Then Error 3
            
            'mnuInfo_About
            strKey = "mnuInfo_About"
            strBuffer = LoadIni(strFile, strSection, strKey)
            If strBuffer = "" Then Error 2
            frmMain.mnuInfo_About.Caption = strBuffer
            If frmMain.mnuInfo_About.Caption <> strBuffer Then Error 3
        
            'cmdProj button
            strKey = "cmdProj"
            strBuffer = LoadIni(strFile, strSection, strKey)
            If strBuffer = "" Then Error 2
            frmMain.cmdProj.Caption = strBuffer
            If frmMain.cmdProj.Caption <> strBuffer Then Error 3
            
            'cmdCnnx button
            strKey = "cmdCnnx"
            strBuffer = LoadIni(strFile, strSection, strKey)
            If strBuffer = "" Then Error 2
            frmMain.cmdCnnx.Caption = strBuffer
            If frmMain.cmdCnnx.Caption <> strBuffer Then Error 3
            
            'framLocal framework
            strKey = "framLocal"
            strBuffer = LoadIni(strFile, strSection, strKey)
            If strBuffer = "" Then Error 2
            frmMain.framLocal.Caption = strBuffer
            If frmMain.framLocal.Caption <> strBuffer Then Error 3
            
            'lstLocalRemote header 1
            strKey = "lstLocalRemote_1"
            strBuffer = LoadIni(strFile, strSection, strKey)
            If strBuffer = "" Then Error 2
            frmMain.lstLocal.ColumnHeaders(1).Text = strBuffer
            If frmMain.lstLocal.ColumnHeaders(1).Text <> strBuffer Then Error 3
            
            frmMain.lstRemote.ColumnHeaders(1).Text = strBuffer
            If frmMain.lstRemote.ColumnHeaders(1).Text <> strBuffer Then Error 3
            
            'lstLocalRemote header 2
            strKey = "lstLocalRemote_2"
            strBuffer = LoadIni(strFile, strSection, strKey)
            If strBuffer = "" Then Error 2
            frmMain.lstLocal.ColumnHeaders(2).Text = strBuffer
            If frmMain.lstLocal.ColumnHeaders(2).Text <> strBuffer Then Error 3
            
            frmMain.lstRemote.ColumnHeaders(2).Text = strBuffer
            If frmMain.lstRemote.ColumnHeaders(2).Text <> strBuffer Then Error 3
            
            'lstLocalRemote header 3
            strKey = "lstLocalRemote_3"
            strBuffer = LoadIni(strFile, strSection, strKey)
            If strBuffer = "" Then Error 2
            frmMain.lstLocal.ColumnHeaders(3).Text = strBuffer
            If frmMain.lstLocal.ColumnHeaders(3).Text <> strBuffer Then Error 3
            
            frmMain.lstRemote.ColumnHeaders(3).Text = strBuffer
            If frmMain.lstRemote.ColumnHeaders(3).Text <> strBuffer Then Error 3
            
            'lstLocalRemote header 4
            strKey = "lstLocalRemote_4"
            strBuffer = LoadIni(strFile, strSection, strKey)
            If strBuffer = "" Then Error 2
            frmMain.lstLocal.ColumnHeaders(4).Text = strBuffer
            If frmMain.lstLocal.ColumnHeaders(4).Text <> strBuffer Then Error 3
            
            frmMain.lstRemote.ColumnHeaders(4).Text = strBuffer
            If frmMain.lstRemote.ColumnHeaders(4).Text <> strBuffer Then Error 3
            
            'lstLocalRemote header 5
            strKey = "lstLocalRemote_5"
            strBuffer = LoadIni(strFile, strSection, strKey)
            If strBuffer = "" Then Error 2
            frmMain.lstLocal.ColumnHeaders(5).Text = strBuffer
            If frmMain.lstLocal.ColumnHeaders(5).Text <> strBuffer Then Error 3
            
            frmMain.lstRemote.ColumnHeaders(5).Text = strBuffer
            If frmMain.lstRemote.ColumnHeaders(5).Text <> strBuffer Then Error 3
            
            'lstLocalRemote header 6
            strKey = "lstLocalRemote_6"
            strBuffer = LoadIni(strFile, strSection, strKey)
            If strBuffer = "" Then Error 2
            frmMain.lstLocal.ColumnHeaders(6).Text = strBuffer
            If frmMain.lstLocal.ColumnHeaders(6).Text <> strBuffer Then Error 3
            
            frmMain.lstRemote.ColumnHeaders(6).Text = strBuffer
            If frmMain.lstRemote.ColumnHeaders(6).Text <> strBuffer Then Error 3
            
            'lblLocalArts label
            strKey = "lblLocalArts"
            strBuffer = LoadIni(strFile, strSection, strKey)
            If strBuffer = "" Then Error 2
            frmMain.lblLocalArts.Caption = strBuffer
            If frmMain.lblLocalArts.Caption <> strBuffer Then Error 3

            'cmdMkLocalArt button
            strKey = "cmdMkLocalArt"
            strBuffer = LoadIni(strFile, strSection, strKey)
            If strBuffer = "" Then Error 2
            frmMain.cmdMkLocalArt.Caption = strBuffer
            If frmMain.cmdMkLocalArt.Caption <> strBuffer Then Error 3
            
            'cmdCopyLocalArt button
            strKey = "cmdCopyLocalArt"
            strBuffer = LoadIni(strFile, strSection, strKey)
            If strBuffer = "" Then Error 2
            frmMain.cmdCopyLocalArt.Caption = strBuffer
            If frmMain.cmdCopyLocalArt.Caption <> strBuffer Then Error 3
            
            'cmdEditLocalArt button
            strKey = "cmdEditLocalArt"
            strBuffer = LoadIni(strFile, strSection, strKey)
            If strBuffer = "" Then Error 2
            frmMain.cmdEditLocalArt.Caption = strBuffer
            If frmMain.cmdEditLocalArt.Caption <> strBuffer Then Error 3
            
            'cmdDelLocalArt button
            strKey = "cmdDelLocalArt"
            strBuffer = LoadIni(strFile, strSection, strKey)
            If strBuffer = "" Then Error 2
            frmMain.cmdDelLocalArt.Caption = strBuffer
            If frmMain.cmdDelLocalArt.Caption <> strBuffer Then Error 3
            
            'cmdRenLocalArt button
            strKey = "cmdRenLocalArt"
            strBuffer = LoadIni(strFile, strSection, strKey)
            If strBuffer = "" Then Error 2
            frmMain.cmdRenLocalArt.Caption = strBuffer
            If frmMain.cmdRenLocalArt.Caption <> strBuffer Then Error 3
            
            'cmdUploadArt(0) button
            strKey = "cmdUploadArt_0"
            strBuffer = LoadIni(strFile, strSection, strKey)
            If strBuffer = "" Then Error 2
            frmMain.cmdUploadArt(0).Caption = strBuffer
            If frmMain.cmdUploadArt(0).Caption <> strBuffer Then Error 3
            
            'cmdUploadArt(1) button
            strKey = "cmdUploadArt_1"
            strBuffer = LoadIni(strFile, strSection, strKey)
            If strBuffer = "" Then Error 2
            frmMain.cmdUploadArt(1).Caption = strBuffer
            If frmMain.cmdUploadArt(1).Caption <> strBuffer Then Error 3
            
            'cmdDownloadArt(0) button
            strKey = "cmdDownloadArt_0"
            strBuffer = LoadIni(strFile, strSection, strKey)
            If strBuffer = "" Then Error 2
            frmMain.cmdDownloadArt(0).Caption = strBuffer
            If frmMain.cmdDownloadArt(0).Caption <> strBuffer Then Error 3
            
            'cmdDownloadArt(1) button
            strKey = "cmdDownloadArt_1"
            strBuffer = LoadIni(strFile, strSection, strKey)
            If strBuffer = "" Then Error 2
            frmMain.cmdDownloadArt(1).Caption = strBuffer
            If frmMain.cmdDownloadArt(1).Caption <> strBuffer Then Error 3
            
            'cmdBackupBlog button
            strKey = "cmdBackupBlog"
            strBuffer = LoadIni(strFile, strSection, strKey)
            If strBuffer = "" Then Error 2
            frmMain.cmdBackupBlog.Caption = strBuffer
            If frmMain.cmdBackupBlog.Caption <> strBuffer Then Error 3
            
            'cmdCleanupBlog button
            strKey = "cmdCleanupBlog"
            strBuffer = LoadIni(strFile, strSection, strKey)
            If strBuffer = "" Then Error 2
            frmMain.cmdCleanupBlog.Caption = strBuffer
            If frmMain.cmdCleanupBlog.Caption <> strBuffer Then Error 3
            
            'framRemote framework
            strKey = "framRemote"
            strBuffer = LoadIni(strFile, strSection, strKey)
            If strBuffer = "" Then Error 2
            frmMain.framRemote.Caption = strBuffer
            If frmMain.framRemote.Caption <> strBuffer Then Error 3
            
            'lblRemoteCats label
            strKey = "lblRemoteCats"
            strBuffer = LoadIni(strFile, strSection, strKey)
            If strBuffer = "" Then Error 2
            frmMain.lblRemoteCats.Caption = strBuffer
            If frmMain.lblRemoteCats.Caption <> strBuffer Then Error 3
            
            'cmdMkRemoteCat button
            strKey = "cmdMkRemoteCat"
            strBuffer = LoadIni(strFile, strSection, strKey)
            If strBuffer = "" Then Error 2
            frmMain.cmdMkRemoteCat.Caption = strBuffer
            If frmMain.cmdMkRemoteCat.Caption <> strBuffer Then Error 3
            
            'cmdRmRemoteCat button
            strKey = "cmdRmRemoteCat"
            strBuffer = LoadIni(strFile, strSection, strKey)
            If strBuffer = "" Then Error 2
            frmMain.cmdRmRemoteCat.Caption = strBuffer
            If frmMain.cmdRmRemoteCat.Caption <> strBuffer Then Error 3
            
            'cmdRenRemoteCat button
            strKey = "cmdRenRemoteCat"
            strBuffer = LoadIni(strFile, strSection, strKey)
            If strBuffer = "" Then Error 2
            frmMain.cmdRenRemoteCat.Caption = strBuffer
            If frmMain.cmdRenRemoteCat.Caption <> strBuffer Then Error 3
            
            'lblRemoteArts label
            strKey = "lblRemoteArts"
            strBuffer = LoadIni(strFile, strSection, strKey)
            If strBuffer = "" Then Error 2
            frmMain.lblRemoteArts.Caption = strBuffer
            If frmMain.lblRemoteArts.Caption <> strBuffer Then Error 3
            
            'cmdSeeRemoteArt button
            strKey = "cmdSeeRemoteArt"
            strBuffer = LoadIni(strFile, strSection, strKey)
            If strBuffer = "" Then Error 2
            frmMain.cmdSeeRemoteArt.Caption = strBuffer
            If frmMain.cmdSeeRemoteArt.Caption <> strBuffer Then Error 3
            
            'cmdDelRemoteArt button
            strKey = "cmdDelRemoteArt"
            strBuffer = LoadIni(strFile, strSection, strKey)
            If strBuffer = "" Then Error 2
            frmMain.cmdDelRemoteArt.Caption = strBuffer
            If frmMain.cmdDelRemoteArt.Caption <> strBuffer Then Error 3
            
            'cmdRenRemoteArt button
            strKey = "cmdRenRemoteArt"
            strBuffer = LoadIni(strFile, strSection, strKey)
            If strBuffer = "" Then Error 2
            frmMain.cmdRenRemoteArt.Caption = strBuffer
            If frmMain.cmdRenRemoteArt.Caption <> strBuffer Then Error 3
            
            'optStatusArt(0) option
            strKey = "optStatusArt_0"
            strBuffer = LoadIni(strFile, strSection, strKey)
            If strBuffer = "" Then Error 2
            frmMain.optStatusArt(0).Caption = strBuffer
            If frmMain.optStatusArt(0).Caption <> strBuffer Then Error 3
            
            'optStatusArt(1) option
            strKey = "optStatusArt_1"
            strBuffer = LoadIni(strFile, strSection, strKey)
            If strBuffer = "" Then Error 2
            frmMain.optStatusArt(1).Caption = strBuffer
            If frmMain.optStatusArt(1).Caption <> strBuffer Then Error 3
            
            'barStatus_Cnnx panel
            strKey = "barStatus_Cnnx"
            strBuffer = LoadIni(strFile, strSection, strKey)
            If strBuffer = "" Then Error 2
            frmMain.barStatus.Panels.item("cnnx").Text = strBuffer
            If frmMain.barStatus.Panels.item("cnnx").Text <> strBuffer Then Error 3
        
        Case "about"
            'The 'About Interface' section about frmAbout
            '***LATER : extend localization to all about-box elements
            strSection = "About Interface"
            
            'title-bar
            strKey = "Title"
            strBuffer = LoadIni(strFile, strSection, strKey)
            If strBuffer = "" Then Error 2
            frmAbout.Caption = strBuffer
            If frmAbout.Caption <> strBuffer Then Error 3
            
        Case "edit"
            strSection = "Edit Interface"
        
            'title-bar
            strKey = "Title"
            strBuffer = LoadIni(strFile, strSection, strKey)
            If strBuffer = "" Then Error 2
            frmEdit.Caption = strBuffer
            If frmEdit.Caption <> strBuffer Then Error 3
        
            'gallery label
            strKey = "lblGallery"
            strBuffer = LoadIni(strFile, strSection, strKey)
            If strBuffer = "" Then Error 2
            frmEdit.lblGallery.Caption = strBuffer
            If frmEdit.lblGallery.Caption <> strBuffer Then Error 3
        
            'cancel button
            strKey = "cmdCancel"
            strBuffer = LoadIni(strFile, strSection, strKey)
            If strBuffer = "" Then Error 2
            frmEdit.cmdCancel.Caption = strBuffer
            If frmEdit.cmdCancel.Caption <> strBuffer Then Error 3
        
        Case "progress"
            strSection = "Progress Interface"
            '(nothing to do until know, but could change)
        
        Case "project"
            strSection = "Project Interface"
            
            'lblTitle label
            strKey = "lblTitle"
            strBuffer = LoadIni(strFile, strSection, strKey)
            If strBuffer = "" Then Error 2
            frmProject.lblTitle.Caption = strBuffer
            If frmProject.lblTitle.Caption <> strBuffer Then Error 3
            
            'lblBlosxomURL label
            strKey = "lblBlosxomURL"
            strBuffer = LoadIni(strFile, strSection, strKey)
            If strBuffer = "" Then Error 2
            frmProject.lblBlosxomURL.Caption = strBuffer
            If frmProject.lblBlosxomURL.Caption <> strBuffer Then Error 3
            
            'framCnnx framework
            strKey = "framCnnx"
            strBuffer = LoadIni(strFile, strSection, strKey)
            If strBuffer = "" Then Error 2
            frmProject.framCnnx.Caption = strBuffer
            If frmProject.framCnnx.Caption <> strBuffer Then Error 3
            
            'lblHostPort label
            strKey = "lblHostPort"
            strBuffer = LoadIni(strFile, strSection, strKey)
            If strBuffer = "" Then Error 2
            frmProject.lblHostPort.Caption = strBuffer
            If frmProject.lblHostPort.Caption <> strBuffer Then Error 3
            
            'lblUser label
            strKey = "lblUser"
            strBuffer = LoadIni(strFile, strSection, strKey)
            If strBuffer = "" Then Error 2
            frmProject.lblUser.Caption = strBuffer
            If frmProject.lblUser.Caption <> strBuffer Then Error 3
            
            'lblPass label
            strKey = "lblPass"
            strBuffer = LoadIni(strFile, strSection, strKey)
            If strBuffer = "" Then Error 2
            frmProject.lblPass.Caption = strBuffer
            If frmProject.lblPass.Caption <> strBuffer Then Error 3
            
            'cmdShowPass button
            strKey = "cmdShowPass"
            strBuffer = LoadIni(strFile, strSection, strKey)
            If strBuffer = "" Then Error 2
            frmProject.cmdShowPass.Caption = strBuffer
            If frmProject.cmdShowPass.Caption <> strBuffer Then Error 3
            
            'framTree framework
            strKey = "framTree"
            strBuffer = LoadIni(strFile, strSection, strKey)
            If strBuffer = "" Then Error 2
            frmProject.framTree.Caption = strBuffer
            If frmProject.framTree.Caption <> strBuffer Then Error 3
            
            'lblArtRoot label
            strKey = "lblArtRoot"
            strBuffer = LoadIni(strFile, strSection, strKey)
            If strBuffer = "" Then Error 2
            frmProject.lblArtRoot.Caption = strBuffer
            If frmProject.lblArtRoot.Caption <> strBuffer Then Error 3
            
            'lblExcluded label
            strKey = "lblExcluded"
            strBuffer = LoadIni(strFile, strSection, strKey)
            If strBuffer = "" Then Error 2
            frmProject.lblExcluded.Caption = strBuffer
            If frmProject.lblExcluded.Caption <> strBuffer Then Error 3
            
            'lblArtExt label
            strKey = "lblArtExt"
            strBuffer = LoadIni(strFile, strSection, strKey)
            If strBuffer = "" Then Error 2
            frmProject.lblArtExt.Caption = strBuffer
            If frmProject.lblArtExt.Caption <> strBuffer Then Error 3
            
            'lblFlavExt label
            strKey = "lblFlavExt"
            strBuffer = LoadIni(strFile, strSection, strKey)
            If strBuffer = "" Then Error 2
            frmProject.lblFlavExt.Caption = strBuffer
            If frmProject.lblFlavExt.Caption <> strBuffer Then Error 3
            
            'framArt framework
            strKey = "framArt"
            strBuffer = LoadIni(strFile, strSection, strKey)
            If strBuffer = "" Then Error 2
            frmProject.framArt.Caption = strBuffer
            If frmProject.framArt.Caption <> strBuffer Then Error 3
            
            'lblImgMax label
            strKey = "lblImgMax"
            strBuffer = LoadIni(strFile, strSection, strKey)
            If strBuffer = "" Then Error 2
            frmProject.lblImgMax.Caption = strBuffer
            If frmProject.lblImgMax.Caption <> strBuffer Then Error 3
            
            'lblArtEncode label
            strKey = "lblArtEncode"
            strBuffer = LoadIni(strFile, strSection, strKey)
            If strBuffer = "" Then Error 2
            frmProject.lblArtEncode.Caption = strBuffer
            If frmProject.lblArtEncode.Caption <> strBuffer Then Error 3
                       
            'framStatus framework
            strKey = "framStatus"
            strBuffer = LoadIni(strFile, strSection, strKey)
            If strBuffer = "" Then Error 2
            frmProject.framStatus.Caption = strBuffer
            If frmProject.framStatus.Caption <> strBuffer Then Error 3
            
            'chkPrev option
            strKey = "chkPrev"
            strBuffer = LoadIni(strFile, strSection, strKey)
            If strBuffer = "" Then Error 2
            frmProject.chkPrev.Caption = strBuffer
            If frmProject.chkPrev.Caption <> strBuffer Then Error 3
            
            'lblPrevPrefix label
            strKey = "lblPrevPrefix"
            strBuffer = LoadIni(strFile, strSection, strKey)
            If strBuffer = "" Then Error 2
            frmProject.lblPrevPrefix.Caption = strBuffer
            If frmProject.lblPrevPrefix.Caption <> strBuffer Then Error 3
            
            'lblPrevPass label
            strKey = "lblPrevPass"
            strBuffer = LoadIni(strFile, strSection, strKey)
            If strBuffer = "" Then Error 2
            frmProject.lblPrevPass.Caption = strBuffer
            If frmProject.lblPrevPass.Caption <> strBuffer Then Error 3
            
            'lblPrevPlugin label
            strKey = "lblPrevPlugin"
            strBuffer = LoadIni(strFile, strSection, strKey)
            If strBuffer = "" Then Error 2
            frmProject.lblPrevPlugin.Caption = strBuffer
            If frmProject.lblPrevPlugin.Caption <> strBuffer Then Error 3
            
            'framCSS framework
            strKey = "framCSS"
            strBuffer = LoadIni(strFile, strSection, strKey)
            If strBuffer = "" Then Error 2
            frmProject.framCSS.Caption = strBuffer
            If frmProject.framCSS.Caption <> strBuffer Then Error 3
            
            'lblCSS(0) label
            strKey = "lblCSS_0"
            strBuffer = LoadIni(strFile, strSection, strKey)
            If strBuffer = "" Then Error 2
            frmProject.lblCSS(0).Caption = strBuffer
            If frmProject.lblCSS(0).Caption <> strBuffer Then Error 3
            
            'lblCSS(1) label
            strKey = "lblCSS_1"
            strBuffer = LoadIni(strFile, strSection, strKey)
            If strBuffer = "" Then Error 2
            frmProject.lblCSS(1).Caption = strBuffer
            If frmProject.lblCSS(1).Caption <> strBuffer Then Error 3
            
            'framChapo framework
            strKey = "framChapo"
            strBuffer = LoadIni(strFile, strSection, strKey)
            If strBuffer = "" Then Error 2
            frmProject.framChapo.Caption = strBuffer
            If frmProject.framChapo.Caption <> strBuffer Then Error 3
            
            'chkChapo option
            strKey = "chkChapo"
            strBuffer = LoadIni(strFile, strSection, strKey)
            If strBuffer = "" Then Error 2
            frmProject.chkChapo.Caption = strBuffer
            If frmProject.chkChapo.Caption <> strBuffer Then Error 3
            
            'lblChapo label
            strKey = "lblChapo"
            strBuffer = LoadIni(strFile, strSection, strKey)
            If strBuffer = "" Then Error 2
            frmProject.lblChapo.Caption = strBuffer
            If frmProject.lblChapo.Caption <> strBuffer Then Error 3
            
            'lblChapoPlugin label
            strKey = "lblChapoPlugin"
            strBuffer = LoadIni(strFile, strSection, strKey)
            If strBuffer = "" Then Error 2
            frmProject.lblChapoPlugin.Caption = strBuffer
            If frmProject.lblChapoPlugin.Caption <> strBuffer Then Error 3
            
            'cmdOK button
            strKey = "cmdOK"
            strBuffer = LoadIni(strFile, strSection, strKey)
            If strBuffer = "" Then Error 2
            frmProject.cmdOK.Caption = strBuffer
            If frmProject.cmdOK.Caption <> strBuffer Then Error 3
            
            'cmdCancel button
            strKey = "cmdCancel"
            strBuffer = LoadIni(strFile, strSection, strKey)
            If strBuffer = "" Then Error 2
            frmProject.cmdCancel.Caption = strBuffer
            If frmProject.cmdCancel.Caption <> strBuffer Then Error 3
        
        Case "select"
            strSection = "Select Interface"
        
            'title-bar
            strKey = "Title"
            strBuffer = LoadIni(strFile, strSection, strKey)
            If strBuffer = "" Then Error 2
            frmSelect.Caption = strBuffer
            If frmSelect.Caption <> strBuffer Then Error 3
            
            'cmdOK button
            strKey = "cmdOK"
            strBuffer = LoadIni(strFile, strSection, strKey)
            If strBuffer = "" Then Error 2
            frmSelect.cmdOK.Caption = strBuffer
            If frmSelect.cmdOK.Caption <> strBuffer Then Error 3
            
            'cmdCancel button
            strKey = "cmdCancel"
            strBuffer = LoadIni(strFile, strSection, strKey)
            If strBuffer = "" Then Error 2
            frmSelect.cmdCancel.Caption = strBuffer
            If frmSelect.cmdCancel.Caption <> strBuffer Then Error 3
        
        Case Else
            'wrong argIN has been passed
            Error 4
        End Select
        
    LoadLangFile = True
LoadLangFile_End:
    On Error GoTo 0
    Exit Function
    
LoadLangFile_Err:
    LoadLangFile = False
    
    Select Case Err
        Case 1
            'file absent
            strMsg = "The required language file is absent : '" & strFile & "'"
        Case 2
            'invalid section
            strMsg = "Invalid " & strSection & " section (" & strKey & ") in language file : '" & strFile & "'"
        Case 3
            'section loading failure
            strMsg = "Unable to load the " & strSection & " section from language file : '" & strFile & "'"
        Case 4
            'absent section
            strMsg = "Absent " & strSection & " section in language file : '" & strFile & "'"
        Case 20
            'invalid file
            strMsg = arMsg(2) & " : '" & strFile & "'"
        Case 30
            'loading failure
            strMsg = arMsg(3) & " : '" & strFile & "'"
        Case Else
            strMsg = Err.Description
    End Select
    
    MsgBox strMsg, vbExclamation
    Resume LoadLangFile_End
End Function

Function SetLang(strLang As String, strPart As String) As Boolean
    'Applies the given language (LoadLangFile wrapper) ; ".lng" file must exist for this language
    'IN : same as for LoadLangFile()
    'NB : if language cannot be applied, there are two cases :
    '     1) if it's during first load, appli will be terminated ASAP (on activate event) using a RightLoading=False
    '     2) if it's later, during app life (using menu) the language will be not applied and menu unchanged
    Dim strLangFile As String
    
    strLangFile = App.Path & "\" & LCase(strLang) & ".lng"
    
    'file loading
    If LoadLangFile(strLangFile, strPart) = False Then
        SetLang = False
        Exit Function
    End If
        
    'synchronize the menu
    If strLang = "EN" Then
        frmMain.mnuLang_EN.Checked = True
        frmMain.mnuLang_FR.Checked = False
    Else 'FR
        frmMain.mnuLang_EN.Checked = False
        frmMain.mnuLang_FR.Checked = True
    End If

    SetLang = True
End Function

Function GetLang() As String
    'Determines the current active language in the application
    'OUT : international code of current language (upper case)
    If frmMain.mnuLang_EN.Checked = True Then
        GetLang = "EN"
    ElseIf frmMain.mnuLang_FR.Checked = True Then
        GetLang = "FR"
    Else
        '(shouldn't arrive but provided to think to adapt this fct if I extend the lang list a day)
        GetLang = "(unknown)"
    End If
End Function
