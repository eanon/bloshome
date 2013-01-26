Attribute VB_Name = "basConf"
'BlosHome (c) FFh Lab / Eric Lequien, 2009-2013 - http://ffh-lab.com
'This module contains the mechanism to manage the configuration INI-file

Option Explicit

Dim strConfFile As String

Sub CheckMandatory()
    'Ensures minimal required environment for the application
    'NB : will be useful
    '     - on first app launching since setup doesn't create prerequisite subdirs
    '     - on subsequent launchings to avoid ugly error if user deletes an app's subdir
    Dim strPath As String
    
    'creates log directory if it doesn't yet exist
    strPath = App.Path & "\log"
    If Dir(strPath, vbDirectory) = "" Then MkDir (strPath)

    'creates data directory if it doesn't yet exist
    strPath = App.Path & "\data"
    If Dir(strPath, vbDirectory) = "" Then MkDir (strPath)
End Sub

Function LoadAllSettings() As Boolean
    'Loads the different settings from ini-file and defines the strConfFile global var
    'NB : we don't force any default value if a value is empty or absent in the INI file
    '***LATER : add checking/validation about values loaded from every section and key
    Dim strSection As String
    Dim strKey As String
    Dim strBuffer As String
    
    strConfFile = App.Path & "\" & App.EXEName & ".ini"

    If Dir(strConfFile) = "" Then
        'INI file is absent => we attempt to load the English language file by default
        If SetLang("EN", "main") = False Then
            LoadAllSettings = False
            Exit Function
        End If
    Else
        'INI file is present => we attempt to load config and language as defined in it
    
        'The 'Language' section ("EN/FR" only w/ "EN" by default)
        strSection = "UI"
        strKey = "Lang"
        strBuffer = LoadIni(strConfFile, strSection, strKey)
        If UCase(strBuffer) <> "EN" And UCase(strBuffer) <> "FR" Then
            strBuffer = "EN"
        End If
        If SetLang(strBuffer, "main") = False Then
            LoadAllSettings = False
            Exit Function
        End If
    End If
    
    LoadAllSettings = True
End Function

Sub SaveAllSettings()
    'Saves the current settings toward ini-file
    '(considering LoadAllSettings being called on Form/Load, strConfFile is already defined)
    Dim strSection As String
    Dim strKey As String
    
    Dim strBuffer As String
    
    'The 'Interface' section
    strSection = "UI"
    
        'Language (EN/FR only w/ EN by default)
        strKey = "Lang"
        If frmMain.mnuLang_FR.Checked Then
            strBuffer = "FR"
        Else
            strBuffer = "EN"
        End If
        SaveIni strConfFile, strSection, strKey, strBuffer
End Sub
