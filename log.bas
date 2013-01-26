Attribute VB_Name = "basLog"
'BlosHome (c) FFh Lab / Eric Lequien, 2009-2013 - http://ffh-lab.com
'This module contains the material to manage a log file

Option Explicit

Public Const LOGFILE_EXT = ".log"
Public Const LOGFILES_DIR = "log"

Private nLogID As Integer

Sub BeginLog()
    'Init a new log file for the session
    Dim strLog As String
    
    strLog = App.Path & "\" & LOGFILES_DIR & "\" & CheckAndMkFilename(Date, True, True) & "_" & CheckAndMkFilename(Time, True, True) & LOGFILE_EXT
    
    nLogID = FreeFile
    Open strLog For Append As #nLogID
    DoLog App.title & " " & App.Major & "." & App.Minor & " rev." & App.Revision & " EN/FR (c) " & App.LegalCopyright & " - " & App.CompanyName
    DoLog vbNewLine & "Begin of session"
End Sub

Sub DoLog(strInfo As String)
    'Add the given info to the current log file
    If bRightLoading = True Then Print #nLogID, strInfo
End Sub

Sub EndLog()
    'Ends the current log file
    DoLog "End of session"
    Close #nLogID
End Sub
