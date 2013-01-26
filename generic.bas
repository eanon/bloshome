Attribute VB_Name = "basGeneric"
'BlosHome (c) FFh Lab / Eric Lequien, 2009-2013 - http://ffh-lab.com
'This module contains functions which are eventually usable/portable toward other projects

Option Explicit

Function LoadIni(strFile As String, strSection As String, strKey As String) As String
    'Loads a given value from a key under a section inside an ini-file
    Dim strValue As String * 256

    GetPrivateProfileString strSection, strKey, vbNullString, strValue, Len(strValue), strFile
    LoadIni = Left$(strValue, InStr(1, strValue, Chr$(0)) - 1)
End Function

Sub SaveIni(strFile As String, strSection As String, strKey As String, strValue As String)
    'Saves a given value with a key under a section inside an ini-file
    'TRICK : pass vbNull or null as strValue to delete a key
    WritePrivateProfileString strSection, strKey, strValue, strFile
End Sub

Sub RemoveIniSection(strFile As String, strSection As String)
    'Deletesa given section in a given ini-file
    WritePrivateProfileSection strSection, vbNullString, strFile
End Sub

Function OpenMIME(strPath As String)
    'GOAL : try to open a given file with the Windows associated software (manage an hourglass during launching)
    'IN : path toward the file to try to open (can be be a local file, UNC, URL, LNK)
    'NB : the associated software is the one registered in the MIME base registry for the "open" verb
    'REF : this fct is an extended mix from the ProspEDI 1.0 ß one and the PerFORM 1.0 one, both by FFh Lab
    'USE : for example, passing an URL, system will try to reach-it using the default Web browser
    Dim hProg As Long
    On Error GoTo OpenMIME_Err

    Screen.MousePointer = vbHourglass
    
    If strPath = "" Then
        MsgBox arMsg(0), vbOKOnly
        GoTo OpenMIME_Exit
    End If
        
    hProg = ShellExecute(0, "open", strPath, "", "", 1)
    If hProg < 32 Then
        MsgBox arMsg(1) & " '" & strPath & "'", vbOKOnly
    End If

OpenMIME_Exit:
    Screen.MousePointer = vbDefault
    On Error GoTo 0
    Exit Function

OpenMIME_Err:
    MsgBox Err.Description
    Resume OpenMIME_Exit
End Function

Function CheckAndMkFilename(strIn As String, bConvOrReject As Boolean, bNoSpace As Boolean, _
                            Optional bConvPercent As Boolean = False, _
                            Optional bPath As Boolean = False) As String
    'Checks or adapts the given string to be usable as valid path/filename
    'IN :
    '- strIn is a string to check and/or transform to valid filename
    '- bConvOrReject indicates wether to convert the forbidden signs by acceptable ones
    '              or reject the entire strIn string at first invalid sign (i.e. returns "")
    '- bNoSpace indicates wether to tranform the spaces as underscore ("_")
    '- bConvPercent indicates wether to convert any %xx to its equivalent character or not
    '- bPath indicates that it is a path rather than a simple isolated filename
    '  (then ':' will be allowed at seconde position and the "\/" will be possible)
    'OUT : a valid path/filename
    '      or an empty string ("") if it failed to provide one (i.e. if bConvOrReject = false)
    'USE : strRet = CheckAndMkFilename(strName, false, false) to check validity of given name
    '     strNewName = CheckAndMkFilename(strName, true, [true|false]) to build a valid filename from given one
    'NB : work too for directory name ; which is, from Windows point of view, a file with a specific attribut
    'NB#2 : note taht all path and name will be returned with ANSI encoding
    'REF : adapted from MakeFileName() in PerFORM 1.0 (c) 2003 FFh Lab
    Dim strTmp As String
    Dim nStart As Integer
    Dim nIdx As Integer
    Dim nPos As Integer
    
    ReDim arMap(1 To 2, 1 To 9) As String
    
    Const DOTx2 = "[DOUBLEDOT]"
    
    strTmp = strIn
    
    'paths and filenames are managed in ANSI
    If IsUTF8(strTmp) = True Then strTmp = Decode_UTF8(strTmp)

    'converts the coded characters to %xx if required ; ***LATER : extend to others %xx
    'NB : we convert ahead to do resulting characters participate to subsequent fct treatment
    If bConvPercent = True Then
        strTmp = Replace(strTmp, "%20", " ", , , vbTextCompare)
    End If
    
    'protects a ":" in second position of a path
    If bPath = True And Mid$(strTmp, 2, 1) = ":" Then
        strTmp = Left$(strTmp, 1) & DOTx2 & Right$(strTmp, Len(strTmp) - 2)
    End If
    
    'forbidden characters in a filename : substitute strings (contening allowed characters only)
    '***NB : this table of conversion is modifiable as needed !
    arMap(1, 1) = "/": arMap(2, 1) = "-" '(these two firsts will be ignored if bPath is true)
    arMap(1, 2) = "\": arMap(2, 2) = "-"
    
    arMap(1, 3) = ":": arMap(2, 3) = "-"
    arMap(1, 4) = "*": arMap(2, 4) = "x"
    arMap(1, 5) = "<": arMap(2, 5) = "-"
    arMap(1, 6) = ">": arMap(2, 6) = "-"
    arMap(1, 7) = "|": arMap(2, 7) = "-"
    arMap(1, 8) = "?": arMap(2, 8) = "!"
    arMap(1, 9) = Chr(34): arMap(2, 9) = "'"
    
    If bNoSpace = True Then
        ReDim Preserve arMap(1 To 2, 1 To 10)
        arMap(1, 10) = " ": arMap(2, 10) = "_"
    End If
    
    'tune-up the analysis range
    If bPath = True Then
        nStart = 3 'slash & antislash autorisés ds un chemin
    Else
        nStart = 1
    End If
    
    'effective substitutions
    For nIdx = nStart To UBound(arMap, 2)
        nPos = 1
        Do While nPos <> 0
            nPos = InStr(nPos, strTmp, arMap(1, nIdx), vbTextCompare)
            If nPos <> 0 Then
                If bConvOrReject = True Or (bConvOrReject = False And nIdx = 10) Then
                    'substitution
                    strTmp = Left$(strTmp, nPos - 1) & arMap(2, nIdx) & Right$(strTmp, Len(strTmp) - nPos)
                Else
                    'reject entire string
                    strTmp = ""
                End If
            End If
        Loop
    Next
    
    'restores an eventual ":" at second position of an absolute path
    If bPath = True Then strTmp = Replace(strTmp, DOTx2, ":", , 1, vbTextCompare)
    
    CheckAndMkFilename = strTmp
End Function

Public Function UnAccent(ByRef strIn As String) As String
    'Strip-out all eventual accents in the given string
    'IN : string to analize
    'OUT : resulting unaccented string
    Dim nIdx As Integer
    Dim strTmp As String
    Dim strChar As String * 1
    
    Dim strAccents As String
    Dim strEquival As String
    
    strAccents = "ÀÁÂÃÄÅàáâãäåÒÓÔÕÖØòóôõöøÈÉÊËèéêëÌÍÎÏìíîïÙÚÛÜùúûüÿÑñÇç"
    strEquival = "AAAAAAaaaaaaOOOOOOooooooEEEEeeeeIIIIiiiiUUUUuuuuyNnCc"
    
    strTmp = strIn
    For nIdx = 1 To Len(strAccents)
        strChar = Mid(strAccents, nIdx, 1)
        If InStr(strTmp, strChar) > 0 Then
            strTmp = Replace(strTmp, strChar, Mid$(strEquival, nIdx, 1))
        End If
    Next
    
    UnAccent = strTmp
End Function

Public Function FormatRGBString(nVal As Long) As String
    'Formats a long integer coming from CommonDialog color dialog to an hexadecimal value string as "#RRGGBB"
    Dim strColor, strR, strG, strB As String
    Dim nPad As Integer
        
    strColor = Hex(nVal) 'converts to hex
    
    'determines how many zeros to pad in front of converted hex value
    nPad = 6 - Len(strColor)
    If nPad > 0 Then strColor = String(nPad, "0") & strColor
        
    'extracts the RGB parts
    strR = Right$(strColor, 2)
    strG = Mid$(strColor, 3, 2)
    strB = Left$(strColor, 2)
    
    'swaps R and B, color dialog returning BGR instead of RGB
    FormatRGBString = "#" & strR & strG & strB
End Function

Function SaveText(strFile As String, strData As String, _
                  Optional bAppend As Boolean = False) As Boolean
    'Records strData the Ascii file strFile (return false if something wrong)
    'NB : depending of bAppend, strData will overwrite eventual existing strFile OR will be added inside it
    Dim nHandle As Integer
    
    On Error GoTo SaveText_Error
    nHandle = FreeFile
    
    If bAppend = True Then
        Open strFile For Append As #nHandle
    Else
        Open strFile For Output As #nHandle
    End If
    
    Print #nHandle, strData;
    Close #nHandle
    SaveText = True
    Exit Function

SaveText_Error:
    SaveText = False
    On Error GoTo 0
End Function

Function LoadText(strFile As String) As String
    'Returns the content of a given Ascii file
    'Warning : entire file content being loaded in memory, take care to have RAM enough
    Dim nHandle As Integer
    Dim strLine As String
    Dim strBuffer As String
    
    On Error GoTo LoadText_Error
    nHandle = FreeFile
    
    Open strFile For Input Access Read Shared As #nHandle
    Do While Not EOF(nHandle)
        Line Input #nHandle, strLine
        strBuffer = strBuffer & strLine & vbNewLine 'le dernier retour sera supprimé
    Loop
    Close #nHandle
    
    'preserves an eventual existing CRLF at the end of the file
    If Len(strBuffer) > FileLen(strFile) Then
        LoadText = Left$(strBuffer, Len(strBuffer) - Len(vbNewLine))
    Else
        LoadText = strBuffer
    End If
    
    Exit Function

LoadText_Error:
    LoadText = ""
    On Error GoTo 0
End Function

Sub ArrangeINISections(strIni As String)
    'Takes care to add an empty ligne between every sections of a given INI file
    'Warning : not any checking of file type is done ; caller is reponsible to pass a path to a valid INI one
    Dim strData As String
    
    strData = LoadText(strIni)
    
    '(cancellation then regeneration of intersections prevents accumulation from call to call)
    '***LATER : current treatment can corrects 1 or 2 existing extra empty lines before "[", but it would be
    '           better to cover any case (ie. n number of empty lines) through a do/while loop
    strData = Replace(strData, vbNewLine & vbNewLine & vbNewLine & "[", vbNewLine & "[", vbTextCompare)
    strData = Replace(strData, vbNewLine & vbNewLine & "[", vbNewLine & "[", vbTextCompare)
    strData = Replace(strData, vbNewLine & "[", vbNewLine & vbNewLine & "[", vbTextCompare)
   
    SaveText strIni, strData
End Sub

Sub GetImgDims(oFrm As Form, strPath As String, ByRef nWidth As Integer, _
                ByRef nHeight As Integer, Optional nScale As Integer = vbPixels)
    'Determine the dimensions of an image in a given unit
    'IN :
    ' - oFrm : caller form (to compute in the given unit)
    ' - strPath : path to the image file (bmp, cur, ico, rle, wmf, emf, gif, jpeg)
    ' - nWidth&, nHeight& : addresses of variables which will receive checked dimensions
    ' - [nScale] : unit used to inform nWidth & nHeight (e.g. vbTwips ; see ScaleX)
    'OUT : value are stored in the given variables nWidth & nHeight passed by reference
    Dim oPic As New StdPicture
    
    Set oPic = LoadPicture(strPath)
    nWidth = oFrm.ScaleX(oPic.Width, vbHimetric, nScale)
    nHeight = oFrm.ScaleY(oPic.Height, vbHimetric, nScale)
    Set oPic = Nothing
End Sub

Sub CalcThumbDims(nMaxThSide As Integer, nOrgWidth As Integer, nOrgHeight As Integer, _
                    ByRef nThWidth As Integer, ByRef nThHeight As Integer)
    'Determine the dimensions of a thumbnail, for the purpose to avoid distortion
    'IN :
    ' - nMaxThSide : maximal size of a thumbnail side (in pixels)
    ' - nOrgWidth, nOrgHeight : original dimensions of the image
    ' - nThWidth&, nThHeight& : addresses of variables which will receive the computed thumbnail dimensions
    'REF : adapted from Perl code written for GOLB 1.0 (c) FFh Lab
    Dim nCoefReduc As Integer

    If (nOrgWidth > nMaxThSide) And (nOrgHeight <= nMaxThSide) Then
        'image is just too large for thumbnail
        nCoefReduc = nOrgWidth / nMaxThSide
        nThWidth = Int(nOrgWidth / nCoefReduc)
        nThHeight = Int(nOrgHeight / nCoefReduc)
    ElseIf (nOrgWidth <= nMaxThSide) And (nOrgHeight > nMaxThSide) Then
        'image is just too high for thumbnail
        nCoefReduc = nOrgHeight / nMaxThSide
        nThWidth = Int(nOrgWidth / nCoefReduc)
        nThHeight = Int(nOrgHeight / nCoefReduc)
    ElseIf (nOrgWidth > nMaxThSide) And (nOrgHeight > nMaxThSide) Then
        'image is too large and too high for thumbnail
        '(we compute reduction from the farest value)
        If nOrgWidth > nOrgHeight Then
            nCoefReduc = nOrgWidth / nMaxThSide
            nThWidth = Int(nOrgWidth / nCoefReduc)
            nThHeight = Int(nOrgHeight / nCoefReduc)
        Else
            nCoefReduc = nOrgHeight / nMaxThSide
            nThWidth = Int(nOrgWidth / nCoefReduc)
            nThHeight = Int(nOrgHeight / nCoefReduc)
        End If
    Else '(nOrgWidth <= nMaxThSide) And (nOrgHeight <= nMaxThSide)
         'image can be directly a thumbnail w/ simple reduction
        nThWidth = nOrgWidth
        nThHeight = nOrgHeight
    End If
End Sub

Function RevInStr(str1 As String, str2 As String, nSensitivity As Integer) As Integer
    'GOAL : searches the position of the last occurrence of string2 in string1
    'IN : the two strings and a flag saying if we treat case sensitively or not ; same values as in InStr()
    'OUT : return the position in string1 ; zero if string2 not found
    'NB : Reverse of VB function InStr()
    'REF : comes from ProsEDI 1.0 (c) FFh Lab
    Dim nPos As Integer
    Dim nPrevPos As Integer
    
    nPos = 0
    Do
        nPrevPos = nPos
        nPos = InStr(nPos + 1, str1, str2, nSensitivity)
    Loop While nPos <> 0
    
    RevInStr = nPrevPos
End Function

Function MsgBoxEx(strPrompt As String, Optional vbButtons As VbMsgBoxStyle = vbOKOnly, _
                    Optional strTitle As String = "BlosHome", _
                    Optional strHelpFile As String = "", _
                    Optional nContext As Integer = 0) As VbMsgBoxResult
    'Extended MsgBox(), forcing standard cursor during display and restoring previous one after
    'IN & OUT : same as the MsgBox() ones
    '           (default strTitle has to be adapted according to current project ; here "BlosHome")
    'USE : useful to provide an arrow on entire UI area rather than on msg-box area only (MsgBox() behavior),
    '      when we have to display a message during framework of a process w/ global or app-related hourglass
    Dim nOldCursor As Integer
    Dim vbRet As VbMsgBoxResult
    
    nOldCursor = Screen.MousePointer
    Screen.MousePointer = vbDefault
    
    vbRet = MsgBox(strPrompt, vbButtons, strTitle, strHelpFile, nContext)
    
    Screen.MousePointer = nOldCursor
    MsgBoxEx = vbRet
End Function

Function CountOccurr(strText As String, strSearched As String) As Integer
    'Returns the number of occurrences of strSearched in strText text
    Dim nPos As Integer
    Dim nCount As Integer
    
    Do
        nPos = InStr(nPos + 1, strText, strSearched, vbTextCompare)
        If nPos > 0 Then nCount = nCount + 1
    Loop While nPos > 0
    
    CountOccurr = nCount
End Function

Sub StartStopTimer(oTimer As Timer, bOp As Boolean, Optional nInterval As Long, _
                    Optional ByRef nCounterToReset As Long)
    'Starts or stops the given timer with given interval
    'IN : the timer to manage, the sense of the operation, an interval in ms if bOp=True
    '     an optional given variable to reset if we maintain a pass counter somewhere
    'NB : ensures interval countdown will always restart from 0 at every timer (re)start
    'NB#2 : if we have to manage a counter (see nCounterToReset), remember that the timer
    '       event will occurs first when timer is enabled (not any interval being spent) ;
    '       so, the number of intervals spent will always be (nCounter - 1) times
    If bOp = True Then
        If nCounterToReset <> Empty Then nCounterToReset = 0
        oTimer.Interval = nInterval
        oTimer.Enabled = True
    Else
        oTimer.Interval = 0
        oTimer.Enabled = False
    End If
End Sub

Public Function IsLoaded(ByVal strForm As String) As Boolean
    'Indicate if a form with the given name is currently loaded (not necessarily shown, but loaded in memory)
    Dim frm As Form

    IsLoaded = False
    For Each frm In Forms
        If frm.Name = strForm Then
            IsLoaded = True
            Exit For
        End If
    Next frm
End Function

Public Function IsUTF8(strData As String) As Boolean
    'Determines if a given string is actually UTF-8 encoded
    '  Char. number range (hexadecimal) | UTF-8 bytes sequence (binary)
    '  ---------------------------------+------------------------------------
    '  0000 0000-0000 007F              | 0xxxxxxx
    '  0000 0080-0000 07FF              | 110xxxxx 10xxxxxx
    '  0000 0800-0000 FFFF              | 1110xxxx 10xxxxxx 10xxxxxx
    '  0001 0000-0010 FFFF              | 11110xxx 10xxxxxx 10xxxxxx 10xxxxxx
    'REF : cyberpat92 @ forum.hardware.fr/hfr/Programmation/VB-VBA-VBS/code-conversion-ansi-sujet_79551_1.htm
    Dim c0, c1, c2, c3, n
     
    IsUTF8 = True
    n = 1
    Do While n <= Len(strData)
        c0 = Asc(Mid$(strData, n, 1))
        
        If n <= Len(strData) - 1 Then
            c1 = Asc(Mid$(strData, n + 1, 1))
        Else
            c1 = 0
        End If
        
        If n <= Len(strData) - 2 Then
            c2 = Asc(Mid$(strData, n + 2, 1))
        Else
            c2 = 0
        End If
        
        If n <= Len(strData) - 3 Then
            c3 = Asc(Mid$(strData, n + 3, 1))
        Else
            c3 = 0
        End If
         
        If (c0 And 240) = 240 Then
            If (c1 And 128) = 128 And (c2 And 128) = 128 And (c3 And 128) = 128 Then
                n = n + 4
            Else
                IsUTF8 = False
                Exit Function
            End If
        ElseIf (c0 And 224) = 224 Then
            If (c1 And 128) = 128 And (c2 And 128) = 128 Then
                n = n + 3
            Else
                IsUTF8 = False
                Exit Function
            End If
        ElseIf (c0 And 192) = 192 Then
            If (c1 And 128) = 128 Then
                n = n + 2
            Else
                IsUTF8 = False
                Exit Function
            End If
        ElseIf (c0 And 128) = 0 Then
            n = n + 1
        Else
            IsUTF8 = False
            Exit Function
        End If
    Loop
End Function

Public Function Encode_UTF8(astr As String) As String
    'Encodes a given ANSI string to UTF-8 encoding
    '   Char. number range (hexadecimal) | UTF-8 octet sequence (binary)
    '   ---------------------------------+------------------------------------
    '   0000 0000-0000 007F              | 0xxxxxxx
    '   0000 0080-0000 07FF              | 110xxxxx 10xxxxxx
    '   0000 0800-0000 FFFF              | 1110xxxx 10xxxxxx 10xxxxxx
    '   0001 0000-0010 FFFF              | 11110xxx 10xxxxxx 10xxxxxx 10xxxxxx
    'REF : adapted from cyberpat92 code
    '      @ forum.hardware.fr/hfr/Programmation/VB-VBA-VBS/code-conversion-ansi-sujet_79551_1.htm
    Dim c
    Dim n
    Dim utftext
     
    utftext = ""
    n = 1
    Do While n <= Len(astr)
        c = AscW(Mid(astr, n, 1))
        If c < 0 Then
            'we explicitely ignore the impossible characters (added in 1.0.4 ; fixes bug giving error 5)
            utftext = utftext '(does nothing)
        ElseIf c < 128 Then
            utftext = utftext + Chr(c)
        ElseIf ((c >= 128) And (c < 2048)) Then
            utftext = utftext + Chr(((c \ 64) Or 192))
            utftext = utftext + Chr(((c And 63) Or 128))
        ElseIf ((c >= 2048) And (c < 65536)) Then
            utftext = utftext + Chr(((c \ 4096) Or 224))
            utftext = utftext + Chr((((c \ 64) And 63) Or 128))
            utftext = utftext + Chr(((c And 63) Or 128))
        Else ' c >= 65536
            utftext = utftext + Chr(((c \ 262144) Or 240))
            utftext = utftext + Chr(((((c \ 4096) And 63)) Or 128))
            utftext = utftext + Chr((((c \ 64) And 63) Or 128))
            utftext = utftext + Chr(((c And 63) Or 128))
        End If
        n = n + 1
    Loop
    Encode_UTF8 = utftext
End Function
 
Public Function Decode_UTF8(astr As String) As String
    'Decodes a given UTF-8 string to ANSI encoding
    '   Char. number range (hexadecimal)  | UTF-8 octet sequence (binary)
    '   ----------------------------------+------------------------------------
    '   0000 0000-0000 007F               | 0xxxxxxx
    '   0000 0080-0000 07FF               | 110xxxxx 10xxxxxx
    '   0000 0800-0000 FFFF               | 1110xxxx 10xxxxxx 10xxxxxx
    '   0001 0000-0010 FFFF               | 11110xxx 10xxxxxx 10xxxxxx 10xxxxxx
    'REF : cyberpat92 @ forum.hardware.fr/hfr/Programmation/VB-VBA-VBS/code-conversion-ansi-sujet_79551_1.htm
    Dim c0, c1, c2, c3
    Dim n
    Dim unitext
     
    If IsUTF8(astr) = False Then
        Decode_UTF8 = astr
        Exit Function
    End If
     
    unitext = ""
    n = 1
    Do While n <= Len(astr)
        c0 = Asc(Mid(astr, n, 1))
        If n <= Len(astr) - 1 Then
            c1 = Asc(Mid(astr, n + 1, 1))
        Else
            c1 = 0
        End If
        If n <= Len(astr) - 2 Then
            c2 = Asc(Mid(astr, n + 2, 1))
        Else
            c2 = 0
        End If
        If n <= Len(astr) - 3 Then
            c3 = Asc(Mid(astr, n + 3, 1))
        Else
            c3 = 0
        End If
         
        If (c0 And 240) = 240 And (c1 And 128) = 128 And (c2 And 128) = 128 And (c3 And 128) = 128 Then
            unitext = unitext + ChrW((c0 - 240) * 65536 + (c1 - 128) * 4096) + (c2 - 128) * 64 + (c3 - 128)
            n = n + 4
        ElseIf (c0 And 224) = 224 And (c1 And 128) = 128 And (c2 And 128) = 128 Then
            unitext = unitext + ChrW((c0 - 224) * 4096 + (c1 - 128) * 64 + (c2 - 128))
            n = n + 3
        ElseIf (c0 And 192) = 192 And (c1 And 128) = 128 Then
            unitext = unitext + ChrW((c0 - 192) * 64 + (c1 - 128))
            n = n + 2
        ElseIf (c0 And 128) = 128 Then
            unitext = unitext + ChrW(c0 And 127)
            n = n + 1
        Else ' c0 < 128
            unitext = unitext + ChrW(c0)
            n = n + 1
        End If
    Loop
 
    Decode_UTF8 = unitext
End Function

Public Sub SetIcon(ByVal hwnd As Long, ByVal sIconResName As String, _
                                Optional ByVal bSetAsAppIcon As Boolean = True)
    'Associates an icon resource to a given window and eventually its entire application
    'NB : using largely the Win32 API, it allows to override the default VB6 behavior which
    '     doesn't allow icon with more than 256 colors (16 colors for the desktop one) and 32x32 pixels
    '     (so, here, with this fct, we can bind any multi-resolutions icon using any colors depth and size)
    'REF : http://www.vbaccelerator.com/home/vb/tips/setting_the_app_icon_correctly/article.asp
    'WARNING : without effect when app. is launched from within the IDE ; ie. it works in compiled EXE only
    Dim lhWndTop As Long
    Dim lhWnd As Long
    Dim cx As Long
    Dim cy As Long
    Dim hIconLarge As Long
    Dim hIconSmall As Long
      
    If (bSetAsAppIcon) Then
       ' Find VB's hidden parent window:
       lhWnd = hwnd
       lhWndTop = lhWnd
       Do While Not (lhWnd = 0)
          lhWnd = GetWindow(lhWnd, GW_OWNER)
          If Not (lhWnd = 0) Then
             lhWndTop = lhWnd
          End If
       Loop
    End If
    
    cx = GetSystemMetrics(SM_CXICON)
    cy = GetSystemMetrics(SM_CYICON)
    hIconLarge = LoadImageAsString(App.hInstance, sIconResName, IMAGE_ICON, cx, cy, LR_SHARED)
    If (bSetAsAppIcon) Then
       SendMessageLong lhWndTop, WM_SETICON, ICON_BIG, hIconLarge
    End If
    SendMessageLong hwnd, WM_SETICON, ICON_BIG, hIconLarge
    
    cx = GetSystemMetrics(SM_CXSMICON)
    cy = GetSystemMetrics(SM_CYSMICON)
    hIconSmall = LoadImageAsString(App.hInstance, sIconResName, IMAGE_ICON, cx, cy, LR_SHARED)
    If (bSetAsAppIcon) Then
       SendMessageLong lhWndTop, WM_SETICON, ICON_SMALL, hIconSmall
    End If
    SendMessageLong hwnd, WM_SETICON, ICON_SMALL, hIconSmall
End Sub

Function GetFilenamePrefix(strFilename As String) As String
    'Returns the part before extension of a filename
    'IN : a filename with or without path
    'OUT : the part before extension and after eventual path
    'USE : GetFilenamePrefix("C:/data/sub/something_here.txt") will return "something_here"
    '      GetFilenamePrefix("C:/data/sub/something_here.dat.txt") will return "something_here.dat"
    '      GetFilenamePrefix("C:/data/sub/something_here") will return "something_here"
    '      GetFilenamePrefix("..\something_here.html") will return "something_here"
    '***LATER : gather GetFilenamePrefix() and GetFilename() in a GetFilenamePart() with a parameter
    '           indicating the part to return : path, filename, prefix or extension
    Dim strPrefix As String
    Dim arPathSign(1 To 2) As String
    Dim nIdx As Integer
    Dim nPos As Integer
    
    arPathSign(1) = "/"
    arPathSign(2) = "\"
    
    strPrefix = strFilename
    
    For nIdx = 1 To 2
        nPos = RevInStr(strPrefix, arPathSign(nIdx), False)
        If nPos <> 0 Then
            strPrefix = Right$(strPrefix, Len(strPrefix) - nPos)
        End If
    Next

    nPos = RevInStr(strPrefix, ".", False)
    If nPos <> 0 Then
        strPrefix = Left$(strPrefix, nPos - 1)
    End If
    
    GetFilenamePrefix = strPrefix
End Function

Function GetTmpFile(strPrefix As String) As String
    'Returns a valid temporary path and filename according to Windows standard
    'IN : filename prefix to use (eases visual recognition of the ownership during debug)
    '     (e.g. "BH" to designate that it's about a temporary file belonging to BlosHome)
    'OUT : valid full path OR "" in case of failure
    'WARNING : this function doesn't create any file, but just returns a usable string
    Dim strTmpPath As String
    Dim strTmpFile As String
    Dim nTmpPathLen As Long
    
    GetTmpFile = ""
    
    strTmpPath = String$(MAX_PATH, 0)
    strTmpFile = String$(MAX_PATH, 0)
    
    nTmpPathLen = GetTempPath(Len(strTmpPath), strTmpPath)
    If nTmpPathLen = 0 Then Exit Function
    strTmpPath = Left$(strTmpPath, nTmpPathLen)
    GetTempFileName strTmpPath, strPrefix, 0, strTmpFile
    
    GetTmpFile = strTmpFile
End Function

Function GetTmpPath(Optional bWithFinalASlash = False) As String
    'Returns a valid temporary path according to Windows standard
    'IN : a flag indicating if we want a final "\" or not (of course, ignored for the case of root)
    'OUT : resulting path or "" in case of failure
    'NB : a final "/" will be considered as "\", both syntax being Windows compatible
    'WARNING : this function doesn't create any directory, but returns a usable string
    Dim strTmpPath As String
    Dim nTmpPathLen As Long
    
    Dim arSepar(1 To 2) As String
    Dim strFinalChar As String
    
    arSepar(1) = "\"
    arSepar(2) = "/"
    
    GetTmpPath = ""
    
    strTmpPath = String$(MAX_PATH, 0)
    nTmpPathLen = GetTempPath(Len(strTmpPath), strTmpPath)
    If nTmpPathLen = 0 Then Exit Function
    strTmpPath = Left$(strTmpPath, nTmpPathLen)
    
    If Len(strTmpPath) > 3 Then
        strFinalChar = Right$(strTmpPath, 1)
    
        If bWithFinalASlash = True Then
            If strFinalChar <> arSepar(1) And strFinalChar <> arSepar(2) Then
                strTmpPath = strTmpPath & "\"
            End If
        Else
            If strFinalChar = arSepar(1) Or strFinalChar = arSepar(2) Then
                strTmpPath = Left$(strTmpPath, Len(strTmpPath) - 1)
            End If
        End If
    End If
    
    GetTmpPath = strTmpPath
End Function

Function GetFileName(strFile As String) As String
    'Returns the filename part of given path (absolute or relative one, doesn't matter)
    'IN : a filename with path ahead (work also if there's not any path ; in this case, the fct does nothing)
    'OUT : filename alone (without path ahead)
    'USE : GetFilename("C:/data/sub/something_here.txt") will return "something_here.txt"
    '      GetFilename("C:/data/sub/something_here.dat.txt") will return "something_here.dat.txt"
    '      GetFilename("C:/data/sub/something_here") will return "something_here"
    '      GetFilename("..\something_here.html") will return "something_here.html"
    '      GetFilename("something_here.html") will return "something_here.html"
    Dim strFilename As String
    Dim arPathSign(1 To 2) As String
    Dim nIdx As Integer
    Dim nPos As Integer
    
    arPathSign(1) = "/"
    arPathSign(2) = "\"
    
    strFilename = strFile
    
    For nIdx = 1 To 2
        nPos = RevInStr(strFilename, arPathSign(nIdx), False)
        If nPos <> 0 Then
            strFilename = Right$(strFilename, Len(strFilename) - nPos)
        End If
    Next

    GetFileName = strFilename
End Function

Function RevStrArray(arIN() As String) As String()
    'Returns a reversed version of a given strings array
    'IN: a strings array
    'OUT: resulting reversed array
    '***LATER : expands to others type of arrays (analyze type via TypeOf OR going with generic variant type)
    '           allows multi-dimensional arrays
    Dim nMin As Long
    Dim nMax As Long
    Dim arTmp() As String
    
    Dim nIdx As Long
    Dim nIdx2 As Long
    
    nMin = LBound(arIN)
    nMax = UBound(arIN)
    
    ReDim arTmp(nMin To nMax)
    nIdx2 = nMin

    For nIdx = nMax To nMin Step -1
        arTmp(nIdx2) = arIN(nIdx)
        nIdx2 = nIdx2 + 1
    Next nIdx
    
    RevStrArray = arTmp()
End Function

Function SetOptionBase(arIN() As String, nBase As Integer) As String()
    'Forces the option base of a strings array
    'IN : the array to fix and the new option base to apply
    '***LATER : expands to others type of arrays (analyze type via TypeOf OR going with generic variant type)
    '           allows multi-dimensional arrays
    Dim nIdx As Long
    Dim nIdx2 As Long
    Dim nElts As Long
    Dim arOUT() As String
    
    SetOptionBase = arIN()
    If LBound(arIN) = nBase Then Exit Function 'retourne un tableau inchangé
    
    nElts = UBound(arIN) - LBound(arIN) + 1
    ReDim arOUT(nBase To nBase + nElts - 1)
    
    nIdx2 = nBase
    
    For nIdx = LBound(arIN) To UBound(arIN)
        arOUT(nIdx2) = arIN(nIdx)
        nIdx2 = nIdx2 + 1
    Next
    
    SetOptionBase = arOUT()
End Function
