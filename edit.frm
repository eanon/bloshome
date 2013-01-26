VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Article"
   ClientHeight    =   10830
   ClientLeft      =   360
   ClientTop       =   645
   ClientWidth     =   13050
   ControlBox      =   0   'False
   LinkTopic       =   "edit"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10830
   ScaleWidth      =   13050
   ShowInTaskbar   =   0   'False
   Begin SHDocVwCtl.WebBrowser wbGallery 
      CausesValidation=   0   'False
      Height          =   6705
      Left            =   10275
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   3405
      Width           =   2625
      ExtentX         =   4630
      ExtentY         =   11827
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   0
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.DirListBox lstDir 
      Height          =   2565
      Left            =   10260
      TabIndex        =   3
      Top             =   735
      Width           =   2670
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Quitter sans enregistrer"
      Height          =   375
      Left            =   10275
      TabIndex        =   5
      Top             =   10230
      Width           =   2670
   End
   Begin SHDocVwCtl.WebBrowser wbEditor 
      CausesValidation=   0   'False
      Height          =   10665
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   30
      Width           =   10140
      ExtentX         =   17886
      ExtentY         =   18812
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   0
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.DriveListBox lstDrive 
      Height          =   315
      Left            =   10260
      TabIndex        =   2
      Top             =   345
      Width           =   2670
   End
   Begin VB.Label lblGallery 
      AutoSize        =   -1  'True
      Caption         =   "Gallery :"
      Height          =   195
      Left            =   10260
      TabIndex        =   1
      Top             =   45
      Width           =   570
   End
End
Attribute VB_Name = "frmEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'BlosHome (c) FFh Lab / Eric Lequien, 2009-2013 - http://ffh-lab.com
'This frame handles all required (including an embedded TinyMCE editor) to edit blog article

Option Explicit
Option Base 1

Private bInitializedEditor As Boolean  'init of wbEditor during load, then activates if failed on load
Private bInitializedGallery As Boolean 'init of wbGallery during load, then activates if failed on load

Private Sub cmdCancel_Click()
    'Simply hides the editor (remains in RAM to avoid slow javascript init each time we need-it)
    Me.Hide
    ResetEditor 'prepares-it for next call
End Sub

Private Sub Form_Activate()
    'Initializes UI
    'Algo : wb* initialized here if failure on load, then next attempt on paint if failure here because of err #91
    Static bFirstCallDone As Boolean
    
    Screen.MousePointer = vbHourglass
    
    If bFirstCallDone = False Then
        'first-time init only
        If bInitializedEditor = False Then
            wbEditor.Navigate "about:blank"
            bInitializedEditor = ResetEditor
        End If
        
        'init of lstDrive -> lstDir and wbGallery through InitGallery
        If bInitializedGallery = False Then
            wbGallery.Navigate "about:blank"
            InitGallery 'bInitializedGallery will be set on lstDir_Change selon PopulateGallery
        End If
        
        bFirstCallDone = True
    End If
       
    EditCurrArt
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    'Initializes the interface (including important properties for look n' feel)
    'NB : frmEdit is pre-loaded with frmMain, then just shown/hidden when needed (at every edit request)
    With wbEditor
        .MenuBar = False
        .AddressBar = False
        .Toolbar = False
        .StatusBar = True   'allows triggering of statusbar_change which is used for saving
        .Offline = True
        .Silent = True      'avoids javascript error if TinyMCE is undefined (because of modified tree and/or files)
    End With
    
    With wbGallery
        .MenuBar = False
        .AddressBar = False
        .Toolbar = False
        .StatusBar = False
        .Offline = True
        .Silent = True
    End With
    
    'try to init the wb* ; next on activate and paint if fails here
    bInitializedEditor = ResetEditor
    InitGallery 'bInitializedGallery will be set on subsequent lstDir_Change according to PopulateGallery
End Sub

Sub DoShow(oOwner As Form)
    'Adjusts the language, then shows the frame
    'NB#1 : if not any valid language is currently selected, we'll keep design-time state ; i.e. English
    'NB#2 : frmEdit remaining loaded in memory during all BlosHome session, this fct has to be called
    '       everytime frmEdit is invoked (i.e. shown) ; concretly from frmMain/cmd[MkLocalArt|EditLocalArt]
    Dim strLang As String
    
    strLang = GetLang()
    If strLang <> "(unknown)" Then SetLang strLang, "edit"
    
    'Some UI-string adjustements afterward
    If IsNumeric(structProj.img_max) = True Then
        lblGallery = lblGallery & " (max. " & structProj.img_max & " KB) :"
    Else
        lblGallery = lblGallery + " :"
    End If
    
    Show 1, oOwner
End Sub

Private Sub Form_Paint()
    'Engages init of wbGallery on first call if not succeeded during form load and activate
    'NB : done here too since wbGallery may be not ready on load/activate & may generate error #91 (object not set)
    Static bFirstCallDone As Boolean
    
    If bFirstCallDone = False Then
        'quick init of wbEditor on first call, prior to long one with EditCurrArt
        If bInitializedEditor = False Then
            wbEditor.Navigate "about:blank"
            bInitializedEditor = ResetEditor
        End If
        
        'init of lstDrive -> lstDir and wbGallery through InitGallery
        If bInitializedGallery = False Then
            wbGallery.Navigate "about:blank"
            InitGallery 'bInitializedGallery will be set on lstDir_Change according to PopulateGallery
        End If
        
        bFirstCallDone = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Cleans the environment (i.e. temporary files which have been created on Form_Load)
    Dim strContentCSS As String
    Dim strEditorCurrent As String
    
    strContentCSS = App.Path & "\" & EDITOR_DIR & "\" & EDITOR_CONTENT_CSS
    If Dir(strContentCSS) <> "" Then Kill strContentCSS
    
    strEditorCurrent = App.Path & "\" & EDITOR_DIR & "\" & EDITOR_CURRENT
    If Dir(strEditorCurrent) <> "" Then Kill strEditorCurrent
End Sub

Private Sub lstDir_Change()
    'Updates the images list from current selected directory (three attempts in case wbGallery would be not ready)
    Dim nAttempts As Integer
    Dim bRet As Boolean
    
    Do
        bRet = PopulateGallery()
        nAttempts = nAttempts + 1
    Loop While bRet = False And nAttempts < 3
    bInitializedGallery = bRet
End Sub

Private Sub lstDrive_Change()
    'Updates the directories list from current selected unit
    lstDir.Path = Left$(lstDrive.Drive, 2) & "\"
End Sub

Private Sub wbEditor_StatusTextChange(ByVal Text As String)
    'Saves the article and its title in the strCurrArt file, then exit if required
    'Algo : clean & save process was intiated from within the TinyMCE webpage itself, through javascript initSave()
    'NB : includes uniformization of the different elements to the encoding defined in project settings (ANSI/UTF8)
    Dim strTitle As String
    Dim strContent As String
    
    'Saves if asked in status-bar
    'IMPORTANT : TinyMCE being designed for online client/server relationship only, this trick going through
    '            status-bar allows to pass request-for-saving from within embedded TinyMCE to VB6-app "host" :o)
    If Left$(Text, 12) = "save_trigger" Then
        'extracts current title from wbEditor
        strTitle = wbEditor.Document.All.title.innerHTML
        
        If UCase(structProj.art_encode) = "UTF-8" Then
            If IsUTF8(strTitle) = False Then strTitle = Encode_UTF8(strTitle)
        Else 'ANSI (default)
            If IsUTF8(strTitle) = True Then strTitle = Decode_UTF8(strTitle)
        End If
        
        If strTitle = "" Then
            MsgBox arMsg(113) & " !", vbExclamation
            Exit Sub
        End If
        
        'extracts current content from TinyMCE in wbEditor
        strContent = wbEditor.Document.All.buffer.innerHTML
        
        'fixes object tag since IE - here as ActiveX - may remove src param
        strContent = FixObjectTag(strContent)
        
        'adds a generator sight at the beginning if there's no one anywhere
        If InStr(1, strContent, EDITOR_GENERATOR_SIGHT, vbTextCompare) = 0 Then
            strContent = EDITOR_GENERATOR_SIGHT & vbNewLine & vbNewLine & strContent
        End If
        
        'adds a witness mark of current encoding at the end if there's no one anywhere
        'NB : it cancels eventual effect of HTML-entity conversion done by TinyMCE
        strContent = Replace(strContent, EDITOR_ENCODING_WITNESS_AS_ENTITE, EDITOR_ENCODING_WITNESS)
        If InStr(1, strContent, EDITOR_ENCODING_WITNESS, vbTextCompare) = 0 Then
            strContent = strContent & EDITOR_ENCODING_WITNESS
        End If
        
        'then re-encodes if necessary
        If UCase(structProj.art_encode) = "UTF-8" Then
            If IsUTF8(strContent) = False Then strContent = Encode_UTF8(strContent)
        Else 'ANSI (default)
            If IsUTF8(strContent) = True Then strContent = Decode_UTF8(strContent)
        End If
        
        'sets article title as first line, then one empty line as required by Blosxom
        strContent = strTitle & vbNewLine & vbNewLine & strContent
        
        SaveText strCurrArt, strContent
        strCurrArt = ""
    End If
    
    'Closes if requested from statut-bar
    'NB : same trick as for saving above : info is passed from webpage to app via keywords in status-bar ^^
    If Right$(Text, 12) = "exit_trigger" Then
        strCurrArt = ""
        PopulateClearLocalList True  'takes care of eventual encoding change
        Me.Hide                      'editor is hidden, but remains in memory
        ResetEditor                  'editor is prepared to be ready for next time (ie. next article edit)
    End If
End Sub

Function ResetEditor() As Boolean
          'Resets the editor prior to every edition, first included (takes care of known error #91)
          'OUT : true if succeeded, false otherwise
          'USE : called on frmEdit load, then on activate if failed on load, then at every frmEdit hiding)
          'WARNING : this fct being a sensitive point, we provide line numbers to manage accurate error messages
          Dim strMsg As String
          Static bExistingDoc As Boolean
          
1         On Error GoTo ResetEditor_Error

2         strMsg = "<html><body><div style='font: 12pt MS Sans Serif; color: #808080'> Chargement en cours...</div></body></html>"
3         wbEditor.Document.write strMsg 'reset UI en prévision d'un prochain appel
4         wbEditor.Document.Close

5         ResetEditor = True
ResetEditor_End:
6        On Error GoTo 0
7        Exit Function

ResetEditor_Error:
8         If Err.Number = 91 And bExistingDoc = False Then
              'wbEditor.Document being not created, we try to initialize one prior to first .write
9             wbEditor.Navigate "about:blank"
10            DoEvents
11            bExistingDoc = True
12            Resume
13        End If

14        MsgBox "Error #" & Err.Number & "@ bloshome/frmEdit/ResetEditor/#" & Erl & " : " & Err.Description, vbExclamation
15        ResetEditor = False
16        Resume ResetEditor_End
End Function

Function InitGallery() As Boolean
    'Initializes the gallery prior to first edition (called on form load, then on activate if failed on load)
    'OUT : true if succeeded, false otherwise
    'Algo : triggers lstDir update which will, at its turn, trigger PopulateGallery()
    lstDrive.Drive = "C:"
    lstDir.Path = "C:\" 'bypasses the usual lstDrive_Change that auto-updates lstDir (just to be sure)
End Function

Function PopulateGallery() As Boolean
          'Populates wbGallery with image files from current selected directory (takes care of known error #91)
          'OUT : true if succeeded, false otherwise (allows to run several times if caller whishes it)
          'WARNING : this fct being a sensitive point, we provide line numbers to manage accurate error messages
          Dim nIdx As Integer
          Dim bFirstLoop As Boolean
          Dim bSelectFile As Boolean
          
          Dim strPath As String
          Dim strFilename As String
          Dim arExt(1 To 3) As String
          Dim strFullPath As String
          
          Dim bHasImgMax As Boolean
          Dim nSize As Long

          Dim strHTML As String
          
          Static bExistingDoc As Boolean
          
          'initialization
1         On Error GoTo PopulateGallery_Error

2         arExt(1) = ".jpg"
3         arExt(2) = ".jpeg"
4         arExt(3) = ".gif"
          
5         strHTML = "<html><head>" & _
                    "<style>" & _
                    "body {font: 8pt sans-serif}" & _
                    "</style>" & _
                    "</head><body><center>"
          
          'path checking (includes its standardization)
6         strPath = lstDir.Path
              
7         If strPath = "" Then
8             wbGallery.Document.All.write ""
9             Exit Function
10        Else
11            If Right$(strPath, 1) <> "\" Then strPath = strPath & "\"
12        End If
          
          'effective research of the images
13        bFirstLoop = True
14        Do
              'file by file
15            If bFirstLoop = True Then
16                strFilename = Dir(strPath & "*.*", vbNormal)
17                bFirstLoop = False
18            Else
19                strFilename = Dir
20            End If
              
              'if it presents the right extension and an acceptable size (according to project settings)
21            bHasImgMax = IsNumeric(structProj.img_max)
22            For nIdx = 1 To UBound(arExt)
23                If InStr(1, strFilename, arExt(nIdx), vbTextCompare) <> 0 Then
24                    strFullPath = strPath & strFilename
25                    nSize = FileLen(strFullPath) / 1024 'in KB
                      
26                    If bHasImgMax = True Then
27                        If nSize <= structProj.img_max Then
28                            bSelectFile = True
29                            Exit For
30                        Else
31                            bSelectFile = False
32                        End If
33                    Else
34                        bSelectFile = True
35                        Exit For
36                    End If
37                Else
38                    bSelectFile = False
39                End If
40            Next
              
              'we add a thumbnail in the gallery
41            If bSelectFile = True Then
                  Dim nThSideMax As Integer
                  Dim nWidth As Integer
                  Dim nHeight As Integer
                  Dim nThWidth As Integer
                  Dim nThHeight As Integer
                  Dim strSRC As String
                  
42                nThSideMax = 96 ' in pixels
43                Call GetImgDims(Me, strFullPath, nWidth, nHeight)
44                Call CalcThumbDims(nThSideMax, nWidth, nHeight, nThWidth, nThHeight)
                  
                  '(trick to do all URL be OK, corecting incompatible encoding for .write)
45                strSRC = Replace(strFullPath, "\", "/")
46                strSRC = Replace(strSRC, "%C9", "É")   '***LATER : do a sub to restore accented capitals
47                strSRC = Replace(strSRC, "'", "&#39;") '***LATER : do a sub to encode HTML entities
                  
48                strHTML = strHTML & "<img src='" & strSRC & "' border=0 " & _
                                      "width='" & nThWidth & "px' " & _
                                      "height='" & nThHeight & "px'></img> " & _
                                      "<br> " & strFilename & _
                                      "<br>" & nWidth & "x" & nHeight & " - " & nSize & "KB" & _
                                      "<br><br>"
49            End If
50        Loop While strFilename <> ""
          
          'informs user that there's no image in current directory
51        If Right(strHTML, 4) <> "<br>" Then
52            strHTML = strHTML & arMsg(142) & "."
53        End If
          
          'finalization
54        strHTML = strHTML & "</center></body></html>"
55        wbGallery.Document.write strHTML
56        wbGallery.Document.Close

57        PopulateGallery = True
PopulateGallery_End:
58       On Error GoTo 0
59       Exit Function

PopulateGallery_Error:
60        If Err.Number = 91 And bExistingDoc = False Then
              'wbGallery.Document being not created, we try to initialize one prior to first .write
61            wbGallery.Navigate "about:blank"
62            DoEvents
63            bExistingDoc = True
64            Resume
65        End If

66        MsgBox "Error #" & Err.Number & "@ bloshome/frmEdit/ResetEditor/#" & Erl & " : " & Err.Description, vbExclamation
67        PopulateGallery = False
68        Resume PopulateGallery_End
End Function

Sub EditCurrArt()
    'Loads the current strCurrArt article in the editor (which is already in loaded in memory)
    'NB : includes uniformization of different data - CSS/HTML/.ART - to the right encoding
    Dim strEditorTemplate As String
    Dim strEditorCurrent As String
    Dim strContentCSS As String
    
    Dim strHTML As String
    Dim strCSS As String
    Dim strTitleCSS As String
    Dim strCharset As String
    Dim strLang As String
    
    Dim strArt As String
    Dim strTitle As String
    Dim strBody As String
    
    Dim nPos As Integer
    Dim nStart As Integer
    Dim nEnd As Integer
    
    'init and checks
    strEditorTemplate = App.Path & "\" & EDITOR_DIR & "\" & EDITOR_TEMPLATE
    strEditorCurrent = App.Path & "\" & EDITOR_DIR & "\" & EDITOR_CURRENT
    strContentCSS = App.Path & "\" & EDITOR_DIR & "\" & EDITOR_CONTENT_CSS
    
    strArt = LoadText(strCurrArt)
    
    'adjust UI
    Caption = Caption & " - " & strCurrArt
    
    'defines the charset
    If UCase(structProj.art_encode) = "UTF-8" Then
        strCharset = "utf-8"
    Else
        strCharset = "iso-8859-1" 'default
    End If
    
    'defines the language
    strLang = LCase(GetLang())
    
    'prepares the CSS file (extracts -> encodes -> saves)
     If Dir(structProj.css) <> "" Then
        strCSS = LoadText(structProj.css)
    
        If UCase(structProj.art_encode) = "UTF-8" Then
            If IsUTF8(strCSS) = False Then strCSS = Encode_UTF8(strCSS)
        Else 'ANSI (default)
            If IsUTF8(strCSS) = True Then strCSS = Decode_UTF8(strCSS)
        End If
    
        SaveText strContentCSS, strCSS 'overwrites if exists
    End If
    
    'isolates an eventual #title from global CSS (will be added to #title inline-style of editor.html)
    strTitleCSS = ""
    nPos = InStr(1, strCSS, "#title", vbTextCompare)
    If nPos <> 0 Then
        nStart = InStr(nPos, strCSS, "{", vbTextCompare)
        nEnd = InStr(nPos, strCSS, "}", vbTextCompare)
        If nStart <> 0 And nEnd <> 0 And nEnd > nStart Then
            strTitleCSS = Mid$(strCSS, nStart + 1, nEnd - nStart - 1)
        End If
    End If
    
    'loads the title (1st ligne followed by an empty line ; see wbEditor_StatusTextChange)
    If strArt <> "" Then
        nPos = InStr(1, strArt, vbNewLine & vbNewLine, vbTextCompare)
        strTitle = Left$(strArt, nPos - 1)
    Else
        strTitle = ""
    End If
    
    strHTML = LoadText(strEditorTemplate)
    strHTML = Replace(strHTML, EDITOR_TITLE_CSS_HOLDER, strTitleCSS)
    strHTML = Replace(strHTML, EDITOR_TITLE_HOLDER, strTitle)
    
    'loads the article itself
    If strArt <> "" Then strBody = Right$(strArt, Len(strArt) - Len(strTitle) - (Len(vbNewLine) * 2))
    
    strHTML = Replace(strHTML, EDITOR_CHARSET_HOLDER, strCharset)
    strHTML = Replace(strHTML, EDITOR_LANG_HOLDER, Replace(EDITOR_LANG_HOLDER, "fr", strLang))
    strHTML = Replace(strHTML, EDITOR_CSS_HOLDER, EDITOR_CONTENT_CSS)
    strHTML = Replace(strHTML, EDITOR_CONTENT_HOLDER, strBody)
    
    'considers eventual chapo plugin (depending of project settings)
    If structProj.with_chapo = True Then
        Dim strChapo As String
        Dim strChapoTag As String
        Dim strChapoBtn As String
                
        strChapo = ",pagebreak"
        strChapoTag = "pagebreak_separator : " & Chr(34) & "<!-- " & structProj.chapo_limit & " -->" & Chr(34) & ","
        strChapoBtn = ",|,pagebreak"
        
        strHTML = Replace(strHTML, PAGEBREAK_HOLDER, strChapo)
        strHTML = Replace(strHTML, PAGEBREAK_TAG_HOLDER, strChapoTag)
        strHTML = Replace(strHTML, PAGEBREAK_BTN_HOLDER, strChapoBtn)
    Else
        strHTML = Replace(strHTML, PAGEBREAK_HOLDER, "")
        strHTML = Replace(strHTML, PAGEBREAK_TAG_HOLDER, "")
        strHTML = Replace(strHTML, PAGEBREAK_BTN_HOLDER, "")
    End If
    
    '(re)encoding if necessary
    If UCase(structProj.art_encode) = "UTF-8" Then
        If IsUTF8(strHTML) = False Then strHTML = Encode_UTF8(strHTML)
    Else 'ANSI (default)
        If IsUTF8(strHTML) = True Then strHTML = Decode_UTF8(strHTML)
    End If
    
    'opening
    If Dir(strEditorCurrent) <> "" Then Kill strEditorCurrent
    SaveText strEditorCurrent, strHTML
    wbEditor.Navigate strEditorCurrent
End Sub
