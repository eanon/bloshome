VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.UserControl FtpCli 
   ClientHeight    =   1515
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2325
   InvisibleAtRuntime=   -1  'True
   Picture         =   "FtpCli.ctx":0000
   ScaleHeight     =   1515
   ScaleWidth      =   2325
   ToolboxBitmap   =   "FtpCli.ctx":030A
   Begin MSWinsockLib.Winsock ftpData 
      Left            =   840
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock ftpDialog 
      Left            =   240
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "FtpCli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim FtpProtocolData As String
Dim FtpTransfertData As String
Dim FtpCode As Integer
Dim TransfertPort As Long
Dim FileNameToSend As String
Dim FileNameToGet As String
Dim SendOk As Boolean

Dim m_DirList As Variant
Dim m_RemoteHost As String
Dim m_RemotePort As Long
Dim m_UserName As String
Dim m_Password As String
Dim m_RemoteDir As String
Dim m_TimeOutDelay As Integer

Public Event ProtocolArrival(Data As String)
Public Event ProtocolSend(Data As String)
Public Event DataArrival(Data As String)
Public Event SendProgress(Pcent As Integer, Size As Long)
Public Event RecieveProgress(Size As Long)


'Variables de propriétés:
Dim m_RemoteFiles As RemoteFiles
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)







'===========================================
' UserControl events
'===========================================
Private Sub UserControl_Initialize()
  Set m_RemoteFiles = New RemoteFiles
End Sub
Private Sub UserControl_InitProperties()
  m_RemoteHost = "Localhost"
  m_RemotePort = 21
  m_UserName = "anonymous"
  m_Password = "jean-luc@delbeke.fr"
  m_TimeOutDelay = 60
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  m_RemoteHost = PropBag.ReadProperty("RemoteHost", "")
  m_RemotePort = PropBag.ReadProperty("RemotePort", 21)
  m_UserName = PropBag.ReadProperty("UserName", "ANONYMOUS")
  m_Password = PropBag.ReadProperty("Password", "anyone@anyserver.com")
  m_TimeOutDelay = PropBag.ReadProperty("TimeOutDelay", 60)
End Sub

Private Sub UserControl_Resize()
  Size 500, 500
End Sub

Private Sub UserControl_Terminate()
  If ftpDialog.State <> 0 Then
    Ftp_QUIT
    Do While m_RemoteFiles.Count > 0
      m_RemoteFiles.Remove 1
    Loop
  End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  PropBag.WriteProperty "RemoteHost", m_RemoteHost, ""
  PropBag.WriteProperty "m_RemotePort", m_RemotePort, 21
  PropBag.WriteProperty "UserName", m_UserName, "ANONYMOUS"
  PropBag.WriteProperty "Password", m_Password, "anyone@anyserver.com"
  PropBag.WriteProperty "TimeOutDelay", m_TimeOutDelay, 60
End Sub

'===========================================
'Winsock events
'===========================================
Private Sub ftpdialog_DataArrival(ByVal bytesTotal As Long)
  Dim varTemp As Variant
  Dim iPnt As Integer
  Static strTemp As String
  Dim ftpString As String
  ftpDialog.GetData ftpString
  'split ftpstring
  Debug.Print ftpString
  varTemp = Split(ftpString, vbCrLf)
  'raiseevents for protocol data arrival
  For iPnt = 0 To UBound(varTemp) - 1
    RaiseEvent ProtocolArrival(CStr(varTemp(iPnt)))
'    If Mid(varTemp(UBound(varTemp) - 1), 4, 1) <> "-" Then
    If Mid(varTemp(iPnt), 4, 1) = " " And IsNumeric(Left(varTemp(iPnt), 3)) Then
        FtpCode = Left(varTemp(UBound(varTemp) - 1), 3)
        FtpProtocolData = strTemp & CStr(varTemp(iPnt))
        strTemp = ""
        DoEvents
        'Debug.Print "!"; FtpProtocolData; "!"
    Else
      strTemp = strTemp & CStr(varTemp(iPnt))
    End If
  Next
End Sub

Private Sub ftpData_Close()
  On Error Resume Next
  ftpData.Close
  On Error GoTo 0
End Sub
Private Sub ftpData_ConnectionRequest(ByVal requestID As Long)
  ftpData.Accept requestID
End Sub

Private Sub ftpdata_DataArrival(ByVal bytesTotal As Long)
  Dim ftpDataReceived As String
  ' On récupère la réponse du serveur ( attention elle peut être multiligne)
  ftpData.GetData ftpDataReceived
  FtpTransfertData = FtpTransfertData & ftpDataReceived
  RaiseEvent DataArrival(ftpDataReceived)
  'Debug.Print FtpTransfertData
End Sub

Private Sub ftpData_SendComplete()
  SendOk = True
End Sub
'===========================================
' Properties
'===========================================
Public Property Get RemoteHost() As String
  RemoteHost = m_RemoteHost
End Property
Public Property Let RemoteHost(ByVal vNewValue As String)
  m_RemoteHost = vNewValue
  PropertyChanged "RemoteHost"
End Property

Public Property Get RemotePort() As Long
  RemotePort = m_RemotePort
End Property
Public Property Let RemotePort(ByVal vNewValue As Long)
  m_RemotePort = vNewValue
  PropertyChanged "RemotePort"
End Property

Public Property Get UserName() As Variant
  UserName = m_UserName
End Property
Public Property Let UserName(ByVal vNewValue As Variant)
  m_UserName = vNewValue
  PropertyChanged "UserName"
End Property

Public Property Get Password() As Variant
  Password = m_Password
End Property
Public Property Let Password(ByVal vNewValue As Variant)
  m_Password = vNewValue
  PropertyChanged "Password"
End Property

Public Property Get DirList() As Variant
Attribute DirList.VB_MemberFlags = "400"
  DirList = m_DirList
End Property

Public Static Property Get TimeOutDelay() As Integer
  TimeOutDelay = m_TimeOutDelay
End Property
Public Static Property Let TimeOutDelay(ByVal vNewValue As Integer)
  m_TimeOutDelay = vNewValue
  If vNewValue = 0 Then
    m_TimeOutDelay = 32565
  End If
  PropertyChanged "TimeOutDelay"
End Property

Public Property Get RemoteDir() As String
  RemoteDir = m_RemoteDir
End Property

Public Property Get RemoteFiles() As RemoteFiles
  Set RemoteFiles = m_RemoteFiles
End Property

'===========================================
' Public Methods
'===========================================
'**************
'Meta commands
'*************
Public Function Met_CONNECT() As Boolean
  Met_CONNECT = False
  If m_RemoteHost = "" Then
    Met_CONNECT = False
    MsgBox "No remote host"
    Exit Function
  End If
  'Met_Connect ftp
  If Met_OPEN(m_RemoteHost, m_RemotePort) = False Then
    Met_CONNECT = False
    Exit Function
  End If
  'Send User
  Select Case Ftp_USER(m_UserName)
  Case 0
    Met_CONNECT = False
  Case 1
    'send password
    If Ftp_PASS(m_Password) = False Then
      Met_CONNECT = False
    Else
      Met_CONNECT = True
      Ftp_PWD
    End If
  Case -1
    Met_CONNECT = True
    Ftp_PWD
  End Select
End Function
Public Function Met_BYE() As Boolean
  Met_BYE = Ftp_QUIT
End Function
Public Function Met_CLOSE() As Boolean
  Met_CLOSE = Ftp_QUIT
End Function
Public Function Met_DISCONNECT() As Boolean
  Met_DISCONNECT = Ftp_QUIT
End Function
Public Function Met_DIR()
  'List directory
  Met_DIR = True
  If Not Ftp_PASV Then
    MsgBox "Error entering passive mode"
    Met_DIR = False
    Exit Function
  End If
  If Not Ftp_LIST Then
    Met_DIR = False
    'MsgBox "Error in List Command"
  End If
End Function
Public Function Met_ASCII() As Boolean
  'ASCII mode for transfert
  Met_ASCII = Ftp_TYPE("A")
End Function
Public Function Met_BINARY() As Boolean
  'Binary mode for transfert
  Met_BINARY = Ftp_TYPE("I")
End Function
Public Function Met_CD(DirectoryName As String) As Boolean
  'change directory
  Met_CD = Ftp_CWD(DirectoryName)
  Ftp_PWD
End Function
Public Function Met_CDUP() As Boolean
  'change directory
  Met_CDUP = Ftp_CDUP()
  Ftp_PWD
End Function
Public Function Met_OPEN(Host As String, Optional Port As Long = 21) As Boolean
  Met_OPEN = Ftp_OPEN(Host, Port)
End Function
Public Function Met_PUT(FileName As String, Optional RemoteFileName As String) As Boolean
  Dim iPosit As Integer
  FileNameToSend = FileName
  If RemoteFileName = "" Then
    iPosit = InStrRev(FileName, "\")
    RemoteFileName = Mid$(FileName, iPosit + 1)
  End If
  If Not Ftp_TYPE("I") Then
      Met_PUT = False
      Exit Function
  End If
  If Not Ftp_PASV Then
      Met_PUT = False
      Exit Function
  End If
  
  Met_PUT = Ftp_STOR(RemoteFileName)
  
End Function
Public Function Met_QUOTE(Param As String) As Boolean
  Met_QUOTE = Ftp_QUOTE(Param)
End Function
Public Function Met_GET(FileName As String, Optional DestFilename As String = "") As Boolean
  Dim RemoteFileName As String
  Dim iPosit As Integer
  FileNameToSend = FileName
  iPosit = InStrRev(FileName, "\")
  RemoteFileName = Mid$(FileName, iPosit + 1)
  If DestFilename = "" Then
    FileNameToGet = App.Path & "\" & FileName
  Else
    FileNameToGet = DestFilename
  End If
  If Not Ftp_TYPE("I") Then
    Met_GET = False
    Exit Function
  End If
  If Not Ftp_PASV Then
    Met_GET = False
    Exit Function
  End If
  If Not Ftp_RETR(RemoteFileName) Then
    Met_GET = False
  Else
    Met_GET = True
  End If
End Function
Public Function Met_SPLITPDF(FileName As String, StartPage As Long, EndPage As Long) As Boolean

  'fonction non suportée par un serveur FTP standard
  
  Dim RemoteFileName As String
  Dim iPosit As Integer
  Dim RetFileName As String
  If Not Met_PUT(FileName) Then
    Met_SPLITPDF = False
    Exit Function
  End If
  'get the remote filename
  iPosit = InStrRev(FileName, "\")
  RemoteFileName = Mid$(FileName, iPosit + 1)
  'get the remote splitted filename
  RetFileName = Left(RemoteFileName, InStrRev(RemoteFileName, ".") - 1) & "_split.pdf"

  SendProtocol "FORMULARY /D= /U=LotusEmail /X=PdfSplitter /P=[/F=*\" & RemoteFileName & " /P=" & CStr(StartPage) & " /T=" & CStr(EndPage) & "]", False
  If FtpCode <> 230 Then
    Met_SPLITPDF = False
    Exit Function
  End If
  'wait for end of work
  FtpCode = 0
  Do While FtpCode = 0
    DoEvents
  Loop
  If FtpCode <> 250 Then
    Met_SPLITPDF = False
    Exit Function
  End If
  Met_GET RetFileName
  If FtpCode <> 226 Then
    Met_SPLITPDF = False
    Exit Function
  End If
  Ftp_DELE RemoteFileName
  Ftp_DELE RetFileName
  Met_SPLITPDF = True
End Function
Public Function Met_DELETE(FileName As String) As Boolean
  'Delete file on server
  Dim RemoteFileName As String
  Dim iPosit As Integer
  iPosit = InStrRev(FileName, "\")
  RemoteFileName = Mid$(FileName, iPosit + 1)
  Met_DELETE = Ftp_DELE(RemoteFileName)
End Function

Public Function Met_RENAME(RemoteFileName As String, DestFilename As String) As Boolean
  'Rename Filename to destfilename
  If Not Ftp_RNFR(RemoteFileName) Then
    Met_RENAME = False
    Exit Function
  End If
  If Not Ftp_RNTO(DestFilename) Then
    Met_RENAME = False
  Else
    Met_RENAME = True
  End If
End Function
Public Function Met_Exec(Cmd As String) As Boolean
  SendProtocol Cmd
End Function

'**************
'Basic commands
'*************
Public Function Ftp_ABOR() As Boolean
  'abort current transfert
  SendProtocol "ABOR"
  If FtpCode = 226 Then
    Ftp_ABOR = True
  Else
    Ftp_ABOR = False
  End If
End Function
Public Function Ftp_APPPE(RemoteFileName As String) As Boolean
  'append a file
  Dim iPosit As Integer
  Dim hFich As Integer
  Dim Buffer As String
  Dim BlockSize As Long
  Dim Rest As Long
  Dim FileSize As Long
  Dim Send As Long
  If FileNameToSend = "" Then
    FileNameToSend = App.Path & "\" & RemoteFileName
  End If
  BlockSize = 4096
  Send = 0
  FileSize = FileLen(FileNameToSend)
  Rest = FileSize
  SendProtocol "APPE " & RemoteFileName
  If FtpCode <> 150 And FtpCode <> 125 Then
    Ftp_APPPE = False
    Exit Function
  End If
  hFich = FreeFile
  Open FileNameToSend For Binary As #hFich
    Do While Rest > 0
      If BlockSize > Rest Then
        BlockSize = Rest
      End If
      Buffer = String(BlockSize, 0)
      Get #hFich, , Buffer
      SendOk = False
      ftpData.SendData Buffer
      'wait for send complete
      Do While Not SendOk
        DoEvents
      Loop
      Send = Send + BlockSize
      Rest = FileSize - Send
    Loop
  Close #hFich
  FtpCode = 0
  CloseFtpData
  Do While FtpCode <> 226
    DoEvents
  Loop
    
  DoEvents
  FileNameToSend = ""
  Ftp_APPPE = True
End Function
Public Function Ftp_APPEND(DestFilename As String, FileName As String)
  'add a file FileName to DestFileName
End Function
Public Function Ftp_BYE()
  'same as Disconnect
  Met_DISCONNECT
End Function
Public Function Ftp_CDUP() As Boolean
  'move to parent directory
  Dim iPosit As String
  SendProtocol "CDUP"
  If FtpCode <> 250 Then
    Ftp_CDUP = False
  Else
    Ftp_CDUP = True
  End If
End Function

Public Function Ftp_CWD(DirectoryName As String) As Boolean
  'change current directory
  Dim iPosit As String
  SendProtocol "CWD " & Trim(DirectoryName)
  If FtpCode <> 250 Then
    Ftp_CWD = False
  Else
    Ftp_CWD = True
    m_RemoteDir = "/" & DirectoryName
End If
End Function
Public Function Ftp_CLOSE() As Boolean
' same as disconnect
  Ftp_CLOSE = Met_DISCONNECT
End Function
Public Function Ftp_DELE(RemoteFileName As String) As Boolean
  'delete a file
  SendProtocol "DELE " & RemoteFileName
  If FtpCode <> 250 Then
    Ftp_DELE = False
  Else
    Ftp_DELE = True
  End If
End Function
Public Function Ftp_HELP() As Boolean
  'help from server
  SendProtocol "HELP"
  If FtpCode <> 214 Then
    Ftp_HELP = False
  Else
    Ftp_HELP = True
  End If
End Function
Public Function Ftp_LIST() As Boolean
  'Get List directory
  Dim iPnt As Integer
  Dim jPnt As Integer
  Dim FilePart As Variant
  Dim FlType As FileTypes
  Dim FlName As String
  Dim FlDate As Date
  Dim FlSize As Long
  Dim FlMonth As Integer
  Dim MaxDate As Date
  Dim TimeOut As Boolean
  'open a connection with a server
  MaxDate = DateAdd("s", m_TimeOutDelay, Now)
  TimeOut = False
  
  FtpTransfertData = ""
  Ftp_LIST = False
  If ftpData.State <> 0 Then
    SendProtocol "LIST", False
    If FtpCode = 150 Or FtpCode = 125 Then
      FtpCode = 0
      Do While FtpCode = 0 And Not TimeOut
        If Now > MaxDate Then
          TimeOut = True
        End If
        DoEvents
      Loop
      If TimeOut Then
        'RaiseEvent ProtocolSend("Time out on Connecting to " & Host & " Port " & CStr(Port))
        Ftp_LIST = False
        Exit Function
      End If
    Else
      Ftp_LIST = False
      Exit Function
    End If
    If FtpCode <> 226 Then
      Ftp_LIST = False
    Else
      Ftp_LIST = True
      m_DirList = Split(FtpTransfertData, vbCrLf)
      Do While m_RemoteFiles.Count > 0
        m_RemoteFiles.Remove 1
      Loop
      For iPnt = 0 To UBound(m_DirList)
        'Debug.Print m_DirList(iPnt)
        Do While InStr(m_DirList(iPnt), "  ")
          m_DirList(iPnt) = Replace(m_DirList(iPnt), "  ", " ")
        Loop
        If Trim(m_DirList(iPnt)) <> "" Then
          FilePart = Split(m_DirList(iPnt), " ")
          'For jPnt = 0 To UBound(FilePart)
          '  Debug.Print FilePart(jPnt); "!";
          'Next
          'Debug.Print
          If UBound(FilePart) >= 8 Then
            Select Case LCase(Left(FilePart(0), 1))
            Case "d"
              FlType = FtpDirectory
            Case "-"
              FlType = FtpFile
            Case "l"
              FlType = FtpLink
            Case Else
              FlType = FtpUnkwnon
            End Select
            FlName = ""
            For jPnt = 8 To UBound(FilePart)
              FlName = FlName & FilePart(jPnt) & " "
            Next
            FlName = RTrim(FlName)
            Select Case LCase(FilePart(5))
            Case "jan"
              FlMonth = 1
            Case "feb"
              FlMonth = 2
            Case "mar"
              FlMonth = 3
            Case "apr"
              FlMonth = 4
            Case "may"
              FlMonth = 5
            Case "jun"
              FlMonth = 6
            Case "jul"
              FlMonth = 7
            Case "aug"
              FlMonth = 8
            Case "sep"
              FlMonth = 9
            Case "oct"
              FlMonth = 10
            Case "nov"
              FlMonth = 11
            Case Else
              FlMonth = 12
            End Select
            On Error Resume Next
            FlDate = Format(FilePart(6) & "/" & CStr(FlMonth) & "/" & CStr(Year(Date)) & " " & FilePart(7), "dd/mm/yyyy hh:nn")
            FlSize = FilePart(4)
            m_RemoteFiles.Add FlName, FlSize, FlDate, FlType, CStr(m_DirList(iPnt)), "K" & FlName
          Else
            'ligne ne comprenant pas 8 elements
            'For jPnt = 0 To UBound(FilePart)
            '  Debug.Print FilePart(jPnt); "!";
            'Next
            'Debug.Print
          End If
        End If
      Next
    End If
  End If
End Function
Public Function Ftp_LOGPROG(ProgName As String) As Boolean

  'fonction non suportée par un serveur FTP standard
  
    SendProtocol "LOGPROG " & ProgName, True
    If FtpCode <> 230 Then
      Ftp_LOGPROG = False
    Else
      Ftp_LOGPROG = True
    End If
End Function
Public Function Ftp_MDTM(RemoteFileName As String) As Boolean
  SendProtocol "MDTM " & RemoteFileName
  If FtpCode <> 213 Then
    Ftp_MDTM = False
  Else
    Ftp_MDTM = True
  End If
End Function
Public Function Ftp_MKD(RemoteDirectory As String) As Boolean
  SendProtocol "MKD " & RemoteDirectory
  If FtpCode <> 257 Then
    Ftp_MKD = False
  Else
    Ftp_MKD = True
  End If
End Function
Public Function Ftp_NOOP() As Boolean
  SendProtocol "NOOP"
  If FtpCode <> 200 Then
    Ftp_NOOP = False
  Else
    Ftp_NOOP = True
  End If
End Function
Public Function Ftp_OPEN(Host As String, Optional Port As Long = 21) As Boolean
  Dim MaxDate As Date
  Dim TimeOut As Boolean
  'open a connection with a server
  MaxDate = DateAdd("s", m_TimeOutDelay, Now)
  TimeOut = False
  ftpDialog.Close
  RaiseEvent ProtocolSend("Connecting to " & Host & " Port " & CStr(Port))
  FtpCode = 0
  ftpDialog.Connect Host, Port
  'wait for response
  Do While FtpCode = 0 And Not TimeOut
    If Now > MaxDate Then
      TimeOut = True
    End If
    DoEvents
  Loop
  If TimeOut Then
    RaiseEvent ProtocolSend("Time out on Connecting to " & Host & " Port " & CStr(Port))
    Ftp_OPEN = False
    ftpDialog.Close
    Exit Function
  End If
  'control response
  If FtpCode <> 220 Then
    Ftp_OPEN = False
  Else
    Ftp_OPEN = True
  End If
End Function
Public Function Ftp_PASS(Pass As String) As Boolean
  SendProtocol "PASS " & Pass
  If FtpCode <> 230 Then
    Ftp_PASS = False
  Else
    Ftp_PASS = True
  End If
End Function
Public Function Ftp_PASV() As Boolean
  'enter passive mode
  Dim varTemp As Variant
  SendProtocol "PASV"
  If FtpCode <> 227 Then
    Ftp_PASV = False
  Else
    varTemp = Split(FtpProtocolData, ")")
    varTemp = Split(varTemp(0), "(")
    varTemp = Split(varTemp(1), ",")
    TransfertPort = CLng(varTemp(4)) * 256 + CLng(varTemp(5))
    CloseFtpData
    ftpData.LocalPort = 0
    ftpData.Connect m_RemoteHost, TransfertPort
    'wait for connectcomplete
    Sleep 500
    DoEvents
    Do While ftpData.State <> 7
      DoEvents
    Loop
    Ftp_PASV = True
  End If
End Function
Public Function Ftp_PORT(HostNumber As Long, PortNumber As Long)
  TransfertPort = PortNumber
  'force Port fordata transfert
  CloseFtpData
  '
  ftpData.LocalPort = TransfertPort
  ftpData.Listen
  SendProtocol "PORT " & Replace(HostNumber, ".", ",") & "," & CStr(PortNumber \ 256) & "," & CStr(Fix(PortNumber Mod 256))
  
  If FtpCode <> 200 Then
    Ftp_PORT = False
  Else
    Ftp_PORT = True
  End If
End Function
Public Function Ftp_PWD() As Boolean
  Dim iPosit As Integer
  SendProtocol "PWD"
  If FtpCode <> 257 Then
    Ftp_PWD = False
  Else
    Ftp_PWD = True
    m_RemoteDir = Mid(FtpProtocolData, 5)
    iPosit = InStrRev(m_RemoteDir, Chr$(34))
    m_RemoteDir = Left(m_RemoteDir, iPosit)
    m_RemoteDir = Replace(m_RemoteDir, Chr$(34), "")
  End If
End Function
Public Function Ftp_QUIT()
  'SendProtocol "QUIT"
  RaiseEvent ProtocolSend("QUIT")
  ftpDialog.SendData "QUIT" & vbCrLf
  DoEvents
  ftpDialog.Close
  CloseFtpData
  Ftp_QUIT = True
End Function
Public Function Ftp_QUOTE(Param As String) As Boolean
  SendProtocol Param
  Ftp_QUOTE = True
  'les reponses négatives commencent à partir de 400
  'on ne peut pas etre plus precis dans la commande quote
  If FtpCode >= 400 Then
    Ftp_QUOTE = False
  Else
    Ftp_QUOTE = True
  End If
End Function
Public Function Ftp_REMOTEHELP() As Boolean
  'help from server
End Function
Public Function Ftp_RENAME(FileName As String, DestFilename As String)
  'Rename Filename to destfilename
End Function
Public Function Ftp_REST(ReStartPos As Long) As Boolean
  'restart position for retr
  SendProtocol "REST " & CStr(ReStartPos)
  If FtpCode <> 350 Then
    Ftp_REST = False
  Else
    Ftp_REST = True
  End If
End Function

Public Function Ftp_RETR(RemoteFileName As String) As Boolean
  'retrieve file (donwload)
  Dim Recieved As Long
  Dim hFich As Integer
  If FileNameToGet = "" Then
    FileNameToGet = App.Path & "\" & RemoteFileName
  End If
  hFich = FreeFile
  Recieved = 0
  Open FileNameToGet For Output As #hFich
  
  
  FtpTransfertData = ""
  Ftp_RETR = False
  If ftpData.State <> 0 Then
    SendProtocol "RETR " & RemoteFileName, False
    If FtpCode <> 150 And FtpCode <> 125 Then
      Close #hFich
      Ftp_RETR = False
      Exit Function
    End If
    FtpCode = 0
    Do
      If FtpTransfertData <> "" Then
        Print #hFich, FtpTransfertData;
        Recieved = Recieved + Len(FtpTransfertData)
        RaiseEvent RecieveProgress(Recieved)
        FtpTransfertData = ""
      End If
      DoEvents
    Loop While FtpCode = 0
    'sécure
      If FtpTransfertData <> "" Then
        Print #hFich, FtpTransfertData;
        Recieved = Recieved + Len(FtpTransfertData)
        RaiseEvent RecieveProgress(Recieved)
        FtpTransfertData = ""
      End If
    '
    If FtpCode <> 226 Then
      Ftp_RETR = False
    Else
      Ftp_RETR = True
    End If
  End If
  Close #hFich
End Function
Public Function Ftp_RMDIR(RemoteDirectory As String)
  'REMOVE directory
  Ftp_RMD RemoteDirectory
End Function
Public Function Ftp_RMD(RemoteDirectory As String) As Boolean
  'remove directory
  SendProtocol "RMD " & RemoteDirectory
  If FtpCode <> 250 Then
    Ftp_RMD = False
  Else
    Ftp_RMD = True
  End If
End Function
Public Function Ftp_RNFR(RemoteFileName As String) As Boolean
  'restart position for retr
  SendProtocol "RNFR " & RemoteFileName
  If FtpCode <> 350 Then
    Ftp_RNFR = False
  Else
    Ftp_RNFR = True
  End If
End Function

Public Function Ftp_RNTO(RemoteFileName As String) As Boolean
  'restart position for retr
  SendProtocol "RNTO " & RemoteFileName
  If FtpCode <> 250 Then
    Ftp_RNTO = False
  Else
    Ftp_RNTO = True
  End If
End Function
Public Function Ftp_SHEL(RemoteProg As String, Optional WindowsStyle As VbAppWinStyle = vbNormalFocus) As Boolean
  SendProtocol "SHEL " & RemoteProg & "," & CStr(WindowsStyle)
  If FtpCode <> 220 Then
    Ftp_SHEL = False
  Else
    Ftp_SHEL = True
  End If
End Function
Public Function Ftp_SIZE(RemoteFileName As String) As Boolean
  SendProtocol "SIZE " & RemoteFileName
  If FtpCode <> 213 Then
    Ftp_SIZE = False
  Else
    Ftp_SIZE = True
  End If
End Function
Public Function FTP_STAT() As Boolean
  SendProtocol "STAT"
  If FtpCode <> 221 Then
    FTP_STAT = False
  Else
    FTP_STAT = True
  End If
End Function
Public Function Ftp_STOR(RemoteFileName As String) As Boolean
  Dim iPosit As Integer
  Dim hFich As Integer
  Dim Buffer As String
  Dim BlockSize As Long
  Dim Rest As Long
  Dim FileSize As Long
  Dim Send As Long
  Dim Pcent As Integer
  If FileNameToSend = "" Then
    FileNameToSend = App.Path & "\" & RemoteFileName
  End If
  BlockSize = 4096
  Send = 0
  FileSize = FileLen(FileNameToSend)
  Rest = FileSize
  SendProtocol "STOR " & RemoteFileName
  If FtpCode <> 150 And FtpCode <> 125 Then
    Ftp_STOR = False
    Exit Function
  End If
  hFich = FreeFile
  Pcent = 0
  On Error GoTo ErrOpen
  Open FileNameToSend For Binary As #hFich
    Do While Rest > 0
      If BlockSize > Rest Then
        BlockSize = Rest
      End If
      Buffer = String(BlockSize, 0)
      Get #hFich, , Buffer
      SendOk = False
      ftpData.SendData Buffer
      'wait for send complete
      Do While Not SendOk
        DoEvents
      Loop
      Send = Send + BlockSize
      Pcent = Fix(Send * 100 / FileSize)
      RaiseEvent SendProgress(Pcent, Send)
      Rest = FileSize - Send
    Loop
  Ftp_STOR = True
Sortie:
  Close #hFich
  FtpCode = 0
  CloseFtpData
  Do While FtpCode <> 226
    DoEvents
  Loop
    
  DoEvents
  FileNameToSend = ""
  Exit Function
ErrOpen:
  Resume Sortie
End Function
Public Function FTP_SYST() As Boolean
  SendProtocol "SYST"
  If FtpCode <> 215 Then
    FTP_SYST = False
  Else
    FTP_SYST = True
  End If
End Function
Public Function Ftp_TYPE(TransMode As String) As Boolean
  'change transfert mode to TransMode
  SendProtocol "TYPE " & TransMode
  If FtpCode <> 200 Then
    Ftp_TYPE = False
  Else
    Ftp_TYPE = True
  End If
End Function
Public Function Ftp_UNLOGPROG() As Boolean

  'fonction non suportée par un serveur FTP standard
  
  SendProtocol "UNLOGPROG", True
  
End Function
Public Function Ftp_USER(User As String) As Integer
  'Set User
  SendProtocol "USER " & User
  If FtpCode = 331 Then
    'server ask for password
    Ftp_USER = 1
  ElseIf FtpCode = 230 Then
    Ftp_USER = -1
  Else
    Ftp_USER = 0
  End If
End Function
'===========================================
' Private Methods
'===========================================
Private Function SendProtocol(Cmd As String, Optional WithTimeOut As Boolean = True) As Integer
  'Send an FTP command
  'Return :
  '   -1 il all is OK
  '   0 if connection not openned
  '   1 if time out
  
  Dim MaxDate As Date
  Dim TimeOut As Boolean
  SendProtocol = -1
  If ftpDialog.State = 0 Or ftpDialog.State = 9 Then
    SendProtocol = 0
    Exit Function
  End If
  TimeOut = False
  MaxDate = DateAdd("s", m_TimeOutDelay, Now)
  FtpCode = 0
  RaiseEvent ProtocolSend(Cmd)
  On Error GoTo GetError
  ftpDialog.SendData Cmd & vbCrLf
  On Error GoTo 0
  Do While FtpCode = 0 And Not TimeOut
    If WithTimeOut Then
      If Now > MaxDate Then
        TimeOut = True
        SendProtocol = 1
      End If
    End If
    DoEvents
  Loop
  Exit Function
GetError:
  MsgBox "Err n° " & CStr(Err) & vbCrLf & _
          Err.Description
End Function


Private Sub CloseFtpData()
  If ftpData.State <> 0 Then
      ftpData.Close
      'wait for close complete
      Do While ftpData.State <> 0
        DoEvents
      Loop
      
  End If

End Sub
