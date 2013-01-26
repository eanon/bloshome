Attribute VB_Name = "basWinAPI"
'BlosHome (c) FFh Lab / Eric Lequien, 2009-2013 - http://ffh-lab.com
'This module gather all the necessary Windows API declarations

Option Explicit

'Global purpose
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'INI Management
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long

'Temporary directory/file management
Public Const MAX_PATH = 260
Public Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

'Error management (when useful, we'll provide Err.LastDllError in // to error indicated by VB6 itself)
Public Declare Function GetLastError Lib "kernel32" () As Long

'Processes detection
Public Declare Function Process32First Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Public Declare Function Process32Next Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Public Declare Function CloseHandle Lib "kernel32.dll" (ByVal Handle As Long) As Long
Public Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long
Public Declare Function EnumProcesses Lib "psapi.dll" (ByRef lpidProcess As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Public Declare Function GetModuleFileNameExA Lib "psapi.dll" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long
Public Declare Function EnumProcessModules Lib "psapi.dll" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Public Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Public Declare Function GetVersionExA Lib "kernel32" (lpVersionInformation As OSVERSIONINFO) As Integer

Public Type PROCESSENTRY32
   dwSize As Long
   cntUsage As Long
   th32ProcessID As Long            ' this process
   th32DefaultHeapID As Long
   th32ModuleID As Long             ' associated exe
   cntThreads As Long
   th32ParentProcessID As Long      ' this process's parent process
   pcPriClassBase As Long           ' base priority of process threads
   dwFlags As Long
   szExeFile As String * 260        ' MAX_PATH
End Type

Public Type OSVERSIONINFO
   dwOSVersionInfoSize As Long
   dwMajorVersion As Long
   dwMinorVersion As Long
   dwBuildNumber As Long
   dwPlatformId As Long             '1:Win95, 2:WinNT
   szCSDVersion As String * 128
End Type

Public Const PROCESS_QUERY_INFORMATION = 1024
Public Const PROCESS_VM_READ = 16
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Public Const SYNCHRONIZE = &H100000
Public Const PROCESS_ALL_ACCESS = &H1F0FFF
Public Const TH32CS_SNAPPROCESS = &H2&

'Icon resource management ; utile à SetIcon()
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Public Const SM_CXICON = 11
Public Const SM_CYICON = 12

Public Const SM_CXSMICON = 49
Public Const SM_CYSMICON = 50
   
Public Declare Function LoadImageAsString Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, _
                            ByVal lpsz As String, ByVal uType As Long, ByVal cxDesired As Long, _
                            ByVal cyDesired As Long, ByVal fuLoad As Long) As Long

Public Const LR_DEFAULTCOLOR = &H0
Public Const LR_MONOCHROME = &H1
Public Const LR_COLOR = &H2
Public Const LR_COPYRETURNORG = &H4
Public Const LR_COPYDELETEORG = &H8
Public Const LR_LOADFROMFILE = &H10
Public Const LR_LOADTRANSPARENT = &H20
Public Const LR_DEFAULTSIZE = &H40
Public Const LR_VGACOLOR = &H80
Public Const LR_LOADMAP3DCOLORS = &H1000
Public Const LR_CREATEDIBSECTION = &H2000
Public Const LR_COPYFROMRESOURCE = &H4000
Public Const LR_SHARED = &H8000&

Public Const IMAGE_ICON = 1

Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, _
                            ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Const WM_SETICON = &H80

Public Const ICON_SMALL = 0
Public Const ICON_BIG = 1

Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Const GW_OWNER = 4
