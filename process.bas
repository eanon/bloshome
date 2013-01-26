Attribute VB_Name = "basProcess"
'BlosHome (c) FFh Lab / Eric Lequien, 2009-2013 - http://ffh-lab.com
'This module contains the mechanism to list running processes or detect a specific one
'REF : comes from EB2MWFL 1.0 (c) FFh Lab, itself adapted from "Active processus" article in Microsoft KB
'***LATER : uniformize all to the same naming convention as the rest of BlosHome code

Option Explicit
Option Base 0

Private Function StrZToStr(s As String) As String
    StrZToStr = Left$(s, Len(s) - 1)
End Function

Private Function getVersion() As Long
    Dim osinfo As OSVERSIONINFO
    Dim retvalue As Integer
    
    osinfo.dwOSVersionInfoSize = 148
    osinfo.szCSDVersion = Space$(128)
    retvalue = GetVersionExA(osinfo)
    getVersion = osinfo.dwPlatformId
End Function

Public Function ListProcess() As String()
    'Returns a list of current active Windows processus (i.e. disk path of executables actually in memory)
    Dim nCurr As Integer
    ReDim arProcess(1) As String
    
    nCurr = 0
    
    Select Case getVersion()
        Case 1
            'Windows 9x
            Dim f As Long
            Dim sname As String
            Dim hSnap As Long
            Dim hNull As Long
            Dim proc As PROCESSENTRY32
            
            hSnap = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
            If hSnap = hNull Then Exit Function
            proc.dwSize = Len(proc)
            
            'iterates through the processes
            f = Process32First(hSnap, proc)
            Do While f
                sname = StrZToStr(proc.szExeFile)
                ReDim Preserve arProcess(nCurr)
                arProcess(nCurr) = sname
                nCurr = nCurr + 1
                f = Process32Next(hSnap, proc)
            Loop
        Case 2
            'Windows NT
            Dim cb As Long
            Dim cbNeeded As Long
            Dim NumElements As Long
            Dim ProcessIDs() As Long
            Dim cbNeeded2 As Long
            Dim Modules(1 To 200) As Long
            Dim lRet As Long
            Dim ModuleName As String
            Dim nSize As Long
            Dim hProcess As Long
            Dim I As Long
            
            'gets the array containing the process id's for each process object
            cb = 8
            cbNeeded = 96
            Do While cb <= cbNeeded
                cb = cb * 2
                ReDim ProcessIDs(cb / 4) As Long
                lRet = EnumProcesses(ProcessIDs(1), cb, cbNeeded)
            Loop
            NumElements = cbNeeded / 4
            
            For I = 1 To NumElements
                'gets a handle to the Process
                hProcess = OpenProcess(PROCESS_QUERY_INFORMATION _
                Or PROCESS_VM_READ, 0, ProcessIDs(I))
                
                'got a Process handle
                If hProcess <> 0 Then
                    'get an array of the module handles for the specified process
                    lRet = EnumProcessModules(hProcess, Modules(1), 200, cbNeeded2)
                    
                    'if the Module Array is retrieved, Get the ModuleFileName
                    If lRet <> 0 Then
                        ModuleName = Space(MAX_PATH)
                        nSize = 500
                        lRet = GetModuleFileNameExA(hProcess, Modules(1), ModuleName, nSize)
                        ReDim Preserve arProcess(nCurr)
                        arProcess(nCurr) = Left$(ModuleName, lRet)
                        nCurr = nCurr + 1
                    End If
                End If
                
                'closes the handle to the process
                lRet = CloseHandle(hProcess)
            Next
        End Select
        ListProcess = arProcess 'must be last instruction for speed (see annexe/Undocumented trick to speed up functions that return array)
End Function

Public Function HowManyRunning(strSample) As Integer
    'GOAL : indicates how many processes containing strName in its path is actually running
    'IN : a sample string which should case-insensitivelly appear in the filename or path of the running executable
    'OUT : the number of running process which contains strSample
    'REF : adapted from EB2MWFL/IsRunning() by FFh Lab
    'TIP : to limit to one instance of an app, just unload it if this fct return more than 1
    'USE : to known if "MailWasher.exe" is running, we could call IsRunning("mailwasher") or IsRunning("ailwa")
    '      .Also, we could add something from the path if we know it could exist several exes of same name on disk
    '      .We could add ".exe" too if it exist a dll, ocx, vbx or else with same prefix... All decision is open to you !
    Dim arProc() As String
    Dim nIdx As Integer
    Dim strCurr As String
    Dim nOccur As Integer

    HowManyRunning = 0
    arProc = ListProcess()
    
    For nIdx = LBound(arProc) To UBound(arProc)
        strCurr = LCase(arProc(nIdx))
        If InStr(1, strCurr, strSample, vbTextCompare) <> 0 Then
            nOccur = nOccur + 1
        End If
    Next nIdx
    
    HowManyRunning = nOccur
End Function
