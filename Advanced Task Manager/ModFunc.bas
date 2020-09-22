Attribute VB_Name = "ModFunc"
'Advanced Task Manager
'By Alper ESKIKILIC
'odesayazilim@gmail.com
'www.odesayazilim.com



Public Type PROCESS_MEMORY_COUNTERS
   cb As Long
   PageFaultCount As Long
   PeakWorkingSetSize As Long
   WorkingSetSize As Long
   QuotaPeakPagedPoolUsage As Long
   QuotaPagedPoolUsage As Long
   QuotaPeakNonPagedPoolUsage As Long
   QuotaNonPagedPoolUsage As Long
   PagefileUsage As Long
   PeakPagefileUsage As Long
End Type

Public Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Public Type MODULEINFO
   lpBaseOfDLL As Long
   SizeOfImage As Long
   EntryPoint As Long
End Type

Public Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

Public Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long
Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Public Declare Function CreateToolHelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Public Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Public Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function GetModuleFileNameExA Lib "PSAPI.DLL" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long
Public Declare Function EnumProcessModules Lib "PSAPI.DLL" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Public Declare Function GetModuleInformation Lib "PSAPI.DLL" (ByVal hProcess As Long, ByVal hModule As Long, lpmodinfo As MODULEINFO, ByVal cb As Long) As Long
Public Declare Function GetProcessMemoryInfo Lib "PSAPI.DLL" (ByVal hProcess As Long, ppsmemCounters As PROCESS_MEMORY_COUNTERS, ByVal cb As Long) As Long
Public Declare Function GetProcessTimes Lib "kernel32" (ByVal hProcess As Long, lpCreationTime As FILETIME, lpExitTime As FILETIME, lpKernelTime As FILETIME, lpUserTime As FILETIME) As Long
Public Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Public Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long

Public Const PROCESS_VM_READ = &H10
Public Const PROCESS_QUERY_INFORMATION = &H400
Public Const PROCESS_ALL_ACCESS As Long = &H1F0FFF
Public Const MAX_PATH As Integer = 260
Public Type PROCESSENTRY32
       dwSize As Long
       cntUsage As Long
       th32ProcessID As Long
       th32DefaultHeapID As Long
       th32ModuleID As Long
       cntThreads As Long
       th32ParentProcessID As Long
       pcPriClassBase As Long
       dwFlags As Long
       szExeFile As String * MAX_PATH
End Type

Public i%

Public Function GetModules(PID As Long, strName As String) As String
Dim lOpen As Long, hModules As Long, sModules(1 To 1000) As Long, cbNeeded As Long, lFileName As Long, sName As String
Dim NameLen As String, sImgSize As String, MODINFO As MODULEINFO, lModInfo As Long


frmModules.ListView1.ListItems.Clear
lOpen = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, PID)

hModules = EnumProcessModules(lOpen, sModules(1), 1024 * 4, cbNeeded)

hModules = cbNeeded

For i = 1 To hModules - 1
sName = Space(MAX_PATH)
    lFileName = GetModuleFileNameExA(lOpen, sModules(i), sName, MAX_PATH)
        NameLen = Len(strName)
            lModInfo = GetModuleInformation(lOpen, sModules(i), MODINFO, Len(MODINFO))
                If Left(sName, NameLen) <> strName Then
                    sImgSize = FileSize(MODINFO.SizeOfImage)
                With frmModules.ListView1.ListItems.Add(, , sName)
                .SubItems(1) = sImgSize
                .SubItems(2) = FileDateTime(sName)
            End With
        End If
    Next i
    
Call CloseHandle(lOpen)
Call CloseHandle(hModules)

End Function

Public Function ProcessLoad()
Dim hSnapshot As Long, lNext As Long, PID As Long, szExename As String, uProcess As PROCESSENTRY32
Dim lOpen As Long, lName As Long, sBuff As String, sExePath As String, SYSTIME As SYSTEMTIME
Dim PROMEM As PROCESS_MEMORY_COUNTERS, lProMem As Long, sUsage As String, lProTime As Long, FT As FILETIME, sTime As Long, tExit As FILETIME, tUser As FILETIME, tKernel As FILETIME
Dim Threads As Integer

hSnapshot = CreateToolHelpSnapshot(2&, 0&)
If hSnapshot = 0 Then Exit Function
uProcess.dwSize = Len(uProcess)
lNext = ProcessFirst(hSnapshot, uProcess)

Do While lNext
    i = InStr(1, uProcess.szExeFile, Chr(0))
    szExename = LCase$(Left$(uProcess.szExeFile, i - 1))
    PID = uProcess.th32ProcessID
    Threads = uProcess.cntThreads
    lOpen = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, PID)
    sBuff = Space(MAX_PATH)
    lName = GetModuleFileNameExA(lOpen, 0, sBuff, MAX_PATH)
    sExePath = Left(sBuff, lName)
    lProMem = GetProcessMemoryInfo(lOpen, PROMEM, Len(PROMEM))
    sUsage = FileSize(PROMEM.WorkingSetSize)
    lProTime = GetProcessTimes(lOpen, FT, tExit, tKernel, tUser)
    FileTimeToLocalFileTime FT, FT
    FileTimeToSystemTime FT, SYSTIME
    If sExePath = vbNullString Then sExePath = "System"
    With frmViewer.ListView1.ListItems.Add(, , szExename)
        .SubItems(1) = PID
        .SubItems(2) = sExePath
        .SubItems(3) = sUsage
        .SubItems(4) = Threads
        .SubItems(5) = SYSTIME.wHour & ":" & SYSTIME.wMinute & ":" & SYSTIME.wSecond
        .SubItems(6) = SYSTIME.wDay & "/" & SYSTIME.wMonth & "/" & SYSTIME.wYear
    End With
    lNext = ProcessNext(hSnapshot, uProcess)

Loop
    Call CloseHandle(hSnapshot)
    Call CloseHandle(lOpen)
End Function


Public Function FileSize(ByVal StrSize As String) As String
    If StrSize$ < 1024 Then
        FileSize = StrSize$ & " Bytes"
    ElseIf StrSize$ < 1048576 Then
        FileSize = Format(StrSize$ / 1024#, "###0.00") & " KB"
    ElseIf StrSize$ > 1048576 Then
        FileSize = Format(StrSize$ / 1024# ^ 2, "###0.00") & " MB"
    End If
End Function

