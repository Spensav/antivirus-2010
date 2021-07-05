VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsScanProc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32.dll" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Private Declare Function EnumProcesses Lib "psapi.dll" (ByRef lpidProcess As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Private Declare Function EnumProcessModules Lib "psapi.dll" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef lpcbNeeded As Long) As Long
Private Declare Function FormatMessage Lib "kernel32.dll" Alias "FormatMessageA" (ByVal dwFlags As Long, ByRef lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef Arguments As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32.dll" () As Long
Private Declare Function GetModuleBaseName Lib "psapi.dll" Alias "GetModuleBaseNameA" (ByVal hProcess As Long, ByVal hModule As Long, ByVal lpBaseName As String, ByVal nSize As Long) As Long
Private Declare Function GetModuleFileNameEx Lib "psapi.dll" Alias "GetModuleFileNameExA" (ByVal hProcess As Long, ByVal hModule As Long, ByVal lpFilename As String, ByVal nSize As Long) As Long
Private Declare Function GetQueueStatus Lib "user32.dll" (ByVal fuFlags As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32.dll" Alias "GetVersionExA" (ByRef lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32.dll" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function lstrlen Lib "kernel32.dll" Alias "lstrlenA" (ByVal lpString As String) As Long
Private Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, ByRef TokenHandle As Long) As Long
Private Declare Function Process32First Lib "kernel32.dll" (ByVal hSnapshot As Long, ByRef lppe As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "kernel32.dll" (ByVal hSnapshot As Long, ByRef lppe As PROCESSENTRY32) As Long

' Modified API Declaration
Private Declare Function AdjustTokenPrivileges Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, ByRef NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, ByVal PreviousState As Long, ByVal ReturnLength As Long) As Long
Private Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, ByRef lpLuid As LUID) As Long
Private Declare Function TerminateProcess32 Lib "kernel32.dll" Alias "TerminateProcess" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long

' API Constants
Private Const ANYSIZE_ARRAY                 As Long = 1
Private Const FORMAT_MESSAGE_FROM_SYSTEM    As Long = &H1000
Private Const MAX_DESCRIPTION               As Long = 1024
Private Const MAX_MODULE_NAME32             As Long = 255
Private Const MAX_PATH                      As Long = 260
Private Const PROCESS_QUERY_INFORMATION     As Long = (&H400)
Private Const PROCESS_VM_READ               As Long = (&H10)
Private Const SE_DEBUG_NAME                 As String = "SeDebugPrivilege"
Private Const SE_PRIVILEGE_ENABLED          As Long = &H2
Private Const STANDARD_RIGHTS_REQUIRED      As Long = &HF0000
Private Const SYNCHRONIZE                   As Long = &H100000
Private Const TH32CS_SNAPPROCESS            As Long = &H2
Private Const TOKEN_ADJUST_PRIVILEGES       As Long = &H20
Private Const TOKEN_QUERY                   As Long = &H8
Private Const PROCESS_ALL_ACCESS            As Long = (STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &HFFF)
Private Const VER_PLATFORM_WIN32_NT         As Long = 2

'GetQueueStatus Flag
Private Const QS_ALLEVENTS As Long = &HBF

' API Types
Private Type LARGE_INTEGER
    lowpart     As Long
    highpart    As Long
End Type

Private Type LUID
    lowpart     As Long
    highpart    As Long
End Type

Private Type LUID_AND_ATTRIBUTES
    pLuid       As LUID
    Attributes  As Long
End Type

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion      As Long
    dwMinorVersion      As Long
    dwBuildNumber       As Long
    dwPlatformId        As Long
    szCSDVersion        As String * 128 ' Maintenance string for PSS usage
End Type

Private Type PROCESSENTRY32
    dwSize              As Long
    cntUsage            As Long
    th32ProcessID       As Long
    th32DefaultHeapID   As Long
    th32ModuleID        As Long
    cntThreads          As Long
    th32ParentProcessID As Long
    pcPriClassBase      As Long
    dwFlags             As Long
    szExeFile           As String * MAX_PATH
End Type

Private Type TOKEN_PRIVILEGES
    PrivilegeCount              As Long
    Privileges(ANYSIZE_ARRAY)   As LUID_AND_ATTRIBUTES
End Type

' Variable Declarations
Private m_bCancel       As Boolean
Private m_bErrorOnFail  As Boolean
Private m_bOldMethod    As Boolean
Private m_bPaused       As Boolean
Private m_bProcModules  As Boolean
Private m_bScanning     As Boolean
Private m_bSystemProc   As Boolean
Private m_bTerminated   As Boolean
Private m_lNumProcess   As Long
Private m_lProcessID    As Long
Private m_lWinPlatform  As Long

' //-- Class Events --//

Public Event CurrentModule(Process As String, ID As Long, Module As String, File As String)
'Occurs everytime a process is scanned for modules.
Public Event CurrentProcess(Name As String, File As String, ID As Long, Modules As Long)
'Occurs everytime a process is scanned.
Public Event DoneScanning(TotalProcess As Long)
'Occures after scanning all processes.

#If False Then
    ' Trick to preserve casing of these variables when used in VB IDE
    Private Process, Module, File, Name, Path, ID, Modules, TotalProcess
#End If

' //-- Properties --//

Public Property Get ErrorOnFail() As Boolean
'Returns/sets whether to raise an error if process termination fails.
    ErrorOnFail = m_bErrorOnFail
End Property

Public Property Let ErrorOnFail(Value As Boolean)
    m_bErrorOnFail = Value
End Property

Public Property Get ForceOldMethod() As Boolean
'Force to use the older method of enumerating processes for newer Windows systems.
    ForceOldMethod = m_bOldMethod
End Property

Public Property Let ForceOldMethod(Value As Boolean)
    m_bOldMethod = Value
End Property

Public Property Get ProcessModules() As Boolean
'Scan for modules(dll,ocx,etc..) used by a process other than its main executable.
    ProcessModules = m_bProcModules
End Property

Public Property Let ProcessModules(Value As Boolean)
    m_bProcModules = Value
End Property

Public Property Get ProcessTerminated() As Boolean
'Returns the boolean result for a process terminated in the class event.
    ProcessTerminated = m_bTerminated
End Property

Public Property Get Scanning() As Boolean
'Returns True if class is currently on scanning state.
    Scanning = m_bScanning
End Property

Public Property Get SystemProcesses() As Boolean
'Returns/sets whether to include scanning for system processes.
    SystemProcesses = m_bSystemProc
End Property

Public Property Let SystemProcesses(Value As Boolean)
    m_bSystemProc = Value
End Property

Public Property Get TotalProcesses() As Integer
'Returns the current or the total number of processes scanned.
    TotalProcesses = m_lNumProcess
End Property

' //-- Procedures --//

Public Sub BeginScanning()
'Start scanning for running processes in the system.
    If (m_bPaused) Then
        ResumeScanning ' Resume scanning instead
        Exit Sub
    End If
   
    If (m_bScanning) Then
        Exit Sub ' avoid cascading scans
    End If
   
    m_bCancel = False
    m_lNumProcess = 0
   
    m_bScanning = True
    ScanForProcesses ' scan the system for running processes
    m_bScanning = False
   
    m_lProcessID = -1
    RaiseEvent DoneScanning(m_lNumProcess)
End Sub

Public Sub CancelScanning()
'Abort scanning for running processes.
    m_bCancel = True
    ResumeScanning ' Resume if scanning has been paused
End Sub

Public Sub PauseScanning()
'Temporarily stop scanning process.
    If (m_bScanning) Then
        m_bPaused = True
    End If
End Sub

Public Sub ResumeScanning()
'Resume paused scanning process.
    If (m_bPaused) Then
        m_bPaused = False
    End If
End Sub

Public Function TerminateProcess(Optional lProcessID As Long = -1) As Boolean
'Terminate a running process using the specified process ID.
    If (lProcessID = -1) Then
        lProcessID = m_lProcessID ' Get process ID of currently scanned process
       
        ' If process ID is not given, it will attempt to terminate the current process
        If (lProcessID = -1) Then
            Exit Function
        End If
    End If
   
    Dim p_lProcess      As Long
    Dim p_lToken        As Long
    Dim p_tPrivileges   As TOKEN_PRIVILEGES
   
    ' Windows NT/2000 requires special treatment to ensure that the
    ' calling process has enough privileges to perform the instruction.
   
    If (m_lWinPlatform = VER_PLATFORM_WIN32_NT) Then
        ' Open token of the defined process
        If (OpenProcessToken(GetCurrentProcess, TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, p_lToken) = 0) Then
            GoTo End_Function
        End If
       
        ' Get LUID used to locally represent the specified privilege name
        If (LookupPrivilegeValue("", SE_DEBUG_NAME, p_tPrivileges.Privileges(ANYSIZE_ARRAY).pLuid) = 0) Then
            GoTo End_Function
        End If
       
        p_tPrivileges.PrivilegeCount = 1
        p_tPrivileges.Privileges(ANYSIZE_ARRAY).Attributes = SE_PRIVILEGE_ENABLED
       
        ' Attempt to acquire debug privilege for the process
        If (AdjustTokenPrivileges(p_lToken, 0&, p_tPrivileges, 0&, 0&, 0&) = 0) Then
            GoTo End_Function
        End If
    End If
   
    ' Finally, open the defined process
    p_lProcess = OpenProcess(PROCESS_ALL_ACCESS, 0&, lProcessID)
   
    If (p_lProcess) Then
        ' Attempt to terminate the process
        TerminateProcess = (TerminateProcess32(p_lProcess, 0&) <> 0)
        CloseHandle p_lProcess
       
        If (Not TerminateProcess) Then
            If (m_bErrorOnFail) Then
                err.Raise err.LastDllError, , ErrorDescription(err.LastDllError)
            End If
        End If
       
        If (GetQueueStatus(QS_ALLEVENTS)) Then ' check for events
            DoEvents
        End If
    End If
   
    If (m_lWinPlatform = VER_PLATFORM_WIN32_NT) Then
        ' Restore original privilege
        p_tPrivileges.Privileges(ANYSIZE_ARRAY).Attributes = 0
        AdjustTokenPrivileges p_lToken, 0&, p_tPrivileges, 0&, 0&, 0&
       
End_Function:
        If (p_lToken) Then
            CloseHandle p_lToken
        End If
    End If
End Function

' //-- Private Procedures --//

Private Sub DeepProcessScan(ID As Long, Name As String)
    If (m_bCancel Or m_lWinPlatform <> VER_PLATFORM_WIN32_NT) Then
        If (Not m_bCancel) Then
            m_lNumProcess = m_lNumProcess + 1
            RaiseEvent CurrentProcess(Name, "", ID, 1)
        End If
       
        Exit Sub ' Abort the whole scanning process
    End If
   
    Dim i                       As Long
    Dim p_lModuleCount          As Long
    Dim p_lModules(1 To 1024)   As Long
    Dim p_lNeeded               As Long
    Dim p_lProcess              As Long
    Dim p_lRet                  As Long
    Dim p_sBaseName             As String
    Dim p_sBuffer               As String
    Dim p_sProcessPath          As String
   
    p_lProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0&, ID)
   
    If (p_lProcess) Then
        p_lRet = EnumProcessModules(p_lProcess, p_lModules(1), 1024 * 4, p_lNeeded)
       
        If (p_lRet) Then
            p_lModuleCount = p_lNeeded \ 4
           
            For i = 1 To p_lModuleCount
                If (GetQueueStatus(QS_ALLEVENTS)) Then ' check for events
                    DoEvents ' processes system events
                End If
               
                If (m_bPaused) Then
                    m_bScanning = False
                    While (m_bPaused)
                        DoEvents ' Don't execute next instructions until resumed
                    Wend
                    m_bScanning = True
                End If
               
                If (m_bCancel) Then
                    Exit For ' attempt to abort the whole scanning process
                End If
               
                p_sBuffer = String$(MAX_MODULE_NAME32, 0)
                p_lRet = GetModuleBaseName(p_lProcess, p_lModules(i), p_sBuffer, MAX_MODULE_NAME32)
               
                If (p_lRet > 0) Then
                    p_sBaseName = Left$(p_sBuffer, p_lRet)
                End If
               
                p_sBuffer = String$(MAX_MODULE_NAME32, 0)
                p_lRet = GetModuleFileNameEx(p_lProcess, p_lModules(i), p_sBuffer, MAX_MODULE_NAME32)
               
                If (p_lRet > 0) Then
                    p_sProcessPath = Left$(p_sBuffer, p_lRet)
                End If
               
                ValidatePath p_sProcessPath
               
                If (i = 1) Then
                    If (Len(p_sProcessPath) = 0) Then
                        ' Consider as a system process if file path is empty
                        p_sProcessPath = "SYSTEM"
                       
                        If (Not m_bSystemProc) Then
                            Exit For ' Excluding system processes
                        End If
                    End If
                   
                    m_lNumProcess = m_lNumProcess + 1
                    RaiseEvent CurrentProcess(p_sBaseName, p_sProcessPath, ID, p_lModuleCount)
                   
                    If (Not m_bProcModules) Or (p_lModuleCount = 0) Then
                        Exit For ' Dont scan preceding process modules
                    End If
                Else
                    RaiseEvent CurrentModule(Name, ID, p_sBaseName, p_sProcessPath)
                End If
            Next
           
            CloseHandle p_lProcess
            Exit Sub
        Else
            If (Len(Name) = 0) And (ID) Then
                Name = "System" ' Assume this is a system process
            End If
        End If
       
        CloseHandle p_lProcess
    End If
   
    If (m_bSystemProc = True) Then
        If (Len(Name) = 0) And (ID) Then
            Name = "[System Process]" ' a.k.a "System Idle Process"
        End If
       
        m_lNumProcess = m_lNumProcess + 1
        RaiseEvent CurrentProcess(Name, "SYSTEM", ID, 1)
    End If
End Sub

Private Function ErrorDescription(nError As Long) As String
    Dim p_lLen As Long
    Dim p_sBuffer As String * MAX_DESCRIPTION
   
    p_lLen = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, _
                           ByVal 0&, _
                           nError, _
                           0&, _
                           p_sBuffer, _
                           MAX_DESCRIPTION, _
                           0&)
    If (p_lLen > 0) Then
        ErrorDescription = Left$(p_sBuffer, p_lLen)
    End If
End Function

Private Sub ScanForProcesses()
    Dim p_eProcessEntry As PROCESSENTRY32
    Dim p_lNeeded       As Long
    Dim p_lProcess      As Long
    Dim p_lProcesses()  As Long
    Dim p_lProcessID    As Long
    Dim p_lSnapshot     As Long
    Dim p_sExeFile      As String
   
    ' Windows 2000/ME/XP or later
    If (m_lWinPlatform = VER_PLATFORM_WIN32_NT) And (Not m_bOldMethod) Then
        ReDim p_lProcesses(1 To 1024) As Long
       
        If (EnumProcesses(p_lProcesses(1), 1024 * 4, p_lNeeded)) Then
            p_lNeeded = p_lNeeded \ 4
           
            For p_lProcess = 1 To p_lNeeded
                If (m_bCancel) Then
                    Exit For
                End If
               
                DeepProcessScan p_lProcesses(p_lProcess), ""
            Next
           
            Exit Sub
        Else
            ' Raise an error
            err.Raise err.LastDllError, , ErrorDescription(err.LastDllError)
            ' Then attempt to enumerate processes using the other way below
        End If
    End If
   
    ' Windows 95/98 (Old Method)
    p_lSnapshot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0&)
   
    If (p_lSnapshot) Then
        p_eProcessEntry.dwSize = Len(p_eProcessEntry)
        p_lProcess = Process32First(p_lSnapshot, p_eProcessEntry)
       
        Do While (p_lProcess) ' iterate through all processes
            If (GetQueueStatus(QS_ALLEVENTS)) Then ' check for events
                DoEvents ' processes system events
            End If
           
            If (m_bPaused) Then
                m_bScanning = False
                While (m_bPaused)
                    DoEvents ' Don't execute next instructions until resumed
                Wend
                m_bScanning = True
            End If
           
            If (m_bCancel) Then
                Exit Do ' abort the whole scanning process
            End If
           
            p_lProcessID = p_eProcessEntry.th32ProcessID
            m_lProcessID = p_lProcessID
           
            p_sExeFile = p_eProcessEntry.szExeFile
            p_sExeFile = Left$(p_sExeFile, lstrlen(p_sExeFile))
           
            ' We need to get process name from path, because process name
            ' from PROCESSENTRY32 type is not complete for long filenames.
            DeepProcessScan p_lProcessID, p_sExeFile
Next_Process:
            p_eProcessEntry.szExeFile = String$(MAX_PATH, 0)
            p_lProcess = Process32Next(p_lSnapshot, p_eProcessEntry)
        Loop
    End If
End Sub

Private Sub ValidatePath(ByRef Path As String)
     ' UNC File names
    If (InStr(1, Path, "\?\UNC\", vbTextCompare)) Then
        Path = Replace$(Path, "\?\UNC\", "", 1, 1)
    End If
   
    ' \\?\ tells Windows to turn off File parsing
    If (InStr(1, Path, "\??\", vbTextCompare)) Then
        Path = Replace$(Path, "\??\", "", 1, 1)
    End If
   
     ' Only the first instances will be replaced
    If (InStr(1, Path, "\SystemRoot\", vbTextCompare)) Then
        Path = Replace$(Path, "\SystemRoot\", WindowsDirectory, 1, 1)
    End If
End Sub

Private Function WindowsDirectory()
    Dim p_lLen      As Long
    Dim p_sBuffer   As String * MAX_PATH
   
    p_lLen = GetWindowsDirectory(p_sBuffer, MAX_PATH)
    If (p_lLen > 0) Then
        WindowsDirectory = Left$(p_sBuffer, p_lLen)
    End If
   
    If (WindowsDirectory Like "*\") Then
        ' Just do nothing
    Else
        WindowsDirectory = WindowsDirectory & "\"
    End If
End Function

Private Function WindowsPlatform() As Long
    Dim p_tOSInfo As OSVERSIONINFO
   
    p_tOSInfo.dwOSVersionInfoSize = Len(p_tOSInfo)
    GetVersionEx p_tOSInfo
    WindowsPlatform = p_tOSInfo.dwPlatformId
End Function

' //-- Class Procedures --//

Private Sub Class_Initialize()
    ' unless these properties are set, these would be their default values
    m_bSystemProc = False
    m_lProcessID = -1
    m_lWinPlatform = WindowsPlatform
End Sub

Private Sub Class_Terminate()
    If (m_bScanning Or m_bPaused) Then
        CancelScanning
    End If
End Sub

' Created by Noel A. Dacara | Copyright © 2003-2005 Davao City, Philippines

'######Selesai Disini######

