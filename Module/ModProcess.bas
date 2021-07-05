Attribute VB_Name = "ModProcess"
Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Private Const TH32CS_SNAPHEAPLIST = &H1
Private Const TH32CS_SNAPPROCESS = &H2
Private Const TH32CS_SNAPTHREAD = &H4
Private Const TH32CS_SNAPMODULE = &H8
Private Const TH32CS_SNAPALL = (TH32CS_SNAPHEAPLIST Or TH32CS_SNAPPROCESS Or TH32CS_SNAPTHREAD Or TH32CS_SNAPMODULE)
Private Const TH32CS_INHERIT = &H80000000
Private Const MAX_PATH As Integer = 260
Private Type PROCESSENTRY32
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
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Private Declare Function Process32First Lib "kernel32" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32.dll" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hHandle As Long) As Long
Private Const PROCESS_ALL_ACCESS = &H1F0FFF
'Enum the path
Private Const PROCESS_QUERY_INFORMATION As Long = &H400
Private Const PROCESS_VM_READ = &H10
Private Declare Function EnumProcessModules Lib "psapi.dll" ( _
    ByVal hProcess As Long, _
    ByRef lphModule As Long, _
    ByVal cb As Long, _
    ByRef cbNeeded As Long) As Long
Private Declare Function GetModuleFileNameExA Lib "psapi.dll" ( _
    ByVal hProcess As Long, _
    ByVal hModule As Long, _
    ByVal ModuleName As String, _
    ByVal nSize As Long) As Long
 Private ProcessID(100) As Long
 Private path(100) As String
 Private jmlProcess As Integer
Private Declare Function MoveFile Lib "kernel32" Alias "MoveFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Public Sub Bunuh(NamaFile As String)
'procedure ini berfungsi untuk menghentikan proses dari sebuah program
Dim a As Long
For a = 1 To jmlProcess
    If path(a) = NamaFile Then
    TerminateProcess OpenProcess(PROCESS_ALL_ACCESS, 1, ProcessID(a)), 0
    Exit For
    Call List_Process
    End If
Next a
End Sub
'fungsi dibawah ini untuk mendapatkan program-program apa yang sedang dalam proses
Public Sub List_Process()
    jmlProcess = 1
    Dim hSnapShot As Long, uProcess As PROCESSENTRY32
        hSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPALL, 0&)
        'Mendapatkan informasi tentang semua proses yang sedang dijalankan
    uProcess.dwSize = Len(uProcess)
    r = Process32First(hSnapShot, uProcess)
        'Mendapatkan informasi tentang proses yang pertama
    Do While r
        'perulangan selama r <> 0
        
        'List1.AddItem Left$(uProcess.szExeFile, InStr(1, uProcess.szExeFile, Chr$(0), vbTextCompare) - 1)
            'Memasukkan nama aplikasi pada List1
        ProcessID(jmlProcess) = uProcess.th32ProcessID
        path(jmlProcess) = PathByPID(ProcessID(jmlProcess))
            'Memasukkan Process ID untuk masing-masing aplikasi
        r = Process32Next(hSnapShot, uProcess)
            'Mendapatkan informasi dari proses selanjutnya pada windows
    jmlProcess = jmlProcess + 1
    Loop
    jmlProcess = jmlProcess - 1
    CloseHandle hSnapShot
End Sub
Public Function PathByPID(pid As Long) As String
'Fungsi dibawah ini berfungsi untuk mencari path atau lokasi dari
'program yang sedang berjalan
'Kode ini dapat dilihat di :
'http://support.microsoft.com/default.aspx?scid=kb;en-us;187913
    Dim cbNeeded As Long
    Dim Modules(1 To 200) As Long
    Dim ret As Long
    Dim ModuleName As String
    Dim nSize As Long
    Dim hProcess As Long
    
    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION _
        Or PROCESS_VM_READ, 0, pid)
    
    If hProcess <> 0 Then
        
        ret = EnumProcessModules(hProcess, Modules(1), _
            200, cbNeeded)
        
        If ret <> 0 Then
            ModuleName = Space(MAX_PATH)
            nSize = 500
            ret = GetModuleFileNameExA(hProcess, _
                Modules(1), ModuleName, nSize)
            PathByPID = Left(ModuleName, ret)
        End If
    End If
    
    ret = CloseHandle(hProcess)
    
    If PathByPID = "" Then
        PathByPID = ""
    End If
    
    If Left(PathByPID, 4) = "\??\" Then
        PathByPID = ""
    End If
    

    If Left(PathByPID, 12) = "\SystemRoot\" Then
        PathByPID = ""
    End If
End Function
