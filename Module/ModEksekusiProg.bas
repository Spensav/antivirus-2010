Attribute VB_Name = "ModEksekusiProg"
      Public Declare Function Process32First Lib "kernel32" ( _
         ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long

      Public Declare Function Process32Next Lib "kernel32" ( _
         ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long

      Public Declare Function CloseHandle Lib "Kernel32.dll" _
         (ByVal Handle As Long) As Long

      Public Declare Function OpenProcess Lib "Kernel32.dll" _
        (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, _
            ByVal dwProcId As Long) As Long

      Public Declare Function EnumProcesses Lib "psapi.dll" _
         (ByRef lpidProcess As Long, ByVal cb As Long, _
            ByRef cbNeeded As Long) As Long

      Public Declare Function GetModuleFileNameExA Lib "psapi.dll" _
         (ByVal hProcess As Long, ByVal hModule As Long, _
            ByVal ModuleName As String, ByVal nSize As Long) As Long

      Public Declare Function EnumProcessModules Lib "psapi.dll" _
         (ByVal hProcess As Long, ByRef lphModule As Long, _
            ByVal cb As Long, ByRef cbNeeded As Long) As Long

      Public Declare Function CreateToolhelp32Snapshot Lib "kernel32" ( _
         ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long

      Public Declare Function GetVersionExA Lib "kernel32" _
         (lpVersionInformation As OSVERSIONINFO) As Integer

      Public Type PROCESSENTRY32
         dwSize As Long
         cntUsage As Long
         th32ProcessID As Long           ' This process
         th32DefaultHeapID As Long
         th32ModuleID As Long            ' Associated exe
         cntThreads As Long
         th32ParentProcessID As Long     ' This process's parent process
         pcPriClassBase As Long          ' Base priority of process threads
         dwFlags As Long
         szExeFile As String * 260       ' MAX_PATH
      End Type

      Public Type OSVERSIONINFO
         dwOSVersionInfoSize As Long
         dwMajorVersion As Long
         dwMinorVersion As Long
         dwBuildNumber As Long
         dwPlatformId As Long           '1 = Windows 95.
                                        '2 = Windows NT

         szCSDVersion As String * 128
      End Type

      Public Const PROCESS_QUERY_INFORMATION = 1024
      Public Const PROCESS_VM_READ = 16
      Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
      Public Const SYNCHRONIZE = &H100000
      'STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &HFFF
      Public Const PROCESS_ALL_ACCESS = &H1F0FFF
      Public Const TH32CS_SNAPPROCESS = &H2&
      Public Const hNull = 0

      Function StrZToStr(s As String) As String
         StrZToStr = Left$(s, Len(s) - 1)
      End Function

      Public Function getVersion() As Long
         Dim osinfo As OSVERSIONINFO
         Dim retvalue As Integer
         osinfo.dwOSVersionInfoSize = 148
         osinfo.szCSDVersion = Space$(128)
         retvalue = GetVersionExA(osinfo)
         getVersion = osinfo.dwPlatformId
      End Function


