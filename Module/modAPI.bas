Attribute VB_Name = "modAPI"
Public Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesW" (ByVal lpFileName As Long, ByVal dwFileAttributes As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Public Declare Function PathIsDirectory Lib "shlwapi.dll" Alias "PathIsDirectoryA" (ByVal pszPath As String) As Long
Public Declare Function CopyFile Lib "Kernel32.dll" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Public Declare Function MoveFile Lib "Kernel32.dll" Alias "MoveFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long
Public Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Public Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long
Public Declare Function GetSystemDirectory Lib "Kernel32.dll" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetWindowsDirectory Lib "Kernel32.dll" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Public Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Private Declare Function RealizePalette Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Public Declare Function VirtualAlloc Lib "kernel32" (lpAddress As Any, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Public Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long
Public Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As Any) As Long
Public Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Public Declare Function SetEndOfFile Lib "kernel32" (ByVal hFile As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Public Declare Function GetVolumeInformationW Lib "kernel32" (ByVal pv_lpRootPathName As Long, ByVal pv_lpVolumeNameBuffer As Long, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal pv_lpFileSystemNameBuffer As Long, ByVal nFileSystemNameSize As Long) As Long
Public Declare Function GetDriveTypeW Lib "kernel32" (ByVal nDrive As Long) As Long
Public Declare Function GetLogicalDriveStringsW Lib "Kernel32.dll" (ByVal nBufferLength As Long, ByVal lpBuffer As Long) As Long
Public Const DRIVE_REMOVABLE As Long = 2
Public mLastDrives As String
Public detik As pewaktu, menit As pewaktu, jam As pewaktu
Public Type pewaktu
    i As Integer
    s As String
End Type

Const ERROR_SUCCESS = 0
Const REG_SZ = 1
Const REG_DWORD = 4
Const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_CLASSES_ROOT = &H80000000
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_USERS = &H80000003
Const HKEY_CURRENT_CONFIG = &H80000005
Const HKEY_PERFORMANCE_DATA = &H80000004
Const HKEY_DYN_DATA = &H80000006
Public Const GENERIC_READ = &H80000000
Public Const GENERIC_WRITE = &H40000000
Public Const FILE_SHARE_READ = &H1
Public Const OPEN_EXISTING = 3
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const INVALID_HANDLE_VALUE = -1
Public Const FILE_END = 2
Public Const FILE_BEGIN = 0
Public Const FILE_CURRENT = 1
Public Const LWA_COLORKEY = &H1
Public Const GWL_EXSTYLE = (-20)
Public Const WS_EX_LAYERED = &H80000
Public Const MAX_PATH = 260
Public Const SW_SHOWNORMAL = 1

Public Type FILETIME
    dwLowDateTime       As Long
    dwHighDateTime      As Long
End Type

Public Type WIN32_FIND_DATA
    dwFileAttributes    As Long
    ftCreationTime      As FILETIME
    ftLastAccessTime    As FILETIME
    ftLastWriteTime     As FILETIME
    nFileSizeHigh       As Long
    nFileSizeLow        As Long
    dwReserved0         As Long
    dwReserved1         As Long
    cFileName           As String * MAX_PATH
    cAlternate          As String * 14
End Type

Type BROWSEINFO
     hOwner As Long
     pidlRoot As Long
     pszDisplayName As String
     lpszTitle As String
     ulFlags As Long
     lpfn As Long
     lParam As Long
     iImage As Long
End Type
'#SysTray=============================================================================
Public Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Public Const NIF_ICON As Long = &H2
Public Const NIF_MESSAGE As Long = &H1
Public Const NIF_TIP As Long = &H4
Public Const NIM_ADD As Long = &H0
Public Const NIM_DELETE As Long = &H2
Public Const NIM_MODIFY As Long = &H1
Public Const NIF_INFO As Long = &H10
Public Const WM_MOUSEMOVE As Long = &H200
Public Const WM_LBUTTONDBLCLK As Long = &H203
Public Const WM_LBUTTONDOWN As Long = &H201
Public Const WM_LBUTTONUP As Long = &H202
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_RBUTTONDOWN As Long = &H204
Public Const WM_RBUTTONUP As Long = &H205
Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 128
    dwState As Long
    dwStateMask As Long
    szInfo As String * 256
    uTimeout As Long
    szInfoTitle As String * 64
    dwInfoFlags As Long
End Type
Public Enum TypeBallon
    NIIF_NONE = &H0
    NIIF_WARNING = &H2
    NIIF_ERROR = &H3
    NIIF_INFO = &H1
    NIIF_GUID = &H4
End Enum
Public Function DriveLabel(ByVal sDrive As String) As String
    Dim sDriveName          As String
    Dim nDriveNameLen       As Long
        nDriveNameLen = 128
        sDriveName = String$(nDriveNameLen, 0)
    sDrive = Left$(sDrive, 1) & ":\"
    If GetVolumeInformationW(StrPtr(sDrive), StrPtr(sDriveName), nDriveNameLen, ByVal 0, ByVal 0, ByVal 0, 0, 0) Then
        DriveLabel = Left$(sDriveName, InStr(1, sDriveName, ChrW$(0)) - 1)
    Else
        DriveLabel = vbNullString
    End If
    If Len(DriveLabel) > 0 Then Exit Function
    Select Case GetDriveType(sDrive)
    Case 2: DriveLabel = "Removable Disk "
    Case 3: DriveLabel = "Local Disk"
    Case 5: DriveLabel = "CD/DVD-Drive"
    End Select
End Function
Public Sub FormOnTop(ByVal Frm As Form, ByVal State As Boolean)
SetWindowPos Frm.hwnd, IIf(State = True, -1, -2), 0, 0, 0, 0, &H1 Or &H2
End Sub
'########## BrowseForFolder ##########'
Public Function BrowseFolder(ByVal aTitle As String, ByVal aForm As Form) As String
Dim bInfo As BROWSEINFO
Dim rtn&, pidl&, path$, POS%
Dim BrowsePath As String
bInfo.hOwner = aForm.hwnd
bInfo.lpszTitle = aTitle
bInfo.ulFlags = &H1
pidl& = SHBrowseForFolder(bInfo)
path = Space(512)
t = SHGetPathFromIDList(ByVal pidl&, ByVal path)
POS% = InStr(path$, Chr$(0))
BrowseFolder = Left(path$, POS - 1)
If Right$(Browse, 1) = "\" Then
    BrowseFolder = BrowseFolder
    Else
    BrowseFolder = BrowseFolder + "\"
End If
If Right(BrowseFolder, 2) = "\\" Then BrowseFolder = Left(BrowseFolder, Len(BrowseFolder) - 1)
If BrowseFolder = "\" Then BrowseFolder = ""
End Function

'########## Yang ini aku gag tahu..... :) ##########'
Public Function StripNulls(ByVal OriginalStr As String) As String
    If (InStr(OriginalStr, Chr$(0)) > 0) Then
        OriginalStr = Left$(OriginalStr, InStr(OriginalStr, Chr(0)) - 1)
    End If
    StripNulls = OriginalStr
End Function

'########## fungsi untuk menentukan file script atau bukan ##########'
Public Function IsScript(Filename As String) As Boolean
IsScript = False
ext = Split("|vbs|vbe", "|")
For i = 1 To UBound(ext)
If LCase(Right(Filename, 3)) = LCase(ext(i)) Then IsScript = True
Next
End Function
Function RESFile()
frmMain.Image5.Picture = LoadResPicture(101, vbResBitmap)
frmMain.Image6.Picture = LoadResPicture(101, vbResBitmap)
frmMain.Image7.Picture = LoadResPicture(106, vbResBitmap)
If frmMain.SettingPelindung(0).Checked = True Then
frmMain.Image8.Picture = LoadResPicture(107, vbResBitmap) 'secure
Else
frmMain.Image8.Picture = LoadResPicture(108, vbResBitmap) 'not secure
End If
frmMain.Image9.Picture = LoadResPicture(102, vbResBitmap)
frmMain.Image10.Picture = LoadResPicture(103, vbResBitmap)
End Function

'\\\\\\\\\\\\\\\----------UNTUK STARTUP MANAGER----------//////////////////////
Private Sub getVal(START As Key)
'Doan code nay co nguon goc tu PSC
Dim Cnt As Long, Buf As String, Buf2 As String, retdata As Long, typ As Long
    'List1.Clear
    Dim KeyName As String
    Dim KeyPath As String
    Buf = Space(BUFFER_SIZE)
    Buf2 = Space(BUFFER_SIZE)
    ret = BUFFER_SIZE
    retdata = BUFFER_SIZE
    Cnt = 0
    RegOpenKeyEx START, Pathkey, 0, KEY_ALL_ACCESS, Result
    While RegEnumValue(Result, Cnt, Buf, ret, 0, typ, ByVal Buf2, retdata) <> ERROR_NO_MORE_ITEMS
        If typ = REG_DWORD Then
            KeyName = Left(Buf, ret)
            KeyPath = ChuoiGiaTri(Left(Asc(Buf2), retdata - 1))
        Else
            KeyName = Left(Buf, ret)
            KeyPath = ChuoiGiaTri(Left(Buf2, retdata - 1))
        End If
        
            Pic.Cls
            GetIcon KeyPath, frmMain.Pic
            ima.ListImages.Add LV.ListItems.Count + 1, , Pic.Image
            Dim lsv As ListItem
                  Set lsv = LV.ListItems.Add(, , KeyName, , LV.ListItems.Count + 1)
                  lsv.SubItems(1) = KeyPath
                  
        Cnt = Cnt + 1
        Buf = Space(BUFFER_SIZE)
        Buf2 = Space(BUFFER_SIZE)
        ret = BUFFER_SIZE
        retdata = BUFFER_SIZE
    Wend
    RegCloseKey Result
End Sub
Private Sub getValSta(START As Key)
'Sorry cac ban nhe, Dung Coi nhac qua nen copy y chang lai doan code nay
Dim Cnt As Long, Buf As String, Buf2 As String, retdata As Long, typ As Long
    'List1.Clear
    Dim KeyName As String
    Dim KeyPath As String
    Buf = Space(BUFFER_SIZE)
    Buf2 = Space(BUFFER_SIZE)
    ret = BUFFER_SIZE
    retdata = BUFFER_SIZE
    Cnt = 0
    RegOpenKeyEx START, Pathkey, 0, KEY_ALL_ACCESS, Result
    While RegEnumValue(Result, Cnt, Buf, ret, 0, typ, ByVal Buf2, retdata) <> ERROR_NO_MORE_ITEMS
            If typ = REG_DWORD Then
                KeyName = Left(Buf, ret)
                If Trim(Buf2) <> "" Then KeyPath = ChuoiGiaTri(Left(Asc(Buf2), retdata - 1))
            Else
                            'Debug.Print Buf
                KeyName = Left(Buf, ret)
                If Trim(Buf2) <> "" Then KeyPath = ChuoiGiaTri(Left(Buf2, retdata - 1))
            End If
                Dim lsv As ListItem
                      Set lsv = frmMain.LVV.ListItems.Add()
                      lsv.Text = KeyName
                      lsv.SubItems(1) = KeyPath
                        If START = aa Then
                            lsv.SubItems(2) = "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
                        ElseIf START = BB Then
                            lsv.SubItems(2) = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
                        End If
                            'lsv.Checked = True
        Cnt = Cnt + 1
        Buf = Space(BUFFER_SIZE)
        Buf2 = Space(BUFFER_SIZE)
        ret = BUFFER_SIZE
        retdata = BUFFER_SIZE
    Wend
    RegCloseKey Result
End Sub

Public Sub GetStartup()
    ThietLap frmMain.LVV, frmMain.ima, frmMain.Pic
    
        GetKeySta HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion\Winlogon", "Userinit"
        'Xu ly key Explorer
        GetKeySta HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion\Winlogon", "Shell"
        'Xu ly cac key dac biet khac
        GetKeySta HKEY_CURRENT_USER, "Software\Microsoft\Windows NT\CurrentVersion\Windows", "Load"
        GetKeySta HKEY_CURRENT_USER, "Software\Microsoft\Windows NT\CurrentVersion\Windows", "Run"
        If (FileExists(frmMain.LVV.ListItems(1).SubItems(1)) = False) Or (FileExists(frmMain.LVV.ListItems(1).SubItems(1)) = False) Then
            'cmdRepair.Visible = True
        End If
    getValSta aa
    getValSta BB
    If frmMain.LVV.ListItems.Count <> 0 Then GetIcons frmMain.LVV, frmMain.ima, frmMain.Pic
End Sub
Private Sub GetKeySta(hKey As Key, kPath As String, kName As String)
On Error Resume Next
    Dim PathExp As String
    Dim t As Byte
    Dim t1 As Byte
    Dim lsv1 As ListItem
        Dim tmpStr As String
        PathExp = GetString1(hKey, kPath, kName)
        Do While InStr(1, PathExp, ".exe") Or InStr(1, PathExp, ".pif") Or InStr(1, PathExp, ".htm")
            t = InStr(1, PathExp, ".", vbBinaryCompare)
            tmpStr = Left(PathExp, t + 3)
            tmpStr = ChuoiGiaTri(tmpStr)
                    Set lsv1 = frmMain.LVV.ListItems.Add()
                    lsv1.Text = kName
                    lsv1.SubItems(1) = tmpStr
                    If hKey = BB Then
                        lsv1.SubItems(2) = "HKEY_LOCAL_MACHINE\" & kPath ' & kName
                    ElseIf hKey = aa Then
                        lsv1.SubItems(2) = "HKEY_CURRENT_USER\" & kPath ' & kName
                    End If
                    'lsv1.Checked = True
            If Len(PathExp) >= t + 4 Then
                PathExp = Right(PathExp, Len(PathExp) - t - 4)
            Else
                PathExp = ""
            End If
        Loop
End Sub
Public Function NormalkanAtribut(sPath As String)
On Error Resume Next
If GetFileAttributes(StrPtr(sPath)) = 4 Then ' system
   SetFileAttributes StrPtr(sPath), 0
ElseIf GetFileAttributes(StrPtr(sPath)) = 6 Then ' hidden + system
   SetFileAttributes StrPtr(sPath), 0
ElseIf GetFileAttributes(StrPtr(sPath)) = 2 Then '
   SetFileAttributes sPath, 0
ElseIf GetFileAttributes(StrPtr(sPath)) = 38 Then '
   SetFileAttributes StrPtr(sPath), 0
ElseIf GetFileAttributes(StrPtr(sPath)) = 39 Then '
   SetFileAttributes StrPtr(sPath), 0
End If
End Function
Public Function HapusFile(sPath As String) As Boolean
On Error GoTo Falsex
NormalkanAtribut sPath
If DeleteFile(StrPtr(sPath)) = 1 Then
   HapusFile = True
Else
   If DeleteFile(StrPtr("\\.\" & sPath)) = 1 Then
      HapusFile = True
   End If
End If
If ValidFile(sPath) = True Then GoTo Falsex
Exit Function
Falsex:
HapusFile = False
End Function
Public Function ValidFile(ByRef sFile As String) As Boolean ' Memvalidasi file
If PathIsDirectory(StrPtr(sFile)) = 0 And PathFileExists(StrPtr(sFile)) = 1 Then
    ValidFile = True
Else
    ValidFile = False
End If
End Function
'########## Memilih folder untuk di scan ##########'
'Dim Pathnya As String
'Pathnya = ""
'Pathnya = BrowseFolder("Pilih folder yang akan di Scan:", Me)
'If Pathnya <> "" Then
'txtPath.Text = Pathnya
'End If
'End Sub
