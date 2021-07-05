Attribute VB_Name = "ModStartupMgr4"
'vnAntiVirus 1.0

'Author : Dung Le Nguyen
'Email : dungcoivb@gmail.com
'This is a software open source

'Code this module  from PSC
Option Explicit

Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Public Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As Any, phkResult As Long, lpdwDisposition As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As Any) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long

Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_CURRENT_USER = &H80000001

Public Const HKEY_DYN_DATA = &H80000006
Public Const HKEY_USERS = &H80000003
Public Const REG_OPTION_NON_VOLATILE = 0
Public Const STANDARD_RIGHTS_ALL = &H1F0000
Public Const SYNCHRONIZE = &H100000
Public Const READ_CONTROL = &H20000
'Module nay co nguon goc tu PSC
Public Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Public Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
Public Const KEY_CREATE_LINK = &H20
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_SET_VALUE = &H2
Public Const REG_NONE As Long = 0
Public Const REG_SZ As Long = 1
Public Const REG_EXPAND_SZ As Long = 2
Public Const REG_BINARY As Long = 3
Public Const REG_DWORD As Long = 4
Public Const REG_LINK As Long = 6
Public Const REG_MULTI_SZ As Long = 7
Public Const REG_RESOURCE_LIST As Long = 8
Public Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
Public Const ERROR_NO_MORE_ITEMS = 259&
Public Const BUFFER_SIZE As Long = 255
Public Result As Long
Public Ret As Long
'Public val As String
Public START As Long
Public curKey As String
Public hCurKey As Long

Public Enum Key
    aa = HKEY_CURRENT_USER
    BB = HKEY_LOCAL_MACHINE
End Enum
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Public Const Pathkey  As String = "Software\Microsoft\Windows\CurrentVersion\Run"

Public Function ChuoiGiaTri(Chuoi As String) As String
'Debug.Print Chuoi
    Dim i, j, t
    Dim strTMP As String
    Chuoi = Replace(Chuoi, ",", "", 1, , vbBinaryCompare)
    If FileExists(CheckPath(Chuoi)) = True Then Chuoi = CheckPath(Chuoi)
'Thao tac xu ly cac chuoi gia tri tu regedit de dua ra ket qua chuan
    If FileExists(Chuoi) = True Then GoTo Okie
    If FileExists(CheckPath(Chuoi)) = True Then Chuoi = CheckPath(Chuoi): GoTo Okie
    
    If InStr(1, Chuoi, "%systemroot%", vbTextCompare) <> 0 Then
        Chuoi = Replace(Chuoi, "%systemroot%", WindowsDir, 1, , vbTextCompare)
    Else
        Chuoi = Chuoi
    End If
        If FileExists(Chuoi) = True Then GoTo Okie
        
    strTMP = Right(Chuoi, InStr(1, StrReverse(Chuoi), "\", vbTextCompare))
    t = InStr(1, strTMP, " ", vbTextCompare)
    
    If t <> 0 Then strTMP = Left(Chuoi, Len(Chuoi) - (Len(strTMP) - t + 1))
    If FileExists(strTMP) = True Then Chuoi = strTMP: GoTo Okie
    If FileExists(CheckPath(strTMP)) = True Then Chuoi = CheckPath(strTMP): GoTo Okie
    
    If InStr(1, Chuoi, Chr(34), vbTextCompare) <> 0 Then
        strTMP = StrReverse(Chuoi)

    i = InStr(1, strTMP, ".", vbTextCompare)
    j = InStr(1, Left(strTMP, i), " ", vbTextCompare)
    strTMP = StrReverse(Right(strTMP, Len(strTMP) - j))
            
    End If
    
    If FileExists(strTMP) = True Then Chuoi = strTMP: GoTo Okie
        
        Dim Buf As String
        Dim Path As String
        Dim vt1 As Byte
        Dim vt2 As Byte
        If InStr(1, Chuoi, Chr(34), vbTextCompare) <> 0 Then
            vt1 = InStr(1, Chuoi, Chr(34), vbTextCompare)
            Buf = Right(Chuoi, Len(Chuoi) - vt1)
            vt2 = InStr(1, Buf, Chr(34), vbTextCompare)
            strTMP = Left(Buf, vt2 - 1)
            If FileExists(strTMP) = True Then Chuoi = strTMP: GoTo Okie
        End If
    'Chuoi = StrReverse(Chuoi)
    
    Chuoi = CheckPath(Chuoi)
Okie:
    ChuoiGiaTri = Chuoi
End Function
Public Function GetString1(hKey As Long, strPath As String, strValue As String) As String
    Dim lngValueType As Long
    Dim strBuffer As String
    Dim lngDataBufferSize As Long
    Dim intZeroPos As Integer
    RegOpenKey hKey, strPath, hCurKey
    RegQueryValueEx hCurKey, strValue, 0&, lngValueType, ByVal 0&, lngDataBufferSize
    If lngValueType = REG_SZ Then
        strBuffer = String(lngDataBufferSize, " ")
        RegQueryValueEx hCurKey, strValue, 0&, 0&, ByVal strBuffer, lngDataBufferSize
        intZeroPos = InStr(strBuffer, Chr$(0))
        If intZeroPos > 0 Then
            GetString1 = Left$(strBuffer, intZeroPos - 1)
        Else
            GetString1 = strBuffer
        End If
    End If
    RegCloseKey hCurKey
End Function
Public Sub SaveString(hKey As Long, strPath As String, strValue As String, strData As String)
    Dim Ret
    RegCreateKey hKey, strPath, Ret
    RegSetValueEx Ret, strValue, 0, REG_SZ, ByVal strData, Len(strData)
    RegCloseKey Ret
End Sub
Public Sub DelSetting(hKey As Long, strPath As String, strValue As String)
    Dim Ret
    RegCreateKey hKey, strPath, Ret
    RegDeleteValue Ret, strValue
    RegCloseKey Ret
End Sub
Public Function WindowsDir() As String
    Dim WindirS As String * 255
    Dim Temp
    Dim Result
    Temp = GetWindowsDirectory(WindirS, 255)
    Result = Left(WindirS, Temp)
    WindowsDir = Result
End Function
Public Function CheckPath(Chuoi As String) As String

    If FileExists(Chuoi) = True Then
    ElseIf FileExists(WindowsDir & "\" & Chuoi) = True Then
        CheckPath = WindowsDir & "\" & Chuoi
    ElseIf FileExists(WindowsDir & "\system32\" & Chuoi) = True Then
        CheckPath = WindowsDir & "\system32\" & Chuoi
    ElseIf FileExists(Chuoi & ".exe") = True Then
        CheckPath = Chuoi & ".exe"
    Else
        CheckPath = "Not found " & Chuoi
    End If
End Function



