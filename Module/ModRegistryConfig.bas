Attribute VB_Name = "ModRegistryConfig"
Enum TypeBase
    TypeHexadecimal
    TypeDecimal
End Enum
Enum RegistryKeys
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_CURRENT_USER = &H80000001
    HKEY_DYN_DATA = &H80000006
    HKEY_PERFORMANCE_DATA = &H80000004
    HKEY_USERS = &H80000003
End Enum
Enum RegDataTypes
    REG_SZ = 1                         ' Unicode nul terminated string
    REG_EXPAND_SZ = 2                  ' Unicode nul terminated string
    REG_DWORD = 4                      ' 32-bit number
    REG_MULTI_SZ = 7
End Enum
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Public Declare Function RegQueryValueExA Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByRef lpData As Long, lpcbData As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.

Public Sub CleanReg()

    On Error Resume Next
    CreateDwordValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableRegistryTools", 0
    CreateDwordValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System", "DisableRegistryTools", 0
    CreateDwordValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System\", "DisableTaskMgr", 0
    CreateDwordValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System\", "DisableTaskMgr", 0
    CreateDwordValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableCMD", 0
    CreateDwordValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFolderOptions", 0
    CreateDwordValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFolderOptions", 0
    DeleteValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "Bron-Spizaetus"
    DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", "Tok-Cirrhatus"
    DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "Shell"
    CreateStringValue HKEY_CLASSES_ROOT, "exefile\shell\open\command", REG_SZ, "", Chr(&H22) & "%1" & Chr(&H22) & " %*"
    CreateStringValue HKEY_CLASSES_ROOT, "lnkfile\shell\open\command", REG_SZ, "", Chr(&H22) & "%1" & Chr(&H22) & " %*"
    CreateStringValue HKEY_CLASSES_ROOT, "piffile\shell\open\command", REG_SZ, "", Chr(&H22) & "%1" & Chr(&H22) & " %*"
    CreateStringValue HKEY_CLASSES_ROOT, "batfile\shell\open\command", REG_SZ, "", Chr(&H22) & "%1" & Chr(&H22) & " %*"
    CreateStringValue HKEY_CLASSES_ROOT, "comfile\shell\open\command", REG_SZ, "", Chr(&H22) & "%1" & Chr(&H22) & " %*"
    CreateStringValue HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\SafeBoot\", REG_SZ, "AlternateShell", "cmd.exe"
    CreateStringValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", REG_SZ, "Userinit", GetSystemPath & "userinit.exe"
    CreateStringValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\AeDebug", REG_SZ, "Debugger", Chr(&H22) & Left(GetWindowsPath, 3) & "Program Files\Microsoft Visual Studio\Common\MSDev98\Bin\msdev.exe" & Chr(&H22) & " -p %ld -e %ld"
    CreateStringValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\AeDebug", REG_SZ, "Auto", "0"
    CreateDwordValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\WinOldApp\", "Disabled", 0
    CreateDwordValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\WinOldApp\", "Disabled", 0
    CreateStringValue HKEY_CLASSES_ROOT, "exefile", REG_SZ, "", "Application"
    CreateDwordValue HKEY_LOCAL_MACHINE, "SOFTWARE\Policies\Microsoft\Windows NT\SystemRestore", "DisableConfig", 0
    CreateDwordValue HKEY_LOCAL_MACHINE, "SOFTWARE\Policies\Microsoft\Windows NT\SystemRestore", "DisableSR", 0
    CreateDwordValue HKEY_LOCAL_MACHINE, "SOFTWARE\Policies\Microsoft\Windows\Installer", "LimitSystemRestoreCheckpointing", 0
    CreateDwordValue HKEY_LOCAL_MACHINE, "SOFTWARE\Policies\Microsoft\Windows\Installer", "DisableMSI", 0
    
End Sub

Public Function DeleteKey(hKey As RegistryKeys, SubKey As String) As Long

    On Error Resume Next
    DeleteKey = RegDeleteKey(hKey, SubKey)
    RegCloseKey ret
    
End Function

Public Function DeleteValue(hKey As RegistryKeys, SubKey As String, lpValName As String) As Long
    
    Dim ret As Long
    On Error Resume Next
    RegOpenKey hKey, SubKey, ret
    DeleteValue = RegDeleteValue(ret, lpValName)
    RegCloseKey ret
    
End Function

Public Function CreateDwordValue(hKey As RegistryKeys, SubKey As String, strValueName As String, dwordData As Long) As Long
    Dim ret As Long
    On Error Resume Next
    RegCreateKey hKey, SubKey, ret
    CreateDwordValue = RegSetValueEx(ret, strValueName, 0, REG_DWORD, dwordData, 4)
    RegCloseKey ret
    
End Function

Public Function CreateStringValue(hKey As RegistryKeys, SubKey As String, RTypeStringValue As RegDataTypes, strValueName As String, strData As String) As Long
    Dim ret As Long
    On Error Resume Next
    RegCreateKey hKey, SubKey, ret
    CreateStringValue = RegSetValueEx(ret, strValueName, 0, RTypeStringValue, ByVal strData, Len(strData))
    RegCloseKey ret
    
End Function
Public Function GetDWORDValue(hKey As RegistryKeys, SubKey As String, Entry As String)
    Dim ret As Long
    rtn = RegOpenKeyEx(hKey, SubKey, 0, KEY_READ, ret)
    If rtn = ERROR_SUCCESS Then
        rtn = RegQueryValueExA(ret, Entry, 0, REG_DWORD, lBuffer, 4)
        If rtn = ERROR_SUCCESS Then
            rtn = RegCloseKey(ret)
            GetDWORDValue = lBuffer
        Else
            GetDWORDValue = "Error"
        End If
    Else
        GetDWORDValue = "Error"
    End If
End Function
Public Function ReadValue(hKey As RegistryKeys, SubKey As String, strValueName As String) As String

    Dim RootKey As Long
    Dim isi As String
    Dim lDataBufSize As Long
    Dim lValueType As Long
    On Error Resume Next
    X = RegOpenKey(hKey, SubKey, RootKey)
    ret = RegQueryValueEx(RootKey, strValueName, 0, lValueType, 0, lDataBufSize)
    isi = String(lDataBufSize, Chr$(0))
    ret = RegQueryValueEx(RootKey, strValueName, 0, 0, ByVal isi, lDataBufSize)
    
    ReadValue = Left$(isi, InStr(1, isi, Chr$(0)) - 1)
    RegCloseKey RootKey

End Function


'used for LoadRegistry

Function GetString(hKey As Long, strPath As String, strValue As String)
'----------------------------------------------------------------------------
'Argument       :   Handlekey, path from the root , Name of the Value in side the key
'Return Value   :   String
'Function       :   To fetch the value from a key in the Registry
'Comments       :   on Success , returns the Value else empty String
'----------------------------------------------------------------------------

    Dim ret
    'Open  key
    RegOpenKey hKey, strPath, ret
    'Get content
    GetString = RegQueryStringValue(ret, strValue)
    'Close the key
    RegCloseKey ret
End Function

Function RegQueryStringValue(ByVal hKey As Long, ByVal strValueName As String) As String
'----------------------------------------------------------------------------
'Argument       :   Handlekey, Name of the Value in side the key
'Return Value   :   String
'Function       :   To fetch the value from a key in the Registry
'Comments       :   on Success , returns the Value else empty String
    '----------------------------------------------------------------------------
    Dim lResult As Long, lValueType As Long, strBuf As String, lDataBufSize As Long
    
    lResult = RegQueryValueEx(hKey, strValueName, 0, lValueType, ByVal 0, lDataBufSize)
    If lResult = 0 Then
        If lValueType = REG_SZ Then
    
            strBuf = String(lDataBufSize, Chr$(0))
            'retrieve the key's value
            lResult = RegQueryValueEx(hKey, strValueName, 0, 0, ByVal strBuf, lDataBufSize)
            If lResult = 0 Then
    
                RegQueryStringValue = Left$(strBuf, InStr(1, strBuf, Chr$(0)) - 1)
            End If
        ElseIf lValueType = REG_BINARY Then
            Dim strData As Integer
            'retrieve the key's value
            lResult = RegQueryValueEx(hKey, strValueName, 0, 0, strData, lDataBufSize)
            If lResult = 0 Then
                RegQueryStringValue = strData
            End If
         ElseIf lValueType = REG_DWORD Then
           
            'retrieve the key's value
            lResult = RegQueryValueEx(hKey, strValueName, 0, 0, strData, lDataBufSize)
            If lResult = 0 Then
                RegQueryStringValue = strData
            End If
            
        End If
    End If
End Function
