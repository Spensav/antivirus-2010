Attribute VB_Name = "ModRegistryScan"
Private Const READ_CONTROL = &H20000
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_CREATE_LINK = &H20
Private Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL

' Ini Yang Paling Penting....
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_USERS = &H80000003
Private Const HKEY_CURRENT_CONFIG = &H80000005
Private Const HKEY_PERFORMANCE_DATA = &H80000004
Private Const HKEY_DYN_DATA = &H80000006

Private Const ERROR_SUCCESS = 0
Private Const REG_SZ = 1                         ' Unicode nul terminated string
Private Const REG_DWORD = 4                      ' 32-bit number

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long



Private Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim i As Long                                           ' Loop Counter
    Dim Rc As Long                                          ' Return Code
    Dim hKey As Long                                        ' Handle To An Open Registry Key
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' Data Type Of A Registry Key
    Dim tmpVal As String                                    ' Tempory Storage For A Registry Key Value
    Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
    '------------------------------------------------------------
    ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    Rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key
    
    If (Rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Error...
    
    tmpVal = String$(1024, 0)                             ' Allocate Variable Space
    KeyValSize = 1024                                       ' Mark Variable Size
    
    '------------------------------------------------------------
    ' Retrieve Registry Key Value...
    '------------------------------------------------------------
    Rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value
                        
    If (Rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 Adds Null Terminated String...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Null Found, Extract From String
    Else                                                    ' WinNT Does NOT Null Terminate String...
        tmpVal = Left(tmpVal, KeyValSize)                   ' Null Not Found, Extract String Only
    End If
    '------------------------------------------------------------
    ' Determine Key Value Type For Conversion...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' Search Data Types...
    Case REG_SZ                                             ' String Registry Key Data Type
        KeyVal = tmpVal                                     ' Copy String Value
    Case REG_DWORD                                          ' Double Word Registry Key Data Type
        For i = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Build Value Char. By Char.
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' Convert Double Word To String
    End Select
    
    GetKeyValue = True                                      ' Return Success
    Rc = RegCloseKey(hKey)                                  ' Close Registry Key
    Exit Function                                           ' Exit
    
GetKeyError:      ' Cleanup After An Error Has Occured...
    KeyVal = ""                                             ' Set Return Val To Empty String
    GetKeyValue = False                                     ' Return Failure
    Rc = RegCloseKey(hKey)                                  ' Close Registry Key
End Function

Function Cek_Value()
   On Error Resume Next
   Dim G As ListItem
   Dim Path As String
   Dim SubKey As String
   Dim value As String
   frmMain.List2.AddItem " "
   'DisableTaskMgr
   If GetKeyValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableTaskMgr", value) Then
   If value = 0 Then
   Else
   frmMain.List2.AddItem "DisableTaskMgr"
   Set G = frmMain.lvReg.ListItems.Add(, , "DisableTaskMgr", , frmMain.ImgListView.ListImages(1).Index)
   G.SubItems(1) = "Seharusnya Nilai Value 0"
   G.SubItems(2) = "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"
   End If
   End If
   'DisableRegistryTools
   If GetKeyValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableRegistryTools", value) Then
   If value = 0 Then
   Else
   frmMain.List2.AddItem "DisableRegistryTools"
   Set G = frmMain.lvReg.ListItems.Add(, , "DisableRegistryTools", , frmMain.ImgListView.ListImages(1).Index)
   G.SubItems(1) = "Seharusnya Nilai Value 0"
   G.SubItems(2) = "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"
   End If
   End If
   'DisableConfig
   If GetKeyValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\SystemRestore", "DisableConfig", value) Then
   If value = 0 Then
   Else
   frmMain.List2.AddItem "DisableConfig"
   Set G = frmMain.lvReg.ListItems.Add(, , "DisableConfig", , frmMain.ImgListView.ListImages(1).Index)
   G.SubItems(1) = "Seharusnya Nilai Value 0"
   G.SubItems(2) = "HKLM\SOFTWARE\Microsoft\Windows NT\SystemRestore"
   End If
   End If
   'Userinit
   If GetKeyValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "Userinit", value) Then
   If value = "userinit.exe" Then
   Else
   frmMain.List2.AddItem "Userinit"
   Set G = frmMain.lvReg.ListItems.Add(, , "Userinit", , frmMain.ImgListView.ListImages(1).Index)
   G.SubItems(1) = "Seharusnya Nilai Value userinit.exe"
   G.SubItems(2) = "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"
   End If
   End If
   'HKU-DisableTaskMgr
   If GetKeyValue(HKEY_USERS, ".DEFAULT\Software\Microsoft\Windows\CurrentVersion\Policies\system", "DisableTaskMgr", value) Then
   If value = "0" Then
   Else
   frmMain.List2.AddItem "HKU-DisableTaskMgr"
   Set G = frmMain.lvReg.ListItems.Add(, , "DisableTaskMgr", , frmMain.ImgListView.ListImages(1).Index)
   G.SubItems(1) = "Seharusnya Nilai Value 0"
   G.SubItems(2) = "HKU\.DEFAULT\Software\Microsoft\Windows\CurrentVersion\Policies\system"
   End If
   End If
   'HKU-DisableRegistryTools
   If GetKeyValue(HKEY_USERS, ".DEFAULT\Software\Microsoft\Windows\CurrentVersion\Policies\system", "DisableRegistryTools", value) Then
   If value = "0" Then
   Else
   frmMain.List2.AddItem "HKU-DisableRegistryTools"
   Set G = frmMain.lvReg.ListItems.Add(, , "DisableRegistryTools", , frmMain.ImgListView.ListImages(1).Index)
   G.SubItems(1) = "Seharusnya Nilai Value 0"
   G.SubItems(2) = "HKU\.DEFAULT\Software\Microsoft\Windows\CurrentVersion\Policies\system"
   End If
   End If
End Function
Public Function ClearAutorun()
    Dim i As Long
    Dim tmp As Long
    Select Case frmMain.List2.Text
                Case "  "
                'sfsf
                Case "DisableTaskMgr"
                    CreateDwordValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableTaskMgr", 0
                Case "DisableRegistryTools"
                    CreateDwordValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableRegistryTools", 0
                Case "DisableConfig"
                    CreateDwordValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\SystemRestore", "DisableConfig", 0
                Case "Userinit"
                    CreateStringValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", REG_SZ, "Userinit", "userinit.exe"
                Case "HKU-DisableTaskMgr"
                    CreateDwordValue HKEY_USERS, ".DEFAULT\Software\Microsoft\Windows\CurrentVersion\Policies\system", "DisableTaskMgr", 0
                Case "HKU-DisableRegistryTools"
                    CreateDwordValue HKEY_USERS, ".DEFAULT\Software\Microsoft\Windows\CurrentVersion\Policies\system", "DisableRegistryTools", 0
    End Select
    Set fso = Nothing
End Function
Public Function Perbaiki()
CreateDwordValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableTaskMgr", 0
CreateDwordValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableRegistryTools", 0
CreateDwordValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\SystemRestore", "DisableConfig", 0
CreateStringValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", REG_SZ, "Userinit", "userinit.exe"
CreateDwordValue HKEY_USERS, ".DEFAULT\Software\Microsoft\Windows\CurrentVersion\Policies\system", "DisableTaskMgr", 0
CreateDwordValue HKEY_USERS, ".DEFAULT\Software\Microsoft\Windows\CurrentVersion\Policies\system", "DisableRegistryTools", 0
End Function
