Attribute VB_Name = "ModSaveSetting"
Private Declare Function GetPrivateProfileString Lib "Kernel32.dll" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "Kernel32.dll" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Function INIFileName() As String
    INIFileName = App.path & "\Pengaturan$.ini" 'atur lokasi file Konfigurasi.ini disini
End Function

Public Function Setting(ByVal Section As String, ByVal Key As String, ByVal DefaultValue)
    Dim s As String, L As Long
    s = String(255, 0)
    L = GetPrivateProfileString(Section, Key, DefaultValue, s, 255, INIFileName)
    SetinganChoy = Left(s, L)
End Function

Public Function Simpan(ByVal Section As String, ByVal Key As String, ByVal Setting As String)
    If Setting = "" Then Setting = vbNullChar
    Simpan = WritePrivateProfileString(Section, Key, Setting, INIFileName)
End Function
Public Function ceksetting()
frmMain.SettingPelindung(0).Checked = Setting("SPENSAV", "Proteksi", True)
frmMain.SettingPelindung(1).Checked = Setting("SPENSAV", "Firewall", False)
frmMain.SettingPelindung(2).Checked = Setting("SPENSAV", "Behavor", True)
frmMain.SettingPelindung(3).Checked = Setting("SPENSAV", "ScanArchive", True)
frmMain.SettingUser(0).Checked = Setting("SPENSAV", "tembus", False)
frmMain.SettingUser(1).Checked = Setting("SPENSAV", "startup", False)
frmMain.SettingUser(2).Checked = Setting("SPENSAV", "Behavor1", True)
frmMain.SettingUser(3).Checked = Setting("SPENSAV", "runtop", False)
End Function
Public Function simpansetting()
Simpan "SPENSAV", "Proteksi", frmMain.SettingPelindung(0).Checked
Simpan "SPENSAV", "Firewall", frmMain.SettingPelindung(1).Checked
Simpan "SPENSAV", "Behavor", frmMain.SettingPelindung(2).Checked
Simpan "SPENSAV", "ScanArchive", frmMain.SettingPelindung(3).Checked
Simpan "SPENSAV", "tembus", frmMain.SettingUser(0).Checked
Simpan "SPENSAV", "startup", frmMain.SettingUser(1).Checked
Simpan "SPENSAV", "Behavor1", frmMain.SettingUser(2).Checked
Simpan "SPENSAV", "runtop", frmMain.SettingUser(3).Checked
End Function
