Attribute VB_Name = "ModBuffer"
Private WFD As WIN32_FIND_DATA
Private Declare Function FindFirstFileW Lib "KERNEL32" (ByVal lpFilename As Long, ByVal lpFindFileData As Long) As Long
Private Declare Function FindNextFileW Lib "KERNEL32" (ByVal hFindFile As Long, ByVal lpFindFileData As Long) As Long
Private Pindai As Boolean, Pause As Boolean, SmartScan As Boolean
Public jmlFiles As Long, jmlDirs As Long, totalFiles1 As Long
Private VirName As String, TypeVir As String
Private LokasiScan, InfoPath As String, ScanInfo As String
Private Const vbStart = "*"
Private Const vbAllFiles = "*.*"
Public Const vbBackSlash As String = "\"
Private Const vbKeyDot = 46
Private Extension As String
Private Const ExtensionDefault As String = "*.exe;*.dll;*.ocx;*.bin;*.sys;*.com;*.cmd;*.bat;*.vbs;*.pif;*.dat;*.ini;*.txt;*.inf;*.htt;*.cpl;*.scr;*.chm;*.lnk;*.pcx;*.ico;*.htm;*.xml;*.css;*.php;*.jsp;*.asp;*.pdf;*.doc;*.tif;*.asm"
Private Function StripNulls(ByVal OriginalStr As String) As String
If (InStr(OriginalStr, Chr$(0)) > 0) Then
    OriginalStr = Left$(OriginalStr, InStr(OriginalStr, Chr$(0)) - 1)
End If
StripNulls = OriginalStr
End Function
Public Function AnalisaFiles(ByVal Path As String) 'u/ menghitung file
Dim i As Long, Cari As Long, dirCount As Long
Dim Direktori() As String, NamaFile As String
DoEvents
On Error Resume Next
Cari = FindFirstFileW(StrPtr(Path & ChrW$(42)), VarPtr(WFD))
If Not Cari = INVALID_HANDLE_VALUE Then
    Do
    If Pindai = True Then Exit Do
    If (WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) Then
        If Asc(WFD.cFileName) <> vbKeyDot Then
            If (dirCount Mod 10) = 0 Then ReDim Preserve Direktori(dirCount + 10)
            dirCount = dirCount + 1
            Direktori(dirCount) = StripNulls(WFD.cFileName)
        End If
    Else
        totalFiles1 = totalFiles1 + 1
    End If

    Loop While FindNextFileW(Cari, VarPtr(WFD))
    FindClose (Cari)
    For i = 1 To dirCount
        AnalisaFiles Path & Direktori(i) & vbBackSlash
    Next i
    jmlDirs = jmlDirs + 1
    frmMain.Label17.Caption = ": " & jmlDirs
    frmMain.Text6.Text = "Menghitung " & totalFiles1 & " File Dan " & jmlDirs & " Folder..."
    ScanInfo = "Analyzing " & totalFiles1 & " Files and " & jmlDirs & " Directories..."
End If
End Function

