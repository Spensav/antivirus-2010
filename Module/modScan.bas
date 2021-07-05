Attribute VB_Name = "modScan"
'Pakai Engine Ansav dan Spensav
Public Const AN_CLEAN = 0
Public Const AN_VIRUS_DETECTED = 1
Public Const AN_SCAN_ERROR = 3
Public Const ERROR_FILE_NOT_FOUND = 2
Public Const NO_MORE_VIRUSES = 0
Public Const MAX_VIRUSNAME_LENGTH = 30

Public Type ANSAV_VERSION_INFO
dwMajor As Long
dwMinor As Long
dwRevision As Long
End Type

Public Declare Function AnsavVirusCount Lib "ansavcore.dll" () As Long
Public Declare Function AnsavGetVersion Lib "ansavcore.dll" (lpAVI As ANSAV_VERSION_INFO) As Long
Public Declare Function AnsavVirusFirst Lib "ansavcore.dll" (ByVal lpszVirusName As String, nSize As Long) As Long
Public Declare Function AnsavVirusNext Lib "ansavcore.dll" (ByVal hFind As Long, ByVal lpszVirusName As String, nSize As Long) As Long
Public Declare Function AnsavVirusClose Lib "ansavcore.dll" (ByVal hFind As Long) As Long
Public Declare Function CheckWithAnsav Lib "ansavcore.dll" (ByVal Filename As String, ByVal VirusName As String, nSize As Long) As Long

Public jumlahDir As Long, jumlahFile As Long, jumlahVirus As Long
Public StopScan As Boolean

Public Function CariVirusDenganHeur(FilePath As String) As String
CariVirusDenganHeur = ""
For i = 1 To UBound(VirusDB)
If GetChecksum(FilePath) = Split(VirusDB(i), "|")(1) Then
CariVirusDenganHeur = VirusDB(i)
Exit Function
End If
Next
'If FileLen(FilePath) / 1024 <= 512 Then
'CekVirus = CekHeuristic(FilePath)
'End If
End Function
Public Sub ScanWithSpensav(ByVal lpFolderName As String, ByVal SubDirs As Boolean)
    Dim i As Long
    Dim hSearch As Long, WFD As WIN32_FIND_DATA
    Dim Result As Long, CurItem As String
    Dim tempDir() As String, dirCount As Long
    Dim RealPath As String
    Dim CekDulu As String
    Dim RetVal As Long
    Dim nSize As Long
    Dim Buff As String * MAX_VIRUSNAME_LENGTH
    Dim NamaVirus As String
    Dim a
    CekDulu = ""
    dirCount = -1
    
    ScanInfo = "Scan File"
    
    If Right$(lpFolderName, 1) = "\" Then
        RealPath = lpFolderName
    Else
        RealPath = lpFolderName & "\"
    End If
    
    hSearch = FindFirstFile(RealPath & "*", WFD)
    If Not hSearch = INVALID_HANDLE_VALUE Then
        Result = True
        Do While Result
            DoEvents
            If StopScan = True Then Exit Do
            CurItem = StripNulls(WFD.cFileName)
            If Not CurItem = "." And Not CurItem = ".." Then
                If PathIsDirectory(RealPath & CurItem) <> 0 Then
                    jumlahDir = jumlahDir + 1
                    frmMain.Text6.Text = jumlahFile & " Jumlah File  " & jumlahDir & " Direktori"
                    If SubDirs = True Then
                        dirCount = dirCount + 1
                        ReDim Preserve tempDir(dirCount) As String
                        tempDir(dirCount) = RealPath & CurItem
                    End If
                Else
                    jumlahFile = jumlahFile + 1
                    a = (100 / totalFiles1) * jumlahFile
                    If a <= 100 Then
                        frmUSB.ProgressBar1RTP.value = a
                        frmMain.ProgressBar1.value = a
                        frmUSB.Text11RTP.Text = frmMain.ProgressBar1.value & " %"
                        frmMain.Text11.Text = frmMain.ProgressBar1.value & " %"
                    End If
                    

                    frmMain.txtPindaiSel.SelStart = Len(frmMain.txtPindaiSel.Text)
                    frmMain.Label21.Caption = ": " & frmMain.txtPindaiSel.SelStart & " /s"
                    frmMain.Text6.Text = jumlahFile & " Jumlah File  " & jumlahDir & " Direktori"
                    
                    frmMain.txtPindaiSel.Text = RealPath & CurItem
                    If Len(RealPath & CurItem) > 50 Then 'jika panjang nama file > 50
                    If Len(CurItem) < 15 Then
                    frmUSB.txtPindaiUSB.Text = Mid(RealPath, 1, InStr(4, RealPath, "\")) & "..." & "\" & CurItem
                    frmMain.txtPindai.Text = Mid(RealPath, 1, InStr(4, RealPath, "\")) & "..." & "\" & CurItem
                    Else
                    frmUSB.txtPindaiUSB.Text = Mid(RealPath, 1, InStr(4, RealPath, "\")) & "..." & "\" & "..." & Right(CurItem, 15)
                    frmMain.txtPindai.Text = Mid(RealPath, 1, InStr(4, RealPath, "\")) & "..." & "\" & "..." & Right(CurItem, 15)
                    End If
                    End If
                    
                    Dim xa As Integer
                    For xa = 1 To frmMain.ListView3.ListItems.Count
                    'Exeption List
                    If CurItem = frmMain.ListView3.ListItems(xa).Text Then
                    GoTo NggakUsah
                    End If
                    Next

                    If (WFD.nFileSizeLow) / 1024 >= 750 Or (WFD.nFileSizeHigh) / 1024 >= 750 Then
                    GoTo NggakUsah ' Jika ukuran besar, tidak usah dicek
                    End If
                    
                    If WFD.nFileSizeLow > 5120 Or WFD.nFileSizeHigh > 5120 Then
                        CekDulu = CariVirusDenganHeur(RealPath & CurItem)
                        If CekDulu <> "" Then
                        LvwSubStyle frmMain.lvVirus, "W32:Elektrik.A", RealPath & CurItem, "Internal Virus"
                        jumlahVirus = jumlahVirus + 1
                        End If
                    End If
            nSize = MAX_VIRUSNAME_LENGTH
            'RetVal = CheckWithAnsav(RealPath & CurItem, Buff, nSize)
            'If RetVal = 1 Then
            'NamaVirus = Left$(Buff, nSize)
            'LvwSubStyle frmMain.lvVirus, NamaVirus, RealPath & CurItem, "Engine Detektor"
            'End If
            
            isFileVirus CStr(RealPath & CurItem)
            MeiPattern CStr(RealPath & CurItem), 202
            
            'Ini Untuk By.USER Virus
            Dim s As Integer
            If frmMain.ListView2.ListItems.Count <> 0 Then
            For s = 1 To frmMain.ListView2.ListItems.Count
            If CurItem = GetFileName(frmMain.ListView2.ListItems(s).Text) Then
            LvwSubStyle frmMain.lvVirus, frmMain.Text2.Text, RealPath & CurItem, "Dari User"
            jumlahVirus = jumlahVirus + 1
            End If
            Next
            End If
            
            
                End If
            End If
NggakUsah:
            Result = FindNextFile(hSearch, WFD)
        Loop
        FindClose hSearch
        
        If SubDirs = True Then
            If dirCount <> -1 Then
                For i = 0 To dirCount
                    ScanWithSpensav tempDir(i), True
                Next i
            End If
        End If
    End If
End Sub
Public Function LvwSubStyle(lvStyle As ListView, Data As String, RumahVirus As String, TipeVirus As String)
Dim G As ListItem
Set G = lvStyle.ListItems.Add(, , Data, , frmMain.ImgListView.ListImages(2).Index)
G.Icon = 1
G.SubItems(1) = TipeVirus
G.SubItems(2) = RumahVirus
End Function
