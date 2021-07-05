Attribute VB_Name = "ModScanRTP"
'Engine Spensav
Public JumlahFolder As Long, JumlahDatanya As Long, jumlahVirus1 As Long
Public BerhentiRTP As Boolean
Public Sub ScanRTP(ByVal lpFolderName As String, ByVal SubDirs As Boolean)
    Dim i As Long
    Dim hSearch As Long, WFD As WIN32_FIND_DATA
    Dim Result As Long, CurItem As String
    Dim tempDir() As String, dirCount As Long
    Dim RealPath As String, GetViri As String
    Dim RetVal As Long
    Dim nSize As Long
    Dim Buff As String * MAX_VIRUSNAME_LENGTH
    Dim NamaVirus As String
    Dim a
    GetViri = ""
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
            If BerhentiRTP = True Then Exit Do
            CurItem = StripNulls(WFD.cFileName)
            If Not CurItem = "." And Not CurItem = ".." Then
                If PathIsDirectory(RealPath & CurItem) <> 0 Then
                    JumlahFolder = JumlahFolder + 1
                    If SubDirs = True Then
                        dirCount = dirCount + 1
                        ReDim Preserve tempDir(dirCount) As String
                        tempDir(dirCount) = RealPath & CurItem
                    End If
                Else
                    JumlahDatanya = JumlahDatanya + 1
                    Dim xa As Integer
                    For xa = 1 To frmMain.ListView3.ListItems.Count
                    'Exeption List
                    If CurItem = frmMain.ListView3.ListItems(xa).Text Then
                    GoTo NggakUsah
                    End If
                    Next
                    If UntukQuickScan(RealPath & CurItem) = "Tak diketahui." Then
                    GoTo NggakUsah
                    End If
                    If (WFD.nFileSizeLow) / 1024 >= 750 Or (WFD.nFileSizeHigh) / 1024 >= 750 Then
                    GoTo NggakUsah ' Jika ukuran besar, tidak usah dicek
                    End If
                    frmRTP.Label1.Caption = RealPath & CurItem
                    If WFD.nFileSizeLow > 5120 Or WFD.nFileSizeHigh > 5120 Then
                        GetViri = CariVirusDenganHeur(RealPath & CurItem)
                        If GetViri <> "" Then
                        LvwSubStyle frmRTP.lvVirus, "W32:Elektrik.A", RealPath & CurItem, "Internal Virus"
                        End If
                    End If
            nSize = MAX_VIRUSNAME_LENGTH
            RetVal = CheckWithAnsav(RealPath & CurItem, Buff, nSize)
            If RetVal = 1 Then
            NamaVirus = Left$(Buff, nSize)
            LvwSubStyle frmRTP.lvVirus, NamaVirus, RealPath & CurItem, "Engine Detektor"
            End If
            
            RTPCekDBExternal CStr(RealPath & CurItem)
            MeiPattern CStr(RealPath & CurItem), 202
            
            'Ini Untuk By.USER Virus
            Dim s As Integer
            If frmMain.ListView2.ListItems.Count <> 0 Then
            For s = 1 To frmMain.ListView2.ListItems.Count
            If CurItem = GetFileName(frmMain.ListView2.ListItems(s).Text) Then
            LvwSubStyle frmRTP.lvVirus, frmMain.Text2.Text, RealPath & CurItem, "Dari User"
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
                    ScanRTP tempDir(i), True
                Next i
            End If
        End If
    End If
End Sub
