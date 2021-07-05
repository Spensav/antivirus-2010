Attribute VB_Name = "ModReadFile"
Private Declare Function DeleteFile Lib "KERNEL32" Alias "DeleteFileW" (ByVal lpFilename As Long) As Long
Private Declare Function SetFileAttributes Lib "KERNEL32" Alias "SetFileAttributesW" (ByVal lpFilename As Long, ByVal dwFileAttributes As Long) As Long
Public Declare Function GetFileAttributes Lib "KERNEL32" Alias "GetFileAttributesW" (ByVal lpFilename As Long) As Long
Dim RDF As New clsFile
Public Function ReadUnicodeFile(ByRef sFilePath As String) As String
On Error GoTo TERAKHIR
Dim zFileName   As String
Dim hFile       As Long 'nomor file handle, valid jika > 0;
Dim nFileLen    As Long
Dim nOperation  As Long

    zFileName = sFilePath

    hFile = RDF.VbOpenFile(zFileName, FOR_BINARY_ACCESS_READ, LOCK_NONE)
    'selanjutnya:
    If hFile > 0 Then 'jika berhasil membuka file hFile/Handel file > 0;
        'cari tahu ukuran filenya:
        nFileLen = RDF.VbFileLen(hFile)
        If nFileLen > 140000000 Then Exit Function ' nyerah aja klo file-nya lebih dari 130.000.000 B
        Dim BufData()   As Byte
            nOperation = RDF.VbReadFileB(hFile, 1, nFileLen, BufData)
            ReadUnicodeFile = StrConv(BufData, vbUnicode) ' Ralat pada buku tadinya Cstr(buffdata)
            RDF.VbCloseFile hFile 'harus tutup handle ke file setelah mengaksesnya !!!
        Erase BufData()
    Else 'jika gagal membuka file;
            GoTo TERAKHIR
    End If
Exit Function

TERAKHIR:
End Function
Public Function ReadAnsiFile(sFile As String) As String
Dim sTemp As String
Open sFile For Binary As #1
    sTemp = Space(LOF(1))
    Get #1, , sTemp
Close #1
ReadAnsiFile = sTemp
End Function
Public Function NormalizeAttribute(spath As String)
On Error Resume Next
If GetFileAttributes(StrPtr(spath)) = 4 Then ' system
   SetFileAttributes StrPtr(spath), 0
ElseIf GetFileAttributes(StrPtr(spath)) = 6 Then ' hidden + system
   SetFileAttributes StrPtr(spath), 0
ElseIf GetFileAttributes(StrPtr(spath)) = 2 Then '
   SetFileAttributes spath, 0
ElseIf GetFileAttributes(StrPtr(spath)) = 38 Then '
   SetFileAttributes StrPtr(spath), 0
ElseIf GetFileAttributes(StrPtr(spath)) = 39 Then '
   SetFileAttributes StrPtr(spath), 0
End If
End Function

