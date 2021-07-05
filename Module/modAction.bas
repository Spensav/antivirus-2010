Attribute VB_Name = "modAction"
Private Type Signature
    sampel(2000) As String
    hash(1000) As String
    NamaVirus(2000) As String
End Type
Private sign As Signature
Public Function GetFileName(PathFile As String) As String
Dim i As Long
Dim DirString As Long
    For i = 1 To Len(PathFile)
        If Mid$(PathFile, i, 1) = "\" Then DirString = i
    Next i
    GetFileName = Right$(PathFile, Len(PathFile) - DirString)
End Function
Public Function GetExt(ByVal lpFileName As String)
Dim sTemp As String
Dim i As Long
sTemp = GetFileName(lpFileName)
    If InStr(lpFileName, ".") Then
        For i = 0 To Len(sTemp) - 1
            If Mid$(sTemp, Len(sTemp) - i, 1) = "." Then
                GetExt = Mid$(sTemp, Len(sTemp) - i, i)
                Exit Function
            End If
        Next i
    End If
End Function
Public Function DatabaseEx()
i = 1
'Mengambil signature dari file
Open App.path & "\Signature.dat" For Input As #1
    Do
    Input #1, sign.sampel(i)
    sign.NamaVirus(i) = Mid(sign.sampel(i), InStr(1, sign.sampel(i), "+") + 1, Len(Mid(sign.sampel(i), InStr(1, sign.sampel(i), "+") + 1)))
    If sign.NamaVirus(i) = "" Then Exit Do
    lsDB.AddItem (i & ". " & sign.NamaVirus(i))
    i = i + 1
    Loop Until i = i + 1
Close #1
frmMain.Label9(1).Caption = ": " & lsDB.ListCount & " Virus"
End Function
