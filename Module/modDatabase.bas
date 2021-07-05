Attribute VB_Name = "modDatabaseSpensav"
Public sMD5() As String
Public sNamaVirus() As String
Public JumlahViruss As Integer

Public VirusDB(13), IconDB(48), Bahaya(5) As String
Public Function BacaDatabase(sPath As String)
Static sTemp As String
Static sTmp() As String
Static sTmp2() As String
Static pisah As String
Static iCount As Integer
Static iTemp As Integer
'Dim sXOR As New clsSimpleXOR
'sXOR.DecryptFile sPath, sPath, "spensas 01"

pisah = Chr(13)
sTemp = ReadAnsiFile(sPath) ' boleh diganti fungsi ReadUnicodeFile
sTmp() = Split(sTemp, pisah)

iTemp = UBound(sTmp()) - 1 ' untuk jumlah virus

ReDim sMD5(iTemp) As String
ReDim sNamaVirus(iTemp) As String

For iCount = 1 To iTemp
    sTmp2() = Split(sTmp(iCount), "+")
    sMD5(iCount) = Mid(sTmp2(0), 2)
    sNamaVirus(iCount) = sTmp2(1)
Next
JumlahViruss = iTemp

End Function
Public Function isFileVirus(sPath As String) As Boolean
Static iCount As Integer
Static MD5file As String
MD5file = frmMain.Text5.Text

For iCount = 1 To JumlahViruss
    If sMD5(iCount) = MD5file Then ' jika virus didapet
    LvwSubStyle frmMain.lvVirus, sNamaVirus(iCount), sPath, "External Virus"
        'MsgBox "Virus Terdeteksi >>> " & sPath, vbCritical, "AVBsOFT"
        isFileVirus = True
        Exit Function
    End If
Next
isFileVirus = False
End Function
Public Function RTPCekDBExternal(sPath As String) As Boolean
Static iCount As Integer
Static MD5file As String
MD5file = frmMain.Text5.Text

For iCount = 1 To JumlahViruss
    If sMD5(iCount) = MD5file Then ' jika virus didapet
    LvwSubStyle frmRTP.lvVirus, sNamaVirus(iCount), sPath, "External Virus"
        'MsgBox "Virus Terdeteksi >>> " & sPath, vbCritical, "AVBsOFT"
        RTPCekDBExternal = True
        Exit Function
    End If
Next
RTPCekDBExternal = False
End Function
Public Sub BuildDatabase()
Call Checksum_DB
Call IconCompare_DB
Call Script_DB
End Sub
Private Sub Checksum_DB()
VirusDB(1) = "Alman.A|8911D290F723"
VirusDB(2) = "Malingsi.A|A6292EA60230"
VirusDB(3) = "Conficker.A|9EC112ABB2F3"
VirusDB(4) = "N4B3.A|B5CCD36CDB98"
VirusDB(5) = "N4B3.B|A1FE6D6DBE07"
VirusDB(6) = "N4B3.C|B1AB1975C444"
VirusDB(7) = "Hakalan|82DCED79B484"
VirusDB(8) = "V3M0.A|B08A258298EF"
VirusDB(9) = "V3M0.B|A11EF57BF704"
VirusDB(10) = "V3M0.C|B03F9E78D587"
VirusDB(11) = "Yuyun.A|3CB6323AD9AE"
VirusDB(12) = "Recycle.A|C46EC255155E"
VirusDB(13) = "Shemale|20DB9D207A8C"
End Sub
Public Sub IconCompare_DB()
On Error Resume Next
IconDB(1) = "20938B2"
IconDB(2) = "19F4ED6"
IconDB(3) = "133BE0B"
IconDB(4) = "18EDEAE"
IconDB(5) = "1EF89C2"
IconDB(6) = "1C915FF"
IconDB(7) = "24563C4"
IconDB(8) = "1B2DB74"
IconDB(9) = "208EA72"
IconDB(10) = "22A064D"
IconDB(11) = "19B64EE"
IconDB(12) = "1D4B7E1"
IconDB(13) = "2087762"
IconDB(14) = "29C7258"
IconDB(15) = "1B18705"
IconDB(16) = "1B5FCAB"
IconDB(17) = "126D4CF"
IconDB(18) = "1C58E5C"
IconDB(19) = "15D7730"
IconDB(20) = "1FB82B7"
IconDB(21) = "112763E"
IconDB(22) = "2165AF9"
IconDB(23) = "25F46BE"
IconDB(24) = "206556B"
IconDB(25) = "22A8D69"
IconDB(26) = "19237F8"
IconDB(27) = "15022B4"
IconDB(28) = "1D8B4EB"
IconDB(29) = "1DBC1EA"
IconDB(30) = "2333F5D"
IconDB(31) = "1F37C2F"
IconDB(32) = "1C9CCA4"
IconDB(33) = "1DFDFB4"
IconDB(34) = "1C1283E"
IconDB(35) = "1F6598C"
IconDB(36) = "27F4C1A"
IconDB(37) = "22F92E0"
IconDB(38) = "191DBDC"
IconDB(39) = "27BFE4A"
IconDB(40) = "20E0907"
IconDB(46) = "2FA4C88"
IconDB(47) = "25AA630"
IconDB(48) = "1DE28E2"
End Sub
Public Sub Script_DB()
On Error Resume Next
    Bahaya(1) = "Scripting.FileSystemObject|Wscript.ScriptFullName|WScript.Shell|.regwrite|.copy"
    Bahaya(2) = "Wscript.ScriptFullName|createobject|strreverse|.regwrite"
    Bahaya(3) = "createobject|Wscript.ScriptFullName|.regwrite|[autorun]"
    Bahaya(4) = "createobject|Wscript.ScriptFullName|specialfolder|.regwrite"
    Bahaya(5) = "chr(asc(mid(|createobject|Wscript.ScriptFullName|.GetFolder|.RegWrite"
End Sub
Public Function MeiPattern(Path As String, Box As Long) As String ' Box = 202
' ini yang lebih lengkap
On Error Resume Next
Dim Temp() As Byte
Dim Temp2() As Byte
Dim X, X2 As Long
Dim Num, num2 As Long
 ReDim Temp(Box) As Byte
 ReDim Temp2(Box) As Byte
 If FileLen(Path) >= 4520 Then
      Open Path For Binary As #1
          Get #1, 3911, Temp
          Get #1, 4311, Temp2
      Close #1
      For Num = 1 To UBound(Temp)
          X = X + Temp(Num) ^ 3
      Next
      For num2 = 1 To UBound(Temp2)
          X2 = X2 + Temp2(num2) ^ 3
      Next
 Else
      Open Path For Binary As #1
          Get #1, , Temp
          Get #1, 401, Temp2
      Close #1
      For Num = 1 To 202
          X = X + Temp(Num) ^ 3
      Next
      For num2 = 1 To 202
          X2 = X2 + Temp2(num2) ^ 3
      Next
 End If
'Ini Untuk Scannya :D
frmMain.Text5.Text = Hex(X) & Hex(X2)
End Function
