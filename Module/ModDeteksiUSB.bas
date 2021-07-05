Attribute VB_Name = "ModDeteksiUSB"
Public Function DeteksiUSBSekarang()
Dim Drive() As String
Dim Drives  As String
Dim i       As Long
On Error GoTo Out
Drive = GetDrives
Drives = Join(Drive, "|")
If Len(Drives) = Len(mLastDrives) Then
    Exit Function
ElseIf Len(Drives) < Len(mLastDrives) Then
    mLastDrives = Drives
    Exit Function
End If
For i = 0 To UBound(Drive)
    If InStr(1, mLastDrives, Drive(i)) = 0 Then
        If GetDriveTypeW(StrPtr(Drive(i))) = DRIVE_REMOVABLE Then
            frmUSB.Label4.Caption = ": " & Drive(i) & "-" & DriveLabel(Drive(i))
            frmUSB.Label5.Caption = ": " & GetDiskSpace(Drive(i))
            frmUSB.Label10.Caption = Drive(i)
            frmUSB.Show
            GradientPic frmUSB.Frame1, &H80&, &HC0&, gmVertical
            GradientPic frmUSB.Frame2, &H80&, &HC0&, gmVertical
            'MsgBox "Ada flashdisk masuk gan -> " & Drive(i), vbInformation
        End If
    End If
Next i
mLastDrives = Drives
DoEvents
Out:
End Function
Private Function GetDrives() As String()
Dim lpBuffer  As String
Dim nLong     As Long
lpBuffer = String$(MAX_PATH, 0)
nLong = GetLogicalDriveStringsW(MAX_PATH, StrPtr(lpBuffer))
lpBuffer = Left$(lpBuffer, nLong - 1)
GetDrives = Split(lpBuffer, vbNullChar)
End Function
