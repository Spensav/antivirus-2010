Attribute VB_Name = "modChecksumPattern"
Public Function GetChecksum(FilePath As String) As String
Dim CheckSum(1 To 2) As String
CheckSum(1) = CalcBinary(FilePath, 499, 4500)
CheckSum(2) = CalcBinary(FilePath, 499, 4000)
GetChecksum = CheckSum(1) & CheckSum(2)
End Function
Public Function CalcBinary(ByVal lpFilename As String, ByVal lpByteCount As Long, Optional ByVal StartByte As Long = 0) As String
On Error GoTo err
Dim Bin() As Byte
Dim ByteSum As Long
Dim I As Long
ReDim Bin(lpByteCount) As Byte
Open lpFilename For Binary As #1
    If StartByte = 0 Then
        Get #1, , Bin
    Else
        Get #1, StartByte, Bin
    End If
Close #1
For I = 0 To lpByteCount
    ByteSum = ByteSum + Bin(I) ^ 2
Next I
CalcBinary = Hex$(ByteSum)
Exit Function
err:
CalcBinary = "00"
End Function
