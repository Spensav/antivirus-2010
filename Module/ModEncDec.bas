Attribute VB_Name = "ModEncrpytDescypt"
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hrgn As Long, ByVal bRedraw As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32.dll" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetClassLong Lib "user32" Alias "GetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Const CS_DROPSHADOW As Long = &H20000
Private Const GCL_STYLE     As Long = -26
Public Sub BuatOval(frm As Form, Optional ByVal Kurva As Double = 35)
    On Error Resume Next
    Dim hrgn      As Long, hrgn2 As Long
    Dim X1        As Long, Y1 As Long
    X1 = frm.Width / Screen.TwipsPerPixelX
    Y1 = frm.Height / Screen.TwipsPerPixelY
    hrgn = CreateRoundRectRgn(0, 0, X1, Y1, Kurva, Kurva)
    SetWindowRgn frm.hwnd, hrgn, True
    DeleteObject hrgn
End Sub
Public Sub DropShadow(ByVal hwnd As Long)
On Error Resume Next
    Call SetClassLong(hwnd, GCL_STYLE, GetClassLong(hwnd, GCL_STYLE) Or CS_DROPSHADOW)
End Sub
Public Function EncodeFile(SourceFile As String, DestFile As String)
    Dim ByteArray() As Byte, Filenr As Integer
    Filenr = FreeFile
    Open SourceFile For Binary As #Filenr
        ReDim ByteArray(0 To LOF(Filenr) - 1)
        Get #Filenr, , ByteArray()
    Close #Filenr
    Call Coder(ByteArray())
    If (PathFileExists(DestFile)) <> 0 Then DeleteFile DestFile
    Open DestFile For Binary As #Filenr
        Put #Filenr, , ByteArray()
    Close #Filenr
End Function
Public Function DecodeFile(SourceFile As String, DestFile As String)
    Dim ByteArray() As Byte, Filenr As Integer
    Filenr = FreeFile
    Open SourceFile For Binary As #Filenr
        ReDim ByteArray(0 To LOF(Filenr) - 1)
        Get #Filenr, , ByteArray()
    Close #Filenr
    Call DeCoder(ByteArray())
    If (PathFileExists(DestFile)) <> 0 Then DeleteFile DestFile
    Open DestFile For Binary As #Filenr
        Put #Filenr, , ByteArray()
    Close #Filenr
End Function
Private Sub Coder(ByteArray() As Byte)
    Dim X As Long
    Dim value As Integer
    value = 0
    For X = 0 To UBound(ByteArray)
        value = value + ByteArray(X)
        If value > 255 Then value = value - 256
        ByteArray(X) = value
    Next
End Sub
Private Sub DeCoder(ByteArray() As Byte)
    Dim X As Long
    Dim value As Integer
    Dim NewValue As Integer
    NewValue = 0
    For X = 0 To UBound(ByteArray)
        value = NewValue
        NewValue = ByteArray(X)
        value = ByteArray(X) - value
        If value < 0 Then value = value + 256
        ByteArray(X) = value
    Next
End Sub

