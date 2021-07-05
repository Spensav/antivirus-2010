Attribute VB_Name = "modHeur"
Private Declare Function ExtractIconEx Lib "shell32.dll" Alias "ExtractIconExA" (ByVal lpszFile As String, ByVal nIconIndex As Long, ByRef phiconLarge As Long, ByRef phiconSmall As Long, ByVal nIcons As Long) As Long
Private Declare Function DestroyIcon Lib "user32.dll" (ByVal hIcon As Long) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Boolean
Private Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExefileName As String, ByVal nIconIndex As Long) As Long
Private Const DI_MASK = &H1
Private Const DI_IMAGE = &H2
Private Const DI_NORMAL = &H3
Private Const DI_COMPAT = &H4
Private Const DI_DEFAULTSIZE = &H8
Private Const SHGFI_DISPLAYNAME = &H200
Private Const SHGFI_EXETYPE = &H2000
Private Const SHGFI_SYSICONINDEX = &H4000
Private Const SHGFI_LARGEICON = &H0
Private Const SHGFI_SMALLICON = &H1
Private Const ILD_TRANSPARENT = &H1
Private Const SHGFI_SHELLICONSIZE = &H4
Private Const SHGFI_TYPENAME = &H400
Private Const BASIC_SHGFI_FLAGS = SHGFI_TYPENAME Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE
Private SIconInfo As SHFILEINFO
Private SectionHeaders() As IMAGE_SECTION_HEADER
Dim i As Integer
Dim j As Integer
Public Function CekHeuristic(Filename As String)
CekHeuristic = ""
On Error GoTo hError
Dim hFile As Long, bRW As Long
Dim DOSheader As IMAGE_DOS_HEADER
Dim NTHeaders As IMAGE_NT_HEADERS
Dim Filedata As String
DOS_HEADER_INFO = ""
NT_HEADERS_INFO = ""
hFile = CreateFile(Filename, ByVal (GENERIC_READ Or GENERIC_WRITE), FILE_SHARE_READ, ByVal 0&, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, ByVal 0)
ReadFile hFile, DOSheader, Len(DOSheader), bRW, ByVal 0&
SetFilePointer hFile, DOSheader.e_lfanew, 0, 0
ReadFile hFile, NTHeaders, Len(NTHeaders), bRW, ByVal 0&
If NTHeaders.Signature <> IMAGE_NT_SIGNATURE Then
'###### file tidak valid PE ######'
If IsScript(Filename) = True Then
Open Filename For Binary As #1
Filedata = Space$(LOF(1))
Get #1, , Filedata
Close #1
CekHeuristic = CekHeur(Filedata)
End If
Exit Function
End If
'######## PE file Valid(executable) ##########'
CekHeuristic = CekIconBinary(Filename)
hError:
End Function
Private Function CekHeur(Data As String)
Dim hsl, asl As Integer
strasli = LCase(Replace(Data, vbNewLine, "$"))
For i = 1 To UBound(Bahaya)
hsl = 0
strData = Split(Bahaya(i), "|")
asl = 0
For K = 0 To UBound(strData)
xxx = LCase(strData(K))
If InStr(strasli, xxx) > 0 Then hsl = hsl + 1
asl = asl + 1
Next
If hsl = asl Then
CekHeur = "Malicious-Script"
Exit Function
End If
Next
CekHeur = ""
End Function
Private Function CekIconBinary(PathFile As String)
Dim q As Integer
Dim IconIDNow As String
        CekIconBinary = ""
    IconIDNow = CalcIcon(PathFile)
    If IconIDNow = "" Then Exit Function
        For q = 1 To UBound(IconDB)
            If IconDB(q) = IconIDNow Then
                CekIconBinary = "Malicious-Icon"
                Exit Function
            End If
        Next q
End Function
Private Function CalcBinary(ByVal lpFileName As String, ByVal lpByteCount As Long, Optional ByVal StartByte As Long = 0) As String
Dim Bin() As Byte
Dim ByteSum As Long
Dim i As Long
ReDim Bin(lpByteCount) As Byte
Open lpFileName For Binary As #1
    If StartByte = 0 Then
        Get #1, , Bin
    Else
        Get #1, StartByte, Bin
    End If
Close #1
For i = 0 To lpByteCount
    ByteSum = ByteSum + Bin(i) ^ 2
Next i
CalcBinary = Hex$(ByteSum)
End Function
Private Function CalcIcon(ByVal lpFileName As String) As String
Dim PicPath As String
Dim ByteSum As String
Dim IconExist As Long
Dim hIcon As Long
IconExist = ExtractIconEx(lpFileName, 0, ByVal 0&, hIcon, 1)
If IconExist <= 0 Then
    IconExist = ExtractIconEx(lpFileName, 0, hIcon, ByVal 0&, 1)
    If IconExist <= 0 Then Exit Function
End If
frmMain.sIcon.BackColor = vbWhite
DrawIconEx frmMain.sIcon.hdc, 0, 0, hIcon, 0, 0, 0, 0, DI_NORMAL
DestroyIcon hIcon
PicPath = Environ$("windir") & "\tmp.tmp"
SavePicture frmMain.sIcon.Image, PicPath
ByteSum = CalcBinary(PicPath, FileLen(PicPath))
DeleteFile PicPath
CalcIcon = ByteSum
End Function

