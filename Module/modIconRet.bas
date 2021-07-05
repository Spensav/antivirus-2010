Attribute VB_Name = "modIconRet"
'fungsi" yang dibutuhkan untuk icon compare
Option Explicit
Private Const SHGFI_DISPLAYNAME = &H200, SHGFI_EXETYPE = &H2000, SHGFI_SYSICONINDEX = &H4000, SHGFI_LARGEICON = &H0, SHGFI_SMALLICON = &H1, SHGFI_SHELLICONSIZE = &H4, SHGFI_TYPENAME = &H400, ILD_TRANSPARENT = &H1, BASIC_SHGFI_FLAGS = SHGFI_TYPENAME Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE
Public Type SHFILEINFO
    hIcon As Long: iIcon As Long: dwAttributes As Long: szDisplayName As String * MAX_PATH: szTypeName As String * 80
End Type
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbSizeFileInfo As Long, ByVal uFlags As Long) As Long
Private Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl As Long, ByVal i As Long, ByVal hdcDest As Long, ByVal X As Long, ByVal Y As Long, ByVal flags As Long) As Long
Private shinfo As SHFILEINFO, sshinfo As SHFILEINFO
Private Declare Function DrawIconEx Lib "user32.dll" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private SIconInfo As SHFILEINFO

Public Enum IconRetrieve
    ricnLarge = 32
    ricnSmall = 16
End Enum

Public Sub RetrieveIcon(fName As String, DC As PictureBox, icnSize As IconRetrieve)
    Dim hImgSmall, hImgLarge As Long
    Debug.Print fName
    Select Case icnSize
    Case ricnSmall
        hImgSmall = SHGetFileInfo(fName$, 0&, shinfo, Len(shinfo), BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
        Call ImageList_Draw(hImgSmall, shinfo.iIcon, DC.hdc, 0, 0, ILD_TRANSPARENT)
    Case ricnLarge
        hImgLarge& = SHGetFileInfo(fName$, 0&, shinfo, Len(shinfo), BASIC_SHGFI_FLAGS Or SHGFI_LARGEICON)
        Call ImageList_Draw(hImgLarge, shinfo.iIcon, DC.hdc, 0, 0, ILD_TRANSPARENT)
    End Select
End Sub
Public Function ExtractIcon(Filename As String, AddtoImageList As ImageList, PictureBox As PictureBox, PixelsXY As IconRetrieve, iKey As String) As Long
    Dim SmallIcon As Long
    Dim NewImage As ListImage
    Dim IconIndex As Integer
    On Error GoTo Load_New_Icon
    If iKey <> "Application" And iKey <> "Shortcut" Then
        ExtractIcon = AddtoImageList.ListImages(iKey).Index
        Exit Function
    End If
Load_New_Icon:
    On Error GoTo Reset_Key
    RetrieveIcon Filename, PictureBox, PixelsXY
    IconIndex = AddtoImageList.ListImages.Count + 1
    Set NewImage = AddtoImageList.ListImages.Add(IconIndex, iKey, PictureBox.Image)
    ExtractIcon = IconIndex
    Exit Function
Reset_Key:
    iKey = ""
    Resume
End Function
Public Sub GetLargeIcon(icPath$, pDisp As PictureBox)
Dim hImgLrg&: hImgLrg = SHGetFileInfo(icPath$, 0&, SIconInfo, Len(SIconInfo), BASIC_SHGFI_FLAGS Or SHGFI_LARGEICON)
 ImageList_Draw hImgLrg, SIconInfo.iIcon, pDisp.hdc, 0, 0, ILD_TRANSPARENT
End Sub



