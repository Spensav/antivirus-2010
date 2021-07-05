Attribute VB_Name = "ModStartupMgr2"
Const RC_PALETTE As Long = &H100
Const SIZEPALETTE As Long = 104
Const RASTERCAPS As Long = 38

Private Type PALETTEENTRY
    peRed As Byte
    peGreen As Byte
    peBlue As Byte
    peFlags As Byte
End Type

Private Type LOGPALETTE
    palVersion As Integer
    palNumEntries As Integer
    palPalEntry(255) As PALETTEENTRY
End Type

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

Private Type PicBmp
    Size As Long
    Type As Long
    hBmp As Long
    hPal As Long
    Reserved As Long
End Type


Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal iCapabilitiy As Long) As Long
Private Declare Function GetSystemPaletteEntries Lib "gdi32" (ByVal hdc As Long, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long
Private Declare Function CreatePalette Lib "gdi32" (lpLogPalette As LOGPALETTE) As Long
Private Declare Function SelectPalette Lib "gdi32" (ByVal hdc As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long
Private Declare Function RealizePalette Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long

Rem
Const SHGFI_DISPLAYNAME = &H200
Const SHGFI_TYPENAME = &H400
Const SHFI_ICON = &H100
Const SHGFI_SMALLICON = &H1
Const MAX_PATH = 260
Const DefaultPath = "C:\WinNt"


Private Type SHFILEINFO
    hIcon As Long
    iIcon As Long
    dwAttributes As Long
    szDisplayName As String * MAX_PATH
    szTypeName As String * 80
End Type

Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long
Private FI As SHFILEINFO
Public dlv As Integer
Public Sub GetIconsDe(lstView As ListView, imaList As ImageList, picTmp As PictureBox)
    On Error Resume Next
        Dim lsv As ListItem
    For Each lsv In lstView.ListItems
            picTmp.Cls
            GetIcon lsv.SubItems(2), picTmp
            If lsv.Index = dlv Then imaList.ListImages.Add lsv.Index, , picTmp.Image
    Next
        
    With lstView
      .SmallIcons = imaList
      For Each lsv In .ListItems
        lsv.SmallIcon = lsv.Index
      Next
    End With
End Sub
Private Function StripTerminator(sInput As String) As String
    Dim ZeroPos As Integer
    ZeroPos = InStr(1, sInput, vbNullChar)
    If ZeroPos > 0 Then
        StripTerminator = Left$(sInput, ZeroPos - 1)
    Else
        StripTerminator = sInput
    End If
End Function

Public Function ConvertIcon(hIcon) As Picture
  On Error GoTo errore
      If hIcon = 0 Then Exit Function
          Dim NewPic As Picture, PicConv As PicBmp, IGuid As GUID
              With PicConv
                .Size = Len(PicConv)
                .Type = vbPicTypeIcon
                .hBmp = hIcon
              End With
                  IGuid.Data1 = &H20400
                  IGuid.Data4(0) = &HC0
                  IGuid.Data4(7) = &H46
      Call OleCreatePictureIndirect(PicConv, IGuid, True, NewPic)
        Set ConvertIcon = NewPic
            Exit Function
errore:
 MsgBox err.Description, vbCritical, App.EXEName & ":Errore di conversione"
End Function
Public Function GetPictureFromIkon(ByVal OpenFile As String, ByVal Size As Long, ByRef picDefaultIcon As Picture) As Picture

  Dim lResult As Picture

  Select Case Size
    Case 16
      SHGetFileInfo OpenFile, 0, FI, Len(FI), SHGFI_DISPLAYNAME Or SHGFI_TYPENAME Or SHFI_ICON Or _
                    SHGFI_SMALLICON
    Case 32
      SHGetFileInfo OpenFile, 0, FI, Len(FI), SHGFI_DISPLAYNAME Or SHGFI_TYPENAME Or SHFI_ICON
  End Select
  
  If FI.hIcon <> 0 Then
    Set GetPictureFromIkon = ConvertIcon(FI.hIcon)
  Else
    Set GetPictureFromIkon = picDefaultIcon
  End If
  'Set GetPictureFromIkon = ctlResult
  'Set lResult = Nothing

End Function




