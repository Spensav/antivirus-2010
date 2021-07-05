VERSION 5.00
Begin VB.UserControl XPFrame 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C9F1FC&
   ClientHeight    =   2700
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3525
   ControlContainer=   -1  'True
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   180
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   235
   ToolboxBitmap   =   "xpframe.ctx":0000
End
Attribute VB_Name = "XPFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
'*************************************************************************
'* XP-Style Gradient Container Control                                   *
'* Original Control - Cameron Groves                                     *
'* API modifications, curvature mods - Jim Jose                          *
'* Gradient - Redbird77 txtCodeID=59020                                  *
'* XP-Icon code - Carles P.V.                                            *
'* Whining, crying, begging for help, a tiny bit of coding               *
'* and assemblage -  Matthew R. Usner                                    *
'*                                                                       *
'* Carles P.V. supplied the .Res file and code to allow the              *
'* display of XP Icons.  Thanks so much Carles.  Now maybe               *
'* "enmity" will get off my back. :)                                     *
'*************************************************************************

Option Explicit

Private Declare Sub RtlMoveMemory Lib "kernel32" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function MoveTo Lib "gdi32" Alias "MoveToEx" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, lpPoint As Any) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long

Private Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal hPalette As Long, pccolorref As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function DrawText Lib "user32.dll" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hrgn As Long, ByVal bRedraw As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32.dll" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function FillRgn Lib "gdi32.dll" (ByVal hDC As Long, ByVal hrgn As Long, ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function CombineRgn Lib "gdi32.dll" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long

Private Type RECT
   Left    As Long
   Top     As Long
   Right   As Long
   Bottom  As Long
End Type

Public Enum XPContainerStyles
   [Header Visible] = 0
   [Header Invisible] = 1
End Enum

' Property Variables
Private m_HeaderAngle       As Integer
Private m_BackAngle         As Integer ' background gradient display angle
Private m_Style             As XPContainerStyles
Private m_HeaderLightColor  As OLE_COLOR
Private m_HeaderDarkColor   As OLE_COLOR
Private m_BackLightColor    As OLE_COLOR
Private m_BackDarkColor     As OLE_COLOR
Private m_BorderColor       As OLE_COLOR
Private m_TextColor         As OLE_COLOR
Private m_Caption           As String
Private m_HeaderHeight      As Long
Private m_HeaderFont        As StdFont
Private m_Alignment         As AlignmentConstants
Private m_hMod              As Long
Private m_Curvature         As Long
Private m_IconResID         As Variant

'[Default Property Values/Constants]
Private Const m_def_HeaderAngle = 180     ' init to horizontal header gradient
Private Const m_def_BackAngle = 180       ' init to horizontal bg gradient
Private Const m_DEF_HeaderLightColor = &HF7E0D3
Private Const m_DEF_HeaderDarkColor = &HEDC5A7
Private Const m_DEF_BackLightColor = &HFCF4EF
Private Const m_DEF_BackDarkColor = &HFAE8DC
Private Const M_DEF_Caption = "XPFrame"
Private Const m_DEF_BorderColor = &HDCC1AD
Private Const m_DEF_Align = vbLeftJustify
Private Const m_DEF_TextColor = &H7B2D02
Private Const m_DEF_Curvature = 10
Private Const m_DEF_hHeight = 25
Private Const CLR_INVALID = -1
Private Const M_DEF_STYLE = 0

'[ Events ]
Public Event Click()
Public Event Resize()
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseClick(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

Private Declare Function CreateDIBSection32 Lib "gdi32" Alias "CreateDIBSection" (ByVal hDC As Long, lpBitsInfo As BITMAPINFOHEADER, ByVal wUsage As Long, lpBits As Long, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function GetObjectType Lib "gdi32" (ByVal hgdiobj As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function StretchDIBits Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As Any, ByVal wUsage As Long, ByVal dwRop As Long) As Long

Private Declare Function FindResourceStr Lib "kernel32" Alias "FindResourceA" (ByVal hInstance As Long, ByVal lpName As String, ByVal lpType As Long) As Long
Private Declare Function LoadLibraryEx Lib "kernel32" Alias "LoadLibraryExA" (ByVal lpLibFileName As String, ByVal hFile As Long, ByVal dwFlags As Long) As Long
Private Declare Function LoadResource Lib "kernel32" (ByVal hInstance As Long, ByVal hResInfo As Long) As Long
Private Declare Function LockResource Lib "kernel32" (ByVal hResData As Long) As Long
Private Declare Function SizeofResource Lib "kernel32" (ByVal hInstance As Long, ByVal hResInfo As Long) As Long

Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (lpDst As Any, ByVal Length As Long)
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, lpSrc As Any, ByVal Length As Long)
Private Declare Function VarPtrArray Lib "msvbvm50" Alias "VarPtr" (Ptr() As Any) As Long

Private Const DIB_RGB_COLORS           As Long = 0
Private Const LOAD_LIBRARY_AS_DATAFILE As Long = &H2
Private Const RT_BITMAP                As Long = 2

Private Type BITMAPINFOHEADER
    biSize          As Long
    biWidth         As Long
    biHeight        As Long
    biPlanes        As Integer
    biBitCount      As Integer
    biCompression   As Long
    biSizeImage     As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed       As Long
    biClrImportant  As Long
End Type

Private Type BITMAP
    bmType       As Long
    bmWidth      As Long
    bmHeight     As Long
    bmWidthBytes As Long
    bmPlanes     As Integer
    bmBitsPixel  As Integer
    bmBits       As Long
End Type

Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound   As Long
End Type

Private Type SAFEARRAY1D
    cDims      As Integer
    fFeatures  As Integer
    cbElements As Long
    cLocks     As Long
    pvData     As Long
    Bounds     As SAFEARRAYBOUND
End Type

'- Private variables (for 32-bit DIB alpha 'icon')
Private m_uBIH    As BITMAPINFOHEADER
Private m_hDC     As Long
Private m_hDIB    As Long
Private m_hOldDIB As Long
Private m_lpBits  As Long

Private Sub RedrawControl()
   UserControl.Cls
   If Style = [Header Visible] Then
      SetBackGround
      SetHeader
      SetBorder
   Else
      SetBackGround
      SetBorder
   End If
End Sub

Private Function TranslateColor(ByVal oClr As OLE_COLOR, Optional hPal As Long = 0) As Long
   If OleTranslateColor(oClr, hPal, TranslateColor) Then
      TranslateColor = CLR_INVALID
   End If
End Function

Private Sub UserControl_Initialize()
   m_hMod = LoadLibrary("shell32.dll") ' Used to prevent crashes on Windows XP
End Sub

Private Sub UserControl_InitProperties()
    m_HeaderAngle = m_def_HeaderAngle
    m_BackAngle = m_def_BackAngle
    m_HeaderLightColor = m_DEF_HeaderLightColor
    m_HeaderDarkColor = m_DEF_HeaderDarkColor
    m_BackLightColor = m_DEF_BackLightColor
    m_BackDarkColor = m_DEF_BackDarkColor
    m_BorderColor = m_DEF_BorderColor
    m_TextColor = m_DEF_TextColor
    m_Caption = M_DEF_Caption
    m_Style = M_DEF_STYLE
    m_Alignment = vbLeftJustify
    m_Curvature = m_DEF_Curvature
    m_HeaderHeight = m_DEF_hHeight
    Set m_HeaderFont = UserControl.Font
    m_IconResID = 0
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
   m_HeaderAngle = PropBag.ReadProperty("HeaderAngle", m_def_HeaderAngle)
   m_BackAngle = PropBag.ReadProperty("BackAngle", m_def_BackAngle)
   m_HeaderLightColor = PropBag.ReadProperty("HeaderLightColor", m_DEF_HeaderLightColor)
   m_HeaderDarkColor = PropBag.ReadProperty("HeaderDarkColor", m_DEF_HeaderDarkColor)
   m_BackLightColor = PropBag.ReadProperty("BackLightColor", m_DEF_BackLightColor)
   m_BackDarkColor = PropBag.ReadProperty("BackDarkColor", m_DEF_BackDarkColor)
   m_BorderColor = PropBag.ReadProperty("BorderColor", m_DEF_BorderColor)
   m_TextColor = PropBag.ReadProperty("TextColor", m_DEF_TextColor)
   m_Caption = PropBag.ReadProperty("Caption", M_DEF_Caption)
   m_Style = PropBag.ReadProperty("Style", M_DEF_STYLE)
   m_Curvature = PropBag.ReadProperty("Curvature", m_DEF_Curvature)
   m_Alignment = PropBag.ReadProperty("CaptionAlignment", m_DEF_Align)
   m_HeaderHeight = PropBag.ReadProperty("HeaderHeight", m_DEF_hHeight)
   Set m_HeaderFont = PropBag.ReadProperty("HeaderFont", UserControl.Font)
   m_IconResID = PropBag.ReadProperty("IconResID", 0)
   If m_IconResID <> 0 Then
      GetResourceBitmap m_IconResID
   End If
End Sub

Private Sub UserControl_Show()
   UserControl_Resize
End Sub

Private Sub UserControl_Terminate()
   DestroyDIB32
   FreeLibrary m_hMod ' Used to prevent crashes on Windows XP
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
   Call PropBag.WriteProperty("HeaderAngle", m_HeaderAngle, m_def_HeaderAngle)
   Call PropBag.WriteProperty("BackAngle", m_BackAngle, m_def_BackAngle)
   Call PropBag.WriteProperty("HeaderLightColor", m_HeaderLightColor, m_DEF_HeaderLightColor)
   Call PropBag.WriteProperty("HeaderDarkColor", m_HeaderDarkColor, m_DEF_HeaderDarkColor)
   Call PropBag.WriteProperty("BackLightColor", m_BackLightColor, m_DEF_BackLightColor)
   Call PropBag.WriteProperty("BackDarkColor", m_BackDarkColor, m_DEF_BackDarkColor)
   Call PropBag.WriteProperty("BorderColor", m_BorderColor, m_DEF_BorderColor)
   Call PropBag.WriteProperty("TextColor", m_TextColor, m_DEF_TextColor)
   Call PropBag.WriteProperty("Caption", m_Caption, M_DEF_Caption)
   Call PropBag.WriteProperty("Style", m_Style, M_DEF_STYLE)
   Call PropBag.WriteProperty("CaptionAlignment", m_Alignment, vbLeftJustify)
   Call PropBag.WriteProperty("HeaderHeight", m_HeaderHeight, m_DEF_hHeight)
   Call PropBag.WriteProperty("HeaderFont", m_HeaderFont, UserControl.Font)
   Call PropBag.WriteProperty("Curvature", m_Curvature, m_DEF_Curvature)
   Call PropBag.WriteProperty("IconResID", m_IconResID, 0)
End Sub

Public Property Get HeaderAngle() As Integer
   HeaderAngle = m_HeaderAngle
End Property

Public Property Let HeaderAngle(ByVal New_HeaderAngle As Integer)
   m_HeaderAngle = New_HeaderAngle
   PropertyChanged "HeaderAngle"
   RedrawControl
End Property

Public Property Get BackAngle() As Integer
   BackAngle = m_BackAngle
End Property

Public Property Let BackAngle(ByVal New_BackAngle As Integer)
   m_BackAngle = New_BackAngle
   PropertyChanged "BackAngle"
   RedrawControl
End Property

Public Property Get BackDarkColor() As OLE_COLOR
   BackDarkColor = m_BackDarkColor
End Property

Public Property Let BackDarkColor(ByVal New_BackDarkColor As OLE_COLOR)
   m_BackDarkColor = New_BackDarkColor
   PropertyChanged "BackDarkColor"
   RedrawControl
End Property

Public Property Get BackLightColor() As OLE_COLOR
   BackLightColor = m_BackLightColor
End Property

Public Property Let BackLightColor(ByVal New_BackLightColor As OLE_COLOR)
   m_BackLightColor = New_BackLightColor
   PropertyChanged "BackLightColor"
   RedrawControl
End Property

Public Property Get BorderColor() As OLE_COLOR
   BorderColor = m_BorderColor
End Property

Public Property Let BorderColor(ByVal New_BorderColor As OLE_COLOR)
   m_BorderColor = New_BorderColor
   PropertyChanged "BorderColor"
   RedrawControl
End Property

Public Property Get Caption() As String
   Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
   m_Caption = New_Caption
   PropertyChanged "Caption"
   RedrawControl
End Property

Public Property Get HeaderDarkColor() As OLE_COLOR
   HeaderDarkColor = m_HeaderDarkColor
End Property

Public Property Let HeaderDarkColor(ByVal New_HeaderDarkColor As OLE_COLOR)
   m_HeaderDarkColor = New_HeaderDarkColor
   PropertyChanged "HeaderDarkColor"
   RedrawControl
End Property

Public Property Get HeaderLightColor() As OLE_COLOR
   HeaderLightColor = m_HeaderLightColor
End Property

Public Property Let HeaderLightColor(ByVal New_HeaderLightColor As OLE_COLOR)
   m_HeaderLightColor = New_HeaderLightColor
   PropertyChanged "HeaderLightColor"
   RedrawControl
End Property

Public Property Get Style() As XPContainerStyles
   Style = m_Style
End Property

Public Property Let Style(ByVal New_Style As XPContainerStyles)
   m_Style = New_Style
   PropertyChanged "Style"
   RedrawControl
End Property

Public Property Get TextColor() As OLE_COLOR
   TextColor = m_TextColor
End Property

Public Property Let TextColor(ByVal New_TextColor As OLE_COLOR)
   m_TextColor = New_TextColor
   PropertyChanged "TextColor"
   RedrawControl
End Property

Public Property Get HeaderHeight() As Long
   HeaderHeight = m_HeaderHeight
End Property

Public Property Let HeaderHeight(ByVal vNewHeight As Long)
   m_HeaderHeight = vNewHeight
   PropertyChanged "HeaderHeight"
   RedrawControl
End Property

Public Property Get HeaderFont() As Font
   Set HeaderFont = m_HeaderFont
End Property

Public Property Set HeaderFont(ByVal vNewHeaderFont As Font)
   Set m_HeaderFont = vNewHeaderFont
   PropertyChanged "HeaderFont"
   RedrawControl
End Property

Public Property Get CaptionAlignment() As AlignmentConstants
   CaptionAlignment = m_Alignment
End Property

Public Property Let CaptionAlignment(ByVal vNewAlignment As AlignmentConstants)
   m_Alignment = vNewAlignment
   PropertyChanged "CaptionAlignment"
   RedrawControl
End Property

Public Property Get IconResID() As Variant
   IconResID = m_IconResID
End Property

Public Property Let IconResID(ByVal vNewIconResID As Variant)
   m_IconResID = vNewIconResID
   If m_IconResID = 0 Then
      Call DestroyDIB32
     Else
       If GetResourceBitmap(IconResID) <> 0 Then
          RedrawControl
         Else
          Err.Raise 1000, , "Resource ID not found or resource file not compiled"
       End If
   End If
   PropertyChanged "IconResID"
End Property

Public Property Get Curvature() As Long
   Curvature = m_Curvature
End Property

Public Property Let Curvature(ByVal vNewCurvature As Long)
   m_Curvature = vNewCurvature
   PropertyChanged "Curvature"
   RedrawControl
End Property

' events
Private Sub UserControl_Resize()
   RedrawControl
   RaiseEvent Resize
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_Click()
   RaiseEvent Click
End Sub

Private Sub SetBackGround()

'*************************************************************************
'* displays the control's background gradient.                           *
'*************************************************************************

   On Error GoTo ErrHandler

   Dim lColor As Long, lColor2 As Long

'  Get the control colors
   lColor = TranslateColor(m_BackLightColor)
   lColor2 = TranslateColor(m_BackDarkColor)
'  Apply the gradients
   DrawGradient hDC, UserControl.ScaleWidth, UserControl.ScaleHeight, lColor, lColor2, m_BackAngle

ErrHandler:
End Sub

Private Sub SetBorder()

'*************************************************************************
'* draws the border around the control, using appropriate curvature      *
'*************************************************************************

   On Error GoTo ErrHandler

   Dim BdrCol As Long
   Dim hBrush As Long
   Dim hrgn1 As Long
   Dim hrgn2 As Long

'  Get the border color
   BdrCol = TranslateColor(m_BorderColor)

'  Define the regions
   hrgn1 = CreateRoundRectRgn(0, 0, ScaleWidth + 1, ScaleHeight + 1, m_Curvature, m_Curvature)
   hrgn2 = CreateRoundRectRgn(1, 1, ScaleWidth, ScaleHeight, m_Curvature, m_Curvature)
   CombineRgn hrgn2, hrgn1, hrgn2, 3

'  Create/Apply the ColorBrush
   hBrush = CreateSolidBrush(BdrCol)
   FillRgn hDC, hrgn2, hBrush

'  Set the control region
   SetWindowRgn hWnd, hrgn1, 0

'  Free the memory
   DeleteObject hrgn1
   DeleteObject hrgn2
   DeleteObject hBrush

ErrHandler:
End Sub

Private Sub SetHeader()

'*************************************************************************
'* displays the header gradient, caption text and an icon if used        *
'*************************************************************************

   On Error GoTo ErrHandler

   Dim lColor  As Long
   Dim lColor2 As Long

'  Get color / Fill gradients
   lColor = TranslateColor(m_HeaderLightColor)
   lColor2 = TranslateColor(m_HeaderDarkColor)
   DrawGradient hDC, UserControl.ScaleWidth, m_HeaderHeight, lColor, lColor2, m_HeaderAngle

'  draw the caption
   Dim R           As RECT
   Dim tHeight     As Long
   Dim tWidth      As Long
   Dim Clearance   As Long

'  Apply the font/Forecolor
   Set UserControl.Font = m_HeaderFont
   tHeight = TextHeight(m_Caption)
   tWidth = TextWidth(m_Caption)
   UserControl.ForeColor = TranslateColor(m_TextColor)

'  make the left clearance one letter width
   Clearance = TextWidth("A")
   With R 'Define the drawing rectangle size
      If m_Alignment = vbCenter Then
         .Left = (ScaleWidth - TextWidth(m_Caption)) / 2
      ElseIf m_Alignment = vbLeftJustify Then
         If m_hDIB <> 0 Then
            .Left = TextWidth("A") + m_HeaderHeight
         Else
            .Left = Clearance
         End If
      Else
         .Left = (ScaleWidth - TextWidth(m_Caption)) - Clearance
      End If
      .Top = (m_HeaderHeight - TextHeight(m_Caption)) / 2
      .Bottom = R.Top + tHeight
      .Right = .Left + tWidth
   End With

'  Draw the caption using API
   DrawText hDC, m_Caption, -1, R, 0

    If m_hDIB <> 0 Then
       AlphaBlend hDC, 2, 2
   End If

ErrHandler:
End Sub

Public Sub DrawGradient(ByVal hDC As Long, ByVal lWidth As Long, ByVal lHeight As Long, _
                        ByVal lCol1 As Long, ByVal lCol2 As Long, ByVal zAngle As Single)

   Dim xStart  As Long, yStart As Long
   Dim xEnd    As Long, yEnd   As Long
   Dim X1      As Long, Y1     As Long
   Dim X2      As Long, Y2     As Long
   Dim lRange  As Long
   Dim iQ      As Integer
   Dim bVert   As Boolean
   Dim lPtr    As Long, lInc   As Long
   Dim lCols() As Long
   Dim hPO     As Long, hPN    As Long
   Dim R       As Long
   Dim X       As Long, xUp    As Long
   Dim b1(2)   As Byte, b2(2)  As Byte
   Dim p       As Single, ip   As Single

   lInc = 1
   xEnd = lWidth - 1
   yEnd = lHeight - 1

'  Positive angles are measured counter-clockwise; negative angles clockwise.
   zAngle = zAngle Mod 360
   If zAngle < 0 Then zAngle = 360 + zAngle

'  Get angle's quadrant (0 - 3).
   iQ = zAngle \ 90

'  Is angle more horizontal or vertical?
   bVert = ((iQ + 1) * 90) - zAngle > 45
   If (iQ Mod 2 = 0) Then bVert = Not bVert

'  Convert angle in degrees to radians.
   zAngle = zAngle * Atn(1) / 45

'  Get start and end y-positions (if vertical), x-positions (if horizontal).
   If bVert Then
      If zAngle Then xStart = lHeight / Abs(Tan(zAngle))
      lRange = lWidth + xStart - 1

      Y1 = IIf(iQ Mod 2, 0, yEnd)
      Y2 = IIf(Y1, -1, lHeight)

      If iQ > 1 Then
         lPtr = lRange: lInc = -1
      End If
   Else
      yStart = lWidth * Abs(Tan(zAngle))
      lRange = lHeight + yStart - 1

      X1 = IIf(iQ Mod 2, 0, xEnd)
      X2 = IIf(X1, -1, lWidth)

      If iQ = 1 Or iQ = 2 Then
         lPtr = lRange: lInc = -1
      End If
   End If

'  -------------------------------------------------------------------
'  Fill in the color array with the interpolated color values.
'  -------------------------------------------------------------------
   ReDim lCols(lRange)

   ' Get the r, g, b components of each color.
   RtlMoveMemory b1(0), lCol1, 3
   RtlMoveMemory b2(0), lCol2, 3

   xUp = UBound(lCols)

   For X = 0 To xUp
      ' Get the position and the 1 - position.
      p = X / xUp
      ip = 1 - p
      ' Interpolate the value at the current position.
      lCols(X) = RGB(b1(0) * ip + b2(0) * p, b1(1) * ip + b2(1) * p, b1(2) * ip + b2(2) * p)
   Next

'  -------------------------------------------------------------------
'  Draw the lines of the gradient at user-specified angle.
'  -------------------------------------------------------------------
   If bVert Then
      For X1 = -xStart To xEnd
         hPN = CreatePen(0, 1, lCols(lPtr))
         hPO = SelectObject(hDC, hPN)
         MoveTo hDC, X1, Y1, ByVal 0&
         LineTo hDC, X2, Y2
         R = SelectObject(hDC, hPO): R = DeleteObject(hPN)
         lPtr = lPtr + lInc
         X2 = X2 + 1
      Next
   Else
      For Y1 = -yStart To yEnd
         hPN = CreatePen(0, 1, lCols(lPtr))
         hPO = SelectObject(hDC, hPN)
         MoveTo hDC, X1, Y1, ByVal 0&
         LineTo hDC, X2, Y2
         R = SelectObject(hDC, hPO): R = DeleteObject(hPN)
         lPtr = lPtr + lInc
         Y2 = Y2 + 1
      Next
   End If
End Sub

Private Function CreateDIB32(ByVal Width As Long, _
                             ByVal Height As Long _
                             ) As Long

    '-- Destroy previous
    Call DestroyDIB32
    '-- Define DIB header
    With m_uBIH
        .biSize = Len(m_uBIH)
        .biPlanes = 1
        .biBitCount = 32
        .biWidth = Width
        .biHeight = Height
        .biSizeImage = (4 * .biWidth) * .biHeight
    End With
    '-- Create DIB and select into a DC
    m_hDC = CreateCompatibleDC(0)
    If (m_hDC <> 0) Then
        m_hDIB = CreateDIBSection32(m_hDC, m_uBIH, DIB_RGB_COLORS, m_lpBits, 0, 0)
        If (m_hDIB <> 0) Then
            m_hOldDIB = SelectObject(m_hDC, m_hDIB)
          Else
            Call DestroyDIB32
        End If
    End If
    '-- Success
    CreateDIB32 = m_hDIB
End Function

Private Sub DestroyDIB32()
    
    '-- Destroy DIB
    If (m_hDC <> 0) Then
        If (m_hDIB <> 0) Then
            Call SelectObject(m_hDC, m_hOldDIB)
            Call DeleteObject(m_hDIB)
        End If
        Call DeleteDC(m_hDC)
    End If
    '-- Reset BIH structure
    Call ZeroMemory(m_uBIH, Len(m_uBIH))
    '-- Reset DIB vars.
    m_hDC = 0: m_hDIB = 0: m_hOldDIB = 0: m_lpBits = 0
End Sub

Private Function GetResourceBitmap(ByVal ResID As Variant _
                                   ) As Long
  Dim sFilename As String
  Dim hInstance As Long
  Dim hInfo     As Long
  Dim hData     As Long
  Dim lSize     As Long

  Dim uBIH      As BITMAPINFOHEADER
  Dim lpResHDR  As Long
  Dim lpResBMP  As Long
  
    '-- Get app. EXE full path
    sFilename = App.Path & IIf(Right$(App.Path, 1) = "\", vbNullString, "\")
    sFilename = sFilename & App.exename & ".exe"

    '-- File exists [?]
    On Error GoTo ErrHandler
    If (FileLen(sFilename)) Then
    On Error GoTo 0
        '-- Get handle to the mapped executable module
        hInstance = LoadLibraryEx(sFilename, 0, LOAD_LIBRARY_AS_DATAFILE)
        If (hInstance) Then
            '-- Get resource info handle
            hInfo = FindResourceStr(hInstance, IIf(IsNumeric(ResID), "#", vbNullString) & ResID, RT_BITMAP)
            If (hInfo) Then
                '-- Get handle to DIB data
                hData = LoadResource(hInstance, hInfo)
                If (hData) Then
                    '-- Get size of DIB data
                    lSize = SizeofResource(hInstance, hInfo)
                    '-- Get pointer to first byte of DIB data (header)
                    lpResHDR = LockResource(hData)
                    '-- Extract DIB info header
                    Call CopyMemory(uBIH, ByVal lpResHDR, Len(uBIH))
                    '-- 32-bit?
                    If (uBIH.biBitCount = 32) Then
                        '-- Create DIB / fill data
                        If (CreateDIB32(uBIH.biWidth, uBIH.biHeight)) Then
                            lpResBMP = lpResHDR + Len(m_uBIH)
                            With m_uBIH
                                Call CopyMemory(ByVal m_lpBits, ByVal lpResBMP, .biSizeImage)
                            End With
                            '-- Success
                            GetResourceBitmap = m_hDIB
                        End If
                    End If
                End If
            End If
            Call FreeLibrary(hInstance)
        End If
    End If
ErrHandler:
End Function

Private Function AlphaBlend( _
                 ByVal hDC As Long, _
                 ByVal X As Long, _
                 ByVal Y As Long _
                 ) As Long
  
  Dim uBI       As BITMAP
  Dim uBIH      As BITMAPINFOHEADER
  
  Dim lhDC      As Long
  Dim lhDIB     As Long
  Dim lhDIBOld  As Long
  
  Dim R         As Long
  Dim G         As Long
  Dim B         As Long
  Dim A1        As Long
  Dim A2        As Long
  
  Dim uSSA      As SAFEARRAY1D
  Dim aSBits()  As Byte
  Dim uDSA      As SAFEARRAY1D
  Dim aDBits()  As Byte
  Dim lpData    As Long
  
  Dim i         As Long
  Dim iIn       As Long
    
    If (m_hDIB <> 0) Then
        With m_uBIH
            '-- Create a temporary DIB section, select into a DC, and
            '   bitblt destination DC area
            lhDC = CreateCompatibleDC(0)
            lhDIB = CreateDIBSection32(lhDC, m_uBIH, DIB_RGB_COLORS, lpData, 0, 0)
            lhDIBOld = SelectObject(lhDC, lhDIB)
            Call BitBlt(lhDC, 0, 0, .biWidth, .biHeight, hDC, X, Y, vbSrcCopy)
            '-- Map destination color data
            Call pvMapDIBits(uDSA, aDBits(), lpData, .biSizeImage)
            '-- Map source color data
            Call pvMapDIBits(uSSA, aSBits(), m_lpBits, .biSizeImage)
            '-- Blend with destination
            For i = 3 To .biSizeImage - 1 Step 4
                A1 = aSBits(i)
                A2 = &HFF - A1
                iIn = i - 1
                aDBits(iIn) = (A1 * aSBits(iIn) + A2 * aDBits(iIn)) \ &HFF: iIn = iIn - 1
                aDBits(iIn) = (A1 * aSBits(iIn) + A2 * aDBits(iIn)) \ &HFF: iIn = iIn - 1
                aDBits(iIn) = (A1 * aSBits(iIn) + A2 * aDBits(iIn)) \ &HFF
            Next i
            '-- Paint alpha-blended
            AlphaBlend = StretchDIBits(hDC, X, Y, .biWidth, .biHeight, 0, 0, .biWidth, .biHeight, ByVal lpData, m_uBIH, DIB_RGB_COLORS, vbSrcCopy)
        End With
        '-- Unmap
        Call pvUnmapDIBits(aDBits())
        Call pvUnmapDIBits(aSBits())
        '-- Clean up
        Call SelectObject(lhDC, lhDIBOld)
        Call DeleteObject(lhDIB)
        Call DeleteDC(lhDC)
    End If
End Function

Private Sub pvMapDIBits(uSA As SAFEARRAY1D, aBits() As Byte, ByVal lpData As Long, ByVal lSize As Long)
    With uSA
        .cbElements = 1
        .cDims = 1
        .Bounds.lLbound = 0
        .Bounds.cElements = lSize
        .pvData = lpData
    End With
    Call CopyMemory(ByVal VarPtrArray(aBits()), VarPtr(uSA), 4)
End Sub

Private Sub pvUnmapDIBits(aBits() As Byte)
    Call CopyMemory(ByVal VarPtrArray(aBits()), 0&, 4)
End Sub
