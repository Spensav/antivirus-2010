VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGetIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function lstrlen Lib "kernel32.dll" Alias "lstrlenA" (ByVal lpString As String) As Long
Private Declare Sub OleCreatePictureIndirect Lib "oleaut32.dll" (ByRef lpPictDesc As PictDesc, ByRef riid As Guid, ByVal fOwn As Long, ByRef lplpvObj As IPictureDisp)
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, ByRef psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As ESHGetFileInfoFlagConstants) As Long

' API Constants
Private Const ERRORAPI As Long = 0
Private Const MAX_PATH As Long = 260

' API Types
Private Type Guid
    Data1           As Long
    data2           As Integer
    Data3           As Integer
    Data4(0 To 7)   As Byte
End Type

Private Type PictDesc
    cbSizeofStruct  As Long
    picType         As Long
    hImage          As Long
    xExt            As Long
    yExt            As Long
End Type

Private Type SHFILEINFO
    hIcon           As Long ' : icon
    iIcon           As Long ' : icondex
    dwAttributes    As Long ' : SFGAO_ flags
    szDisplayName   As String * MAX_PATH ' : display name (or path)
    szTypeName      As String * 80 ' : type name
End Type

'User-Defined API Enum
Private Enum ESHGetFileInfoFlagConstants
    SHGFI_ATTRIBUTES = &H800        ' get file attributes
    SHGFI_DISPLAYNAME = &H200       ' get display name
    SHGFI_EXETYPE = &H2000          ' get exe type
    SHGFI_ICON = &H100              ' get icon handle and index
    SHGFI_LARGEICON = &H0           ' get file's large icon
    SHGFI_LINKOVERLAY = &H8000      ' add link overlay on the icon
    SHGFI_OPENICON = &H2            ' get file's open icon
    SHGFI_SELECTED = &H10000        ' blend icon with the system highlight color
    SHGFI_SHELLICONSIZE = &H4       ' get shell-sized icon
    SHGFI_SMALLICON = &H1           ' get file's small icon
    SHGFI_SYSICONINDEX = &H4000     ' get icon index from system image list
    SHGFI_TYPENAME = &H400          ' get file type description
    SHGFI_USEFILEATTRIBUTES = &H10  ' use dwFileAttributes parameter
End Enum

Public Enum EFileIconTypes
    LargeIcon = 0
    SmallIcon = 1
End Enum

Public Enum EFileExeTypes
    MSDosApp = 2        ' MS-DOS .EXE, .COM or .BAT file
    NonExecutable = 0   ' Nonexecutable file or an error condition
    Win32Console = 3    ' Win32 console application
    WindowsApp = 1      ' Windows application
End Enum

' Variable Declarations
Private m_bOpenState    As Boolean
Private m_bOverlay      As Boolean
Private m_bSelected     As Boolean
Private m_eIconType     As EFileIconTypes
Private m_lHandle       As Long
Private m_sFile         As String

' //-- Properties --//

Public Property Get DisplayName(Optional File) As String
'Returns the display name of the specified file.
    Dim p_Ret   As Long
    Dim p_SHFI  As SHFILEINFO
   
    If (IsMissing(File)) Then
        File = m_sFile
    End If
   
    p_Ret = SHGetFileInfo(CStr(File), 0&, p_SHFI, Len(p_SHFI), SHGFI_DISPLAYNAME)
   
    If (p_Ret <> ERRORAPI) Then
        DisplayName = p_SHFI.szDisplayName
        DisplayName = Left$(DisplayName, lstrlen(DisplayName))
    End If
End Property

Public Property Get ExeType(Optional File) As EFileExeTypes
'Returns the display name of the specified file.
    Dim p_Ret   As Long
    Dim p_SHFI  As SHFILEINFO
   
    If (IsMissing(File)) Then
        File = m_sFile
    End If
   
    p_Ret = SHGetFileInfo(CStr(File), 0&, p_SHFI, Len(p_SHFI), SHGFI_EXETYPE)
   
    If (p_Ret <> ERRORAPI) Then
        If (HiWord(p_Ret) > 0) Then ' NE 0x00004E45 or PE 0x00005045
            ExeType = WindowsApp
        Else
            Select Case LoWord(p_Ret)
                Case 23117 ' MZ 0x00004D5A
                    ExeType = MSDosApp
                Case 17744 ' PE 0x00005045
                    ExeType = Win32Console
            End Select
        End If
    End If
End Property

Public Property Get File() As String
'Returns/sets the complete file path to be used.
    File = m_sFile
End Property

Public Property Let File(Value As String)
    m_sFile = Value
End Property

Public Property Get Handle() As Long
'Returns/sets the icon handle to be used by the IconEx property.
    Handle = m_lHandle
End Property

Public Property Let Handle(Value As Long)
    m_lHandle = Value
End Property

Public Property Get IconType() As EFileIconTypes
'Returns/sets the type of icon to retrieve.
    IconType = m_eIconType
End Property

Public Property Let IconType(Value As EFileIconTypes)
    m_eIconType = Value
End Property

Public Property Get Icon(Optional File, Optional IconType) As IPictureDisp
'Returns the icon of the specified file.
    If (IsMissing(File)) Then
        File = m_sFile
    End If
   
    If (IsMissing(IconType)) Then
        IconType = m_eIconType
    End If
   
    Dim p_Flags As ESHGetFileInfoFlagConstants
    Dim p_hIcon As Long
    Dim p_Ret   As Long
    Dim p_SHFI  As SHFILEINFO
   
    If (m_eIconType = LargeIcon) Then
        p_Flags = SHGFI_ICON Or SHGFI_LARGEICON
    Else
        p_Flags = SHGFI_ICON Or SHGFI_SMALLICON
    End If
   
    If (m_bOverlay) Then
        p_Flags = p_Flags Or SHGFI_LINKOVERLAY
    End If
   
    If (m_bSelected) Then
        p_Flags = p_Flags Or SHGFI_SELECTED
    Else
        p_Flags = p_Flags And Not SHGFI_SELECTED
    End If
   
    If (m_bOpenState) Then
        p_Flags = p_Flags Or SHGFI_OPENICON
    Else
        p_Flags = p_Flags And Not SHGFI_OPENICON
    End If
   
    p_Ret = SHGetFileInfo(CStr(File), 0&, p_SHFI, Len(p_SHFI), p_Flags)
   
    If (p_Ret <> ERRORAPI) Then
        p_hIcon = p_SHFI.hIcon
       
        If (p_hIcon) Then
            Set Icon = IconEx(p_hIcon)
        End If
    End If
End Property

Public Property Get IconEx(Optional hIcon As Long) As IPictureDisp
'Returns the file's icon using the specified icon handle.
    If (hIcon = 0) Then
        hIcon = m_lHandle
       
        If (hIcon = 0) Then
            Exit Property
        End If
    End If
   
    Dim p_Picture   As IPictureDisp
    Dim p_PicDesc   As PictDesc
    Dim p_Guid      As Guid
   
    p_PicDesc.cbSizeofStruct = Len(p_PicDesc)
    p_PicDesc.picType = vbPicTypeIcon
    p_PicDesc.hImage = hIcon
   
    ' IPicture GUID {7BF80980-BF32-101A-8BBB-00AA00300CAB}
    With p_Guid
        .Data1 = &H7BF80980
        .data2 = &HBF32
        .Data3 = &H101A
        .Data4(0) = &H8B
        .Data4(1) = &HBB
        .Data4(2) = &H0
        .Data4(3) = &HAA
        .Data4(4) = &H0
        .Data4(5) = &H30
        .Data4(6) = &HC
        .Data4(7) = &HAB
    End With
    ' From vbAccelerator... (http://www.vbaccelerator.com)
   
    OleCreatePictureIndirect p_PicDesc, p_Guid, True, p_Picture
   
    Set IconEx = p_Picture
End Property

Public Property Get LinkOverlay() As Boolean
'Returns/sets a value to determine if a linkoverlay icon is displayed on the icon.
    LinkOverlay = m_bOverlay
End Property

Public Property Let LinkOverlay(Value As Boolean)
    m_bOverlay = Value
End Property

Public Property Get OpenState() As Boolean
'Returns/sets a value to determine if the icon will be in open state. (Ex. Folders)
    OpenState = m_bOpenState
End Property

Public Property Let OpenState(Value As Boolean)
    m_bOpenState = Value
End Property

Public Property Get Selected() As Boolean
'Returns/sets a value to determine if the icon is in selected state.
    Selected = m_bSelected
End Property

Public Property Let Selected(Value As Boolean)
    m_bSelected = Value
End Property

Public Property Get TypeName(Optional File) As String
'Returns the type name of the specified file.
    Dim p_Ret   As Long
    Dim p_SHFI  As SHFILEINFO
   
    If (IsMissing(File)) Then
        File = m_sFile
    End If
   
    p_Ret = SHGetFileInfo(CStr(File), 0&, p_SHFI, Len(p_SHFI), SHGFI_TYPENAME)
   
    If (p_Ret <> ERRORAPI) Then
        TypeName = p_SHFI.szTypeName
        TypeName = Left$(TypeName, lstrlen(TypeName))
    End If
End Property

' //-- Private properties --//

Private Property Get HiWord(DWord As Long) As Long
    HiWord = (DWord And &HFFFF0000) \ &H10000
End Property

Private Property Get LoWord(DWord As Long) As Long
    If (DWord And &H8000&) Then
        LoWord = DWord Or &HFFFF0000
    Else
        LoWord = DWord And &HFFFF&
    End If
End Property
'######Selesai Disini######
