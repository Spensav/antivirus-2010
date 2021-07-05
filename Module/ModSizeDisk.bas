Attribute VB_Name = "ModSizeDisk"
Option Explicit

' advanced information about devices

Private Declare Sub CpyMem Lib "KERNEL32" _
Alias "RtlMoveMemory" ( _
    pDst As Any, pSrc As Any, ByVal cb As Long _
)

Private Declare Function lstrcpy Lib "KERNEL32" Alias "lstrcpyA" ( _
    lpString1 As Any, lpString2 As Any _
) As Long

Private Declare Function lstrlen Lib "KERNEL32" Alias "lstrlenA" ( _
    lpString As Any _
) As Long

Private Declare Function CreateFile Lib "kernel32.dll" _
Alias "CreateFileA" ( _
    ByVal lpFilename As String, ByVal dwDesiredAccess As Long, _
    ByVal dwShareMode As Long, lpSecurityAttributes As Any, _
    ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, _
    ByVal hTemplateFile As Long _
) As Long

Private Declare Function DeviceIoControl Lib "KERNEL32" ( _
    ByVal hDevice As Long, ByVal dwIoControlCode As Long, _
    lpInBuffer As Any, ByVal nInBufferSize As Long, _
    lpOutBuffer As Any, ByVal nOutBufferSize As Long, _
    lpBytesReturned As Long, lpOverlapped As Any _
) As Long

Private Declare Function CloseHandle Lib "KERNEL32" ( _
    ByVal hObject As Long _
) As Long

Private Const OPEN_EXISTING                 As Long = 3&
Private Const FILE_SHARE_READ               As Long = &H1&
Private Const FILE_SHARE_WRITE              As Long = &H2&
Private Const GENERIC_READ                  As Long = &H80000000
Private Const IOCTL_STORAGE_QUERY_PROPERTY  As Long = &H2D1400

Private Type STORAGE_PROPERTY_QUERY
    PropertyId                              As STORAGE_PROPERTY_ID
    QueryType                               As STORAGE_QUERY_TYPE
    AdditionalParameters                    As Byte
End Type

Public Type DEVICE_INFORMATION
    Valid                                   As Boolean
    BusType                                 As STORAGE_BUS_TYPE
    Removable                               As Boolean
    VendorID                                As String
    ProductID                               As String
    ProductRevision                         As String
End Type

Private Type STORAGE_DEVICE_DESCRIPTOR
    Version                                 As Long
    Size                                    As Long
    DeviceType                              As Byte
    DeviceTypeModifier                      As Byte
    RemovableMedia                          As Byte
    CommandQueueing                         As Byte
    VendorIdOffset                          As Long
    ProductIdOffset                         As Long
    ProductRevisionOffset                   As Long
    SerialNumberOffset                      As Long
    BusType                                 As Integer
    RawPropertiesLength                     As Long
    RawDeviceProperties                     As Byte
End Type

Public Enum STORAGE_BUS_TYPE
    BusTypeUnknown = 0
    BusTypeScsi
    BusTypeAtapi
    BusTypeAta
    BusType1394
    BusTypeSsa
    BusTypeFibre
    BusTypeUsb
    BusTypeRAID
    BusTypeMaxReserved = &H7F
End Enum

Private Enum STORAGE_PROPERTY_ID
    StorageDeviceProperty = 0
    StorageAdapterProperty
    StorageDeviceIdProperty
End Enum

Private Enum STORAGE_QUERY_TYPE
    PropertyStandardQuery = 0
    PropertyExistsQuery
    PropertyMaskQuery
    PropertyQueryMaxDefined
End Enum

Private Type LARGE_INTEGER
    lowpart As Long
    highpart As Long
End Type
Private Declare Function GetDiskFreeSpaceEx Lib "KERNEL32" Alias "GetDiskFreeSpaceExA" (ByVal lpRootPathName As String, lpFreeBytesAvailableToCaller As LARGE_INTEGER, lpTotalNumberOfBytes As LARGE_INTEGER, lpTotalNumberOfFreeBytes As LARGE_INTEGER) As Long

Public Function GetDiskSpace(sDrive As String) As String
    Dim lResult As Long
    Dim liAvailable As LARGE_INTEGER
    Dim liTotal As LARGE_INTEGER
    Dim liFree As LARGE_INTEGER
    Dim dblAvailable As Double
    Dim dblTotal As Double
    Dim dblFree As Double
    If Right(sDrive, 1) <> "" Then sDrive = sDrive & ""
    'Determine the Available Space, Total Size and Free Space of a drive
    lResult = GetDiskFreeSpaceEx(sDrive, liAvailable, liTotal, liFree)
   
    'Convert the return values from LARGE_INTEGER to doubles
    dblAvailable = CLargeInt(liAvailable.lowpart, liAvailable.highpart)
    dblTotal = CLargeInt(liTotal.lowpart, liTotal.highpart)
    dblFree = CLargeInt(liFree.lowpart, liFree.highpart)
   
    'Display the results
    GetDiskSpace = Format(dblTotal / 1024 ^ 3, "0.00") & " GB"
End Function
Public Function GetDiskSpace1(sDrive As String) As String
    Dim lResult As Long
    Dim liAvailable As LARGE_INTEGER
    Dim liTotal As LARGE_INTEGER
    Dim liFree As LARGE_INTEGER
    Dim dblAvailable As Double
    Dim dblTotal As Double
    Dim dblFree As Double
    If Right(sDrive, 1) <> "" Then sDrive = sDrive & ""
    'Determine the Available Space, Total Size and Free Space of a drive
    lResult = GetDiskFreeSpaceEx(sDrive, liAvailable, liTotal, liFree)
   
    'Convert the return values from LARGE_INTEGER to doubles
    dblAvailable = CLargeInt(liAvailable.lowpart, liAvailable.highpart)
    dblTotal = CLargeInt(liTotal.lowpart, liTotal.highpart)
    dblFree = CLargeInt(liFree.lowpart, liFree.highpart)
   
    'Display the results
    GetDiskSpace1 = "Available Space on " & sDrive & ":  " & dblAvailable & " bytes (" & _
                Format(dblAvailable / 1024 ^ 3, "0.00") & " G) " & vbCr & _
                "Total Space on " & sDrive & ":      " & dblTotal & " bytes (" & _
                Format(dblTotal / 1024 ^ 3, "0.00") & " G) " & vbCr & _
                "Free Space on " & sDrive & ":       " & dblFree & " bytes (" & _
                Format(dblFree / 1024 ^ 3, "0.00") & " G) "
End Function
Public Function GetDiskSpaceFree(sDrive As String) As String
    Dim lResult As Long
    Dim liAvailable As LARGE_INTEGER
    Dim liTotal As LARGE_INTEGER
    Dim liFree As LARGE_INTEGER
    Dim dblAvailable As Double
    Dim dblTotal As Double
    Dim dblFree As Double
    If Right(sDrive, 1) <> "" Then sDrive = sDrive & ""
    'Determine the Available Space, Total Size and Free Space of a drive
    lResult = GetDiskFreeSpaceEx(sDrive, liAvailable, liTotal, liFree)
   
    'Convert the return values from LARGE_INTEGER to doubles
    dblAvailable = CLargeInt(liAvailable.lowpart, liAvailable.highpart)
    dblTotal = CLargeInt(liTotal.lowpart, liTotal.highpart)
    dblFree = CLargeInt(liFree.lowpart, liFree.highpart)
   
    'Display the results
    GetDiskSpaceFree = Format(dblFree / 1024 ^ 3, "0.00") & " GB"
End Function
Private Function CLargeInt(Lo As Long, Hi As Long) As Double
    'This function converts the LARGE_INTEGER data type to a double
    Dim dblLo As Double, dblHi As Double
   
    If Lo < 0 Then
        dblLo = 2 ^ 32 + Lo
    Else
        dblLo = Lo
    End If
   
    If Hi < 0 Then
        dblHi = 2 ^ 32 + Hi
    Else
        dblHi = Hi
    End If
    CLargeInt = dblLo + dblHi * 2 ^ 32
End Function
Public Function GetDevInfo(ByVal strDrive As String) As DEVICE_INFORMATION
    Dim hDrive          As Long
    Dim udtQuery        As STORAGE_PROPERTY_QUERY
    Dim dwOutBytes      As Long
    Dim lngResult       As Long
    Dim btBuffer(9999)  As Byte
    Dim udtOut          As STORAGE_DEVICE_DESCRIPTOR
    
    hDrive = CreateFile("\\.\" & Left$(strDrive, 1) & ":", 0, _
                        FILE_SHARE_READ Or FILE_SHARE_WRITE, _
                        ByVal 0&, OPEN_EXISTING, 0, 0)

    If hDrive = -1 Then Exit Function
    
    With udtQuery
        .PropertyId = StorageDeviceProperty
        .QueryType = PropertyStandardQuery
    End With
    
    lngResult = DeviceIoControl(hDrive, IOCTL_STORAGE_QUERY_PROPERTY, _
                                udtQuery, LenB(udtQuery), _
                                btBuffer(0), UBound(btBuffer) + 1, _
                                dwOutBytes, ByVal 0&)
        
    If lngResult Then
        CpyMem udtOut, btBuffer(0), Len(udtOut)
        
        With GetDevInfo
            .Valid = True
            .BusType = udtOut.BusType
            .Removable = CBool(udtOut.RemovableMedia)
            
            If udtOut.ProductIdOffset > 0 Then _
                .ProductID = StringCopy(VarPtr(btBuffer(udtOut.ProductIdOffset)))
            If udtOut.ProductRevisionOffset > 0 Then _
                .ProductRevision = StringCopy(VarPtr(btBuffer(udtOut.ProductRevisionOffset)))
            If udtOut.VendorIdOffset > 0 Then
                .VendorID = StringCopy(VarPtr(btBuffer(udtOut.VendorIdOffset)))
            End If
        End With
    Else
        GetDevInfo.Valid = False
    End If
    
    CloseHandle hDrive
End Function

Private Function StringCopy(ByVal pBuffer As Long) As String
    Dim tmp As String
    
    tmp = Space(lstrlen(ByVal pBuffer))
    lstrcpy ByVal tmp, ByVal pBuffer
    StringCopy = Trim$(tmp)
End Function




