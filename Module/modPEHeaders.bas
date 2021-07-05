Attribute VB_Name = "modPEPilus"
'########################################################
'####           SEPENSAP PE HEADER                   ####
'########################################################

Public Type IMAGE_DOS_HEADER
    e_magic As Integer
    e_cblp As Integer
    e_cp As Integer
    e_crlc As Integer
    e_cparhdr As Integer
    e_minalloc As Integer
    e_maxalloc As Integer
    e_ss As Integer
    e_sp As Integer
    e_csum As Integer
    e_ip As Integer
    e_cs As Integer
    e_lfarlc As Integer
    e_ovno As Integer
    e_res(1 To 4) As Integer
    e_oemid As Integer
    e_oeminfo As Integer
    e_res2(1 To 10)    As Integer
    e_lfanew As Long
End Type

Public Type IMAGE_SECTION_HEADER
    nameSec As String * 6
    PhisicalAddress As Integer
    
    VirtualSize As Long
    VirtualAddress As Long
    SizeOfRawData As Long
    PointerToRawData As Long
    PointerToRelocations As Long
    PointerToLinenumbers As Long
    NumberOfRelocations As Integer
    NumberOfLinenumbers As Integer
    Characteristics As Long
   
End Type

Public Type IMAGE_DATA_DIRECTORY
    VirtualAddress As Long
    Size As Long
End Type

Public Type IMAGE_OPTIONAL_HEADER
    Magic As Integer
    MajorLinkerVersion As Byte
    MinorLinkerVersion As Byte
    SizeOfCode As Long
    SizeOfInitializedData As Long
    SizeOfUninitializedData As Long
    AddressOfEntryPoint As Long
    BaseOfCode As Long
    BaseOfData As Long
    ImageBase As Long
    SectionAlignment As Long
    FileAlignment As Long
    MajorOperatingSystemVersion As Integer
    MinorOperatingSystemVersion As Integer
    MajorImageVersion As Integer
    MinorImageVersion As Integer
    MajorSubsystemVersion As Integer
    MinorSubsystemVersion As Integer
    Win32VersionValue As Long
    SizeOfImage As Long
    SizeOfHeaders As Long
    CheckSum As Long
    Subsystem As Integer
    DllCharacteristics As Integer
    SizeOfStackReserve As Long
    SizeOfStackCommit As Long
    SizeOfHeapReserve As Long
    SizeOfHeapCommit As Long
    LoaderFlags As Long
    NumberOfRvaAndSizes As Long
    DataDirectory(0 To 15) As IMAGE_DATA_DIRECTORY
End Type

Public Type IMAGE_FILE_HEADER
    Machine As Integer
    NumberOfSections As Integer
    TimeDateStamp As Long
    PointerToSymbolTable As Long
    NumberOfSymbols As Long
    SizeOfOptionalHeader As Integer
    Characteristics As Integer
End Type

Public Type IMAGE_NT_HEADERS
    Signature As Long
    FileHeader As IMAGE_FILE_HEADER
    OptionalHeader As IMAGE_OPTIONAL_HEADER
End Type

Public Type IMAGE_EXPORT_DIRECTORY
    Characteristics As Long
    TimeDateStamp As Long
    MajorVersion As Integer
    MinorVersion As Integer
    Name As Long
    Base As Long
    NumberOfFunctions As Long
    NumberOfNames As Long
    AddressOfFunctions As Long
    AddressOfNames As Long
    AddressOfNameOrdinals As Long
End Type

Public Type IMAGE_IMPORT_DESCRIPTOR
    OriginalFirstThunk As Long
    TimeDateStamp As Long
    ForwarderChain As Long
    Name As Long
    FirstThunk As Long
End Type

Public Type IMAGE_IMPORT_BY_NAME
    Hint As Integer
    Name As String * 255
End Type

Public Const IMAGE_SIZEOF_SECTION_HEADER = 40
Public Const IMAGE_DOS_SIGNATURE = &H5A4D
Public Const IMAGE_NT_SIGNATURE = &H4550
Public Const IMAGE_ORDINAL_FLAG = &H80000000

Public Enum SECTION_CHARACTERISTICS
    IMAGE_SCN_LNK_NRELOC_OVFL = &H1000000   'Section contains extended relocations.
    IMAGE_SCN_MEM_DISCARDABLE = &H2000000   'Section can be discarded.
    IMAGE_SCN_MEM_NOT_CACHED = &H4000000    'Section is not cachable.
    IMAGE_SCN_MEM_NOT_PAGED = &H8000000     'Section is not pageable.
    IMAGE_SCN_MEM_SHARED = &H10000000       'Section is shareable.
    IMAGE_SCN_MEM_EXECUTE = &H20000000      'Section is executable.
    IMAGE_SCN_MEM_READ = &H40000000         'Section is readable.
    IMAGE_SCN_MEM_WRITE = &H80000000        'Section is writeable.
End Enum

Public Enum IMAGE_DIRECTORY
    IMAGE_DIRECTORY_ENTRY_EXPORT = 0           ' Export Directory
    IMAGE_DIRECTORY_ENTRY_IMPORT = 1           ' Import Directory
    IMAGE_DIRECTORY_ENTRY_RESOURCE = 2         ' Resource Directory
    IMAGE_DIRECTORY_ENTRY_EXCEPTION = 3        ' Exception Directory
    IMAGE_DIRECTORY_ENTRY_SECURITY = 4         ' Security Directory
    IMAGE_DIRECTORY_ENTRY_BASERELOC = 5        ' Base Relocation Table
    IMAGE_DIRECTORY_ENTRY_DEBUG = 6            ' Debug Directory
    IMAGE_DIRECTORY_ENTRY_ARCHITECTURE = 7     ' Architecture Specific Data
    IMAGE_DIRECTORY_ENTRY_GLOBALPTR = 8        ' RVA of GP
    IMAGE_DIRECTORY_ENTRY_TLS = 9              ' TLS Directory
    IMAGE_DIRECTORY_ENTRY_LOAD_CONFIG = 10     ' Load Configuration Directory
    IMAGE_DIRECTORY_ENTRY_BOUND_IMPORT = 11    ' Bound Import Directory in headers
    IMAGE_DIRECTORY_ENTRY_IAT = 12             ' Import Address Table
    IMAGE_DIRECTORY_ENTRY_DELAY_IMPORT = 13    ' Delay Load Import Descriptors
    IMAGE_DIRECTORY_ENTRY_COM_DESCRIPTOR = 14  ' COM Runtime descriptor
End Enum
