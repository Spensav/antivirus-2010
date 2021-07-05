Attribute VB_Name = "ModListView"
Option Explicit
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" _
   (ByVal hwnd As Long, _
    ByVal Msg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long) As Long


Public Const LVM_SETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + 54
Public Const LVM_GETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + 55

Public Const LVS_EX_FULLROWSELECT = &H20
Public Const LVS_EX_GRIDLINES = &H1 ' untuk membuat gridlines
Public Const LVS_EX_CHECKBOXES As Long = &H4 ' untuk penambahan checkbox
Public Const LVS_EX_HEADERDRAGDROP = &H10
Public Const LVS_EX_TRACKSELECT = &H8
Public Const LVS_EX_ONECLICKACTIVATE = &H40
Public Const LVS_EX_TWOCLICKACTIVATE = &H80
Public Const LVS_EX_SUBITEMIMAGES = &H2

Public Const LVIF_STATE = &H8
 
Public Const LVM_SETITEMSTATE = (LVM_FIRST + 43)
Public Const LVM_GETITEMSTATE As Long = (LVM_FIRST + 44)
Public Const LVM_GETITEMTEXT As Long = (LVM_FIRST + 45)
Private Const GWL_STYLE        As Long = (-16)
Private Const LVM_GETHEADER    As Long = (LVM_FIRST + 31)
Private Const LVM_ARRANGE      As Long = (LVM_FIRST + 22)
Private Const HDS_BUTTONS      As Long = 2

Public Const LVIS_STATEIMAGEMASK As Long = &HF000

Public Type LVITEM
   mask         As Long
   iItem        As Long
   iSubItem     As Long
   State        As Long
   stateMask    As Long
   pszText      As String
   cchTextMax   As Long
   iImage       As Long
   lParam       As Long
   iIndent      As Long
End Type

Public Const LVM_GETCOLUMN = (LVM_FIRST + 25)
Public Const LVM_GETCOLUMNORDERARRAY = (LVM_FIRST + 59)
Public Const LVCF_TEXT = &H4

Public Type LVCOLUMN
    mask As Long
    fmt As Long
    cx As Long
    pszText  As String
    cchTextMax As Long
    iSubItem As Long
    iImage As Long
    iOrder As Long
End Type

'Public Const LVM_FIRST As Long = &H1000
Public Type LV_ITEM
   mask         As Long
   iItem        As Long
   iSubItem     As Long
   State        As Long
   stateMask    As Long
   pszText      As String
   cchTextMax   As Long
   iImage       As Long
   lParam       As Long
   iIndent      As Long
End Type
Function ListviewCheck(lvStyle As ListView)
    Dim rStyle As Long
    Dim r As Long
    rStyle = SendMessageLong(lvStyle.hwnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0&, 0&)
    rStyle = rStyle Xor LVS_EX_FULLROWSELECT Xor LVS_EX_GRIDLINES Xor LVS_EX_CHECKBOXES
    r = SendMessageLong(lvStyle.hwnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0&, rStyle)
End Function
Function ListviewFlat(lvStyle As ListView)
    Dim rStyle As Long
    Dim r As Long
    rStyle = SendMessageLong(lvStyle.hwnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0&, 0&)
    rStyle = rStyle Xor LVS_EX_FULLROWSELECT Xor LVS_EX_GRIDLINES 'Xor LVS_EX_CHECKBOXES
    r = SendMessageLong(lvStyle.hwnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0&, rStyle)
End Function
Public Sub SetCheckAllItems(bState As Boolean)
   Dim LVV As LV_ITEM
   Dim lvCount As Long
   Dim lvIndex As Long
   Dim lvState As Long
   lvState = IIf(bState, &H2000, &H1000)
   lvCount = frmMain.lvVirus.ListItems.Count - 1
   Do
      With LVV
         .mask = LVIF_STATE
         .State = lvState
         .stateMask = LVIS_STATEIMAGEMASK
      End With
Call SendMessage(frmMain.lvVirus.hwnd, LVM_SETITEMSTATE, lvIndex, LVV)
lvIndex = lvIndex + 1
Loop Until lvIndex > lvCount
End Sub
