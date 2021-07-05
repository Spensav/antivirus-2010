Attribute VB_Name = "ModSysTray"
'******************************************************************************
'Systray Module
'
'Option Explicit
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copyright ©1996-2004 VBnet, Randy Birch, All Rights Reserved.
' Some pages may also contain other copyrights by the author.

Public defwindowproc As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Public Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
Public Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

'flag preventing re-creating the timer
Private tmrRunning As Boolean

'Get/SetWindowLong messages
Private Const GWL_WNDPROC As Long = (-4)
Private Const GWL_HWNDPARENT As Long = (-8)
Private Const GWL_ID As Long = (-12)
Private Const GWL_STYLE As Long = (-16)
Private Const GWL_EXSTYLE As Long = (-20)
Private Const GWL_USERDATA As Long = (-21)

'general windows messages
Private Const WM_USER As Long = &H400
Private Const WM_NOTIFY As Long = &H4E
Private Const WM_COMMAND As Long = &H111
Public Const WM_CLOSE As Long = &H10
Private Const WM_TIMER = &H113

'mouse constants for the callback
Private Const WM_LBUTTONDOWN As Long = &H201
Private Const WM_LBUTTONUP As Long = &H202
Private Const WM_LBUTTONDBLCLK As Long = &H203
Private Const WM_MBUTTONDOWN As Long = &H207
Private Const WM_MBUTTONUP As Long = &H208
Private Const WM_MBUTTONDBLCLK As Long = &H209
Private Const WM_RBUTTONDOWN As Long = &H204
Private Const WM_RBUTTONUP As Long = &H205
Private Const WM_RBUTTONDBLCLK As Long = &H206

'private message the shell_notify api will pass
'to WindowProc when our systray icon is acted upon
Private Const WM_MYHOOK As Long = WM_USER + 1

'ID constant representing this
'application in the systray
Private Const APP_SYSTRAY_ID = 999

'ID constant representing this
'application for SetTimer
Public Const APP_TIMER_EVENT_ID As Long = 998

'const holding number of milliseconds to timeout
'7000=7 seconds
Public Const APP_TIMER_MILLISECONDS As Long = 7000

'balloon tip notification messages
Private Const NIN_BALLOONSHOW = (WM_USER + 2)
Private Const NIN_BALLOONHIDE = (WM_USER + 3)
Private Const NIN_BALLOONTIMEOUT = (WM_USER + 4)
Private Const NIN_BALLOONUSERCLICK = (WM_USER + 5)

'shell version / NOTIFIYICONDATA struct size constants
Private Const NOTIFYICONDATA_V1_SIZE As Long = 88  'pre-5.0 structure size
Private Const NOTIFYICONDATA_V2_SIZE As Long = 488 'pre-6.0 structure size
Private Const NOTIFYICONDATA_V3_SIZE As Long = 504 '6.0+ structure size
Private NOTIFYICONDATA_SIZE As Long

Private Const NOTIFYICON_VERSION = &H3

'shell_notify flags
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const NIF_STATE = &H8
Private Const NIF_INFO = &H10
'shell_notify messages
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIM_SETFOCUS = &H3
Private Const NIM_SETVERSION = &H4
Private Const NIM_VERSION = &H5
'shell_notify styles
Private Const NIS_HIDDEN = &H1
Private Const NIS_SHAREDICON = &H2

Public Type NOTIFYICONDATA
  cbSize As Long
  hwnd As Long
  uID As Long
  uFlags As Long
  uCallbackMessage As Long
  hIcon As Long
  szTip As String * 128
  dwState As Long
  dwStateMask As Long
  szInfo As String * 256
  uTimeoutAndVersion As Long
  szInfoTitle As String * 64
  dwInfoFlags As Long
End Type

'shell_notify icon flags
Private Const NIIF_NONE = &H0
Private Const NIIF_INFO = &H1
Private Const NIIF_WARNING = &H2
Private Const NIIF_ERROR = &H3
Private Const NIIF_GUID = &H5
Private Const NIIF_ICON_MASK = &HF
Private Const NIIF_NOSOUND = &H10

Public Sub ShellTrayIconAdd(hwnd As Long, _
                            hIcon As StdPicture, _
                            sToolTip As String)
   
   Dim nid As NOTIFYICONDATA
   
    NOTIFYICONDATA_SIZE = NOTIFYICONDATA_V3_SIZE
     
   With nid
      .cbSize = NOTIFYICONDATA_SIZE
      .hwnd = hwnd
      .uID = APP_SYSTRAY_ID
      .uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP Or NIF_INFO
      .dwState = NIS_SHAREDICON
      .hIcon = hIcon
      .szTip = sToolTip & vbNullChar
      .uTimeoutAndVersion = NOTIFYICON_VERSION
      .uCallbackMessage = WM_MYHOOK
   End With
   
   If Shell_NotifyIcon(NIM_ADD, nid) = 1 Then
   
      Call Shell_NotifyIcon(NIM_SETVERSION, nid)
      Call SubClass(hwnd)

   End If
       
End Sub


Public Sub ShellTrayIconRemove(hwnd As Long)

Dim nid As NOTIFYICONDATA
   NOTIFYICONDATA_SIZE = NOTIFYICONDATA_V3_SIZE
   With nid
      .cbSize = NOTIFYICONDATA_SIZE
      .hwnd = hwnd
      .uID = APP_SYSTRAY_ID
   End With
   Call Shell_NotifyIcon(NIM_DELETE, nid)

End Sub


Private Sub ShellTrayBalloonTipClose(hwnd As Long)

Dim nid As NOTIFYICONDATA
   
NOTIFYICONDATA_SIZE = NOTIFYICONDATA_V3_SIZE
With nid
      .cbSize = NOTIFYICONDATA_SIZE
      .hwnd = hwnd
      .uID = APP_SYSTRAY_ID
      .uFlags = NIF_TIP Or NIF_INFO
      .szTip = vbNullChar
      .uTimeoutAndVersion = NOTIFYICON_VERSION
   End With
   Call Shell_NotifyIcon(NIM_MODIFY, nid)
   
End Sub


Public Sub ShellTrayBalloonTipShow(hwnd As Long, nIconIndex As Long, sTitle As String, sMessage As String)

Dim nid As NOTIFYICONDATA
   
NOTIFYICONDATA_SIZE = NOTIFYICONDATA_V3_SIZE
With nid
      .cbSize = NOTIFYICONDATA_SIZE
      .hwnd = hwnd
      .uID = APP_SYSTRAY_ID
      .uFlags = NIF_INFO
      .dwInfoFlags = nIconIndex
      .szInfoTitle = sTitle & vbNullChar
      .szInfo = sMessage & vbNullChar
   End With
   Call Shell_NotifyIcon(NIM_MODIFY, nid)

End Sub


Private Sub SubClass(hwnd As Long)

   On Error Resume Next
   defwindowproc = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf WindowProc)
   
End Sub


Public Sub UnSubClass(hwnd As Long)

   If defwindowproc <> 0 Then
      SetWindowLong hwnd, GWL_WNDPROC, defwindowproc
      defwindowproc = 0
   End If
   
End Sub

Public Function WindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
On Error Resume Next
   Select Case hwnd
      Case frmMain.hwnd
         Select Case uMsg
            Case WM_MYHOOK
               Select Case lParam
                  Case WM_RBUTTONUP
                    Call SetForegroundWindow(frmMain.hwnd)
                    frmOther.PopupMenu frmOther.Menu
                  Case WM_LBUTTONDBLCLK
                    Call SetForegroundWindow(frmMain.hwnd)
                    frmMain.Show
               End Select
            Case Else
               WindowProc = CallWindowProc(defwindowproc, hwnd, uMsg, wParam, lParam)
               Exit Function
         End Select
      Case Else
          WindowProc = CallWindowProc(defwindowproc, hwnd, uMsg, wParam, lParam)
   End Select
End Function


Public Sub LoadTray(Index As Integer)
Dim sToolTipku As String
Dim sTitle As String
Dim sMessage As String
Select Case Index
Case 1
sToolTipku = "Spensav AntiVirus"
sTitle = "Spensav AntiVirus"
sMessage = "Spensav AntiVirus Tahun 2013" & vbCrLf & "Create By.Muh.Isfahani Ghiyath.YM"
Case 2
sToolTipku = "Ade Shinichi Simple Protector"
sTitle = "Ade Shinichi Simple Protector"
sMessage = "Ade Shinichi Simple Protector tidak aktif" & vbCrLf & "Komputer anda tidak aman dari beberapa Virus saja."
End Select
Call ShellTrayIconAdd(frmMain.hwnd, frmMain.Icon, sToolTipku)
Call ShellTrayBalloonTipShow(frmMain.hwnd, 1, sTitle, sMessage)
End Sub




