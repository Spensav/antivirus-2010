VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmRTP 
   BorderStyle     =   0  'None
   ClientHeight    =   4470
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7935
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   7935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tmrControl 
      Interval        =   1000
      Left            =   7200
      Top             =   4800
   End
   Begin VB.Timer TimerRTP 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6720
      Top             =   4800
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Abaikan"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      TabIndex        =   5
      Top             =   3840
      Width           =   1575
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Hapus Cek"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   3840
      Width           =   1575
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Hapus Semua"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   3840
      Width           =   1575
   End
   Begin VB.PictureBox Picture20 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   7
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   8175
      TabIndex        =   0
      Top             =   0
      Width           =   8175
      Begin VB.PictureBox Picture24 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   7200
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   7
         Top             =   0
         Width           =   495
         Begin VB.Label Label81 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C0C0&
            Height          =   255
            Left            =   -120
            MouseIcon       =   "frmSetting.frx":0000
            MousePointer    =   99  'Custom
            TabIndex        =   8
            Top             =   0
            Width           =   735
         End
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Virus Terdeteksi"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   5055
      End
   End
   Begin ComctlLib.ListView lvVirus 
      Height          =   2655
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   4683
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      Icons           =   "ImgListView"
      SmallIcons      =   "ImgListView"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Nama Virus"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Tipe"
         Object.Width           =   1482
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Lokasi Virus"
         Object.Width           =   7832
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H000000C0&
      BorderStyle     =   0  'None
      Height          =   3735
      Left            =   0
      ScaleHeight     =   3735
      ScaleWidth      =   8055
      TabIndex        =   9
      Top             =   960
      Width           =   8055
   End
   Begin ComctlLib.ImageList ImgListView 
      Left            =   8520
      Top             =   2880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   5
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSetting.frx":0152
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSetting.frx":032C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSetting.frx":067E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSetting.frx":0998
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSetting.frx":0CEA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   5040
      Width           =   7455
   End
End
Attribute VB_Name = "frmRTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'########## Untuk ListView ##########'
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const LVM_FIRST As Long = &H1000
Private Const LVM_SETEXTENDEDLISTVIEWSTYLE As Long = (LVM_FIRST + 54)
Private Const LVM_GETEXTENDEDLISTVIEWSTYLE As Long = (LVM_FIRST + 55)
Private Const LVIF_STATE = &H8
Private Const LVS_EX_CHECKBOXES As Long = &H4
Private Const LVM_SETITEMSTATE = (LVM_FIRST + 43)
Private Const LVM_GETITEMSTATE As Long = (LVM_FIRST + 44)
Private Const LVM_GETITEMTEXT As Long = (LVM_FIRST + 45)
Private Const LVIS_STATEIMAGEMASK As Long = &HF000
Private Type LV_ITEM
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
Dim L1(100) As String
Dim L2(100) As String
Dim tL1 As Long, tL2 As Long
Dim iewindow As InternetExplorer
Private currentwindows As ShellWindows

Private Sub Command1_Click()
lvVirus.ListItems.Clear
Me.Hide
TimerRTP.Enabled = True
tmrControl.Enabled = True
End Sub

Private Sub Command6_Click()
   Dim i As Integer
   Dim r As Long
   Dim b As String
   Dim SubItemText As String
   Dim LV As LV_ITEM
   For i = 0 To lvVirus.ListItems.Count
         r = SendMessage(lvVirus.hwnd, LVM_GETITEMSTATE, i, ByVal LVIS_STATEIMAGEMASK)
      If r And &H2000& Then
         With LV
           .iSubItem = 2
            .cchTextMax = MAX_PATH
            .pszText = Space$(MAX_PATH)
         End With
         Call SendMessage(lvVirus.hwnd, LVM_GETITEMTEXT, i, LV)
         Dim iCount As Integer
         For iCount = 1 To lvVirus.ListItems.Count
         With lvVirus.ListItems.Item(iCount)
         .SmallIcon = frmMain.ImgListView.ListImages(5).Index
         End With
         Next
         Bunuh (Left$(LV.pszText, InStr(LV.pszText, Chr$(0)) - 1))
         NormalkanAtribut (Left$(LV.pszText, InStr(LV.pszText, Chr$(0)) - 1))
         DeleteFile (Left$(LV.pszText, InStr(LV.pszText, Chr$(0)) - 1))
      End If
   Next
End Sub

Private Sub Command7_Click()
On Error Resume Next
Dim i As Integer
For i = 1 To lvVirus.ListItems.Count
DoEvents
Bunuh (Left$(LV.pszText, InStr(LV.pszText, Chr$(0)) - 1))
NormalkanAtribut (Left$(LV.pszText, InStr(LV.pszText, Chr$(0)) - 1))
DeleteFile lvVirus.ListItems(i).SubItems(2)
With lvVirus.ListItems.Item(i)
.SmallIcon = frmMain.ImgListView.ListImages(5).Index
End With
Next
End Sub

Private Sub Form_Load()
Set currentwindows = New ShellWindows
GradientPic Picture1, &H80&, &HC0&, gmVertical
FormOnTop Me, True
TimerRTP.Enabled = True
BacaDatabase App.path & "\Signature.dat"
End Sub
Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then 'left click
ReleaseCapture
SendMessage Me.hwnd, &HA1, 2, 0
End If
End Sub

Private Sub Label81_Click()
If lvVirus.ListItems.Count <> 0 Then
FormOnTop Me, False
If MsgBox("Yakin Untuk Membunuh Proses terpilih ini ??", vbExclamation + vbYesNo, "Persetujuan") = vbYes Then
lvVirus.ListItems.Clear
Me.Hide
TimerRTP.Enabled = True
tmrControl.Enabled = True
End If
End If
End Sub
Private Sub Picture20_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then 'left click
ReleaseCapture
SendMessage Me.hwnd, &HA1, 2, 0
End If
End Sub

Private Sub TimerRTP_Timer()
Dim i As Long
On Error GoTo TheEnd
If currentwindows.Count > 0 Then
Erase L2
tL2 = 0
    For Each iewindow In currentwindows
        DoEvents
        If iewindow.Busy Then GoTo busysignal
    Dim currentlocation As String
        currentlocation = iewindow.LocationURL
        If Mid$(currentlocation, 1, 7) = "file://" Then
                 currentlocation = Replace(currentlocation, "file:///", "")
                 currentlocation = Replace(currentlocation, "%20", " ")
                 currentlocation = Replace(currentlocation, "/", "\")
                 currentlocation = Replace(currentlocation, "%5B", "[")
                 currentlocation = Replace(currentlocation, "%5D", "]")
         L2(tL2) = currentlocation
         tL2 = inc(tL2)
         Dim K As Long
         For K = 0 To dec(tL1)
            If currentlocation = L1(K) Then GoTo busysignal
         Next K
        'MsgBox currentlocation, vbSystemModal, "ojanblank"
        ScanRTP currentlocation, True
        ScanRTP currentlocation, False
End If
busysignal:
    Next
    Erase L1
    tL1 = 0
    For K = 0 To dec(tL2)
        L1(K) = L2(K)
        tL1 = inc(tL1)
    Next K
    End If
TheEnd:
End Sub
Private Function inc(ByVal a As Long) As Long
a = a + 1
End Function
Private Function dec(ByVal a As Long) As Long
a = a - 1
End Function

Private Sub tmrControl_Timer()
If lvVirus.ListItems.Count <> 0 Then
TimerRTP.Enabled = False
Label8.Caption = "Ada " & lvVirus.ListItems.Count & " Virus Terdeteksi"
frmRTPControls.Show
'Dim I As Integer
'For I = 1 To lvVirus.ListItems.Count
'frmRTPControls.Label7.Caption = lvVirus.ListItems(I)
'frmRTPControls.Label8.Caption = FileLen(lvVirus.ListItems(I))
'frmRTPControls.Label9.Caption = lvVirus.ListItems(I).SubItems(2)
'Next
End If
End Sub
