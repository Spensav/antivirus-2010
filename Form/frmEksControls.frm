VERSION 5.00
Begin VB.Form frmEksControls 
   BorderStyle     =   0  'None
   Caption         =   "frmEksControls"
   ClientHeight    =   3105
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5505
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   5505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   2400
      Left            =   1080
      TabIndex        =   19
      Top             =   3960
      Width           =   4455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1080
      TabIndex        =   18
      Text            =   "Text1"
      Top             =   6480
      Width           =   4455
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1080
      TabIndex        =   17
      Top             =   6840
      Width           =   4455
   End
   Begin VB.ListBox List2 
      Height          =   2790
      ItemData        =   "frmEksControls.frx":0000
      Left            =   5640
      List            =   "frmEksControls.frx":0002
      TabIndex        =   16
      Top             =   4320
      Width           =   3495
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   5640
      TabIndex        =   15
      Top             =   3960
      Width           =   3495
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6600
      Top             =   1440
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C00000&
      BorderStyle     =   0  'None
      Height          =   3135
      Left            =   0
      ScaleHeight     =   209
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   369
      TabIndex        =   0
      Top             =   0
      Width           =   5535
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   3135
         Left            =   4560
         ScaleHeight     =   3135
         ScaleWidth      =   975
         TabIndex        =   2
         Top             =   0
         Width           =   975
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   615
            Left            =   600
            MouseIcon       =   "frmEksControls.frx":0004
            MousePointer    =   99  'Custom
            TabIndex        =   3
            Top             =   120
            Width           =   4095
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Lihat List App..."
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   1
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   1080
         Top             =   2400
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama File"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0C0&
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Ukuran[s]"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0C0&
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Lokasi File"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0C0&
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0C0&
         Height          =   255
         Left            =   1320
         TabIndex        =   11
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0C0&
         Height          =   255
         Left            =   1440
         TabIndex        =   10
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0C0&
         Height          =   255
         Left            =   1320
         TabIndex        =   9
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "aaa"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0C0&
         Height          =   255
         Left            =   1440
         TabIndex        =   8
         Top             =   1200
         Width           =   2895
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "aaa"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0C0&
         Height          =   255
         Left            =   1440
         TabIndex        =   7
         Top             =   1560
         Width           =   2895
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "aaa"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0C0&
         Height          =   375
         Left            =   1440
         TabIndex        =   6
         Top             =   1920
         Width           =   2895
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Baru Saja Spensav Menemukan Program Yang Akan Di Jalankan, Mungkin Bersifat Virus."
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0C0&
         Height          =   615
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   4095
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "EKSEKUSI PROGRAM MALICIOUS !!"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0C0&
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   4095
      End
   End
End
Attribute VB_Name = "frmEksControls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function CreateRemoteThread Lib "kernel32" (ByVal hProcess As Long, lpThreadAttributes As Any, ByVal dwStackSize As Long, lpStartAddress As Long, lpParameter As Any, ByVal dwCreationFlags As Long, lpThreadId As Long) As Long
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Dim LetakAwal As String
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Dim Naik As Boolean

Private Sub Command1_Click()
Naik = False
frmEks.Show
Timer1.Enabled = True
End Sub

Private Sub Form_Load()
GradientPic frmRTPControls.Picture2, &H404040, &H0&, gmHorizontal: GradientPic frmUSB.Picture4, &H404040, &H0&, gmHorizontal: GradientPic frmEksControls.Picture2, &H404040, &H0&, gmHorizontal
    GradientPic Picture1, &H80&, &HC0&, gmVertical
    Top = ((GetSystemMetrics(17) + GetSystemMetrics(4)) * Screen.TwipsPerPixelY)
    Left = (GetSystemMetrics(16) * Screen.TwipsPerPixelX) - Width
    Naik = True
GradientPic Picture1, &H800000, &HFF0000, gmVertical
End Sub

Private Sub Label12_Click()
Naik = False
Timer1.Enabled = True
End Sub

Private Sub List1_Click()
List2.AddItem List1.List(List1.ListCount)
End Sub
Private Function RemoteExitProcess(lProcessID As Long) As Boolean
    Dim lProcess As Long
    Dim lRemThread As Long
    Dim lExitProcess As Long
    
    On Error GoTo errHandle
    
    lProcess = OpenProcess((&HF0000 Or &H100000 Or &HFFF), False, lProcessID) 'PROCESS_ALL_ACCESS
        lExitProcess = GetProcAddress(GetModuleHandleA("kernel32"), "ExitProcess")
        lRemThread = CreateRemoteThread(lProcess, ByVal 0, 0, ByVal lExitProcess, 0, 0, 0)
    CloseHandle lProcess
    
    CloseHandle lRemThread
    RemoteExitProcess = True
    
    Exit Function
errHandle:
    RemoteExitProcess = False
End Function

Private Sub Timer1_Timer()
Const s = 80 'kecepatan gerak / slide
    Dim V As Single
    V = (GetSystemMetrics(17) + GetSystemMetrics(4)) * Screen.TwipsPerPixelY
    
    If Naik = True Then
        If Top - s <= V - Height Then
            Top = Top - (Top - (V - Height))
            Timer1.Enabled = False
        Else
            Top = Top - s
        End If
        
    Else
        Top = Top + s
        If Top >= V Then Unload Me
    End If
End Sub

Private Sub Timer2_Timer()
      List1.Clear
      Select Case getVersion()

      Case 1 'Windows 95/98

         Dim f As Long, sname As String
         Dim hSnap As Long, proc As PROCESSENTRY32
         hSnap = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
         If hSnap = hNull Then Exit Sub
         proc.dwSize = Len(proc)
         ' Iterate through the processes
         f = Process32First(hSnap, proc)
         Do While f
           sname = StrZToStr(proc.szExeFile)
           List1.AddItem sname
           f = Process32Next(hSnap, proc)
         Loop

      Case 2 'Windows NT

         Dim cb As Long
         Dim cbNeeded As Long
         Dim NumElements As Long
         Dim ProcessIDs() As Long
         Dim cbNeeded2 As Long
         Dim NumElements2 As Long
         Dim Modules(1 To 200) As Long
         Dim lRet As Long
         Dim ModuleName As String
         Dim nSize As Long
         Dim hProcess As Long
         Dim i As Long
         'Get the array containing the process id's for each process object
         cb = 8
         cbNeeded = 96
         Do While cb <= cbNeeded
            cb = cb * 2
            ReDim ProcessIDs(cb / 4) As Long
            lRet = EnumProcesses(ProcessIDs(1), cb, cbNeeded)
         Loop
         NumElements = cbNeeded / 4

         For i = 1 To NumElements
            'Get a handle to the Process
            hProcess = OpenProcess(PROCESS_QUERY_INFORMATION _
               Or PROCESS_VM_READ, 0, ProcessIDs(i))
            'Got a Process handle
            If hProcess <> 0 Then
                'Get an array of the module handles for the specified
                'process
                lRet = EnumProcessModules(hProcess, Modules(1), 200, _
                                             cbNeeded2)
                'If the Module Array is retrieved, Get the ModuleFileName
                If lRet <> 0 Then
                   ModuleName = Space(MAX_PATH)
                   nSize = 500
                   lRet = GetModuleFileNameExA(hProcess, Modules(1), _
                                   ModuleName, nSize)
                   List1.AddItem Left(ModuleName, lRet)


                End If
            End If
          'Close the handle to the process
         lRet = CloseHandle(hProcess)
         Next

      End Select
Text1.Text = List1.ListCount
LetakAwal = List1.List(Text1.Text - 1) 'karena kan mulai dari 0 jadi kurangi 1
Text2.Text = LetakAwal

    If Text3.Text = "" Then
Text3.Text = Text2.Text

    ElseIf Text3.Text <> Text2.Text Then
'MsgBox "Ditemukan process baru di komputer anda bernama :" & Text2.Text
'RemoteExitProcess (Text2.Text)
frmEksControls.Show
Dim LV As ListItem
Set LV = frmEks.lvProg.ListItems.Add(, , GetFileName(Text2.Text))
LV.SubItems(1) = FileLen(Text2.Text)
LV.SubItems(2) = Text2.Text

Label7.Caption = GetFileName(Text2.Text)
Label8.Caption = FileLen(Text2.Text)
Label9.Caption = Text2.Text
'INFORMASI : PATHNYA YAITU TEXT2.TEXT Lhoo
Text3.Text = Text2.Text
List2.AddItem LetakAwal
    End If
End Sub
