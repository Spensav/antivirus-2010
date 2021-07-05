VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmUSB 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5520
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer3 
      Interval        =   1
      Left            =   4560
      Top             =   3600
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   3135
      Left            =   4560
      ScaleHeight     =   3135
      ScaleWidth      =   1215
      TabIndex        =   24
      Top             =   0
      Width           =   1215
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
         MouseIcon       =   "frmUSB.frx":0000
         MousePointer    =   99  'Custom
         TabIndex        =   25
         Top             =   120
         Width           =   4095
      End
   End
   Begin VB.PictureBox Frame2 
      BackColor       =   &H00000080&
      BorderStyle     =   0  'None
      Height          =   3135
      Left            =   4560
      ScaleHeight     =   3135
      ScaleWidth      =   4575
      TabIndex        =   15
      Top             =   0
      Width           =   4575
      Begin VB.TextBox Text11RTP 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox txtPindaiUSB 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   600
         Width           =   4095
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Lihat Hasil >>"
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
         Left            =   2520
         TabIndex        =   16
         Top             =   2280
         Width           =   1815
      End
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   240
         Top             =   2280
      End
      Begin ComctlLib.ProgressBar ProgressBar1RTP 
         Height          =   255
         Left            =   1680
         TabIndex        =   19
         Top             =   1320
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   0
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Proses Pemindai :"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0FF&
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Progress Bar :"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0FF&
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Virus :"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0FF&
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0FF&
         Height          =   255
         Left            =   840
         TabIndex        =   20
         Top             =   2160
         Width           =   2055
      End
   End
   Begin VB.PictureBox Frame1 
      BackColor       =   &H00000080&
      BorderStyle     =   0  'None
      Height          =   3135
      Left            =   0
      ScaleHeight     =   3135
      ScaleWidth      =   4575
      TabIndex        =   6
      Top             =   0
      Width           =   4575
      Begin VB.CommandButton Command1 
         Caption         =   "Pindai Flash-Drive"
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
         Left            =   480
         TabIndex        =   8
         Top             =   2400
         Width           =   1815
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Batalkan"
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
         Left            =   2400
         TabIndex        =   7
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "FLASHDISK TERDETEKSI !!"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0FF&
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   240
         Width           =   4095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Ada Flash-Drive Yang Baru Masuk di Sistem Operasi, Silahkan Pilih Tindakan Di Bawah Ini :"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   600
         Width           =   4095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama FlashDrive"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0FF&
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Ukuran FlashDrive"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0FF&
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   1680
         Width           =   2175
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
         ForeColor       =   &H00C0C0FF&
         Height          =   255
         Left            =   2160
         TabIndex        =   10
         Top             =   1320
         Width           =   2175
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
         ForeColor       =   &H00C0C0FF&
         Height          =   255
         Left            =   2160
         TabIndex        =   9
         Top             =   1680
         Width           =   2175
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   3135
      Left            =   600
      ScaleHeight     =   3135
      ScaleWidth      =   5535
      TabIndex        =   5
      Top             =   5040
      Width           =   5535
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   7320
      Top             =   2280
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   2775
      Left            =   4320
      ScaleHeight     =   2775
      ScaleWidth      =   7215
      TabIndex        =   0
      Top             =   5520
      Width           =   7215
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   3135
         Left            =   1440
         ScaleHeight     =   3135
         ScaleWidth      =   4575
         TabIndex        =   4
         Top             =   6240
         Width           =   4575
      End
      Begin VB.PictureBox Picture24 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   5040
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   1
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
            MouseIcon       =   "frmUSB.frx":0152
            MousePointer    =   99  'Custom
            TabIndex        =   2
            Top             =   0
            Width           =   735
         End
      End
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Label10"
      Height          =   255
      Left            =   6240
      TabIndex        =   3
      Top             =   4920
      Width           =   3015
   End
End
Attribute VB_Name = "frmUSB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Dim Naik As Boolean
Private Sub Command1_Click()
ProgressBar1RTP.value = 0: txtPindaiUSB.Text = "": Text11RTP.Text = "": Label9.Caption = 0
Geser Frame2
frmMain.clear_log
frmMain.tmrInformasi.Enabled = True: frmMain.Picture10.Visible = True: frmMain.Picture11.Visible = False: frmMain.Picture9.Visible = False: frmMain.Picture8.Visible = False
StopScan = False
frmMain.Command9.Enabled = True 'Untuk Hentikan Proses Scan
AnalisaFiles Label10.Caption & vbBackSlash  'Untuk Buffer
ScanWithSpensav Label10.Caption, True
frmMain.cmdScan.Enabled = True
frmMain.Command9.Enabled = False  'Untuk Hentikan Proses Scan
frmMain.tmrInformasi.Enabled = False
MsgBox "Scan finished !", vbInformation, ""
End Sub

Private Sub Command2_Click()
Naik = False
Timer2.Enabled = True
End Sub

Private Sub Command4_Click()
frmMain.Show
frmMain.Picture20(7).Visible = False
frmMain.Picture20(8).Visible = False
frmMain.Picture10.Visible = True
frmMain.Picture11.Visible = False
frmMain.Picture9.Visible = False
frmMain.Picture8.Visible = False
Call frmMain.Button
Unload Me
End Sub

Private Sub Form_Load()
GradientPic frmRTPControls.Picture2, &H404040, &H0&, gmHorizontal: GradientPic frmUSB.Picture4, &H404040, &H0&, gmHorizontal: GradientPic frmEksControls.Picture2, &H404040, &H0&, gmHorizontal
    Top = ((GetSystemMetrics(17) + GetSystemMetrics(4)) * Screen.TwipsPerPixelY)
    Left = (GetSystemMetrics(16) * Screen.TwipsPerPixelX) - Width
    Naik = True
End Sub
Sub Geser(Apa As PictureBox) 'Ini adalah sub yg digunakan utk memberikan animasi gerakan pada Step Window
Do Until Apa.Left = 0           'Lakukan sampai step window yg dipilih berada pada posisi paling kiri
DoEvents
    For Each Control In Me.Controls
        If TypeOf Control Is PictureBox Then
        Control.Left = Control.Left - 1         'Gerakkan ke kiri
        End If
    Next Control

    Do Until i = 750            'Hanya Delay, agar animasi tidak terlalu cepat
    i = i + 1
    Loop
    i = 0
Loop
End Sub
Sub UntukMaximize(Apa As PictureBox) 'Ini adalah sub yg digunakan utk memberikan animasi gerakan pada Step Window
Do Until Apa.Left = 0           'Lakukan sampai step window yg dipilih berada pada posisi paling kiri
DoEvents
    For Each Control In Me.Controls
        If TypeOf Control Is PictureBox Then
        Control.Left = Control.Left - 1         'Gerakkan ke kiri
        End If
    Next Control

    Do Until i = 750            'Hanya Delay, agar animasi tidak terlalu cepat
    i = i + 1
    Loop
    i = 0
Loop
frmUSB.Show
End Sub

Private Sub Label12_Click()
Naik = False
Timer2.Enabled = True
End Sub
Private Sub Label81_Click()
Naik = False
Timer2.Enabled = True
End Sub
Private Sub Timer1_Timer()
Label9.Caption = frmMain.lvVirus.ListItems.Count
End Sub

Private Sub Timer2_Timer()
Const s = 80 'kecepatan gerak / slide
    Dim V As Single
    V = (GetSystemMetrics(17) + GetSystemMetrics(4)) * Screen.TwipsPerPixelY
    
    If Naik = True Then
        If Top - s <= V - Height Then
            Top = Top - (Top - (V - Height))
            Timer2.Enabled = False
        Else
            Top = Top - s
        End If
        
    Else
        Top = Top + s
        If Top >= V Then Unload Me
    End If
End Sub

Private Sub Timer3_Timer()
Picture4.Left = 4560
End Sub
