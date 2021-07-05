VERSION 5.00
Begin VB.Form frmRTPControls 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3105
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5505
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleMode       =   0  'User
   ScaleWidth      =   322.436
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BackColor       =   &H000000C0&
      BorderStyle     =   0  'None
      Height          =   3135
      Left            =   0
      ScaleHeight     =   209
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   369
      TabIndex        =   0
      Top             =   0
      Width           =   5535
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   1080
         Top             =   2400
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Lihat List Virus..."
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
         TabIndex        =   8
         Top             =   2400
         Width           =   1815
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   3135
         Left            =   4560
         ScaleHeight     =   3135
         ScaleWidth      =   975
         TabIndex        =   1
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
            MouseIcon       =   "frmRTPControls.frx":0000
            MousePointer    =   99  'Custom
            TabIndex        =   14
            Top             =   120
            Width           =   4095
         End
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "VIRUS TERDETEKSI !!"
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
         TabIndex        =   13
         Top             =   240
         Width           =   4095
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Spensav Menemukan Virus Di Sistem Operasi Anda, Silahkan Lihat List Untuk Mengeksekusi Virus."
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
         Height          =   615
         Left            =   240
         TabIndex        =   12
         Top             =   600
         Width           =   4095
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
         ForeColor       =   &H00C0C0FF&
         Height          =   375
         Left            =   1440
         TabIndex        =   11
         Top             =   1920
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
         ForeColor       =   &H00C0C0FF&
         Height          =   255
         Left            =   1440
         TabIndex        =   10
         Top             =   1560
         Width           =   2895
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
         ForeColor       =   &H00C0C0FF&
         Height          =   255
         Left            =   1440
         TabIndex        =   9
         Top             =   1200
         Width           =   2895
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
         ForeColor       =   &H00C0C0FF&
         Height          =   255
         Left            =   1320
         TabIndex        =   7
         Top             =   1920
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
         ForeColor       =   &H00C0C0FF&
         Height          =   255
         Left            =   1320
         TabIndex        =   6
         Top             =   1560
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
         ForeColor       =   &H00C0C0FF&
         Height          =   255
         Left            =   1320
         TabIndex        =   5
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Lokasi Virus"
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
         TabIndex        =   4
         Top             =   1920
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
         ForeColor       =   &H00C0C0FF&
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Virus"
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
         TabIndex        =   2
         Top             =   1200
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmRTPControls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Dim Naik As Boolean

Private Sub Command1_Click()
Naik = False
frmRTP.Show
Timer1.Enabled = True
End Sub

Private Sub Form_Load()
GradientPic frmRTPControls.Picture2, &H404040, &H0&, gmHorizontal: GradientPic frmUSB.Picture4, &H404040, &H0&, gmHorizontal: GradientPic frmEksControls.Picture2, &H404040, &H0&, gmHorizontal
    GradientPic Picture1, &H80&, &HC0&, gmVertical
    Top = ((GetSystemMetrics(17) + GetSystemMetrics(4)) * Screen.TwipsPerPixelY)
    Left = (GetSystemMetrics(16) * Screen.TwipsPerPixelX) - Width
    Naik = True
End Sub

Private Sub Label12_Click()
Naik = False
Timer1.Enabled = True
End Sub

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
