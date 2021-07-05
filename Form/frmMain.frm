VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   "Spensav-AntiVir"
   ClientHeight    =   7620
   ClientLeft      =   -45
   ClientTop       =   -435
   ClientWidth     =   10350
   LinkTopic       =   "Form1"
   ScaleHeight     =   7620
   ScaleWidth      =   10350
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture29 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   135
      Left            =   11160
      ScaleHeight     =   135
      ScaleWidth      =   8055
      TabIndex        =   150
      Top             =   840
      Width           =   8055
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   2040
         X2              =   2040
         Y1              =   0
         Y2              =   480
      End
   End
   Begin VB.PictureBox Picture28 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   135
      Left            =   10800
      ScaleHeight     =   135
      ScaleWidth      =   10215
      TabIndex        =   140
      Top             =   6120
      Width           =   10215
   End
   Begin Scanner.XPFrame XPFrame2 
      Height          =   5175
      Left            =   120
      TabIndex        =   71
      Top             =   2160
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   9128
      BackLightColor  =   16777215
      BackDarkColor   =   16777215
      BorderColor     =   8421504
      Style           =   1
      Curvature       =   0
      Begin VB.ListBox List2 
         Height          =   960
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   163
         Top             =   3840
         Visible         =   0   'False
         Width           =   1815
      End
      Begin Scanner.jcbutton rButton2 
         Height          =   735
         Left            =   120
         TabIndex        =   164
         Top             =   840
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   1296
         ButtonStyle     =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483633
         Caption         =   "Analyze Virus"
         Picture         =   "frmMain.frx":0000
         PictureAlign    =   0
         UseMaskCOlor    =   -1  'True
      End
      Begin Scanner.jcbutton rButton1 
         Height          =   735
         Left            =   120
         TabIndex        =   165
         Top             =   120
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1296
         ButtonStyle     =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         Caption         =   "Overview"
         Picture         =   "frmMain.frx":00F4
         PictureAlign    =   0
         UseMaskCOlor    =   -1  'True
      End
      Begin Scanner.jcbutton rButton3 
         Height          =   735
         Left            =   120
         TabIndex        =   166
         Top             =   2280
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   1296
         ButtonStyle     =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483633
         Caption         =   "Application"
         Picture         =   "frmMain.frx":01C5
         PictureAlign    =   0
         UseMaskCOlor    =   -1  'True
      End
      Begin Scanner.jcbutton rButton5 
         Height          =   735
         Left            =   120
         TabIndex        =   167
         Top             =   1560
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   1296
         ButtonStyle     =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483633
         Caption         =   "Settings"
         Picture         =   "frmMain.frx":02C7
         PictureAlign    =   0
         UseMaskCOlor    =   -1  'True
      End
      Begin Scanner.jcbutton rButton4 
         Height          =   735
         Left            =   120
         TabIndex        =   168
         Top             =   3000
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   1296
         ButtonStyle     =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483633
         Caption         =   "Database"
         Picture         =   "frmMain.frx":03A7
         PictureAlign    =   0
         UseMaskCOlor    =   -1  'True
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   5175
      Left            =   2160
      ScaleHeight     =   5175
      ScaleWidth      =   8055
      TabIndex        =   0
      Top             =   2160
      Width           =   8055
      Begin VB.PictureBox Picture12 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   5175
         Left            =   0
         ScaleHeight     =   5175
         ScaleWidth      =   8055
         TabIndex        =   55
         Top             =   0
         Width           =   8055
         Begin VB.CommandButton Command10 
            Caption         =   "Pengaturan"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   6000
            TabIndex        =   69
            Top             =   4440
            Width           =   1695
         End
         Begin VB.Image imgAman 
            Height          =   690
            Index           =   3
            Left            =   480
            Picture         =   "frmMain.frx":045E
            Top             =   1200
            Visible         =   0   'False
            Width           =   600
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Register"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   10
            Left            =   5160
            TabIndex        =   160
            Top             =   2040
            Width           =   1095
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   ": 1F4U1215"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   9
            Left            =   5880
            TabIndex        =   159
            Top             =   2040
            Width           =   2055
         End
         Begin VB.Image imgUpdate 
            Height          =   720
            Left            =   4320
            Picture         =   "frmMain.frx":1302
            Top             =   1200
            Width           =   720
         End
         Begin VB.Image imgAman 
            Height          =   690
            Index           =   2
            Left            =   480
            Picture         =   "frmMain.frx":2546
            Top             =   2760
            Width           =   600
         End
         Begin VB.Image imgAman 
            Height          =   600
            Index           =   1
            Left            =   4440
            Picture         =   "frmMain.frx":33EA
            Top             =   2760
            Width           =   600
         End
         Begin VB.Image imgAman 
            Height          =   690
            Index           =   0
            Left            =   480
            Picture         =   "frmMain.frx":40AE
            Top             =   1200
            Width           =   600
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   ": 01 Oktober 2012"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   17
            Left            =   5880
            TabIndex        =   68
            Top             =   1800
            Width           =   2055
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   ": Ansav Core && Internal"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   16
            Left            =   5880
            TabIndex        =   67
            Top             =   1560
            Width           =   2175
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Definisi"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   15
            Left            =   5160
            TabIndex        =   66
            Top             =   1800
            Width           =   1095
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Fitur Yang Dapat Memperbaiki Kerusakan Value/Data Registri Yang Rusak Atau Terinfeksi Virus."
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   855
            Index           =   14
            Left            =   5160
            TabIndex        =   65
            Top             =   3120
            Width           =   2415
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Registry Scanner"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   5160
            MouseIcon       =   "frmMain.frx":4F52
            MousePointer    =   99  'Custom
            TabIndex        =   64
            Top             =   2760
            Width           =   2775
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Engine"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   13
            Left            =   5160
            TabIndex        =   63
            Top             =   1560
            Width           =   1095
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Update Revinition"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   5160
            MouseIcon       =   "frmMain.frx":50A4
            MousePointer    =   99  'Custom
            TabIndex        =   62
            Top             =   1200
            Width           =   2775
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Fitur Yang Dapat Melindungi Operasi Sistem Dari Program Yang Bersifat Infected."
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   855
            Index           =   12
            Left            =   1200
            TabIndex        =   61
            Top             =   3120
            Width           =   2415
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "Process Monitoring"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   1200
            MouseIcon       =   "frmMain.frx":51F6
            MousePointer    =   99  'Custom
            TabIndex        =   60
            Top             =   2760
            Width           =   2775
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Fitur Yang Dapat Mendeteksi Rootkit, Spyware, Malware, Trojan, Worm, Backdoor, Dan Sejenisnya."
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   855
            Index           =   11
            Left            =   1200
            TabIndex        =   59
            Top             =   1560
            Width           =   2415
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Real-Time Protection"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   1200
            MouseIcon       =   "frmMain.frx":5348
            MousePointer    =   99  'Custom
            TabIndex        =   58
            Top             =   1200
            Width           =   2775
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Tampilan Utama"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   9
            Left            =   5400
            TabIndex        =   57
            Top             =   240
            Width           =   2415
         End
         Begin VB.Shape Shape2 
            BorderColor     =   &H00C0C0C0&
            Height          =   495
            Left            =   5400
            Top             =   120
            Width           =   15
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "TKJSI Protektor"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   8
            Left            =   120
            TabIndex        =   56
            Top             =   120
            Width           =   7935
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H00C0C0C0&
            Height          =   15
            Index           =   8
            Left            =   120
            Top             =   600
            Width           =   7815
         End
      End
      Begin VB.PictureBox Picture7 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   5175
         Left            =   0
         ScaleHeight     =   5175
         ScaleWidth      =   8055
         TabIndex        =   40
         Top             =   0
         Width           =   8055
         Begin VB.PictureBox Picture11 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   4455
            Left            =   240
            ScaleHeight     =   4455
            ScaleWidth      =   7575
            TabIndex        =   52
            Top             =   480
            Width           =   7575
            Begin VB.PictureBox Picture27 
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   495
               Left            =   360
               ScaleHeight     =   495
               ScaleWidth      =   495
               TabIndex        =   136
               Top             =   1200
               Width           =   495
               Begin VB.Image Image10 
                  Height          =   495
                  Left            =   0
                  Top             =   0
                  Width           =   495
               End
            End
            Begin VB.PictureBox Picture26 
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   495
               Left            =   360
               ScaleHeight     =   495
               ScaleWidth      =   495
               TabIndex        =   135
               Top             =   240
               Width           =   495
               Begin VB.Image Image9 
                  Height          =   495
                  Left            =   0
                  Top             =   0
                  Width           =   495
               End
            End
            Begin Scanner.DirTree DirTree1 
               Height          =   2175
               Left            =   360
               TabIndex        =   54
               Top             =   2040
               Width           =   6855
               _ExtentX        =   12091
               _ExtentY        =   3836
            End
            Begin VB.PictureBox Pic1 
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   735
               Left            =   960
               ScaleHeight     =   735
               ScaleWidth      =   5415
               TabIndex        =   131
               Top             =   3000
               Width           =   5415
               Begin VB.Label Label27 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Pindai Area Yang Di tentukan Pada List Di Bawah !!"
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
                  Left            =   120
                  TabIndex        =   129
                  Top             =   360
                  Width           =   5535
               End
               Begin VB.Label cmdScan 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Scan Now !"
                  BeginProperty Font 
                     Name            =   "Segoe UI"
                     Size            =   12
                     Charset         =   0
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Left            =   360
                  MouseIcon       =   "frmMain.frx":549A
                  MousePointer    =   99  'Custom
                  TabIndex        =   132
                  Top             =   0
                  Width           =   2175
               End
               Begin VB.Image Image6 
                  Height          =   570
                  Left            =   0
                  Top             =   0
                  Width           =   5370
               End
            End
            Begin VB.PictureBox Pic2 
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   735
               Left            =   960
               ScaleHeight     =   735
               ScaleWidth      =   5415
               TabIndex        =   130
               Top             =   2160
               Width           =   5415
               Begin VB.Label Label25 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Pindai Pada Semua Disk-Drive ( My Computer )"
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
                  TabIndex        =   134
                  Top             =   360
                  Width           =   5535
               End
               Begin VB.Label Label24 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Scan All Drive"
                  BeginProperty Font 
                     Name            =   "Segoe UI"
                     Size            =   12
                     Charset         =   0
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Left            =   360
                  MouseIcon       =   "frmMain.frx":55EC
                  MousePointer    =   99  'Custom
                  TabIndex        =   133
                  Top             =   0
                  Width           =   5295
               End
               Begin VB.Image Image5 
                  Height          =   570
                  Left            =   0
                  Top             =   0
                  Width           =   5370
               End
            End
            Begin VB.CommandButton Command11 
               Caption         =   "S"
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
               Left            =   6840
               TabIndex        =   53
               Top             =   1440
               Width           =   375
            End
            Begin VB.Label Label80 
               BackStyle       =   0  'Transparent
               Caption         =   "Pindai Pada Semua Disk-Drive ( My Computer )"
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   960
               TabIndex        =   128
               Top             =   600
               Width           =   5535
            End
            Begin VB.Label Label79 
               BackStyle       =   0  'Transparent
               Caption         =   "Scan All Drive"
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   960
               MouseIcon       =   "frmMain.frx":573E
               MousePointer    =   99  'Custom
               TabIndex        =   127
               Top             =   240
               Width           =   5295
            End
            Begin VB.Label Label78 
               BackStyle       =   0  'Transparent
               Caption         =   "Pindai Area Yang Di tentukan Pada List Di Bawah !!"
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   960
               TabIndex        =   126
               Top             =   1560
               Width           =   5535
            End
            Begin VB.Label Command123 
               BackStyle       =   0  'Transparent
               Caption         =   "Scan Now !"
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   960
               MouseIcon       =   "frmMain.frx":5890
               MousePointer    =   99  'Custom
               TabIndex        =   125
               Top             =   1200
               Width           =   2175
            End
         End
         Begin VB.PictureBox Picture10 
            BorderStyle     =   0  'None
            Height          =   4455
            Left            =   240
            ScaleHeight     =   297
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   505
            TabIndex        =   49
            Top             =   480
            Width           =   7575
            Begin VB.Timer tmrWaktu 
               Enabled         =   0   'False
               Interval        =   1000
               Left            =   1560
               Top             =   600
            End
            Begin VB.TextBox txtPindaiSel 
               Height          =   285
               Left            =   7320
               TabIndex        =   157
               Text            =   "Text3"
               Top             =   1080
               Visible         =   0   'False
               Width           =   150
            End
            Begin VB.PictureBox Picture25 
               BackColor       =   &H00FFFFFF&
               Height          =   2055
               Left            =   120
               ScaleHeight     =   1995
               ScaleWidth      =   7275
               TabIndex        =   110
               Top             =   2280
               Width           =   7335
               Begin VB.CommandButton Command8 
                  Caption         =   "&Kecil Pemindai"
                  Enabled         =   0   'False
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
                  Left            =   360
                  TabIndex        =   124
                  Top             =   1440
                  Width           =   1815
               End
               Begin VB.CommandButton Command9 
                  Caption         =   "&Hentikan Pemindai"
                  Enabled         =   0   'False
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
                  Left            =   5160
                  TabIndex        =   123
                  Top             =   1440
                  Width           =   1815
               End
               Begin VB.Label Label23 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   ": 00:00:00"
                  BeginProperty Font 
                     Name            =   "Segoe UI"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   6000
                  TabIndex        =   122
                  Top             =   960
                  Width           =   1215
               End
               Begin VB.Label Label22 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   ": 0"
                  BeginProperty Font 
                     Name            =   "Segoe UI"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   6000
                  TabIndex        =   121
                  Top             =   600
                  Width           =   1215
               End
               Begin VB.Label Label21 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   ": 0 /s"
                  BeginProperty Font 
                     Name            =   "Segoe UI"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   6000
                  TabIndex        =   120
                  Top             =   240
                  Width           =   1215
               End
               Begin VB.Label Label20 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Waktu"
                  BeginProperty Font 
                     Name            =   "Segoe UI"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   5040
                  TabIndex        =   119
                  Top             =   960
                  Width           =   1215
               End
               Begin VB.Label Label19 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Jum.File"
                  BeginProperty Font 
                     Name            =   "Segoe UI"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   5040
                  TabIndex        =   118
                  Top             =   600
                  Width           =   1215
               End
               Begin VB.Label Label18 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Kecepatan"
                  BeginProperty Font 
                     Name            =   "Segoe UI"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   5040
                  TabIndex        =   117
                  Top             =   240
                  Width           =   1215
               End
               Begin VB.Label Label17 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   ": 0"
                  BeginProperty Font 
                     Name            =   "Segoe UI"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   1200
                  TabIndex        =   116
                  Top             =   960
                  Width           =   1215
               End
               Begin VB.Label Label16 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   ": 0"
                  BeginProperty Font 
                     Name            =   "Segoe UI"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   1200
                  TabIndex        =   115
                  Top             =   600
                  Width           =   1215
               End
               Begin VB.Label Label7 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   ": 0"
                  BeginProperty Font 
                     Name            =   "Segoe UI"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   1200
                  TabIndex        =   114
                  Top             =   240
                  Width           =   1215
               End
               Begin VB.Label Label6 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Direktori"
                  BeginProperty Font 
                     Name            =   "Segoe UI"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   240
                  TabIndex        =   113
                  Top             =   960
                  Width           =   1215
               End
               Begin VB.Label Label5 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Registri"
                  BeginProperty Font 
                     Name            =   "Segoe UI"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   240
                  TabIndex        =   112
                  Top             =   600
                  Width           =   1215
               End
               Begin VB.Label Label4 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Virus"
                  BeginProperty Font 
                     Name            =   "Segoe UI"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   240
                  TabIndex        =   111
                  Top             =   240
                  Width           =   1215
               End
            End
            Begin ComctlLib.ProgressBar ProgressBar1 
               Height          =   255
               Left            =   1560
               TabIndex        =   85
               Top             =   1080
               Width           =   5895
               _ExtentX        =   10398
               _ExtentY        =   450
               _Version        =   327682
               Appearance      =   0
            End
            Begin VB.TextBox Text11 
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
               Left            =   120
               Locked          =   -1  'True
               TabIndex        =   84
               Top             =   1080
               Width           =   1335
            End
            Begin VB.TextBox Text6 
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
               Left            =   120
               Locked          =   -1  'True
               TabIndex        =   51
               Top             =   1800
               Width           =   7335
            End
            Begin VB.TextBox txtPindai 
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
               Left            =   120
               Locked          =   -1  'True
               TabIndex        =   50
               Top             =   360
               Width           =   7335
            End
            Begin VB.Label Label29 
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Progress Bar :"
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   139
               Top             =   840
               Width           =   2175
            End
            Begin VB.Label Label28 
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Status Pemindai :"
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   138
               Top             =   1560
               Width           =   2175
            End
            Begin VB.Label Label26 
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Proses Pemindai :"
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   137
               Top             =   120
               Width           =   2175
            End
         End
         Begin VB.PictureBox Picture20 
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   615
            Index           =   7
            Left            =   120
            ScaleHeight     =   615
            ScaleWidth      =   7815
            TabIndex        =   143
            Top             =   480
            Visible         =   0   'False
            Width           =   7815
            Begin VB.CheckBox Check4 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               Caption         =   "Cek Semua"
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   6360
               MaskColor       =   &H00FFFFFF&
               TabIndex        =   161
               Top             =   240
               Width           =   1335
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
               Index           =   6
               Left            =   120
               TabIndex        =   144
               Top             =   120
               Width           =   5055
            End
         End
         Begin VB.PictureBox Picture20 
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   615
            Index           =   8
            Left            =   120
            ScaleHeight     =   615
            ScaleWidth      =   7815
            TabIndex        =   141
            Top             =   480
            Visible         =   0   'False
            Width           =   7815
            Begin VB.Label Label8 
               BackStyle       =   0  'Transparent
               Caption         =   "Registri Infeksi/Berubah"
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
               Index           =   5
               Left            =   120
               TabIndex        =   142
               Top             =   120
               Width           =   5055
            End
         End
         Begin VB.PictureBox Picture9 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   4455
            Left            =   240
            ScaleHeight     =   4455
            ScaleWidth      =   7575
            TabIndex        =   44
            Top             =   480
            Width           =   7575
            Begin VB.CommandButton Command18 
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
               Left            =   1680
               TabIndex        =   158
               Top             =   3960
               Width           =   1575
            End
            Begin ComctlLib.ListView lvVirus 
               Height          =   3135
               Left            =   0
               TabIndex        =   45
               Top             =   720
               Width           =   7575
               _ExtentX        =   13361
               _ExtentY        =   5530
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
            Begin VB.CommandButton Command7 
               Caption         =   "Karantina Semua"
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
               Left            =   6000
               TabIndex        =   47
               Top             =   3960
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
               Left            =   0
               TabIndex        =   46
               Top             =   3960
               Width           =   1575
            End
         End
         Begin VB.PictureBox Picture8 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   4455
            Left            =   240
            ScaleHeight     =   4455
            ScaleWidth      =   7575
            TabIndex        =   42
            Top             =   480
            Width           =   7575
            Begin ComctlLib.ListView lvReg 
               Height          =   3135
               Left            =   0
               TabIndex        =   48
               Top             =   720
               Width           =   7575
               _ExtentX        =   13361
               _ExtentY        =   5530
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
                  Text            =   "Nama Value"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
                  SubItemIndex    =   1
                  Key             =   ""
                  Object.Tag             =   ""
                  Text            =   "Status"
                  Object.Width           =   1482
               EndProperty
               BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
                  SubItemIndex    =   2
                  Key             =   ""
                  Object.Tag             =   ""
                  Text            =   "Lokasi Value"
                  Object.Width           =   7832
               EndProperty
            End
            Begin VB.CommandButton Command5 
               Caption         =   "Perbaiki Semua"
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
               Left            =   0
               TabIndex        =   43
               Top             =   3960
               Width           =   1575
            End
         End
         Begin ComctlLib.TabStrip TabStrip3 
            Height          =   4935
            Left            =   120
            TabIndex        =   41
            Top             =   120
            Width           =   7815
            _ExtentX        =   13785
            _ExtentY        =   8705
            _Version        =   327682
            BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
               NumTabs         =   4
               BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
                  Caption         =   " &Path Scanned "
                  Key             =   ""
                  Object.Tag             =   ""
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
                  Caption         =   " &Current Report "
                  Key             =   ""
                  Object.Tag             =   ""
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
                  Caption         =   " &Virus "
                  Key             =   ""
                  Object.Tag             =   ""
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
                  Caption         =   " &Registry "
                  Key             =   ""
                  Object.Tag             =   ""
                  ImageVarType    =   2
               EndProperty
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Segoe UI"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.PictureBox Picture21 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   5175
         Left            =   0
         ScaleHeight     =   5175
         ScaleWidth      =   8055
         TabIndex        =   80
         Top             =   0
         Width           =   8055
         Begin VB.PictureBox Picture20 
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   615
            Index           =   4
            Left            =   120
            ScaleHeight     =   615
            ScaleWidth      =   7815
            TabIndex        =   145
            Top             =   480
            Width           =   7815
            Begin VB.Label Label8 
               BackStyle       =   0  'Transparent
               Caption         =   "Konfigurasi User"
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
               Index           =   11
               Left            =   120
               TabIndex        =   146
               Top             =   120
               Width           =   5055
            End
         End
         Begin VB.PictureBox Picture23 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   4455
            Left            =   240
            ScaleHeight     =   4455
            ScaleWidth      =   7575
            TabIndex        =   83
            Top             =   480
            Width           =   7575
            Begin VB.CommandButton Command16 
               Caption         =   "Simpan Pengaturan"
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   615
               Left            =   5880
               TabIndex        =   155
               Top             =   3720
               Width           =   1575
            End
            Begin Scanner.mm_checkbox SettingUser 
               Height          =   225
               Index           =   0
               Left            =   120
               TabIndex        =   87
               Top             =   840
               Width           =   360
               _ExtentX        =   635
               _ExtentY        =   397
               RoundedValue    =   10
            End
            Begin Scanner.mm_checkbox SettingUser 
               Height          =   225
               Index           =   1
               Left            =   120
               TabIndex        =   88
               Top             =   1200
               Width           =   360
               _ExtentX        =   635
               _ExtentY        =   397
               RoundedValue    =   10
            End
            Begin Scanner.mm_checkbox SettingUser 
               Height          =   225
               Index           =   2
               Left            =   120
               TabIndex        =   89
               Top             =   1560
               Width           =   360
               _ExtentX        =   635
               _ExtentY        =   397
               RoundedValue    =   10
            End
            Begin Scanner.mm_checkbox SettingUser 
               Height          =   225
               Index           =   3
               Left            =   120
               TabIndex        =   90
               Top             =   1920
               Width           =   360
               _ExtentX        =   635
               _ExtentY        =   397
               RoundedValue    =   10
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Aktifkan Always On Top ""Selalu Di Depan Aplikasi"""
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   3
               Left            =   600
               TabIndex        =   94
               Top             =   1920
               Width           =   7215
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Aktifkan Context Menu Pada Sistem Operasi"
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   2
               Left            =   600
               TabIndex        =   93
               Top             =   1560
               Width           =   7215
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Aktifkan Run As Startup ""Berjalan saat Memulai Windows"""
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   1
               Left            =   600
               TabIndex        =   92
               Top             =   1200
               Width           =   7215
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Aktifkan Transparant Pada Spensav"
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   0
               Left            =   600
               TabIndex        =   91
               Top             =   840
               Width           =   7215
            End
         End
         Begin VB.PictureBox Picture22 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   4455
            Left            =   240
            ScaleHeight     =   4455
            ScaleWidth      =   7575
            TabIndex        =   82
            Top             =   480
            Width           =   7575
            Begin VB.CommandButton Command17 
               Caption         =   "Simpan Pengaturan"
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   615
               Left            =   5880
               TabIndex        =   156
               Top             =   3720
               Width           =   1575
            End
            Begin Scanner.mm_checkbox SettingPelindung 
               Height          =   225
               Index           =   0
               Left            =   120
               TabIndex        =   99
               Top             =   840
               Width           =   360
               _ExtentX        =   635
               _ExtentY        =   397
               RoundedValue    =   10
            End
            Begin Scanner.mm_checkbox SettingPelindung 
               Height          =   225
               Index           =   1
               Left            =   120
               TabIndex        =   100
               Top             =   1200
               Width           =   360
               _ExtentX        =   635
               _ExtentY        =   397
               RoundedValue    =   10
            End
            Begin Scanner.mm_checkbox SettingPelindung 
               Height          =   225
               Index           =   2
               Left            =   120
               TabIndex        =   101
               Top             =   1560
               Width           =   360
               _ExtentX        =   635
               _ExtentY        =   397
               RoundedValue    =   10
            End
            Begin Scanner.mm_checkbox SettingPelindung 
               Height          =   225
               Index           =   3
               Left            =   120
               TabIndex        =   102
               Top             =   1920
               Width           =   360
               _ExtentX        =   635
               _ExtentY        =   397
               RoundedValue    =   10
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Aktifkan Pemindai ZIP dan RAR -Engine ""SPENSAVarchvScn.dll"""
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   7
               Left            =   600
               TabIndex        =   98
               Top             =   1920
               Width           =   7215
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Aktifkan Radar Behavor Spensav"
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   6
               Left            =   600
               TabIndex        =   97
               Top             =   1560
               Width           =   7215
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Aktifkan Firewall Sistem Operasi"
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   5
               Left            =   600
               TabIndex        =   96
               Top             =   1200
               Width           =   7215
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Aktifkan Pelindung Spensav-AntiVirus"
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   4
               Left            =   600
               TabIndex        =   95
               Top             =   840
               Width           =   7215
            End
         End
         Begin VB.PictureBox Picture31 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   4455
            Left            =   240
            ScaleHeight     =   4455
            ScaleWidth      =   7575
            TabIndex        =   151
            Top             =   480
            Width           =   7575
            Begin MSComDlg.CommonDialog CDialogPengecualian 
               Left            =   7080
               Top             =   3840
               _ExtentX        =   847
               _ExtentY        =   847
               _Version        =   393216
            End
            Begin VB.CommandButton Command15 
               Caption         =   "Bersihkan List"
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
               Left            =   1680
               TabIndex        =   154
               Top             =   3960
               Width           =   1575
            End
            Begin ComctlLib.ListView ListView3 
               Height          =   3135
               Left            =   0
               TabIndex        =   152
               Top             =   720
               Width           =   7575
               _ExtentX        =   13361
               _ExtentY        =   5530
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
                  Text            =   "Nama File"
                  Object.Width           =   3069
               EndProperty
               BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
                  SubItemIndex    =   1
                  Key             =   ""
                  Object.Tag             =   ""
                  Text            =   "Ukuran[B]"
                  Object.Width           =   1658
               EndProperty
               BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
                  SubItemIndex    =   2
                  Key             =   ""
                  Object.Tag             =   ""
                  Text            =   "Lokasi File"
                  Object.Width           =   7832
               EndProperty
            End
            Begin VB.CommandButton Command14 
               Caption         =   "Tambah File"
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
               Left            =   0
               TabIndex        =   153
               Top             =   3960
               Width           =   1575
            End
         End
         Begin ComctlLib.TabStrip TabStrip4 
            Height          =   4935
            Left            =   120
            TabIndex        =   81
            Top             =   120
            Width           =   7815
            _ExtentX        =   13785
            _ExtentY        =   8705
            _Version        =   327682
            BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
               NumTabs         =   3
               BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
                  Caption         =   " &General "
                  Key             =   ""
                  Object.Tag             =   ""
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
                  Caption         =   " &Protection "
                  Key             =   ""
                  Object.Tag             =   ""
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
                  Caption         =   "&Exeption Controls"
                  Key             =   ""
                  Object.Tag             =   ""
                  ImageVarType    =   2
               EndProperty
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Segoe UI"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   5175
         Left            =   0
         ScaleHeight     =   5175
         ScaleWidth      =   8055
         TabIndex        =   18
         Top             =   0
         Width           =   8055
         Begin VB.PictureBox Picture20 
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   615
            Index           =   3
            Left            =   120
            ScaleHeight     =   615
            ScaleWidth      =   7815
            TabIndex        =   147
            Top             =   480
            Width           =   7815
            Begin VB.Label Label8 
               BackStyle       =   0  'Transparent
               Caption         =   "Tambahkan Virus Dari User"
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
               Index           =   4
               Left            =   120
               TabIndex        =   148
               Top             =   120
               Width           =   5055
            End
         End
         Begin VB.PictureBox Picture6 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   4455
            Left            =   240
            ScaleHeight     =   4455
            ScaleWidth      =   7575
            TabIndex        =   30
            Top             =   480
            Width           =   7575
            Begin ComctlLib.ListView ListView2 
               Height          =   2775
               Left            =   0
               TabIndex        =   39
               Top             =   1680
               Width           =   7575
               _ExtentX        =   13361
               _ExtentY        =   4895
               View            =   3
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               _Version        =   327682
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
                  Text            =   "Lokasi File"
                  Object.Width           =   7832
               EndProperty
               BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
                  SubItemIndex    =   1
                  Key             =   ""
                  Object.Tag             =   ""
                  Text            =   "Nama File"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
                  SubItemIndex    =   2
                  Key             =   ""
                  Object.Tag             =   ""
                  Text            =   "Risk"
                  Object.Width           =   1658
               EndProperty
            End
            Begin MSComDlg.CommonDialog CommonDialog 
               Left            =   7560
               Top             =   480
               _ExtentX        =   847
               _ExtentY        =   847
               _Version        =   393216
            End
            Begin VB.CommandButton Command4 
               Caption         =   "Tambah Virus"
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   5760
               TabIndex        =   38
               Top             =   1200
               Width           =   1815
            End
            Begin VB.CommandButton Command3 
               Caption         =   "..."
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   7080
               TabIndex        =   37
               Top             =   840
               Width           =   495
            End
            Begin VB.TextBox Text2 
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1440
               TabIndex        =   36
               Top             =   1200
               Width           =   4215
            End
            Begin VB.TextBox Text1 
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1440
               TabIndex        =   35
               Top             =   840
               Width           =   5535
            End
            Begin VB.Label Label9 
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
               Height          =   255
               Index           =   8
               Left            =   1320
               TabIndex        =   34
               Top             =   1200
               Width           =   1455
            End
            Begin VB.Label Label9 
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
               Height          =   255
               Index           =   7
               Left            =   1320
               TabIndex        =   33
               Top             =   840
               Width           =   1455
            End
            Begin VB.Label Label9 
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
               Height          =   255
               Index           =   6
               Left            =   0
               TabIndex        =   32
               Top             =   1200
               Width           =   1455
            End
            Begin VB.Label Label9 
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
               Height          =   255
               Index           =   5
               Left            =   0
               TabIndex        =   31
               Top             =   840
               Width           =   1455
            End
         End
         Begin VB.PictureBox Picture19 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   4455
            Left            =   240
            ScaleHeight     =   4455
            ScaleWidth      =   7575
            TabIndex        =   72
            Top             =   480
            Width           =   7575
            Begin ComctlLib.ListView LVV 
               Height          =   3135
               Left            =   0
               TabIndex        =   74
               Top             =   720
               Width           =   7575
               _ExtentX        =   13361
               _ExtentY        =   5530
               View            =   3
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               _Version        =   327682
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
                  Text            =   "Nama"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
                  SubItemIndex    =   1
                  Key             =   ""
                  Object.Tag             =   ""
                  Text            =   "Lokasi"
                  Object.Width           =   6068
               EndProperty
               BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
                  SubItemIndex    =   2
                  Key             =   ""
                  Object.Tag             =   ""
                  Text            =   "Value"
                  Object.Width           =   7832
               EndProperty
            End
            Begin VB.CommandButton Command13 
               Caption         =   "Hapus Value"
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
               Left            =   1680
               TabIndex        =   76
               Top             =   3960
               Width           =   1575
            End
            Begin VB.CommandButton Command12 
               Caption         =   "Segarkan"
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
               Left            =   0
               TabIndex        =   75
               Top             =   3960
               Width           =   1575
            End
            Begin VB.PictureBox Pic 
               AutoRedraw      =   -1  'True
               Height          =   300
               Left            =   960
               ScaleHeight     =   240
               ScaleWidth      =   240
               TabIndex        =   73
               Top             =   1200
               Width           =   300
            End
            Begin ComctlLib.ImageList ima 
               Left            =   600
               Top             =   1440
               _ExtentX        =   1005
               _ExtentY        =   1005
               BackColor       =   -2147483643
               MaskColor       =   12632256
               _Version        =   327682
            End
         End
         Begin VB.PictureBox Picture5 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   4455
            Left            =   240
            ScaleHeight     =   4455
            ScaleWidth      =   7575
            TabIndex        =   26
            Top             =   480
            Width           =   7575
            Begin ComctlLib.ListView ListView1 
               Height          =   3135
               Left            =   0
               TabIndex        =   27
               Top             =   720
               Width           =   7575
               _ExtentX        =   13361
               _ExtentY        =   5530
               View            =   3
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               _Version        =   327682
               Icons           =   "ImageList1"
               SmallIcons      =   "ImageList1"
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
               NumItems        =   4
               BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
                  Key             =   ""
                  Object.Tag             =   ""
                  Text            =   "Nama Proses"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
                  SubItemIndex    =   1
                  Key             =   ""
                  Object.Tag             =   ""
                  Text            =   "Lokasi Process"
                  Object.Width           =   6068
               EndProperty
               BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
                  SubItemIndex    =   2
                  Key             =   ""
                  Object.Tag             =   ""
                  Text            =   "PID"
                  Object.Width           =   1482
               EndProperty
               BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
                  SubItemIndex    =   3
                  Key             =   ""
                  Object.Tag             =   ""
                  Text            =   "Module"
                  Object.Width           =   1658
               EndProperty
            End
            Begin VB.CommandButton CmdTerminate 
               Caption         =   "Matikan Proses"
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
               Left            =   0
               TabIndex        =   79
               Top             =   3960
               Width           =   1575
            End
            Begin VB.CommandButton CmdRefresh 
               Caption         =   "Segarkan"
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
               Left            =   1680
               TabIndex        =   78
               Top             =   3960
               Width           =   1575
            End
            Begin VB.CommandButton CmdExplore 
               Caption         =   "Lihat Lokasi..."
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
               TabIndex        =   77
               Top             =   3960
               Width           =   1455
            End
            Begin VB.CheckBox Check2 
               Caption         =   "Check2"
               Height          =   255
               Left            =   3600
               TabIndex        =   29
               Top             =   1800
               Width           =   1695
            End
            Begin VB.CheckBox Check1 
               Caption         =   "Check1"
               Height          =   255
               Left            =   3600
               TabIndex        =   28
               Top             =   1440
               Width           =   1695
            End
            Begin ComctlLib.ImageList ImageList1 
               Left            =   720
               Top             =   1680
               _ExtentX        =   1005
               _ExtentY        =   1005
               BackColor       =   -2147483643
               ImageWidth      =   16
               ImageHeight     =   16
               MaskColor       =   12632256
               _Version        =   327682
            End
         End
         Begin VB.PictureBox Picture4 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   4455
            Left            =   240
            ScaleHeight     =   4455
            ScaleWidth      =   7575
            TabIndex        =   20
            Top             =   480
            Width           =   7575
            Begin VB.CheckBox Check3 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Cek Semua"
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   0
               TabIndex        =   149
               Top             =   720
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.ListBox List1 
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   3210
               Left            =   0
               Style           =   1  'Checkbox
               TabIndex        =   21
               Top             =   1080
               Width           =   5775
            End
            Begin VB.DirListBox Dir1 
               Height          =   1440
               Left            =   360
               TabIndex        =   25
               Top             =   1080
               Width           =   2055
            End
            Begin VB.FileListBox File1 
               Height          =   1455
               Left            =   2520
               TabIndex        =   24
               Top             =   1080
               Width           =   2295
            End
            Begin VB.CommandButton Command2 
               Caption         =   "&Kembalikan File"
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   5880
               TabIndex        =   23
               Top             =   1560
               Width           =   1695
            End
            Begin VB.CommandButton Command1 
               Caption         =   "&Segarkan List"
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   5880
               TabIndex        =   22
               Top             =   1080
               Width           =   1695
            End
         End
         Begin ComctlLib.TabStrip TabStrip2 
            Height          =   4935
            Left            =   120
            TabIndex        =   19
            Top             =   120
            Width           =   7815
            _ExtentX        =   13785
            _ExtentY        =   8705
            _Version        =   327682
            BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
               NumTabs         =   4
               BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
                  Caption         =   " &Use Virus By.User "
                  Key             =   ""
                  Object.Tag             =   ""
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
                  Caption         =   " &Startup Manager "
                  Key             =   ""
                  Object.Tag             =   ""
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
                  Caption         =   " &Process Manager "
                  Key             =   ""
                  Object.Tag             =   ""
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
                  Caption         =   " &Quarantine "
                  Key             =   ""
                  Object.Tag             =   ""
                  ImageVarType    =   2
               EndProperty
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Segoe UI"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   5175
         Left            =   0
         ScaleHeight     =   5175
         ScaleWidth      =   8055
         TabIndex        =   2
         Top             =   0
         Width           =   8055
         Begin VB.ListBox lsDB 
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4740
            Left            =   240
            TabIndex        =   3
            Top             =   240
            Width           =   3735
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   ": SPENSAV AntiVirus"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   5880
            TabIndex        =   17
            Top             =   3240
            Width           =   1815
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   ": isfa"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   5880
            TabIndex        =   16
            Top             =   3000
            Width           =   1815
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   ": M.Isfahani Ghiyath.YM"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   5880
            TabIndex        =   15
            Top             =   2760
            Width           =   1935
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Programmer"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   4200
            TabIndex        =   14
            Top             =   3240
            Width           =   1815
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "a.k.a"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   4200
            TabIndex        =   13
            Top             =   3000
            Width           =   1815
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Nama"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   4200
            TabIndex        =   12
            Top             =   2760
            Width           =   1815
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H00C0C0C0&
            Height          =   15
            Index           =   1
            Left            =   4200
            Top             =   2520
            Width           =   3615
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Informasi Programmer"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   4200
            TabIndex        =   11
            Top             =   2040
            Width           =   3615
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   ": 500 Variant Virus"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   5880
            TabIndex        =   10
            Top             =   1440
            Width           =   1815
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   ": 16 September 2012"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   5880
            TabIndex        =   9
            Top             =   1200
            Width           =   1815
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   ": ..."
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   5880
            TabIndex        =   8
            Top             =   960
            Width           =   1815
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Variant Virus"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   4200
            TabIndex        =   7
            Top             =   1440
            Width           =   1815
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Waktu Perbaharui"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   4200
            TabIndex        =   6
            Top             =   1200
            Width           =   1815
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah Def.Virus"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   4200
            TabIndex        =   5
            Top             =   960
            Width           =   1815
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H00C0C0C0&
            Height          =   15
            Index           =   0
            Left            =   4200
            Top             =   720
            Width           =   3615
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Informasi DataBase"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   4200
            TabIndex        =   4
            Top             =   240
            Width           =   3615
         End
      End
      Begin ComctlLib.TabStrip TabStrip1 
         Height          =   5295
         Left            =   0
         TabIndex        =   1
         Top             =   6000
         Visible         =   0   'False
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   9340
         _Version        =   327682
         BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
            NumTabs         =   4
            BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "&Overview"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "&Scanner"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "&Application"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "&Information"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.PictureBox PicHeader 
      BackColor       =   &H000000C0&
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   12000
      ScaleHeight     =   1695
      ScaleWidth      =   10095
      TabIndex        =   108
      Top             =   1080
      Width           =   10095
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   120
         Top             =   1080
      End
      Begin VB.Image Image8 
         Height          =   1695
         Left            =   6480
         Top             =   480
         Width           =   4215
      End
      Begin VB.Image Image7 
         Height          =   1695
         Left            =   720
         Top             =   480
         Width           =   10095
      End
   End
   Begin VB.PictureBox Picture24 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   9600
      ScaleHeight     =   375
      ScaleWidth      =   495
      TabIndex        =   105
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
         MouseIcon       =   "frmMain.frx":59E2
         MousePointer    =   99  'Custom
         TabIndex        =   106
         Top             =   0
         Width           =   735
      End
   End
   Begin VB.PictureBox Picture18 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   9120
      ScaleHeight     =   375
      ScaleWidth      =   495
      TabIndex        =   103
      Top             =   0
      Width           =   495
      Begin VB.Label Label82 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   -120
         MouseIcon       =   "frmMain.frx":5B34
         MousePointer    =   99  'Custom
         TabIndex        =   104
         Top             =   0
         Width           =   735
      End
   End
   Begin VB.PictureBox sIcon 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   11760
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   86
      Top             =   3600
      Width           =   255
   End
   Begin VB.Timer tmrInformasi 
      Interval        =   1000
      Left            =   11760
      Top             =   3120
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   10680
      TabIndex        =   70
      Text            =   "Ini Checksumnya :D"
      Top             =   6840
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.PictureBox Picture17 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   7815
      Left            =   11640
      ScaleHeight     =   521
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   697
      TabIndex        =   107
      Top             =   1800
      Width           =   10455
      Begin VB.PictureBox Picture13 
         Height          =   1815
         Left            =   0
         ScaleHeight     =   1815
         ScaleWidth      =   15
         TabIndex        =   162
         Top             =   0
         Width           =   15
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Registered : Spensas"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   109
         Top             =   7440
         Width           =   8415
      End
   End
   Begin VB.Image Image1 
      Height          =   7785
      Left            =   0
      Picture         =   "frmMain.frx":5C86
      Top             =   -120
      Width           =   10440
   End
   Begin ComctlLib.ImageList ImgListView 
      Left            =   10560
      Top             =   3000
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
            Picture         =   "frmMain.frx":10E5E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":10E7BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":10EB0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":10EE26
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":10F178
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents SCANPROC As clsScanProc
Attribute SCANPROC.VB_VarHelpID = -1
Private Declare Function LockWindowUpdate Lib "user32.dll" (ByVal hwndLock As Long) As Long
Dim m_sProcess  As String
Dim m_sTime     As Single
Dim FILEICON    As clsGetIcon
'########## Untuk ListView ##########'
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, lParam As Any) As Long
'Untuk Registry
Const ERROR_SUCCESS = 0
Const REG_SZ = 1
Const REG_DWORD = 4
Const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_CLASSES_ROOT = &H80000000
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_USERS = &H80000003
Const HKEY_CURRENT_CONFIG = &H80000005
Const HKEY_PERFORMANCE_DATA = &H80000004
Const HKEY_DYN_DATA = &H80000006
'Pengumuman variabel
Private a As Integer, b As Integer
Dim Qfolder As String
Dim strRes(100) As String
Private path$, hSectionSharedArr&
Private FileToScan As Long
Private Function Quarantine1(strPath As String)
On Error Resume Next
Dim splitFile() As String
Dim rndFile As String
Dim i As Long
'Membuat Nilai Acak utk penamaan File
rndFile = Rnd * Val(Time)
rndFile = Round(Rnd * rndFile, 3)
SaveSetting "JHelp", "Quarantine", rndFile, strPath 'Simpan Setting di Registry
NormalizeAttribute strPath  'Normalkan file yang berattribute Hidden, System dan Read Only
'Bisa dimasukkan Routine Encrypt File disini, kemudian copy file yang terencrypt
DoEvents
FileCopy strPath, Qfolder & "\" & rndFile 'Copy File ke Folder Quarantine dengan nama Acak
Kill strPath 'Hapus File asli
End Function
Private Function Restore(rndPath As String)
On Error Resume Next
'Gunakan DirListBox dan FileListBox, biar lebih mudah
Dim Res As String
Dim ResFolder() As String
Dim Folder As String
Res = GetSetting("JHelp", "Quarantine", rndPath) 'Ambil Setting dari Registry
DeleteSetting "JHelp", "Quarantine", rndPath 'Hapus Setting dari Registry
ResFolder = Split(Res, "\") 'Memisahkan Nilai dengan Pemisah "\"
Folder = Replace(Res, ResFolder(UBound(ResFolder)), "") 'Ambil Hanya Directory dari Nama File tsb
If Dir(Folder, vbDirectory) = "" Then
    MkDir Folder 'Jika Tidak ada, dibuatkan
End If
'Bisa dimasukkan Routine Decrypt File disini, kemudian di Restore/Copy ke Path asal
DoEvents
FileCopy Qfolder & "\" & rndPath, Res 'Copy File dari Folder Quarantine ke Folder asli-nya
Kill Qfolder & "\" & rndPath 'Hapus File Quarantine
End Function
Function GetList()
On Error Resume Next
Dim i As Long
Dim getFile As String
File1.Refresh
For i = 0 To File1.ListCount - 1
    strRes(i) = File1.List(i) 'Ambil Nama File yang ada di Folder [Quarantine]
    getFile = GetSetting("JHelp", "Quarantine", strRes(i)) 'Ambil Setting dengan key nama Folder dari [Quarantine]
    List1.AddItem getFile
Next i
End Function
Private Sub Check3_Click()
On Error Resume Next
Dim i As Long
For i = 1 To List1.ListCount - 1
If Check3.value = 1 Then
List1.Selected(i) = True
Else
List1.Selected(i) = False
End If
Next
End Sub

Private Sub Check4_Click()
If Check4.value = 1 Then
SetCheckAllItems True
Else
SetCheckAllItems False
End If
End Sub
Private Sub CmdExplore_Click()
Shell "Explorer.exe " & Left(ListView1.SelectedItem.SubItems(1), _
Len(ListView1.SelectedItem.SubItems(1)) - Len(ListView1.SelectedItem)), _
vbNormalFocus
End Sub

Private Sub CmdRefresh_Click()
Call Check1_Click
End Sub

Private Sub cmdScan_Click()
Dim lstCek    As Collection
Set lstCek = New Collection
Dim iCount As Long
DirTree1.OutPutPath lstCek
For iCount = 1 To lstCek.Count
cmdScan.Enabled = False
clear_log
Cek_Value 'Untuk Scan Registri
tmrInformasi.Enabled = True
Picture10.Visible = True
Picture11.Visible = False
Picture9.Visible = False
Picture8.Visible = False
StopScan = False
Command9.Enabled = True 'Untuk Hentikan Proses Scan
AnalisaFiles lstCek(iCount) & vbBackSlash 'Untuk Buffer
Command8.Enabled = True 'Ini Untuk Minimize
tmrWaktu.Enabled = True
ScanWithSpensav lstCek(iCount), True
Command8.Enabled = False 'Ini Untuk Minimize
cmdScan.Enabled = True
Command9.Enabled = False  'Untuk Hentikan Proses Scan
Next
tmrWaktu.Enabled = False
tmrInformasi.Enabled = False
MsgBox "Scan finished !", vbInformation, ""
End Sub

Private Sub CmdTerminate_Click()
If MsgBox("Yakin Untuk Membunuh Proses terpilih ini ??", vbExclamation + vbYesNoCancel + vbDefaultButton2, "Terminate Process") = vbYes Then
        If (SCANPROC.TerminateProcess(ListView1.SelectedItem.SubItems(2)) = True) Then
            Call Check1_Click
        End If
    End If
End Sub
Private Sub Command1_Click() 'Untuk Refresh List Quarantine
GetList
End Sub

Private Sub Command11_Click()
DirTree1.LoadTreeDir False
End Sub

Private Sub Command12_Click()
GetStartup
End Sub
Private Sub Command123_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Pic1.Top = 1200
End Sub
Private Sub Command13_Click()
If Left(LVV.SelectedItem.SubItems(2), Len("HKEY_LOCAL_MACHINE\Software\Microsoft\Windows NT\CurrentVersion\Winlogon")) <> "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows NT\CurrentVersion\Winlogon" Then
    If Left(LVV.SelectedItem.SubItems(2), 18) = "HKEY_LOCAL_MACHINE" Then
        DelSetting HKEY_LOCAL_MACHINE, Right(LVV.SelectedItem.SubItems(2), Len(LVV.SelectedItem.SubItems(2)) - 19), LVV.SelectedItem.Text
    ElseIf Left(LVV.SelectedItem.SubItems(2), 17) = "HKEY_CURRENT_USER" Then
        DelSetting HKEY_CURRENT_USER, Right(LVV.SelectedItem.SubItems(2), Len(LVV.SelectedItem.SubItems(2)) - 18), LVV.SelectedItem.Text
    End If
        GetStartup
Else
End If
End Sub

Private Sub Command14_Click()
On Error Resume Next
CDialogPengecualian.ShowOpen
If CDialogPengecualian.Filename <> "" Then
Dim LV As ListItem
Set LV = ListView3.ListItems.Add(, , GetFileName(CDialogPengecualian.Filename), , ImgListView.ListImages(3).Index)
LV.SubItems(1) = FileLen(CDialogPengecualian.Filename)
LV.SubItems(2) = CDialogPengecualian.Filename
End If
End Sub

Private Sub Command16_Click()
Call simpansetting
MsgBox "Konfigurasi Di Terapkan", vbInformation, "Informasi"
End Sub

Private Sub Command17_Click()
Call simpansetting
MsgBox "Konfigurasi Di Terapkan", vbInformation, "Informasi"
End Sub

Private Sub Command18_Click()
Dim i As Integer
For i = 1 To lvVirus.ListItems.Count - 1
DoEvents
Call List_Process
Bunuh lvVirus.ListItems(i).SubItems(2)
NormalkanAtribut lvVirus.ListItems(i).SubItems(2)
DeleteFile lvVirus.ListItems(i).SubItems(2)
DeleteFile lvVirus.ListItems(i).SubItems(2)
DeleteFile lvVirus.ListItems(i).SubItems(2)
DeleteFile lvVirus.ListItems(i).SubItems(2)
DeleteFile lvVirus.ListItems(i).SubItems(2)
DeleteFile lvVirus.ListItems(i).SubItems(2)
With lvVirus.ListItems.Item(i)
.SmallIcon = frmMain.ImgListView.ListImages(5).Index
End With
Next
End Sub
Private Sub Command2_Click()
On Error GoTo err
Dim i As Long

For i = 0 To List1.ListCount - 1
    If List1.Selected(i) = True Then
        Restore strRes(i) 'Panggil Nama Acak yang dibuat, sebagai key dari Setting
        GetList
    End If
Next i
GetList
Exit Sub
err:
MsgBox "Silahkan Ulangi !!", vbInformation
GetList
End Sub

Private Sub Command3_Click()

CommonDialog.ShowOpen
Text1.Text = CommonDialog.Filename
End Sub

Private Sub Command4_Click()
Dim LV As ListItem
If Text2.Text <> "" Then
Set LV = ListView2.ListItems.Add(, , Text1.Text)
LV.SubItems(1) = Text2.Text
LV.SubItems(2) = FileLen(Text1.Text)
Else
MsgBox "Not Found!", vbCritical, ""
End If
End Sub

Private Sub Command5_Click()
Call Perbaiki
End Sub
Private Sub Command6_Click()
'On Error Resume Next
   Dim i As Integer
   Dim r As Long
   Dim b As String
   Dim SubItemText As String
   Dim LV As LV_ITEM
    If MsgBox("Yakin Untuk Menghapus Beberapa Objek Virus ??", vbYesNo) = vbYes Then
   For i = 0 To lvVirus.ListItems.Count - 1
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
         DeleteFile (Left$(LV.pszText, InStr(LV.pszText, Chr$(0)) - 1))
         DeleteFile (Left$(LV.pszText, InStr(LV.pszText, Chr$(0)) - 1))
         DeleteFile (Left$(LV.pszText, InStr(LV.pszText, Chr$(0)) - 1))
         DeleteFile (Left$(LV.pszText, InStr(LV.pszText, Chr$(0)) - 1))
         DeleteFile (Left$(LV.pszText, InStr(LV.pszText, Chr$(0)) - 1))
      End If
   Next
   End If
End Sub
Public Function ListView_GetItemText(i As Long, iSubItem As Long) As String
    Dim lpPitem As LVITEM
    Dim SubItemText As String
    SubItemText = String$(MAX_PATH, 0)
    lpPitem.iSubItem = iSubItem
    lpPitem.cchTextMax = MAX_PATH
    lpPitem.pszText = SubItemText
     
    Call SendMessage(lvVirus.hwnd, LVM_GETITEMTEXT, ByVal i, lpPitem)
    ListView_GetItemText = Left$(lpPitem.pszText, InStr(lpPitem.pszText, vbNullChar) - 1)
End Function
Private Sub Command7_Click()
Dim ii As Integer
For ii = 1 To lvVirus.ListItems.Count
DoEvents
Quarantine1 lvVirus.ListItems(ii).SubItems(2)
GetList
With lvVirus.ListItems.Item(ii)
.SmallIcon = frmMain.ImgListView.ListImages(5).Index
End With
Next ii
End Sub
Private Sub Command8_Click()
GradientPic frmUSB.Frame2, &H800000, &HFF0000, gmVertical
frmUSB.Label6.ForeColor = &HFFC0C0: frmUSB.Label7.ForeColor = &HFFC0C0: frmUSB.Label8.ForeColor = &HFFC0C0: frmUSB.Label9.ForeColor = &HFFC0C0
Me.Hide
frmUSB.UntukMaximize frmUSB.Frame2
End Sub
Private Sub Command9_Click()
StopScan = True
End Sub

'########## Saat Aplikasi dijalankan akan membuat folder karantina dan meload database virus##########'
Private Sub Form_Load()
On Error Resume Next
DropShadow frmMain.hwnd
BuatOval Me, 25
frmEksControls.Timer2.Enabled = True 'untuk eksekusi program
Call ceksetting 'cek dulu
Call RESFile
GradientPic PicHeader, &H80&, &HC0&, gmHorizontal
GetStartup
BacaDatabase App.path & "\Signature.dat"
Call ListView32 'Untuk Use Theme ListView
Call DatabaseEx
DirTree1.LoadTreeDir False
Set SCANPROC = New clsScanProc
Set FILEICON = New clsGetIcon
Call Check1_Click
MkDir "karantina"
BuildDatabase
'Routine ambil Nama Drive System
Qfolder = Environ("windir") 'Ambil Folder System
Qfolder = Replace(UCase(Qfolder), "WINDOWS", "") 'Dapatkan Nama Drive system
Qfolder = Qfolder & "Sys.Folder" 'Tambahkan "[Quarantine]", untuk nama Folder
'Membuat Folder [Quarantine], jika belum ada
If Dir(Qfolder, vbDirectory) = "" Then MkDir Qfolder
Dir1 = Qfolder
File1 = Dir1
GetList
End Sub
Sub clear_log()
List2.Clear 'Nggak tau, lupa :D
detik.i = 0: menit.i = 0: jam.i = 0: detik.s = 0: menit.s = 0: jam.s = 0 'Untuk Waktu
jmlFiles = 0: jmlDirs = 0: totalFiles1 = 0 'Untuk Jumlah Pada Buffer
txtPindai.Text = "": Text6.Text = "" 'Untuk Teks
lvVirus.ListItems.Clear: lvReg.ListItems.Clear 'Untuk ListView
jumlahDir = 0: jumlahFile = 0: jumlahVirus = 0 'Untuk Jumlah
End Sub
Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then 'left click
ReleaseCapture
SendMessage Me.hwnd, &HA1, 2, 0
End If
End Sub

Private Sub Image7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then 'left click
ReleaseCapture
SendMessage Me.hwnd, &HA1, 2, 0
End If
End Sub
Private Sub Image8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then 'left click
ReleaseCapture
SendMessage Me.hwnd, &HA1, 2, 0
End If
End Sub

Private Sub Label12_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label8(9).Caption = "Real-Time Protection"
End Sub
Private Sub Label14_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label8(9).Caption = "Update Revinition"
End Sub
Private Sub Label13_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label8(9).Caption = "Web Protection"
End Sub
Private Sub Label15_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label8(9).Caption = "Registry Scanner"
End Sub
Private Sub Label79_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Pic2.Top = 240
End Sub

Private Sub Label81_Click()
Me.Hide
End Sub

Private Sub Label82_Click()
Me.WindowState = 1
End Sub
Private Sub Picture11_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Pic2.Top = 2160
Pic1.Top = 2880
End Sub
Private Sub Picture17_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then 'left click
ReleaseCapture
SendMessage Me.hwnd, &HA1, 2, 0
End If
End Sub

Private Sub rButton1_Click()
rButton1.Width = 2055: rButton2.Width = 1935: rButton3.Width = 1935: rButton4.Width = 1935: rButton5.Width = 1935
rButton1.BackColor = &HFFFFFF: rButton2.BackColor = &H8000000F: rButton3.BackColor = &H8000000F: rButton4.BackColor = &H8000000F: rButton5.BackColor = &H8000000F
Picture12.Visible = True: Picture7.Visible = False: Picture3.Visible = False: Picture2.Visible = False: Picture21.Visible = False
End Sub
Private Sub rButton2_Click()
rButton2.Width = 2055: rButton1.Width = 1935: rButton3.Width = 1935: rButton4.Width = 1935: rButton5.Width = 1935
rButton2.BackColor = &HFFFFFF: rButton1.BackColor = &H8000000F: rButton3.BackColor = &H8000000F: rButton4.BackColor = &H8000000F: rButton5.BackColor = &H8000000F
Picture7.Visible = True: Picture12.Visible = False: Picture3.Visible = False: Picture2.Visible = False: Picture21.Visible = False
End Sub
Public Sub Button()
Call rButton2_Click
End Sub
Private Sub rButton3_Click()
rButton3.Width = 2055: rButton2.Width = 1935: rButton1.Width = 1935: rButton4.Width = 1935: rButton5.Width = 1935
rButton3.BackColor = &HFFFFFF: rButton2.BackColor = &H8000000F: rButton1.BackColor = &H8000000F: rButton4.BackColor = &H8000000F: rButton5.BackColor = &H8000000F
Picture3.Visible = True: Picture12.Visible = False: Picture7.Visible = False: Picture2.Visible = False: Picture21.Visible = False
End Sub

Private Sub rButton4_Click()
rButton4.Width = 2055: rButton2.Width = 1935: rButton3.Width = 1935: rButton1.Width = 1935: rButton5.Width = 1935
rButton4.BackColor = &HFFFFFF: rButton2.BackColor = &H8000000F: rButton3.BackColor = &H8000000F: rButton1.BackColor = &H8000000F: rButton5.BackColor = &H8000000F
Label9(1).Caption = ": " & lsDB.ListCount & " Virus"
Picture2.Visible = True: Picture12.Visible = False: Picture3.Visible = False: Picture7.Visible = False: Picture21.Visible = False
End Sub
Private Sub rButton5_Click()
rButton5.Width = 2055: rButton2.Width = 1935: rButton1.Width = 1935: rButton4.Width = 1935: rButton3.Width = 1935
rButton5.BackColor = &HFFFFFF: rButton2.BackColor = &H8000000F: rButton3.BackColor = &H8000000F: rButton4.BackColor = &H8000000F: rButton1.BackColor = &H8000000F
Picture21.Visible = True: Picture12.Visible = False: Picture7.Visible = False: Picture2.Visible = False: Picture3.Visible = False
End Sub
Private Sub SCANPROC_CurrentModule(Process As String, ID As Long, Module As String, File As String)
    Dim lsv As ListItem
    Set lsv = ListView1.ListItems.Add(, , Module)
    With lsv
        .SubItems(1) = File 'Ini Untuk Module
    End With
End Sub

Private Sub SCANPROC_CurrentProcess(Name As String, File As String, ID As Long, Modules As Long)
    Dim p_HasImage As Boolean
    If (File <> "SYSTEM") Then
        On Error Resume Next
        ImageList1.ListImages(Name).Tag = ""   'Nggak Tau Apa nihh XD
        If (err.Number <> 0) Then
            err.Clear
            ImageList1.ListImages.Add , Name, FILEICON.Icon(File, SmallIcon)
            p_HasImage = (err.Number = 0)
        Else
            p_HasImage = True
        End If
    End If
    Dim lsv As ListItem
    If (p_HasImage = True) Then
        Set lsv = ListView1.ListItems.Add(, "#" & Name & ID, Name, , Name)
    Else
        Set lsv = ListView1.ListItems.Add(, "#" & Name & ID, Name)
    End If
    With lsv
        .SubItems(1) = File
        .SubItems(2) = ID
        .SubItems(3) = Modules
'        .EnsureVisible
    End With
   
    If (m_sProcess <> "#" & Name & ID) Then
        Modules = 0
    End If
End Sub
Private Sub SCANPROC_DoneScanning(TotalProcess As Long)
    Dim p_Elapsed As Single
    p_Elapsed = Timer - m_sTime
    LockWindowUpdate 0& ' Fungsikan ListView Repaint
    'NUMPROC = TotalProcess
End Sub
Private Sub ListView1_ItemClick(ByVal Item As ComctlLib.ListItem)
    If (Check2.value <> vbChecked) Then
        'PopupMenu mnEx 'If Button = 2 Then PopupMenu mnEx
        Exit Sub ' What for?
    End If
    Dim i As Long
    i = Item.Index
    If (m_sProcess = Item.Key) Then
        m_sProcess = ""
    Else
        m_sProcess = Item.Key
    End If
    Call Check1_Click
    On Error Resume Next
    ListView1.ListItems(i).Selected = True
    ListView1.SelectedItem.EnsureVisible
End Sub
Private Sub Check1_Click()
    ListView1.ListItems.Clear
    SCANPROC.SystemProcesses = (Check1.value = vbChecked)
    SCANPROC.ProcessModules = (Check2.value = vbChecked)
    m_sTime = Timer
    LockWindowUpdate ListView1.hwnd ' Prevent listview repaints
    SCANPROC.BeginScanning
End Sub
Private Sub Check2_Click()
    If (Check2.value = vbChecked) Then
        Command2.Caption = "Refresh"
    Else
        Command2.Caption = "Refresh"
    End If
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyEscape) Then
        SCANPROC.CancelScanning
        Unload Me
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set SCANPROC = Nothing
    Set FILEICON = Nothing
End Sub
Private Sub SettingPelindung_Click(Index As Integer)
'Untuk RTP
If SettingPelindung(0).Checked = True Then
frmRTP.TimerRTP.Enabled = True: frmRTP.tmrControl.Enabled = True: frmMain.Image8.Picture = LoadResPicture(107, vbResBitmap) 'secure
Else
frmRTP.TimerRTP.Enabled = False: frmRTP.tmrControl.Enabled = False: frmMain.Image8.Picture = LoadResPicture(108, vbResBitmap) 'not secure
End If
'Untuk FireWall Windows
On Error Resume Next
Dim Firewall, Digital
If SettingPelindung(1).Checked = True Then
Set Firewall = CreateObject("HNetCfg.fwMgr")
Set Digital = Firewall.LocalPolicy.CurrentProfile
Digital.firewallenabled = True
Else
Set Firewall = CreateObject("HNetCfg.fwMgr")
Set Digital = Firewall.LocalPolicy.CurrentProfile
Digital.firewallenabled = False
End If
End Sub
Private Sub SettingUser_Click(Index As Integer)
Dim REG1
'Untuk Transparant
If SettingUser(0).Checked = True Then: SetTrans Me, 127: Else: SetTrans Me, 254
'Untuk Startup
On Error Resume Next
Dim s
If SettingUser(1).Checked = True Then
s = Replace(App.path & "\" & App.EXEName & ".exe", "\\", "\")
Set REG1 = CreateObject("WScript.Shell")
REG1.RegWrite "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\" & App.EXEName, s
Else
Set REG1 = CreateObject("WScript.Shell")
REG1.RegDelete "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\" & App.EXEName
End If
'Untuk Always On Top
If SettingUser(3).Checked = True Then: FormOnTop Me, True: Else: FormOnTop Me, False
End Sub
Private Sub TabStrip1_Click()
Select Case TabStrip1.SelectedItem.Index
Case 1: Picture12.Visible = True: Picture7.Visible = False: Picture3.Visible = False: Picture2.Visible = False
Case 2: Picture7.Visible = True: Picture12.Visible = False: Picture3.Visible = False: Picture2.Visible = False
Case 3: Picture3.Visible = True: Picture12.Visible = False: Picture7.Visible = False: Picture2.Visible = False
Case 4: Label9(1).Caption = ": " & lsDB.ListCount & " Virus": Picture2.Visible = True: Picture12.Visible = False: Picture3.Visible = False: Picture7.Visible = False
End Select
End Sub
Private Sub TabStrip2_Click()
Select Case TabStrip2.SelectedItem.Index
Case 1: Label8(4).Caption = "Tambahkan Virus Dari User"
Picture6.Visible = True: Picture5.Visible = False: Picture4.Visible = False: Picture19.Visible = False
Case 2: Label8(4).Caption = "Startup Kontrol"
Picture19.Visible = True: Picture6.Visible = False: Picture4.Visible = False: Picture5.Visible = False
Case 3: Label8(4).Caption = "Proses Kontrol"
Picture5.Visible = True: Picture6.Visible = False: Picture4.Visible = False: Picture19.Visible = False
Case 4: Check3.Visible = True 'Untuk Cek Quarantine
Label8(4).Caption = "List Karantina": Picture4.Visible = True: Picture5.Visible = False: Picture6.Visible = False: Picture19.Visible = False
End Select
End Sub
Private Sub TabStrip3_Click()
Select Case TabStrip3.SelectedItem.Index
Case 1: Picture20(7).Visible = False: Picture20(8).Visible = False: Picture11.Visible = True: Picture10.Visible = False: Picture9.Visible = False: Picture8.Visible = False
Case 2: Picture20(7).Visible = False: Picture20(8).Visible = False: Picture10.Visible = True: Picture11.Visible = False: Picture9.Visible = False: Picture8.Visible = False
Case 3: Picture20(7).Visible = True: Picture20(8).Visible = False: Picture9.Visible = True: Picture10.Visible = False: Picture11.Visible = False: Picture8.Visible = False
Case 4: Picture8.Visible = True: Picture10.Visible = False: Picture9.Visible = False: Picture11.Visible = False: Picture20(7).Visible = False: Picture20(8).Visible = True
End Select
End Sub
Function ListView32()
ListviewCheck lvVirus: ListviewCheck frmRTP.lvVirus: ListviewCheck lvReg: ListviewCheck frmEks.lvProg: ListviewFlat ListView1: ListviewFlat ListView2: ListviewFlat ListView3: ListviewFlat LVV
End Function
Private Sub TabStrip4_Click()
Select Case TabStrip4.SelectedItem.Index
Case 1: Label8(11).Caption = "Konfigurasi User"
Picture23.Visible = True: Picture22.Visible = False: Picture31.Visible = False
Case 2: Label8(11).Caption = "Konfigurasi Pelindung"
Picture22.Visible = True: Picture23.Visible = False: Picture31.Visible = False
Case 3: Label8(11).Caption = "Lewati File Aman"
Picture31.Visible = True: Picture23.Visible = False: Picture22.Visible = False
End Select
End Sub

Private Sub Timer1_Timer()
Call DeteksiUSBSekarang
End Sub
Private Sub tmrInformasi_Timer()
Label7.Caption = ": " & lvVirus.ListItems.Count
Label16.Caption = ": " & lvReg.ListItems.Count
End Sub
Private Sub tmrWaktu_Timer()
detik.i = detik.i + 1
If detik.i > 59 Then
    menit.i = menit.i + 1
    detik.i = 0
End If

If menit.i > 59 Then
    jam.i = jam.i + 1
    menit.i = 0
End If

detik.s = detik.i
menit.s = menit.i
jam.s = jam.i

If Len(detik.s) = 1 Then
    detik.s = "0" & detik.s
End If

If Len(menit.s) = 1 Then
    menit.s = "0" & menit.s
End If

If Len(jam.s) = 1 Then
    jam.s = "0" & jam.s
End If
Label23.Caption = ": " & jam.s & ":" & menit.s & ":" & detik.s
End Sub
