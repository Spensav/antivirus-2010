VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmEks 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4440
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7920
   LinkTopic       =   "Form1"
   ScaleHeight     =   4440
   ScaleWidth      =   7920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture20 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   7
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   8175
      TabIndex        =   2
      Top             =   0
      Width           =   8175
      Begin VB.PictureBox Picture24 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   7200
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   3
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
            MouseIcon       =   "frmEks.frx":0000
            MousePointer    =   99  'Custom
            TabIndex        =   4
            Top             =   0
            Width           =   735
         End
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Program Eksekusi"
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
         TabIndex        =   5
         Top             =   360
         Width           =   5055
      End
   End
   Begin VB.CommandButton Command7 
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
      Left            =   2040
      TabIndex        =   1
      Top             =   3840
      Width           =   1575
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Jalankan Program"
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
      TabIndex        =   0
      Top             =   3840
      Width           =   1695
   End
   Begin ComctlLib.ListView lvProg 
      Height          =   2655
      Left            =   240
      TabIndex        =   6
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
         Text            =   "Nama Program"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Ukuran[b]"
         Object.Width           =   1482
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Lokasi Program"
         Object.Width           =   7832
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   3735
      Left            =   0
      ScaleHeight     =   3735
      ScaleWidth      =   8055
      TabIndex        =   7
      Top             =   960
      Width           =   8055
   End
End
Attribute VB_Name = "frmEks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command7_Click()
Unload Me
End Sub
Private Sub Form_Load()
GradientPic Picture1, &H800000, &HFF0000, gmVertical
End Sub
Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then 'left click
ReleaseCapture
SendMessage Me.hwnd, &HA1, 2, 0
End If
End Sub
Private Sub Label81_Click()
Unload Me
End Sub

Private Sub Picture20_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then 'left click
ReleaseCapture
SendMessage Me.hwnd, &HA1, 2, 0
End If
End Sub
