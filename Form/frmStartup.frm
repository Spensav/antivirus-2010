VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmStartup 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3810
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   ScaleHeight     =   3810
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   3375
      Left            =   0
      ScaleHeight     =   3375
      ScaleWidth      =   6255
      TabIndex        =   0
      Top             =   0
      Width           =   6255
      Begin ComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   1320
         TabIndex        =   2
         Top             =   2400
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   3375
         Left            =   0
         ScaleHeight     =   3375
         ScaleWidth      =   975
         TabIndex        =   1
         Top             =   0
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmStartup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
LoadTray 1
End Sub
