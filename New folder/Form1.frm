VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   2145
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Interval        =   13000
      Left            =   840
      Top             =   480
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   240
      Top             =   480
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1560
      TabIndex        =   0
      Top             =   1320
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Dim Naik As Boolean


Private Sub Form_Load()

    Top = ((GetSystemMetrics(17) + GetSystemMetrics(4)) * Screen.TwipsPerPixelY)
    Left = (GetSystemMetrics(16) * Screen.TwipsPerPixelX) - Width
    Naik = True
    
End Sub

Private Sub Command1_Click()
    Naik = False
    Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
    Const s = 150 'kecepatan gerak / slide
    Dim v As Single
    v = (GetSystemMetrics(17) + GetSystemMetrics(4)) * Screen.TwipsPerPixelY
    
    If Naik = True Then
        If Top - s <= v - Height Then
            Top = Top - (Top - (v - Height))
            Timer1.Enabled = False
        Else
            Top = Top - s
        End If
        
    Else
        Top = Top + s
        If Top >= v Then Unload Me
    End If
End Sub

Private Sub Timer2_Timer()
    Naik = False
    Timer1.Enabled = True
End Sub
