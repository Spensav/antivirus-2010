VERSION 5.00
Begin VB.Form frmOther 
   Caption         =   "Form2"
   ClientHeight    =   3030
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   4560
   LinkTopic       =   "Form2"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu Menu 
      Caption         =   "Menu"
      Begin VB.Menu mnSpensav 
         Caption         =   "Buka Tampilan Spensav"
      End
      Begin VB.Menu mnTentang 
         Caption         =   "Tentang Spensav"
      End
      Begin VB.Menu MnBatas 
         Caption         =   "-"
      End
      Begin VB.Menu mnExit 
         Caption         =   "Keluar Aplikasi"
      End
   End
End
Attribute VB_Name = "frmOther"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnExit_Click()
End
End Sub

Private Sub mnSpensav_Click()
frmMain.Show
End Sub
