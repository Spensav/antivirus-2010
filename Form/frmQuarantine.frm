VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmQuarantine 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4830
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   ScaleHeight     =   4830
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Scanner.XPFrame XPFrame1 
      Height          =   4815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   8493
      Caption         =   "Quarantine (0)"
      CaptionAlignment=   2
      Begin ComctlLib.ListView lvQ 
         Height          =   2055
         Left            =   1680
         TabIndex        =   5
         Top             =   1200
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   3625
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   327682
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin Scanner.jcbutton cmdRestore 
         Height          =   495
         Index           =   1
         Left            =   4440
         TabIndex        =   4
         Top             =   4200
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   873
         ButtonStyle     =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16765357
         Caption         =   "Restore To ..."
         UseMaskCOlor    =   -1  'True
      End
      Begin Scanner.jcbutton cmdRestore 
         Height          =   495
         Index           =   0
         Left            =   2280
         TabIndex        =   3
         Top             =   4200
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   873
         ButtonStyle     =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16765357
         Caption         =   "Restore"
         UseMaskCOlor    =   -1  'True
      End
      Begin Scanner.jcbutton cmdDelete 
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   4200
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   873
         ButtonStyle     =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16765357
         Caption         =   "Delete"
         UseMaskCOlor    =   -1  'True
      End
      Begin Scanner.jcbutton cmdBack 
         Height          =   255
         Left            =   6240
         TabIndex        =   1
         Top             =   60
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         ButtonStyle     =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16765357
         Caption         =   "X"
      End
   End
End
Attribute VB_Name = "frmQuarantine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Dat As String

'########## Update list file yang di karantina ##########'
Private Sub UpdateQ()
lvQ.ListItems.Clear
Dim Data() As String
If PathFileExists(Dat) = 0 Then Exit Sub
Open Dat For Input As #1
Input #1, isi
Close #1
Data = Split(isi, "|")
For i = 1 To UBound(Data)
With lvQ.ListItems.Add(, , Split(Data(i), "?")(0))
.SubItems(1) = Split(Data(i), "?")(2)
.SubItems(2) = Split(Data(i), "?")(1)
End With
Next
XPFrame1.Caption = "Quarantine (" & lvQ.ListItems.Count & ")"
End Sub

'########## kembali pada form utama ##########'
Private Sub cmdBack_Click()
frmQuarantine.Visible = False
frmMain.Enabled = True
End Sub

'########## menghapus file yang di karantina ##########'
Private Sub cmdDelete_Click()
If lvQ.ListItems.Count = 0 Then Exit Sub
Dim Data() As String
If PathFileExists(Dat) <> 0 Then
Open Dat For Input As #1
Input #1, isi
Close #1
DeleteFile Dat
Else
isi = ""
End If
Data = Split(isi, "|")
For i = 1 To UBound(Data)
namafile = lvQ.SelectedItem.SubItems(2)
If namafile <> Split(Data(i), "?")(1) Then
nyu = nyu & "|" & Data(i)
End If
Next
DeleteFile AppPath & "karantina\" & lvQ.SelectedItem.SubItems(1)
Open Dat For Output As #2
Print #2, nyu
Close #2
MsgBox "Success Deleting File !!!", vbInformation, ""
UpdateQ
End Sub

'########## melakukan restore pada folder asli/folder pilihan ##########'
Private Sub cmdRestore_Click(Index As Integer)
If lvQ.ListItems.Count = 0 Then Exit Sub
Select Case Index
Case 0
DecodeFile AppPath & "karantina\" & lvQ.SelectedItem.SubItems(1), lvQ.SelectedItem.SubItems(2)
MsgBox "File Restored to " & Chr(34) & lvQ.SelectedItem.SubItems(2) & Chr(34) & " !!!", vbInformation, ""
Case 1
sTitle = "Select path:" & vbNewLine & "Select path to restore file."
ThePath = BrowseFolder(sTitle, Me)
If ThePath <> "" Then
DecodeFile AppPath & "karantina\" & lvQ.SelectedItem.SubItems(1), ThePath & GetFileName(lvQ.SelectedItem.SubItems(2))
MsgBox "File Restored to " & Chr(34) & ThePath & GetFileName(lvQ.SelectedItem.SubItems(2)) & Chr(34) & " !!!", vbInformation, ""
End If
End Select
End Sub

'########## Meload daftar file yang di karantina
Private Sub Form_Load()
frmMain.Enabled = False
Dat = AppPath & "karantina\ASE-Q.dat"
UpdateQ
End Sub


