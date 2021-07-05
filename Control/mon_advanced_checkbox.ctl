VERSION 5.00
Begin VB.UserControl mm_checkbox 
   AutoRedraw      =   -1  'True
   ClientHeight    =   1380
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1710
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   ScaleHeight     =   92
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   114
   ToolboxBitmap   =   "mon_advanced_checkbox.ctx":0000
   Begin VB.PictureBox pic4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   900
      Picture         =   "mon_advanced_checkbox.ctx":0312
      ScaleHeight     =   435
      ScaleWidth      =   720
      TabIndex        =   3
      Top             =   840
      Width           =   720
   End
   Begin VB.PictureBox pic3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   855
      Picture         =   "mon_advanced_checkbox.ctx":13A4
      ScaleHeight     =   435
      ScaleWidth      =   720
      TabIndex        =   2
      Top             =   255
      Width           =   720
   End
   Begin VB.PictureBox pic2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   150
      Picture         =   "mon_advanced_checkbox.ctx":2436
      ScaleHeight     =   210
      ScaleWidth      =   345
      TabIndex        =   1
      Top             =   780
      Width           =   345
   End
   Begin VB.PictureBox pic1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   150
      Picture         =   "mon_advanced_checkbox.ctx":2868
      ScaleHeight     =   210
      ScaleWidth      =   345
      TabIndex        =   0
      Top             =   465
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000C0&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H008080FF&
      FillStyle       =   4  'Upward Diagonal
      Height          =   405
      Left            =   75
      Shape           =   4  'Rounded Rectangle
      Top             =   60
      Visible         =   0   'False
      Width           =   285
   End
End
Attribute VB_Name = "mm_checkbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit

'EVENTS.
Public Event Click()
Public Event DoubleClick()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseEnters(ByVal X As Long, ByVal Y As Long)
Public Event MouseLeaves(ByVal X As Long, ByVal Y As Long)


Private udtPoint As POINTAPI
Private bolMouseDown As Boolean
Private bolMouseOver As Boolean
'Private bolHasFocus As Boolean
Private bolEnabled As Boolean
Private bolChecked As Boolean
Private bolSmall As Boolean
Private lonRoundValue As Long 'Rounded corners value.
Private lonRect As Long
Private button_clique As Integer

Private Type POINTAPI
        X As Long
        Y As Long
End Type
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function RoundRect Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Function PointInControl(X As Single, Y As Single) As Boolean
  If X >= 0 And X <= UserControl.ScaleWidth And _
    Y >= 0 And Y <= UserControl.ScaleHeight Then
    PointInControl = True
  End If
End Function

Private Sub PaintControl()
    On Error Resume Next

    Dim Round1 As Integer
    
    pic1.Visible = False
    pic2.Visible = False
    pic3.Visible = False
    pic4.Visible = False
    
    
    If Small Then
'        Round1 = 10
        UserControl.Width = (pic1.Width + 1) * Screen.TwipsPerPixelX
        UserControl.Height = (pic1.Height + 1) * Screen.TwipsPerPixelY
        If Checked Then UserControl.Picture = pic1.Picture Else UserControl.Picture = pic2.Picture
    Else
'        Round1 = 26
        UserControl.Width = (pic3.Width + 1) * Screen.TwipsPerPixelX
        UserControl.Height = (pic3.Height + 1) * Screen.TwipsPerPixelY
        If Checked Then UserControl.Picture = pic3.Picture Else UserControl.Picture = pic4.Picture
    End If
    
End Sub

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Button Enabled/Disable."
Enabled = bolEnabled
End Property

Public Property Get Small() As Boolean
Small = bolSmall
End Property
Public Property Get Checked() As Boolean
Checked = bolChecked
End Property

Public Property Let Enabled(ByVal NewValue As Boolean)
bolEnabled = NewValue
PropertyChanged "Enabled"


If bolSmall Then
    Shape1.Width = UserControl.Width / Screen.TwipsPerPixelX - 5: Shape1.Height = 4 '7
Else
    Shape1.Width = UserControl.Width / Screen.TwipsPerPixelX - 10: Shape1.Height = 8 '14
End If

If bolEnabled Then
    Shape1.Visible = False
Else
    'Shape1.Left = UserControl.Width / 2
    'Shape1.Top = UserControl.Height / 2
    Shape1.Left = ((UserControl.Width / 2) - (Shape1.Width * Screen.TwipsPerPixelX / 2)) / Screen.TwipsPerPixelX
    Shape1.Top = ((UserControl.Height / 2) - (Shape1.Height * Screen.TwipsPerPixelY / 2)) / Screen.TwipsPerPixelY
    Shape1.Visible = True
End If

UserControl.Enabled = bolEnabled

'PaintControl
End Property

Public Property Let Small(ByVal NewValue As Boolean)
bolSmall = NewValue
PropertyChanged "Small"

PaintControl

If Small = True Then
    RoundedValue = 10
Else
    RoundedValue = 26
End If


End Property
Public Property Let Checked(ByVal NewValue As Boolean)
bolChecked = NewValue
PropertyChanged "Checked"
PaintControl
End Property
Public Property Get RoundedValue() As Long
Attribute RoundedValue.VB_Description = "Button Border Rounded Value."
RoundedValue = lonRoundValue
End Property

Public Property Let RoundedValue(ByVal NewValue As Long)
lonRoundValue = NewValue
PropertyChanged "RoundedValue"
'PaintControl

lonRect = CreateRoundRectRgn(0, 0, ScaleWidth, ScaleHeight, lonRoundValue, lonRoundValue)     '- 1
SetWindowRgn UserControl.hWnd, lonRect, True

End Property

Private Sub pic1_Click()
UserControl_Click
End Sub

Private Sub pic1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call UserControl_MouseDown(Button, Shift, X, Y)
End Sub
Private Sub pic2_Click()
UserControl_Click
End Sub
Private Sub pic2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call UserControl_MouseDown(Button, Shift, X, Y)
End Sub
Private Sub pic3_Click()
UserControl_Click
End Sub
Private Sub pic3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call UserControl_MouseDown(Button, Shift, X, Y)
End Sub
Private Sub pic4_Click()
UserControl_Click
End Sub
Private Sub pic4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call UserControl_MouseDown(Button, Shift, X, Y)
End Sub
Private Sub UserControl_Click()
If bolEnabled = True Then
    If button_clique = 1 Then
        
        Checked = Not Checked
        PaintControl
        
        RaiseEvent Click
        RaiseEvent MouseLeaves(0, 0)
    End If
End If
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If bolEnabled = True Then
    button_clique = Button
    If Button = 1 Then
        bolMouseDown = True
        RaiseEvent MouseDown(Button, Shift, X, Y)
'        PaintControl
    End If
End If

End Sub
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If bolEnabled = False Then Exit Sub
    RaiseEvent MouseMove(Button, Shift, X, Y)
    SetCapture hWnd
    If PointInControl(X, Y) Then
        'pointer on control
        If Not bolMouseOver Then
            bolMouseOver = True
            RaiseEvent MouseEnters(udtPoint.X, udtPoint.Y)
        End If
    Else
        'pointer out of control
        bolMouseOver = False
        bolMouseDown = False
        ReleaseCapture
        RaiseEvent MouseLeaves(udtPoint.X, udtPoint.Y)
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If bolEnabled = True Then
    button_clique = Button
    If Button = 1 Then
        RaiseEvent MouseUp(Button, Shift, X, Y)
        bolMouseDown = False
    End If
End If
End Sub

Private Sub UserControl_Paint()
PaintControl
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
'On Error Resume Next
With PropBag
    
    Let Enabled = .ReadProperty("Enabled", True)
    Let Checked = .ReadProperty("Checked", False)
    Let Small = .ReadProperty("Small", True)
    Let RoundedValue = .ReadProperty("RoundedValue", 5)
End With
End Sub
Private Sub UserControl_Resize()
    PaintControl
End Sub
Private Sub UserControl_Terminate()
bolMouseDown = False
bolMouseOver = False
'bolHasFocus = False
'UserControl.Cls
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
'On Error Resume Next
With PropBag
    .WriteProperty "Enabled", bolEnabled, True
    .WriteProperty "Checked", bolChecked, False
    .WriteProperty "Small", bolSmall, True
    .WriteProperty "RoundedValue", lonRoundValue, 5
End With
End Sub
Private Sub UserControl_InitProperties()
Let Enabled = True
Let Checked = False
Let Small = True
Let RoundedValue = 10 '5
End Sub
