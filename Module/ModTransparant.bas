Attribute VB_Name = "ModTransparant"
Private Declare Function GetLayeredWindowAttributes Lib "user32.dll" (ByVal hwnd As Long, ByRef crKey As Long, ByRef bAlpha As Byte, ByRef dwFlags As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal Color As Long, ByVal x As Byte, ByVal Alpha As Long) As Boolean
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long

Public Enum TransType
  LWA_OPAQUE = 0
  LWA_COLORKEY = 1
  LWA_ALPHA = 2
End Enum

Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000

Private Const zFormOrPictBoxStr = "Must pass in the name of either a Form or a PictureBox."

Public Function isTransparent(zForm As Form) As TransType
  On Local Error Resume Next
  Dim vTrans As Byte, Alpha As TransType, cKey As Long
  GetLayeredWindowAttributes zForm.hwnd, cKey, vTrans, Alpha
  If err Then
    isTransparent = -1
  Else
    isTransparent = Alpha
  End If
End Function

Public Function GetTrans(zForm As Form) As Long
  On Local Error Resume Next
  Dim vTrans As Byte, Alpha As TransType, cKey As Long
  GetLayeredWindowAttributes zForm.hwnd, cKey, vTrans, Alpha
  If Alpha = LWA_ALPHA Then
    GetTrans = vTrans
  ElseIf Alpha = LWA_COLORKEY Then
    GetTrans = cKey
  Else
    GetTrans = -1
  End If
  If err Then
    GetTrans = -1
  End If
End Function

Public Function FadeIn(zForm As Form, Optional ByVal Final As Byte = 255, Optional ByVal vStep As Single = 2) As Boolean
  On Local Error Resume Next
  Dim vTrans As Long, ZFE As Boolean, VarTmp As Single
  vTrans = isTransparent(zForm)
  If vTrans <> LWA_ALPHA Then SetTrans zForm, 0
  vTrans = GetTrans(zForm)
  If vTrans = -1 Then
    SetTrans zForm, 0
    vTrans = 0
  End If
  If vTrans > Final Then
    FadeIn = False
    Exit Function
  End If
  If zForm.Visible = False Then zForm.Show
  ZFE = zForm.Enabled
  If ZFE = True Then zForm.Enabled = False
  VarTmp = vTrans
  While VarTmp < Final
    DoEvents
    VarTmp = VarTmp + vStep
    If VarTmp > Final Then VarTmp = Final
    SetTrans zForm, CByte(VarTmp)
  Wend
  If ZFE = True Then zForm.Enabled = True
  If err Then
    FadeIn = False
  Else
    FadeIn = True
  End If
End Function

Public Function FadeOut(zForm As Form, Optional ByVal Final As Byte = 0, Optional ByVal vStep As Single = 2) As Boolean
  On Local Error Resume Next
  Dim vTrans As Long, ZFE As Boolean, VarTmp As Single
  vTrans = isTransparent(zForm)
  If vTrans <> LWA_ALPHA Then SetTrans zForm, 255
  vTrans = GetTrans(zForm)
  If vTrans = -1 Then
    SetTrans zForm, 255
    vTrans = 255
  End If
  If vTrans < Final Then
    FadeOut = False
    Exit Function
  End If
  If zForm.Visible = False Then zForm.Show
  ZFE = zForm.Enabled
  If ZFE = True Then zForm.Enabled = False
  VarTmp = vTrans
  While VarTmp > Final
    DoEvents
    VarTmp = VarTmp - vStep
    If VarTmp < Final Then VarTmp = Final
    SetTrans zForm, CByte(VarTmp)
  Wend
  If ZFE = True Then zForm.Enabled = True
  If Final = 0 Then zForm.Hide
  If err Then
    FadeOut = False
  Else
    FadeOut = True
  End If
End Function

Public Function SetTrans(zForm As Form, Optional ByVal vTrans As Byte = 127) As Boolean
  On Local Error Resume Next
  Dim Msg As Long
  Msg = GetWindowLong(zForm.hwnd, GWL_EXSTYLE)
  Msg = Msg Or WS_EX_LAYERED
  SetWindowLong zForm.hwnd, GWL_EXSTYLE, Msg
  SetLayeredWindowAttributes zForm.hwnd, 0, vTrans, LWA_ALPHA
  If err Then
    SetTrans = False
  Else
    SetTrans = True
  End If
End Function
Public Function SetTrans1(zForm As PictureBox, Optional ByVal vTrans As Byte = 127) As Boolean
  On Local Error Resume Next
  Dim Msg As Long
  Msg = GetWindowLong(zForm.hwnd, GWL_EXSTYLE)
  Msg = Msg Or WS_EX_LAYERED
  SetWindowLong zForm.hwnd, GWL_EXSTYLE, Msg
  SetLayeredWindowAttributes zForm.hwnd, 0, vTrans, LWA_ALPHA
  If err Then
    SetTrans1 = False
  Else
    SetTrans1 = True
  End If
End Function
Public Function MakeTransparent(hwnd As Long, Perc As Integer) As Long
    Dim Msg As Long
    On Error Resume Next
    If Perc < 0 Or Perc > 255 Then
        MakeTransparent = 1
    Else
        Msg = GetWindowLong(hwnd, GWL_EXSTYLE)
        Msg = Msg Or WS_EX_LAYERED
        SetWindowLong hwnd, GWL_EXSTYLE, Msg
        SetLayeredWindowAttributes hwnd, 0, Perc, LWA_ALPHA
        MakeTransparent = 0
    End If
    If err Then
        MakeTransparent = 2
    End If
End Function
Public Sub FadeForm(ByVal Frm As Form, ByVal Level As Byte)
On Error Resume Next
Dim Msg As Long

Msg = GetWindowLong(Frm.hwnd, -20) Or &H80000
SetWindowLong Frm.hwnd, -20, Msg
SetLayeredWindowAttributes Frm.hwnd, 0, Level, &H2
End Sub



