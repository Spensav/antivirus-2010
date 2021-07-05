Attribute VB_Name = "ModGradient"
Enum GradMode
gmHorizontal = 0
gmVertical = 1
End Enum

Public Function GradientForm(ByVal Frm As Form, ByVal StartColor As Long, ByVal Endcolor As Long, ByVal Mode As GradMode)
Dim Rs As Integer, Gs As Integer, Bs As Integer
Dim Re As Integer, Ge As Integer, Be As Integer
Dim Rk As Single, Gk As Single, Bk As Single
Dim r As Integer, G As Integer, b As Integer
Dim i As Integer, j As Single

On Error Resume Next
Frm.AutoRedraw = True
Frm.ScaleMode = vbPixels

Rs = StartColor And (Not &HFFFFFF00)
Gs = (StartColor And (Not &HFFFF00FF)) \ &H100&
Bs = (StartColor And (Not &HFF00FFFF)) \ &HFFFF&
Re = Endcolor And (Not &HFFFFFF00)
Ge = (Endcolor And (Not &HFFFF00FF)) \ &H100&
Be = (Endcolor And (Not &HFF00FFFF)) \ &HFFFF&

j = IIf(Mode = gmHorizontal, Frm.ScaleWidth, Frm.ScaleHeight)
Rk = (Rs - Re) / j: Gk = (Gs - Ge) / j: Bk = (Bs - Be) / j

For i = 0 To j
r = Rs - i * Rk: G = Gs - i * Gk: b = Bs - i * Bk
If Mode = gmHorizontal Then
Frm.Line (i, 0)-(i - 1, Frm.ScaleHeight), RGB(r, G, b), B
Else
Frm.Line (0, i)-(Frm.ScaleWidth, i - 1), RGB(r, G, b), B
End If
Next
End Function
Public Function GradientPic(ByVal Frm As PictureBox, ByVal StartColor As Long, ByVal Endcolor As Long, ByVal Mode As GradMode)
Dim Rs As Integer, Gs As Integer, Bs As Integer
Dim Re As Integer, Ge As Integer, Be As Integer
Dim Rk As Single, Gk As Single, Bk As Single
Dim r As Integer, G As Integer, b As Integer
Dim i As Integer, j As Single

On Error Resume Next
Frm.AutoRedraw = True
Frm.ScaleMode = vbPixels

Rs = StartColor And (Not &HFFFFFF00)
Gs = (StartColor And (Not &HFFFF00FF)) \ &H100&
Bs = (StartColor And (Not &HFF00FFFF)) \ &HFFFF&
Re = Endcolor And (Not &HFFFFFF00)
Ge = (Endcolor And (Not &HFFFF00FF)) \ &H100&
Be = (Endcolor And (Not &HFF00FFFF)) \ &HFFFF&

j = IIf(Mode = gmHorizontal, Frm.ScaleWidth, Frm.ScaleHeight)
Rk = (Rs - Re) / j: Gk = (Gs - Ge) / j: Bk = (Bs - Be) / j

For i = 0 To j
r = Rs - i * Rk: G = Gs - i * Gk: b = Bs - i * Bk
If Mode = gmHorizontal Then
Frm.Line (i, 0)-(i - 1, Frm.ScaleHeight), RGB(r, G, b), B
Else
Frm.Line (0, i)-(Frm.ScaleWidth, i - 1), RGB(r, G, b), B
End If
Next
End Function
