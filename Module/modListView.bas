Attribute VB_Name = "modListView"
'##### mendapatkan file mana saja yang di pilih pada listview ########'
Public Function GetSelected(TheLV As ListView)
Dim Sel As String
For i = 1 To TheLV.ListItems.Count
If TheLV.ListItems.Item(i).Checked = True Then
Sel = Sel & "|" & TheLV.ListItems.Item(i).SubItems(1)
End If
Next
GetSelected = Sel
End Function

'########### Memilih semua file pada listview ##########'
Public Function SelectedAll(TheLV As ListView)
For i = 1 To TheLV.ListItems.Count
TheLV.ListItems.Item(i).Checked = True
Next
End Function

'########### tidak memilih apapun pada listview ##########'
Public Function SelectedNone(TheLV As ListView)
For i = 1 To TheLV.ListItems.Count
TheLV.ListItems.Item(i).Checked = False
Next
End Function

'######## mendapatkan index pada listview ########'
Public Function GetIndex(TheLV As ListView, Data As String) As Integer
For i = 1 To TheLV.ListItems.Count
If TheLV.ListItems.Item(i).SubItems(1) = Data Then
GetIndex = i
End If
Next
End Function

'######### tidak memilih file yang sudah dihapus/dikarantina ##########'
Public Function UnSelect(TheLV As ListView, Data As String)
For i = 1 To TheLV.ListItems.Count
If TheLV.ListItems.Item(i).SubItems(3) = Data Then
TheLV.ListItems.Item(i).Checked = False
End If
Next
End Function

'######### menambahkan file yang terdeteksi kedalam listview #######'
Public Function AddDetect(TheLV As ListView, FilePath As String, VirData As String)
With TheLV
If Left(VirData, 9) <> "Malicious" Then
Set lvItm = .ListItems.Add(, , Split(VirData, "|")(0), , frmMain.ImgSmall.ListImages(1).Index)
lvItm.SubItems(1) = FilePath
lvItm.SubItems(2) = Split(VirData, "|")(1)
lvItm.SubItems(3) = "Virus File"
Else
Set lvItm = .ListItems.Add(, , VirData, , frmMain.ImgSmall.ListImages(1).Index)
lvItm.SubItems(1) = FilePath
lvItm.SubItems(2) = GetChecksum(FilePath)
lvItm.SubItems(3) = "Virus File"
End If
End With
End Function
