Attribute VB_Name = "ModStartupMgr3"
'vnAntiVirus 1.0

'Author : Dung Le Nguyen
'Email : dungcoivb@gmail.com
'This is a software open source

Public tb As Boolean

Public ichkUSB As Boolean
Public ichkScanI As Boolean
Public ichkAutoIT As Boolean
Public ichkSam As Boolean
Public ichkDec As Boolean
Public PathDec As String

Public ioptVie As Boolean
Public ichkShow As Boolean
Public ichkSystemTray As Boolean

Public LoadMon As Boolean

Public PathWScan As String
Public SeeSta As Boolean

Public SPro As Boolean
Public SSta As Boolean
Public SAll As Boolean
Public Sub GetIcons(lstView As ListView, imaList As ImageList, picTmp As PictureBox)

        Dim lsv As ListItem
For Each lsv In lstView.ListItems
        picTmp.Cls
        GetIcon lsv.SubItems(1), picTmp
        imaList.ListImages.Add lsv.Index, , picTmp.Image
Next
    
With lstView
  .SmallIcons = imaList
  For Each lsv In .ListItems
    lsv.SmallIcon = lsv.Index
  Next
End With
End Sub
Public Sub ThietLap(lstView As ListView, imaList As ImageList, picTmp As PictureBox)
'On Local Error Resume Next
    picTmp.Cls
    picTmp.BackColor = vbWhite
    lstView.BackColor = vbWhite
    lstView.ListItems.Clear
    lstView.SmallIcons = Nothing
    imaList.ListImages.Clear
End Sub
Public Function FileExists(sFileName As String) As Boolean
    '##############################################################################################
    'Returns True if the specified file exists
    'Ham nay neu su dung de kiem tra file tren USB se lap tuc gay ra loi
    '##############################################################################################
    
    Dim WFD As WIN32_FIND_DATA
    Dim lResult As Long
    
    lResult = FindFirstFile(sFileName, WFD)
    If lResult <> INVALID_HANDLE_VALUE Then
        If (WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY Then
            FileExists = False
        Else
            FileExists = True
        End If
    End If
End Function



