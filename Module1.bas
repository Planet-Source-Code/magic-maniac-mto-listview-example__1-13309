Attribute VB_Name = "Module1"
Function SplitPath$(ByVal Path$, ByVal ReturnType$)
  Dim Drv$
  Dim DirPath$
  Dim File$
  Dim Ext$
  Dim PathLength%
  Dim Offset%
  Dim ThisLength%
  If Len(Path) < 2 Then
    SplitPath = ""
    Exit Function
  End If
  If Mid(Path, 2, 1) = ":" Then
    Drv = Left(Path, 2)
    Path = Right(Path, Len(Path) - 2)
  Else
    Path = "__\" & Path
  End If
  PathLength = Len(Path)
  For Offset = PathLength To 1 Step -1
    Select Case Mid(Path, Offset, 1)
      Case "."
        If Ext = "" Then
          ThisLength = Len(Path) - Offset
          Ext = Right(Path, ThisLength)
          Path = Left(Path, Offset - 1)
        End If
      Case "\", "/"
        ThisLength = Len(Path) - Offset
        If ThisLength >= 1 Then
          File = Right(Path, ThisLength)
          Path = Left(Path, Offset)
          DirPath = Path
          Exit For
        End If
      Case Else
    End Select
  Next Offset
  SplitPath = Drv & Path & File & "." & Ext
  Select Case UCase$(ReturnType)
    Case "DRIVE": SplitPath = Drv
    Case "PATH", "DIR": SplitPath = Path
    Case "FILE", "NAME": SplitPath = File
    Case "EXT": SplitPath = Ext
    Case "DRIVE+PATH", "DRIVE+DIR": SplitPath = Drv & Path
    Case "FILE+EXT", "NAME+EXT": SplitPath = File & "." & Ext
    Case "DRIVE+PATH+FILE", "DRIVE+PATH+NAME", "DRIVE+DIR+FILE", "DRIVE+DIR+NAME": SplitPath = Drv & Path & File
  End Select
End Function

Public Function sAttr(Attr As VbFileAttribute) As String
  Dim sStr1 As String
  sStr1 = ""
  If Attr And vbReadOnly Then sStr1 = "r" Else sStr1 = "-"
  If Attr And vbArchive Then sStr1 = sStr1 + "a" Else sStr1 = sStr1 + "-"
  If Attr And vbHidden Then sStr1 = sStr1 + "h" Else sStr1 = sStr1 + "-"
  If Attr And vbSystem Then sStr1 = sStr1 + "s" Else sStr1 = sStr1 + "-"
  sAttr = sStr1
End Function

Public Sub SetColumnHeader(LV As ListView, ImgLstHeader As ImageList)
  Dim LvIndex As Long
  On Error Resume Next
    LvIndex = LV.Index
    If Err Then
      LvIndex = 0
    End If
  On Error GoTo 0
  LV.Sorted = False
  LV.ColumnHeaderIcons = ImgLstHeader
  LV.SortKey = GetSetting(App.CompanyName, App.Title, "SortKey" & LV.Name & LvIndex, 0)
  LV.SortOrder = GetSetting(App.CompanyName, App.Title, "SortOrder" & LV.Name & LvIndex, lvwAscending)
  LV.ColumnHeaders.Item(LV.SortKey + 1).Icon = LV.SortOrder + 1
  LV.Sorted = True
End Sub

Public Sub ColumnHeaderClick(LV As ListView, ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  Dim LvIndex As Long
  On Error Resume Next
    LvIndex = LV.Index
    If Err Then
      LvIndex = 0
    End If
  On Error GoTo 0
  LV.ColumnHeaders.Item(LV.SortKey + 1).Icon = 0
  If LV.SortKey = ColumnHeader.Index - 1 Then
    If LV.SortOrder = lvwAscending Then
      LV.SortOrder = lvwDescending
    Else
      LV.SortOrder = lvwAscending
    End If
    SaveSetting App.CompanyName, App.Title, "SortOrder" & LV.Name & LvIndex, LV.SortOrder
  Else
    LV.SortKey = ColumnHeader.Index - 1
    SaveSetting App.CompanyName, App.Title, "SortKey" & LV.Name & LvIndex, LV.SortKey
  End If
  LV.ColumnHeaders.Item(LV.SortKey + 1).Icon = LV.SortOrder + 1
  LV.SetFocus
  DoEvents
  If LV.ListItems.Count > 0 Then
    LV.ListItems(LV.SelectedItem.Index).EnsureVisible
  End If
End Sub

