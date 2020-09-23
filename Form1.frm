VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7230
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   7230
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1215
      Top             =   3285
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   45
      Top             =   3285
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   8
      ImageHeight     =   8
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":014A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":01EE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      Caption         =   "Options"
      Height          =   3120
      Left            =   5805
      TabIndex        =   3
      Top             =   45
      Width           =   1365
      Begin VB.CommandButton CmdEdit 
         Caption         =   "&Edit"
         Height          =   330
         Left            =   90
         TabIndex        =   7
         Top             =   675
         Width           =   1185
      End
      Begin VB.CommandButton CmdRemove 
         Caption         =   "&Remove"
         Height          =   330
         Left            =   90
         TabIndex        =   6
         Top             =   1080
         Width           =   1185
      End
      Begin VB.CommandButton CmdAdd 
         Caption         =   "&Add"
         Height          =   330
         Left            =   90
         TabIndex        =   5
         Top             =   270
         Width           =   1185
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "E&xit"
         Height          =   330
         Left            =   90
         TabIndex        =   4
         Top             =   1485
         Width           =   1185
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Files"
      Height          =   3120
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   5685
      Begin MSComctlLib.ListView ListView1 
         Height          =   2580
         Left            =   90
         TabIndex        =   1
         Top             =   225
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   4551
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Total: 0/0 Files"
         Height          =   240
         Left            =   90
         TabIndex        =   2
         Top             =   2835
         Width           =   5460
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   675
      Top             =   3285
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'# -------------------------------------
'# LISTVIEW EXAMPLE
'#
'# Version 1.1 ( File Collection )
'# -------------------------------------
'# CODED BY:
'#
'# MAGiC MANiAC^mTo ( mto@kabelfoon.nl )
'#
'# MORTAL OBSESSiON
'# http://home.kabelfoon.nl/~mto
'# -------------------------------------
'# RELEASED 04-Dec-2000 ON:
'#
'# www.planet-source-code.com
'# -------------------------------------

Dim lItem As ListItem

Sub Form_Load()
  Me.Caption = App.Title

  With Me.ListView1
    .MousePointer = ccDefault
    .View = lvwReport
    .Arrange = lvwNone
    .LabelEdit = lvwManual
    .BorderStyle = ccFixedSingle
    .Appearance = cc3D
    .OLEDragMode = ccOLEDragManual
    .OLEDropMode = ccOLEDropNone
    
    .HideColumnHeaders = False
    .HideSelection = False
    .LabelWrap = False
    .MultiSelect = False
    .Enabled = True
    .AllowColumnReorder = True
    .Checkboxes = False
    .FlatScrollBar = False
    .FullRowSelect = True
    .GridLines = True
    .HotTracking = False
    .HoverSelection = False
    
    .Sorted = True
    .SortKey = 0
    .SortOrder = lvwAscending
    
    .ColumnHeaders.Add , , "File", 1500
    .ColumnHeaders.Add , , "Size", 1000, lvwColumnRight
    .ColumnHeaders.Add , , "Date & Time", 1700
    .ColumnHeaders.Add , , "Attr", 700
    .ColumnHeaders.Add , , "Dir", 10000
    
    SetColumnHeader Me.ListView1, Me.ImageList1
      
    Set lItem = .ListItems.Add(, , "autoexec.bat")
      lItem.ListSubItems.Add , , Format(FileLen("c:\autoexec.bat"), "###,###,##0")
      lItem.ListSubItems.Add , , Format(FileDateTime("c:\autoexec.bat"), "DD-MM-YYYY HH:MM:SS")
      lItem.ListSubItems.Add , , sAttr(GetAttr("c:\autoexec.bat"))
      lItem.ListSubItems.Add , , "c:\"
    
    Set lItem = .ListItems.Add(, , "command.com")
      lItem.ListSubItems.Add , , Format(FileLen("c:\command.com"), "###,###,##0")
      lItem.ListSubItems.Add , , Format(FileDateTime("c:\command.com"), "DD-MM-YYYY HH:MM:SS")
      lItem.ListSubItems.Add , , sAttr(GetAttr("c:\command.com"))
      lItem.ListSubItems.Add , , "c:\"
    
    Set lItem = .ListItems.Add(, , "config.sys")
      lItem.ListSubItems.Add , , Format(FileLen("c:\config.sys"), "###,###,##0")
      lItem.ListSubItems.Add , , Format(FileDateTime("c:\config.sys"), "DD-MM-YYYY HH:MM:SS")
      lItem.ListSubItems.Add , , sAttr(GetAttr("c:\config.sys"))
      lItem.ListSubItems.Add , , "c:\"
  End With
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  ColumnHeaderClick Me.ListView1, ColumnHeader
End Sub

Private Sub ListView1_DblClick()
  CmdEdit_Click
End Sub

Private Sub CmdAdd_Click()
  Dim lTmp1 As Long
  Dim bFound As Boolean
  Dim sFile As String
  Dim sDir As String
  Me.CommonDialog1.InitDir = App.Path
  Me.CommonDialog1.Filter = "All Files (*.*)|*.*|Picture Files (*.bmp;*.gif;*.jpg)|*.BMP;*.GIF;*.JPG|HTML Files (*.htm;*.html;*.mht)|*.htm;*.html;*.mht|Text Files (*.txt;*.doc)|*.txt;*.doc|Sound Files (*.wav;*.mp3)|*.wav;*.mp3"
  Me.CommonDialog1.CancelError = True
  On Error Resume Next
    Me.CommonDialog1.ShowOpen
    If Err Then
      GoTo CommonDialog1Canceled
    End If
  On Error GoTo 0
  sFile = SplitPath(CommonDialog1.FileName, "FILE+EXT")
  sDir = SplitPath(CommonDialog1.FileName, "DRIVE+DIR")
  With Me.ListView1
    bFound = False
    For lTmp1 = 1 To .ListItems.Count
      If CommonDialog1.FileName = .ListItems(lTmp1).SubItems(4) + .ListItems(lTmp1).Text Then
        bFound = True
        Exit For
      End If
    Next
    If bFound Then
      .ListItems(lTmp1).Selected = True
      MsgBox "You can't add the new file because the file is already exist in the listview!", , App.Title + " - Add"
    Else
      Set lItem = .ListItems.Add(, , sFile)
        lItem.ListSubItems.Add , , Format(FileLen(CommonDialog1.FileName), "###,###,##0")
        lItem.ListSubItems.Add , , Format(FileDateTime(CommonDialog1.FileName), "DD-MM-YYYY HH:MM:SS")
        lItem.ListSubItems.Add , , sAttr(GetAttr(CommonDialog1.FileName))
        lItem.ListSubItems.Add , , sDir
        lItem.Selected = True
    End If
    .SetFocus
  End With
  Exit Sub
CommonDialog1Canceled:
  MsgBox "CommonDialog1 is canceled by user!"
End Sub

Private Sub CmdEdit_Click()
  Dim lTmp1 As Long
  Dim bFound As Boolean
  Dim sFile As String
  Dim sDir As String
  Me.CommonDialog1.InitDir = App.Path
  Me.CommonDialog1.Filter = "All Files (*.*)|*.*|Picture Files (*.bmp;*.gif;*.jpg)|*.BMP;*.GIF;*.JPG|HTML Files (*.htm;*.html;*.mht)|*.htm;*.html;*.mht|Text Files (*.txt;*.doc)|*.txt;*.doc|Sound Files (*.wav;*.mp3)|*.wav;*.mp3"
  Me.CommonDialog1.CancelError = True
  On Error Resume Next
    Me.CommonDialog1.ShowOpen
    If Err Then
      GoTo CommonDialog1Canceled
    End If
  On Error GoTo 0
  sFile = SplitPath(CommonDialog1.FileName, "FILE+EXT")
  sDir = SplitPath(CommonDialog1.FileName, "DRIVE+DIR")
  With Me.ListView1
    bFound = False
    For lTmp1 = 1 To .ListItems.Count
      If lTmp1 <> .SelectedItem.Index And CommonDialog1.FileName = .ListItems(lTmp1).SubItems(4) + .ListItems(lTmp1).Text Then
        bFound = True
        Exit For
      End If
    Next
    If bFound Then
      MsgBox "You can't change the file because the file is already exist in the listview!", , App.Title + " - Edit"
    Else
      .ListItems(.SelectedItem.Index).Text = sFile
      .ListItems(.SelectedItem.Index).ListSubItems(1) = Format(FileLen(CommonDialog1.FileName), "###,###,##0")
      .ListItems(.SelectedItem.Index).ListSubItems(2) = Format(FileDateTime(CommonDialog1.FileName), "DD-MM-YYYY HH:MM:SS")
      .ListItems(.SelectedItem.Index).ListSubItems(3) = sAttr(GetAttr(CommonDialog1.FileName))
      .ListItems(.SelectedItem.Index).ListSubItems(4) = sDir
      .ListItems(.SelectedItem.Index).Selected = True
    End If
    .SetFocus
  End With
  Exit Sub
CommonDialog1Canceled:
  MsgBox "CommonDialog1 is canceled by user!"
End Sub

Private Sub CmdRemove_Click()
  With Me.ListView1
    If .ListItems.Count > 0 Then
      .ListItems.Remove .SelectedItem.Index
    End If
    If .ListItems.Count > 0 Then
      .ListItems(.SelectedItem.Index).Selected = True
    End If
    .SetFocus
  End With
End Sub

Private Sub CmdExit_Click()
  Unload Me
End Sub

Private Sub Timer1_Timer()
  Dim lTmp1 As Long
  With Me.ListView1
    Me.CmdEdit.Enabled = .ListItems.Count > 0
    Me.CmdRemove.Enabled = .ListItems.Count > 0
    lTmp1 = 0
    If .ListItems.Count > 0 Then
      lTmp1 = .SelectedItem.Index
    End If
    Me.Label1.Caption = "Total: " & lTmp1 & "/" & .ListItems.Count & " Files"
  End With
End Sub
