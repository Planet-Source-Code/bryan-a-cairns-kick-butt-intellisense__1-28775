VERSION 5.00
Object = "{665BF2B8-F41F-4EF4-A8D0-303FBFFC475E}#2.0#0"; "CMCS21.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Code Editor"
   ClientHeight    =   6450
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7155
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6450
   ScaleWidth      =   7155
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   2520
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   360
      Width           =   2415
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form1.frx":058A
      Left            =   0
      List            =   "Form1.frx":059A
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   360
      Width           =   2415
   End
   Begin MSComctlLib.ImageList IMGIntellisence 
      Left            =   5760
      Top             =   3600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":05C0
            Key             =   ""
            Object.Tag             =   "class"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0B5C
            Key             =   ""
            Object.Tag             =   "constant"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":10F8
            Key             =   ""
            Object.Tag             =   "function"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1694
            Key             =   ""
            Object.Tag             =   "sub2"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1C30
            Key             =   ""
            Object.Tag             =   "sub"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":21CC
            Key             =   ""
            Object.Tag             =   "property"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2328
            Key             =   ""
            Object.Tag             =   "method"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2484
            Key             =   ""
            Object.Tag             =   "variable"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2A20
            Key             =   ""
            Object.Tag             =   "map"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2B7C
            Key             =   ""
            Object.Tag             =   "control"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   6195
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Text            =   "Line:"
            TextSave        =   "Line:"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Text            =   "Col:"
            TextSave        =   "Col:"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6985
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   12
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "New"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Open"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Save"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Undo"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Redo"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Copy"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cut"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Paste"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Help"
            ImageIndex      =   9
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5280
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2CD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2E34
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2F90
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":30EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3248
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":33A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3500
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":365C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":37B8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin CodeSenseCtl.CodeSense RT 
      Height          =   4455
      Left            =   0
      OleObjectBlob   =   "Form1.frx":3914
      TabIndex        =   2
      Top             =   720
      Width           =   5055
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuNew 
         Caption         =   "New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save As"
      End
      Begin VB.Menu mnuFileSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuUndo 
         Caption         =   "Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuRedo 
         Caption         =   "Redo"
         Shortcut        =   ^Y
      End
      Begin VB.Menu mnuEditSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy"
      End
      Begin VB.Menu mnuCut 
         Caption         =   "Cut"
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "Paste"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete"
      End
      Begin VB.Menu mnuEditSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFind 
         Caption         =   "Find"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuFindNext 
         Caption         =   "Find Next"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuReplace 
         Caption         =   "Replace"
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuGoLine 
         Caption         =   "Go To Line"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuEditSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelAll 
         Caption         =   "Select All"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuSelLine 
         Caption         =   "Select Current Line"
      End
   End
   Begin VB.Menu mnuCode 
      Caption         =   "Code"
      Begin VB.Menu mnuCodes 
         Caption         =   ""
         Index           =   0
      End
   End
   Begin VB.Menu mnuBookMarks 
      Caption         =   "Bookmarks"
      Begin VB.Menu mnuBToggle 
         Caption         =   "Bookmark On / Off"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuBClearALL 
         Caption         =   "Clear All Bookmarks"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuBJumpFirst 
         Caption         =   "Jump to First Bookmark"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuBJumpLast 
         Caption         =   "Jump to Last Bookmark"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuBNext 
         Caption         =   "Go to Next Bookmark"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuBGoPrev 
         Caption         =   "Go to Previous Bookmark"
         Shortcut        =   {F9}
      End
   End
   Begin VB.Menu mnuHelpMe 
      Caption         =   "Help"
      Begin VB.Menu mnuHelp 
         Caption         =   "Help"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sLastWord As String
Dim sIntellText As String
Dim LBoxPos As Long
Public Sub EditorSetVals()
  'Use the color data
  RT.Language = "Basic"
  
  RT.SetColor cmClrBookmark, ClrData(0).frClr
  RT.SetColor cmClrBookmarkBk, ClrData(0).bgClr
  RT.SetColor cmClrCommentBk, ClrData(1).bgClr
  RT.SetColor cmClrComment, ClrData(1).frClr
  RT.SetColor cmClrHDividerLines, ClrData(2).frClr
  RT.SetColor cmClrVDividerLines, ClrData(3).frClr
  RT.SetColor cmClrHighlightedLine, ClrData(4).frClr
  RT.SetColor cmClrKeyword, ClrData(5).frClr
  RT.SetColor cmClrKeywordBk, ClrData(5).bgClr
  RT.SetColor cmClrLeftMargin, ClrData(6).frClr
  RT.SetColor cmClrLineNumber, ClrData(7).frClr
  RT.SetColor cmClrLineNumberBk, ClrData(7).bgClr
  RT.SetColor cmClrNumber, ClrData(8).frClr
  RT.SetColor cmClrNumberBk, ClrData(8).bgClr
  RT.SetColor cmClrOperator, ClrData(9).frClr
  RT.SetColor cmClrOperatorBk, ClrData(9).bgClr
  RT.SetColor cmClrScopeKeyword, ClrData(10).frClr
  RT.SetColor cmClrScopeKeywordBk, ClrData(10).bgClr
  RT.SetColor cmClrString, ClrData(11).frClr
  RT.SetColor cmClrStringBk, ClrData(11).bgClr
  RT.SetColor cmClrTagElementName, ClrData(12).frClr
  RT.SetColor cmClrTagElementNameBk, ClrData(12).bgClr
  RT.SetColor cmClrTagEntity, ClrData(13).frClr
  RT.SetColor cmClrTagEntityBk, ClrData(13).bgClr
  RT.SetColor cmClrTagAttributeName, ClrData(14).frClr
  RT.SetColor cmClrTagAttributeNameBk, ClrData(14).bgClr
  RT.SetColor cmClrTagText, ClrData(15).frClr
  RT.SetColor cmClrTagTextBk, ClrData(15).bgClr
  RT.SetColor cmClrText, ClrData(16).frClr
  RT.SetColor cmClrTextBk, ClrData(16).bgClr
  RT.SetColor cmClrWindow, ClrData(17).frClr
  
  'Setup font styles
  RT.SetFontStyle cmStyComment, txtProp(ClrData(1).fntProp)
  RT.SetFontStyle cmStyLineNumber, txtProp(ClrData(7).fntProp)
  RT.SetFontStyle cmStyNumber, txtProp(ClrData(8).fntProp)
  RT.SetFontStyle cmStyOperator, txtProp(ClrData(9).fntProp)
  RT.SetFontStyle cmStyScopeKeyword, txtProp(ClrData(10).fntProp)
  RT.SetFontStyle cmStyString, txtProp(ClrData(11).fntProp)
  RT.SetFontStyle cmStyTagAttributeName, txtProp(ClrData(12).fntProp)
  RT.SetFontStyle cmStyTagAttributeName, txtProp(ClrData(13).fntProp)
  RT.SetFontStyle cmStyTagEntity, txtProp(ClrData(14).bgClr)
  RT.SetFontStyle cmStyKeyword, txtProp(ClrData(5).fntProp)
  RT.SetFontStyle cmStyTagText, txtProp(ClrData(15).fntProp)
  RT.SetFontStyle cmStyNumber, txtProp(ClrData(8).fntProp)
  RT.SetFontStyle cmStyText, txtProp(ClrData(16).fntProp)
  
  'get the extra properties
  Dim iHG As Integer
iHG = CInt(GetSetting(App.EXEName, "EditOptions", "Highlight", "1"))
If iHG = 0 Then
RT.HighlightedLine = -1
End If
RT.LineNumbering = CBool(GetSetting(App.EXEName, "EditOptions", "linenumber", "1"))
RT.DisplayLeftMargin = CBool(GetSetting(App.EXEName, "EditOptions", "leftmargin", "1"))
RT.DisplayWhitespace = CBool(GetSetting(App.EXEName, "EditOptions", "whitespace", "0"))
RT.SmoothScrolling = CBool(GetSetting(App.EXEName, "EditOptions", "smoothscroll", "1"))
RT.LineNumberStart = 1
  'Set drag and drop
  RT.EnableDragDrop = True
  'Lets load some font data up :)
  RT.Font.Bold = False
  RT.Font.Italic = False
  RT.Font.Name = "Courier New"
  RT.Font.Size = 11.25
  RT.Font.Strikethrough = False
  RT.Font.Underline = False
  RT.ExpandTabs = True
  'RT.SetColor cmClrLineNumber, 16777215
  'RT.SetColor cmClrLineNumberBk, 8421504
  RT_SelChange RT

End Sub



Private Sub Combo1_Click()
LoadSubs Combo1.Text
End Sub

Private Sub Form_Load()
ResetAllEditVals
GetEditColors

EditorSetVals
StatusBar1.Panels(1).Text = "Line: 1"
StatusBar1.Panels(2).Text = "Col: 1"
Combo1.ListIndex = 0
RT.Text = ""
LoadCodeList
'for this demo
'remove if you want to use the program
RT.Text = "To see the Intellisense..." & vbCrLf & "Type in " & Chr(34) & "Device" & Chr(34) & " followed by a period " & Chr(34) & "." & Chr(34)
''''''''''''''''''''''''''''''''''''''
End Sub

Public Sub DoHighLight()
  On Error Resume Next
  Dim R As CodeSenseCtl.Range
  Set R = RT.GetSel(True)
  If CInt(GetSetting(App.EXEName, "EditOptions", "Highlight", "1")) = 1 Then
    RT.HighlightedLine = R.EndLineNo
  End If
End Sub

Private Sub Form_Resize()
On Error Resume Next
RT.Width = Me.Width - 120
RT.Height = (Me.Height - RT.Top) - 950
Combo2.Width = (Me.Width - Combo2.Left) - 120
End Sub

Private Sub List1_DblClick()
AddIntellWord
End Sub


Private Sub mnuBClearALL_Click()
On Error Resume Next
RT.DisplayLeftMargin = True
RT.ExecuteCmd cmCmdBookmarkClearAll
End Sub

Private Sub mnuBGoPrev_Click()
On Error Resume Next
RT.DisplayLeftMargin = True
RT.ExecuteCmd cmCmdBookmarkPrev
End Sub

Private Sub mnuBJumpFirst_Click()
On Error Resume Next
RT.DisplayLeftMargin = True
RT.ExecuteCmd cmCmdBookmarkJumpToFirst
End Sub

Private Sub mnuBJumpLast_Click()
On Error Resume Next
RT.DisplayLeftMargin = True
RT.ExecuteCmd cmCmdBookmarkJumpToLast
End Sub

Private Sub mnuBNext_Click()
On Error Resume Next
RT.DisplayLeftMargin = True
RT.ExecuteCmd cmCmdBookmarkNext
End Sub

Private Sub mnuBToggle_Click()
On Error Resume Next
RT.DisplayLeftMargin = True
RT.ExecuteCmd cmCmdBookmarkToggle
End Sub

Private Sub mnuCodes_Click(Index As Integer)
InsertCodes mnuCodes(Index).Caption
End Sub

Private Sub mnuCopy_Click()
  On Error Resume Next
  Clipboard.Clear
  Clipboard.SetText RT.SelText
End Sub

Private Sub mnuCut_Click()
  On Error Resume Next
  Clipboard.Clear
  Clipboard.SetText RT.SelText
  RT.SelText = ""
End Sub

Private Sub mnuDelete_Click()
  On Error Resume Next
  RT.SelText = ""
End Sub

Private Sub mnuFind_Click()
On Error Resume Next
RT.ExecuteCmd cmCmdFind
End Sub

Private Sub mnuFindNext_Click()
On Error Resume Next
RT.ExecuteCmd cmCmdFindNext
End Sub

Private Sub mnuGoLine_Click()
On Error Resume Next
RT.ExecuteCmd cmCmdGotoLine, -1
End Sub

Private Sub mnuHelp_Click()
'
End Sub

Private Sub mnuNew_Click()
'
End Sub

Private Sub mnuOpen_Click()
'
End Sub

Private Sub mnuPaste_Click()
  On Error Resume Next
  RT.Paste
End Sub

Private Sub mnuRedo_Click()
On Error Resume Next
RT.Redo
End Sub

Private Sub mnuReplace_Click()
On Error Resume Next
RT.ExecuteCmd cmCmdFindReplace
End Sub

Private Sub mnuSave_Click()
'
End Sub

Private Sub mnuSelAll_Click()
On Error Resume Next
RT.ExecuteCmd cmCmdSelectAll
End Sub

Private Sub mnuSelLine_Click()
On Error Resume Next
RT.ExecuteCmd cmCmdSelectLine
End Sub

Private Sub mnuUndo_Click()
On Error Resume Next
RT.Undo
End Sub

Private Sub RT_Change(ByVal Control As CodeSenseCtl.ICodeSense)
'
End Sub

Private Function RT_KeyPress(ByVal Control As CodeSenseCtl.ICodeSense, ByVal KeyAscii As Long, ByVal Shift As Long) As Boolean
GetRange

End Function

Private Function RT_KeyUp(ByVal Control As CodeSenseCtl.ICodeSense, ByVal KeyCode As Long, ByVal Shift As Long) As Boolean
'= is 187
'. is 190

Dim R As CodeSenseCtl.Range

If KeyCode = 9 Or KeyCode = 13 Then
    AddIntellWord
End If
If RT.CurrentWord <> "." Then
sLastWord = RT.CurrentWord
End If
If KeyCode = 190 Then

Set R = RT.GetSel(False)
LBoxPos = R.EndColNo
RT.ExecuteCmd cmCmdCodeList
End If
End Function
Private Function RT_CodeList(ByVal Control As CodeSenseCtl.ICodeSense, ByVal ListCtrl As CodeSenseCtl.ICodeList) As Boolean
'Me.Caption = sLastWord
ListCtrl.hImageList = IMGIntellisence.hImageList
RT_CodeList = LoadIntellList(sLastWord, ListCtrl)
End Function

Private Function RT_CodeListCancel(ByVal Control As CodeSenseCtl.ICodeSense, ByVal ListCtrl As CodeSenseCtl.ICodeList) As Boolean
'MsgBox "Canceled", vbInformation, "Code List Cancel"
AddIntellWord
RT_CodeListCancel = False
End Function
Private Function RT_CodeListChar(ByVal Control As CodeSenseCtl.ICodeSense, ByVal ListCtrl As CodeSenseCtl.ICodeList, ByVal wChar As Long, ByVal lKeyData As Long) As Boolean
'MsgBox wChar, vbInformation, "Code List Char"
RT_CodeListChar = False
End Function
Private Function RT_CodeListSelChange(ByVal Control As CodeSenseCtl.ICodeSense, ByVal ListCtrl As CodeSenseCtl.ICodeList, ByVal lItem As Long) As String

sIntellText = ListCtrl.GetItemText(lItem)
RT_CodeListSelChange = ""
End Function
Private Function RT_CodeListSelMade(ByVal Control As CodeSenseCtl.ICodeSense, ByVal ListCtrl As CodeSenseCtl.ICodeList) As Boolean
AddIntellWord
RT_CodeListSelMade = False
End Function
Private Function RT_CodeListSelWord(ByVal Control As CodeSenseCtl.ICodeSense, ByVal ListCtrl As CodeSenseCtl.ICodeList, ByVal lItem As Long) As Boolean
'MsgBox lItem, vbCritical, "SELWORD"
RT_CodeListSelWord = True
End Function
Private Function RT_CodeTip(ByVal Control As CodeSenseCtl.ICodeSense) As CodeSenseCtl.cmToolTipType
RT_CodeTip = cmToolTipTypeNormal
End Function
Private Function RT_CodeTipCancel(ByVal Control As CodeSenseCtl.ICodeSense, ByVal ToolTipCtrl As CodeSenseCtl.ICodeTip) As Boolean
'
End Function
Private Sub RT_CodeTipUpdate(ByVal Control As CodeSenseCtl.ICodeSense, ByVal ToolTipCtrl As CodeSenseCtl.ICodeTip)
'
End Sub

Private Function RT_MouseDown(ByVal Control As CodeSenseCtl.ICodeSense, ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long) As Boolean

GetRange
If Button = 2 Then
Me.PopupMenu Me.mnuEdit
End If
End Function

Private Function RT_MouseUp(ByVal Control As CodeSenseCtl.ICodeSense, ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long) As Boolean
GetRange
End Function

Private Sub RT_SelChange(ByVal Control As CodeSenseCtl.ICodeSense)
DoHighLight
End Sub

Private Sub GetRange()
Dim R As CodeSenseCtl.Range
Dim LLine As Long
Dim LCurrent As Long
Set R = RT.GetSel(False)
LLine = R.EndLineNo
LCurrent = R.EndColNo
LLine = LLine + 1
LCurrent = LCurrent + 1
StatusBar1.Panels(1).Text = "Line: " & LLine
StatusBar1.Panels(2).Text = "Col: " & LCurrent
End Sub

Private Sub AddIntellWord()
'add the selected word to the editor
Dim R As CodeSenseCtl.Range
    If sIntellText <> "" Then
   'MsgBox LBoxPos
        Set R = RT.GetSel(False)
        'set the position
        R.StartColNo = LBoxPos
        R.EndColNo = R.EndColNo
        RT.SetSel R, False
    'MsgBox RT.SelText
        'change the text
        R.StartColNo = R.EndColNo + Len(sIntellText)
        R.EndColNo = R.EndColNo + Len(sIntellText)
        RT.SelText = sIntellText
        RT.SetSel R, False

        sIntellText = ""
End If
End Sub
Public Function LoadIntellList(sWord As String, ByVal ListCtrl As CodeSenseCtl.ICodeList) As Boolean

On Error GoTo EH
Dim L As Long
Dim Ndb As Database
Dim TNdb As TableDef
Dim FLDnssc As Field
Dim RECndb As Recordset
Dim wrkDefault As Workspace
Dim I As Long
Dim bFound As Boolean
If CheckFile(GetDatabaseFilename) = False Then
MsgBox "Could not find main database!", vbCritical, "Error"
LoadIntellList = False
Exit Function
End If


Screen.MousePointer = 11
Set Ndb = OpenDatabase(GetDatabaseFilename)
Set RECndb = Ndb.OpenRecordset("CodeEdit", dbOpenDynaset)
If RECndb.BOF = True And RECndb.EOF = True Then
MsgBox "no records"
Else
    Do While Not RECndb.EOF
    DoEvents
    'See if this is a record we want
    If LCase(RECndb.Fields("type").Value) = "editor" Then
            If LCase(RECndb.Fields("header").Value) = LCase(sWord) Then
                For I = 1 To IMGIntellisence.ListImages.Count
                    If LCase(IMGIntellisence.ListImages(I).Tag) = LCase(RECndb.Fields("icontype").Value) Then
                        lImage = I - 1
                        Exit For
                    End If
                Next I
                ListCtrl.AddItem RECndb.Fields("data").Value, lImage
            End If
    End If
    RECndb.MoveNext
    Loop
End If
RECndb.Close
Ndb.Close
LoadIntellList = True
Screen.MousePointer = 0
Exit Function
EH:
Screen.MousePointer = 0
LoadIntellList = False
MsgBox Err.Description, vbCritical, "Load Intell List"
Exit Function
End Function

Public Function ReturnImage(sImage As String) As Integer
On Error GoTo EH
Dim I As Integer
For I = 1 To IMGIntellisence.ListImages.Count
If LCase(IMGIntellisence.ListImages(I).Tag) = LCase(sImage) Then
ReturnImage = I - 1
Exit Function
End If
Next I
ReturnImage = 0
Exit Function
EH:
MsgBox Err.Description, vbCritical, "Return Image"
Exit Function
End Function

Private Sub LoadSubs(sType As String)
On Error GoTo EH
Dim L As Long
Dim Ndb As Database
Dim TNdb As TableDef
Dim FLDnssc As Field
Dim RECndb As Recordset
Dim wrkDefault As Workspace
Dim I As Long
Dim bFound As Boolean

Combo2.Clear

If CheckFile(GetDatabaseFilename) = False Then
MsgBox "Could not find main database!", vbCritical, "Error"
Exit Sub
End If


Screen.MousePointer = 11
Set Ndb = OpenDatabase(GetDatabaseFilename)
Set RECndb = Ndb.OpenRecordset("CodeEdit", dbOpenDynaset)
If RECndb.BOF = True And RECndb.EOF = True Then
MsgBox "no records"
Else
    Do While Not RECndb.EOF
    DoEvents
    'See if this is a record we want
    If LCase(RECndb.Fields("type").Value) = "subs" Then
            If LCase(RECndb.Fields("header").Value) = LCase(sType) Then
                Combo2.AddItem RECndb.Fields("data").Value, lImage
            End If
    End If
    RECndb.MoveNext
    Loop
End If
RECndb.Close
Ndb.Close
Screen.MousePointer = 0
Exit Sub
EH:
Screen.MousePointer = 0
MsgBox Err.Description, vbCritical, "Load Subs"
Exit Sub
End Sub

Private Sub LoadCodeList()
On Error GoTo EH
Dim L As Long
Dim Ndb As Database
Dim TNdb As TableDef
Dim FLDnssc As Field
Dim RECndb As Recordset
Dim wrkDefault As Workspace
Dim I As Long
Dim bFound As Boolean



If CheckFile(GetDatabaseFilename) = False Then
MsgBox "Could not find main database!", vbCritical, "Error"
Exit Sub
End If


Screen.MousePointer = 11
Set Ndb = OpenDatabase(GetDatabaseFilename)
Set RECndb = Ndb.OpenRecordset("CodeEdit", dbOpenDynaset)
If RECndb.BOF = True And RECndb.EOF = True Then
MsgBox "no records"
Else
    Do While Not RECndb.EOF
    DoEvents
    'See if this is a record we want
    If LCase(RECndb.Fields("type").Value) = "statements" Then
    I = mnuCodes.UBound + 1
    Load mnuCodes(I)
    mnuCodes(I).Caption = RECndb.Fields("header").Value
    End If
    RECndb.MoveNext
    Loop
End If
RECndb.Close
Ndb.Close
mnuCodes(0).Visible = False
Screen.MousePointer = 0
Exit Sub
EH:
Screen.MousePointer = 0
MsgBox Err.Description, vbCritical, "Load Codes"
Exit Sub
End Sub

Private Sub InsertCodes(sWhat As String)
On Error GoTo EH
Dim L As Long
Dim Ndb As Database
Dim TNdb As TableDef
Dim FLDnssc As Field
Dim RECndb As Recordset
Dim wrkDefault As Workspace
Dim I As Long
Dim bFound As Boolean

If CheckFile(GetDatabaseFilename) = False Then
MsgBox "Could not find main database!", vbCritical, "Error"
Exit Sub
End If

Screen.MousePointer = 11
Set Ndb = OpenDatabase(GetDatabaseFilename)
Set RECndb = Ndb.OpenRecordset("CodeEdit", dbOpenDynaset)
If RECndb.BOF = True And RECndb.EOF = True Then
MsgBox "no records"
Else
    Do While Not RECndb.EOF
    DoEvents
    'See if this is a record we want
    If LCase(RECndb.Fields("type").Value) = "statements" Then
    If LCase(RECndb.Fields("header").Value) = LCase(sWhat) Then
    RT.SelText = RECndb.Fields("data").Value
    End If
    End If
    RECndb.MoveNext
    Loop
End If
RECndb.Close
Ndb.Close
mnuCodes(0).Visible = False
Screen.MousePointer = 0
Exit Sub
EH:
Screen.MousePointer = 0
MsgBox Err.Description, vbCritical, "Insert Codes"
Exit Sub
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case Is = 1 'new
mnuNew_Click
Case Is = 2 'open
mnuOpen_Click
Case Is = 3 'save
mnuSave_Click
Case Is = 4 'seperator
Case Is = 5 'undo
mnuUndo_Click
Case Is = 6 'redo
mnuRedo_Click
Case Is = 7 'seperator
Case Is = 8 'copy
mnuCopy_Click
Case Is = 9 'cut
mnuCut_Click
Case Is = 10 'paste
mnuPaste_Click
Case Is = 11 'seperator
Case Is = 12 'help
mnuHelp_Click
End Select
End Sub
