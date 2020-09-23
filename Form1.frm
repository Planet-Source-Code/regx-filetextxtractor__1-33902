VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FileTextXtractor"
   ClientHeight    =   7830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11145
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7830
   ScaleWidth      =   11145
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   1815
      Left            =   6720
      ScaleHeight     =   1755
      ScaleWidth      =   4155
      TabIndex        =   24
      Top             =   720
      Width           =   4215
      Begin VB.CheckBox ChkCaseSensative 
         Caption         =   "Case Sensitive"
         Height          =   375
         Left            =   120
         TabIndex        =   29
         Top             =   1440
         Value           =   1  'Checked
         Width           =   3375
      End
      Begin VB.TextBox txtsearchstr 
         Height          =   330
         Left            =   120
         TabIndex        =   26
         Top             =   1080
         Width           =   3975
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Form1.frx":0442
         Left            =   120
         List            =   "Form1.frx":045B
         TabIndex        =   25
         Text            =   "Select a predefined regx"
         Top             =   480
         Width           =   3975
      End
      Begin VB.Label Label4 
         Caption         =   "The search string must be contained in (parenthesis) To find out why and for more regular expresion help click here."
         Height          =   495
         Left            =   0
         TabIndex        =   28
         Top             =   0
         Width           =   4215
      End
      Begin VB.Label Label3 
         Caption         =   "Search String (Regular Expression)"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   840
         Width           =   2895
      End
   End
   Begin VB.PictureBox picstatus1 
      Align           =   2  'Align Bottom
      BackColor       =   &H00808080&
      ForeColor       =   &H8000000E&
      Height          =   300
      Left            =   0
      ScaleHeight     =   240
      ScaleWidth      =   11085
      TabIndex        =   22
      Top             =   7260
      Width           =   11145
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3570
      Left            =   0
      TabIndex        =   21
      Top             =   3360
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   6297
      View            =   3
      SortOrder       =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "File"
         Object.Tag             =   "File"
         Text            =   "File"
         Object.Width           =   13231
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "Match"
         Object.Tag             =   "Match"
         Text            =   "Match"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "Count"
         Object.Tag             =   "Count"
         Text            =   "Count"
         Object.Width           =   2646
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3480
      Top             =   6840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Open File"
      Filter          =   "Text Files(*.txt,*.rtf)|*.txt;*.rtf|MS Office(*.doc,*.mdb,*.xls) |*.doc;*.mdb;*.xls|All Files(*.*)|*.*"
   End
   Begin RichTextLib.RichTextBox RTF1 
      Height          =   2220
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   5250
      _ExtentX        =   9260
      _ExtentY        =   3916
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"Form1.frx":0577
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3060
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   5398
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      TabMaxWidth     =   3881
      MouseIcon       =   "Form1.frx":0658
      TabCaption(0)   =   "Extract From File"
      TabPicture(0)   =   "Form1.frx":0674
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lblscanfiletarget"
      Tab(0).Control(1)=   "cmdSelectFile"
      Tab(0).Control(2)=   "cmdScanFile"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Extract From Folder"
      TabPicture(1)   =   "Form1.frx":0BB6
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lblscanfoldertarget"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label1"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cmdRemoveExt"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cmdAddExt"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "txtext"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "lstExt"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "cmdScanFolder"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "cmbRecursionlvl"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "cmdStopScanningFolder"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "cmdSelectFolder"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).ControlCount=   11
      Begin VB.CommandButton cmdScanFile 
         BackColor       =   &H00008000&
         Caption         =   "Scan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -64920
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   2640
         Width           =   855
      End
      Begin VB.CommandButton cmdSelectFile 
         BackColor       =   &H0000FFFF&
         Caption         =   "Select File"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -74880
         TabIndex        =   17
         Top             =   360
         Width           =   1395
      End
      Begin VB.CommandButton cmdSelectFolder 
         BackColor       =   &H0000FFFF&
         Caption         =   "Select Folder"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   1395
      End
      Begin VB.CommandButton cmdStopScanningFolder 
         BackColor       =   &H000000C0&
         Caption         =   "Stop"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   10080
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   2640
         Width           =   855
      End
      Begin VB.ComboBox cmbRecursionlvl 
         Height          =   315
         ItemData        =   "Form1.frx":10F8
         Left            =   5400
         List            =   "Form1.frx":110E
         TabIndex        =   13
         Text            =   "0 - unlimited"
         Top             =   960
         Width           =   1290
      End
      Begin VB.CommandButton cmdScanFolder 
         BackColor       =   &H00008000&
         Caption         =   "Scan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   9120
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   2640
         Width           =   855
      End
      Begin VB.ListBox lstExt 
         Height          =   840
         ItemData        =   "Form1.frx":1130
         Left            =   5400
         List            =   "Form1.frx":1137
         MultiSelect     =   2  'Extended
         TabIndex        =   10
         Top             =   1560
         Width           =   1290
      End
      Begin VB.TextBox txtext 
         Height          =   285
         Left            =   5400
         TabIndex        =   9
         Top             =   2400
         Width           =   795
      End
      Begin VB.CommandButton cmdAddExt 
         Caption         =   "Add"
         Height          =   240
         Left            =   6240
         TabIndex        =   8
         Top             =   2400
         Width           =   435
      End
      Begin VB.CommandButton cmdRemoveExt 
         Caption         =   "Remove"
         Height          =   255
         Left            =   5400
         TabIndex        =   7
         Top             =   2685
         Width           =   1275
      End
      Begin VB.Label Label1 
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Recursion Level"
         ForeColor       =   &H80000014&
         Height          =   255
         Left            =   5400
         TabIndex        =   23
         Top             =   720
         Width           =   1290
      End
      Begin VB.Label lblscanfiletarget 
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   -73440
         TabIndex        =   18
         Top             =   360
         Width           =   9375
      End
      Begin VB.Label lblscanfoldertarget 
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   1545
         TabIndex        =   16
         Top             =   360
         Width           =   9390
      End
      Begin VB.Label Label2 
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Extensions"
         ForeColor       =   &H80000014&
         Height          =   255
         Left            =   5400
         TabIndex        =   11
         Top             =   1320
         Width           =   1290
      End
   End
   Begin VB.CommandButton cmdremoveselectedemails 
      Caption         =   "Remove Selected"
      Height          =   285
      Left            =   7680
      TabIndex        =   5
      Top             =   6960
      Width           =   1575
   End
   Begin VB.CommandButton cmdsavelist 
      Caption         =   "Save to File"
      Height          =   285
      Left            =   9360
      TabIndex        =   4
      Top             =   6960
      Width           =   1695
   End
   Begin VB.CommandButton cmdClearEmailList 
      Caption         =   "Clear List"
      Height          =   285
      Left            =   6600
      TabIndex        =   2
      Top             =   6960
      Width           =   975
   End
   Begin VB.PictureBox picstatus2 
      Align           =   2  'Align Bottom
      BackColor       =   &H00404040&
      ForeColor       =   &H8000000E&
      Height          =   270
      Left            =   0
      ScaleHeight     =   210
      ScaleWidth      =   11085
      TabIndex        =   1
      Top             =   7560
      Width           =   11145
   End
   Begin VB.Label lblscancount 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   30
      TabIndex        =   20
      Top             =   3060
      Width           =   5055
   End
   Begin VB.Label lbllistcount 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   6960
      Width           =   3075
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' VeryFast File-Email Extractor
' Copyright 2002 DGS
'Written by Gary Varnell
'=============================================
'Needs reference to:
'Microsoft Scripting Runtime
'Microsoft VBScript Regular Expressions 5.5
'download at http://msdn.microsoft.com/downloads/default.asp?URL=/downloads/sample.asp?url=/msdn-files/027/001/733/msdncompositedoc.xml
'=============================================
Option Explicit
Option Compare Binary
Dim extensions As Dictionary
Dim matchcache As Dictionary
Dim regx1 As RegExp
Dim Matches As MatchCollection
Dim Match As Match
Public WithEvents DGSDirScan1 As DGSDirScanner
Attribute DGSDirScan1.VB_VarHelpID = -1
Dim scancount As Long



Private Sub ChkCaseSensative_Click()
txtsearchstr_Change
End Sub

Private Sub txtsearchstr_Change()
On Error GoTo bail
Dim ignorecase As Boolean
If ChkCaseSensative.Value = 1 Then
    ignorecase = False
Else
    ignorecase = True
End If
'regx for emails
    Set regx1 = New RegExp   ' Create Regular expresion to extract valid email addresses
    regx1.Pattern = txtsearchstr   ' Set pattern.
    regx1.ignorecase = ignorecase   ' Set case insensitivity.
    regx1.Global = True        ' Set global applicability.
Exit Sub
bail:
MsgBox Err.Number & " " & Err.Description
End Sub

Private Sub cmdScanFolder_Click()
Set matchcache = New Dictionary
' make sure we have a search string
    If Me.txtsearchstr = "" Then
        MsgBox "Please select or type a search string", vbInformation, "No Search String"
        Exit Sub
    End If


' make sure user selected a folder
If Me.lblscanfoldertarget & "" = "" Then
    MsgBox "Please select a folder to scan", vbOKOnly, "Nothing to do"
    Exit Sub
End If

' make sure folder exist
Dim fs As FileSystemObject
Set fs = New FileSystemObject
If fs.FolderExists(Me.lblscanfoldertarget) = False Then
    MsgBox "The folder you selected does not exist!", vbCritical, "Error"
    Exit Sub
End If

' create a dictionary object and add all the extensions to be scanned
Set extensions = New Dictionary
Dim a As Long
extensions.RemoveAll
For a = 0 To lstExt.listcount - 1
    extensions.Add lstExt.List(a), lstExt.List(a)
Next

scancount = 0
' Start recursive directory scan
' Fires the new folder and new file event (below) for every file/folder
DGSDirScan1.Scan Me.lblscanfoldertarget
status1 "All Done =)"
status2 ""
Beep

End Sub

Private Sub Combo1_Click()
If InStr(1, Combo1, ": ") Then
    Me.txtsearchstr = Mid(Combo1, InStr(1, Combo1, ": ") + 2)
End If
End Sub

Private Sub DGSDirScan1_newdir(d As Scripting.Folder)
status1 "Scanning Dir " & d.path
End Sub

Private Sub DGSDirScan1_newfile(f As Scripting.IFile)
On Error Resume Next
'On Error GoTo bail
Set matchcache = New Dictionary
status2 f.Name
Dim ext As String
Dim x As String
x = InStrRev(f.path, ".")
If x < 1 Then Exit Sub
ext = Mid(f.path, x)

If extensions.Exists(ext) = True Or extensions.Exists(".*") = True Then
    status2 "Opening File " & f.path
    Me.RTF1.LoadFile f.path
    ExtractText (f.path)
End If
Exit Sub
bail:
Select Case MsgBox("Error: " & Err.Number & " " & Err.Description, vbAbortRetryIgnore, "Error")
Case vbAbort
Me.DGSDirScan1.Cancel
Exit Sub
Case vbRetry
Resume
Case vbIgnore
Resume Next
End Select
End Sub

Private Sub cmdremoveExt_Click()
On Error Resume Next
If Me.lstExt.SelCount = 0 Then
    MsgBox "Please select an extension to remove", vbOKOnly, "Nothing to do!"
End If
Dim x As Long
x = 0
While lstExt.SelCount > 0
If lstExt.Selected(x) = True Then
    lstExt.RemoveItem x
Else
    x = x + 1
End If
Wend

End Sub

Private Sub cmdremoveselectedemails_Click()
On Error Resume Next
Dim x As Long
x = 1
While x < ListView1.ListItems.Count + 1
    If ListView1.ListItems.Item(x).Selected = True Then
        ListView1.ListItems.Remove x
    Else
        x = x + 1
    End If
Wend
End Sub

Private Sub cmdsavelist_Click()
Form2.Show vbModal
End Sub

Private Sub cmbRecursionlvl_Change()
Me.DGSDirScan1.Scandepth = cmbRecursionlvl
End Sub

Private Sub cmbRecursionlvl_Click()
Me.DGSDirScan1.Scandepth = Mid(cmbRecursionlvl, 1, 1)
End Sub

Private Sub cmdSelectFolder_Click()
Me.lblscanfoldertarget = getFolder & ""
lblscancount.Caption = ""
status1 " Press the scan button to scan the selected folder."
status2 ""
End Sub

Private Sub cmdStopScanningFolder_Click()
Me.DGSDirScan1.Cancel
End Sub

Private Sub cmdAddExt_Click()
If txtext = "" Then
     MsgBox "You must type an extension to add.", vbOKOnly, "No Extension"
    txtext = Right(txtext, 4)
ElseIf InStr(1, txtext, ".") < 1 Then
    MsgBox "Extension must be in the form .txt (begining with a "".""" & vbCrLf & "To scan all extensions type .*", vbOKOnly, "Invalid Extension"
    Exit Sub
End If
lstExt.AddItem txtext
txtext = ""
End Sub

Private Sub cmdSelectFile_Click()
On Error GoTo bail

Me.CommonDialog1.ShowOpen
If Me.CommonDialog1.FileName & "" <> "" Then
    ' check that file exist
    Dim fs As FileSystemObject
    Set fs = New FileSystemObject

    If fs.FileExists(CommonDialog1.FileName) Then
        status2 "Opening File " & CommonDialog1.FileName
        Me.RTF1.LoadFile CommonDialog1.FileName
        status1 "Press the scan button to scan this file."
        status2 ""
        lblscancount.Caption = ""
        lblscanfiletarget.Caption = CommonDialog1.FileName
    Else
    MsgBox CommonDialog1.FileName & " doesn't exist", vbCritical, "File not found"
    End If

End If
Exit Sub
bail:
MsgBox Err.Number & " " & Err.Description, vbCritical, "Unexpected error"
Err.Clear
End Sub

Private Sub cmdClearEmailList_Click()
ListView1.ListItems.Clear
End Sub

Private Sub cmdScanFile_Click()
' make sure we have a search string
    If Me.txtsearchstr = "" Then
        MsgBox "Please select or type a search string", vbInformation, "No Search String"
        Exit Sub
    End If
Set matchcache = New Dictionary
scancount = 0
ExtractText lblscanfiletarget.Caption
    status1 "All Done =)"
    status2 ""
    Beep
End Sub

Private Sub ExtractText(path As String)
        On Error Resume Next ' or we could implicitly handle duplicates
        status2 "Finding Matches"
        Set Matches = regx1.Execute(RTF1.Text)     ' Execute search.
        For Each Match In Matches
            'add to cache
            If matchcache.Exists(CStr(path & Match.SubMatches(0))) Then
                matchcache(CStr(path & Match.SubMatches(0))) = matchcache(CStr(path & Match.SubMatches(0))) + 1
                Me.ListView1.ListItems(CStr(path & Match.SubMatches(0))).ListSubItems.Item(2) = matchcache(CStr(path & Match.Value))
            Else
                    matchcache.Add CStr(path & Match.SubMatches(0)), 1
                ' add to listview
                Me.ListView1.ListItems.Add , CStr(path & Match.SubMatches(0)), path
                Me.ListView1.ListItems(CStr(path & Match.SubMatches(0))).ListSubItems.Add , "Match", Match.SubMatches(0)
                Me.ListView1.ListItems(CStr(path & Match.SubMatches(0))).ListSubItems.Add , "Count", 1
            End If
            ' update listcounter
            lbllistcount.Caption = ListView1.ListItems.Count & " items in list"
        Next
        Set Matches = Nothing
        scancount = scancount + 1
        lblscancount.Caption = scancount & " files scanned"
End Sub

Private Sub Form_Load()
Me.Caption = Me.Caption & " " & App.Major & "." & App.Minor & "." & App.Revision
Set DGSDirScan1 = New DGSDirScanner


'load lstext from ini
lstExt.Clear
Dim ext As String
Dim listcount As Long
listcount = CLng(GetIni("extractor.ini", "lstExt", "ListCount", "1"))
Dim x As Long
For x = 0 To listcount - 1
    ext = GetIni("extractor.ini", "lstExt", CStr(x), ".txt")
    lstExt.AddItem ext, x
Next x
End Sub

Private Sub Form_Unload(Cancel As Integer)
Me.DGSDirScan1.Cancel

' save lstExt to ini
PutIni "extractor.ini", "lstExt", "ListCount", CStr(lstExt.listcount)
Dim x As Long
For x = 0 To lstExt.listcount - 1
    PutIni "extractor.ini", "lstExt", CStr(x), lstExt.List(x)
Next x

End Sub

Private Sub status2(msg As String)
picstatus2.Cls
picstatus2.Print msg
End Sub

Private Sub status1(msg As String)
picstatus1.Cls
picstatus1.Print msg
End Sub


Private Sub Label4_Click()
' goto help web page
gotoweb "http://www.2dgs.com/prog/filetextxtractor/"
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
' Sort by clicked column
' reverse sort order if current column is sortkey
If ListView1.SortKey = ColumnHeader.Index - 1 Then
    If ListView1.SortOrder = lvwAscending Then
        ListView1.SortOrder = lvwDescending
    Else
        ListView1.SortOrder = lvwAscending
    End If
Else ' set selected column as sort column
    ListView1.SortOrder = lvwAscending
    ListView1.SortKey = ColumnHeader.Index - 1
End If
ListView1.Sorted = True
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
Form3.update Item.Text
Form3.Show vbModal
End Sub



