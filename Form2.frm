VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form2 
   Caption         =   "Save to File"
   ClientHeight    =   1560
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6015
   LinkTopic       =   "Form2"
   ScaleHeight     =   1560
   ScaleWidth      =   6015
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picstatus 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      ScaleHeight     =   225
      ScaleWidth      =   5955
      TabIndex        =   6
      Top             =   1275
      Width           =   6015
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3360
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "txt"
      DialogTitle     =   "Save File"
      Filter          =   "Text(*.txt,*.rtf)|*.txt;*.rtf|All Files(*.*)|*.*"
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "Cancel"
      Height          =   405
      Left            =   3360
      TabIndex        =   5
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   405
      Left            =   4680
      TabIndex        =   4
      Top             =   720
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Save Options"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3015
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Form2.frx":0000
         Left            =   2160
         List            =   "Form2.frx":000D
         TabIndex        =   3
         Text            =   ";"
         Top             =   600
         Width           =   735
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Delimited By"
         Height          =   255
         Left            =   840
         TabIndex        =   2
         Top             =   600
         Width           =   1935
      End
      Begin VB.OptionButton Option1 
         Caption         =   "1 Eail item per line"
         Height          =   255
         Left            =   840
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   2055
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdcancel_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
Dim delimiter
Dim fs As FileSystemObject
Dim ts As TextStream

Set fs = New FileSystemObject

If Option1 = True Then
    delimiter = vbCrLf
ElseIf Me.Combo1 & "" = "" Then
    MsgBox "Please select or type a delimiter", vbOKOnly, "No delimiter selected!"
    Exit Sub
ElseIf Me.Combo1 = "Tab" Then
    delimiter = vbTab
Else
    delimiter = Me.Combo1
End If

CommonDialog1.ShowSave
If CommonDialog1.FileName & "" <> "" Then
    If fs.FileExists(CommonDialog1.FileName) = True Then
        If MsgBox("File Exist, Replace File?", vbYesNo, "Replace File") = vbNo Then
            Exit Sub
        End If
    End If
    picstatus.Print "Saving File, Please Wait!"
    Set ts = fs.CreateTextFile(CommonDialog1.FileName, True)
    Dim x As Long
    For x = 1 To Form1.ListView1.ListItems.Count
        ts.Write Form1.ListView1.ListItems.Item(x).SubItems(1) & delimiter
    Next
    
    ts.Close
    Set fs = Nothing
    Set ts = Nothing
    Unload Me
End If
End Sub
