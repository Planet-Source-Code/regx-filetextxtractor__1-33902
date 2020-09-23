VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{683364A1-B37D-11D1-ADC5-006008A5848C}#1.0#0"; "DHTMLED.OCX"
Begin VB.Form Form3 
   Caption         =   "Preview"
   ClientHeight    =   4320
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9375
   LinkTopic       =   "Form3"
   ScaleHeight     =   4320
   ScaleWidth      =   9375
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   4215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   7435
      _Version        =   393216
      TabOrientation  =   1
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   5
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "Text Preview"
      TabPicture(0)   =   "Form3.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "RichTextBox1"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "HTML Preview"
      TabPicture(1)   =   "Form3.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "DHTMLEdit1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin DHTMLEDLibCtl.DHTMLEdit DHTMLEdit1 
         CausesValidation=   0   'False
         Height          =   3255
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   9015
         ActivateApplets =   0   'False
         ActivateActiveXControls=   0   'False
         ActivateDTCs    =   0   'False
         ShowDetails     =   0   'False
         ShowBorders     =   0   'False
         Appearance      =   1
         Scrollbars      =   -1  'True
         ScrollbarAppearance=   1
         SourceCodePreservation=   -1  'True
         AbsoluteDropMode=   0   'False
         SnapToGrid      =   0   'False
         SnapToGridX     =   50
         SnapToGridY     =   50
         BrowseMode      =   -1  'True
         UseDivOnCarriageReturn=   0   'False
      End
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   3615
         Left            =   -74880
         TabIndex        =   1
         Top             =   120
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   6376
         _Version        =   393217
         TextRTF         =   $"Form3.frx":0038
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub update(path As String)
On Error Resume Next
Me.Caption = path
Me.RichTextBox1.LoadFile path
Me.DHTMLEdit1.LoadDocument path

End Sub



Private Sub Form_Resize()
With Me.SSTab1
   .Top = 0
   .Left = 0
   .Width = Me.ScaleWidth
   .Height = Me.ScaleHeight
End With

With Me.RichTextBox1
   .Top = 10
   .Left = 30
   .Width = SSTab1.Width - 60
   .Height = SSTab1.Height - 500
End With

With Me.DHTMLEdit1
   .Top = 10
   .Left = 30
   .Width = SSTab1.Width - 60
   .Height = SSTab1.Height - 500
End With


End Sub

