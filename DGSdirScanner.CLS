VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DGSDirScanner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"walkdir"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' DirScanner Class Module by RegX
' Copyright 2002 DGS
' You may freely use this code as long as
' All Copyright information remains intact
'=============================================
'Needs reference to:
'Microsoft VBScript Regular Expressions 5.5
'download at http://msdn.microsoft.com/downloads/default.asp?URL=/downloads/sample.asp?url=/msdn-files/027/001/733/msdncompositedoc.xml
'=============================================
'
'Usage ---------------------------------------------------------------
'Dim WithEvents ds As DGSDirScanner
'Set ds = New dirScanner
'ds.Scan Dir1.Path
'
'Scandepth: is used to set the recursion level like so ds.ScanDepth=0
'0 = Infinate recursion (default)
'1 = No dir recursion (All files in scandir will be scanned)
'2 = 1 level of dir recursion + files
'3 = 2 levels of dir recursion + files
'etc
'
'use Cancel to Hault scanning
'---------------------------------------------------------------------
' Asside from being extremely simple, this recursive scanner exposes
' all folder and file attributes making it extremely useful.
'---------------------------------------------------------------------

Public Event newdir(d As Folder)
Attribute newdir.VB_Description = "Fires every time a directory is encountered.\r\nExposes folder object as d"
'Fires every time a directory is encountered exposes folder object as d

Public Event newfile(f As File)
Attribute newfile.VB_Description = "Fires every time a file is encountered.\r\nexposes file object as f"
' Fires every time a file is encountered exposes file object as f

'local variable(s) to hold property value(s)
Private mvarScandepth As Long 'local copy
Private hault As Boolean
Public Property Let Scandepth(ByVal vData As Long)
Attribute Scandepth.VB_Description = "Sets the recursion depth for scan.\r\n0 =  infinate recursion.\r\n1 = No dir recursion, all files in the start dir will be scanned.\r\n2 = 1 level dir recursion + files\r\n3 = 2 levels of dir recursion + files\r\n4 = 3 levels of dir recursion + files\r\netc."
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Scandepth = 5
    mvarScandepth = vData
End Property

Public Property Get Scandepth() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Scandepth
    Scandepth = mvarScandepth
End Property
Public Sub Cancel() ' Haults scanning execution
Attribute Cancel.VB_Description = "Haults scanning."
    hault = True
End Sub
Public Sub Scan(startdir) ' Start scanning
Attribute Scan.VB_Description = "Begins scanning.\r\nusage: scanobject.scan(path)"
' check that start dir was given
If startdir & "" = "" Then
    MsgBox "No scan target given", vbExclamation, "Error"
    Exit Sub
End If
' check that scan target has a trailing backslash
If Right(startdir, 1) <> "\" Then startdir = startdir & "\"

' dim variables
Dim fs As New FileSystemObject
Dim fld1, fld2 As Folder
Dim fil As File
Dim tmpdepth As Long
hault = False

Dim dirs As New Collection ' create collection for dirs
dirs.Add startdir
Do While dirs.Count > 0 And hault = False
    ' get sub folders
    Set fld1 = fs.getFolder(dirs.Item(1))
    tmpdepth = tmpdepth + 1
    If Scandepth = 0 Or tmpdepth < Scandepth Then
        For Each fld2 In fld1.SubFolders
            If hault = True Then Exit Sub
            RaiseEvent newdir(fld2)
            If Right(fld2.path, 1) <> "\" Then
                dirs.Add fld2.path & "\"
            Else
                dirs.Add fld2.path
            End If
            DoEvents: DoEvents
        Next
    End If
    ' get files
    For Each fil In fld1.Files
        If hault = True Then Exit Sub
        RaiseEvent newfile(fil)
            DoEvents: DoEvents
    Next
    dirs.Remove 1
Loop

Set fs = Nothing
Set dirs = Nothing
End Sub

