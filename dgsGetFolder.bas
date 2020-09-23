Attribute VB_Name = "dgsGetFolder"
' GetFolder Module by RegX
' Copyright 2002 DGS
' You may freely use this code as long as
' All Copyright information remains intact
'-----------------------------------------
' This module simply opens a browse for folder dialog

Option Explicit

Private Type BrowseInfo
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type
Const BIF_RETURNONLYFSDIRS = 1
Const MAX_PATH = 260
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long

Public Function getFolder() As String
    Dim iNull As Integer, lpIDList As Long, lResult As Long
    Dim sPath As String, udtBI As BrowseInfo

    With udtBI
        'Set the owner window
        '.hWndOwner = Me.hWnd
        'lstrcat appends the two strings and returns the memory address
        .lpszTitle = lstrcat("C:\", "")
        'Return only if the user selected a directory
        .ulFlags = BIF_RETURNONLYFSDIRS
    End With

    'Show the 'Browse for folder' dialog
    lpIDList = SHBrowseForFolder(udtBI)
    If lpIDList Then
        sPath = String$(MAX_PATH, 0)
        'Get the path from the IDList
        SHGetPathFromIDList lpIDList, sPath
        'free the block of memory
        CoTaskMemFree lpIDList
        iNull = InStr(sPath, vbNullChar)
        If iNull Then
            sPath = Left$(sPath, iNull - 1)
        End If
    End If
    
        getFolder = sPath & ""
End Function
