Attribute VB_Name = "ini"
' ini.bas Module by RegX
' Copyright 2002 DGS
' You may freely use this code as long as
' All Copyright information remains intact
'-----------------------------------------
' For easier ini manipulation

Option Compare Text
Option Explicit

Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyName As Any, ByVal lsString As Any, ByVal lplFilename As String) As Long
Public Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPriviteProfileIntA" (ByVal lpApplicationname As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Function AppPath() As String
AppPath = App.path & IIf(Right(App.path, 1) = "\", "", "\")
End Function
Public Sub PutIni(iniFile As String, iniHead As String, iniKey As String, iniVal As String)
Dim IniFileName As String
IniFileName = AppPath & iniFile
WritePrivateProfileString iniHead, iniKey, iniVal, IniFileName
End Sub

Public Function GetIni(iniFile As String, iniHead As String, iniKey As String, iniDefault As String) As String
Dim IniFileName As String
IniFileName = AppPath & iniFile
Dim Temp As String
Temp = "                                                                               "
GetPrivateProfileString iniHead, iniKey, iniDefault, Temp, Len(Temp), IniFileName
GetIni = Trim(Temp)
GetIni = Left(GetIni, Len(GetIni) - 1)
End Function
