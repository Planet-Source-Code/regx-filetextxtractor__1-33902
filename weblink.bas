Attribute VB_Name = "weblink"
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const SW_SHOWNORMAL = 1

Public Sub gotoweb(url As String)
Dim Success As Long

Success = ShellExecute(0&, vbNullString, url, vbNullString, "C:\", SW_SHOWNORMAL)

End Sub
