Attribute VB_Name = "modAutoSaver"
Public Declare Function InitCommonControls Lib "COMCTL32.DLL" () As Long

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const SW_SHOWNORMAL = 1

Sub main()
InitCommonControls
Load MainForm
If InStr(1, Command$, "-silent") > 0 Then
    MainForm.mnuMShow.Caption = "&Show"
Else
    MainForm.Show
End If
End Sub
