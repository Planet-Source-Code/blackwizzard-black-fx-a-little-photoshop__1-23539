Attribute VB_Name = "OnTop"
Option Explicit
Declare Function SetSysColors Lib "user32" (ByVal nChanges As Long, lpSysColor As Long, lpColorValues As Long) As Long
Public Const COLOR_ACTIVECAPTION = 2
'API nécessaire pour le mode "toujours visible"
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long
'toujours visible
Public Function forward(who As Form) 'who correspond au nom de la form  | exemple: form1
Dim Resultat As Long
Const Flags = &H2 Or &H1 Or &H40 Or &H10
Resultat = SetWindowPos(who.hwnd, -1, 0, 0, 0, 0, Flags)
End Function



'annuler toujours visible
Public Function backward(who As Form)
Dim Resultat As Long
Const Flags = &H2 Or &H1 Or &H40 Or &H10
Resultat = SetWindowPos(who.hwnd, -2, 0, 0, 0, 0, Flags)
End Function

Public Function WeB(WebPage As String, actualfrmHWND As String)
On Error Resume Next
Dim cod
cod = ShellExecute(actualfrmHWND, vbNullString, WebPage, "", vbNullString, 1)
End Function



