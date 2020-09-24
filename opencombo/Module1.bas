Attribute VB_Name = "Module1"
Option Explicit

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Const CB_SHOWDROPDOWN = &H14F
Public Const CB_GETDROPPEDSTATE = &H157

Public Sub OpenCloseCombo(chwnd As Long)
    Dim rc As Long
    rc = SendMessage(chwnd, CB_GETDROPPEDSTATE, 0, 0)
    If rc = 0 Then
        SendMessage chwnd, CB_SHOWDROPDOWN, True, 0
    Else
        SendMessage chwnd, CB_SHOWDROPDOWN, False, 0
    End If
End Sub

