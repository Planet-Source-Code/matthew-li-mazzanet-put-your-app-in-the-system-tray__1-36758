Attribute VB_Name = "systray"
' ************************************
' * Systray Module                   *
' * Copyright (c) Matthew Li 2002    *
' * You may use this in your own     *
' * applications as long as you give *
' * credit to me                     *
' ************************************
Option Explicit
Public OldWindowProc As Long
Public frmTargetForm As Form
Public mnuMenu As Menu
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Public Const WM_USER = &H400
Public Const WM_LBUTTONUP = &H202
Public Const WM_MBUTTONUP = &H208
Public Const WM_RBUTTONUP = &H205
Public Const TRAY_CALLBACK = (WM_USER + 1001&)
Public Const GWL_WNDPROC = (-4)
Public Const GWL_USERDATA = (-21)
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const NIM_ADD = &H0
Public Const NIF_MESSAGE = &H1
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2

Public Type NOTIFYICONDATA
cbSize As Long
hwnd As Long
uID As Long
uFlags As Long
uCallbackMessage As Long
hIcon As Long
szTip As String * 64
End Type

Private IconData As NOTIFYICONDATA
Public Function NewWindowProc(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
If Msg = TRAY_CALLBACK Then
If lParam = WM_LBUTTONUP Then
If frmTargetForm.WindowState = vbMinimized Then frmTargetForm.WindowState = 0
frmTargetForm.SetFocus
Exit Function
End If
If lParam = WM_RBUTTONUP Then
frmTargetForm.PopupMenu mnuMenu
Exit Function
End If
End If
NewWindowProc = CallWindowProc(OldWindowProc, hwnd, Msg, wParam, lParam)
End Function
Public Sub AddToTray(frm As Form, mnu As Menu)
Set frmTargetForm = frm
Set mnuMenu = mnu
OldWindowProc = SetWindowLong(frm.hwnd, GWL_WNDPROC, AddressOf NewWindowProc)
IconData.uID = 0
IconData.hwnd = frm.hwnd
IconData.cbSize = Len(IconData)
IconData.hIcon = frm.Icon.Handle
IconData.uFlags = NIF_ICON
IconData.uCallbackMessage = TRAY_CALLBACK
IconData.uFlags = IconData.uFlags Or NIF_MESSAGE
IconData.cbSize = Len(IconData)
Shell_NotifyIcon NIM_ADD, IconData
End Sub
Public Sub RemoveFromTray()
On Error Resume Next
IconData.uFlags = 0
Shell_NotifyIcon NIM_DELETE, IconData
SetWindowLong frmTargetForm.hwnd, GWL_WNDPROC, OldWindowProc
End Sub
Public Sub SetTrayTip(tip As String)
IconData.szTip = tip & vbNullChar
IconData.uFlags = NIF_TIP
Shell_NotifyIcon NIM_MODIFY, IconData
End Sub
Public Sub SetTrayIcon(pic As Picture)
If pic.Type <> vbPicTypeIcon Then Exit Sub
IconData.hIcon = pic.Handle
IconData.uFlags = NIF_ICON
Shell_NotifyIcon NIM_MODIFY, IconData
End Sub

Public Sub ShowTray()
RemoveFromTray
AddToTray frmMain, frmMain.mnuTray
SetTrayTip frmMain.Text1
SetTrayIcon frmMain.Picture1.Picture
End Sub
Public Sub HideTray()
RemoveFromTray
End Sub
