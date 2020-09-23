Attribute VB_Name = "mSysTray"
'--------------------------------------------------
' E-Mail Checker
' By H G Laughland
'
' TrayIcon Module
' Declarations and procedures for the tray icon.
'--------------------------------------------------

Public Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

Public Type NOTIFYICONDATA
 cbSize As Long
 hwnd As Long
 uID As Long
 uFlags As Long
 uCallbackMessage As Long
 hIcon As Long
 szTip As String * 64
End Type

Public Const NIM_ADD = &H0
Public Const NIM_DELETE = &H2
Public Const NIM_MODIFY = &H1

Public Const NIF_ICON = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_TIP = &H4

Public Const WM_MOUSEMOVE = &H200
Public Const WM_RBUTTONUP = &H205
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_LBUTTONDBLCLK = &H203

Global TrayIcon As NOTIFYICONDATA

Public Sub SetTrayIcon(frm As Form, strToolTip As String, Icon, iAction As Integer)
On Error Resume Next
With TrayIcon
 .cbSize = Len(TrayIcon)
 .hwnd = frm.hwnd
 .szTip = strToolTip & vbNullChar
 .hIcon = Icon
 .uID = vbNull
 .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
 .uCallbackMessage = WM_MOUSEMOVE
End With
Select Case iAction
 Case 1 'Show Icon
  Shell_NotifyIcon NIM_ADD, TrayIcon
 Case 2 'Modify Icon
  Shell_NotifyIcon NIM_MODIFY, TrayIcon
End Select
End Sub

Public Sub RemoveTrayIcon()
 Shell_NotifyIcon NIM_DELETE, TrayIcon
End Sub
