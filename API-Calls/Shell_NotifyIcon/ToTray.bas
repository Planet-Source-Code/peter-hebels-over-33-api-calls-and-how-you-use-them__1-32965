Attribute VB_Name = "Totray"
Private Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Dim TrayI As NOTIFYICONDATA

Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4

Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_RBUTTONUP = &H205

Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

'Add to tray
Sub Into_Tray(hWnd As Long, picture As IPictureDisp, traytip As String)
 TrayI.cbSize = Len(TrayI)
 TrayI.hWnd = hWnd
 TrayI.uId = 1&
 TrayI.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
 TrayI.ucallbackMessage = WM_LBUTTONDOWN
 TrayI.hIcon = picture
 TrayI.szTip = traytip & Chr$(0)
 Shell_NotifyIcon NIM_ADD, TrayI
End Sub

'Delete the trayicon
Sub Del_TrayIcon(hWnd As Long)
 TrayI.cbSize = Len(TrayI)
 TrayI.hWnd = hWnd
 TrayI.uId = 1&
 Shell_NotifyIcon NIM_DELETE, TrayI
End Sub

'The Tooltip for the trayicon
Sub Modify_TrayIcon(traytip As String, picture As IPictureDisp)
 TrayI.hIcon = picture
 TrayI.szTip = traytip & Chr$(0)
 Shell_NotifyIcon NIM_MODIFY, TrayI
End Sub


