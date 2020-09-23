VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Shell_NotifyIcon"
   ClientHeight    =   45
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   45
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu Mnutray 
      Caption         =   "TrayMen"
      Begin VB.Menu MnuAbout 
         Caption         =   "About"
      End
      Begin VB.Menu MnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'Shell_NotifyIcon API-Call project by Peter Hebels, Website "www.phsoft.cjb.net"          *
'Iam not responsible for any damages may caused by this project                           *
'******************************************************************************************

Private Sub Form_Load()
'Add the trayicon and hide the form
Totray.Into_Tray Me.hWnd, Me.Icon, "Tray Example"
Me.Hide
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Look where the mouse pointer is, and if right click then show the menu
    Msg = X / Screen.TwipsPerPixelX
    If Msg = WM_RBUTTONUP Then
        Me.PopupMenu Form1.Mnutray
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Delete the tray icon
Totray.Del_TrayIcon Me.hWnd
End Sub

Private Sub MnuAbout_Click()
'MnuAbout is clicked
MsgBox "This project is created by Peter Hebels", vbInformation, "About"
End Sub

Private Sub MnuExit_Click()
'MnuExit is clicked
Unload Me
End
End Sub
