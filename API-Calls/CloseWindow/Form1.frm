VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CloseWindow"
   ClientHeight    =   1110
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   1935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1110
   ScaleWidth      =   1935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Close window"
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'CloseWindow API call project by Peter Hebels, Website "www.phsoft.cjb.net"               *
'Iam not responsible for any damages may caused by this project                           *
'******************************************************************************************

Private Sub Command1_Click()
'Close Form1
CloseWindow Me.hwnd
End Sub

Private Sub Form_Unload(Cancel As Integer)
'End the app
Unload Me
End
End Sub
