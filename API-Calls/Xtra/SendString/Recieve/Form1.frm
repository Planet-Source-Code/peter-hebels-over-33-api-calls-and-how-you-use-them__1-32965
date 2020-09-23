VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Testapp"
   ClientHeight    =   1830
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4695
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1830
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Data send from other app:"
      Height          =   1215
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   4695
      Begin VB.Label Label1 
         Caption         =   """"""
         Height          =   855
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4455
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'Send string project by Peter Hebels, Website "www.phsoft.cjb.net"                        *
'Iam not responsible for any damages may caused by this project                           *
'******************************************************************************************
      
'To test this project you have to run both project's from the 'Send' And
''Recieve' directory's

'This is the Recieveing project

Private Sub Command1_Click()
  Unload Form1
End Sub

Private Sub Form_Load()
  gHW = Me.hwnd
  Hook
  Me.Caption = "Testapp"
  Me.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Unhook
  Unload Form1
End Sub

