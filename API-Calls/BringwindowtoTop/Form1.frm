VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BringWindowToTop"
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   3135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Bring me to top"
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   2895
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Bring Form2 to top"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'BringWindowToTop API call project by Peter Hebels, Website "www.phsoft.cjb.net"          *
'Iam not responsible for any damages may caused by this project                           *
'******************************************************************************************

Private Sub Command1_Click()
'Bring form2 to top
BringWindowToTop Form2.hwnd
End Sub

Private Sub Command2_Click()
'Bring form1 to top
BringWindowToTop Form1.hwnd
End Sub

Private Sub Command3_Click()
'End the app
Unload Form2
Unload Me
End
End Sub

Private Sub Form_Load()
'Show form2
Form2.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
'End the app
Unload Form2
Unload Me
End
End Sub
