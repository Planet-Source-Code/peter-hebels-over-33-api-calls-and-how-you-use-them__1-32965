VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "AppendMenu"
   ClientHeight    =   840
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5550
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   840
   ScaleWidth      =   5550
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   855
      Left            =   4080
      TabIndex        =   0
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Right-Clik on my icon in the taskbar.."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'AppendMenu API call project by Peter Hebels, Website "www.phsoft.cjb.net"                *
'Iam not responsible for any damages may caused by this project                           *
'******************************************************************************************

Const IDM_ABOUT = 1010
Dim hSysMenu As Long

Private Sub Command2_Click()
'Close the app
Unload Me
End
End Sub

Private Sub Form_Load()
'Get the hwnd of the System menu
hSysMenu = GetSystemMenu(hwnd, 0&)

'API calls to add some items to the systemmenu.
'Note: I don't know how to use the messages sended from the menu!
Call AppendMenu(hSysMenu, MF_SEPARATOR, 0&, 0&)
Call AppendMenu(hSysMenu, MF_STRING, IDM_ABOUT, "About...")
Call AppendMenu(hSysMenu, MF_STRING, IDM_TOMENU, "Added to menu...")

End Sub
