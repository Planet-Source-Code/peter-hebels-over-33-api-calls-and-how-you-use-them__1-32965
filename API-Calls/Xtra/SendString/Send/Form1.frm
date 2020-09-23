VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "RunningPlugin"
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5835
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   5835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Stop"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   2535
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Pause"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Play"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Song Info"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "SongInfo"
      Height          =   1455
      Left            =   2760
      TabIndex        =   1
      Top             =   120
      Width           =   2895
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

'This is the sending project

      'Memory copy data structure
      Private Type COPYDATASTRUCT
              dwData As Long
              cbData As Long
              lpData As Long
      End Type

      'Copy memory data
      Private Const WM_COPYDATA = &H4A

      'Find a window on the desktop
      Private Declare Function FindWindow Lib "user32" Alias _
         "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName _
         As String) As Long

      'Used to send messages between app's
      Private Declare Function SendMessage Lib "user32" Alias _
         "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal _
         wParam As Long, lParam As Any) As Long

      'Copies a block of memory from one location to another.
      Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
         (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

      Private Sub SendCommand(TheCom As String)
          Dim cds As COPYDATASTRUCT
          Dim ThWnd As Long
          Dim buf(1 To 255) As Byte

        
     
         ThWnd = FindWindow(vbNullString, "Mp3CoolPlay")
      
         a$ = TheCom
      
         Call CopyMemory(buf(1), ByVal a$, Len(a$))
          cds.dwData = 3
          cds.cbData = Len(a$) + 1
          cds.lpData = VarPtr(buf(1))
      
         i = SendMessage(ThWnd, WM_COPYDATA, Me.hwnd, cds)
      '***********************************************************
           
            
End Sub
     

Private Sub Command1_Click()
SendCommand "getsonginfo"
End Sub

Private Sub Command2_Click()
SendCommand "mpcpplay"
End Sub

Private Sub Command3_Click()
SendCommand "mpcppause"
End Sub

Private Sub Command4_Click()
SendCommand "mpcpstop"
End Sub

Private Sub Form_Load()
  gHW = Me.hwnd
  Hook
  Me.Caption = "RunningPlugin"
  Me.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Unhook
  Unload Form1
End Sub
