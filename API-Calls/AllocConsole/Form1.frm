VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "AllocConsole"
   ClientHeight    =   1965
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   2295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Open Console"
      Height          =   735
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Write to it"
      Enabled         =   0   'False
      Height          =   735
      Left            =   360
      TabIndex        =   0
      Top             =   1080
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'AllocConsole API call project by Peter Hebels, Website "www.phsoft.cjb.net"              *
'Iam not responsible for any damages may caused by this project                           *
'******************************************************************************************

Dim hConsole As Long 'Console window handler

Private Sub Command1_Click()
   Dim Result As Long 'Errorchecking
   Dim sOut As String 'Console output handler
   Dim cWritten As Long 'How many chars are written
     
     sOut = "Welcome to ConsoleTest" & vbCrLf 'Write to the console
     Result = WriteConsole(hConsole, ByVal sOut, Len(sOut), cWritten, ByVal 0&)
     Shell App.Path & "\Dir.bat" 'Execute (Dir.bat)
   End Sub

Private Sub Command2_Click()
     'Look if the console is created
     If AllocConsole() Then
       hConsole = GetStdHandle(STD_OUTPUT_HANDLE)
       If hConsole = 0 Then MsgBox "Couldn't allocate STDOUT"
     Else
       MsgBox "Couldn't allocate console"
     End If

Command1.Enabled = True
End Sub

   Private Sub Form_Unload(Cancel As Integer)
     'close the console
     CloseHandle hConsole
     FreeConsole
     'And end the app
     Unload Me
     End
   End Sub


