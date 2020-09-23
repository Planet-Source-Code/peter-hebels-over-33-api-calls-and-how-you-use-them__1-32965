VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6555
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   6555
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   615
      Left            =   4800
      TabIndex        =   5
      Top             =   1680
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   6255
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   6255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Copy"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "From:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "To:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'CopyFile API call project by Peter Hebels, Website "www.phsoft.cjb.net"                  *
'Iam not responsible for any damages may caused by this project                           *
'******************************************************************************************

Dim CopyFrom As String
Dim CopyTo As String


Function CopyAFile(ExistingFile As String, NewFile As String, IfFileExists As Long)
CopyFile ExistingFile, NewFile, IfFileExists
End Function

Private Sub Command1_Click()
CopyFrom = Text1.Text
CopyTo = Text2.Text

CopyDone = CopyAFile(CopyFrom, CopyTo, 1)
MsgBox "File copied from: " & CopyFrom & vbCrLf & "to: " & CopyTo
End Sub

Private Sub Command2_Click()
Unload Me
End
End Sub

Private Sub Form_Load()
Text1.Text = App.Path & "\Project1.vbp"
Text2.Text = App.Path & "\NewProject1.vbp"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
End
End Sub

Private Sub Text1_Change()
If Text1.Text = "" Then Command1.Enabled = False
End Sub

Private Sub Text2_Change()
If Text2.Text = "" Then Command1.Enabled = False
End Sub
