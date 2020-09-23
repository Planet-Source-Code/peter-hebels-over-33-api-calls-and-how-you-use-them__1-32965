VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "WritePrivateProfileString"
   ClientHeight    =   3885
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   3270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   120
      TabIndex        =   7
      Top             =   360
      Width           =   3015
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Text            =   "TestString"
      Top             =   2520
      Width           =   3015
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Text            =   "TestKey"
      Top             =   1800
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "TestSection"
      Top             =   1080
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Write INI file"
      Height          =   735
      Left            =   480
      TabIndex        =   0
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Label Label4 
      Caption         =   "File name:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "Key String:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Key Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Section to write key to:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'WritePrivateProfileString API call project by Peter Hebels, Website "www.phsoft.cjb.net" *
'Iam not responsible for any damages may caused by this project                           *
'******************************************************************************************

Dim TxtSection As String
Dim TxtKeyName As String
Dim TxtKeyString As String
Dim TxtFileName As String

Private Sub Command1_Click()
'Things to write to the ini file
TxtFileName = Text4.Text
TxtSection = Text1.Text
TxtKeyName = Text2.Text
TxtKeyString = Text3.Text

'Look if all the information is entered
If TxtFileName = "" Then GoTo ErrHand
If TxtSection = "" Then GoTo ErrHand
If TxtKeyName = "" Then GoTo ErrHand
If TxtKeyString = "" Then GoTo ErrHand

'Write the file
WritePrivateProfileString TxtSection, TxtKeyName, TxtKeyString, TxtFileName

'Show a message
MsgBox "INI file written", vbInformation, "Done"
Exit Sub

'Error handler
ErrHand:
MsgBox "Information not complete!", vbInformation, "Error"
End Sub

Private Sub Form_Load()
'Add a filename to the textbox
Text4.Text = App.Path & "\Test.ini"
End Sub
