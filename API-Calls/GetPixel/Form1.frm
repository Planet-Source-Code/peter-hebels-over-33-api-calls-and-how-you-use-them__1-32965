VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "GetPixel"
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4815
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Color"
      Height          =   1935
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1335
      Begin VB.PictureBox Picture2 
         AutoRedraw      =   -1  'True
         Height          =   615
         Left            =   120
         ScaleHeight     =   555
         ScaleWidth      =   1035
         TabIndex        =   7
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "B="
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "G="
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "R="
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   2520
      Width           =   3135
   End
   Begin VB.Frame Frame1 
      Caption         =   "Go with mousepointer over picture"
      Height          =   2295
      Left            =   1560
      TabIndex        =   0
      Top             =   120
      Width           =   3135
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   1725
         Left            =   360
         Picture         =   "Form1.frx":0000
         ScaleHeight     =   111
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   156
         TabIndex        =   1
         Top             =   360
         Width           =   2400
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'GetPixel API-Call project by Peter Hebels, Website "www.phsoft.cjb.net"                  *
'Iam not responsible for any damages may caused by this project                           *
'******************************************************************************************

Dim PixCol As Long

Private Sub Command1_Click()
'End the app
Unload Me
End
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

'Get the pixel color
PixCol = GetPixel(Picture1.hdc, X, Y)

'Convert to RGB
r = PixCol Mod 256
b = Int(PixCol / 65536)
g = (PixCol - (b * 65536) - r) / 256

'Add to the labels
Label1.Caption = "R=" & r
Label2.Caption = "G=" & g
Label3.Caption = "B=" & b

'Change Picture2's backcolor
Picture2.BackColor = RGB(r, g, b)

End Sub
