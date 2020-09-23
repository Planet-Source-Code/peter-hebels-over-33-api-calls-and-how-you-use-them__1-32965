VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BitBlt"
   ClientHeight    =   2760
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2085
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   2085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   825
      Index           =   1
      Left            =   600
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   51
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   825
      Index           =   0
      Left            =   0
      Picture         =   "Form1.frx":0C3A
      ScaleHeight     =   51
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.PictureBox picture1 
      BackColor       =   &H80000003&
      Height          =   1095
      Left            =   360
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   85
      TabIndex        =   2
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   615
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "BitBlt"
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'BitBlt API call project by Peter Hebels, Website "www.phsoft.cjb.net"                    *
'Iam not responsible for any damages may caused by this project                           *
'******************************************************************************************


Private Sub Command1_Click()
Dim XP, YP, HP, WP As Long

'Variables for positioning the picture
XP = picture1.ScaleWidth / 4
YP = picture1.ScaleHeight / 8
WP = Picture3(1).ScaleWidth
HP = Picture3(1).ScaleHeight

BitBlt picture1.hDC, XP, YP, WP, HP, Picture3(1).hDC, 0, 0, MERGEPAINT
BitBlt picture1.hDC, XP, YP, WP, HP, Picture3(0).hDC, 0, 0, SRCAND
'BitBlt picture1.hDC, XP = The Left position, YP = the top position , WP = width, HP = height, Picture3(1).hDC, 0, 0, MERGEPAINT
'BitBlt picture1.hDC, XP = The left position of the mask, YP = the top position of the mask, WP, HP, Picture3(0).hDC, 0, 0, SRCAND
End Sub

Private Sub Command2_Click()
'Close the app
Unload Me
End
End Sub
