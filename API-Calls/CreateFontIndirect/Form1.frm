VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CreateFontIndirect"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Spin font"
      Height          =   615
      Left            =   2040
      TabIndex        =   3
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   360
      Top             =   240
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1935
      LargeChange     =   100
      Left            =   3840
      TabIndex        =   2
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Draw Font"
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   2280
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   1935
      Left            =   240
      ScaleHeight     =   1875
      ScaleWidth      =   3435
      TabIndex        =   0
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'CreateFontIndirect API call project by Peter Hebels, Website "www.phsoft.cjb.net"        *
'Iam not responsible for any damages may caused by this project                           *
'******************************************************************************************
Dim I As Integer

Sub Command1_Click()
   Dim font As LOGFONT
   Dim prevFont As Long, hFont As Long, ret As Long
   Const FONTSIZE = 10 'Desired point size of font
     Picture1.Cls
     font.lfEscapement = 1800    '180-degree rotation
     font.lfFaceName = "Arial" & Chr$(0) 'Null character at end
     'Windows expects the font size to be in pixels and to
     'be negative if you are specifying the character height
     'you want.
     font.lfHeight = (FONTSIZE * -20) / Screen.TwipsPerPixelY
     hFont = CreateFontIndirect(font)
     prevFont = SelectObject(Picture1.hdc, hFont)
     Picture1.CurrentX = Picture1.Left + Picture1.Width / 2
     Picture1.CurrentY = Picture1.ScaleHeight / 2
     ret = SelectObject(Picture1.hdc, prevFont)
     ret = DeleteObject(hFont)
     Picture1.CurrentY = Picture1.ScaleHeight / 2
     'Print the font to the picturebox
     Picture1.Print "Normal Text"
End Sub

Private Sub Command2_Click()
If Timer1.Enabled = False Then
  Timer1.Enabled = True
  Command2.Caption = "Stop spin"
  Command1.Enabled = False
  VScroll1.Enabled = False

Else

  Timer1.Enabled = False
  Command2.Caption = "Spin font"
  Command1.Enabled = True
  VScroll1.Enabled = True
End If
End Sub

Private Sub Form_Load()
  VScroll1.Max = 1800 * 2
End Sub

Private Sub Timer1_Timer()
Dim font As LOGFONT
Dim prevFont As Long, hFont As Long, ret As Long

I = I + 10

If I >= 1800 * 2 Then I = 0

Const FONTSIZE = 10 'Desired point size of font
     Picture1.Cls
     font.lfEscapement = I    'Rotate the text automaticly
     font.lfFaceName = "Arial" & Chr$(0) 'Null character at end

     font.lfHeight = (FONTSIZE * -20) / Screen.TwipsPerPixelY
     hFont = CreateFontIndirect(font)
     prevFont = SelectObject(Picture1.hdc, hFont)
     Picture1.CurrentX = Picture1.Left + Picture1.Width / 2
     Picture1.CurrentY = Picture1.ScaleHeight / 2
     'Print the font to the picturebox
     Picture1.Print "Rotated Text"
End Sub

Private Sub VScroll1_Change()
Dim font As LOGFONT
Dim prevFont As Long, hFont As Long, ret As Long
Const FONTSIZE = 10 'Desired point size of font
     'Clear the picturebox
     Picture1.Cls
     font.lfEscapement = VScroll1.Value  'Rotate the text depending on Vscroll1
     font.lfFaceName = "Arial" & Chr$(0) 'Null character at end

     font.lfHeight = (FONTSIZE * -20) / Screen.TwipsPerPixelY
     hFont = CreateFontIndirect(font)
     prevFont = SelectObject(Picture1.hdc, hFont)
     Picture1.CurrentX = Picture1.Left + Picture1.Width / 2
     Picture1.CurrentY = Picture1.ScaleHeight / 2
     'Print the font to the picturebox
     Picture1.Print "Rotated Text"
End Sub
