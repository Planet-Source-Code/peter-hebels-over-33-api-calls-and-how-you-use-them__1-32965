VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SetTimer"
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4815
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Timer3"
      Height          =   1215
      Left            =   120
      TabIndex        =   8
      Top             =   2880
      Width           =   4575
      Begin VB.CommandButton Command6 
         Caption         =   "Stop"
         Height          =   495
         Left            =   2280
         TabIndex        =   11
         Top             =   600
         Width           =   1935
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Start"
         Height          =   495
         Left            =   360
         TabIndex        =   10
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   360
         TabIndex        =   9
         Text            =   "0"
         Top             =   240
         Width           =   3855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Timer2"
      Height          =   1335
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   4575
      Begin VB.CommandButton Command4 
         Caption         =   "Stop"
         Height          =   495
         Left            =   2280
         TabIndex        =   7
         Top             =   720
         Width           =   1935
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Start"
         Height          =   495
         Left            =   360
         TabIndex        =   6
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   360
         TabIndex        =   5
         Text            =   "0"
         Top             =   240
         Width           =   3855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Timer1"
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4575
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   360
         TabIndex        =   3
         Text            =   "0"
         Top             =   240
         Width           =   3855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Start"
         Height          =   495
         Left            =   360
         TabIndex        =   2
         Top             =   720
         Width           =   1815
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Stop"
         Height          =   495
         Left            =   2280
         TabIndex        =   1
         Top             =   720
         Width           =   1935
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'SetTimer API-Call project by Peter Hebels, Website "www.phsoft.cjb.net"                  *
'Iam not responsible for any damages may caused by this project                           *
'******************************************************************************************

'This project shows you how to use timers without timer controls!
'WARNING this project uses some nasty functions like "AddressOf" Don't use them wrong
'Or VB will crash with the result of losing work in progress :(!!
'Also don't forget to shutdown started timers!

Dim Timer1 As Long            'Declare your timer ID's
Dim Timer2 As Long
Dim Timer3 As Long

'----------------------------------------------------------------------
'Timer1 function
Function StartTimer1()
Timer1 = KillTimer(0, Timer1)
Timer1 = SetTimer(0, 0, 100, AddressOf TimerSub1) 'This calls the module sub named "TimerSub1",
End Function                                      'The "AddressOf" function can only call to modules

Function StopTimer1()         'This sub will stop the timer
Timer1 = KillTimer(0, Timer1) 'Stop the timer, always do this before closing your app otherwise
End Function                  'VB will crash!!

'Timer2 function
Function StartTimer2()        'This sub will start the timer
Timer2 = KillTimer(0, Timer2) 'First kill a started timer because if a timer is
                              'started twice you can't stop it anymore with will result in a crash!!
Timer2 = SetTimer(0, 0, 100, AddressOf TimerSub2) 'This calls the module sub named "TimerSub2",
'/|\                    /|\
'Timer ID               This number is the interval
End Function

Function StopTimer2()
Timer2 = KillTimer(0, Timer2)
End Function

'Timer3 function
Function StartTimer3()
Timer3 = KillTimer(0, Timer3)
Timer3 = SetTimer(0, 0, 100, AddressOf TimerSub3) 'This calls the module sub named "TimerSub3",
'/|\                    /|\
'Timer ID               This number is the interval
End Function

Function StopTimer3()
Timer3 = KillTimer(0, Timer3)
End Function
'----------------------------------------------------------------------



Public Function TimedSub1() 'This sub is called from the module
Form1.Text1.Text = Form1.Text1.Text + 1 'just count as you would do by a normal timer control
End Function

Public Function TimedSub2() 'This sub is called from the module
Form1.Text2.Text = Form1.Text2.Text + 1
End Function

Public Function TimedSub3() 'This sub is called from the module
Form1.Text3.Text = Form1.Text3.Text + 1
End Function


'The start and stop buttons
Private Sub Command1_Click()
StartTimer1 'Start timer1
End Sub

Private Sub Command2_Click()
StopTimer1 'stop timer1
End Sub

Private Sub Command3_Click()
StartTimer2 'Start timer2
End Sub

Private Sub Command4_Click()
StopTimer2 'stop timer2
End Sub

Private Sub Command5_Click()
StartTimer3 'Start timer3
End Sub

Private Sub Command6_Click()
StopTimer3 'Stop timer3
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Don't forget to kill the timers!!
StopTimer1
StopTimer2
StopTimer3
End Sub
