Attribute VB_Name = "Module1"
Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
'/|\
' |-----This are the api calls for the timer.

'this are the subs called from the timer
Sub TimerSub1()
Form1.TimedSub1 'you can only call to a form object this way because of the "AddressOf" function.
End Sub

Sub TimerSub2()
Form1.TimedSub2 'you can only call to a form object this way because of the "AddressOf" function.
End Sub

Sub TimerSub3()
Form1.TimedSub3 'you can only call to a form object this way because of the "AddressOf" function.
End Sub

