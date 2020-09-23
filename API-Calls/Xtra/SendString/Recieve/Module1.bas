Attribute VB_Name = "Module1"
Type COPYDATASTRUCT
     dwData As Long
     cbData As Long
     lpData As Long
End Type

Public Const GWL_WNDPROC = (-4)
Public Const WM_COPYDATA = &H4A
Global lpPrevWndProc As Long
Global gHW As Long

      'Copies a block of memory from one location to another.
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
         (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

Declare Function CallWindowProc Lib "user32" Alias _
         "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As _
         Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As _
         Long) As Long

Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
         (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As _
         Long) As Long

Public Sub Hook()
   lpPrevWndProc = SetWindowLong(gHW, GWL_WNDPROC, _
   AddressOf WindowProc)
End Sub

Public Sub Unhook()
   Dim temp As Long
   temp = SetWindowLong(gHW, GWL_WNDPROC, lpPrevWndProc)
End Sub

Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, _
   ByVal wParam As Long, ByVal lParam As Long) As Long
     If uMsg = WM_COPYDATA Then
     Call mySub(lParam)
     End If
   WindowProc = CallWindowProc(lpPrevWndProc, hw, uMsg, wParam, lParam)
End Function

Sub mySub(lParam As Long)
          Dim cds As COPYDATASTRUCT
          Dim buf(1 To 255) As Byte

          Call CopyMemory(cds, ByVal lParam, Len(cds))

          Select Case cds.dwData
           Case 1
              
           Case 2
              
           Case 3
              Call CopyMemory(buf(1), ByVal cds.lpData, cds.cbData)
              a$ = StrConv(buf, vbUnicode)
              a$ = Left$(a$, InStr(1, a$, Chr$(0)) - 1)
              Form1.Label1.Caption = a$
          End Select
End Sub


