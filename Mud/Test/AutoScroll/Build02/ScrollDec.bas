Attribute VB_Name = "ScrollDeclare"
Option Explicit

Declare Function PutFocus Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long

Function ScrollText&(TextBox As Control, vLines As Integer)

Dim Success As Long
Dim SavedWnd As Long
Dim R As Long
Const EM_LINESCROLL = &HB6

Dim Lines&

' Get the window handle of the control that currently has the
'  focus, Command1 or Command2.
SavedWnd = Screen.ActiveControl.hwnd
Lines& = vLines

' Set the focus to the passed control (text control).
TextBox.SetFocus       ' Scroll the lines.
Success = SendMessage(TextBox.hwnd, EM_LINESCROLL, 0, Lines&)

' Restore the focus to the original control, Command1 or
'  Command2.       R = PutFocus(SavedWnd)
' Return the number of lines actually scrolled.
ScrollText& = Success

End Function
