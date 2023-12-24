Attribute VB_Name = "Declares"
Option Explicit

Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Global Const EM_LINESCROLL = &H406
Global Const EM_SCROLL = &HB5 'test

Declare Function PutFocus% Lib "user" Alias "SetFocus" (ByVal hWd%)

Function ScrollText&(TextBox As Control, vLines As Integer, hLines As Integer)
Dim Lines&, SavedWnd%, Success&, r%

    'Const EM_LINESCROLL = &H406

    ' Place the number of horizontal columns to scroll in the high-
    ' order 2 bytes of Lines&. The vertical lines to scroll is
    ' placed in the low-order 2 bytes.
    Lines& = CLng(&H10000 * hLines) + vLines

    ' Get the window handle of the control that currently has the focus
    SavedWnd% = Screen.ActiveControl.hwnd

    ' Set the focus to the passed control (text control).
    TextBox.SetFocus

    ' Scroll the lines.
    Success& = SendMessage(TextBox.hwnd, EM_LINESCROLL, 0, Lines&)

    'txtSend.SetFocus

    ' Return the number of lines actually scrolled.
    ScrollText& = Success&

End Function
