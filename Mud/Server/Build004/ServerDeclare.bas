Attribute VB_Name = "ServerDeclare"
Option Explicit

Global Const MAXPLAYERS = 10

Type PlayerInfo
    Name As String
    SpriteNum As Integer
    
    Direction As Integer
    Busy As Boolean
    SprX As Integer
    
    X As Single
    Y As Single
    Map As Integer
End Type

Public Player(1 To MAXPLAYERS) As PlayerInfo

'FUNCTIONS: Strings********************************************
Declare Function PutFocus Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long

Public SavedWnd As Long

'********************* System Tray API **********************
'user defined type required by Shell_NotifyIcon API call
Public Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uId As Long
    uFlags As Long
    uCallBackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

'constants required by Shell_NotifyIcon API call:
Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201     'Button down
Public Const WM_LBUTTONUP = &H202       'Button up
Public Const WM_LBUTTONDBLCLK = &H203   'Double-click
Public Const WM_RBUTTONDOWN = &H204     'Button down
Public Const WM_RBUTTONUP = &H205       'Button up
Public Const WM_RBUTTONDBLCLK = &H206   'Double-click

Public Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

Public nid As NOTIFYICONDATA

Public Const vbPopupMenuLeftAlign = 0
Public Const vbPopupMenuCenterAlig = 4
Public Const vbPopupMenuRightAlign = 8
'********************************************************

Function ScrollText(TextBox As Control, vLines As Integer)

Dim Success As Long

Dim R As Long
Const EM_LINESCROLL = &HB6

Dim Lines

' Get the window handle of the control that currently has the
'  focus, Command1 or Command2.
'SavedWnd = Screen.ActiveControl.hwnd
Lines = vLines

' Set the focus to the passed control (text control).
R = PutFocus(TextBox.hWnd)       ' Scroll the lines.
Success = SendMessage(TextBox.hWnd, EM_LINESCROLL, 0, Lines)

' Restore the focus to the original control, Command1 or
'  Command2.
R = PutFocus(SavedWnd)

' Return the number of lines actually scrolled.
ScrollText = Success

End Function
