VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmServer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Insomnia Game Server"
   ClientHeight    =   4620
   ClientLeft      =   3330
   ClientTop       =   2460
   ClientWidth     =   5820
   Icon            =   "frmServer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   308
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   388
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   1  'Minimized
   Begin VB.TextBox txtEvents 
      Height          =   2040
      Left            =   75
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   75
      Width           =   5640
   End
   Begin MSWinsockLib.Winsock WinSock 
      Index           =   0
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      LocalPort       =   6767
   End
   Begin VB.TextBox txtSend 
      Height          =   330
      Left            =   75
      TabIndex        =   1
      Top             =   4200
      Width           =   5640
   End
   Begin VB.TextBox txtRec 
      Height          =   1890
      Left            =   75
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   2325
      Width           =   5640
   End
   Begin VB.Menu mPopupSys 
      Caption         =   "&SysTray"
      Visible         =   0   'False
      Begin VB.Menu mPopRestore 
         Caption         =   "&Restore"
      End
      Begin VB.Menu mPopExit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public HR
Dim KillApp As Boolean

Const vbKeyReturn = 13 '<Enter>

'**********************************************
' ***************** Initializations ******************
'**********************************************

Private Sub Initializations()
Dim SockCreate As Integer
Dim ErrHandler
Dim Dummy

On Error GoTo ErrHandler
KillApp = False

'[String and Player Variables]
HR = Chr(13) + Chr(10) 'Hard Return
'MAXPLAYERS = 10

Dim R
R = PutFocus(txtSend.hWnd)

'[Display Server's IP Address]
txtRec.Text = txtRec.Text + frmServer.Caption + " initiated. (" + WinSock(0).LocalIP + ")"
frmServer.Caption = frmServer.Caption + " - [ " + WinSock(0).LocalIP + " ]"

'[Load Sockets]
WinSock(0).Listen 'Socket is listening
For SockCreate = 1 To MAXPLAYERS
    Load WinSock(SockCreate) 'Load a new socket with +1 index number
Next
Exit Sub 'Avoid ErrHandler

ErrHandler:
KillApp = True
MsgBox "Server is already running. Ctrl+Alt+Del to find it.", vbOKOnly, "Second Server Not Allowed"

End Sub

Private Sub Tray_Init()

'The form must be fully visible before calling Shell_NotifyIcon
Me.Show
Me.Refresh

With nid
    .cbSize = Len(nid)
    .hWnd = Me.hWnd
    .uId = vbNull
    .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    .uCallBackMessage = WM_MOUSEMOVE
    .hIcon = Me.Icon 'Change to any icon in project for different icons
    .szTip = frmServer.Caption & vbNullChar
End With

Shell_NotifyIcon NIM_ADD, nid

End Sub

'**********************************************
' ***************** System Tray ******************
'**********************************************

Private Sub Form_Click()
Dim R
R = PutFocus(txtSend.hWnd)
End Sub

Private Sub Form_Load()

Call Initializations
Call Tray_Init

If KillApp = True Then: Unload Me

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'This procedure receives the callbacks from the System Tray icon.

Dim Result As Long
Dim Msg As Long

'The value of X will vary depending upon the scalemode setting
If Me.ScaleMode = vbPixels Then
    Msg = X
Else
    Msg = X / Screen.TwipsPerPixelX
End If

Select Case Msg
    'Case WM_LBUTTONUP        '514 restore form window
    '    Me.WindowState = vbNormal
    '    Result = SetForegroundWindow(Me.hwnd)
    '    Me.Show
    Case WM_LBUTTONDBLCLK    '515 restore form window
        Me.WindowState = vbNormal
        Result = SetForegroundWindow(Me.hWnd)
        Me.Show
    Case WM_RBUTTONUP
        '517 display popup menu
        Result = SetForegroundWindow(Me.hWnd)
        Me.PopupMenu Me.mPopupSys, vbPopupMenuRightAlign
End Select

End Sub

Private Sub Form_Resize()
    'This is necessary to assure that the minimized window is hidden
     If Me.WindowState = vbMinimized Then Me.Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim Msg As String, Response, EndServer
Dim NumPlayers As Integer, PCountLoop As Integer

'Dim AvoidUnloads
If KillApp = True Then: Exit Sub 'GoTo AvoidUnloads

NumPlayers = 0
For PCountLoop = 1 To MAXPLAYERS
    If WinSock(PCountLoop).State = 7 Then: NumPlayers = NumPlayers + 1
Next PCountLoop

Msg = "If you close this server, the" + Str(NumPlayers) + " players currently" + HR
Msg = Msg + "connected will be not be able to play." + HR + HR
Msg = Msg + "Would you like to keep the server running?"
Response = MsgBox(Msg, vbExclamation + vbYesNo, "Authorization Required")
Select Case Response
    Case vbYes   ' Don't allow close.
        Cancel = -1: Exit Sub
    Case vbNo
        GoTo EndServer
End Select

EndServer:
Shell_NotifyIcon NIM_DELETE, nid 'Removes icon from the system tray

'AvoidUnloads:

End Sub

Private Sub mPopExit_Click()
    Unload Me
End Sub
      
Private Sub mPopRestore_Click()
    Dim Result
    Me.WindowState = vbNormal
    Result = SetForegroundWindow(Me.hWnd)
    Me.Show
End Sub

Private Sub txtEvents_Change()
    If frmServer.WindowState = 0 Then
        SavedWnd = Screen.ActiveControl.hWnd
    End If
    Dim Num
    Num = ScrollText(txtEvents, 16000)
End Sub

Private Sub txtRec_Change()
    If frmServer.WindowState = 0 Then
        SavedWnd = Screen.ActiveControl.hWnd
    End If
    Dim Num
    Num = ScrollText(txtRec, 16000)
End Sub

'**********************************************
'******************** TCP/IP ********************
'**********************************************

Private Sub txtSend_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Then
    
    txtRec.Text = txtRec.Text + HR + "<Server> " + txtSend.Text
    Call SocketSend("Text", 0, "<Server> " + txtSend.Text, False, 0)
    txtSend.Text = ""
    
End If

End Sub

Private Sub txtSend_KeyPress(KeyAscii As Integer)

Select Case KeyAscii
    Case 167
        KeyAscii = False 'Stop the key from being presseed for §'s
    Case 13
        KeyAscii = 0 'Stop <enter>'s from beeping
End Select

End Sub

Private Sub WinSock_Close(Index As Integer)

WinSock(Index).Close
txtRec.Text = txtRec.Text + HR + Player(Index).Name + " disconnected from socket " + Str(Index) + "."

Call SocketSend("Text", Str(Index), "hears another realm beckoning....and leaves this world.", False, 0)

End Sub

Private Sub Winsock_ConnectionRequest(Index As Integer, ByVal requestID As Long)

Dim Temp As Integer

For Temp = 1 To MAXPLAYERS
    If WinSock(Temp).State <> sckConnected Then 'And WinSock(Temp).State <> 7
        
        WinSock(Temp).Accept requestID 'Accept requestID
        Exit For
    
    End If
Next Temp

End Sub

Private Sub WinSock_DataArrival(Index As Integer, ByVal bytesTotal As Long)

Static DataRead As String
Dim LineString As String
Dim CommandString As String

WinSock(Index).GetData DataRead

While InStr(DataRead, "§") 'Make sure the EOF was recieved

    CommandString = Left$(DataRead, InStr(DataRead, " ") - 1) 'Get command
    LineString = Mid$(DataRead, InStr(DataRead, " ") + 1) 'Get data w/ EOF
    LineString = Left$(LineString, InStr(LineString, "§") - 1) 'Remove EOF
    
    Select Case CommandString
    
        '[Communication and Movement]
        Case "Chat"

            txtRec.Text = txtRec.Text + HR + Player(Index).Name + ": " + LineString
            
            Call SocketSend("Chat", Index, LineString, False, 0)
            
        Case "Emote"
        
        Case "Move"
        
            Select Case LineString
                Case "D": Player(Index).Y = Player(Index).Y + 1
                Case "U": Player(Index).Y = Player(Index).Y - 1
                Case "R": Player(Index).X = Player(Index).X + 1
                Case "L": Player(Index).X = Player(Index).X - 1
            End Select
            
            
            Call SocketSend("Move", Index, LineString, False, Index)
            txtEvents.Text = txtEvents.Text + HR + Player(Index).Name + " MoveTo- [ " + Str(Player(Index).X) + "," + Str(Player(Index).Y) + " ]"
    
        '[Player Initializations]
        Case "InitPosition"

            Player(Index).X = Int(Val(Left(LineString, InStr(LineString, " ") - 1)))
            Player(Index).Y = Int(Val(Mid(LineString, InStr(LineString, " ") + 1)))
            
            Call SocketSend("InitPosition", Index, Str(Player(Index).X) + Str(Player(Index).Y), True, Index)
            
        Case "InitImage"
        
            Player(Index).SpriteNum = Int(Val(LineString))
            Call SocketSend("InitImage", Index, Str(Player(Index).SpriteNum), True, Index)

        Case "Name"
        
            Player(Index).Name = LineString
            txtRec.Text = txtRec.Text + HR + Player(Index).Name + " connected on socket" + Str(Index) + "."
            
            Dim InitLoop As Integer
            For InitLoop = 1 To MAXPLAYERS
            
                If WinSock(InitLoop).State = 7 And (InitLoop <> Index) Then
                    Dim SendPos As String, SendImg As String, SendName As String
                    
                    '[Send Position]
                    SendPos = "InitPosition" + Str(InitLoop) + Str(Player(InitLoop).X) + Str(Player(InitLoop).Y) + "§"
                    DoEvents: WinSock(Index).SendData SendPos: DoEvents
                    '[Send Image Number]
                    SendImg = "InitImage" + Str(InitLoop) + Str(Player(InitLoop).SpriteNum) + "§"
                    DoEvents: WinSock(Index).SendData SendImg: DoEvents
                    '[Send Name]
                    SendName = "Name" + Str(InitLoop) + " " + Player(InitLoop).Name + "§"
                    DoEvents: WinSock(Index).SendData SendName: DoEvents
                End If
                
            Next InitLoop
            
            Dim SendStart
            SendStart = "Start" + Str(Index) + " Welcome to Insomnia, " + Player(Index).Name + "!§"
            DoEvents
            WinSock(Index).SendData SendStart
            DoEvents
            
            Call SocketSend("Name", Index, Player(Index).Name, False, Index)
            
            txtEvents.Text = txtEvents.Text + HR + Player(Index).Name + " Init- [" + Str(Player(Index).X) + "," + Str(Player(Index).Y) + " ]"

        '[Enemies and Attacking]
        Case "Attack"

        '[Accidental Non-Legal Command(s)]
        Case Else
        
            txtRec.Text = txtRec.Text + HR + "Illigal Command attempt by " + Player(Index).Name + " [" + CommandString + "]"
    
    End Select

    DataRead = Mid$(DataRead, InStr(DataRead, "§") + 1) 'Clear EOF + data

Wend

Dim R
R = PutFocus(txtSend.hWnd)
    
End Sub

Private Sub SocketSend(Command As String, ClientIndex As Integer, Info As String, InfoIsNum As Boolean, ClientOmit As Integer)

Dim SendLoop As Integer
Dim DataToSend As String

If InfoIsNum = True Then 'Info already has a preceding space
    DataToSend = Command + Str(ClientIndex) + Info + "§"
Else 'Info needs a space before it
    DataToSend = Command + Str(ClientIndex) + " " + Info + "§"
End If

For SendLoop = 1 To MAXPLAYERS 'Check conn. players to send to
    If WinSock(SendLoop).State = 7 And (SendLoop <> ClientOmit) Then
        DoEvents
        WinSock(SendLoop).SendData DataToSend 'Send the string
        DoEvents
    End If
Next

End Sub

Private Sub WinSock_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

Dim Msg As String

Msg = Msg + "WINSOCK ERROR!!! - From: " + Player(Index).Name + HR
Msg = Msg + "   Error Number: " + Str(Number) + HR
Msg = Msg + "   Description: " + Description + HR
Msg = Msg + "   S-Code: " + Str(Scode) + HR
Msg = Msg + "   Source: " + Source + HR
Msg = Msg + "   Help File: " + HelpFile + HR + HR
txtRec.Text = txtRec.Text + HR + HR + Msg

End Sub
