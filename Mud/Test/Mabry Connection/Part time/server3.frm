VERSION 5.00
Object = "{7A245FC3-9D37-11CF-840F-444553540000}#5.0#0"; "ASOCK32.OCX"
Begin VB.MDIForm MDIForm1 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000C&
   Caption         =   "Banana Chat"
   ClientHeight    =   6405
   ClientLeft      =   3135
   ClientTop       =   3360
   ClientWidth     =   8340
   Icon            =   "server3.frx":0000
   LinkTopic       =   "MDIForm1"
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1515
      Left            =   0
      ScaleHeight     =   1515
      ScaleWidth      =   8340
      TabIndex        =   0
      Top             =   0
      Width           =   8340
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   7650
         Top             =   780
      End
      Begin VB.TextBox txtLocalAddress 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1335
         TabIndex        =   3
         Text            =   "0.0.0.0"
         Top             =   135
         Width           =   1425
      End
      Begin VB.TextBox txtLocalPort 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3870
         TabIndex        =   2
         Text            =   "1024"
         Top             =   150
         Width           =   555
      End
      Begin VB.ListBox lstStatus 
         Appearance      =   0  'Flat
         Height          =   810
         Left            =   1350
         TabIndex        =   1
         Top             =   465
         Width           =   6120
      End
      Begin AsocketLib.AsyncSocket AsyncSocket1 
         Left            =   7560
         Top             =   240
         _Version        =   327680
         _ExtentX        =   847
         _ExtentY        =   847
         _StockProps     =   0
         ReceiveBufferSize=   8192
         SendBufferSize  =   8192
         BroadcastEnabled=   0   'False
         LingerEnabled   =   0   'False
         RouteEnabled    =   -1  'True
         KeepAliveEnabled=   0   'False
         OutOfBandEnabled=   0   'False
         ReuseAddressEnabled=   0   'False
         TCPNoDelayEnabled=   0   'False
         LingerMode      =   0
         LingerTime      =   0
         EventMask       =   63
         LocalPort       =   0
         RemotePort      =   0
         SocketType      =   0
         LocalAddress    =   ""
         RemoteName      =   ""
         RemoteAddress   =   ""
         ReceiveTimeout  =   -1
         SendTimeout     =   -1
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Local Address:"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   0
         Left            =   165
         TabIndex        =   6
         Top             =   150
         Width           =   1155
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Local Port:"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   1
         Left            =   2940
         TabIndex        =   5
         Top             =   150
         Width           =   840
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Status:"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   435
         TabIndex        =   4
         Top             =   450
         Width           =   855
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub AsyncSocket1_OnAccept(ByVal SocketHandle As Long, ByVal ErrorCode As Integer)
    Dim foo As New frmSession
    'Declare a dummy variable to represent frmSession

    foo.textSocketHandle = CStr(SocketHandle)
    'Makes tSH into a string with SH in it
    lstStatus.AddItem "OnAccept" & Str(ErrorCode)
    lstStatus.AddItem "Socket Address: " & foo.ClientSocket.SocketAddress
    If Trim$(foo.ClientSocket.RemoteAddress) <> "" Then
        foo.lblClientIP = foo.ClientSocket.RemoteAddress
    Else
        foo.lblClientIP = "Unknown IP"
    End If
    foo.lblClientPort = CStr(foo.ClientSocket.RemotePort)
    'The above statements show us who is connecting
    'by writing their IP to the Status Window

End Sub

Private Sub AsyncSocket1_OnClose(ByVal ErrorCode As Integer)
    lstStatus.AddItem "Close" & Str(ErrorCode)
    'Tells us someone's left
    
    AsyncSocket1.Action = ASocketClose
    'Close the socket they were using
    
    AsyncSocket1.Action = ASocketCreate
    AsyncSocket1.Action = ASocketListen
    'Create a new socket in its place and listen w/ it

End Sub

Private Sub AsyncSocket1_OnConnect(ByVal ErrorCode As Integer)
    lstStatus.AddItem "Connect" & Str(ErrorCode)
    'Tell us someone's connected

    AsyncSocket1.LocalPort = 1024
    AsyncSocket1.LocalAddress = AsyncSocket1.SocketAddress
    'Assign the port and Address
    
    txtLocalAddress = AsyncSocket1.LocalAddress
    txtLocalPort = AsyncSocket1.LocalPort
    'Label the text boxes with their addy and port
    
    AsyncSocket1.Action = ASocketClose
    AsyncSocket1.Action = ASocketCreate
    AsyncSocket1.Action = ASocketListen
    'Close the 1st socket ('cuz sum1's connected) +
    're-creat it to listen on again.

End Sub

Private Sub mnuAbout_Click()

Dim MsgResponse, Msg1, Msg2, MsgAbout
' Dim the variables for the message box

Msg1 = "This chat client was programmed by "
Msg2 = "DragonLord Xian. E-mail him at std@leo.net.gull-lake.sk.ca"
'The above variables are so we can see both lines

MsgAbout = Msg1 + Msg2
'combines the variables

MsgResponse = MsgBox(MsgAbout, 0, "DLX Software", 0, 1000)
'Show a messagebox containing the info about the app

End Sub

Private Sub mnuFileExit_Click()
    AsyncSocket1.Action = ASocketClose
    Unload frmSession
    Unload Me
End Sub

Private Sub MDIForm_Load()

    AsyncSocket1.LocalPort = 0
    AsyncSocket1.LocalAddress = "0.0.0.0"
    '
    ' Winsock won't tell us what our ip address is until
    ' after we connect to some remote socket.  Here I've
    ' chosen to use mit.edu.
    '
    AsyncSocket1.RemotePort = 13

    AsyncSocket1.RemoteNameAddrXlate = True
    AsyncSocket1.RemoteName = "mit.edu"
    AsyncSocket1.RemoteNameAddrXlate = False
    'AsyncSocket1.RemoteAddress = "18.72.2.1"
    'The above address (18.72.2, etc.) is in case you can't
    'do a name lookup

    AsyncSocket1.Action = ASocketCreate
    On Error Resume Next
    'Creates the socket to listen on

    AsyncSocket1.Action = ASocketConnect
    If (Err <> 0 And Err <> 10035) Then
        MsgBox Error
    End If
    On Error GoTo 0
    
End Sub

Private Sub MDIForm_Resize()
    'This procedure resizes the Status box appropriatly
    lstStatus.Width = Me.ScaleWidth - (lstStatus.Left + 5 * Screen.TwipsPerPixelX)
End Sub
