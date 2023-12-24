VERSION 4.00
Begin VB.Form UDPPeerB 
   Caption         =   "UDPPeerB"
   ClientHeight    =   3780
   ClientLeft      =   5355
   ClientTop       =   1515
   ClientWidth     =   4020
   Height          =   4185
   KeyPreview      =   -1  'True
   Left            =   5295
   LinkTopic       =   "UDPPeerB"
   LockControls    =   -1  'True
   ScaleHeight     =   3780
   ScaleWidth      =   4020
   Top             =   1170
   Width           =   4140
   Begin VB.ListBox lstTranscript 
      Height          =   2400
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   3735
   End
   Begin VB.TextBox txtSend 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   3240
      Width           =   3735
   End
   Begin VB.Label Label1 
      Caption         =   "Transcript:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Enter message here then press Enter:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   3000
      Width           =   2655
   End
   Begin AsocketLib.AsyncSocket ASocket1 
      Left            =   3120
      Top             =   2760
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
      LocalPort       =   1002
      RemotePort      =   1001
      SocketType      =   1
      LocalAddress    =   ""
      RemoteName      =   ""
      RemoteAddress   =   ""
      ReceiveTimeout  =   -1
      SendTimeout     =   -1
   End
End
Attribute VB_Name = "UDPPeerB"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Form_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
        If (txtSend.Text <> "") Then
            lstTranscript.AddItem "send: " & txtSend.Text
            ASocket1.SendBuffer = txtSend.Text
            ASocket1.SendTo
            txtSend.Text = ""
        End If
        KeyAscii = 0
    End If
End Sub

Private Sub ASocket1_OnReceive(ByVal ErrorCode As Integer)
    lstTranscript.AddItem "recv: " & ASocket1.Receive
End Sub

Private Sub Form_Load()
    ASocket1.Create
    '
    ' Workaround for version 5.00.004
    '
    If (ASocket1.Version = "5.00.004") Then
        On Error Resume Next
        ASocket1.Bind
        If (Err And Err <> 10022) Then
            MsgBox "Unexpected error: " & Error
        End If
        On Error Resume Next
    Else
        ASocket1.Bind
    End If
End Sub
