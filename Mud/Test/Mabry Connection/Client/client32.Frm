VERSION 4.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Socket Client"
   ClientHeight    =   3615
   ClientLeft      =   1515
   ClientTop       =   1950
   ClientWidth     =   6855
   Height          =   4020
   Left            =   1455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   Top             =   1605
   Width           =   6975
   Begin VB.TextBox textMessageEntry 
      Height          =   900
      Left            =   1560
      TabIndex        =   0
      Top             =   2475
      Width           =   5055
   End
   Begin VB.ListBox listTranscript 
      Height          =   1815
      Left            =   240
      MultiSelect     =   1  'Simple
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   330
      Width           =   6375
   End
   Begin AsocketLib.AsyncSocket AsyncSocket1 
      Left            =   960
      Top             =   2880
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
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Enter Message and Press Enter:"
      Height          =   570
      Left            =   240
      TabIndex        =   1
      Top             =   2445
      Width           =   1215
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   240
      X2              =   6600
      Y1              =   2310
      Y2              =   2310
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   240
      X2              =   6600
      Y1              =   2340
      Y2              =   2340
   End
   Begin VB.Label Label1 
      Caption         =   "Transcript:"
      Height          =   225
      Left            =   240
      TabIndex        =   2
      Top             =   75
      Width           =   900
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit

Private Sub AsyncSocket1_OnClose(ByVal ErrorCode As Integer)
   ' Close the program if our server goes away
   Unload Me
   End Sub

Private Sub AsyncSocket1_OnReceive(ByVal ErrorCode As Integer)
   Dim t As String
   
   ' Get the text received
   t = AsyncSocket1.Receive()
   
   ' If all we have is a carriage return ...
   If (Len(t) = 1 And Asc(Left(t, 1)) = 13) Then
      ' ... just use a blank string.
      t = ""
      End If
      
   ' Add our new string to the transcript
   listTranscript.AddItem t
   listTranscript.ListIndex = listTranscript.ListCount - 1
   End Sub
   
Private Sub Form_Load()
   Me.Show
   
   ' Get the server's IP address.
   frmEnterIP.Show 1
   
   ' If the user pressed Esc/Cancel exit.
   If AsyncSocket1.RemoteAddress = "" Then
      End
      End If
   
   ' Set up the socket control.
   AsyncSocket1.RemoteNameAddrXlate = False
   AsyncSocket1.RemotePort = 1024
   AsyncSocket1.LocalAddress = "0.0.0.0"
   AsyncSocket1.LocalPort = 0
   AsyncSocket1.Create
   
   ' Connect to the server.
   On Error Resume Next
   AsyncSocket1.Connect
   If (Err <> 0 And Err <> 10035) Then
      MsgBox Error
      End If
   On Error GoTo 0
   End Sub
   
Private Sub Form_Unload(Cancel As Integer)
   ' Disconnect from test server
   AsyncSocket1.Close
   End Sub


Private Sub textMessageEntry_KeyPress(KeyAscii As Integer)
   Dim t As String
   
   ' Did the user press Enter?
   If (KeyAscii = 13) Then
      t = textMessageEntry.Text
      If (textMessageEntry.Text = "") Then
         t = Chr$(13)
         End If
      '
      ' If you want to send a string containing binary data
      ' you must assign it to SendBuffer and then use Send
      ' without any arguments.  If you use Send strTemp VB
      ' will truncate the string at the first chr$(0) if it
      ' contains any.
      '
      AsyncSocket1.SendBuffer = t
      AsyncSocket1.Send
      textMessageEntry.Text = ""
      KeyAscii = 0
      End If
   End Sub
