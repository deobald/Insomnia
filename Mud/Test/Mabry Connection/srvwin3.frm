VERSION 5.00
Object = "{7A245FC3-9D37-11CF-840F-444553540000}#5.0#0"; "ASOCK32.OCX"
Begin VB.Form frmSession 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Socket Server Session"
   ClientHeight    =   3405
   ClientLeft      =   2085
   ClientTop       =   2460
   ClientWidth     =   6165
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
   Icon            =   "srvwin3.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3405
   ScaleWidth      =   6165
   Begin VB.TextBox textSocketHandle 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   600
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   3000
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.ListBox listTranscript 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1590
      Left            =   105
      MultiSelect     =   1  'Simple
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   600
      Width           =   5940
   End
   Begin VB.TextBox textMessageEntry 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   1380
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   2520
      Width           =   4665
   End
   Begin AsocketLib.AsyncSocket AsyncSocket1 
      Left            =   0
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
      LocalPort       =   0
      RemotePort      =   0
      SocketType      =   0
      LocalAddress    =   ""
      RemoteName      =   ""
      RemoteAddress   =   ""
      ReceiveTimeout  =   -1
      SendTimeout     =   -1
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Client Port:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3600
      TabIndex        =   8
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Client IP:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3600
      TabIndex        =   7
      Top             =   0
      Width           =   975
   End
   Begin VB.Label lblClientPort 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4680
      TabIndex        =   6
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label lblClientIP 
      Caption         =   "0.0.0.0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4680
      TabIndex        =   5
      Top             =   0
      Width           =   1335
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Transcript:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   75
      TabIndex        =   2
      Top             =   240
      Width           =   810
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   -15
      X2              =   7800
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   7815
      Y1              =   2360
      Y2              =   2360
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Enter Message:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   105
      TabIndex        =   3
      Top             =   2520
      Width           =   1215
   End
End
Attribute VB_Name = "frmSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub AsyncSocket1_OnClose(ByVal ErrorCode As Integer)
    AsyncSocket1.Action = ASocketClose
    Unload Me
End Sub

Private Sub AsyncSocket1_OnReceive(ByVal ErrorCode As Integer)
    Dim t As String
    '
    ' Receive data available, get it
    '
    AsyncSocket1.Action = ASocketReceive
    t = AsyncSocket1.ReceiveBuffer
    '
    ' Echo to client
    '
    AsyncSocket1.SendBuffer = t
    AsyncSocket1.Action = ASocketSend
    '
    ' Add to transcript
    '
    If (Len(t) = 1 And Asc(t) = 13) Then
        t = ""
    End If
    listTranscript.AddItem "recv: " & t
    listTranscript.ListIndex = listTranscript.ListCount - 1
End Sub

Private Sub textMessageEntry_KeyPress(KeyAscii As Integer)
    Dim t As String

    If (KeyAscii = 13) Then
        t = textMessageEntry.Text
        listTranscript.AddItem "send: " & t
        listTranscript.ListIndex = listTranscript.ListCount - 1
        If (t = "") Then
            t = Chr$(KeyAscii)
        End If
        AsyncSocket1.SendBuffer = t
        AsyncSocket1.Action = ASocketSend
        textMessageEntry.Text = ""
        KeyAscii = 0
    End If
End Sub

Private Sub textSocketHandle_Change()
    AsyncSocket1.Socket = CLng(textSocketHandle.Text)
End Sub

