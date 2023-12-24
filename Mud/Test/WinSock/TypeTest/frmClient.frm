VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmClient 
   Caption         =   "Client - Insomnia Test"
   ClientHeight    =   3585
   ClientLeft      =   3555
   ClientTop       =   2025
   ClientWidth     =   4965
   LinkTopic       =   "Form1"
   ScaleHeight     =   3585
   ScaleWidth      =   4965
   Begin VB.CommandButton btnDisconnect 
      Caption         =   "Disconnect"
      Height          =   315
      Left            =   2775
      TabIndex        =   7
      Top             =   375
      Width           =   990
   End
   Begin VB.TextBox txtName 
      Height          =   315
      Left            =   900
      TabIndex        =   6
      Top             =   375
      Width           =   1815
   End
   Begin VB.TextBox txtIP 
      Height          =   315
      Left            =   900
      TabIndex        =   3
      Top             =   75
      Width           =   1815
   End
   Begin VB.TextBox txtSend 
      Height          =   315
      Left            =   75
      TabIndex        =   2
      Top             =   3225
      Width           =   4815
   End
   Begin VB.TextBox txtRec 
      Height          =   2490
      Left            =   75
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   750
      Width           =   4815
   End
   Begin VB.CommandButton btnConnect 
      Caption         =   "Connect!"
      Height          =   315
      Left            =   2775
      TabIndex        =   0
      Top             =   75
      Width           =   990
   End
   Begin MSWinsockLib.Winsock WinSock 
      Left            =   4350
      Top             =   150
      _ExtentX        =   741
      _ExtentY        =   741
      RemotePort      =   6767
   End
   Begin VB.Label lblName 
      Caption         =   "Name:"
      Height          =   240
      Left            =   75
      TabIndex        =   5
      Top             =   450
      Width           =   765
   End
   Begin VB.Label lblIP 
      Caption         =   "IP-Addy:"
      Height          =   240
      Left            =   75
      TabIndex        =   4
      Top             =   150
      Width           =   840
   End
End
Attribute VB_Name = "frmClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public HR
Public Namo As String

Const vbKeyReturn = 13

Private Sub btnConnect_Click()

If WinSock.State = sckConnected Then
    Exit Sub
End If

If txtName.Text > "" And txtIP.Text > "" Then
    WinSock.Connect (txtIP.Text) 'Connect with the IP in the textbox
    Namo = txtName.Text
Else
    MsgBox "Please enter a name and IP address before connecting. Thanks."
End If

End Sub

Private Sub btnDisconnect_Click()

WinSock.Close
txtRec.Text = txtRec.Text + "You have disconnected." + HR

End Sub

Private Sub Form_Load()

txtIP.Text = "204.83.231.18"  'Show default IP
HR = Chr(13) + Chr(10) 'Hard Return

End Sub

Private Sub Form_Unload(Cancel As Integer)

WinSock.Close

End Sub

Private Sub txtSend_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Then

    Dim SendHolder As DataPack
    Dim SendString As String
    
    txtRec.Text = txtRec.Text + txtSend.Text + HR
    
    SendHolder.Type = 0
    SendHolder.Data = txtSend.Text 'Apply the text to the variable
    SendString = Str(SendHolder)
    WinSock.SendData SendString 'Send the string variable
    
    txtSend.Text = "" 'Clear the sendbox

End If

End Sub

Private Sub Winsock_Close()

MsgBox "Server Dead. Closing Client."
Unload frmClient 'Close Client

End Sub

Private Sub Winsock_Connect()

Dim SendName As DataPack
Dim SendString As String

SendName.Type = 3
SendName.Data = txtName.Text
SendString = Str(SendName)
WinSock.SendData SendString

txtRec.Text = txtRec.Text + "Connection Successfull!" + HR

End Sub

Private Sub Winsock_DataArrival(ByVal bytesTotal As Long)

Dim DataHolder As String

WinSock.GetData DataHolder, vbString
txtRec.Text = txtRec.Text + DataHolder + HR

'FIND A WAY TO PUT THE SCROLLBAR AT THE BOTTOM

End Sub

Public Function FuncTimeOut(TOInterval As Single)

Dim TOStart As Single 'TimeOut initial time
TOStart = Timer 'Set initial time to Timer's value
Do: DoEvents: Loop Until Timer - TOStart >= TOInterval
'Loop until Timer has advanced as far as TOInterval

End Function
