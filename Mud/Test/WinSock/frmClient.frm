VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmClient 
   Caption         =   "Client - Insomnia Test"
   ClientHeight    =   3585
   ClientLeft      =   390
   ClientTop       =   1335
   ClientWidth     =   4965
   LinkTopic       =   "Form1"
   ScaleHeight     =   3585
   ScaleWidth      =   4965
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

If txtName.Text > "" And txtIP.Text > "" Then
    WinSock.Connect txtIP.Text 'Connect with the IP in the textbox
    Namo = txtName.Text
    btnConnect.Enabled = False
Else
    MsgBox "Please enter a name and IP address before connecting. Thanks."
End If

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

    Dim SendHolder As String

    SendHolder = "Chat " + txtSend.Text + "§"
    WinSock.SendData SendHolder 'Send the string variable
    
    'txtRec.Text = txtRec.Text + txtSend.Text + HR
    txtSend.Text = "" 'Clear the sendbox

End If

End Sub

Private Sub txtSend_KeyPress(KeyAscii As Integer)

If KeyAscii = 167 Then KeyAscii = False

End Sub

Private Sub Winsock_Close()

MsgBox "Server Dead. Closing Client."
Unload frmClient 'Close Client

End Sub

Private Sub Winsock_Connect()

Dim SendName As String

SendName = "Name " + Namo + "§"
WinSock.SendData SendName

txtRec.Text = txtRec.Text + "Connection Successfull!" + HR

End Sub

Private Sub Winsock_DataArrival(ByVal bytesTotal As Long)
'FIND A WAY TO PUT THE SCROLLBAR AT THE BOTTOM!!!!

Static DataRead As String
Dim LineString As String
Dim CommandString As String

WinSock.GetData DataRead

While InStr(DataRead, "§") 'If there's still the EOF left in DataRead

    CommandString = Left$(DataRead, InStr(DataRead, " ") - 1)
    'InStr(DataRead, " ") - Find a space in DataRead
    'Left$(DataRead, ABOVE - 1) - CommandString = All the the left of the space

    LineString = Mid$(DataRead, InStr(DataRead, " ") + 1)
    'InStr(DataRead, " ") - Find a space in DataRead
    'Mid$(DataRead, ABOVE + 1) - LineString = All the the right of the space

    LineString = Left$(LineString, InStr(LineString, "§") - 1)
    'InStr(LineString, "!") - Find a ! in LineString
    'Left$(LineString, ABOVE - 1) - Take all to the left of the EOF

    
    Select Case CommandString
    
        Case "Chat"
        txtRec.Text = txtRec.Text + LineString + HR
            
        Case "Move"
    
        Case "Attack"
    
        Case Else
        txtRec.Text = txtRec.Text + "Illigal Server Data Sent!"
    
    End Select

    DataRead = Mid$(DataRead, InStr(DataRead, "§") + 1)
    'InStr(DataRead, "!") - Find EOF in DataString
    'Mid$(DataRead, ABOVE + 1) - DataRead = To the right of EOF (nothing)

Wend

End Sub
