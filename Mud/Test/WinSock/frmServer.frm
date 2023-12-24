VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmServer 
   Caption         =   "Server - Insomnia Test"
   ClientHeight    =   2715
   ClientLeft      =   3255
   ClientTop       =   5610
   ClientWidth     =   6165
   LinkTopic       =   "Form1"
   ScaleHeight     =   2715
   ScaleWidth      =   6165
   Begin VB.TextBox txtSend 
      Height          =   330
      Left            =   75
      TabIndex        =   3
      Top             =   2325
      Width           =   5340
   End
   Begin VB.TextBox txtRec 
      Height          =   1365
      Left            =   75
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   675
      Width           =   5340
   End
   Begin MSWinsockLib.Winsock WinSock 
      Index           =   0
      Left            =   5625
      Top             =   450
      _ExtentX        =   741
      _ExtentY        =   741
      LocalPort       =   6767
   End
   Begin VB.Label lblSend 
      Caption         =   "Send:"
      Height          =   240
      Left            =   75
      TabIndex        =   2
      Top             =   2100
      Width           =   1365
   End
   Begin VB.Label lblRecieved 
      Caption         =   "Recieved:"
      Height          =   240
      Left            =   75
      TabIndex        =   1
      Top             =   450
      Width           =   1590
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public HR
Public MaxPlayers

Const vbKeyReturn = 13 '<Enter>

Private Sub Form_Load()

Dim Temp As Integer

HR = Chr(13) + Chr(10) 'Hard Return
MaxPlayers = 10

WinSock(0).Listen 'Socket is listening
For Temp = 1 To MaxPlayers
    Load WinSock(Temp) 'Load a new socket with +1 index number
Next

End Sub

Private Sub txtSend_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Then

    Dim SendHolder As String
    Dim Temp As Integer

    SendHolder = "Chat <Server> " + txtSend.Text + "§"
    
    For Temp = 1 To MaxPlayers 'Check conn. players to send to
        If WinSock(Temp).State = 7 Then
            DoEvents
            WinSock(Temp).SendData SendHolder 'Send the string
        End If
    Next
    
    txtRec.Text = txtRec.Text + "<Server> " + txtSend.Text + HR
    txtSend.Text = ""
    
End If

End Sub

Private Sub WinSock_Close(Index As Integer)

WinSock(Index).Close
txtRec.Text = txtRec.Text + PlayerName(Index) + " Disconnected" + HR

Dim Temp As Integer
Dim SendName As String
SendName = "Chat " + PlayerName(Index) + " Disconnected!§"
For Temp = 1 To MaxPlayers 'Check conn. players to send to
    If WinSock(Temp).State = 7 Then
        DoEvents
        WinSock(Temp).SendData SendName 'Send the string
    End If
Next

End Sub

Private Sub Winsock_ConnectionRequest(Index As Integer, ByVal requestID As Long)

Dim Temp As Integer

For Temp = 1 To MaxPlayers
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

While InStr(DataRead, "§") 'If there's still the EOF left in DataRead

    CommandString = Left$(DataRead, InStr(DataRead, " ") - 1)
    'Cut out the command left of the 1st space.
    LineString = Mid$(DataRead, InStr(DataRead, " ") + 1)
    'Use what's left (w/ EOF) as the data
    LineString = Left$(LineString, InStr(LineString, "§") - 1)
    'Get rid of EOF in the data
   
    Select Case CommandString
    
        Case "Chat"
        txtRec.Text = txtRec.Text + PlayerName(Index) + ": " + LineString + HR
        
        Dim ChatTemp As Integer
        Dim SendHolder As String
        SendHolder = "Chat " + PlayerName(Index) + ": " + LineString + "§"
        For ChatTemp = 1 To MaxPlayers 'Check conn. players to send to
            If WinSock(ChatTemp).State = 7 Then
                DoEvents
                WinSock(ChatTemp).SendData SendHolder 'Send the string
            End If
        Next
    
        Case "Move"
        'FOR/NEXT WINSOCK SEND TO CLIENTS
    
        Case "Attack"
    
        Case "Name"
        PlayerName(Index) = LineString 'Assign the name to an index
        txtRec.Text = txtRec.Text + PlayerName(Index) + " Connected!" + HR
        
        Dim NameTemp As Integer
        Dim SendName As String
        SendHolder = "Chat " + PlayerName(Index) + " Connected!§"
        For NameTemp = 1 To MaxPlayers 'Check conn. players to send to
            If WinSock(NameTemp).State = 7 Then
                DoEvents
                WinSock(NameTemp).SendData SendName 'Send the string
            End If
        Next
    
        Case Else
        txtRec.Text = txtRec.Text + "Illigal type attempt by " + PlayerName(Index)
    
    End Select

    DataRead = Mid$(DataRead, InStr(DataRead, "§") + 1)
    'InStr(DataRead, "!") - Find EOF in DataString
    'Mid$(DataRead, ABOVE + 1) - DataRead = To the right of EOF (nothing)

Wend
    
End Sub
