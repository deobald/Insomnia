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

Winsock(0).Listen 'Socket is listening
For Temp = 1 To MaxPlayers
    Load Winsock(Temp) 'Load a new socket with +1 index number
Next

End Sub

Private Sub txtSend_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Then

    Dim SendHolder As String
    Dim Temp As Integer

    txtRec.Text = txtRec.Text + "<Server> " + txtSend.Text + HR
    SendHolder = "<Server> " + txtSend.Text 'Apply the text to the variable
    txtSend.Text = "" 'Clear the textbox
    
    For Temp = 1 To MaxPlayers
        If Winsock(Temp).State = 7 Then
            DoEvents
            Winsock(Temp).SendData SendHolder 'Send the string
        End If
    Next
    
End If

End Sub

Private Sub WinSock_Close(Index As Integer)

Winsock(Index).Close
txtRec.Text = txtRec.Text + PlayerName(Index) + " Disconnected" + HR

End Sub

Private Sub Winsock_ConnectionRequest(Index As Integer, ByVal requestID As Long)

Dim Temp As Integer

For Temp = 1 To MaxPlayers
    If Winsock(Temp).State <> sckConnected Then 'And WinSock(Temp).State <> 7
        
        Winsock(Temp).Accept requestID 'Accept requestID
        Exit For
    
    End If
Next Temp

End Sub

Private Sub WinSock_DataArrival(Index As Integer, ByVal bytesTotal As Long)

Dim DataHolder
Dim TempName As String

Winsock(Index).GetData DataHolder, vbString

Select Case Left(DataHolder, 10)
    
    Case "!SENDNAME!"
    TempName = Mid(DataHolder, 1) = "          "
    PlayerName(Index) = LTrim(TempName)
    txtRec.Text = txtRec.Text + PlayerName(Index) + " Connected!" + HR
    
    Case Else
    txtRec.Text = txtRec.Text + "<" + PlayerName(Index) + "> " + DataHolder + HR
    'txtRec.ScrollBars.Value = Something + 1
    'FOR/NEXT WINSOCK SEND TO CLIENTS
        
End Select
    
End Sub
