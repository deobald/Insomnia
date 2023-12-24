VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2205
   ClientLeft      =   3210
   ClientTop       =   3165
   ClientWidth     =   3840
   LinkTopic       =   "Form1"
   ScaleHeight     =   2205
   ScaleWidth      =   3840
   Begin VB.TextBox txtInfo 
      Height          =   315
      Left            =   1050
      TabIndex        =   2
      Top             =   825
      Width           =   1965
   End
   Begin VB.CommandButton btnStop 
      Caption         =   "Stop"
      Height          =   435
      Left            =   1260
      TabIndex        =   1
      Top             =   1230
      Width           =   1215
   End
   Begin VB.CommandButton btnStart 
      Caption         =   "Start"
      Height          =   435
      Left            =   1260
      TabIndex        =   0
      Top             =   330
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim MusicPath As String


Private Sub btnStart_Click()
Dim MidFile
MidFile = App.Path & "\ins.mid"
Dim Dummy As Integer

' The following will open the sequencer with the CANYON.MID
' file. Canyon is the device_id.

MsgBox MidFile
Dummy = mciSendString("open ins.mid type sequencer alias midi", 0&, 0, 0)

'D:\PROGRA~1\DevStudio\VB\Mud\Test\MidiAP~1\
'" + MusicPath + "duke2.mid - Add this code to change midis

' The wait tells the MCI command to complete before returning
' control to the application.

Dummy = mciSendString("play midi", 0&, 0, 0)

End Sub

Private Sub btnStop_Click()
Dim Dummy As Integer

'Close *.MID file and sequencer device
Dummy = mciSendString("close midi", 0&, 0, 0)

End Sub

Private Sub Form_Load()

MusicPath = App.Path
If Right(MusicPath, 1) <> "\" Then
   MusicPath = MusicPath + "\"
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

Dim Dummy As Integer

'Close *.MID file and sequencer device - cannot be thru [] button in VB
Dummy = mciSendString("close midi", 0&, 0, 0)

End Sub
