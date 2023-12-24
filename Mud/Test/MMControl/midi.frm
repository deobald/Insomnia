VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form frmWave 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1380
   ClientLeft      =   2745
   ClientTop       =   4545
   ClientWidth     =   5715
   LinkMode        =   1  'Source
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1380
   ScaleWidth      =   5715
   Begin VB.HScrollBar hsbWave 
      Height          =   255
      LargeChange     =   3
      Left            =   240
      Max             =   100
      TabIndex        =   0
      Top             =   330
      Width           =   5295
   End
   Begin MCI.MMControl mciWave 
      Height          =   495
      Left            =   2250
      TabIndex        =   1
      Top             =   780
      Width           =   780
      _ExtentX        =   1376
      _ExtentY        =   873
      _Version        =   327680
      BorderStyle     =   0
      PlayEnabled     =   -1  'True
      StopEnabled     =   -1  'True
      PrevVisible     =   0   'False
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.Menu AL_FILE 
      Caption         =   "&File"
      Begin VB.Menu AI_EXIT 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "frmWave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const conInterval = 50
Const conIntervalPlus = 55

Dim CurrentValue As Double

Private Sub AI_EXIT_Click()
    Unload Me
End Sub

Private Sub Form_Load()

MMControl1.Filename = Filename
MMControl1.Type = "Sequencer"
MMControl1.Command = "Open"
MMControl1.Command = "Play"

If NotifyCode = Successful Then
MMControl1.Command = "Prev"
MMControl1.Command = "Play"
End Sub

Dim Filename As String

    ' Allow return of control to go back to the app as soon as the midi begins
    mciWave.Wait = True
    mciWave.DeviceType = "Sequencer"

'[[[[[]]]]]

'IMPORTANT
    Song = "\test.mid"
    Filename = App.Path & Song
    
    ' If the device is open, close it.
    If Not mciWave.Mode = vbMCIModeNotOpen Then
        mciWave.Command = "Close"
    End If

    ' Open the device with the new filename.
    mciWave.Filename = Filename
    On Error GoTo MCI_ERROR
    mciWave.Command = "Open"
    On Error GoTo 0
    Caption = DialogCaption + mciWave.Filename

    ' Set the scrollbar values.
    hsbWave.value = 0
    CurrentValue = 0#
    Exit Sub

MCI_ERROR:
    DisplayErrorMessageBox
    Resume MCI_EXIT

MCI_EXIT:
    Unload frmWave
'[[[[[]]]]]
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'frmMCITest.Show
End Sub


Private Sub hsbWave_Change()

End Sub

Private Sub mciWave_StatusUpdate()
    Dim value As Integer

    ' If the device is not playing, reset to the beginning.
    If Not mciWave.Mode = vbMCIModePlay Then
        hsbWave.value = hsbWave.Max
        mciWave.UpdateInterval = 0
        Exit Sub
    End If
    
    ' Determine how much of the file has played.  Set a
    ' value of the scrollbar between 0 and 100.
    CurrentValue = CurrentValue + conIntervalPlus
    value = CInt((CurrentValue / mciWave.Length) * 100)
    
    If value > hsbWave.Max Then
        value = 100
    End If

    hsbWave.value = value
End Sub

