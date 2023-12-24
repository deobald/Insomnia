VERSION 5.00
Begin VB.Form frmScrollDown 
   Caption         =   "Auto-Scrolling Example"
   ClientHeight    =   3360
   ClientLeft      =   3150
   ClientTop       =   3420
   ClientWidth     =   5685
   LinkTopic       =   "Form1"
   ScaleHeight     =   3360
   ScaleWidth      =   5685
   Begin VB.TextBox txtSend 
      Height          =   315
      Left            =   450
      TabIndex        =   1
      Top             =   1575
      Width           =   4290
   End
   Begin VB.TextBox txtRec 
      Height          =   1215
      Left            =   450
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   375
      Width           =   4290
   End
End
Attribute VB_Name = "frmScrollDown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    'Call InitializeTextBox
Dim Num As Integer

    For Num = 1 To 20 Step 1
        txtRec.Text = txtRec.Text + Chr(13) + Chr(10) + "Blaaaaaaaaahhhhh!!!! (" + Str(Num) + ")."
    Next Num
        
End Sub

Private Sub txtSend_KeyDown(KeyCode As Integer, Shift As Integer)
Dim NumofLines, Lines As Integer, hLines As Integer, vLines As Integer
Dim Num

If KeyCode = vbKeyReturn Then
    txtRec.Text = txtRec.Text + txtSend.Text + Chr(13) + Chr(10)
    txtSend.Text = ""

    hLines = 0
    vLines = -3
    Lines = CLng(&H10000 * hLines) + vLines
    'SendMessage (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
    'NumofLines = SendMessage(txtRec.hwnd, EM_LINESCROLL, 0, Lines)
    'num = ScrollText(txtRec, -3, 0)
    'DoEvents 'Just to stop progy
End If

End Sub

Sub InitializeTextBox()
    Text1.Text = ""
    For i% = 1 To 50
        Text1.Text = Text1.Text + "This is line " + Str$(i%)
        
        ' Add 15 words to a line of text.
        For j% = 1 To 10
            Text1.Text = Text1.Text + " Word " + Str$(j%)
        Next j%

        ' Force a carriage return (CR) and linefeed (LF).

        Text1.Text = Text1.Text + Chr$(13) + Chr$(10)

        x% = DoEvents()
    Next i%
End Sub

Private Sub txtSend_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then

    Dim NumofLines, Lines As Integer
    Dim hLines As Integer, vLines As Integer
    Dim Num

    hLines = 0
    vLines = -3
    Lines = CLng(&H10000 * hLines) + vLines
    'SendMessage (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
    NumofLines = SendMessage(txtRec.hwnd, EM_LINESCROLL, 0, Lines)

End If

End Sub
