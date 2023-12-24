VERSION 5.00
Begin VB.Form frmScroll 
   Caption         =   "Auto Scroller"
   ClientHeight    =   2910
   ClientLeft      =   2025
   ClientTop       =   3225
   ClientWidth     =   9330
   LinkTopic       =   "Form1"
   ScaleHeight     =   2910
   ScaleWidth      =   9330
   Begin VB.CommandButton Command1 
      Caption         =   "Vertical"
      Height          =   315
      Left            =   150
      TabIndex        =   1
      Top             =   2400
      Width           =   840
   End
   Begin VB.TextBox Text1 
      Height          =   2190
      Left            =   150
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   150
      Width           =   8865
   End
End
Attribute VB_Name = "frmScroll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Sub InitializeTextBox()
Dim I As Integer, J As Integer

Text1.Text = ""
For I = 1 To 50
    Text1.Text = Text1.Text + "This is line " + Str$(I)

    ' Add 15 words to a line of text.
    For J = 1 To 10
        Text1.Text = Text1.Text + " Word " + Str$(J)
    Next J
    
    ' Force a carriage return (CR) and linefeed (LF).
    Text1.Text = Text1.Text + Chr(13) + Chr(10)
    DoEvents

Next I

End Sub

Private Sub Form_Load()
Call InitializeTextBox
End Sub

Sub Command1_Click()        ' Scroll text 5 vertical lines upward.
Dim Num
Num = ScrollText(Text1, 60)
End Sub

