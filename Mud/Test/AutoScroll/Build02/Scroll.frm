VERSION 5.00
Begin VB.Form frmScroll 
   Caption         =   "Auto Scroller"
   ClientHeight    =   5325
   ClientLeft      =   3315
   ClientTop       =   1935
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   5325
   ScaleWidth      =   6585
   Begin VB.CommandButton Command1 
      Caption         =   "Vertical"
      Height          =   765
      Left            =   975
      TabIndex        =   1
      Top             =   2850
      Width           =   1590
   End
   Begin VB.TextBox Text1 
      Height          =   1665
      Left            =   675
      MultiLine       =   -1  'True
      ScrollBars      =   1  'Horizontal
      TabIndex        =   0
      Top             =   1125
      Width           =   2115
   End
End
Attribute VB_Name = "frmScroll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Sub InitializeTextBox()
Dim i As Integer, j As Integer, x As Integer

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

Private Sub Form_Load()
Call InitializeTextBox
End Sub

Sub Command1_Click()        ' Scroll text 5 vertical lines upward.
Dim Num
Num = ScrollText&(Text1, 5)
End Sub

