VERSION 5.00
Begin VB.Form frmGlowBall 
   Caption         =   "Glow Ball Test"
   ClientHeight    =   2580
   ClientLeft      =   3300
   ClientTop       =   4035
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   ScaleHeight     =   2580
   ScaleWidth      =   4800
   Begin VB.PictureBox picDraw 
      Height          =   840
      Left            =   2700
      ScaleHeight     =   52
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   51
      TabIndex        =   5
      Top             =   225
      Width           =   825
   End
   Begin VB.PictureBox picHolder 
      Height          =   765
      Left            =   75
      Picture         =   "GlowBall.frx":0000
      ScaleHeight     =   705
      ScaleWidth      =   4530
      TabIndex        =   4
      Top             =   1725
      Visible         =   0   'False
      Width           =   4590
   End
   Begin VB.Frame fraSpeed 
      Caption         =   "Animation Speed"
      Height          =   1515
      Left            =   150
      TabIndex        =   0
      Top             =   150
      Width           =   2040
      Begin VB.OptionButton optSpeed4 
         Caption         =   "Very Slow"
         Height          =   240
         Left            =   150
         TabIndex        =   6
         Top             =   1200
         Width           =   1440
      End
      Begin VB.OptionButton optSpeed3 
         Caption         =   "Slow"
         Height          =   240
         Left            =   150
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   900
         Width           =   915
      End
      Begin VB.OptionButton optSpeed2 
         Caption         =   "Medium"
         Height          =   240
         Left            =   150
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   600
         Width           =   915
      End
      Begin VB.OptionButton optSpeed1 
         Caption         =   "Fast"
         Height          =   240
         Left            =   150
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   300
         Width           =   915
      End
   End
End
Attribute VB_Name = "frmGlowBall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Frame As Integer
Dim Animate As Boolean

Private Sub Form_Load()
    Frame = 0
    picDraw.ScaleMode = 3
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Animate = False
End Sub

Private Sub optSpeed1_Click()
    Animate = True
    Call AnimateBall(56)
End Sub

Private Sub optSpeed2_Click()
    Animate = True
    Call AnimateBall(36)
End Sub

Private Sub optSpeed3_Click()
    Animate = True
    Call AnimateBall(28)
End Sub

Private Sub optSpeed4_Click()
    Animate = True
    Call AnimateBall(14)
End Sub

Public Sub AnimateBall(Speed)
Dim Dummy

While Animate = True

    If Frame = 5 Then
        Dummy = BitBlt(picDraw.hDC, 0, 0, 50, 50, picHolder.hDC, (Frame * 50), 0, SRCCOPY)
        Frame = 0
    Else
        Dummy = BitBlt(picDraw.hDC, 0, 0, 50, 50, picHolder.hDC, (Frame * 50), 0, SRCCOPY)
        Frame = Frame + 1
    End If
    
    Call FuncTimeOut((1 / Speed) * 5)

Wend

End Sub

Public Function FuncTimeOut(TOInterval As Single)

Dim TOStart As Single
TOStart = Timer
Do: DoEvents: Loop Until Timer - TOStart >= TOInterval

End Function

