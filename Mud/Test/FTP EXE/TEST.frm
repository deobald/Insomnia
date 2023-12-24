VERSION 5.00
Begin VB.Form frmTEST 
   Caption         =   "Close This Window!"
   ClientHeight    =   1590
   ClientLeft      =   2865
   ClientTop       =   1545
   ClientWidth     =   4335
   LinkTopic       =   "Form1"
   ScaleHeight     =   1590
   ScaleWidth      =   4335
   Begin VB.Timer Timer1 
      Interval        =   30000
      Left            =   1725
      Top             =   975
   End
   Begin VB.Label lblWhy 
      Caption         =   "Steven Deobald made this program to test FTP connections which run EXEs remotely. Feel free to close it. Thanks."
      Height          =   690
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   4065
   End
End
Attribute VB_Name = "frmTEST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Beep
End Sub

Private Sub Timer1_Timer()
Unload Me
End Sub
