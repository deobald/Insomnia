VERSION 5.00
Begin VB.Form frmStartUp 
   Caption         =   "Start Up"
   ClientHeight    =   5715
   ClientLeft      =   2865
   ClientTop       =   1545
   ClientWidth     =   6630
   LinkTopic       =   "Form1"
   ScaleHeight     =   5715
   ScaleWidth      =   6630
   Begin VB.CommandButton btnConnect 
      Caption         =   "Connect"
      Height          =   1815
      Left            =   1800
      TabIndex        =   0
      Top             =   1425
      Width           =   2640
   End
   Begin VB.Menu mnuGame 
      Caption         =   "Game"
      Visible         =   0   'False
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmStartUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnConnect_Click()
    frmStartUp.Visible = False
    Load frmDirectX
    frmDirectX.Show
End Sub

Private Sub mnuExit_Click()
    Unload frmDirectX
    Unload frmStartUp
End Sub
