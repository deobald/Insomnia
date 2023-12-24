VERSION 5.00
Begin VB.Form frmDirectX 
   BorderStyle     =   0  'None
   Caption         =   "Main"
   ClientHeight    =   5715
   ClientLeft      =   2370
   ClientTop       =   1545
   ClientWidth     =   6630
   LinkTopic       =   "Form1"
   ScaleHeight     =   480
   ScaleMode       =   0  'User
   ScaleWidth      =   442
   ShowInTaskbar   =   0   'False
   Begin VB.Line Line1 
      X1              =   5
      X2              =   445
      Y1              =   25.197
      Y2              =   25.197
   End
   Begin VB.Label lblmnuFile 
      Caption         =   "Game"
      Height          =   165
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   465
   End
End
Attribute VB_Name = "frmDirectX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

frmDirectX.ScaleMode = 3
frmDirectX.Width = 640
frmDirectX.Height = 480

'Needed: dixu.bas, sprite.Cls, directx5.tlb
Call dixuInit(dixuInitFullScreen, frmDirectX, 640, 480, 16)


End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call dixuDone
End Sub

Private Sub lblmnuFile_Click()
    PopupMenu frmStartUp.mnuGame
End Sub
