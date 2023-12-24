VERSION 5.00
Begin VB.Form frmEncrypt 
   Caption         =   "Encrypt"
   ClientHeight    =   2745
   ClientLeft      =   4065
   ClientTop       =   4260
   ClientWidth     =   2610
   LinkTopic       =   "Form1"
   ScaleHeight     =   2745
   ScaleWidth      =   2610
   Begin VB.CommandButton btnDecrypt 
      Caption         =   "Open and Decrypt"
      Height          =   315
      Left            =   450
      TabIndex        =   1
      Top             =   1050
      Width           =   1665
   End
   Begin VB.CommandButton btnEncrypt 
      Caption         =   "Open and Encrypt"
      Height          =   315
      Left            =   450
      TabIndex        =   0
      Top             =   750
      Width           =   1665
   End
End
Attribute VB_Name = "frmEncrypt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnEncrypt_Click()
  Dim NextByte As Byte

  Close #1, #2
  Open (App.Path + "\Test.bmp") For Binary As #1
  Open (App.Path + "\Encrypt.bmp") For Binary As #2

  Do While Not EOF(1)
    Get #1, , NextByte
    NextByte = NextByte Xor 5
    Put #2, , NextByte
  Loop

  Close #1, #2
End Sub

Private Sub btnDecrypt_Click()
  Dim NextByte As Byte

  Close #3, #4
  Open (App.Path + "\Encrypt.bmp") For Binary As #3
  Open (App.Path + "\Decrypt.bmp") For Binary As #4

  Do While Not EOF(3)
    Get #3, , NextByte
    NextByte = NextByte Xor 5
    Put #4, , NextByte
  Loop

  Close #3, #4
End Sub
