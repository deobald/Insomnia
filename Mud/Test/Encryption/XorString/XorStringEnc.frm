VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5325
   ClientLeft      =   3315
   ClientTop       =   1935
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   5325
   ScaleWidth      =   6585
   Begin VB.TextBox txtDecrypt 
      Height          =   330
      Left            =   1725
      TabIndex        =   3
      Top             =   825
      Width           =   1590
   End
   Begin VB.CommandButton btnDecrypt 
      Caption         =   "Decrypt string here:"
      Height          =   315
      Left            =   150
      TabIndex        =   2
      Top             =   825
      Width           =   1515
   End
   Begin VB.TextBox txtStringToEncrypt 
      Height          =   315
      Left            =   1725
      TabIndex        =   1
      Top             =   450
      Width           =   1590
   End
   Begin VB.CommandButton btnEncrypt 
      Caption         =   "Encrypt this string:"
      Height          =   315
      Left            =   150
      TabIndex        =   0
      Top             =   450
      Width           =   1515
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub Encrypt(Secret As String, PassWord As String)
' secret$ = the string you wish to encrypt or decrypt.
Dim X, L, Char

' PassWord$ = the password with which to encrypt the string.
L = Len(PassWord)
For X = 1 To Len(Secret)
   Char = Asc(Mid(PassWord, (X Mod L) - L * ((X Mod L) = 0), 1))
   Mid$(Secret, X, 1) = Chr$(Asc(Mid(Secret, X, 1)) Xor Char)
Next

End Sub

Private Sub Form_Load()
Dim Secret As String, PassWord As String

Form1.Show  ' Must Show form in Load event before Print is visible.
Secret = "This is the string that will be encrypted."
PassWord = "password"

MsgBox "Encryption will now begin"

Call Encrypt(Secret, PassWord)     'Encrypt the string.
Print " After encrypting it once: "  'Print the result.
Print Secret
Print

MsgBox "Decryption will now begin"

Call Encrypt(Secret, PassWord)  'A second encryption decrypts it.
Print "After a second encryption: "
Print Secret

End Sub
