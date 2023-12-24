VERSION 4.00
Begin VB.Form frmEnterIP 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Enter Server IP Address"
   ClientHeight    =   2085
   ClientLeft      =   1380
   ClientTop       =   3945
   ClientWidth     =   5430
   ControlBox      =   0   'False
   Height          =   2490
   Left            =   1320
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   5430
   ShowInTaskbar   =   0   'False
   Top             =   3600
   Width           =   5550
   Begin VB.TextBox txtIP 
      Height          =   285
      Left            =   240
      TabIndex        =   3
      Text            =   "Enter IP Address Here"
      Top             =   480
      Width           =   3255
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3840
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   $"enterip.frx":0000
      Height          =   975
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   3255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Server IP Address:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "frmEnterIP"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
   Form1.AsyncSocket1.RemoteAddress = ""
   Unload Me
   End Sub


Private Sub cmdOK_Click()
   Form1.AsyncSocket1.RemoteAddress = Trim(txtIP)
   Unload Me
   End Sub


