VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   3000
   ClientLeft      =   2445
   ClientTop       =   2865
   ClientWidth     =   6000
   Icon            =   "Splash.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   200
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Path As String 'File path

Private Sub Form_Load()

'[Set paths to game location on Hard Disk]
If (Right(App.Path, 1) <> "\") Then
    Path = App.Path & "\"
        Else
Path = App.Path
End If
'[End of Set Paths]

'[Banner Initialization]
Me.ScaleMode = 3
Me.Visible = True
Me.Picture = LoadPicture(Path + "BnrStart.bmp")
Call FuncTimeOut(4)

'[Main Form Init]
frmMap.Visible = False
Load frmMap
frmMap.Show
Unload Me

End Sub
