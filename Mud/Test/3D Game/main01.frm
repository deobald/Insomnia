VERSION 5.00
Object = "{9A4D6F83-5291-101C-96E6-0020AF38F4BB}#1.0#0"; "TEGFLR32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Wisp: Realm of Dreams"
   ClientHeight    =   4560
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6240
   LinkTopic       =   "Form1"
   ScaleHeight     =   4560
   ScaleWidth      =   6240
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtCols 
      Height          =   345
      Left            =   2160
      TabIndex        =   6
      Top             =   4140
      Width           =   675
   End
   Begin VB.TextBox txtRows 
      Height          =   345
      Left            =   2160
      TabIndex        =   5
      Top             =   3780
      Width           =   675
   End
   Begin VB.CommandButton btnChange 
      Default         =   -1  'True
      Height          =   345
      Left            =   3120
      TabIndex        =   1
      Top             =   3840
      Width           =   525
   End
   Begin VB.TextBox txtChange 
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   3390
      Width           =   2775
   End
   Begin FloorLibCtl.Floor flrMain 
      Left            =   30
      Top             =   30
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
   End
   Begin VB.Label lblCols 
      Caption         =   "Columns:"
      Height          =   285
      Left            =   1440
      TabIndex        =   4
      Top             =   4170
      Width           =   675
   End
   Begin VB.Label lblRows 
      Caption         =   "Rows:"
      Height          =   255
      Left            =   1620
      TabIndex        =   3
      Top             =   3810
      Width           =   495
   End
   Begin VB.Label lblFloor 
      Caption         =   "Floor:"
      Height          =   255
      Left            =   1620
      TabIndex        =   2
      Top             =   3450
      Width           =   495
   End
   Begin VB.Menu mnuGame 
      Caption         =   "&Game"
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'All variables must be declared
Option Explicit

'Declare variables
Dim OpenResult As Integer
Dim Message As String
Dim Path As String

Dim CurrentFloor As String
Dim CurrentRows As Integer
Dim CurrentCols As Integer
Dim PosX As Integer
Dim PosY As Integer

Private Sub btnChange_Click()
CurrentFloor = txtChange.Text
CurrentRows = txtRows.Text
CurrentCols = txtCols.Text
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

Select Case KeyCode
       
   Case 37, 100
        'Left key (37) or 4 key (100) was pressed.
        flrMain.Angle = flrMain.Angle + 6
       
   Case 39, 102
        'Right key (39) or 6 key (102) was pressed.
        flrMain.Angle = flrMain.Angle - 6
     
   Case 38, 104
        'Up key (38) or 8 key (104) was pressed.
        flrMain.Advance 40
      
   Case 40, 98
        'Down key (40) or 2 key (98) was pressed.
        flrMain.Advance -40

End Select

' Display the 3D view.
flrMain.Display3D

End Sub

Private Sub Form_Load()
  
' Get the name of the directory where the
' program resides. If the last letter isn't a '/',
' then add one to it.
Path = App.Path
If Right(Path, 1) <> "\" Then
   Path = Path + "\"
End If

'*Define Variables Below*
CurrentFloor = "floor001.flr"
CurrentRows = 50
CurrentCols = 50
PosX = 2
PosY = 2
'4 * flrMain.CellWidth
'This was the original position for both X + Y

' Open the FLOOR50.FLR file.
flrMain.filename = Path + CurrentFloor
flrMain.hWndDisplay = Me.hWnd
flrMain.NumOfRows = CurrentRows
flrMain.NumOfCols = CurrentCols
OpenResult = flrMain.Open

' If FLR file could not be opened, terminate
' the program.
If OpenResult <> 0 Then
    Message = "Unable to open file: " + flrMain.filename
    Message = Message + Chr(13) + Chr(10)
    Message = Message + "Error Code: " + Str(OpenResult)
    MsgBox Message, vbCritical, "Error"
    End
End If

' Set the initial user's position and viewing angle.
flrMain.X = PosX
flrMain.Y = PosY
flrMain.Angle = 0

' Set the colors of the walls, ceiling, and floor.
flrMain.WallColorA = 7    ' White
flrMain.WallColorB = 4    ' Red
flrMain.CeilingColor = 11 ' Light Cyan
flrMain.FloorColor = 2    ' Green
flrMain.StripeColor = 0   ' Black

' Load the sprites. *COMMENTED OUT*
'flrMain.SpritePath = Path
'flrMain.Sprite(65) = "TREE.BMP"   ' 65 = ASCII of "A"
'flrMain.Sprite(66) = "LIGHT.BMP"  ' 66 = ASCII of "B"
'flrMain.Sprite(67) = "EX1.BMP"    ' 67 = ASCII of "C"
'flrMain.Sprite(68) = "EX2.BMP"    ' 68 = ASCII OF "D"
'flrMain.Sprite(69) = "JOG1.BMP"   ' 69 = ASCII OF "E"
'flrMain.Sprite(70) = "JOG2.BMP"   ' 70 = ASCII OF "F"
'flrMain.Sprite(71) = "JOG3.BMP"   ' 71 = ASCII OF "G"
'flrMain.Sprite(72) = "JOG4.BMP"   ' 72 = ASCII OF "H"

' Set sprite number 66 (the Light sprite)
' as a soft sprite.
' (i.e. the user can walk through this sprite).
'flrMain.SetSpriteSoft (66)

End Sub

Private Sub Form_Paint()

' If the form is minimized, terminate this procedure to save RAM.
If Me.WindowState = 1 Then Exit Sub

' Display the 3D view.
flrMain.Display3D

End Sub

Private Sub mnuExit_Click()

' Terminate the program.
Unload Me

End Sub
