VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H80000004&
   ClientHeight    =   7200
   ClientLeft      =   975
   ClientTop       =   930
   ClientWidth     =   9600
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   480
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   640
   Begin VB.Timer tmrPlayers 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2775
      Top             =   2325
   End
   Begin VB.TextBox txtIP 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4575
      TabIndex        =   0
      Text            =   "204.83.231."
      Top             =   2325
      Width           =   2115
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4575
      TabIndex        =   1
      Text            =   "Xian"
      Top             =   2625
      Width           =   2115
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4575
      TabIndex        =   2
      Text            =   "password"
      Top             =   2925
      Width           =   2115
   End
   Begin VB.TextBox txtSend 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   150
      TabIndex        =   3
      Top             =   6900
      Visible         =   0   'False
      Width           =   9315
   End
   Begin MCI.MMControl MMControl 
      Height          =   615
      Left            =   8100
      TabIndex        =   6
      Top             =   4800
      Visible         =   0   'False
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   1085
      _Version        =   327680
      PrevEnabled     =   -1  'True
      PlayEnabled     =   -1  'True
      PauseEnabled    =   -1  'True
      StopEnabled     =   -1  'True
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.Timer tmrDraw 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   8100
      Top             =   4350
   End
   Begin VB.TextBox txtRec 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF57A4&
      Height          =   1740
      Left            =   150
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   5175
      Visible         =   0   'False
      Width           =   9315
   End
   Begin MSWinsockLib.Winsock WinSock 
      Left            =   8550
      Top             =   4350
      _ExtentX        =   741
      _ExtentY        =   741
      RemotePort      =   6767
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4800
      Left            =   2325
      ScaleHeight     =   320
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   350
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   150
      Visible         =   0   'False
      Width           =   5250
   End
   Begin VB.PictureBox picTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1515
      Left            =   1800
      Picture         =   "Main.frx":000C
      ScaleHeight     =   101
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   401
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   150
      Width           =   6015
   End
   Begin VB.Label lblXY 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Maiandra GD"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF57A4&
      Height          =   390
      Left            =   7875
      TabIndex        =   15
      Top             =   300
      Width           =   1365
   End
   Begin VB.Label lblIP 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "IP Address:"
      ForeColor       =   &H00FF57A4&
      Height          =   240
      Left            =   3675
      TabIndex        =   14
      Top             =   2400
      Width           =   840
   End
   Begin VB.Label lblUserName 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "User Name:"
      ForeColor       =   &H00FF57A4&
      Height          =   240
      Left            =   3675
      TabIndex        =   13
      Top             =   2700
      Width           =   840
   End
   Begin VB.Label lblPassword 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Password:"
      ForeColor       =   &H00FF57A4&
      Height          =   240
      Left            =   3750
      TabIndex        =   12
      Top             =   3000
      Width           =   765
   End
   Begin VB.Label lblStatus 
      BackColor       =   &H00000000&
      Caption         =   "Status: Not Connected"
      ForeColor       =   &H00FF57A4&
      Height          =   240
      Left            =   3150
      TabIndex        =   11
      Top             =   3750
      Width           =   3465
   End
   Begin VB.Label lblConnect 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Connect"
      BeginProperty Font 
         Name            =   "Maiandra GD"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF57A4&
      Height          =   540
      Left            =   2775
      TabIndex        =   9
      Top             =   5025
      Width           =   2265
   End
   Begin VB.Label lblEnter 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Enter"
      BeginProperty Font 
         Name            =   "Maiandra GD"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF57A4&
      Height          =   540
      Left            =   3150
      TabIndex        =   8
      Top             =   5025
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.Label lblExit 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Maiandra GD"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF57A4&
      Height          =   540
      Left            =   5700
      TabIndex        =   7
      Top             =   5025
      Width           =   1065
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'[TEMPORARY VARIABLES]

'[String]
Dim HR
Dim KeepFocusOn As String

'[Sprite]
Dim PicCharSelect As tArea
Dim Char As tSprite   'declare the Char as user defined sprite
Dim PicSprite As tArea
Dim PicSpriteMask As tArea

Dim CharInit As CharData

'[Tile]
Dim PicTile As tArea 'Tile Hdc
Dim PicObject As tArea  'Overlay tiles
Dim PicObjMask As tArea 'Overlay tiles' masks

'[Work Area/Map]
Dim PicTemp As tArea 'Area that map data is compiled to before seen
Dim PicBottomWork As tArea 'Ground tile map area
Dim PicTopWork As tArea 'Object tile map area
Dim PicTopMask As tArea

'[Tile, map, and movement declarations]
Const TileSize = 32      'size of each sprite/map tile
Const MoveSize = (TileSize / 4) 'The ammount of movement in pixels
Const ScreenWidth = 352
Const ScreenHeight = 320
Const MapSize = 1600

Dim KeyBusy As Boolean
Dim RightIsSolid As Boolean
Dim LeftIsSolid As Boolean
Dim UpIsSolid As Boolean
Dim DownIsSolid As Boolean

Dim MapOffSetX As Integer
Dim MapOffSetY As Integer
Dim MapWidth As Long 'The width of the map, which the mapfile will set
Dim MapHeight As Long 'The height of the map, which the mapfile will set
Dim MapHolder(1 To 50, 1 To 50) As Record 'The Variable for maps is a user-defined record
Dim MapFile As String

'**********************************************************
'******************** INITIALIZATIONS *********************
'**********************************************************

Private Sub Var_Init()
Dim Dummy As Long

'[String Initializations]
HR = Chr(13) + Chr(10)

'[Set paths to game location on Hard Disk]
Path = App.Path
If (Right(Path, 1) <> "\") Then
    Path = Path & "\"
End If
PathArchive = Path + "Archive\"
PathMap = Path + "Maps\"
PathMusic = Path + "Music\"
PathSound = Path + "SFX\"
'[End of Set Paths]

'[Form Initializations]
Me.KeyPreview = True  'Look for keystrokes
Me.BackColor = BLACK

'[Background PictureBox Settings]
picMain.ScaleMode = 3
picMain.Width = ScreenWidth
picMain.Height = ScreenHeight
picMain.Picture = LoadPicture(PathArchive & "PaletteTile.bmp")
    'Tile Palette Bitmap
    Dim Y As Integer
    Dim X As Integer
    For Y = 0 To picMain.Height Step TileSize
        For X = 0 To picMain.Width Step TileSize
            picMain.PaintPicture picMain.Picture, X, Y
        Next X
    Next Y

End Sub

Private Sub hDC_Init()
Dim Dummy As Long

'*********************** WORK AREAS ***********************
'[TEMPORARY WORK AREA - Mem Hdc]
PicTemp.hdc = 0
PicTemp.Left = 0
PicTemp.Top = 0
PicTemp.Width = ScreenWidth
PicTemp.Height = ScreenHeight
PicTemp.hdc = CreateMemHdc(picMain.hdc, PicTemp.Width, PicTemp.Height)
Dummy = SelectPalette(PicTemp.hdc, picMain.Picture.hPal, False)
Dummy = RealizePalette(PicTemp.hdc)

'[BOTTOM TILE WORK AREA]
PicBottomWork.hdc = 0
PicBottomWork.Left = 0
PicBottomWork.Top = 0
PicBottomWork.Width = MapSize
PicBottomWork.Height = MapSize
PicBottomWork.hdc = CreateMemHdc(picMain.hdc, MapSize, MapSize)
Dummy = SelectPalette(PicBottomWork.hdc, picMain.Picture.hPal, False)
Dummy = RealizePalette(PicBottomWork.hdc)

'[TOP WORK AREA]
PicTopWork.hdc = 0
PicTopWork.Left = 0
PicTopWork.Top = 0
PicTopWork.Width = MapSize
PicTopWork.Height = MapSize
PicTopWork.hdc = CreateMemHdc(picMain.hdc, MapSize, MapSize)
Dummy = SelectPalette(PicBottomWork.hdc, picMain.Picture.hPal, False)
Dummy = RealizePalette(PicBottomWork.hdc)

'[TOP MASK WORK AREA]
PicTopMask.hdc = 0
PicTopMask.Left = 0
PicTopMask.Top = 0
PicTopMask.Width = MapSize
PicTopMask.Height = MapSize
PicTopMask.hdc = CreateMemHdc(picMain.hdc, MapSize, MapSize)
Dummy = SelectPalette(PicBottomWork.hdc, picMain.Picture.hPal, False)
Dummy = RealizePalette(PicBottomWork.hdc)
'**********************************************************

'************************** TILES *************************
'[MAPTILES - Mem Hdc]
PicTile.hdc = 0
PicTile.Left = 0
PicTile.Top = 0
PicTile.Width = TileSize 'Width of the memory holder
PicTile.Height = TileSize 'Height of the memory holder
PicTile.hdc = CreateMemHdc(picMain.hdc, TileSize, TileSize)
Call LoadBmpToHdc(PicTile.hdc, "BotTiles.bmp")

'[OBJECT TILES - Mem Hdc]
PicObject.hdc = 0
PicObject.Left = 0
PicObject.Top = 0
PicObject.Width = TileSize 'Width of the memory holder
PicObject.Height = TileSize 'Height of the memory holder
PicObject.hdc = CreateMemHdc(picMain.hdc, TileSize, TileSize)
Call LoadBmpToHdc(PicObject.hdc, "TopTiles.bmp")

'[OBJECT MASK TILES - Mem Hdc]
PicObjMask.hdc = 0
PicObjMask.Left = 0
PicObjMask.Top = 0
PicObjMask.Width = TileSize 'Width of the memory holder
PicObjMask.Height = TileSize 'Height of the memory holder
PicObjMask.hdc = CreateMemHdc(picMain.hdc, TileSize, TileSize)
Call LoadBmpToHdc(PicObjMask.hdc, "TopMask.bmp")
'**********************************************************

'************************ SPRITES *************************
'[SPRITES Mem Hdc]
PicSprite.hdc = 0
PicSprite.Left = 0
PicSprite.Top = 0
PicSprite.Width = TileSize 'Width of the memory holder
PicSprite.Height = (TileSize * 2) 'Height of the memory holder
PicSprite.hdc = CreateMemHdc(picMain.hdc, TileSize, (TileSize * 2))
Call LoadBmpToHdc(PicSprite.hdc, "Sprite.bmp")

'[SPRITE MASK - Mem Hdc]
PicSpriteMask.hdc = 0
PicSpriteMask.Left = 0
PicSpriteMask.Top = 0
PicSpriteMask.Width = TileSize 'Width of the mask memory holder
PicSpriteMask.Height = (TileSize * 2) 'Height of the mask memory holder
PicSpriteMask.hdc = CreateMemHdc(picMain.hdc, TileSize, (TileSize * 2))
Call LoadBmpToHdc(PicSpriteMask.hdc, "SpriteMsk.bmp")

'[CHAR SELECTION Mem hDC]
PicCharSelect.hdc = 0
PicCharSelect.Left = 0
PicCharSelect.Top = 0
PicCharSelect.Width = TileSize 'Width of the memory holder
PicCharSelect.Height = (TileSize * 2) 'Height of the memory holder
PicCharSelect.hdc = CreateMemHdc(picMain.hdc, TileSize, (TileSize * 2))
Char.SpriteNum = 3
Dummy = SelectPalette(PicTemp.hdc, picMain.Picture.hPal, False)
Dummy = RealizePalette(PicTemp.hdc)
'**********************************************************

'[Clear off palette-setters]
picMain.Picture = LoadPicture()

End Sub

Private Sub Char_Init()

'[Character Animation Values]
Char.Width = TileSize         'width of sprite
Char.Height = (TileSize * 2)  'height of sprite
Char.Frame = 1     'Which frame in a 4-frame sequence
Char.Direction = 0 '0 = Down, 96 = Up, 192 = Right, 288 = Left
Char.SprX = 32     'Which frame, horizontally. Start with standing.

'[Source Area of Main Character Bitmap]
Char.src.hdc = PicSprite.hdc
Char.src.Left = PicSprite.Left
Char.src.Top = PicSprite.Top
Char.src.Width = PicSprite.Width
Char.src.Height = PicSprite.Height

'[Mask Area for Main Charmask Bitmap]
Char.mask.hdc = PicSpriteMask.hdc
Char.mask.Left = PicSpriteMask.Left
Char.mask.Top = PicSpriteMask.Top
Char.mask.Width = PicSpriteMask.Width
Char.mask.Height = PicSpriteMask.Height

'[Make the Char Background the Main View]
Char.bkg.hdc = PicTemp.hdc
Char.bkg.Left = PicTemp.Left
Char.bkg.Top = PicTemp.Top
Char.bkg.Width = PicTemp.Width
Char.bkg.Height = PicTemp.Height

End Sub

Private Sub Initializations()
Dim R

'[Draw the player's current map]
Call Map_Draw("Pond2.map")

GameScreen = "Connect"
lblConnect.Visible = True
lblIP.Visible = True
lblUserName.Visible = True
lblPassword.Visible = True
lblStatus.Visible = True
txtIP.Visible = True
txtName.Visible = True
txtPassword.Visible = True
R = PutFocus(txtName.hWnd)

End Sub

'**********************************************************
'******************** PIXEL / DRAWING *********************
'**********************************************************

Private Sub Mask_Make(dest As tArea, src As tArea, Width As Integer, Height As Integer)
'MAKE SURE ALL FORMS AND OBJECTS HAVE PIXEL AS THEIR SCALEMODE!

Dim X As Integer    'x pixel pos
Dim Y As Integer    'y pixel pos
Dim Color As Long   'current color of pixel
Dim Dummy As Long   'dummy return code needed for blit

'Make mask pixel by pixel (FORM AND ALL OBJs MUST HAVE 3 FOR SCALEMODE)
For Y = 0 To Height    'do until Width
    For X = 0 To Width    'do until Height
        Color = GetPixel(src.hdc, X, Y)  'check pixel color
        If Color = BLACK Then  'Black: make it white
            Dummy = SetPixel(dest.hdc, X, Y, WHITE)
        Else
            Dummy = SetPixel(dest.hdc, X, Y, BLACK) 'Color: make it black
        End If
    Next X
Next Y

End Sub

Private Sub Board_Refresh()
Dim Dummy As Long
Dim HeadLoop As Integer
Dim CheckIfOver As Integer
Dim PDrawLoop As Integer

Dim AvoidHeadBlit
Dim AvoidPlayerHeadBlit
Dim AvoidCharUnderlay
Dim AvoidPlayerUnderlay

'[Map Bottom Layer Blit]
Dummy = BitBlt(PicTemp.hdc, 0, 0, ScreenWidth, ScreenHeight, PicBottomWork.hdc, MapOffSetX, MapOffSetY, SRCCOPY)

'[Map Edge Blackness Code]
Dim xBlack As Single, yBlack As Single, MaxBlack As Single
Dim AvoidBlackX1, AvoidBlackX2, AvoidBlackY1, AvoidBlackY2

If Char.X < 6 Then
    
    Select Case Char.X 'Find the char's x position and invert it for drawing the black
        Case 5.75: MaxBlack = 0.25
        Case 5.5: MaxBlack = 0.5
        Case 5.25: MaxBlack = 0.75
        Case 5: MaxBlack = 1

        Case 4.75: MaxBlack = 1.25
        Case 4.5: MaxBlack = 1.5
        Case 4.25: MaxBlack = 1.75
        Case 4: MaxBlack = 2

        Case 3.75: MaxBlack = 2.25
        Case 3.5: MaxBlack = 2.5
        Case 3.25: MaxBlack = 2.75
        Case 3: MaxBlack = 3
        
        Case 2.75: MaxBlack = 3.25
        Case 2.5: MaxBlack = 3.5
        Case 2.25: MaxBlack = 3.75
        Case 2: MaxBlack = 4
        
        Case 1.75: MaxBlack = 4.25
        Case 1.5: MaxBlack = 4.5
        Case 1.25: MaxBlack = 4.75
        Case 1: MaxBlack = 5
        
        Case Else: GoTo AvoidBlackX1
    End Select
    
    For xBlack = -1 To (MaxBlack - 1) Step 0.25
        For yBlack = 0 To (ScreenHeight / 32)
            Dummy = BitBlt(PicTemp.hdc, (xBlack * 32), (yBlack * 32), TileSize, TileSize, PicObject.hdc, (32 * 79), 0, SRCCOPY)
        Next yBlack
    Next xBlack
AvoidBlackX1:
    
ElseIf Char.X > 45 Then

    Select Case Char.X 'Find the char's x position and invert it for drawing the black
        Case 45.25: MaxBlack = 0.25
        Case 45.5: MaxBlack = 0.5
        Case 45.75: MaxBlack = 0.75
        Case 46: MaxBlack = 1

        Case 46.25: MaxBlack = 1.25
        Case 46.5: MaxBlack = 1.5
        Case 46.75: MaxBlack = 1.75
        Case 47: MaxBlack = 2

        Case 47.25: MaxBlack = 2.25
        Case 47.5: MaxBlack = 2.5
        Case 47.75: MaxBlack = 2.75
        Case 48: MaxBlack = 3

        Case 48.25: MaxBlack = 3.25
        Case 48.5: MaxBlack = 3.5
        Case 48.75: MaxBlack = 3.75
        Case 49: MaxBlack = 4

        Case 49.25: MaxBlack = 4.25
        Case 49.5: MaxBlack = 4.5
        Case 49.75: MaxBlack = 4.75
        Case 50: MaxBlack = 5

        Case Else: GoTo AvoidBlackX2
    End Select

    For xBlack = ((ScreenWidth / 32) - MaxBlack) To (ScreenWidth / 32) Step 0.25
        For yBlack = 0 To (ScreenHeight / 32)
            Dummy = BitBlt(PicTemp.hdc, (xBlack * 32), (yBlack * 32), TileSize, TileSize, PicObject.hdc, (32 * 79), 0, SRCCOPY)
        Next yBlack
    Next xBlack
AvoidBlackX2:

End If

If Char.Y < 6 Then

    Select Case Char.Y 'Find the char's y position and invert it for drawing the black
        Case 5.75: MaxBlack = 0.25
        Case 5.5: MaxBlack = 0.5
        Case 5.25: MaxBlack = 0.75
        Case 5: MaxBlack = 1

        Case 4.75: MaxBlack = 1.25
        Case 4.5: MaxBlack = 1.5
        Case 4.25: MaxBlack = 1.75
        Case 4: MaxBlack = 2

        Case 3.75: MaxBlack = 2.25
        Case 3.5: MaxBlack = 2.5
        Case 3.25: MaxBlack = 2.75
        Case 3: MaxBlack = 3
        
        Case 2.75: MaxBlack = 3.25
        Case 2.5: MaxBlack = 3.5
        Case 2.25: MaxBlack = 3.75
        Case 2: MaxBlack = 4
        
        Case 1.75: MaxBlack = 4.25
        Case 1.5: MaxBlack = 4.5
        Case 1.25: MaxBlack = 4.75
        Case 1: MaxBlack = 5
        
        Case Else: GoTo AvoidBlackY1
    End Select
    
    For xBlack = 0 To (ScreenWidth / 32)
        For yBlack = -1 To (MaxBlack - 1) Step 0.25
            Dummy = BitBlt(PicTemp.hdc, (xBlack * 32), (yBlack * 32), TileSize, TileSize, PicObject.hdc, (32 * 79), 0, SRCCOPY)
        Next yBlack
    Next xBlack
AvoidBlackY1:

ElseIf Char.Y > 45 Then

    Select Case Char.Y 'Find the char's Y position and invert it for drawing the black
        Case 45.25: MaxBlack = 0.25
        Case 45.5: MaxBlack = 0.5
        Case 45.75: MaxBlack = 0.75
        Case 46: MaxBlack = 1

        Case 46.25: MaxBlack = 1.25
        Case 46.5: MaxBlack = 1.5
        Case 46.75: MaxBlack = 1.75
        Case 47: MaxBlack = 2

        Case 47.25: MaxBlack = 2.25
        Case 47.5: MaxBlack = 2.5
        Case 47.75: MaxBlack = 2.75
        Case 48: MaxBlack = 3

        Case 48.25: MaxBlack = 3.25
        Case 48.5: MaxBlack = 3.5
        Case 48.75: MaxBlack = 3.75
        Case 49: MaxBlack = 4

        Case 49.25: MaxBlack = 4.25
        Case 49.5: MaxBlack = 4.5
        Case 49.75: MaxBlack = 4.75
        Case 50: MaxBlack = 5

        Case Else: GoTo AvoidBlackY2
    End Select

    For xBlack = 0 To (ScreenWidth / 32)
        For yBlack = ((ScreenHeight / 32) - (MaxBlack - 1)) To (ScreenHeight / 32) Step 0.25
            Dummy = BitBlt(PicTemp.hdc, (xBlack * 32), (yBlack * 32), TileSize, TileSize, PicObject.hdc, (32 * 79), 0, SRCCOPY)
        Next yBlack
    Next xBlack
AvoidBlackY2:

End If

'[Draw the Local Char]
Dummy = BitBlt(Char.bkg.hdc, (ScreenWidth / 2) - (Char.Width / 2), (ScreenHeight / 2) - (Char.Height / 2), Char.Width, Char.Height, Char.mask.hdc, Char.SprX, (Char.SpriteNum * 64), SRCAND)
Dummy = BitBlt(Char.bkg.hdc, (ScreenWidth / 2) - (Char.Width / 2), (ScreenHeight / 2) - (Char.Height / 2), Char.Width, Char.Height, Char.src.hdc, Char.SprX, (Char.SpriteNum * 64), SRCINVERT)

'[Draw All Other Players]
For PDrawLoop = 1 To MAXPLAYERS 'Draw all other players within viscinity
    If Player(PDrawLoop).X > (MapOffSetX / 32) - 2 And Player(PDrawLoop).X < ((MapOffSetX + ScreenWidth) / 32) + 2 And Player(PDrawLoop).Y > (MapOffSetY / 32) - 2 And Player(PDrawLoop).Y < ((MapOffSetY + ScreenHeight) / 32) + 2 And Player(PDrawLoop).IsConnected = True Then
        Dummy = BitBlt(PicTemp.hdc, (((Player(PDrawLoop).X - 1) * 32) - MapOffSetX), ((((Player(PDrawLoop).Y - 1) * 32) - TileSize) - MapOffSetY), TileSize, (TileSize * 2), PicSpriteMask.hdc, Player(PDrawLoop).SprX, (Player(PDrawLoop).SpriteNum * 64), SRCAND)
        Dummy = BitBlt(PicTemp.hdc, (((Player(PDrawLoop).X - 1) * 32) - MapOffSetX), ((((Player(PDrawLoop).Y - 1) * 32) - TileSize) - MapOffSetY), TileSize, (TileSize * 2), PicSprite.hdc, Player(PDrawLoop).SprX, (Player(PDrawLoop).SpriteNum * 64), SRCINVERT)
    End If
Next PDrawLoop


'***Check for other players' Heads***
    '[Local Char]
If Char.Y < 2 Then: GoTo AvoidHeadBlit
For HeadLoop = 1 To MAXPLAYERS 'Check for players 1 above the local char
    If Player(HeadLoop).IsConnected = True Then
        If Char.Y > Player(HeadLoop).Y And Char.Y < Player(HeadLoop).Y + 2 And Char.X > Player(HeadLoop).X - 2 And Char.X < Player(HeadLoop).X + 2 Then
            Dummy = BitBlt(Char.bkg.hdc, (ScreenWidth / 2) - (Char.Width / 2), (ScreenHeight / 2) - (Char.Height / 2), TileSize, TileSize, Char.mask.hdc, Char.SprX, (Char.SpriteNum * 64), SRCAND)
            Dummy = BitBlt(Char.bkg.hdc, (ScreenWidth / 2) - (Char.Width / 2), (ScreenHeight / 2) - (Char.Height / 2), TileSize, TileSize, Char.src.hdc, Char.SprX, (Char.SpriteNum * 64), SRCINVERT)
            Exit For
        End If
    End If
Next HeadLoop

AvoidHeadBlit:

    '[Other Players]
For PDrawLoop = 1 To MAXPLAYERS 'Draw all other players within viscinity

    If Player(PDrawLoop).Y < 2 Then GoTo AvoidPlayerHeadBlit:
    For CheckIfOver = 1 To MAXPLAYERS
        If Player(CheckIfOver).IsConnected = True Then
            If Player(PDrawLoop).Y > Player(CheckIfOver).Y And Player(PDrawLoop).Y < Player(CheckIfOver).Y + 2 And Player(PDrawLoop).X > Player(CheckIfOver).X - 1 And Player(PDrawLoop).X < Player(CheckIfOver).X + 1 Then
                Dummy = BitBlt(PicTemp.hdc, (((Player(PDrawLoop).X - 1) * 32) - MapOffSetX), ((((Player(PDrawLoop).Y - 1) * 32) - TileSize) - MapOffSetY), TileSize, TileSize, PicSpriteMask.hdc, Player(PDrawLoop).SprX, (Player(PDrawLoop).SpriteNum * 64), SRCAND)
                Dummy = BitBlt(PicTemp.hdc, (((Player(PDrawLoop).X - 1) * 32) - MapOffSetX), ((((Player(PDrawLoop).Y - 1) * 32) - TileSize) - MapOffSetY), TileSize, TileSize, PicSprite.hdc, Player(PDrawLoop).SprX, (Player(PDrawLoop).SpriteNum * 64), SRCINVERT)
            End If
        End If
    Next CheckIfOver
    
AvoidPlayerHeadBlit:

Next PDrawLoop
'***********************************************

'[Map Top Layer Blit]
Dummy = BitBlt(PicTemp.hdc, 0, 0, ScreenWidth, ScreenHeight, PicTopMask.hdc, MapOffSetX, MapOffSetY, SRCAND)
Dummy = BitBlt(PicTemp.hdc, 0, 0, ScreenWidth, ScreenHeight, PicTopWork.hdc, MapOffSetX, MapOffSetY, SRCINVERT)

'***Check for Underlays***
    '[Local Char]
If Char.Y < 2 Then: GoTo AvoidCharUnderlay
If MapHolder(Char.X, Char.Y - 1).TIsUnderLay = True Then
    Dummy = BitBlt(Char.bkg.hdc, (ScreenWidth / 2) - (Char.Width / 2), (ScreenHeight / 2) - (Char.Height / 2), TileSize, TileSize, Char.mask.hdc, Char.SprX, (Char.SpriteNum * 64), SRCAND)
    Dummy = BitBlt(Char.bkg.hdc, (ScreenWidth / 2) - (Char.Width / 2), (ScreenHeight / 2) - (Char.Height / 2), TileSize, TileSize, Char.src.hdc, Char.SprX, (Char.SpriteNum * 64), SRCINVERT)
End If
AvoidCharUnderlay:

    '[Players]
For PDrawLoop = 1 To MAXPLAYERS 'Draw all other players within viscinity
    If Player(PDrawLoop).Y < 2 Then GoTo AvoidPlayerUnderlay:
    If MapHolder(Player(PDrawLoop).X, Player(PDrawLoop).Y - 1).TIsUnderLay = True Then
        Dummy = BitBlt(PicTemp.hdc, (((Player(PDrawLoop).X - 1) * 32) - MapOffSetX), ((((Player(PDrawLoop).Y - 1) * 32) - TileSize) - MapOffSetY), TileSize, TileSize, PicSpriteMask.hdc, Player(PDrawLoop).SprX, (Player(PDrawLoop).SpriteNum * 64), SRCAND)
        Dummy = BitBlt(PicTemp.hdc, (((Player(PDrawLoop).X - 1) * 32) - MapOffSetX), ((((Player(PDrawLoop).Y - 1) * 32) - TileSize) - MapOffSetY), TileSize, TileSize, PicSprite.hdc, Player(PDrawLoop).SprX, (Player(PDrawLoop).SpriteNum * 64), SRCINVERT)
    End If
Next PDrawLoop
AvoidPlayerUnderlay:

'[Main View Blit from Temporary hDC]
Dummy = BitBlt(picMain.hdc, 0, 0, ScreenWidth, ScreenHeight, PicTemp.hdc, 0, 0, SRCCOPY)

End Sub

'**********************************************************
'******************** MAPPING / MUSIC *********************
'**********************************************************

Private Sub Map_Draw(MapToOpen As String)

Dim Dummy As Long
Dim Tile As Integer
Dim X As Integer, Y As Integer
Dim PixelX As Integer, PixelY As Integer
Dim RecordNum As Integer

MapFile = PathMap + MapToOpen 'File to open

Close #1 'Close previous map
Open MapFile For Random As #1 Len = Len(MapHolder(50, 50))
'Open up the map file into MapRecord

PixelX = 0
PixelY = 0
X = 1
Y = 1

For Y = 1 To 50 Step 1
    PixelX = 0
    For X = 1 To 50 Step 1

        RecordNum = (X * (MaxCoord + 1)) + Y
        Get #1, RecordNum, MapHolder(X, Y) 'Get the data from MapRecord
        Dummy = BitBlt(PicBottomWork.hdc, PixelX, PixelY, TileSize, TileSize, PicTile.hdc, (MapHolder(X, Y).BTile - 1) * 32, 0, SRCCOPY)
        Dummy = BitBlt(PicTopMask.hdc, PixelX, PixelY, TileSize, TileSize, PicObjMask.hdc, (MapHolder(X, Y).TTile - 1) * 32, 0, SRCCOPY)
        Dummy = BitBlt(PicTopWork.hdc, PixelX, PixelY, TileSize, TileSize, PicObject.hdc, (MapHolder(X, Y).TTile - 1) * 32, 0, SRCCOPY)
        
        PixelX = PixelX + 32
    Next X
    PixelY = PixelY + 32
Next Y

End Sub

Private Sub Music_Play(FileName As String)
Dim FullFilename As String

FullFilename = PathMusic + FileName
MMControl.FileName = FullFilename
MMControl.Wait = True

If Not MMControl.Mode = mciModeNotOpen Then
    MMControl.Command = "Close"
End If

MMControl.DeviceType = "Sequencer"
MMControl.Command = "Open"
MMControl.Command = "Play"

End Sub

Private Sub Sound_Play(FileName As String)
Dim Dummy

Dummy = sndPlaySound("", SND_ASYNC Or SND_NODEFAULT)
Dummy = sndPlaySound(PathSound + FileName, SND_ASYNC Or SND_NODEFAULT)

End Sub

'**********************************************************
'************************ ANIMATION ***********************
'**********************************************************

Private Sub MoveSprite(Direction As String)

Dim AnimLoop As Integer
Dim MoveChar
Dim StayPut
Dim SendHolder

Select Case Direction
    Case "Down"
        Char.Direction = 0
        Call CheckDown
        If DownIsSolid = False Then
            GoTo MoveChar
        Else: GoTo StayPut: End If
        
    Case "Up"
        Char.Direction = 96
        Call CheckUp
        If UpIsSolid = False Then
            GoTo MoveChar
        Else: GoTo StayPut: End If
        
    Case "Left"
        Char.Direction = 192
        Call CheckLeft
        If LeftIsSolid = False Then
            GoTo MoveChar
        Else: GoTo StayPut: End If
        
    Case "Right"
        Char.Direction = 288
        Call CheckRight
        If RightIsSolid = False Then
            GoTo MoveChar
        Else: GoTo StayPut: End If
End Select

MoveChar:

KeyBusy = True
    
For AnimLoop = 1 To 4 Step 1 'Loop animation 4 times
    
    Select Case Char.Frame
        Case 1
            Char.Frame = 2
            Char.SprX = Char.Direction + 64
        Case 2
            Char.Frame = 3
            Char.SprX = Char.Direction + 32
        Case 3
            Char.Frame = 4
            Char.SprX = Char.Direction + 0
        Case 4
            Char.Frame = 1
            Char.SprX = Char.Direction + 32
    End Select
    
    Select Case Direction
        Case "Down"
            Char.Y = Char.Y + (MoveSize / 32)
        Case "Up"
            Char.Y = Char.Y - (MoveSize / 32)
        Case "Right"
            Char.X = Char.X + (MoveSize / 32)
        Case "Left"
            Char.X = Char.X - (MoveSize / 32)
    End Select
    
    MapOffSetX = ((Char.X - 1) * 32) - (ScreenWidth / 2) + 16
    MapOffSetY = ((Char.Y - 1) * 32) - (ScreenHeight / 2)
    
    Call FuncTimeOut(0.2)
    
Next AnimLoop
lblXY.Caption = Str(Char.X) + "," + Str(Char.Y)

'[Client --> Server movement Selection and position save]
Select Case Direction
    Case "Down"
        SendHolder = "Move " + "D" + "§"
        DoEvents: WinSock.SendData SendHolder: DoEvents
        CharInit.Y = Char.Y
    Case "Up"
        SendHolder = "Move " + "U" + "§"
        DoEvents: WinSock.SendData SendHolder: DoEvents
        CharInit.Y = Char.Y
    Case "Right"
        SendHolder = "Move " + "R" + "§"
        DoEvents: WinSock.SendData SendHolder: DoEvents
        CharInit.X = Char.X
    Case "Left"
        SendHolder = "Move " + "L" + "§"
        DoEvents: WinSock.SendData SendHolder: DoEvents
        CharInit.X = Char.X
End Select

KeyBusy = False
Exit Sub
    
StayPut:
    Char.Frame = 1
    Char.SprX = Char.Direction + 32
    Call CharTimeOut(0.2)
    Exit Sub

End Sub

Private Sub CheckLeft()

If Char.X <= 1 Then: LeftIsSolid = True: Exit Sub
    
If MapHolder(Char.X - 1, Char.Y).BIsWall = True Or MapHolder(Char.X - 1, Char.Y).TIsWall = True Then
    LeftIsSolid = True
Else
    Dim WallLoop
    For WallLoop = 1 To MAXPLAYERS
        If Player(WallLoop).IsConnected And (Char.X - 1 = Player(WallLoop).X) And Char.Y = Player(WallLoop).Y Then
            LeftIsSolid = True: Exit For
        Else
            LeftIsSolid = False
        End If
    Next WallLoop
End If

End Sub

Private Sub CheckRight()

If Char.X = (MapSize / 32) Then: RightIsSolid = True: Exit Sub

If MapHolder(Char.X + 1, Char.Y).BIsWall = True Or MapHolder(Char.X + 1, Char.Y).TIsWall = True Then
    RightIsSolid = True
Else
    Dim WallLoop
    For WallLoop = 1 To MAXPLAYERS
        If Player(WallLoop).IsConnected And (Char.X + 1 = Player(WallLoop).X) And Char.Y = Player(WallLoop).Y Then
            RightIsSolid = True: Exit For
        Else
            RightIsSolid = False
        End If
    Next WallLoop
End If

End Sub

Private Sub CheckUp()

If Char.Y <= 1 Then: UpIsSolid = True: Exit Sub
    
If MapHolder(Char.X, Char.Y - 1).BIsWall = True Or MapHolder(Char.X, Char.Y - 1).TIsWall = True Then
    UpIsSolid = True
Else
    Dim WallLoop
    For WallLoop = 1 To MAXPLAYERS
        If Player(WallLoop).IsConnected And (Char.Y - 1 = Player(WallLoop).Y) And Char.X = Player(WallLoop).X Then
            UpIsSolid = True: Exit For
        Else
            UpIsSolid = False
        End If
    Next WallLoop
End If

End Sub

Private Sub CheckDown()

If Char.Y >= (MapSize / 32) Then: DownIsSolid = True: Exit Sub
    
If MapHolder(Char.X, Char.Y + 1).BIsWall = True Or MapHolder(Char.X, Char.Y + 1).TIsWall = True Then
    DownIsSolid = True
Else
    Dim WallLoop
    For WallLoop = 1 To MAXPLAYERS
        If Player(WallLoop).IsConnected And (Char.Y + 1 = Player(WallLoop).Y) And Char.X = Player(WallLoop).X Then
            DownIsSolid = True: Exit For
        Else
            DownIsSolid = False
        End If
    Next WallLoop
End If

End Sub

'**********************************************************
'********************* GAMEPLAY OBJECTS *******************
'**********************************************************

Private Sub Form_Load()

'[INITIALIZE DIRECTX]
frmMain.ScaleMode = 3
'Call dixuInit(dixuInitFullScreen, frmMain, 640, 480, 16)

'[CALLING SUBROUTINES]
Call Var_Init          'Initialize Variables
Call hDC_Init         'Initialize All hDCs
Call Char_Init         'Initialize Char Data
Call Initializations   'initialize Game data

End Sub

Private Sub Form_Unload(Cancel As Integer)

Call DestroyHdcs 'Clear all Mem DCs out of memory
WinSock.Close

If Not MMControl.Mode = mciModeNotOpen Then
    MMControl.Command = "Close"
End If

If GameScreen = "Main" Then
    Put #2, 1, CharInit
    Close #2
End If

'Call dixuDone 'Close full-screen DirectX

End Sub

Private Sub Form_Click()

If GameScreen = "Main" Then: picMain.SetFocus

End Sub

Private Sub Form_DblClick()
    Unload Me 'Failsafe to avoid crashes
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyBusy = True Then: Exit Sub

Select Case KeyCode
    Case KEY_LEFT
        Call MoveSprite("Left")
    Case KEY_RIGHT
        Call MoveSprite("Right")
    Case KEY_UP
        Call MoveSprite("Up")
    Case KEY_DOWN
        Call MoveSprite("Down")
End Select

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyBusy = False Then: Char.SprX = Char.Direction + 32

End Sub

Private Sub lblExit_Click()
    Unload Me
End Sub

Private Sub MMControl_Done(NotifyCode As Integer)

If NotifyCode = 1 Then 'If finish was successful
    MMControl.Command = "Close" 'Close old control
    MMControl.Command = "Open" 'Reopen
    MMControl.Command = "Prev"
    MMControl.Command = "Play"
End If

End Sub

Private Sub txtRec_Change()
SavedWnd = Screen.ActiveControl.hWnd
Dim Num
Num = ScrollText(txtRec, 16000)
End Sub

Private Sub txtRec_GotFocus()
    KeyBusy = True 'Char won't move if cursor is in textbox
End Sub

Private Sub txtRec_LostFocus()
    KeyBusy = False 'Char will move now that the cursor is gone
End Sub

Private Sub txtSend_GotFocus()
    KeyBusy = True 'Char won't move if cursor is in textbox
End Sub

Private Sub txtSend_LostFocus()
    KeyBusy = False 'Char will move now that the cursor is gone
End Sub

Private Sub tmrDraw_Timer()

Call Board_Refresh

End Sub

'**********************************************************
'******************** CHAR SELECT OBJECTS *****************
'**********************************************************

Private Sub picMain_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Dummy

If GameScreen = "CharSel" Then

    If Button = 1 Then
        If Char.SpriteNum = 14 Then
            Char.SpriteNum = 0
        Else: Char.SpriteNum = Char.SpriteNum + 1
        End If
    End If
    
    If Button = 2 Then
        If Char.SpriteNum = 0 Then
            Char.SpriteNum = 14
        Else: Char.SpriteNum = Char.SpriteNum - 1
        End If
    End If
    
    Dummy = BitBlt(PicCharSelect.hdc, 0, 0, TileSize, TileSize * 2, PicSprite.hdc, 32, (Char.SpriteNum * 64), SRCCOPY)

ElseIf GameScreen = "Main" Then
    picMain.SetFocus 'If they're playing the game, set to the mainview
End If

End Sub

Private Sub lblEnter_Click()
Dim SendPos As String, SendImg As String, SendName As String

SendPos = "InitPosition" + Str(Char.X) + Str(Char.Y) + "§"
DoEvents
WinSock.SendData SendPos
DoEvents

SendImg = "InitImage" + Str(Char.SpriteNum) + "§"
DoEvents
WinSock.SendData SendImg
DoEvents

SendName = "Name " + Namo + "§"
DoEvents
WinSock.SendData SendName
DoEvents

End Sub

Private Sub tmrPlayers_Timer()

If GameScreen = "CharSel" Then
    Dim Dummy
    Dummy = BitBlt(picMain.hdc, (ScreenWidth / 2) - (Char.Width / 2), (ScreenHeight / 2) - (Char.Height / 2), TileSize, TileSize * 2, PicCharSelect.hdc, 0, 0, SRCCOPY)
ElseIf GameScreen = "Main" Then
    Call Players_MoveLoop
End If

End Sub

'**********************************************************
'******************** CLIENT / SERVER *********************
'**********************************************************

Private Sub lblConnect_Click()
Dim Seconds As Integer

Namo = txtName.Text
Passwordo = txtPassword.Text
lblConnect.Enabled = False

If txtIP.Text <> "" And txtName.Text <> "" Then
    WinSock.Connect txtIP.Text 'Connect to Server w/ IP in textbox
    
    '[Show seconds that connection has been going]
    Seconds = 0
    Do Until WinSock.State = 7
        lblStatus.Caption = "Status: Connecting -" + Str(Seconds) + " seconds"
        Call FuncTimeOut(1)
        Seconds = Seconds + 1
    Loop
    lblStatus = "Status: Connection Successful! Please wait for game to load."
    Call FuncTimeOut(3)
    
Else
    Beep 'Make sound
End If
    
End Sub

Private Sub WinSock_Connect()
Dim DataToSend As String

'[Hide all Connection Controls]
GameScreen = "CharSel"
picTitle.Visible = False
lblConnect.Visible = False
lblEnter.Visible = True
lblIP.Visible = False
lblUserName.Visible = False
lblPassword.Visible = False
lblStatus.Visible = False

txtIP.Visible = False
txtName.Visible = False
txtPassword.Visible = False

'[Character Sprite Animation Values]

Close #2 'Close previous Char file if client crashed
Open Namo + ".chr" For Random As #2 Len = Len(CharInit)
Get #2, 1, CharInit
If CharInit.X = 0 And CharInit.Y = 0 Then
    Char.X = 10
    Char.Y = 10
Else
    Char.X = CharInit.X
    Char.Y = CharInit.Y
End If

MapOffSetX = ((Char.X - 1) * 32) - (ScreenWidth / 2) + 16
MapOffSetY = ((Char.Y - 1) * 32) - (ScreenHeight / 2)

'[Char Selection Values]
picMain.Visible = True 'Make main view visible
tmrPlayers.Enabled = True

Dim Dummy
Dummy = BitBlt(PicCharSelect.hdc, 0, 0, TileSize, TileSize * 2, PicSprite.hdc, 32, (Char.SpriteNum * 64), SRCCOPY)

End Sub

Private Sub WinSock_DataArrival(ByVal bytesTotal As Long)

Static DataRead As String
Dim LineString As String
Dim CommandString As String

Dim PIndex
Dim PString

WinSock.GetData DataRead

While InStr(DataRead, "§") 'If there's still the EOF left in DataRead

    CommandString = Left(DataRead, InStr(DataRead, " ") - 1) 'Get command
    
    LineString = Mid(DataRead, InStr(DataRead, " ") + 1) 'Get data w/ EOF
    LineString = Left(LineString, InStr(LineString, "§") - 1) 'Remove EOF
    
    PIndex = Left(LineString, InStr(LineString, " ") - 1) 'Player Index
    PIndex = Int(Val(PIndex))
    PString = Mid(LineString, InStr(LineString, " ") + 1) 'Player Text
    
    Select Case CommandString
    
        '[Communicaton and Movement]
        Case "Chat"
            
            txtRec.Text = txtRec.Text + HR + Player(PIndex).Name + ": " + PString
            'SavedWnd = Screen.ActiveControl.hwnd
            'Dim NumScroll
            'NumScroll = ScrollText(txtRec, 60)
        
        Case "Emote"
        
        Case "Move"
        
            Select Case PString
                Case "D"
                    Player(PIndex).MoveScript = Player(PIndex).MoveScript + "D"
                Case "U"
                    Player(PIndex).MoveScript = Player(PIndex).MoveScript + "U"
                Case "R"
                    Player(PIndex).MoveScript = Player(PIndex).MoveScript + "R"
                Case "L"
                    Player(PIndex).MoveScript = Player(PIndex).MoveScript + "L"
            End Select
            
        '[Server-to-Client ONLY Messages]
        Case "Start"
            
            '[Make all Char Select stuff invis]
            GameScreen = "Main"
            lblEnter.Visible = False
            lblExit.Left = 15
            lblExit.Top = 15
            
            Me.Picture = LoadPicture(PathArchive + "Backdrop.bmp")
            tmrDraw.Enabled = True  'Start Timer
            Player(PIndex).Name = Namo
            txtSend.Visible = True
            txtRec.Visible = True
            Call Sound_Play("Welcome.wav")
            'Call Music_Play("Ins03.Mid")
            txtRec.Text = txtRec.Text + HR + PString
        
        Case "Text"
        
            If PString = "hears another realm beckoning....and leaves this world." Then
                Player(PIndex).IsConnected = False
                txtRec.Text = txtRec.Text + HR + Player(PIndex).Name + " " + PString
            ElseIf Left(PString, 8) = "<Server>" Then
                txtRec.Text = txtRec.Text + HR + PString
            End If
        
        '[Player Initializations]
        Case "InitPosition"
        
            Player(PIndex).X = Int(Val(Left(PString, InStr(PString, " ") - 1)))
            Player(PIndex).Y = Int(Val(Mid(PString, InStr(PString, " ") + 1)))
            Player(PIndex).SprX = 32
            
        Case "InitImage"
        
            Player(PIndex).SpriteNum = Int(Val(PString))
            
        Case "Name"
        
            txtRec.Text = txtRec.Text + HR + PString + " is awake within this world."
            Player(PIndex).Name = PString
            Player(PIndex).Frame = 1
            Player(PIndex).IsConnected = True

        '[Enemies and Attacking]
        Case "Attack"
    
        Case Else
        txtRec.Text = txtRec.Text + HR + "Illigal Server Data Sent! Please submit a bug report to Xian. Thanks!"
    
    End Select

    DataRead = Mid$(DataRead, InStr(DataRead, "§") + 1) 'Remove EOF + data

Wend

End Sub

Private Sub WinSock_Close()
    txtRec.Text = txtRec.Text + HR + "Server Dead. Closing Client."
    Call FuncTimeOut(8)
    Unload frmMain 'Close Client
End Sub

Private Sub txtSend_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Then

    Dim SendHolder As String
    SendHolder = "Chat " + txtSend.Text + "§"
    DoEvents: WinSock.SendData SendHolder: DoEvents 'Send the string variable
    
    txtSend.Text = "" 'Clear the sendbox

End If

End Sub

Private Sub txtSend_KeyPress(KeyAscii As Integer)
    
Select Case KeyAscii
    Case 167
        KeyAscii = False 'Stop the key from being pressed (§)
    Case 13
        KeyAscii = 0 'Stop <enter>'s from beeping
End Select

End Sub

Private Sub Players_MoveLoop()
Dim MoveLoop As Integer

Dim PIndex As Integer 'The index of the player for the drawing routine
Dim CharMove
Dim AvoidMove

For MoveLoop = 1 To MAXPLAYERS 'Rotate from 1 to Maximum # of players
    
    If Player(MoveLoop).IsConnected = True And Player(MoveLoop).Busy = False Then
        
        Select Case Left(Player(MoveLoop).MoveScript, 1) 'Leftmost letter
            Case "D"
                Player(MoveLoop).Direction = 0
                Player(MoveLoop).Busy = True
                Call Players_MoveAnim(MoveLoop)
            Case "U"
                Player(MoveLoop).Direction = 96
                Player(MoveLoop).Busy = True
                Call Players_MoveAnim(MoveLoop)
            Case "L"
                Player(MoveLoop).Direction = 192
                Player(MoveLoop).Busy = True
                Call Players_MoveAnim(MoveLoop)
            Case "R"
                Player(MoveLoop).Direction = 288
                Player(MoveLoop).Busy = True
                Call Players_MoveAnim(MoveLoop)
        End Select
        
    End If
    
Next MoveLoop

End Sub

Private Sub Players_MoveAnim(PIndex As Integer)
Dim AnimLoop
Dim MovePause As Single

Select Case Len(Player(PIndex).MoveScript)
    Case 1: MovePause = 0.2
    Case 2: MovePause = 0.15
    Case 3: MovePause = 0.1
    Case 4: MovePause = 0.05
    Case 5: MovePause = 0.05
    Case 6: MovePause = 0.05
    Case Else
        If Len(Player(PIndex).MoveScript) > 6 Then: MovePause = 0
End Select

For AnimLoop = 1 To 4 Step 1 'Loop animation 4 times
    
    Select Case Player(PIndex).Frame
        Case 1
            Player(PIndex).Frame = 2
            Player(PIndex).SprX = Player(PIndex).Direction + 64
        Case 2
            Player(PIndex).Frame = 3
            Player(PIndex).SprX = Player(PIndex).Direction + 32
        Case 3
            Player(PIndex).Frame = 4
            Player(PIndex).SprX = Player(PIndex).Direction + 0
        Case 4
            Player(PIndex).Frame = 1
            Player(PIndex).SprX = Player(PIndex).Direction + 32
    End Select
    
    Select Case Player(PIndex).Direction 'Change Player's X/Y
        Case 0
            Player(PIndex).Y = Player(PIndex).Y + (MoveSize / 32) 'Change Pos
        Case 96
            Player(PIndex).Y = Player(PIndex).Y - (MoveSize / 32)
        Case 192
            Player(PIndex).X = Player(PIndex).X - (MoveSize / 32)
        Case 288
            Player(PIndex).X = Player(PIndex).X + (MoveSize / 32)
    End Select
    
    Call FuncTimeOut(MovePause) 'Pause in animation
    
Next AnimLoop

Player(PIndex).MoveScript = Mid(Player(PIndex).MoveScript, 2) 'Append
Player(PIndex).Busy = False 'Let it loop to another animation, if needed

End Sub

Private Sub WinSock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Dim Msg As String

Msg = "An Error! You may want to restart Insomnia." + HR
Msg = Msg + "Please send Lord Xian this info:" + HR + HR
Msg = Msg + "   Winsock Error" + HR
Msg = Msg + "   Error Number: " + Str(Number) + HR
Msg = Msg + "   Description: " + Description + HR
Msg = Msg + "   S-Code: " + Str(Scode) + HR
Msg = Msg + "   Source: " + Source + HR
Msg = Msg + "   Help File: " + HelpFile + HR
txtRec.Text = Msg + txtRec.Text
    
End Sub
