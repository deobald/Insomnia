VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Begin VB.Form frmMap 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Insomnia Map Editor"
   ClientHeight    =   6165
   ClientLeft      =   1800
   ClientTop       =   1725
   ClientWidth     =   8655
   Icon            =   "frmMap.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   411
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   577
   Visible         =   0   'False
   Begin ComctlLib.Toolbar ToolBar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList"
      _Version        =   327680
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   6
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Create a New Map File"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Open a map file"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Save this file"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Save this file under a new name"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
      EndProperty
      MouseIcon       =   "frmMap.frx":030A
   End
   Begin VB.OptionButton optBottomDelete 
      Caption         =   "Delete Tile"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   1500
      TabIndex        =   22
      ToolTipText     =   "Delete Ground Tiles"
      Top             =   3075
      Width           =   1365
   End
   Begin VB.Frame fraToolBar 
      Height          =   90
      Left            =   -75
      TabIndex        =   21
      Top             =   375
      Width           =   10365
   End
   Begin VB.OptionButton optMonsterDelete 
      Caption         =   "Delete NPC"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Left            =   1500
      TabIndex        =   19
      ToolTipText     =   "Delete Monsters"
      Top             =   5475
      Width           =   1365
   End
   Begin VB.OptionButton optObjDelete 
      Caption         =   "Delete Obj."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Left            =   1500
      TabIndex        =   18
      ToolTipText     =   "Delete Objects"
      Top             =   4125
      Width           =   1365
   End
   Begin VB.OptionButton optMonster 
      Caption         =   "Use NPCs"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   1500
      TabIndex        =   17
      ToolTipText     =   "Click Here to use this as your current tile"
      Top             =   5175
      Width           =   1365
   End
   Begin VB.OptionButton optTop 
      Caption         =   "Use Objects"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   1500
      TabIndex        =   16
      ToolTipText     =   "Click Here to use this as your current tile"
      Top             =   3825
      Width           =   1365
   End
   Begin VB.OptionButton optBottom 
      Caption         =   "Use Tiles"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   1500
      TabIndex        =   15
      ToolTipText     =   "Click Here to use this as your current tile"
      Top             =   2775
      Value           =   -1  'True
      Width           =   1290
   End
   Begin VB.Frame fraPosition 
      Caption         =   "X/Y Position"
      Height          =   690
      Left            =   150
      TabIndex        =   13
      ToolTipText     =   "Cursor position, relative to map co-ordinates"
      Top             =   675
      Width           =   2790
      Begin VB.Label lblCurPosition 
         Height          =   240
         Left            =   1125
         TabIndex        =   14
         ToolTipText     =   "Cursor position, relative to map co-ordinates"
         Top             =   300
         Width           =   615
      End
   End
   Begin VB.Frame frmCurrentTile 
      Caption         =   "Current Tile"
      Height          =   990
      Left            =   150
      TabIndex        =   9
      Top             =   1425
      Width           =   2790
      Begin VB.PictureBox picCurrentTile 
         Height          =   540
         Left            =   1200
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   10
         ToolTipText     =   "This is the tile (image) you are using on the map right now"
         Top             =   300
         Width           =   540
         Begin VB.Label lblCurrentTile 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   765
            Left            =   -150
            TabIndex        =   20
            ToolTipText     =   "Click on the map to delete tiles"
            Top             =   0
            Visible         =   0   'False
            Width           =   765
         End
      End
   End
   Begin VB.VScrollBar VertBar 
      Height          =   4815
      LargeChange     =   10
      Left            =   8250
      Max             =   41
      Min             =   1
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   975
      Value           =   1
      Width           =   240
   End
   Begin VB.HScrollBar HorizBar 
      Height          =   240
      LargeChange     =   10
      Left            =   3450
      Max             =   41
      Min             =   1
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   5775
      Value           =   1
      Width           =   4815
   End
   Begin VB.PictureBox picMain 
      Height          =   4800
      Left            =   3450
      ScaleHeight     =   316
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   316
      TabIndex        =   0
      Top             =   975
      Width           =   4800
   End
   Begin VB.Frame fraNPC 
      Caption         =   "Monsters/NPCs"
      Enabled         =   0   'False
      Height          =   990
      Left            =   150
      TabIndex        =   6
      Top             =   4875
      Width           =   2790
      Begin VB.PictureBox picMonster 
         Height          =   540
         Left            =   150
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   8
         ToolTipText     =   "Click this box with either mouse button to rotate tiles"
         Top             =   300
         Width           =   540
      End
   End
   Begin VB.Frame fraObjects 
      Caption         =   "Objects"
      Height          =   1290
      Left            =   150
      TabIndex        =   5
      Top             =   3525
      Width           =   2790
      Begin VB.CheckBox chkUnderLay 
         Caption         =   "Under-Lay"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   240
         Left            =   1350
         TabIndex        =   24
         ToolTipText     =   "Check here to make the object UNDERNEATH any players/monsters when the map is drawn in the game."
         Top             =   900
         Width           =   1365
      End
      Begin VB.CommandButton btnTopWalls 
         Caption         =   "OFF"
         Height          =   315
         Left            =   750
         TabIndex        =   11
         ToolTipText     =   "Wall (solid) Tile Toggle for Overlays"
         Top             =   300
         Width           =   540
      End
      Begin VB.PictureBox picTTile 
         Height          =   540
         Left            =   150
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   7
         ToolTipText     =   "Click this box with either mouse button to rotate tiles"
         Top             =   300
         Width           =   540
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   3150
      Top             =   1575
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
   End
   Begin VB.Timer tmrPaint 
      Interval        =   1
      Left            =   3150
      Top             =   1125
   End
   Begin VB.Frame fraTiles 
      Caption         =   "Ground Tiles"
      Height          =   990
      Left            =   150
      TabIndex        =   3
      Top             =   2475
      Width           =   2790
      Begin VB.CommandButton btnBottomWalls 
         Caption         =   "OFF"
         Height          =   315
         Left            =   750
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Wall (solid) Tile Toggle for Lower-layer Tiles"
         Top             =   300
         Width           =   540
      End
      Begin VB.PictureBox picBTile 
         Height          =   540
         Left            =   150
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   4
         ToolTipText     =   "Click this box with either mouse button to rotate tiles"
         Top             =   300
         Width           =   540
      End
   End
   Begin ComctlLib.ImageList ImageList 
      Left            =   3150
      Top             =   2100
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327680
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   4
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMap.frx":0326
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMap.frx":0640
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMap.frx":095A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMap.frx":0C74
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu sepFile1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save &As"
      End
      Begin VB.Menu sepFile2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      NegotiatePosition=   2  'Middle
      Begin VB.Menu mnuUsing 
         Caption         =   "&Using the Editor"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About &DreamEditor"
      End
   End
End
Attribute VB_Name = "frmMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'[Map Declares]
Dim MapFile As String 'The map file to open on startup
Dim MapHolder(1 To 50, 1 To 50) As Record
Dim Path As String
Dim TileSize As Integer
Dim HR

'[Bottom and Top layer image holders]
Dim PicTile As tArea 'Tile Hdc for lower-layer images
Dim PicObjects As tArea 'hDC for object images
Dim PicObjMask As tArea 'hDC for object Masks
Dim PicMaskBW As tArea '32*64 hDC for black and white temp tiles

'[Bottom and Top Mem hDCs]
Dim PicBottomWork As tArea 'TileBot tiles workarea
Dim PicTopWork As tArea 'Top objects workarea
Dim PicTopMask As tArea 'Top objects workarea mask
Dim PicTemp As tArea

'[Hold current properties for Top, Bottom, and NPCs]
Dim TileBot As ImageProperties 'Properties of the bottom, top and
Dim TileTop As ImageProperties 'NPC tiles
Dim Monster As ImageProperties

'[Cursor Stuff]
Dim PicCurBox As tArea 'Cursor box
Dim PicCurMask As tArea 'Cursor box mask
Public CurX, CurY 'Position of cursor on the 32*32 grid
Public CursorVisible As Boolean

Private Sub Initializations()

Dim Dummy

'[Set Variables to values]
TileSize = 32

'[Set paths to game location on Hard Disk]
If (Right(App.Path, 1) <> "\") Then
    Path = App.Path & "\"
        Else
Path = App.Path
End If
'[End of Set Paths]

'[Form Initializations]
frmSplash.Picture = LoadPicture(Path + "BnrLoad.bmp")
frmMap.ScaleMode = 3
HR = Chr(13) + Chr(10)

'[Main View Initializations]
picMain.ScaleMode = 3
picMain.Width = 1600 'ScreenWidth
picMain.Height = 1600 'ScreenHeight
picMain.Cls

'[Mem Hdc to hold Lower-layer maptiles]
PicTile.hDC = 0
PicTile.Left = 0
PicTile.Top = 0
PicTile.Width = TileSize 'Width of the memory holder
PicTile.Height = TileSize 'Height of the memory holder
PicTile.hDC = CreateMemHdc(picMain.hDC, TileSize, TileSize)
Call LoadBmpToHdc(PicTile.hDC, "BottomTiles.bmp")

'[Mem Hdc to hold Lower-layer maptiles]
PicObjects.hDC = 0
PicObjects.Left = 0
PicObjects.Top = 0
PicObjects.Width = TileSize 'Width of the memory holder
PicObjects.Height = TileSize 'Height of the memory holder
PicObjects.hDC = CreateMemHdc(picMain.hDC, TileSize, TileSize)
Call LoadBmpToHdc(PicObjects.hDC, "TopTiles.bmp")

'[Mem Hdc to hold Lower-layer Masks]
PicObjMask.hDC = 0
PicObjMask.Left = 0
PicObjMask.Top = 0
PicObjMask.Width = TileSize 'Width of the memory holder
PicObjMask.Height = TileSize 'Height of the memory holder
PicObjMask.hDC = CreateMemHdc(picMain.hDC, TileSize, TileSize)
Call LoadBmpToHdc(PicObjMask.hDC, "TopTiles.bmp")
Call Mask_Make(PicObjMask, PicObjects, 0, 0, 2560, TileSize)

'[Mem Hdc to hold Black and White temporary tiles]
PicMaskBW.hDC = 0
PicMaskBW.Left = 0
PicMaskBW.Top = 0
PicMaskBW.Width = (TileSize * 2) 'Width of the memory holder
PicMaskBW.Height = TileSize 'Height of the memory holder
PicMaskBW.hDC = CreateMemHdc(picMain.hDC, (TileSize * 2), TileSize)
Call LoadBmpToHdc(PicMaskBW.hDC, "MaskBW.bmp")

'[Mem hDC to hold the map and work on bottom tiles]
picMain.Picture = LoadPicture(Path + "BottomTiles.bmp") 'Set Palette
Dummy = SelectPalette(PicBottomWork.hDC, picMain.Picture.hPal, False)
Dummy = RealizePalette(PicBottomWork.hDC)
picMain.Cls 'Clear palette-setting image
PicBottomWork.hDC = 0
PicBottomWork.Left = 0
PicBottomWork.Top = 0
PicBottomWork.Width = 1600
PicBottomWork.Height = 1600
PicBottomWork.hDC = CreateMemHdc(picMain.hDC, 1600, 1600)

'[Mem hDC to hold the objects and work on top tiles]
picMain.Picture = LoadPicture(Path + "BottomTiles.bmp") 'Set Palette
Dummy = SelectPalette(PicTopWork.hDC, picMain.Picture.hPal, False)
Dummy = RealizePalette(PicTopWork.hDC)
picMain.Cls 'Clear palette-setting image
PicTopWork.hDC = 0
PicTopWork.Left = 0
PicTopWork.Top = 0
PicTopWork.Width = 1600
PicTopWork.Height = 1600
PicTopWork.hDC = CreateMemHdc(picMain.hDC, 1600, 1600)

'[Mem hDC to hold the top tiles mask image for transparency]
picMain.Picture = LoadPicture(Path + "BottomTiles.bmp") 'Set Palette
Dummy = SelectPalette(PicTopMask.hDC, picMain.Picture.hPal, False)
Dummy = RealizePalette(PicTopMask.hDC)
picMain.Cls 'Clear palette-setting image
PicTopMask.hDC = 0
PicTopMask.Left = 0
PicTopMask.Top = 0
PicTopMask.Width = 1600
PicTopMask.Height = 1600
PicTopMask.hDC = CreateMemHdc(picMain.hDC, 1600, 1600)

'[Mem hDC to hold the objects and work on top tiles]
picMain.Picture = LoadPicture(Path + "BottomTiles.bmp") 'Set Palette
Dummy = SelectPalette(PicTemp.hDC, picMain.Picture.hPal, False)
Dummy = RealizePalette(PicTemp.hDC)
picMain.Cls 'Clear palette-setting image
PicTemp.hDC = 0
PicTemp.Left = 0
PicTemp.Top = 0
PicTemp.Width = 320
PicTemp.Height = 320
PicTemp.hDC = CreateMemHdc(picMain.hDC, 320, 320)

'[CursorBox Mem hDC]
PicCurBox.hDC = 0
PicCurBox.Left = 0
PicCurBox.Top = 0
PicCurBox.Width = TileSize 'Width of the memory holder
PicCurBox.Height = TileSize 'Height of the memory holder
PicCurBox.hDC = CreateMemHdc(picMain.hDC, TileSize, TileSize)
Call LoadBmpToHdc(PicCurBox.hDC, "CursorBox.bmp")

'[CursorBox Mask Mem hDC]
PicCurMask.hDC = 0
PicCurMask.Left = 0
PicCurMask.Top = 0
PicCurMask.Width = TileSize 'Width of the memory holder
PicCurMask.Height = TileSize 'Height of the memory holder
PicCurMask.hDC = CreateMemHdc(picMain.hDC, TileSize, TileSize)
Call LoadBmpToHdc(PicCurMask.hDC, "CursorBox.bmp")
Call Mask_Make(PicCurMask, PicCurBox, 0, 0, 32, 32)

'[Return picMain to normal]
picMain.Width = 320 'ScreenWidth
picMain.Height = 320 'ScreenHeight

'[Tile Viewer Inits]
picBTile.ScaleMode = 3
picBTile.ScaleWidth = 32
picBTile.ScaleHeight = 32
TileBot.Tile = 1
TileBot.MaxNumTiles = 80
TileBot.IsWall = False
Dummy = BitBlt(picBTile.hDC, 0, 0, 32, 32, PicTile.hDC, (TileBot.Tile - 1) * 32, 0, SRCCOPY)
Dummy = BitBlt(picCurrentTile.hDC, 0, 0, 32, 32, PicTile.hDC, (TileBot.Tile - 1) * 32, 0, SRCCOPY)

picTTile.ScaleMode = 3
picTTile.ScaleWidth = 32
picTTile.ScaleHeight = 32
TileTop.Tile = 1
TileTop.MaxNumTiles = 80
TileTop.IsWall = False
Dummy = BitBlt(picTTile.hDC, 0, 0, 32, 32, PicObjects.hDC, (TileTop.Tile - 1) * 32, 0, SRCCOPY)

picMonster.ScaleMode = 3
picMonster.ScaleWidth = 32
picMonster.ScaleHeight = 32
Monster.Tile = 1

'[Open a New Map]
Call MakeNewMap

frmMap.Visible = True

End Sub

Private Sub Form_Load()

Call Initializations

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

'If the cursor goes off the viewscreen, make the box invis.
If X <= picMain.Left Or X >= (picMain.Left + picMain.Width) Or Y <= picMain.Top Or Y >= (picMain.Top + picMain.Height) Then
    CursorVisible = False
End If

lblCurPosition.Caption = ""

End Sub

Private Sub Form_Unload(Cancel As Integer)

Close #1 'Close the map file from RAM
Call DestroyHdcs

End Sub

Private Sub btnBottomWalls_Click()

If TileBot.IsWall = False Then 'Toggle bottom layer walls on and off
    TileBot.IsWall = True
    btnBottomWalls.Caption = "ON"
Else: TileBot.IsWall = False: btnBottomWalls.Caption = "OFF"
End If
frmMap.SetFocus

End Sub

Private Sub btnTopWalls_Click()

If TileTop.IsWall = False Then 'Toggle bottom layer walls on and off
    TileTop.IsWall = True
    btnTopWalls.Caption = "ON"
Else: TileTop.IsWall = False: btnTopWalls.Caption = "OFF"
End If
frmMap.SetFocus

End Sub

Private Sub mnuAbout_Click()

Dim Msg
Msg = "Thank you for downloading Insomnia DreamEditor."
Msg = Msg + HR + HR + "Copyright © 1998 OmniSoft All Rights Reserved."
Msg = Msg + HR + "Unpermissable copying of this software in whole or in part"
Msg = Msg + HR + "warrents legal action on the part of OmniSoft."
Msg = Msg + HR + HR + "Developed by:"
Msg = Msg + HR + "Lord Xian ---- Programming"
Msg = Msg + HR + "Sabin -------- Graphics"
Msg = Msg + HR + HR + HR + "OMNISOFT" + HR + "Web: OmniSoft.os.ca" + HR + "E-mail: OmniSoft@os.ca"
MsgBox Msg, vbInformation, "Insomnia DreamEditor Credits"

End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuNew_Click()
    Call MakeNewMap
End Sub

Private Sub MakeNewMap()

Close #1
picMain.Cls 'Clear everything off of screen

Open "New.map" For Random As #1 Len = Len(MapHolder(50, 50))
MapFile = Path + "new.map"
frmMap.Caption = "Insomnia Map Editor - [Generating New Map...]"

Dim Y As Integer, X As Integer
Dim PixelY, PixelX
Dim RecordNum As Long, Dummy
PixelX = 0
PixelY = 0

For Y = 1 To 50
    
    PixelX = 0
    
    For X = 1 To 50

        Dummy = BitBlt(PicBottomWork.hDC, PixelX, PixelY, 32, 32, PicTile.hDC, 0, 0, SRCCOPY)
        MapHolder(X, Y).BIsWall = False
        MapHolder(X, Y).BTile = 1

        PixelX = PixelX + 32

    Next X
    
    PixelY = PixelY + 32
    
Next Y

PixelX = 0
PixelY = 0
For Y = 1 To 50
    PixelX = 0
    For X = 1 To 50
    
        Dummy = BitBlt(PicTopWork.hDC, PixelX, PixelY, 32, 32, PicMaskBW.hDC, 0, 0, SRCCOPY)
        Dummy = BitBlt(PicTopMask.hDC, PixelX, PixelY, 32, 32, PicMaskBW.hDC, 32, 0, SRCCOPY)
        MapHolder(X, Y).TIsWall = False
        MapHolder(X, Y).TTile = 80
        
        PixelX = PixelX + 32
    Next X
    PixelY = PixelY + 32
Next Y
frmMap.Caption = "Insomnia Map Editor - [...New.map]"

VertBar.Value = 1
HorizBar.Value = 1

End Sub

Private Sub OpenMap()

Dim ErrHandler
CommonDialog.CancelError = True
On Error GoTo ErrHandler
CommonDialog.Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNOverwritePrompt
CommonDialog.Filter = "Map Files (*.MAP)|*.MAP|MapData Files (*.DAT)|*.DAT|All Files (*.*)|*.*"
CommonDialog.FilterIndex = 1 'Set to .MAP as default
CommonDialog.ShowOpen 'Action = 1 'Open a "Open File" box
MapFile = CommonDialog.filename 'MapFile equals the filename chosen

Close #1
Open MapFile For Random As #1 Len = Len(MapHolder(50, 50))
frmMap.Caption = "Insomnia Map Editor - [Opening File...]"
'Open selected file and change form caption accordingly.

Dim Y As Integer, X As Integer
Dim PixelY, PixelX
Dim RecordNum As Long
Dim Dummy

PixelX = 0
PixelY = 0
For Y = 1 To 50
    PixelX = 0
    For X = 1 To 50

        RecordNum = (X * (MaxCoord + 1)) + Y
        Get #1, RecordNum, MapHolder(X, Y) 'Get the data from MapRecord
        Dummy = BitBlt(PicBottomWork.hDC, PixelX, PixelY, 32, 32, PicTile.hDC, (MapHolder(X, Y).BTile - 1) * 32, 0, SRCCOPY)
        Dummy = BitBlt(PicTopWork.hDC, PixelX, PixelY, 32, 32, PicObjects.hDC, (MapHolder(X, Y).TTile - 1) * 32, 0, SRCCOPY)
        Dummy = BitBlt(PicTopMask.hDC, PixelX, PixelY, 32, 32, PicObjMask.hDC, (MapHolder(X, Y).TTile - 1) * 32, 0, SRCCOPY)
        
        PixelX = PixelX + 32
    Next X
    PixelY = PixelY + 32
Next Y
frmMap.Caption = "Insomnia Map Editor - [..." + Right$(MapFile, 12) + "]"

VertBar.Value = 1
HorizBar.Value = 1
    Exit Sub

ErrHandler:
    Exit Sub

End Sub

Private Sub mnuOpen_Click()
    Call OpenMap
End Sub

Private Sub SaveMap()

If MapFile = Path + "new.map" Then
    Call SaveMapAs
Else

    Dim X, Y
    Dim RecordNum As Long
    For Y = 1 To 50
        For X = 1 To 50
        RecordNum = (X * (MaxCoord + 1)) + Y
        Put #1, RecordNum, MapHolder(X, Y)
        'Write records to file.
        Next X
    Next Y

End If

End Sub

Private Sub mnuSave_Click()
    Call SaveMap
End Sub

Private Sub SaveMapAs()
Dim MapFile As String

Dim ErrHandler
CommonDialog.CancelError = True
On Error GoTo ErrHandler
CommonDialog.Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNOverwritePrompt
CommonDialog.Filter = "Map Files (*.MAP)|*.MAP|MapData Files (*.DAT)|*.DAT|All Files (*.*)|*.*"
CommonDialog.FilterIndex = 1 'Set to .MAP as default
CommonDialog.DialogTitle = "Save As"
CommonDialog.ShowSave 'Action = 2 'Open a "Save As" box
MapFile = CommonDialog.filename 'MapFile equals the filename chosen

Close #1 'Close the previously edited file
Open MapFile For Random As #1 Len = Len(MapHolder(50, 50))
frmMap.Caption = "Insomnia Map Editor - [..." + Right$(MapFile, 12) + "]"

Dim X, Y, RecordNum As Long
For Y = 1 To 50
    For X = 1 To 50
        RecordNum = (X * (MaxCoord + 1)) + Y
        Put #1, RecordNum, MapHolder(X, Y)
        'Write records to file.
    Next X
Next Y
    Exit Sub
    
ErrHandler:
    Exit Sub

End Sub

Private Sub mnuSaveAs_Click()
    Call SaveMapAs
End Sub

Private Sub mnuUsing_Click()
Dim Message
Message = "Hold you cursor over any item which you want to gain" + HR
Message = Message + "information about to display a help dialog on that" + HR
Message = Message + "perticular tool."
MsgBox Message, vbInformation, "DreamEditor Help"
End Sub

Private Sub optBottomDelete_Click()
TileBot.Tile = 1
End Sub

Private Sub optObjDelete_Click()
TileTop.Tile = 80
End Sub

Private Sub picMain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Dummy
Dim TileY, TileX 'Position # to the cursor (ie. 128 = 4)

'[Change the cursor location to a *32 location]
TileX = Int(X / 32)
TileY = Int(Y / 32)
CurX = TileX * 32
CurY = TileY * 32

lblCurPosition.Caption = Str(HorizBar.Value + TileX) + ", " + Str(VertBar.Value + TileY)
CursorVisible = True

End Sub

Private Sub picMain_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Dummy
Dim PixelX, PixelY 'Position on the 32*32 grid
Dim TileY, TileX 'Position # to the mouseclick (ie. 128 = 4)
Dim RecordNum As Long 'The current record

'[Change the mouse-click location to a *32 location]
TileX = Int(X / 32)
TileY = Int(Y / 32)
PixelX = TileX * 32
PixelY = TileY * 32

If Button = 1 Then 'If Left button is clicked

    If optBottom.Value = True Then
        RecordNum = (TileX + HorizBar.Value * (MaxCoord + 1)) + TileY + VertBar.Value
        Dummy = BitBlt(PicBottomWork.hDC, PixelX + (HorizBar.Value - 1) * 32, PixelY + (VertBar.Value - 1) * 32, 32, 32, PicTile.hDC, (TileBot.Tile - 1) * 32, 0, SRCCOPY)
        MapHolder(TileX + HorizBar.Value, TileY + VertBar.Value).BIsWall = TileBot.IsWall
        MapHolder(TileX + HorizBar.Value, TileY + VertBar.Value).BTile = TileBot.Tile 'Record the tile's image for that position
    ElseIf optTop.Value = True Then
        RecordNum = (TileX + HorizBar.Value * (MaxCoord + 1)) + TileY + VertBar.Value
        Dummy = BitBlt(PicTopWork.hDC, PixelX + (HorizBar.Value - 1) * 32, PixelY + (VertBar.Value - 1) * 32, 32, 32, PicObjects.hDC, (TileTop.Tile - 1) * 32, 0, SRCCOPY)
        Dummy = BitBlt(PicTopMask.hDC, PixelX + (HorizBar.Value - 1) * 32, PixelY + (VertBar.Value - 1) * 32, 32, 32, PicObjMask.hDC, (TileTop.Tile - 1) * 32, 0, SRCCOPY)
        MapHolder(TileX + HorizBar.Value, TileY + VertBar.Value).TIsWall = TileTop.IsWall
        MapHolder(TileX + HorizBar.Value, TileY + VertBar.Value).TTile = TileTop.Tile 'Record the tile's image for that position
        If chkUnderLay.Value = 1 Then
            MapHolder(TileX + HorizBar.Value, TileY + VertBar.Value).TIsUnderLay = True
        Else: MapHolder(TileX + HorizBar.Value, TileY + VertBar.Value).TIsUnderLay = False
        End If
    ElseIf optMonster.Value = True Then
        'MONSTER BLIT CODES
    ElseIf optBottomDelete.Value = True Then
        RecordNum = (TileX + HorizBar.Value * (MaxCoord + 1)) + TileY + VertBar.Value
        Dummy = BitBlt(PicBottomWork.hDC, PixelX + (HorizBar.Value - 1) * 32, PixelY + (VertBar.Value - 1) * 32, 32, 32, PicTile.hDC, (TileBot.Tile - 1) * 32, 0, SRCCOPY)
        MapHolder(TileX + HorizBar.Value, TileY + VertBar.Value).BIsWall = False
        MapHolder(TileX + HorizBar.Value, TileY + VertBar.Value).BTile = TileBot.Tile
    ElseIf optObjDelete.Value = True Then
        RecordNum = (TileX + HorizBar.Value * (MaxCoord + 1)) + TileY + VertBar.Value
        Dummy = BitBlt(PicTopWork.hDC, PixelX + (HorizBar.Value - 1) * 32, PixelY + (VertBar.Value - 1) * 32, 32, 32, PicMaskBW.hDC, (TileTop.Tile - 1) * 32, 0, SRCCOPY)
        Dummy = BitBlt(PicTopMask.hDC, PixelX + (HorizBar.Value - 1) * 32, PixelY + (VertBar.Value - 1) * 32, 32, 32, PicMaskBW.hDC, (TileTop.Tile - 1) * 32, 0, SRCCOPY)
        MapHolder(TileX + HorizBar.Value, TileY + VertBar.Value).TIsWall = False
        MapHolder(TileX + HorizBar.Value, TileY + VertBar.Value).TTile = TileTop.Tile
    ElseIf optMonsterDelete.Value = True Then
        RecordNum = (TileX + HorizBar.Value * (MaxCoord + 1)) + TileY + VertBar.Value
        Dummy = BitBlt(PicTopWork.hDC, PixelX + (HorizBar.Value - 1) * 32, PixelY + (VertBar.Value - 1) * 32, 32, 32, PicMaskBW.hDC, 0, 0, SRCCOPY)
        Dummy = BitBlt(PicTopMask.hDC, PixelX + (HorizBar.Value - 1) * 32, PixelY + (VertBar.Value - 1) * 32, 32, 32, PicMaskBW.hDC, 32, 0, SRCCOPY)
    End If

ElseIf Button = 2 Then 'If Right button is clicked
    MsgBox "Please use the left mouse key to drop a tile here.", vbOKOnly, "Use Left Button"
End If

End Sub
Private Sub Mask_Make(dest As tArea, src As tArea, StartX As Integer, StartY As Integer, Width As Integer, Height As Integer)
'dest = destination object, src = source object
'!!WARNING!! make sure all forms and objects have pixel as their scalemode

Dim X As Integer    'x pixel pos
Dim Y As Integer    'y pixel pos
Dim color As Long   'current color of pixel
Dim Dummy As Long   'dummy return code needed for blit
Dim TransparentColor As Long  'color to be transparent
Dim FG As Long      'foreground mask color
Dim BG As Long      'background mask color

'foreground and backgroung settings
FG = WHITE               'foreground is white
BG = BLACK               'background is black
TransparentColor = BG    'What color is xparent

'pixel by pixel, make an invert/negative (mask)
For Y = 0 To Height    'do until Width
    For X = 0 To Width    'do until Height
        color = GetPixel(src.hDC, X, Y)  'check pixel color
        If color = TransparentColor Then  'Black: make it white
            Dummy = SetPixel(dest.hDC, X, Y, FG)
        Else
            Dummy = SetPixel(dest.hDC, X, Y, BG) 'Color: make it black
        End If
    Next X
Next Y

End Sub

Private Sub Board_Refresh()
Dim Dummy As Long 'dummy variable needed for blit

'[Map TileBot Layer Blit]
Dummy = BitBlt(PicTemp.hDC, 0, 0, 320, 320, PicBottomWork.hDC, (HorizBar.Value - 1) * 32, (VertBar.Value - 1) * 32, SRCCOPY)

'[Map Top Layer Blit]
Dummy = BitBlt(PicTemp.hDC, 0, 0, 320, 320, PicTopMask.hDC, (HorizBar.Value - 1) * 32, (VertBar.Value - 1) * 32, SRCAND)
Dummy = BitBlt(PicTemp.hDC, 0, 0, 320, 320, PicTopWork.hDC, (HorizBar.Value - 1) * 32, (VertBar.Value - 1) * 32, SRCINVERT)

'[Cursor Blit]
If CursorVisible = True Then
    Dummy = BitBlt(PicTemp.hDC, CurX, CurY, 32, 32, PicCurMask.hDC, 0, 0, SRCAND)
    Dummy = BitBlt(PicTemp.hDC, CurX, CurY, 32, 32, PicCurBox.hDC, 0, 0, SRCINVERT)
End If
  
'[EVERYTHING onto Main View Blit]
Dummy = BitBlt(picMain.hDC, 0, 0, 320, 320, PicTemp.hDC, 0, 0, SRCCOPY)

End Sub

Private Sub Tool_Refresh()
Dim Dummy

Dummy = BitBlt(picBTile.hDC, 0, 0, 32, 32, PicTile.hDC, (TileBot.Tile - 1) * 32, 0, SRCCOPY)
Dummy = BitBlt(picTTile.hDC, 0, 0, 32, 32, PicObjects.hDC, (TileTop.Tile - 1) * 32, 0, SRCCOPY)

If optBottom.Value = True Then 'If the bottom layer check is checked.
    lblCurrentTile.Visible = False
    Dummy = BitBlt(picCurrentTile.hDC, 0, 0, 32, 32, PicTile.hDC, (TileBot.Tile - 1) * 32, 0, SRCCOPY)
ElseIf optTop.Value = True Then 'If the top Layer is checked
    lblCurrentTile.Visible = False
    Dummy = BitBlt(picCurrentTile.hDC, 0, 0, 32, 32, PicObjects.hDC, (TileTop.Tile - 1) * 32, 0, SRCCOPY)
ElseIf optMonster.Value = True Then 'If monsters are checked
    lblCurrentTile.Visible = False
    
ElseIf optBottomDelete.Value = True Then
    lblCurrentTile.Visible = True
ElseIf optObjDelete.Value = True Then
    lblCurrentTile.Visible = True
ElseIf optMonsterDelete.Value = True Then
    lblCurrentTile.Visible = True
End If

End Sub

Private Sub picBTile_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Dummy

If Button = 1 Then
    If TileBot.Tile < TileBot.MaxNumTiles Then
            TileBot.Tile = TileBot.Tile + 1
    Else: TileBot.Tile = 1
    End If
ElseIf Button = 2 Then
    If TileBot.Tile > 1 Then
        TileBot.Tile = TileBot.Tile - 1
    Else: TileBot.Tile = TileBot.MaxNumTiles
    End If
End If

Dummy = BitBlt(picBTile.hDC, 0, 0, 32, 32, PicTile.hDC, (TileBot.Tile - 1) * 32, 0, SRCCOPY)
optBottom.Value = True

End Sub

Private Sub picTTile_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Dummy

If Button = 1 Then
    If TileTop.Tile < TileTop.MaxNumTiles Then
            TileTop.Tile = TileTop.Tile + 1
    Else: TileTop.Tile = 1
    End If
ElseIf Button = 2 Then
    If TileTop.Tile > 1 Then
        TileTop.Tile = TileTop.Tile - 1
    Else: TileTop.Tile = TileTop.MaxNumTiles
    End If
End If

Dummy = BitBlt(picTTile.hDC, 0, 0, 32, 32, PicObjects.hDC, (TileTop.Tile - 1) * 32, 0, SRCCOPY)
optTop.Value = True

End Sub

Private Sub tmrPaint_Timer()

Call Board_Refresh
Call Tool_Refresh

End Sub

Private Sub ToolBar_ButtonClick(ByVal Button As ComctlLib.Button)

Select Case Button.Index
    Case 2: Call MakeNewMap
    Case 3: Call OpenMap
    Case 5: Call SaveMap
    Case 6: Call SaveMapAs
    Case Else: Exit Sub
End Select

End Sub

Private Sub Temporary_Space()
Dim ErrHandler
' Set CancelError is True
    CommonDialog.CancelError = True
    On Error GoTo ErrHandler
    ' Set flags
    CommonDialog.Flags = cdlOFNHideReadOnly
    ' Set filters
    CommonDialog.Filter = "All Files (*.*)|*.*|Text Files" & _
    "(*.txt)|*.txt|Batch Files (*.bat)|*.bat"
    ' Specify default filter
    CommonDialog.FilterIndex = 2
    ' Display the Open dialog box
    CommonDialog.ShowOpen
    ' Display name of selected file

    MsgBox CommonDialog.filename
    Exit Sub
    
ErrHandler:
    'User pressed the Cancel button
    Exit Sub

End Sub

