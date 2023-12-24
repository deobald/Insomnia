Attribute VB_Name = "GameInfo"
Option Explicit
                
                '***********************
                '****** Game Codes *****
                '***********************

'DECLARES: STRING**************************************************

'[KeyCode constants]
Global Const vbKeyReturn = 13
Global Const KEY_LEFT = 37
Global Const KEY_RIGHT = 39
Global Const KEY_UP = 38
Global Const KEY_DOWN = 40
Global Const KEY_SPACE = 32
Public SavedWnd As Long

Global Namo As String
Global Passwordo As String

'[Locations of files]
Global Path As String      'Location of the .EXE on the HD
Global PathArchive As String   'Locations of the .DLL files
Global PathMap As String   'Location of .MAP files
Global PathMusic As String 'Location of .MID files
Global PathSound As String 'Locations of .WAV files

'DECLARES: MAP*****************************************************

'[Properties for the map variable]
Type Record
    BTile As Integer
    BIsWall As Boolean
    TTile As Integer
    TIsWall As Boolean
    TIsUnderLay As Boolean
End Type

Global Const MaxCoord = 50 'Maximum tile position

'DECLARES: DRAWING*************************************************

'[Area Data type - Solid image]
Type tArea
    hdc As Long
    Left As Integer
    Top As Integer
    Width As Integer
    Height As Integer
End Type

'[Sprite data type look at Char_init for meanings]
Type tSprite
    Width As Integer
    Height As Integer
    Frame As Integer
    X As Single
    Y As Single
    SprX As Integer
    SpriteNum As Integer
    Direction As Integer
    src As tArea
    bkg As tArea
    mask As tArea
End Type

'[Client / Server Declarations]

Public GameScreen As String 'The view currently used
'Connect, CharSelect, or Main

Global Const MAXPLAYERS = 10

Type PlayerInfo
    Name As String
    IsConnected As Boolean
    SpriteNum As Integer
    
    Direction As Integer
    Busy As Boolean
    SprX As Integer
    Frame As Byte
    
    MoveScript As String
    X As Single
    Y As Single
End Type

Public Player(1 To MAXPLAYERS) As PlayerInfo

Type CharData
    Name As String * 10
    Password As String * 10
    X As Integer
    Y As Integer
    Map As String * 20
End Type

'[Color Constants]
Global Const WHITE = &HFFFFFF
Global Const BLACK = &H0&

'[Windows GDI Bitmap API constants and functions]
Global Const SRCCOPY = &HCC0020
Global Const SRCAND = &H8800C6
Public Const SRCERASE = &H440328
Global Const SRCINVERT = &H660046
Global Const SRCPAINT = &HEE0086

'FUNCTIONS: Strings********************************************

Declare Function PutFocus Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long

'FUNCTIONS: Manipulating Sprites***********************************

Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Integer
Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Declare Function GetFocus Lib "user32" () As Long

'FUNCTIONS: Music and Sounds***************************************

'[Music]
'Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As Any, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Const mciModeNotOpen = 524

'[Sounds]
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
'The parameter lpszSoundName is the path & filename of the WAV file
'you want to play. uFlags is used to specify different ways to play
'the sound. See the constants below.

'Constants used by the uFlags parameter
Public Const SND_SYNC = &H0       'The sound is played synchronously. The
                                  'function does not return until the sound ends.

Public Const SND_ASYNC = &H1      'The sound is played asynchronously. The
                                  'function returns immediately after beginning
                                  'the sound. To terminate an asynchronously
                                  'played sound, call SndPlaySound with
                                  'lpszSoundName set to ""

Public Const SND_NODEFAULT = &H2  'If the specified sound cannot be played, the
                                  'function does not play the default sound.

Public Const SND_MEMORY = &H4     'The parameter specified by lpszSoundName
                                  'points to an in-memory image of a wave-
                                  'form sound. Use this if your sounds are part
                                  'of a resource file.

Public Const SND_LOOP = &H8       'The sound continues to play repeatedly until
                                  'sndPlaySound is called again with
                                  'lpsSoundName set to "". You must also
                                  'specify the SND_ASYNC flag to loop
                                  'sounds.

Public Const SND_NOSTOP = &H10    'If a sound is currently playing, the function
                                  'immediately returns FALSE without playing
                                  'the specified sound.
                                  

            '**********************************
            '****** hDC and Palette Codes *****
            '**********************************

'[User Types]
Type PALETTEENTRY
        peRed As Byte
        peGreen As Byte
        peBlue As Byte
        peFlags As Byte
End Type

Type LOGPALETTE
        palVersion As Integer
        palNumEntries As Integer
        palPalEntry(1) As PALETTEENTRY
End Type

'[Windows GDI API constants and Functions for Temp HDC]
'***************************************************************
Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Declare Function CreatePalette Lib "gdi32" (lpLogPalette As LOGPALETTE) As Long
Declare Function RealizePalette Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function SelectPalette Lib "gdi32" (ByVal hdc As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long
'***************************************************************

'[Arrays that will hold the Hdc's]
Dim MemHdc() As Long
Dim BitmapHdc() As Long
Dim TrashBmpHdc() As Long
Dim NumOfDcs As Integer

Function CreateMemHdc(ScreenHdc As Long, Width As Integer, Height As Integer) As Long
'This function will create a temporary Hdc to blit in and out of
'ScreenHdc = the display DC that we will be compatible
' Width = width of needed bitmap
' Height = height of needed bitmap

ReDim Preserve MemHdc(NumOfDcs)
ReDim Preserve BitmapHdc(NumOfDcs)
ReDim Preserve TrashBmpHdc(NumOfDcs)

MemHdc(NumOfDcs) = CreateCompatibleDC(ScreenHdc)
    If MemHdc(NumOfDcs) Then
        BitmapHdc(NumOfDcs) = CreateCompatibleBitmap(ScreenHdc, Width, Height)
        If BitmapHdc(NumOfDcs) Then
            TrashBmpHdc(NumOfDcs) = SelectObject(MemHdc(NumOfDcs), BitmapHdc(NumOfDcs))
            CreateMemHdc = MemHdc(NumOfDcs)
        End If
    End If
NumOfDcs = NumOfDcs + 1
End Function
Sub DestroyHdcs()
'Free all DCs
Dim Dummy As Long
Dim I As Integer

For I = 0 To NumOfDcs - 1
BitmapHdc(I) = SelectObject(MemHdc(I), TrashBmpHdc(I))
Dummy = DeleteObject(BitmapHdc(I))
Dummy = DeleteDC(MemHdc(I))
Next I

End Sub
Sub LoadBmpToHdc(MHdc As Long, FileN As String)
Dim OrgBmp As Long

OrgBmp = SelectObject(MHdc, LoadPicture(PathArchive & FileN))

If OrgBmp Then
   Call DeleteObject(OrgBmp)
End If

End Sub

Function FuncTimeOut(TOInterval As Single)

Dim TOStart As Single
TOStart = Timer
Do: DoEvents: Loop Until Timer - TOStart >= TOInterval

End Function

Function CharTimeOut(TOInterval As Single)

Dim TOStart As Single
TOStart = Timer
Do: DoEvents: Loop Until Timer - TOStart >= TOInterval

End Function

Function ScrollText(TextBox As Control, vLines As Integer)

Dim Success As Long

Dim R As Long
Const EM_LINESCROLL = &HB6

Dim Lines

' Get the window handle of the control that currently has the
'  focus, Command1 or Command2.
SavedWnd = Screen.ActiveControl.hWnd
Lines = vLines

' Set the focus to the passed control (text control).
R = PutFocus(TextBox.hWnd)       ' Scroll the lines.
Success = SendMessage(TextBox.hWnd, EM_LINESCROLL, 0, Lines)

' Restore the focus to the original control, Command1 or
'  Command2.
R = PutFocus(SavedWnd)

' Return the number of lines actually scrolled.
ScrollText = Success

End Function
