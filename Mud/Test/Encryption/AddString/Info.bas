Attribute VB_Name = "GameInfo"
Option Explicit

'[Locations of files]
Global Path As String      'Location of the .EXE on the HD

'[Dialog Box Constants]
Const cdlOFNFileMustExist = &H1000
Const cdlOFNHideReadOnly = &H4
Const cdlOFNOverwritePrompt = &H2

'DECLARES: DRAWING*************************************************

'[Area Data type - Solid image]
Type tArea
    hDC As Long
    Left As Integer
    Top As Integer
    Width As Integer
    Height As Integer
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

'FUNCTIONS: Manipulating Sprites***********************************

Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Integer
Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Declare Function GetFocus Lib "user32" () As Long

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
Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Declare Function CreatePalette Lib "gdi32" (lpLogPalette As LOGPALETTE) As Long
Declare Function RealizePalette Lib "gdi32" (ByVal hDC As Long) As Long
Declare Function SelectPalette Lib "gdi32" (ByVal hDC As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long
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

Public Function FuncTimeOut(TOInterval As Single)

Dim TOStart As Single
TOStart = Timer
Do: DoEvents: Loop Until Timer - TOStart >= TOInterval

End Function
