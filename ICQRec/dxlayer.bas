Attribute VB_Name = "DirectDraw"
Public Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
Public Declare Function lstrcpy Lib "kernel32" (ByVal lpszDestinationString1 As Any, ByVal lpszSourceString2 As Any) As Long
Public Declare Function waveOutSetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, ByVal dwVolume As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Type POINTAPI
        X As Long
        Y As Long
End Type
Public vbPal(0 To 255) As PALETTEENTRY
Public vbPalette As DirectDrawPalette, dxDisable As Integer
Public Const SND_SYNC = &H0
Public Const SND_ASYNC = &H1
Public Const SND_NODEFAULT = &H2
Public Const SND_LOOP = &H8
Public Const SND_NOSTOP = &H10
Public ddMouseCursor As DirectDrawSurface3
Public ddMouseX As Integer
Public ddMouseY As Integer
Public dxFxDb As DDBLTFX
Public WobbleX As Single, WobbleY As Single
Type ImageData
    asdf As Integer
End Type

Sub apiPlayMidi(midiname$)
ret = mciSendString("open " + midiname$ + " type sequencer alias hoho", 0&, 0, 0)
ret = mciSendString("play hoho", 0&, 0, 0)
End Sub
Sub apiPlayWave(SoundName$)
   wFlags% = SND_ASYNC Or SND_NODEFAULT Or SND_NOSTOP
   X% = sndPlaySound(SoundName$, wFlags%)
End Sub


Sub apiStopMidi()
ret = mciSendString("stop hoho", 0&, 0, 0)
End Sub

' Creates a DirectSoundBuffer from a wave file
Public Sub CreateDSBFromWaveFile(DS As DirectSound, ByVal strFile As String, DSB As DirectSoundBuffer)
    Dim hWave As Long
    Dim pcmwave As WAVEFORMATEX
    Dim lngSize As Long
    Dim lngPosition As Long
    Dim ptr1 As Long, ptr2 As Long, lng1 As Long, lng2 As Long
    Dim aByte() As Byte
    'Dim lngBufferSize As Long
    'Dim strBuffer As String
    ' Byte array to load the whole file
    ReDim aByte(1 To FileLen(strFile))
    hWave = FreeFile
    Open strFile For Binary As hWave
    ' Load the whole file in the byte array
    Get hWave, , aByte
    Close hWave
    ' Search "fmt" tag
    lngPosition = 1
    While Chr$(aByte(lngPosition)) + Chr$(aByte(lngPosition + 1)) + Chr$(aByte(lngPosition + 2)) <> "fmt"
        lngPosition = lngPosition + 1
    Wend
    ' Copy wave header to structure
    CopyMemory VarPtr(pcmwave), VarPtr(aByte(lngPosition + 8)), Len(pcmwave)
    ' Search "data" tag
    While Chr$(aByte(lngPosition)) + Chr$(aByte(lngPosition + 1)) + Chr$(aByte(lngPosition + 2)) + Chr$(aByte(lngPosition + 3)) <> "data"
        lngPosition = lngPosition + 1
    Wend
    ' Get the data size
    CopyMemory VarPtr(lngSize), VarPtr(aByte(lngPosition + 4)), Len(lngSize)
    ' Fill buffer description
    Dim dsbd As DSBUFFERDESC
    With dsbd
        .dwSize = Len(dsbd)
        .dwFlags = DSBCAPS_CTRLDEFAULT Or DSBCAPS_STATIC Or DSBCAPS_LOCSOFTWARE
        .dwBufferBytes = lngSize
        .lpwfxFormat = VarPtr(pcmwave)
    End With
    ' Create the sound buffer
    DS.CreateSoundBuffer dsbd, DSB, Nothing
    ' Lock
    DSB.Lock 0&, lngSize, ptr1, lng1, ptr2, lng2, 0&
    ' Copy data to buffer
    CopyMemory ptr1, VarPtr(aByte(lngPosition + 4 + 4)), lng1
    ' Copy second part if needed
    If lng2 <> 0 Then
        CopyMemory ptr2, VarPtr(aByte(lngPosition + 4 + 4 + lng1)), lng2
    End If
    ' Unlock
    ' Automation error if uncommented !
    'dsb.Unlock ptr1, lng1, ptr2, lng2
End Sub


Sub diFS2Blt(X As Integer, Y As Integer, bltWidth As Integer, bltHeight As Integer, srcSurface As DirectDrawSurface3, srcX As Integer, srcY As Integer)
Dim dxBlitSrcRect As RECT
'On Error Resume Next
dxBlitSrcRect.Top = srcY
dxBlitSrcRect.bottom = srcY + bltHeight
dxBlitSrcRect.Left = srcX
dxBlitSrcRect.Right = srcX + bltWidth

If X + bltWidth <= 2 Or Y + bltHeight <= 2 Or X >= 638 Or Y >= 478 Then Exit Sub
If X < 2 Then dxBlitSrcRect.Left = (dxBlitSrcRect.Left - X) + 2: X = 2 '+ 1: X = 0
If Y < 2 Then dxBlitSrcRect.Top = (dxBlitSrcRect.Top - Y) + 2: Y = 2 '+ 1: Y = 0
If X + bltWidth >= 638 Then dxBlitSrcRect.Right = dxBlitSrcRect.Right - ((X + bltWidth) - 638)
If Y + bltHeight >= 478 Then dxBlitSrcRect.bottom = dxBlitSrcRect.bottom - ((Y + bltHeight) - 478)
Call dixuBackBuffer.BltFast(X, Y, srcSurface, dxBlitSrcRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub

Sub diFSBlt(X As Integer, Y As Integer, bltWidth As Integer, bltHeight As Integer, srcSurface As DirectDrawSurface3, srcX As Integer, srcY As Integer)
Dim dxBlitSrcRect As RECT
'On Error Resume Next
dxBlitSrcRect.Top = srcY
dxBlitSrcRect.bottom = srcY + bltHeight
dxBlitSrcRect.Left = srcX
dxBlitSrcRect.Right = srcX + bltWidth

If X + bltWidth <= 2 Or Y + bltHeight <= 2 Or X >= 638 Or Y >= 478 Then Exit Sub
If X < 2 Then dxBlitSrcRect.Left = (dxBlitSrcRect.Left - X) + 2: X = 2 '+ 1: X = 0
If Y < 2 Then dxBlitSrcRect.Top = (dxBlitSrcRect.Top - Y) + 2: Y = 2 '+ 1: Y = 0
If X + bltWidth >= 638 Then dxBlitSrcRect.Right = dxBlitSrcRect.Right - ((X + bltWidth) - 638)
If Y + bltHeight >= 478 Then dxBlitSrcRect.bottom = dxBlitSrcRect.bottom - ((Y + bltHeight) - 478)
Call dixuBackBuffer.BltFast(X, Y, srcSurface, dxBlitSrcRect, DDBLTFAST_WAIT)
End Sub


Sub d48BitBlt(X As Integer, Y As Integer, bltWidth As Integer, bltHeight As Integer, srcSurface As DirectDrawSurface3, srcX As Integer, srcY As Integer)
Dim dxBlitSrcRect As RECT
'On Error Resume Next
dxBlitSrcRect.Top = srcY
dxBlitSrcRect.bottom = srcY + bltHeight
dxBlitSrcRect.Left = srcX
dxBlitSrcRect.Right = srcX + bltWidth

If X + bltWidth <= 4 Or Y + bltHeight <= 4 Or X >= 635 Or Y >= 475 Then Exit Sub
If X < 4 Then dxBlitSrcRect.Left = (dxBlitSrcRect.Left - X) + 4: X = 4 '+ 1: X = 0
If Y < 4 Then dxBlitSrcRect.Top = (dxBlitSrcRect.Top - Y) + 4: Y = 4 '+ 1: Y = 0
If X + bltWidth >= 475 Then dxBlitSrcRect.Right = dxBlitSrcRect.Right - ((X + bltWidth) - 475)
If Y + bltHeight >= 475 Then dxBlitSrcRect.bottom = dxBlitSrcRect.bottom - ((Y + bltHeight) - 475)
Call dixuBackBuffer.BltFast(X, Y, srcSurface, dxBlitSrcRect, DDBLTFAST_WAIT)
End Sub


Sub d48TransBlt(X As Integer, Y As Integer, bltWidth As Integer, bltHeight As Integer, srcSurface As DirectDrawSurface3, srcX As Integer, srcY As Integer)
Dim dxBlitSrcRect As RECT
'On Error Resume Next
dxBlitSrcRect.Top = srcY
dxBlitSrcRect.bottom = srcY + bltHeight
dxBlitSrcRect.Left = srcX
dxBlitSrcRect.Right = srcX + bltWidth

If X + bltWidth <= 4 Or Y + bltHeight <= 4 Or X >= 635 Or Y >= 475 Then Exit Sub
If X < 4 Then dxBlitSrcRect.Left = (dxBlitSrcRect.Left - X) + 4: X = 4 '+ 1: X = 0
If Y < 4 Then dxBlitSrcRect.Top = (dxBlitSrcRect.Top - Y) + 4: Y = 4 '+ 1: Y = 0
If X + bltWidth >= 475 Then dxBlitSrcRect.Right = dxBlitSrcRect.Right - ((X + bltWidth) - 475)
If Y + bltHeight >= 475 Then dxBlitSrcRect.bottom = dxBlitSrcRect.bottom - ((Y + bltHeight) - 475)
Call dixuBackBuffer.BltFast(X, Y, srcSurface, dxBlitSrcRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub



Sub diNoClipBlt(X As Integer, Y As Integer, bltWidth As Integer, bltHeight As Integer, srcSurface As DirectDrawSurface3, srcX As Integer, srcY As Integer)
Dim dxBlitSrcRect As RECT
'On Error Resume Next
dxBlitSrcRect.Top = srcY
dxBlitSrcRect.bottom = srcY + bltHeight
dxBlitSrcRect.Left = srcX
dxBlitSrcRect.Right = srcX + bltWidth

Call dixuBackBuffer.BltFast(X, Y, srcSurface, dxBlitSrcRect, DDBLTFAST_WAIT)
End Sub

Sub diTransBlt(X As Integer, Y As Integer, bltWidth As Integer, bltHeight As Integer, srcSurface As DirectDrawSurface3, srcX As Integer, srcY As Integer)
'On Error Resume Next
Dim dxBlitSrcRect As RECT
With dxBlitSrcRect
    .Top = srcY
    .bottom = srcY + bltHeight
    .Left = srcX
    .Right = srcX + bltWidth
End With

If X + bltWidth <= 15 Or Y + bltHeight <= 15 Or X >= 400 Or Y >= 400 Then Exit Sub

If X < 15 Then dxBlitSrcRect.Left = (dxBlitSrcRect.Left - X) + 15: X = 15 '+ 1: X = 0
If Y < 15 Then dxBlitSrcRect.Top = (dxBlitSrcRect.Top - Y) + 15: Y = 15 '+ 1: Y = 0
If X + bltWidth >= 400 Then dxBlitSrcRect.Right = dxBlitSrcRect.Right - ((X + bltWidth) - 400)
If Y + bltHeight >= 400 Then dxBlitSrcRect.bottom = dxBlitSrcRect.bottom - ((Y + bltHeight) - 400)

'If dxBlitSrcRect.Left < 0 Then dxBlitSrcRect.Left = 0
'If dxBlitSrcRect.Top < 0 Then dxBlitSrcRect.Top = 0

Call dixuBackBuffer.BltFast(X, Y, srcSurface, dxBlitSrcRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub

Sub dfTransBlt(X As Integer, Y As Integer, bltWidth As Integer, bltHeight As Integer, srcSurface As DirectDrawSurface3, srcX As Integer, srcY As Integer)
'On Error Resume Next
Dim dxBlitSrcRect As RECT
With dxBlitSrcRect
    .Top = srcY
    .bottom = srcY + bltHeight
    .Left = srcX
    .Right = srcX + bltWidth
End With

If X + bltWidth <= 15 Or Y + bltHeight <= 15 Or X >= 465 Or Y >= 465 Then Exit Sub

If X < 15 Then dxBlitSrcRect.Left = (dxBlitSrcRect.Left - X) + 15: X = 15 '+ 1: X = 0
If Y < 15 Then dxBlitSrcRect.Top = (dxBlitSrcRect.Top - Y) + 15: Y = 15 '+ 1: Y = 0
If X + bltWidth >= 465 Then dxBlitSrcRect.Right = dxBlitSrcRect.Right - ((X + bltWidth) - 465)
If Y + bltHeight >= 465 Then dxBlitSrcRect.bottom = dxBlitSrcRect.bottom - ((Y + bltHeight) - 465)

Call dixuBackBuffer.BltFast(X, Y, srcSurface, dxBlitSrcRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub


Sub dlTransBlt(X As Integer, Y As Integer, bltWidth As Integer, bltHeight As Integer, srcSurface As DirectDrawSurface3, srcX As Integer, srcY As Integer)
On Error Resume Next
Dim R As RECT
If X + bltWidth > 399 Then
    If X >= 399 Then
        Exit Sub
    Else
        bltWidth = 399 - X + 1
    End If
End If
If X < 15 Then
    If X + bltWidth <= 15 Then
        Exit Sub
    Else
        srcX = 15 - X
        bltWidth = X + bltWidth - 15
        X = 15
    End If
End If
    If Y + bltHeight > 399 Then
        If Y >= 399 Then
            Exit Sub
        Else
            bltHeight = 399 - Y + 1
        End If
    End If
    If Y < 15 Then
        If Y + bltHeight <= 15 Then
            Exit Sub
        Else
            srcY = 15 - Y
            bltHeight = Y + bltHeight - 15
            Y = 15
        End If
    End If
    With R
        .Left = srcX
        .Top = srcY
        .Right = srcX + bltWidth
        .bottom = srcY + bltHeight
    End With
    dixuBackBuffer.BltFast X, Y, srcSurface, R, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
End Sub

Sub dmTransBlt(X As Integer, Y As Integer, bltWidth As Integer, bltHeight As Integer, srcSurface As DirectDrawSurface3, srcX As Integer, srcY As Integer)
On Error Resume Next
Dim R As RECT
If X + bltWidth > 639 Then
    If X >= 639 Then
        Exit Sub
    Else
        bltWidth = 639 - X + 1
    End If
End If
If X < 1 Then
    If X + bltWidth <= 1 Then
        Exit Sub
    Else
        srcX = 1 - X
        bltWidth = X + bltWidth - 1
        X = 1
    End If
End If
    If Y + bltHeight > 479 Then
        If Y >= 479 Then
            Exit Sub
        Else
            bltHeight = 479 - Y + 1
        End If
    End If
    If Y < 1 Then
        If Y + bltHeight <= 1 Then
            Exit Sub
        Else
            srcY = 1 - Y
            bltHeight = Y + bltHeight - 1
            Y = 1
        End If
    End If
    With R
        .Left = srcX
        .Top = srcY
        .Right = srcX + bltWidth
        .bottom = srcY + bltHeight
    End With
    dixuBackBuffer.BltFast X, Y, srcSurface, R, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
End Sub


Sub dxPlayWave(wavname As String)
    wFlags% = SND_ASYNC Or SND_NODEFAULT
    X% = sndPlaySound(wavname, wFlags%)
End Sub
Sub dxSetWave(volume As Integer)

End Sub
Sub dxCls()
        Dim RD As RECT
        Dim dxFx As DDBLTFX
        With dxFx
            .dwSize = Len(dxFx)
            .dwFillColor = RGB(0, 0, 0)
        End With
        RD.Top = 0
        RD.Left = 0
        RD.bottom = 480
        RD.Right = 640
        dixuBackBuffer.Blt RD, Nothing, RD, DDBLT_COLORFILL Or DDBLT_WAIT, dxFx
End Sub

Sub dxFlip()
dixuPrimarySurface.Flip Nothing, DDFLIP_WAIT
End Sub
Sub diBitBlt(X As Integer, Y As Integer, bltWidth As Integer, bltHeight As Integer, srcSurface As DirectDrawSurface3, srcX As Integer, srcY As Integer)
'On Error Resume Next
Dim dxBlitSrcRect As RECT
With dxBlitSrcRect
    .Top = srcY
    .bottom = srcY + bltHeight
    .Left = srcX
    .Right = srcX + bltWidth
End With

If X + bltWidth <= 15 Or Y + bltHeight <= 15 Or X >= 400 Or Y >= 400 Then Exit Sub

If X < 15 Then dxBlitSrcRect.Left = (dxBlitSrcRect.Left - X) + 15: X = 15 '+ 1: X = 0
If Y < 15 Then dxBlitSrcRect.Top = (dxBlitSrcRect.Top - Y) + 15: Y = 15 '+ 1: Y = 0
If X + bltWidth >= 400 Then dxBlitSrcRect.Right = dxBlitSrcRect.Right - ((X + bltWidth) - 400)
If Y + bltHeight >= 400 Then dxBlitSrcRect.bottom = dxBlitSrcRect.bottom - ((Y + bltHeight) - 400)

Call dixuBackBuffer.BltFast(X, Y, srcSurface, dxBlitSrcRect, DDBLTFAST_WAIT)
End Sub
Sub dfBitBlt(X As Integer, Y As Integer, bltWidth As Integer, bltHeight As Integer, srcSurface As DirectDrawSurface3, srcX As Integer, srcY As Integer)
'On Error Resume Next
Dim dxBlitSrcRect As RECT
With dxBlitSrcRect
    .Top = srcY
    .bottom = srcY + bltHeight
    .Left = srcX
    .Right = srcX + bltWidth
End With

If X + bltWidth <= 15 Or Y + bltHeight <= 15 Or X >= 465 Or Y >= 465 Then Exit Sub

If X < 15 Then dxBlitSrcRect.Left = (dxBlitSrcRect.Left - X) + 15: X = 15 '+ 1: X = 0
If Y < 15 Then dxBlitSrcRect.Top = (dxBlitSrcRect.Top - Y) + 15: Y = 15 '+ 1: Y = 0
If X + bltWidth >= 465 Then dxBlitSrcRect.Right = dxBlitSrcRect.Right - ((X + bltWidth) - 465)
If Y + bltHeight >= 465 Then dxBlitSrcRect.bottom = dxBlitSrcRect.bottom - ((Y + bltHeight) - 465)

Call dixuBackBuffer.BltFast(X, Y, srcSurface, dxBlitSrcRect, DDBLTFAST_WAIT)
End Sub

Sub dxTransBlt(X As Integer, Y As Integer, bltWidth As Integer, bltHeight As Integer, srcSurface As DirectDrawSurface3, srcX As Integer, srcY As Integer)
Dim dxBlitSrcRect As RECT
Dim dxBlitDestRect As RECT
Dim iX As Integer, iY As Integer
On Error Resume Next
iX = X
iY = Y
dxBlitSrcRect.Top = srcY
dxBlitSrcRect.bottom = srcY + bltHeight
dxBlitSrcRect.Left = srcX
dxBlitSrcRect.Right = srcX + bltWidth
dxBlitDestRect.Top = iY
dxBlitDestRect.bottom = iY + bltHeight
dxBlitDestRect.Left = iX
dxBlitDestRect.Right = iX + bltWidth

If iX + bltWidth <= 1 Or iY + bltHeight <= 1 Or iX > 639 Or iY > 479 Then Exit Sub

If iX < 0 Then dxBlitSrcRect.Left = (dxBlitSrcRect.Left - iX) + 1: dxBlitDestRect.Left = 0
If iY < 0 Then dxBlitSrcRect.Top = (dxBlitSrcRect.Top - iY) + 1: dxBlitDestRect.Top = 0
If iX + bltWidth >= 640 Then dxBlitSrcRect.Right = dxBlitSrcRect.Right - ((iX + bltWidth) - 640): dxBlitDestRect.Right = 640
If iY + bltHeight >= 480 Then dxBlitSrcRect.bottom = dxBlitSrcRect.bottom - ((iY + bltHeight) - 480): dxBlitDestRect.bottom = 480

dixuBackBuffer.Blt dxBlitDestRect, srcSurface, dxBlitSrcRect, DDBLT_WAIT Or DDBLT_KEYSRCOVERRIDE, dxFxDb
End Sub

Sub dxTransInit()
dxFxDb.dwSize = Len(dxFxDb)
dxFxDb.ddckSrcColorkey.dwColorSpaceHighValue = 0
dxFxDb.ddckSrcColorkey.dwColorSpaceLowValue = 0
End Sub


Sub prepSrcColorKey(srf As DirectDrawSurface3)
Dim aColorkey As DDCOLORKEY
aColorkey.dwColorSpaceHighValue = 0
aColorkey.dwColorSpaceLowValue = 0
srf.SetColorKey DDCKEY_SRCBLT, aColorkey
End Sub
Sub prepSrcColorKeyA(srf As DirectDrawSurface3)
Dim aColorkey As DDCOLORKEY
aColorkey.dwColorSpaceHighValue = 255
aColorkey.dwColorSpaceLowValue = 255
srf.SetColorKey DDCKEY_SRCBLT, aColorkey
End Sub

Public Sub SetPal(palname$)
ZD% = 9
Open App.Path + "\" + palname$ For Input As #ZD%
Line Input #ZD%, Useless$
Line Input #ZD%, Useless$
Line Input #ZD%, Useless$
For imm% = 0 To 255
Line Input #ZD%, US$

CPos% = InStr(US$, " ")
RedM$ = Left$(US$, CPos% - 1)
US$ = Right$(US$, (Len(US$) - Len(RedM$)) - 1)
Red% = Val(RedM$)

CPos% = InStr(US$, " ")
GreenM$ = Left$(US$, CPos% - 1)
US$ = Right$(US$, (Len(US$) - Len(GreenM$)) - 1)
Green% = Val(GreenM$)

Blue% = Val(US$)

vbPal(imm%).peRed = Red%
vbPal(imm%).peGreen = Green%
vbPal(imm%).peBlue = Blue%
vbPal(imm%).peFlags = 0
Next imm%
Close #ZD%
dixuDDraw.CreatePalette DDPCAPS_8BIT, vbPal(0), vbPalette, Nothing
dixuPrimarySurface.SetPalette vbPalette
Set vbPalette = Nothing
End Sub
Public Sub SetPalS(palname$, sf As DirectDrawSurface3)
ZD% = 9
Open App.Path + "\" + palname$ For Input As #ZD%
Line Input #ZD%, Useless$
Line Input #ZD%, Useless$
Line Input #ZD%, Useless$
For imm% = 0 To 255
Line Input #ZD%, US$

CPos% = InStr(US$, " ")
RedM$ = Left$(US$, CPos% - 1)
US$ = Right$(US$, (Len(US$) - Len(RedM$)) - 1)
Red% = Val(RedM$)

CPos% = InStr(US$, " ")
GreenM$ = Left$(US$, CPos% - 1)
US$ = Right$(US$, (Len(US$) - Len(GreenM$)) - 1)
Green% = Val(GreenM$)

Blue% = Val(US$)

vbPal(imm%).peRed = Red%
vbPal(imm%).peGreen = Green%
vbPal(imm%).peBlue = Blue%
vbPal(imm%).peFlags = 0
Next imm%
Close #ZD%
dixuDDraw.CreatePalette DDPCAPS_8BIT, vbPal(0), vbPalette, Nothing
sf.SetPalette vbPalette
Set vbPalette = Nothing
End Sub

