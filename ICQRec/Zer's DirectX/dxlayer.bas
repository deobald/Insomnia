Attribute VB_Name = "Module2"
Public Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, ByVal source As Long, ByVal Length As Long)
Public Declare Function lstrcpy Lib "kernel32" (ByVal lpszDestinationString1 As Any, ByVal lpszSourceString2 As Any) As Long
Public Declare Function waveOutSetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, ByVal dwVolume As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Public Const SND_SYNC = &H0
Public Const SND_ASYNC = &H1
Public Const SND_NODEFAULT = &H2
Public Const SND_LOOP = &H8
Public Const SND_NOSTOP = &H10
Public ddMouseCursor As DirectDrawSurface3
Public ddMouseX As Integer
Public ddMouseY As Integer
Public dsSound As DirectSound
Public dsSoundBuffer(8) As DirectSoundBuffer

Sub apiPlayMidi(midiname$)
ret = mciSendString("open " + midiname$ + " type sequencer alias hoho", 0&, 0, 0)
ret = mciSendString("play hoho", 0&, 0, 0)
End Sub
Sub apiPlayWave(SoundName$)
   wFlags% = SND_ASYNC Or SND_NODEFAULT
   X% = sndPlaySound(SoundName$, wFlags%)
End Sub


Sub apiStopMidi()
ret = mciSendString("stop hoho", 0&, 0, 0)
End Sub

Public Sub CreateDSBFromWaveFile(ds As DirectSound, ByVal strFile As String, dsb As DirectSoundBuffer)
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
        .dwFlags = DSBCAPS_CTRLDEFAULT 'Or DSBCAPS_STATIC Or DSBCAPS_LOCSOFTWARE
        .dwBufferBytes = lngSize
        .lpwfxFormat = VarPtr(pcmwave)
    End With
    ' Create the sound buffer
    ds.CreateSoundBuffer dsbd, dsb, Nothing
    ' Lock
    dsb.Lock 0&, lngSize, ptr1, lng1, ptr2, lng2, 0&
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

Sub dsLoadWave(strFile As String, channelnumber As Integer)
CreateDSBFromWaveFile dsSound, strFile, dsSoundBuffer(channelnumber)
End Sub


Sub dsPlay(channelnum As Integer)
dsSoundBuffer(channelnum).Play 0, 0, 0
End Sub
Sub dsSetPan(panval As Long, channelnum As Integer)
dsSoundBuffer(channelnum).SetPan panval
End Sub

Sub dsSetVol(vol As Long, channelnum As Integer)
Dim von As Long
von = 100 - vol
dsSoundBuffer(channelnum).SetVolume -Int(von * 100)
End Sub

Sub dsInit()
    DirectSoundCreate ByVal 0&, dsSound, Nothing
    dsSound.SetCooperativeLevel Form1.hWnd, DSSCL_NORMAL
End Sub

Sub dsStop(channelnum As Integer)
dsSoundBuffer(channelnum).Stop
End Sub
Sub dsUninit()
    Set dsSound = Nothing
    Set dsSoundBuffer(0) = Nothing
    Set dsSoundBuffer(1) = Nothing
    Set dsSoundBuffer(2) = Nothing
    Set dsSoundBuffer(3) = Nothing
    Set dsSoundBuffer(4) = Nothing
    Set dsSoundBuffer(5) = Nothing
    Set dsSoundBuffer(6) = Nothing
    Set dsSoundBuffer(7) = Nothing
    Set dsSoundBuffer(8) = Nothing
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
Sub dxBitBlt(X As Integer, Y As Integer, bltWidth As Integer, bltHeight As Integer, srcSurface As DirectDrawSurface3, srcX As Integer, srcY As Integer)
Dim dxBlitSrcRect As RECT
iX = X
iY = Y
dxBlitSrcRect.Top = srcY
dxBlitSrcRect.bottom = srcY + bltHeight
dxBlitSrcRect.Left = srcX
dxBlitSrcRect.Right = srcX + bltWidth
If iX + bltWidth < 0 Or iY + bltHeight < 0 Or iX > 640 Or iY > 480 Then Exit Sub

If iX < 0 Then dxBlitSrcRect.Left = dxBlitSrcRect.Left + Abs(iX): iX = 0
If iY < 0 Then dxBlitSrcRect.Top = dxBlitSrcRect.Top + Abs(iY): iY = 0
If iX + bltWidth > 640 Then dxBlitSrcRect.Right = dxBlitSrcRect.Right - ((iX + bltWidth) - 640)
If iY + bltHeight > 480 Then dxBlitSrcRect.bottom = dxBlitSrcRect.bottom - ((iY + bltHeight) - 480)

dixuBackBuffer.BltFast iX, iY, srcSurface, dxBlitSrcRect, DDBLTFAST_NOCOLORKEY Or DDBLTFAST_WAIT
End Sub
Sub dxTransBlt(X As Integer, Y As Integer, bltWidth As Integer, bltHeight As Integer, srcSurface As DirectDrawSurface3, srcX As Integer, srcY As Integer)
Dim dxBlitSrcRect As RECT
Dim dxBlitDestRect As RECT
Dim dxFx As DDBLTFX
iX = X
iY = Y
dxFx.dwSize = Len(dxFx)
dxFx.ddckSrcColorkey.dwColorSpaceHighValue = 0
dxFx.ddckSrcColorkey.dwColorSpaceLowValue = 0
dxBlitSrcRect.Top = srcY
dxBlitSrcRect.bottom = srcY + bltHeight
dxBlitSrcRect.Left = srcX
dxBlitSrcRect.Right = srcX + bltWidth
dxBlitDestRect.Top = iY
dxBlitDestRect.bottom = iY + bltHeight
dxBlitDestRect.Left = iX
dxBlitDestRect.Right = iX + bltWidth

If iX + bltWidth < 0 Or iY + bltHeight < 0 Or iX > 640 Or iY > 480 Then Exit Sub

If iX < 0 Then dxBlitSrcRect.Left = dxBlitSrcRect.Left + (-iX): dxBlitDestRect.Left = 0
If iY < 0 Then dxBlitSrcRect.Top = dxBlitSrcRect.Top + (-iY): dxBlitDestRect.Top = 0
If iX + bltWidth > 640 Then dxBlitSrcRect.Right = dxBlitSrcRect.Right - ((iX + bltWidth) - 640): dxBlitDestRect.Right = 640
If iY + bltHeight > 480 Then dxBlitSrcRect.bottom = dxBlitSrcRect.bottom - ((iY + bltHeight) - 480): dxBlitDestRect.bottom = 480
dixuBackBuffer.Blt dxBlitDestRect, srcSurface, dxBlitSrcRect, DDBLT_WAIT Or DDBLT_KEYSRCOVERRIDE, dxFx
End Sub
