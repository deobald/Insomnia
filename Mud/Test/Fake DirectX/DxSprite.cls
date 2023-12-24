VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DxSprite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' dixuSprite class v 0.2, Copyright Patrice Scribe, 1997
' http://www.chez.com/scribe
' Changes from v 0.1 :
' - Key property added for dixuSpriteRemove

Option Explicit

Public Key As String                    ' Identification key
Public x As Long                        ' x position
Public y As Long                        ' y position
Public VelocityX As Long                ' Velocity along the X axis (not yet pixels/s)
Public VelocityY As Long                ' Velocity along the Y axis (not yet pixels/s)
Public Surface As DirectDrawSurface3    ' Source for sprite bitmap
Public Width As Long                    ' Width of sprite in pixels
Public Height As Long                   ' Height of sprite in pixels

Private Sub Class_Initialize()
    ' Register the new sprite in a sprites collection
    Key = CStr(Timer)
    dixuSprites.Add Me, Key
End Sub

Private Sub Class_Terminate()
    Set Surface = Nothing
End Sub

Public Sub Move()
    ' Move according to elapsed time
    x = x + VelocityX '/ (1 / (dixuTime - dixuLastTime))
    y = y + VelocityY '/ (1 / (dixuTime - dixuLastTime))
    ' Boucing tests
    dixuBounceTest x, VelocityX, Width, ScreenRect.Right
    dixuBounceTest y, VelocityY, Height, ScreenRect.bottom
End Sub

' Paint the sprite to the back buffer
Public Sub Paint()
    Static Top As Long
    Static Left As Long
    Static MaxWidth As Long
    Static MaxHeight As Long
    Dim R As RECT
    If MaxWidth = 0 Then
        Dim ddsd As DDSURFACEDESC
        ddsd.dwSize = Len(ddsd)
        Surface.GetSurfaceDesc ddsd
        MaxWidth = ddsd.dwWidth
        MaxHeight = ddsd.dwHeight
    End If
    ' Source rect for the sprite image
    R.Top = Top
    R.Left = Left
    R.Right = R.Left + Width
    R.bottom = R.Top + Height
    ' Blit image to the back buffer
    dixuBackBuffer.BltFast x, y, Surface, R, DDBLTFAST_NOCOLORKEY Or DDBLTFAST_WAIT
    ' Next image position
    Left = (Left + Width) Mod MaxWidth
    ' If row was done, next line and back to top if done
    If Left = 0 Then Top = (Top + Height) Mod MaxHeight
End Sub

' ********** Private **********

' Boucing test
Private Sub dixuBounceTest(ByRef Where As Long, ByRef Velocity As Long, ByVal Size As Long, ByVal Limit As Long)
    ' If sprite goes right or down
    If Velocity > 0 Then
        ' If bottom right corner out of screen
        If Where > Limit - Size Then
            ' Bounce
            Where = Limit - Size - (Where - (Limit - Size))
            ' Reverse speed
            Velocity = -Velocity
        End If
    ' Else sprite goes top or left
    Else
        ' If top left corner out of screen
        If Where < 0 Then
            ' Bounce
            Where = -Where
            ' Reverse speed
            Velocity = -Velocity
        End If
    End If
End Sub
