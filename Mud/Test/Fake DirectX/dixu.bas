Attribute VB_Name = "basDixu"
' dixu v 0.2, Copyright Patrice Scribe, 1997
' Changes from v 0.1 :
' - dixuFrontBuffer renamed to dixuPrimarySurface for consistency with DirectX documentation
' - dixuDrawBuffer renamed to dixuBackBufferDraw for consistency
' - support for custom code in the rendering loop (through dixuBackBufferClear)
' - support added for 3D (especially camera movements with collision detection)
' - other minor changes

Option Explicit

' 3D support
' Faces identifiers for rooms
' Note that "front" and "back" terms are related to the initial point of view :
'          Front
'      +-----------+
'      |           |
'      |           |
' Left |     ^     | Right
'      |   Camera  |
'      |           |
'      +-----------+
'           Back

Global Const DIXU_NOSPRITE = True

Global Const SRCAND = &H8800C6
Global Const SRCERASE = &H440328
Global Const SRCINVERT = &H660046
Global Const SRCPAINT = &HEE0086

Global Const dixuFaceDown = 0   ' Down face (floor)
Global Const dixuFaceTop = 1    ' Top face (ceiling)
Global Const dixuFaceFront = 5  ' Front face
Global Const dixuFaceLeft = 3   ' Left face
Global Const dixuFaceRight = 4  ' Right face
Global Const dixuFaceBack = 2   ' Back face (related to the subjective view)

Public dixuAppEnd As Boolean

' Win32 API
Const IMAGE_BITMAP = 0
Const LR_LOADFROMFILE = &H10
Const LR_CREATEDIBSECTION = &H2000
Const SRCCOPY = &HCC0020
Private Type BITMAP
        bmType          As Long
        bmWidth         As Long
        bmHeight        As Long
        bmWidthBytes    As Long
        bmPlanes        As Integer
        bmBitsPixel     As Integer
        bmBits          As Long
End Type

' GDI32
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
' USER32
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
' ...
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlCopyMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Sub FillMemory Lib "Kernerl32" Alias "RtlFillMemory" (ByVal l As Long, ByVal Length As Long, ByVal Fill As Byte)
Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, ByVal W As Long, ByVal E As Long, ByVal O As Long, ByVal W As Long, ByVal I As Long, ByVal u As Long, ByVal S As Long, ByVal C As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal F As String) As Long
Private Declare Function GetSystemPaletteEntries Lib "gdi32" (ByVal hdc As Long, ByVal First As Long, ByVal Size As Long, ptr As Any) As Long

' Flags for dixuInit
Global Const dixuInit3DDevice = 1   ' 3D enabled
Global Const dixuInitFullScreen = 2 ' Full screen mode
Global Const dixuInitRGB = 4        ' RGB driver
Global Const dixuInitMono = 8       ' Mono driver

' Don't compile if dixuSprite not needed
#If DIXU_NOSPRITE = 0 Then
Public dixuSprites As New Collection
#End If
     
Const Pi = 3.14159265358

' DirectDraw public objects
Public dixuDDraw As DirectDraw2
Public dixuPrimarySurface As DirectDrawSurface3
Public dixuBackBuffer As DirectDrawSurface3
Public dixuClipper As DirectDrawClipper
Public ScreenRect As RECT

' Direct3D Retained Mode public objects
Public dixuD3DRM As Direct3DRM
Public dixuD3DRMDevice As Direct3DRMDevice
Public dixuD3DRMViewport As Direct3DRMViewPort

' High-level 3D objects
Public dixuScene As Direct3DRMFrame
Public dixuCamera As Direct3DRMFrame

' Time values
Public dixuTime As Single
Public dixuLastTime As Single

' Private
Private bln3DDevice As Boolean      ' dixuInit3DDevice specified ?
Private blnFullScreen As Boolean    ' dixuInitFullScreen specified ?
Private SpritesRect As RECT         ' Not used yet
Private blnBackBufferClear As Boolean ' Back buffer to clear ?

' Camera values
Private sngCameraStep As Single ' Moving forward (or backward)
Private sngCameraCos As Single  ' For camera rotation
Private sngCameraSin As Single  ' For camera rotation

' Initializes DirectX
Sub dixuInit(ByVal Flags As Long, frm As Form, ByVal Width As Long, ByVal Height As Long, ByVal BitsPerPixel As Long)
    Dim ddsd As DDSURFACEDESC
    Dim ddc As DDSCAPS
    ' Camera default values
    sngCameraStep = 3
    sngCameraCos = Cos(10 * Pi / 180)
    sngCameraSin = Sin(10 * Pi / 180)
    ' 3D enabled ?
    bln3DDevice = (Flags And dixuInit3DDevice) <> 0
    ' Full screen mode ?
    blnFullScreen = (Flags And dixuInitFullScreen) <> 0
    blnFullScreen = True
    ' Initializes DirectDraw if not already done
    If dixuDDraw Is Nothing Then
        DirectDrawCreate ByVal 0&, dixuDDraw, Nothing
    End If
    ' Clear other DirectDraw objects if needed
    Set dixuD3DRMViewport = Nothing
    Set dixuD3DRMDevice = Nothing
    Set dixuBackBuffer = Nothing
    Set dixuPrimarySurface = Nothing
    ' DirectDraw part
    If blnFullScreen Then
        ' Full screen mode : change display mode
        dixuDDraw.SetCooperativeLevel frm.hWnd, DDSCL_EXCLUSIVE Or DDSCL_FULLSCREEN
        dixuDDraw.SetDisplayMode Width, Height, BitsPerPixel, 0, 0
        With ddsd
            ' Structure size
            .dwSize = Len(ddsd)
            ' Use DDSD_CAPS and BackBufferCount
            .dwFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT
            With .DDSCAPS
                .dwCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_FLIP Or DDSCAPS_COMPLEX ' Or DDSCAPS_SYSTEMMEMORY
                ' If 3D enabled
                If bln3DDevice Then
                    .dwCaps = .dwCaps Or DDSCAPS_3DDEVICE
                End If
            End With
            ' One back buffer
            .dwBackBufferCount = 1
        End With
        ' Creates buffers
        dixuDDraw.CreateSurface ddsd, dixuPrimarySurface, Nothing
        ' Retrieve back buffer
        ddc.dwCaps = DDSCAPS_BACKBUFFER
        dixuPrimarySurface.GetAttachedSurface ddc, dixuBackBuffer
        ' Keep screen rect
        ScreenRect.Left = 0
        ScreenRect.Top = 0
        ScreenRect.Right = Width
        ScreenRect.bottom = Height
    Else
        ' Windowed mode
        frm.Move frm.Left, frm.Top, Width, Height
        dixuDDraw.SetCooperativeLevel frm.hWnd, DDSCL_NORMAL
        ' Create the front buffer
        With ddsd
            .dwSize = Len(ddsd)
            .dwFlags = DDSD_CAPS
            .DDSCAPS.dwCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_SYSTEMMEMORY
        End With
        If bln3DDevice Then ddsd.DDSCAPS.dwCaps = ddsd.DDSCAPS.dwCaps Or DDSCAPS_3DDEVICE
        dixuDDraw.CreateSurface ddsd, dixuPrimarySurface, Nothing
        ' Create and attach clipper
        DirectDrawCreateClipper 0, dixuClipper, Nothing
        dixuClipper.SetHWnd 0, frm.hWnd
        dixuPrimarySurface.SetClipper dixuClipper
        ' Create back buffer
        With ddsd
            .dwSize = Len(ddsd)
            .dwFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
            .DDSCAPS.dwCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_3DDEVICE
            .dwWidth = frm.ScaleWidth
            .dwHeight = frm.ScaleHeight
        End With
        dixuDDraw.CreateSurface ddsd, dixuBackBuffer, Nothing
        dixuBackBuffer.SetClipper dixuClipper
        ' Keep screen rect
        ScreenRect.Left = 0
        ScreenRect.Top = 0
        ScreenRect.Right = frm.ScaleWidth - 1
        ScreenRect.bottom = frm.ScaleHeight - 1
        frm.Show
    End If
    ' Now work on Direct3D part
    If bln3DDevice Then
        Dim d3d As Direct3D
        Dim fds As D3DFINDDEVICESEARCH
        Dim fdr As D3DFINDDEVICERESULT
        Dim ddsdFront As DDSURFACEDESC
        If d3d Is Nothing Then
            ' Get D3D interface (QueryInterface stuff for C/C++)
            Set d3d = dixuDDraw
        End If
        ' Search for the color model driver
        fds.dwSize = Len(fds)
        fds.dwFlags = D3DFDS_COLORMODEL
        If (Flags And dixuInitRGB) <> 0 Then
            fds.dcmColorModel = D3DCOLOR_RGB
        Else
            fds.dcmColorModel = D3DCOLOR_MONO
        End If
        fdr.dwSize = Len(fdr)
        ' Find the driver
        d3d.FindDevice fds, fdr
        ' If 256 colors, set the palette
        If BitsPerPixel = 8 Then
            Dim ColorTable(0 To 255) As PALETTEENTRY
            Dim I As Long
            Dim Palette As DirectDrawPalette
            Debug.Print GetSystemPaletteEntries(frm.hdc, 0, 256, ColorTable(0))
            For I = 0 To 255
                ColorTable(I).peFlags = &H40
            Next
            dixuDDraw.CreatePalette 4, ColorTable(0), Palette, Nothing
            dixuBackBuffer.SetPalette Palette
        End If
        ' If needed, create top-level objects
        If dixuD3DRM Is Nothing Then
            Direct3DRMCreate dixuD3DRM
            dixuD3DRM.CreateFrame Nothing, dixuScene
            dixuD3DRM.CreateFrame dixuScene, dixuCamera
        End If
        ' Create the device from the existing DirectDraw back buffer
        dixuD3DRM.CreateDeviceFromSurface fdr.GUID, dixuDDraw, dixuBackBuffer, dixuD3DRMDevice
        dixuD3DRM.CreateViewport dixuD3DRMDevice, dixuCamera, 0, 0, dixuD3DRMDevice.GetWidth, dixuD3DRMDevice.GetHeight, dixuD3DRMViewport
    End If
End Sub
' Clean up objects
Public Sub dixuDone()
    Dim I As Long
    ' Clear sprites (don't compile if dixuSprite not needed)
    #If DIXU_NOSPRITE = 0 Then
    ' Clear sprites
    For I = 1 To dixuSprites.Count
        dixuSprites.Remove 1
    Next
    #End If
    ' Reset display mode
    If blnFullScreen Then
        dixuDDraw.FlipToGDISurface
        dixuDDraw.RestoreDisplayMode
    End If
    dixuDDraw.SetCooperativeLevel 0, DDSCL_NORMAL
    If bln3DDevice Then
        ' Clear 3D objects
        Set dixuCamera = Nothing
        Set dixuScene = Nothing
        Set dixuD3DRMViewport = Nothing
        Set dixuD3DRMDevice = Nothing
        Set dixuD3DRM = Nothing
    End If
    ' Clear DirectDraw objects
    Set dixuClipper = Nothing
    Set dixuBackBuffer = Nothing
    Set dixuPrimarySurface = Nothing
    Set dixuDDraw = Nothing
End Sub

' Clears the back buffer
Public Sub dixuBackBufferClear()
    Dim fx As DDBLTFX
    With fx
        .dwSize = Len(fx)
        .dwFillColor = RGB(0, 0, 0)
    End With
    dixuBackBuffer.Blt ScreenRect, Nothing, ScreenRect, DDBLT_COLORFILL Or DDBLT_WAIT, fx
End Sub
' Draws the back buffer
Public Sub dixuBackBufferDraw()
    On Error GoTo UpdateScreen_Error
    If bln3DDevice Then
        ' Render 3D scene
        dixuScene.Move 1
        dixuD3DRMViewport.Clear
        dixuD3DRMViewport.Render dixuScene
        dixuD3DRMDevice.Update
    Else
        ' Clear the back buffer
        If blnBackBufferClear Then dixuBackBufferClear
    End If
    'Exit Sub
    ' Render sprites (don't compile if dixuSprite not needed)
    #If DIXU_NOSPRITE = 0 Then
    If dixuSprites.Count <> 0 Then dixuSpritesDraw
    #End If
    If blnFullScreen Then
        ' Display frame to screen
        Do
            dixuPrimarySurface.Flip Nothing, 0
        Loop Until Err.Number = 0
    Else
        ' Display frame to window
        Do
            Dim fx As DDBLTFX
            fx.dwSize = Len(fx)
            fx.dwRop = SRCCOPY
            ' BltFast not usable with clippers...
            dixuPrimarySurface.Blt ScreenRect, dixuBackBuffer, ScreenRect, DDBLT_ROP Or DDBLT_WAIT, fx
        Loop Until Err.Number = 0
    End If
    blnBackBufferClear = True
    Exit Sub
UpdateScreen_Error:
    App.LogEvent Err.Description
    Err.Clear
    If Err.Number = DDERR_SURFACELOST Then
        dixuBackBuffer.Restore
        dixuPrimarySurface.Restore
    End If
    Resume
End Sub

' ********** DirectDraw support **********

' Loads a bitmap in a DirectDraw surface (to change for NT compatibility)
Public Function dixuCreateSurfaceFromBitmap(ByVal strFile As String, sysmemFlag As Integer) As DirectDrawSurface3
    Dim hbm As Long                 ' Handle on bitmap
    Dim bm As BITMAP                ' Bitmap header
    Dim ddsd As DDSURFACEDESC       ' Surface description
    Dim dds As DirectDrawSurface3   ' Created surface
    Dim hdcImage As Long            ' Handle on image
    Dim lhdc As Long                ' Handle on surface context
    ' Load bitmap
    hbm = LoadImage(ByVal 0&, strFile, IMAGE_BITMAP, 0, 0, LR_LOADFROMFILE Or LR_CREATEDIBSECTION)
    ' Get bitmap info
    GetObject hbm, Len(bm), bm
    ' Fill surface description
    With ddsd
        .dwSize = Len(ddsd)
        .dwFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
        
        If sysmemFlag = True Then
        .DDSCAPS.dwCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
        Else
        .DDSCAPS.dwCaps = DDSCAPS_OFFSCREENPLAIN 'Or DDSCAPS_SYSTEMMEMORY
        End If
        
        .dwWidth = bm.bmWidth
        .dwHeight = bm.bmHeight
    End With
    ' Create surface
    dixuDDraw.CreateSurface ddsd, dds, Nothing
    ' Create memory device
    hdcImage = CreateCompatibleDC(ByVal 0&)
    ' Select the bitmap in this memory device
    SelectObject hdcImage, hbm
    ' Restore the surface
    dds.Restore
    ' Get the surface's DC
    dds.GetDC lhdc
    ' Copy from the memory device to the DirectDrawSurface
    StretchBlt lhdc, 0, 0, ddsd.dwWidth, ddsd.dwHeight, hdcImage, 0, 0, bm.bmWidth, bm.bmHeight, SRCCOPY
    ' Release the surface's DC
    dds.ReleaseDC lhdc
    ' Release the memory device and the bitmap
    DeleteDC hdcImage
    DeleteObject hbm
    ' Returns the new surface
    Set dixuCreateSurfaceFromBitmap = dds
End Function

' ********** Direct3D Retained Mode support **********

' >>>>>>>>>> Camera

' Set camera steps values
Public Sub dixuSetCameraMoves(ByVal Step As Single, ByVal Angle As Single)
    sngCameraStep = Step
    sngCameraCos = Cos(Angle * Pi / 180)
    sngCameraSin = Sin(Angle * Pi / 180)
End Sub

' Move camera according to KeyCode (features dectection collision code for moving forward - can still pass through walls when moving backward !)
Public Sub dixuCameraMove(ByVal KeyCode As Long)
    On Error GoTo dixuCameraMove_Error
    Select Case KeyCode
        Case vbKeyDown      ' Move backward
            dixuCamera.SetPosition dixuCamera, 0, 0, -sngCameraStep
        Case vbKeyEscape    ' Ends app
            dixuAppEnd = True
        Case vbKeyLeft      ' Turn left
            dixuCamera.SetOrientation dixuCamera, -sngCameraSin, 0, sngCameraCos, 0, 1, 0
        Case vbKeyRight     ' Turn right
            dixuCamera.SetOrientation dixuCamera, sngCameraSin, 0, sngCameraCos, 0, 1, 0
        Case vbKeyUp        ' Move forward
            Dim PickedArray As Direct3DRMPickedArray
            Dim MeshBuilder As Direct3DRMMeshBuilder
            Dim FrameArray As Direct3DRMFrameArray
            Dim PickDesc As D3DRMPICKDESC
            Dim Distance As Single
            Dim CameraPosition As D3DVECTOR
            Dim Screen As D3DRMVECTOR4D ' Screen coordinates
            Dim World As D3DVECTOR      ' World coordinates
            ' Retrieves intersected visuals
            dixuD3DRMViewport.Pick ScreenRect.Right \ 2, ScreenRect.bottom \ 2, PickedArray
            If PickedArray.GetSize <> 0 Then
                ' Retrieve the first visual, parent frame (?), and details
                PickedArray.GetPick 0, MeshBuilder, FrameArray, PickDesc
                ' Copy screen coordinates
                With PickDesc.vPosition
                    Screen.x = .x
                    Screen.y = .y
                    Screen.z = .z
                    Screen.W = 1
                End With
                ' Transform to world coordinates
                dixuD3DRMViewport.InverseTransform World, Screen
                ' Get camera position
                dixuCamera.GetPosition dixuScene, CameraPosition
                ' Compute distance between intersection and camera
                Distance = Sqr((CameraPosition.x - World.x) ^ 2 + (CameraPosition.z - World.z) ^ 2)
                ' Enough distance : move
                If Distance > 2 * sngCameraStep Then dixuCamera.SetPosition dixuCamera, 0, 0, sngCameraStep
            Else
                ' No visual : move
                dixuCamera.SetPosition dixuCamera, 0, 0, sngCameraStep
            End If
            ' Collision detection defeated if key pressed repeatidly and scene not rendered ?!
            dixuBackBufferDraw
    End Select
    Exit Sub
' Trap transient errors (division by zero)
dixuCameraMove_Error:
    Debug.Print Hex$(Err.Number)
    Resume
End Sub


' Creates a room (a cube whose faces are visible from inside)
Public Function dixuCreateRoom() As Direct3DRMMeshBuilder
    Dim aVertices(0 To 8) As D3DVECTOR  ' Vertices array
    Dim aNormals(0) As D3DVECTOR        ' Normal array (not used)
    Dim aFaces(1 To 31) As Long         ' Faces array
    Dim MeshBuilder As Direct3DRMMeshBuilder
    Dim I As Integer
    ' Coordinates for floor vertices
    aVertices(0).x = -0.5
    aVertices(0).y = 0
    aVertices(0).z = -0.5
    aVertices(1).x = -0.5
    aVertices(1).y = 0
    aVertices(1).z = 0.5
    aVertices(2).x = 0.5
    aVertices(2).y = 0
    aVertices(2).z = 0.5
    aVertices(3).x = 0.5
    aVertices(3).y = 0
    aVertices(3).z = -0.5
    ' Copy floor vertices to ceiling vertices
    For I = 0 To 3
        ' Copy vertex
        aVertices(4 + I) = aVertices(I)
        ' Change height
        aVertices(4 + I).y = 1
    Next
    ' Fill faces array (number of vertices and vertex index for each face)
    ' Faces are described clockwise
    ' Floor
    aFaces(1) = 4 ' 4 vertices
    aFaces(2) = 0
    aFaces(3) = 1
    aFaces(4) = 2
    aFaces(5) = 3
    ' Ceiling
    aFaces(6) = 4
    aFaces(7) = 7
    aFaces(8) = 6
    aFaces(9) = 5
    aFaces(10) = 4
    ' Front wall
    aFaces(11) = 4
    aFaces(12) = 1
    aFaces(13) = 5
    aFaces(14) = 6
    aFaces(15) = 2
    ' Left wall
    aFaces(16) = 4
    aFaces(17) = 0
    aFaces(18) = 4
    aFaces(19) = 5
    aFaces(20) = 1
    ' Right wall
    aFaces(21) = 4
    aFaces(22) = 2
    aFaces(23) = 6
    aFaces(24) = 7
    aFaces(25) = 3
    ' Back wall
    aFaces(26) = 4
    aFaces(27) = 3
    aFaces(28) = 7
    aFaces(29) = 4
    aFaces(30) = 0
    ' Terminator
    aFaces(31) = 0
    ' Create and return object
    dixuD3DRM.CreateMeshBuilder MeshBuilder
    MeshBuilder.AddFaces 8, aVertices(0), 0, aNormals(0), aFaces(1), Nothing
    Set dixuCreateRoom = MeshBuilder
End Function

' ********** Sprites support **********
' (don't compile if not needed)

#If DIXU_NOSPRITE = 0 Then
Private Sub dixuSpritesDraw()
    Dim Sprite As dixuSprite
    dixuLastTime = dixuTime
    dixuTime = Timer
    ' If runs at midnight !
    While dixuTime < dixuLastTime
        dixuTime = dixuTime + 86400 ' 1 day
    Wend
    ' Move and paint all sprites
    For Each Sprite In dixuSprites
        Sprite.Move
        Sprite.Paint
    Next
End Sub

Public Sub dixuSpriteRemove(Sprite As dixuSprite)
    dixuSprites.Remove Sprite.Key
    Set Sprite = Nothing
End Sub

#End If

