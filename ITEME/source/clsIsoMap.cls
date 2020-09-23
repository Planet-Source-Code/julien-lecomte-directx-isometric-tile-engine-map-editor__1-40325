VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIsoMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'*******************************************************************************************
' ITEME
'*******************************************************************************************
' Julien Lecomte
' webmaster@amanitamuscaria.org
' http://www.amanitamuscaria.org
' Feel free to use, abuse or distribute. (USUS, FRUCTUS, & ABUSUS)
' If you improve it, tell me !
' Don't take credit for what you didn't create. Thanks.
'*******************************************************************************************
'// Nice references :
'// http://www.gamedev.net/reference/articles/article744.asp
'// http://www.gamedev.net/reference/articles/article747.asp
'// http://www.gamedev.net/reference/articles/article748.asp
'// http://www.mvps.org/directx/articles/blit3d.htm

Public Enum OTHER_MODES
    BLACKEN_VERTICES
    WHITEN_VERTICES
End Enum

Private Type ISO_TILE
    lTileId           As Long       '// Id of tile; 0 means tile is empty
    lHeight(0 To 3)   As Long       '// 0 is basic 'sea level'
    ptVertice(0 To 3) As POINTFLOAT '// Lookup vertice 0=far-left; 1=top; 2=bottom; 3=far-right
End Type

Private Type SURFACE_LOOKUP
    objTexture As DxVBLibA.Direct3DTexture8
    sFileName  As String
End Type
    
Private Type D3DCUSTOMVERTEX
    X        As Single
    Y        As Single
    Z        As Single
    rhw      As Single
    color    As Long
    specular As Long
    tu       As Single
    tv       As Single
End Type

Private Const MAX_HEIGHT = 7& '// 0 based
Private Const DEFAULT_TILE_SIZE = 32&
Private TILE_SIZE&, TILE_FOURTH&, TILE_HALF& '// Used as lookup values, do not individually set

'// Temporary private variables
Private m_lLastTick&
Private m_lMaxFlip&             '// Maximum flips of textures per tile.
Private m_lMapPxWidth!, m_lMapPxHeight! '// In pixels
Private m_lSurfPxWidth!, m_lSurfPxHeight!
Private ptTopLeft As POINTFLOAT     '// Current top-left position of map

'// Lookup tables
Private lu_Blackness(0 To 32) As Long '// White to Black scale

'// Map properties
'// These should be saved in the save routine
Private m_lTileSizeNoZoom&                   '// tile size (as set in default), not always = to TILE_SIZE
Private m_lBackColor&                        '// Color of map background
Private m_lVerticeColorRange(0 To 1) As Long '// Color of vertice at low & high
Private m_losTileCount&                      '// Number of texture surfaces used in map;
Private m_lMapSizeWidth&, m_lMapSizeHeight&  '// Size in tile UNITS !!!
Private m_aTileMap() As ISO_TILE             '// All the tiles used in map
Private osTile() As SURFACE_LOOKUP           '// Texture array

'// DirectX objects
Private oDirectX    As DxVBLibA.DirectX8
Private oDirect3D   As DxVBLibA.Direct3D8
Private oDevice     As DxVBLibA.Direct3DDevice8
Private oD3D8       As DxVBLibA.D3DX8

'-----------------------------------------------------------------------------
'// GDI32
'-----------------------------------------------------------------------------
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long

Private Sub Class_Initialize()
    Dim I&
    Dim lColor&
    
    '// Create lookup
    For I = 0& To 32&
        lColor = I * 8&
        If lColor = 256& Then lColor = 255&
        lu_Blackness(32 - I) = RGB(lColor, lColor, lColor)
    Next
    
    DirectX_Destroy
End Sub

Private Sub Class_Terminate()
    DirectX_Destroy '// Be sure to terminate the DirectX
End Sub

Private Function ddVertex(tu!, tv!) As D3DCUSTOMVERTEX
    '// Vertex wrapper
    With ddVertex
        .color = vbWhite
        .tu = tu
        .tv = tv
    End With
End Function

Public Sub DirectX_Destroy()
    '// This function actually also initializes everything
On Local Error Resume Next
    Dim I&
    
    For I = 1 To m_losTileCount
        Set osTile(I).objTexture = Nothing
    Next
    
    '// Set defaults back
    ptTopLeft.X = 0!
    ptTopLeft.Y = 0!
    m_lMaxFlip = 0&
    m_losTileCount = 0&
    m_lLastTick = -1&
    m_lVerticeColorRange(0) = 0&
    m_lVerticeColorRange(1) = 32&
    m_lBackColor = vbBlack
    lTileSize = DEFAULT_TILE_SIZE
    
    Set oDirect3D = Nothing
    Set oDevice = Nothing
    Set oD3D8 = Nothing
    Set oDirectX = Nothing
    
    Erase osTile
    Erase m_aTileMap
End Sub

Public Function DirectX_Initialize(objSurface As Object, Optional bFullScreen As Boolean) As Boolean
    '// Returns true if success in creation.
    Dim sFilePath$
    Dim oDispMode  As DxVBLibA.D3DDISPLAYMODE
    Dim oD3DWindow As DxVBLibA.D3DPRESENT_PARAMETERS

    '// Initialize
    Set oDirectX = New DxVBLibA.DirectX8
    Set oD3D8 = New DxVBLibA.D3DX8
    Set oDirect3D = oDirectX.Direct3DCreate
    m_lLastTick = 0&

On Local Error GoTo ErrHandler
    '// Display
    
    If bFullScreen Then
        MsgBox "Windowed screen is only permitted right now."
    Else
        oDirect3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, oDispMode
        With oDispMode
            If .Format = D3DFMT_UNKNOWN Then .Format = D3DFMT_R5G6B5 '// 24 bit RGB pixel format
            objSurface.ScaleMode = vbPixels
            .Width = objSurface.ScaleWidth
            .Height = objSurface.ScaleHeight
            .RefreshRate = 0&
        End With
        
        m_lSurfPxWidth = objSurface.ScaleWidth
        m_lSurfPxHeight = objSurface.ScaleHeight
        
        '// Window
        With oD3DWindow
            .Windowed = 1&
            .SwapEffect = D3DSWAPEFFECT_DISCARD 'D3DSWAPEFFECT_COPY_VSYNC
            .BackBufferCount = 1&
            .BackBufferFormat = oDispMode.Format
            '// You should check if this is supported
            '// If it's not, then swapeffect must be changed back to D3DSWAPEFFECT_COPY_VSYNC
            .MultiSampleType = D3DMULTISAMPLE_2_SAMPLES
        End With

        '// Create the device
        '// Maybe we could use D3DCREATE_HARDWARE_VERTEXPROCESSING
        Set oDevice = oDirect3D.CreateDevice(D3DADAPTER_DEFAULT, _
                                         D3DDEVTYPE_HAL, _
                                         objSurface.hwnd, _
                                         D3DCREATE_SOFTWARE_VERTEXPROCESSING, _
                                         oD3DWindow)
    End If

    '// Set device mode
    With oDevice
        .SetVertexShader D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR
        '// Unused stuff
        '// Source & dest blend
'        .SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
'        .SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
        '// Alpha blending; not used anyway (at least not yet)
'        .SetRenderState D3DRS_ALPHABLENDENABLE, True
    End With
    
    DirectX_Initialize = True
ErrHandler:
End Function

Public Function DirectX_LoadSurface(sPath$, sFileName$) As Long
    '// sPath must already contain backslash
    '// Returns -1 if failed, otherwise returns texture's id
    
On Local Error GoTo ErrHandler
    DirectX_LoadSurface = -1&
    m_losTileCount = m_losTileCount + 1&
    ReDim Preserve osTile(1 To m_losTileCount) As SURFACE_LOOKUP
    
    Set osTile(m_losTileCount).objTexture = oD3D8.CreateTextureFromFile(oDevice, sPath & sFileName)
    osTile(m_losTileCount).sFileName = sFileName
    DirectX_LoadSurface = m_losTileCount
ErrHandler:
End Function

Public Function DirectX_LoadSurfaceEx(sPath$, sFileName$) As Long
    '// sPath must already contain backslash
    '// Adds texture only if it didn't exist yet.
    '// Returns -1 if failed, otherwise returns texture's id
    Dim I&
    
On Local Error GoTo ErrHandler
    DirectX_LoadSurfaceEx = -1&
    
    For I = 1 To m_losTileCount
        If osTile(I).sFileName = sFileName Then
            DirectX_LoadSurfaceEx = I&
            Exit Function
        End If
    Next
    
    DirectX_LoadSurfaceEx = DirectX_LoadSurface(sPath$, sFileName$)
ErrHandler:
End Function

Public Function DirectX_Refresh() As Long
    '// Function returns 0 if fps was not calculated,
    '// otherwise it returns the fps
    Dim I&, J&, K&, lX!, lY!
    Dim lCurId&
    Dim lHeight&, lWidth&
    Dim aVertice(0 To 3)  As D3DCUSTOMVERTEX
    Dim lTileId&, lCurFlip&
    Dim lColor&
    Dim bHasVerticeRange As Boolean
    Static lFps&
    Static lFlip&

    '// Precalculate some stuff
    bHasVerticeRange = CBool((m_lVerticeColorRange(0) Or m_lVerticeColorRange(1)) > 0)

    With oDevice
        '// Clear then start drawing
        .Clear 0&, ByVal 0&, D3DCLEAR_TARGET, m_lBackColor, 1#, 0&
        .BeginScene
        
        For J = 0 To m_lMapSizeHeight
            For I = 0 To m_lMapSizeWidth
                '// Get tile to show
                lTileId = m_aTileMap(I, J).lTileId

                '// If tile still exists then draw it
                If lTileId Then
                    .SetTexture 0&, osTile(lTileId).objTexture
                    
                    '// Set up vertices.
                    aVertice(0) = ddVertex(0!, 0!)
                    aVertice(1) = ddVertex(1!, 0!)
                    aVertice(2) = ddVertex(0!, 1!)
                    aVertice(3) = ddVertex(1!, 1!)
                    
                    '// Set up vertice axis X
                    aVertice(0).X = ptTopLeft.X + m_aTileMap(I, J).ptVertice(0).X
                    aVertice(1).X = ptTopLeft.X + m_aTileMap(I, J).ptVertice(1).X
                    aVertice(2).X = ptTopLeft.X + m_aTileMap(I, J).ptVertice(2).X
                    aVertice(3).X = ptTopLeft.X + m_aTileMap(I, J).ptVertice(3).X
                    
                    '// Set up vertice axis Y
                    aVertice(0).Y = ptTopLeft.Y + m_aTileMap(I, J).ptVertice(0).Y - m_aTileMap(I, J).lHeight(0) * TILE_FOURTH
                    aVertice(1).Y = ptTopLeft.Y + m_aTileMap(I, J).ptVertice(1).Y - m_aTileMap(I, J).lHeight(1) * TILE_FOURTH
                    aVertice(2).Y = ptTopLeft.Y + m_aTileMap(I, J).ptVertice(2).Y - m_aTileMap(I, J).lHeight(2) * TILE_FOURTH
                    aVertice(3).Y = ptTopLeft.Y + m_aTileMap(I, J).ptVertice(3).Y - m_aTileMap(I, J).lHeight(3) * TILE_FOURTH
                    
                    '// Set up blackness per vertice height
                    If bHasVerticeRange Then '// Save some time if both are set to white vertices.
                        lHeight = m_aTileMap(I, J).lHeight(0)
                        lColor = (m_lVerticeColorRange(1) * lHeight + m_lVerticeColorRange(0) * (7 - lHeight)) \ 8
                        aVertice(0).color = lu_Blackness(lColor)
                        lHeight = m_aTileMap(I, J).lHeight(1)
                        lColor = (m_lVerticeColorRange(1) * lHeight + m_lVerticeColorRange(0) * (7 - lHeight)) \ 8
                        aVertice(1).color = lu_Blackness(lColor)
                        lHeight = m_aTileMap(I, J).lHeight(2)
                        lColor = (m_lVerticeColorRange(1) * lHeight + m_lVerticeColorRange(0) * (7 - lHeight)) \ 8
                        aVertice(2).color = lu_Blackness(lColor)
                        lHeight = m_aTileMap(I, J).lHeight(3)
                        lColor = (m_lVerticeColorRange(1) * lHeight + m_lVerticeColorRange(0) * (7 - lHeight)) \ 8
                        aVertice(3).color = lu_Blackness(lColor)
                    End If
                    
                    '// We render 2 triangles per call
                    '// Microsoft recommends rendering 300 triangles per call
                    '// even though most AGP-3D cards now support 1000+ triangles per call
                    '// I got no idea (yet) how to do this.
                    '// Well actually I do, but this isn't important yet.
                    .DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2&, aVertice(0), Len(aVertice(0))
                End If
            Next
        Next

        .EndScene
        .Present ByVal 0&, ByVal 0&, 0&, ByVal 0&
    End With
    
    '// Calculate FPS
    If lFps = 30 Then
        If m_lLastTick = -1& Then
            m_lLastTick = GetTickCount
        Else
            '// 1000 * 30 \ ((GetTickCount - m_lLastTick) + 1)
            DirectX_Refresh = 30000 \ (GetTickCount - m_lLastTick + 1)
            m_lLastTick = GetTickCount
        End If
        lFps = 0&
    End If
    lFps = lFps + 1
End Function

Public Function IsValid() As Boolean
    Dim bAllTests As Boolean
    '// This function returns true if all initial parameters have been set
    '//
On Local Error GoTo ErrInvalid
    bAllTests = True
    bAllTests = bAllTests And CBool(m_lMapSizeWidth > 0)
    bAllTests = bAllTests And CBool(m_lMapSizeHeight > 0)
    bAllTests = bAllTests And CBool(TILE_SIZE > 0)
    '// Include all extra tests here
    
    IsValid = bAllTests
ErrInvalid:
End Function

Public Property Let lBackColor(lValue&)
    m_lBackColor = lValue
End Property

Public Property Let lTileSize(lValue&)
    '// Check if tile size is ok
    '// We could allow maybe other sizes, but it's not fully tested
    Select Case lValue
        Case 32, 64, 128, 256
            m_lTileSizeNoZoom = lValue
            Map_Zoom 0&
        Case Else
            Err.Raise 1&, , "Tile size must either be between 32 and 256 pixels"
    End Select
End Property

Private Sub Map_Create(Optional lDefaultTile = -1)
    Dim I&, J&
    Dim lX&, lY&

    '// Create map table
On Local Error GoTo SkipSub
    m_lMapPxWidth = -1
    m_lMapPxHeight = -1
    
    For I = LBound(m_aTileMap, 1) To UBound(m_aTileMap, 1)
        For J = LBound(m_aTileMap, 2) To UBound(m_aTileMap, 2)
            '// Create tile map
            lX = I * TILE_SIZE + (J And 1) * TILE_HALF
            lY = J * TILE_FOURTH
            
            '// get max values
            If lX > m_lMapPxWidth Then m_lMapPxWidth = lX
            If lY > m_lMapPxHeight Then m_lMapPxHeight = lY
            
            '// Set vertices
            m_aTileMap(I, J).ptVertice(0).X = lX
            m_aTileMap(I, J).ptVertice(0).Y = lY + TILE_FOURTH
            m_aTileMap(I, J).ptVertice(1).X = lX + TILE_HALF
            m_aTileMap(I, J).ptVertice(1).Y = lY
            m_aTileMap(I, J).ptVertice(2).X = lX + TILE_HALF
            m_aTileMap(I, J).ptVertice(2).Y = lY + TILE_HALF
            m_aTileMap(I, J).ptVertice(3).X = lX + TILE_SIZE
            m_aTileMap(I, J).ptVertice(3).Y = lY + TILE_FOURTH

            If Not (lDefaultTile = -1) Then m_aTileMap(I, J).lTileId = lDefaultTile
        Next
    Next
    
    m_lMapPxWidth = m_lMapPxWidth + TILE_SIZE
    m_lMapPxHeight = m_lMapPxHeight + TILE_HALF

SkipSub:
End Sub

Public Property Get Map_GetTileXY(X!, Y!, ByRef lX&, ByRef lY&, hdcRefTile&) As Boolean
    Dim lPixel&
    Dim lXRef&, lYRef&
    
    lX = X \ TILE_SIZE
    lY = (Y \ TILE_HALF) * 2
    lXRef = X Mod TILE_SIZE
    lYRef = Y Mod TILE_HALF
    
    lPixel = GetPixel(hdcRefTile, lXRef, lYRef)
    Select Case lPixel
        Case vbRed:   lX = lX - 1: lY = lY - 1
        Case vbYellow:             lY = lY - 1
        Case vbBlue:               lY = lY + 1
        Case vbGreen: lX = lX - 1: lY = lY + 1
    End Select
    
    If lX >= 0 And lY >= 0 Then
        Map_GetTileXY = CBool(lX <= m_lMapSizeWidth And lY <= m_lMapSizeHeight)
    End If
End Property

Public Sub Map_SetSize(lWidth&, lHeight&, Optional lDefaultTile = -1)
    '// Size the map
    m_lMapSizeWidth = lWidth
    m_lMapSizeHeight = lHeight
    ReDim m_aTileMap(0 To m_lMapSizeWidth, 0 To m_lMapSizeHeight) As ISO_TILE
    
    Map_Create lDefaultTile
End Sub

Public Sub Map_SetTileHeight(lX&, lY&, lHeight&, lVertices&)
    '// lVertices is used as bit mask
    '// bit 1 = vertice 0
    '// bit 2 = vertice 1
    '// bit 3 = vertice 2
    '// bit 4 = vertice 3
    '// Thus a value of binary(1111) = dec(1+2+4+8) = dec(15) means all 4 vertices
    '// Thus a value of binary(0101) = dec(0+2+0+8) = dec(10) means all vertices 1 & 3
    Dim ptSearch As POINTFLOAT
    Dim lVertice&
    Dim I&, J&, K&, L&
    Dim lBitCurMask&
    Const lBitMask = &HFF&
    Const lMaskShift = &H100&
    Dim lNewValue&
    
On Local Error Resume Next '// Ignore all subscript out of range errors; faster than error traping
    For I = 0 To 3
        lBitCurMask = 2 ^ I
        If (lVertices And lBitCurMask) = lBitCurMask Then
            Map_SetTileHeightNewValue m_aTileMap(lX, lY).lHeight(I), lHeight '// Byref sub
            
            '// Find all same vertices, and add the lHeight value
            '// Any improvement of this is *very* welcome
            ptSearch = m_aTileMap(lX, lY).ptVertice(I)
            For K = (lX - 2) To (lX + 2)
                For J = (lY - 2) To (lY + 2)
                    If Not ((lX = K) And (lY = J)) Then
                        For L = 0 To 3
                            If ptSearch.X = m_aTileMap(K, J).ptVertice(L).X Then
                                If ptSearch.Y = m_aTileMap(K, J).ptVertice(L).Y Then
                                    Map_SetTileHeightNewValue m_aTileMap(K, J).lHeight(L), lHeight
                                End If
                            End If
                        Next
                    End If
                Next
            Next
        End If
    Next
End Sub

Private Sub Map_SetTileHeightNewValue(ByRef lOldValue&, lAddValue&)
    lOldValue = lOldValue + lAddValue
    '// Prevent any value under 0 and any value over 8
    If lOldValue < 0 Then lOldValue = 0
    If lOldValue > MAX_HEIGHT Then lOldValue = MAX_HEIGHT
End Sub

Public Sub Map_SetTileId(lX&, lY&, lTileId&)
On Local Error GoTo ErrHandler:
    m_aTileMap(lX, lY).lTileId = lTileId
ErrHandler:
End Sub

Public Function Map_Zoom(ByVal lValue&, Optional ByRef lTileSize&) As Boolean
    '// Returns FALSE if zoom failed
    '// If lValue is positive, then map zooms out
    '// If value is 0&, then map is set back to it's default
    Dim lVal&
    
    If lValue = 0 Then
        TILE_SIZE = m_lTileSizeNoZoom
    Else
        '// Check if you can (un)zoom to have a correct TILE_FOURTH size
        lVal = TILE_SIZE
        If (lValue > 0) Then
            lValue = lValue + 1
            If (lVal * lValue) > 256 Then Exit Function
            TILE_SIZE = lVal * lValue
        Else
            lValue = Abs(lValue) + 1
            If (lVal \ lValue) < 8 Then Exit Function
            TILE_SIZE = lVal \ lValue
        End If
    End If
    
    lTileSize = TILE_SIZE
    TILE_HALF = TILE_SIZE \ 2
    TILE_FOURTH = TILE_SIZE \ 4
    Map_Create
    UpdateXY 0, 0 '// Reinitialize if needed
    
    Map_Zoom = True
End Function

Public Sub SetMode_Fill(lMode As CONST_D3DFILLMODE)
    oDevice.SetRenderState D3DRS_FILLMODE, lMode
End Sub

Public Sub SetMode_Render(ByVal lVal&)
    '// lval =0 or lVal =1
    oDevice.SetRenderState D3DRS_MULTISAMPLE_ANTIALIAS, lVal
End Sub

Public Sub SetMode_VerticeColorRange(lScale&, bLow As Boolean)
    '// lColorMax is the maximum color
    If lScale < 0 Or lScale > 32 Then
        Err.Raise 1&, , "Scale must be from 0 to 32"
    Else
        m_lVerticeColorRange(IIf(bLow, 0, 1)) = lScale
    End If
End Sub

Public Sub UpdateXY(lXValue&, lYValue&)
    If (m_lSurfPxWidth - m_lMapPxWidth) < 0 Then
        ptTopLeft.X = ptTopLeft.X + lXValue
        If lXValue > 0 Then
            If ptTopLeft.X > 0 Then ptTopLeft.X = 0 '// Axis is inverted...
        ElseIf lXValue < 0 Then
            If ptTopLeft.X < (m_lSurfPxWidth - m_lMapPxWidth) Then ptTopLeft.X = (m_lSurfPxWidth - m_lMapPxWidth)
        End If
    Else
        ptTopLeft.X = 0 '// Reinitialize if needed
    End If

    If (m_lSurfPxHeight - m_lMapPxHeight) < 0 Then
        ptTopLeft.Y = ptTopLeft.Y + lYValue
        If lYValue > 0 Then
            If ptTopLeft.Y > 0 Then ptTopLeft.Y = 0 '// Axis is inverted...
        ElseIf lYValue < 0 Then
            If ptTopLeft.Y < (m_lSurfPxHeight - m_lMapPxHeight) Then ptTopLeft.Y = (m_lSurfPxHeight - m_lMapPxHeight)
        End If
    Else
        ptTopLeft.Y = 0 '// Reinitialize if needed
    End If
End Sub
