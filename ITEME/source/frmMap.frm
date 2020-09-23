VERSION 5.00
Object = "{AB4F6C60-4898-11D2-9692-204C4F4F5020}#29.0#0"; "Ccrpsld.ocx"
Begin VB.Form frmMap 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Isometric Tile Engine Map Editor"
   ClientHeight    =   6150
   ClientLeft      =   150
   ClientTop       =   750
   ClientWidth     =   9975
   HasDC           =   0   'False
   Icon            =   "frmMap.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   9975
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picTileRef 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   0
      Picture         =   "frmMap.frx":000C
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Frame fraFrame 
      Caption         =   "Vertice color range"
      Height          =   1065
      Left            =   8085
      TabIndex        =   7
      Top             =   105
      Width           =   1800
      Begin CCRSlider.ccrpSlider ccrpSliderVertice 
         Height          =   330
         Index           =   0
         Left            =   105
         TabIndex        =   8
         Top             =   630
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         AutoColor       =   0   'False
         BackColor       =   -2147483633
         LargeChange     =   2
         Max             =   32
         Min             =   0
         MousePointer    =   0
         TickFrequency   =   8
         ThumbLength     =   5
      End
      Begin CCRSlider.ccrpSlider ccrpSliderVertice 
         Height          =   330
         Index           =   1
         Left            =   105
         TabIndex        =   9
         Top             =   315
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         AutoColor       =   0   'False
         BackColor       =   -2147483633
         LargeChange     =   2
         Max             =   32
         Min             =   0
         MousePointer    =   0
         TickStyle       =   1
         TickFrequency   =   8
         ThumbLength     =   5
         Value           =   32
      End
   End
   Begin VB.Frame fraTextures 
      Caption         =   "Textures"
      Height          =   4635
      Left            =   5985
      TabIndex        =   4
      Top             =   1260
      Width           =   2010
      Begin VB.VScrollBar vScroll 
         Height          =   4215
         LargeChange     =   3
         Left            =   1680
         TabIndex        =   5
         Top             =   315
         Width           =   225
      End
      Begin VB.FileListBox fleTextureList 
         Height          =   285
         Left            =   105
         Pattern         =   "*.bmp"
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   4200
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Shape shpSelected 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         Height          =   480
         Left            =   105
         Top             =   210
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image imgTexture 
         Height          =   480
         Index           =   0
         Left            =   105
         Stretch         =   -1  'True
         Top             =   315
         Width           =   480
      End
      Begin VB.Image imgTexture 
         Height          =   480
         Index           =   1
         Left            =   630
         Stretch         =   -1  'True
         Top             =   315
         Width           =   480
      End
      Begin VB.Image imgTexture 
         Height          =   480
         Index           =   2
         Left            =   1155
         Stretch         =   -1  'True
         Top             =   315
         Width           =   480
      End
      Begin VB.Image imgTexture 
         Height          =   480
         Index           =   3
         Left            =   105
         Stretch         =   -1  'True
         Top             =   840
         Width           =   480
      End
      Begin VB.Image imgTexture 
         Height          =   480
         Index           =   4
         Left            =   630
         Stretch         =   -1  'True
         Top             =   840
         Width           =   480
      End
      Begin VB.Image imgTexture 
         Height          =   480
         Index           =   5
         Left            =   1155
         Stretch         =   -1  'True
         Top             =   840
         Width           =   480
      End
      Begin VB.Image imgTexture 
         Height          =   480
         Index           =   6
         Left            =   105
         Stretch         =   -1  'True
         Top             =   1365
         Width           =   480
      End
      Begin VB.Image imgTexture 
         Height          =   480
         Index           =   7
         Left            =   630
         Stretch         =   -1  'True
         Top             =   1365
         Width           =   480
      End
      Begin VB.Image imgTexture 
         Height          =   480
         Index           =   8
         Left            =   1155
         Stretch         =   -1  'True
         Top             =   1365
         Width           =   480
      End
      Begin VB.Image imgTexture 
         Height          =   480
         Index           =   9
         Left            =   105
         Stretch         =   -1  'True
         Top             =   1890
         Width           =   480
      End
      Begin VB.Image imgTexture 
         Height          =   480
         Index           =   10
         Left            =   630
         Stretch         =   -1  'True
         Top             =   1890
         Width           =   480
      End
      Begin VB.Image imgTexture 
         Height          =   480
         Index           =   11
         Left            =   1155
         Stretch         =   -1  'True
         Top             =   1890
         Width           =   480
      End
      Begin VB.Image imgTexture 
         Height          =   480
         Index           =   12
         Left            =   105
         Stretch         =   -1  'True
         Top             =   2415
         Width           =   480
      End
      Begin VB.Image imgTexture 
         Height          =   480
         Index           =   13
         Left            =   630
         Stretch         =   -1  'True
         Top             =   2415
         Width           =   480
      End
      Begin VB.Image imgTexture 
         Height          =   480
         Index           =   14
         Left            =   1155
         Stretch         =   -1  'True
         Top             =   2415
         Width           =   480
      End
      Begin VB.Image imgTexture 
         Height          =   480
         Index           =   15
         Left            =   105
         Stretch         =   -1  'True
         Top             =   2940
         Width           =   480
      End
      Begin VB.Image imgTexture 
         Height          =   480
         Index           =   16
         Left            =   630
         Stretch         =   -1  'True
         Top             =   2940
         Width           =   480
      End
      Begin VB.Image imgTexture 
         Height          =   480
         Index           =   17
         Left            =   1155
         Stretch         =   -1  'True
         Top             =   2940
         Width           =   480
      End
      Begin VB.Image imgTexture 
         Height          =   480
         Index           =   18
         Left            =   105
         Stretch         =   -1  'True
         Top             =   3465
         Width           =   480
      End
      Begin VB.Image imgTexture 
         Height          =   480
         Index           =   19
         Left            =   630
         Stretch         =   -1  'True
         Top             =   3465
         Width           =   480
      End
      Begin VB.Image imgTexture 
         Height          =   480
         Index           =   20
         Left            =   1155
         Stretch         =   -1  'True
         Top             =   3465
         Width           =   480
      End
      Begin VB.Image imgTexture 
         Height          =   480
         Index           =   21
         Left            =   105
         Stretch         =   -1  'True
         Top             =   3990
         Width           =   480
      End
      Begin VB.Image imgTexture 
         Height          =   480
         Index           =   22
         Left            =   630
         Stretch         =   -1  'True
         Top             =   3990
         Width           =   480
      End
      Begin VB.Image imgTexture 
         Height          =   480
         Index           =   23
         Left            =   1155
         Stretch         =   -1  'True
         Top             =   3990
         Width           =   480
      End
   End
   Begin VB.PictureBox picDirectX 
      BorderStyle     =   0  'None
      Height          =   5895
      Left            =   105
      ScaleHeight     =   5895
      ScaleWidth      =   5685
      TabIndex        =   0
      Top             =   105
      Width           =   5685
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Shift-Click to make area higher, Shift-RClick to lower area. F2-F3 for wire/solid mode and F4-F5 for antialias mode."
      ForeColor       =   &H80000017&
      Height          =   1590
      Left            =   8085
      TabIndex        =   10
      Top             =   1365
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   435
      Left            =   5985
      Top             =   735
      Width           =   645
   End
   Begin VB.Label lblFps 
      BackStyle       =   0  'Transparent
      Caption         =   "### fps"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5985
      TabIndex        =   2
      Top             =   420
      Width           =   1170
   End
   Begin VB.Label lblLocation 
      BackStyle       =   0  'Transparent
      Caption         =   "X = #; Y = #"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5985
      TabIndex        =   1
      Top             =   105
      Width           =   1665
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileTiret 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileQuit 
         Caption         =   "&Quit"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuToolsZoom 
         Caption         =   "&Zoom"
         Begin VB.Menu mnuToolsZoomSub 
            Caption         =   "Zoom out"
            Index           =   0
         End
         Begin VB.Menu mnuToolsZoomSub 
            Caption         =   "Zoom in"
            Index           =   1
         End
         Begin VB.Menu mnuToolsZoomSub 
            Caption         =   "-"
            Index           =   2
         End
         Begin VB.Menu mnuToolsZoomSub 
            Caption         =   "Zoom 100 %"
            Index           =   3
         End
      End
      Begin VB.Menu mnuToolsMode 
         Caption         =   "3D &Mode"
         Begin VB.Menu mnuToolsModeSub 
            Caption         =   "&Wireframe"
            Index           =   2
            Shortcut        =   {F2}
         End
         Begin VB.Menu mnuToolsModeSub 
            Caption         =   "&Solid"
            Index           =   3
            Shortcut        =   {F3}
         End
      End
      Begin VB.Menu mnuToolsRender 
         Caption         =   "Render mode"
         Begin VB.Menu mnuToolsRenderOff 
            Caption         =   "Antialias off"
            Index           =   0
            Shortcut        =   {F4}
         End
         Begin VB.Menu mnuToolsRenderOff 
            Caption         =   "Antialias on"
            Index           =   1
            Shortcut        =   {F5}
         End
      End
   End
End
Attribute VB_Name = "frmMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
Option Explicit

Private objIsoMap      As clsIsoMap

Private m_bDirectXInit As Boolean '// Is DirectX running ?
Private m_lSelectedTexture&       '// Selected texture full index
Private m_sLookUpTextPath$        '// Lookup string

'-------------------------------------------------------------------------------------------
' USER32.DLL
'-------------------------------------------------------------------------------------------
Private Declare Function HideCaret Lib "user32" (ByVal hwnd As Long) As Long

Private Sub ccrpSliderVertice_Change(Index As Integer)
    If m_bDirectXInit Then
        objIsoMap.SetMode_VerticeColorRange ccrpSliderVertice(Index).Value, CBool(Index = 0)
    End If
End Sub



Private Sub Form_Activate()
    Textures_Act
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Const MAP_MOVE = 5
    Select Case KeyCode
        Case vbKeyLeft, vbKeyNumpad4:  objIsoMap.UpdateXY MAP_MOVE, 0
        Case vbKeyNumpad7:             objIsoMap.UpdateXY MAP_MOVE, MAP_MOVE
        Case vbKeyUp, vbKeyNumpad8:    objIsoMap.UpdateXY 0, MAP_MOVE
        Case vbKeyNumpad9:             objIsoMap.UpdateXY -MAP_MOVE, 0
        Case vbKeyRight, vbKeyNumpad6: objIsoMap.UpdateXY -MAP_MOVE, 0
        Case vbKeyNumpad3:             objIsoMap.UpdateXY -MAP_MOVE, -MAP_MOVE
        Case vbKeyDown, vbKeyNumpad2:  objIsoMap.UpdateXY 0, -MAP_MOVE
        Case vbKeyNumpad1:             objIsoMap.UpdateXY MAP_MOVE, -MAP_MOVE
    End Select
End Sub

Private Sub Form_Load()
    '// Textures ---------------------------------------------------------------------
    '// Lookup path
    m_sLookUpTextPath = GetPath(PATH_DATA_TEXTURES)
    '// Load all textures from path
    fleTextureList.Path = m_sLookUpTextPath
    '// Disabe/Enable scrollbar
    vScroll.Min = 0
    If fleTextureList.ListCount - imgTexture.UBound > 1 Then
        vScroll.Max = (fleTextureList.ListCount - imgTexture.UBound + 1) \ 3
    Else
        vScroll.Enabled = False
    End If
    '// Size scollbar
    vScroll.Top = imgTexture(imgTexture.LBound).Top
    vScroll.Height = imgTexture(imgTexture.UBound).Top + imgTexture(imgTexture.UBound).Height - vScroll.Top
    '// No tile selected by default
    m_lSelectedTexture = -1&

    'Set objIsoMap = New clsIsoMap
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    m_bDirectXInit = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    m_bDirectXInit = False
    Set objIsoMap = Nothing
End Sub

Private Sub imgTexture_Click(Index As Integer)
    shpSelected.Visible = True
    shpSelected.Tag = imgTexture(Index).Tag
    m_lSelectedTexture = Index + vScroll.Value * 3
    Texture_Select
End Sub

Private Sub mnuFileQuit_Click()
    Unload Me
End Sub

Private Sub mnuNew_Click()
    If m_bDirectXInit Then
        
    End If
    
    '// Clear and recreate a new map
    Set objIsoMap = Nothing
    Set objIsoMap = New clsIsoMap
    Load frmNewMap
    frmNewMap.SetRef objIsoMap
    frmNewMap.Show vbModal, Me
    If objIsoMap.IsValid Then
        RenderLoop
    Else
        '// User canceled
        Set objIsoMap = Nothing
    End If
    
End Sub

Private Sub mnuToolsModeSub_Click(Index As Integer)
    If m_bDirectXInit Then
        objIsoMap.SetMode_Fill CLng(Index)
    End If
End Sub

Private Sub mnuToolsRenderOff_Click(Index As Integer)
    objIsoMap.SetMode_Render Index
End Sub

Private Sub mnuToolsZoomSub_Click(Index As Integer)
    Dim lNewSize&
    If m_bDirectXInit Then
        If Index = 3 Then
            If objIsoMap.Map_Zoom(0&, lNewSize) Then
                picTileRef.Picture = LoadPicture(GetPath(PATH_DATA_RES) & "tile_ref" & CStr(lNewSize) & ".bmp")
            End If
        Else
            If objIsoMap.Map_Zoom(IIf(CBool(Index = 1), 1, -1), lNewSize) Then
                picTileRef.Picture = LoadPicture(GetPath(PATH_DATA_RES) & "tile_ref" & CStr(lNewSize) & ".bmp")
            End If
        End If
    End If
End Sub

'Private Sub mnuToolsOtherModesSub_Click(Index As Integer)
'    If m_bDirectXInit Then
'        Select Case Index
'            Case 0: objIsoMap.SwitchMode_Other BLACKEN_VERTICES
'            Case 0: objIsoMap.SwitchMode_Other WHITEN_VERTICES
'        End Select
'    End If
'End Sub

Private Sub picDirectX_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lX&, lY&
    Dim lHeight&
    Dim lTile&, sTileName$
    
    If m_bDirectXInit Then '// If directx is loaded
        If objIsoMap.Map_GetTileXY(X, Y, lX&, lY&, picTileRef.hdc) Then
            Select Case (Shift And (vbShiftMask Or vbCtrlMask))
                Case (vbShiftMask Or vbCtrlMask)
                    '// NOTHING TO HAPPEN
                Case vbCtrlMask
                    '// NOTHING TO HAPPEN
                Case vbShiftMask
                    '// RESERVED BY MOUSE UP
                Case 0& '// None of those pressed
                    If Button = vbLeftButton Then   '// ADD TEXTURE
                        sTileName = shpSelected.Tag
                        If Len(sTileName) Then
                            lTile = objIsoMap.DirectX_LoadSurfaceEx(m_sLookUpTextPath, sTileName)
                            '// Function above = success
                            If lTile <> -1& Then objIsoMap.Map_SetTileId lX, lY, lTile
                        End If
                    ElseIf Button = vbRightButton Then
                        '
                    End If
            End Select

        End If
    End If
End Sub

Private Sub picDirectX_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lX&, lY&
    
    If m_bDirectXInit Then '// If directX map is loaded
        '// If anything is pressed
        If (Button Or Shift) Then picDirectX_MouseDown Button, Shift, X, Y

        If objIsoMap.Map_GetTileXY(X, Y, lX&, lY&, picTileRef.hdc) Then
            lblLocation = "X = " & lX & "; Y = " & lY
        Else
            lblLocation = "X = #; Y = #"
        End If

    End If
End Sub

Private Sub picDirectX_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lX&, lY&
    Dim lHeight&
    Dim lTile&, sTileName$
    
    If m_bDirectXInit Then '// If directx is loaded
        If objIsoMap.Map_GetTileXY(X, Y, lX&, lY&, picTileRef.hdc) Then
            Select Case (Shift And (vbShiftMask Or vbCtrlMask))
                Case (vbShiftMask Or vbCtrlMask)
                    '// NOTHING TO HAPPEN
                Case vbCtrlMask
                    '// NOTHING TO HAPPEN
                Case vbShiftMask
                    If Button = vbLeftButton Then
                        lHeight = 1
                    ElseIf Button = vbRightButton Then
                        lHeight = -1
                    End If
                    objIsoMap.Map_SetTileHeight lX, lY, lHeight, 15&
                Case 0&
                     '// RESERVED BY MOUSE DOWN
            End Select

        End If
    End If
End Sub

Public Sub RenderLoop()
    Dim lRetVal&, lTileDefault&
    
On Local Error GoTo ErrHandler
    '// Set defaults
    m_bDirectXInit = objIsoMap.DirectX_Initialize(picDirectX)
    lTileDefault = objIsoMap.DirectX_LoadSurface(GetPath(PATH_DATA_TEXTURES), "default.bmp")
    objIsoMap.Map_SetSize 10&, 40&, lTileDefault

    If m_bDirectXInit Then
        Do '// ***Main Loop***
            DoEvents
            lRetVal = objIsoMap.DirectX_Refresh
            If lRetVal Then lblFps.Caption = lRetVal & " fps"

        Loop While m_bDirectXInit '// ***Main Loop***

        objIsoMap.DirectX_Destroy '// Clean stuff up
    End If
ErrHandler:
    m_bDirectXInit = False
End Sub

Private Sub Texture_Select()
    Dim lIndex&
    lIndex = m_lSelectedTexture - vScroll.Value * 3
    '// if texture is visible, then select it.
    If lIndex >= 0 And lIndex <= imgTexture.UBound Then
        shpSelected.Move imgTexture(lIndex).Left, imgTexture(lIndex).Top
        shpSelected.Visible = True
    Else
        shpSelected.Visible = False
    End If
End Sub

Private Sub Textures_Act() '// Wrapper
    Textures_Show
    Texture_Select
End Sub

Private Sub Textures_Show()
    Dim lIndex&
    Dim I&, lW&, lH&
    Dim sPath$, sFile$, sFilePath$

    '// Get primary data
    lIndex = vScroll.Value * 3
    sPath = GetPath(PATH_DATA_TEXTURES)

    '// Show all textures if they can be shown
    For I = imgTexture.LBound To imgTexture.UBound
        If (I + lIndex) <= fleTextureList.ListCount Then
            sFile = fleTextureList.List(I + lIndex)
            sFilePath = sPath & sFile
        Else
            sFile = vbNullString
            sFilePath = vbNullString
        End If

        If IsFile(sFilePath) Then
            imgTexture(I).Picture = LoadPicture(sFilePath)
            
            '// Convert from HIMETRIC to Pixels [Edanmo's website]
            lW = (imgTexture(I).Picture.Width * 1440&) / (2540& * Screen.TwipsPerPixelX)
            lH = (imgTexture(I).Picture.Height * 1440&) / (2540& * Screen.TwipsPerPixelY)

            imgTexture(I).ToolTipText = lW & "x" & lH & " [" & sFile & "]"
            imgTexture(I).Tag = sFile
            imgTexture(I).Visible = True
        Else
            imgTexture(I).Visible = False
        End If
    Next
End Sub

Private Sub vScroll_Change()
    Textures_Act
End Sub

Private Sub vScroll_GotFocus()
    '// Hide the ugly flicker of scrollbar
    HideCaret vScroll.hwnd
End Sub

Private Sub vScroll_Scroll()
    Textures_Act
End Sub
