VERSION 5.00
Begin VB.MDIForm mdiMain 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "Isometric Tile Engine Map Editor"
   ClientHeight    =   6915
   ClientLeft      =   165
   ClientTop       =   765
   ClientWidth     =   8085
   Icon            =   "mdiMain.frx":0000
   LinkTopic       =   "MDIForm1"
   NegotiateToolbars=   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As"
      End
      Begin VB.Menu mnuFileSeparator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileQuit 
         Caption         =   "&Quit"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuToolsMode 
         Caption         =   "&Mode"
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
   End
End
Attribute VB_Name = "mdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
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

Public Sub Enable_Commands()
    Dim lFormsOpen&
    Dim bEnable As Boolean

    lFormsOpen = DoEvents '// DoEvents returns the number of load forms
    '// frmTiles + mdiMain = 2 forms open
    '// 4 forms if frmMap is loaded because frmTools is also loaded
    bEnable = CBool(lFormsOpen > 2)

    mnuFileSave.Enabled = bEnable
    mnuFileSaveAs.Enabled = bEnable
    mnuToolsMode.Enabled = bEnable
End Sub

Private Sub MDIForm_Activate()
    Enable_Commands
End Sub

Private Sub MDIForm_Load()
'    '// frmTiles loading part
'    Load frmTiles
'    frmTiles.Move 0&, 0&
'    frmTiles.Show
End Sub

Private Sub mnuFileNew_Click()
    Load frmMap
    frmMap.Show
    frmMap.RenderLoop
    Enable_Commands
End Sub

Private Sub mnuFileQuit_Click()
    Unload Me
End Sub

'Private Sub mnuToolsModeSub_Click(Index As Integer)
'On Local Error Resume Next
'    frmMap.objIsoMap.DirectX_ModeFill CLng(Index)
'End Sub
Private Sub mnuFileNew_Click()

End Sub
