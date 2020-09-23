VERSION 5.00
Begin VB.Form frmNewMap 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Create map"
   ClientHeight    =   2235
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   3585
   ControlBox      =   0   'False
   Icon            =   "frmNewMap.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   3585
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmbBrowseColor 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3045
      TabIndex        =   10
      Top             =   1050
      Width           =   330
   End
   Begin VB.TextBox txtBackColor 
      Height          =   285
      Left            =   1785
      MaxLength       =   6
      TabIndex        =   9
      Text            =   "000000"
      Top             =   1050
      Width           =   750
   End
   Begin VB.TextBox txtSize 
      Height          =   285
      Index           =   1
      Left            =   2730
      MaxLength       =   3
      TabIndex        =   6
      Text            =   "40"
      Top             =   630
      Width           =   645
   End
   Begin VB.TextBox txtSize 
      Height          =   285
      Index           =   0
      Left            =   1785
      MaxLength       =   3
      TabIndex        =   5
      Text            =   "20"
      Top             =   630
      Width           =   645
   End
   Begin VB.ComboBox cmbTileWidth 
      Height          =   315
      Left            =   1785
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   210
      Width           =   1170
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   330
      Left            =   1260
      TabIndex        =   1
      Top             =   1785
      Width           =   1065
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   330
      Left            =   2415
      TabIndex        =   0
      Top             =   1785
      Width           =   1065
   End
   Begin VB.Shape shpBackColor 
      BackStyle       =   1  'Opaque
      Height          =   330
      Left            =   2625
      Top             =   1050
      Width           =   330
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "Map backcolor :"
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
      Index           =   3
      Left            =   105
      TabIndex        =   8
      Top             =   1050
      UseMnemonic     =   0   'False
      Width           =   1395
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "x"
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
      Index           =   1
      Left            =   2520
      TabIndex        =   7
      Top             =   690
      Width           =   105
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "Tile size :"
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
      Index           =   2
      Left            =   105
      TabIndex        =   3
      Top             =   210
      UseMnemonic     =   0   'False
      Width           =   840
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "Map size (in tiles) :"
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
      Index           =   0
      Left            =   105
      TabIndex        =   2
      Top             =   630
      UseMnemonic     =   0   'False
      Width           =   1620
   End
End
Attribute VB_Name = "frmNewMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private objDuplicateMap As clsIsoMap

Private Sub cmbBrowseColor_Click()
    MsgBox "Implement Color Browser Common Dialog"
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    Dim sMessage$
    Dim lTileSize&
    
    '// Validation code
On Local Error Resume Next
    If Not (IsNumeric(txtSize(0).Text) And IsNumeric(txtSize(1).Text)) Then
        sMessage = "Please enter numeric values."
    ElseIf (txtSize(0).Text < 10) Or (txtSize(1).Text < 10) Or _
           (txtSize(0).Text > 999) Or (txtSize(1).Text > 999) Then
        sMessage = "Please enter a value between 10 and 999 for the map size."
    ElseIf IsError(CLng("&h00" & txtBackColor.Text)) Then
        sMessage = "Please enter a valid hexadecimal background value."
    End If

    If Len(sMessage) Then
        MsgBox sMessage, vbOKOnly + vbExclamation, "Validation error."
    Else
        '// Prepare data
        lTileSize = cmbTileWidth.ItemData(cmbTileWidth.ListIndex)
        frmMap.picTileRef.Picture = LoadPicture(GetPath(PATH_DATA_RES) & "tile_ref" & CStr(lTileSize) & ".bmp")
        objDuplicateMap.lTileSize = lTileSize
        
        objDuplicateMap.lBackColor = CLng("&h00" & txtBackColor.Text)
        objDuplicateMap.Map_SetSize txtSize(0).Text, txtSize(1).Text
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    
    SetNumberBox txtSize(0), True
    SetNumberBox txtSize(1), True
    
    cmbTileWidth.AddItem "32 * 16 px"
    cmbTileWidth.ItemData(cmbTileWidth.NewIndex) = 32
    cmbTileWidth.AddItem "64 * 32 px"
    cmbTileWidth.ItemData(cmbTileWidth.NewIndex) = 64
    cmbTileWidth.AddItem "128 * 64 px"
    cmbTileWidth.ItemData(cmbTileWidth.NewIndex) = 128
    cmbTileWidth.AddItem "256 * 128 px"
    cmbTileWidth.ItemData(cmbTileWidth.NewIndex) = 256

    cmbTileWidth.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objDuplicateMap = Nothing
End Sub

Public Sub SetRef(objMap As clsIsoMap)
    Set objDuplicateMap = objMap
End Sub

Private Sub txtBackColor_Change()
On Local Error Resume Next
    shpBackColor.BackColor = CLng("&h00" & txtBackColor.Text)
End Sub
