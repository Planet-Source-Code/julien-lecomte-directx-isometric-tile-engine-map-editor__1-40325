Attribute VB_Name = "modMain"
Option Explicit

Private Const GWL_STYLE = (-16)

'-----------------------------------------------------------------------------
'// KERNEL32
'-----------------------------------------------------------------------------
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'-------------------------------------------------------------------------------------------
' SHELL32.DLL
'-------------------------------------------------------------------------------------------
Public Declare Function ShellExecuteA Lib "shell32" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'-------------------------------------------------------------------------------------------
' USER.DLL
'-------------------------------------------------------------------------------------------
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function GetWindowLongA Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLongA Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SendMessageA Lib "user32" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Sub Main()
    '// Load the main form
    Screen.MousePointer = vbHourglass
    Load frmMap
    frmMap.Show
'    frmMap.WindowState = vbMaximized
    Screen.MousePointer = vbNormal
End Sub

Public Function IsCompiled() As Boolean
    Static bIsCompiled As Boolean, bHasRun As Boolean

On Local Error Resume Next
    If Not bHasRun Then
        Debug.Print 1 / 0
        bIsCompiled = CBool(Err.Number = 0)
        bHasRun = True
    End If
    
    IsCompiled = bIsCompiled
End Function

Public Sub SetNumberBox(objTextBox As TextBox, bFlag As Boolean)
    Dim lCurStyle&
    Const ES_NUMBER = &H2000&
    
    ' retrieve the window style
    lCurStyle = GetWindowLongA(objTextBox.hwnd, GWL_STYLE)
    If bFlag Then
       lCurStyle = lCurStyle Or ES_NUMBER
    Else
       lCurStyle = lCurStyle And Not ES_NUMBER
    End If
    SetWindowLongA objTextBox.hwnd, GWL_STYLE, lCurStyle
    objTextBox.Refresh
End Sub
