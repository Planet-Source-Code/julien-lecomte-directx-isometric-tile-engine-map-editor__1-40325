Attribute VB_Name = "modFiles"
Option Explicit
Option Base 0
Option Compare Binary

Public Enum DATA_PATH
    PATH_APP                '// %APP%
    PATH_SYS                '// %SYS%
    PATH_DATA               '// %APP%\DATA
    PATH_DATA_RES           '// %APP%\DATA\RESSOURCES
    PATH_DATA_TEXTURES      '// %APP%\DATA\TEXTURES
    PATH_TEMP               '// %TEMP%
    FILE_TEMP               '// %TEMP%\[file.tmp]
    FILE_PROFILES           '// %APP%\DATA\[profiles.dat]
End Enum

'// File paths
Private Const FILE_PATH_DATA = "data\"
Private Const FILE_PATH_RES = "ressources\"
Private Const FILE_PATH_TEXTURES = "textures\"

'// File extensions
Private Const MAX_SIZE = 255

Private Const FILE_EXT_DOT = "."
Private Const FILE_EXT_BMP = "bmp"
Private Const FILE_EXT_MAP = "map"
Public Const FILE_BMP = FILE_EXT_DOT & FILE_EXT_BMP
Public Const FILE_MAP = FILE_EXT_DOT & FILE_EXT_MAP

'-------------------------------------------------------------------------------------------
' SHLWAPI.DLL
'-------------------------------------------------------------------------------------------
Private Declare Function PathIsDirectoryA Lib "shlwapi" (ByVal pszPath As String) As Long
Private Declare Function PathFileExistsA Lib "shlwapi" (ByVal pszPath As String) As Long

'-------------------------------------------------------------------------------------------
' KERNEL32.DLL
'-------------------------------------------------------------------------------------------
Private Declare Function GetTempFileNameA Lib "kernel32" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Private Declare Function GetTempPathA Lib "kernel32" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetSystemDirectoryA Lib "kernel32" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function LoadLibraryA Lib "kernel32" (ByVal lpLibfName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function CreateThread Lib "kernel32" (lpThreadAttributes As Any, ByVal dwStackSize As Long, ByVal lpStartAddress As Long, ByVal lpparameter As Long, ByVal dwCreationFlags As Long, lpThreadID As Long) As Long
Private Declare Function GetExitCodeThread Lib "kernel32" (ByVal hThread As Long, lpExitCode As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Sub ExitThread Lib "kernel32" (ByVal xc As Long)

Private Function AddBackSlash(sString As String) As String
    AddBackSlash = sString & IIf(Right$(sString, 1) = "\", vbNullString, "\")
End Function

Private Function DirCreate(sSource As String, Optional bRaiseErr As Boolean = True) As Boolean
    '// creates a folder, return true for success
    DirCreate = True
On Local Error GoTo ErrorHandler
    If Not (IsDir(sSource)) Then MkDir sSource
    Exit Function
ErrorHandler:
    If Not (bRaiseErr) Then Exit Function
    DirCreate = False
End Function

Public Function FileAddExtension(ByVal sFile$, sExt$) As String
    Dim sRet$
    '// Removes the extension from a filename
    sRet = sFile
    If Right$(UCase$(sRet), Len(FILE_EXT_DOT & sExt)) <> UCase$(FILE_EXT_DOT & sExt) Then
        sRet = sRet & FILE_EXT_DOT & sExt
    End If
    FileAddExtension = sRet
End Function

Private Function GetAppPath() As String
    GetAppPath = AddBackSlash(App.Path) & IIf(IsCompiled, vbNullString, "distribute\")
End Function

Public Function GetPath(ePath As DATA_PATH) As String
    '// Returns a path
    Dim sReturn$
    Dim bIsDir As Boolean
    Select Case ePath
        Case PATH_APP
            sReturn = GetAppPath
        Case PATH_SYS
            sReturn = GetSysPath
        Case PATH_DATA
            sReturn = GetAppPath & FILE_PATH_DATA
            bIsDir = True
        Case PATH_DATA_RES
            sReturn = GetPath(PATH_DATA) & FILE_PATH_RES
            bIsDir = True
        Case PATH_DATA_TEXTURES
            sReturn = GetPath(PATH_DATA) & FILE_PATH_TEXTURES
            bIsDir = True
        Case PATH_TEMP
            sReturn = GetPathTempDir
        Case FILE_TEMP
            sReturn = GetTemporaryFile
    End Select
    
    '// Create file if it doesn't exist
    If bIsDir Then
        If Not IsDir(sReturn) Then DirCreate sReturn
    End If
    
    GetPath = sReturn
End Function

Private Function GetPathTempDir() As String
    '// Returns Temporary directory
    Dim sBuffer$
    Dim lRetVal&

    sBuffer = String$(MAX_SIZE, vbNullChar)
    lRetVal = GetTempPathA(MAX_SIZE, sBuffer)
    GetPathTempDir = AddBackSlash(Left$(sBuffer, lRetVal - 1))
End Function

Private Function GetSysPath() As String
    Dim sBuffer$
    Dim lRetVal&

    sBuffer = String$(MAX_SIZE, vbNullChar)
    lRetVal = GetSystemDirectoryA(sBuffer, MAX_SIZE)
    GetSysPath = AddBackSlash(Left$(sBuffer, lRetVal))
End Function

Private Function GetTemporaryFile() As String
    '// Returns a temporary file
    Dim sPath$
    Dim sBuffer$
    
    sPath = GetPath(PATH_TEMP)
    
    sBuffer = String$(MAX_SIZE, vbNullChar)
    GetTempFileNameA sPath, "tmp", 0&, sBuffer '// Replace header here if needed
    sBuffer = Left$(sBuffer, InStr(1, sBuffer, vbNullChar) - 1)
     
    GetTemporaryFile = sBuffer
End Function

Public Function IsDir(pszPath As String) As Boolean
    '// Checks if is a directory
    If Len(pszPath) Then
        IsDir = CBool(PathIsDirectoryA(pszPath))
    End If
End Function

Public Function IsFile(ByVal pszPath As String) As Boolean
    '// Checks if is a file
    If Len(pszPath) Then
        If Not IsDir(pszPath) Then
            IsFile = CBool(PathFileExistsA(pszPath))
        End If
    End If
End Function

Public Function RegisterActiveX(FName$) As Boolean
    '//Returns true for success.
    Dim regLib&, process&, succeed&
    Dim h1&, xc&, id&
    Const p As String = "DllRegisterServer"

    regLib = LoadLibraryA(FName)
    If regLib Then
        process = GetProcAddress(regLib, p)
        If process Then
            h1 = CreateThread(ByVal 0&, 0&, ByVal process, ByVal 0&, 0&, id)
            If h1 Then
                succeed = (WaitForSingleObject(h1, 10000&) = 0)
                If succeed Then
                    CloseHandle h1
                    RegisterActiveX = True
                Else
                    GetExitCodeThread h1, xc
                    ExitThread xc
                End If
            End If
        End If
        FreeLibrary regLib
    End If
End Function
