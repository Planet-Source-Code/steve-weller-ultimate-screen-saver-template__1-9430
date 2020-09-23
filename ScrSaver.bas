Attribute VB_Name = "modScrSaver"
' Screen Saver template
' by Steve Weller (6/27/2000)
' Use this as a template for screen savers
' Just put in Templates/Projects directory
' and start creating!
' Note: in the make EXE dialog, add .scr
' to the end of the program name, and put
' the screen saver in the Windows directory
Option Explicit

' For showing/hiding mouse cursor
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Private CursorDepth As Long

' Disable/Enable Ctrl+Alt+Del, Alt+Tab, etc.
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Const SPI_SCREENSAVERRUNNING = 97
Public Const SPI_SETSCREENSAVEACTIVE = 17

' For passwords (undocumented)
Private Declare Sub PwdChangePassword Lib "mpr" Alias "PwdChangePasswordA" (ByVal lpcRegKeyName As String, ByVal hWnd As Long, ByVal uReserved1 As Long, ByVal uReserved2 As Long)
Public Declare Function VerifyScreenSavePwd Lib "password.cpl" (ByVal hWnd As Long) As Long

' For preview mode
Public FullScreen As Boolean
Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type
' Functions for preview mode
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
' SetWindowLong constants
Private Const WS_CHILD = &H40000000
Private Const GWL_HWNDPARENT = (-8)
Private Const GWL_STYLE = (-16)
Private Const HWND_TOP = 0&
Private Const SWP_NOZORDER = &H4
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_SHOWWINDOW = &H40

' Is the OS Windows NT?
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" _
  (lpVersionInformation As OSVERSIONINFO) As Long
Private Const VER_PLATFORM_WIN32s = 0
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2
Private Type OSVERSIONINFO
  OSVSize As Long
  dwVerMajor As Long
  dwVerMinor As Long
  dwBuildNumber As Long
  PlatformID As Long
  szCSDVersion As String * 128
End Type
Public Function IsWinNT() As Boolean
  Dim OSV As OSVERSIONINFO
      
  OSV.OSVSize = Len(OSV)
   
  If GetVersionEx(OSV) = 1 Then
    ' PlatformId contains a value representing
    ' the OS, so if its VER_PLATFORM_WIN32_NT,
    ' return true
    IsWinNT = (OSV.PlatformID = VER_PLATFORM_WIN32_NT)
  End If
End Function

Public Sub CtrlAltDel(ByVal Enable As Boolean)
' Enables or disables Ctrl+Alt+Del, Alt+Tab, etc.
SystemParametersInfo SPI_SCREENSAVERRUNNING, Not (Enable), ByVal 0&, 0&
End Sub


Public Sub Cursor(ByVal CursorOn As Boolean)
' Shows or hides cursor depending on CursorOn
Dim CurrentCursorDepth As Integer

If CursorOn Then
  CurrentCursorDepth = ShowCursor(True)
  
  Do While CurrentCursorDepth < CursorDepth
    CurrentCursorDepth = ShowCursor(True)
  Loop
Else
  CurrentCursorDepth = ShowCursor(False)
  
  Do While CurrentCursorDepth > -1
    CurrentCursorDepth = ShowCursor(False)
  Loop
End If
End Sub


Public Sub Main()

' Check command-line parameter
Select Case UCase$(Left$(Command$, 2))
  Case "/P"
    ' Show preview on square in
    ' screen saver dialog box
    Dim WinStyle As Long, PreviewBoxhWnd As Long
    Dim FrmhWnd As Long, Rct As RECT
    
    ' For preview mode
    FullScreen = False
    
    ' Get preview window hWnd and size
    PreviewBoxhWnd = Val(Mid$(Command$, 4))
    GetClientRect PreviewBoxhWnd, Rct
    
    Load frmScrSave
    FrmhWnd = frmScrSave.hWnd
    
    ' Adds child style to form
    WinStyle = GetWindowLong(FrmhWnd, GWL_STYLE)
    WinStyle = WinStyle Or WS_CHILD
    SetWindowLong FrmhWnd, GWL_STYLE, WinStyle
    
    ' Parent window set
    SetParent FrmhWnd, PreviewBoxhWnd
    ' Saves the handle
    SetWindowLong FrmhWnd, GWL_HWNDPARENT, PreviewBoxhWnd
    ' Changes form's dimension and position
    SetWindowPos FrmhWnd, _
     HWND_TOP, 0&, 0&, Rct.Right, Rct.Bottom, _
     SWP_NOZORDER Or SWP_NOACTIVATE Or SWP_SHOWWINDOW
  Case "/C"
    ' Show options
    frmScrSaveOptions.Show vbModal
  Case "/A"
    ' WinNT handles passwords on its own
    If Not (IsWinNT()) Then
      ' Show password dialog
      PwdChangePassword "SCRSAVE", CLng(Mid$(Command$, 4)), 0&, 0&
    End If
  Case "/S"
    ' Show screen saver
    Call Cursor(False)
    Call CtrlAltDel(False)
    SystemParametersInfo SPI_SETSCREENSAVEACTIVE, 0&, ByVal 0&, 0&
    FullScreen = True
    Load frmScrSave
    frmScrSave.Show
End Select
End Sub
