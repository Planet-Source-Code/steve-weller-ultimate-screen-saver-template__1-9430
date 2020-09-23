VERSION 5.00
Begin VB.Form frmScrSave 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   3900
   ClientLeft      =   1095
   ClientTop       =   2055
   ClientWidth     =   6585
   ControlBox      =   0   'False
   Icon            =   "ScrSaver.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   260
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   439
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrScrSave 
      Interval        =   50
      Left            =   960
      Top             =   2040
   End
End
Attribute VB_Name = "frmScrSave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub EndScrSaver()
' End the screen saver
Dim Res As Boolean

' WinNT handles passwords on its own
If Not (IsWinNT()) Then
  ' Check for password
  Res = VerifyScreenSavePwd(Me.hWnd)
  If Res = False Then Exit Sub
End If

' True if password is correct
' or if there is no password
Call Cursor(True)
Call CtrlAltDel(True)
SystemParametersInfo SPI_SETSCREENSAVEACTIVE, 1&, ByVal 0&, 0&
Unload Me
End
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
' end screen saver
If FullScreen Then Call EndScrSaver
End Sub


Private Sub Form_Load()
' Load settings from Registry, etc.
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' end screen saver
If FullScreen Then Call EndScrSaver
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Static OldX%, OldY%

' If this isn't first time
' and there is a big move
If (OldX > 0 And OldY > 0) And _
 (Abs(X - OldX) > 3 Or Abs(Y - OldY) > 3) Then
 If FullScreen Then Call EndScrSaver
End If
OldX = X
OldY = Y
End Sub

Private Sub tmrScrSave_Timer()
' Put code here for timer event
If FullScreen Then
  ' Code for full-screen mode

Else
  ' Code for preview mode

End If
End Sub
