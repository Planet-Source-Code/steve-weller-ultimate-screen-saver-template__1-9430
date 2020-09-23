VERSION 5.00
Begin VB.Form frmScrSaveOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Screen Saver Options"
   ClientHeight    =   3195
   ClientLeft      =   1275
   ClientTop       =   2445
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About..."
      Height          =   375
      Left            =   4680
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmScrSaveOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hWnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long



Private Sub CancelButton_Click()
' unload Setup form
Unload Me
End Sub

Private Sub cmdAbout_Click()
' Show about box
' You can change it if you want
ShellAbout Me.hWnd, "Screen Saver", "", 0
End Sub


Private Sub OKButton_Click()
' put code here to save settings
' (write to Registry, INI file, etc.)
End Sub


Private Sub cmdCancel_Click()
' Unload Setup form
Unload Me
End Sub

Private Sub cmdOK_Click()
' Put code here to save settings
' (write to Registry, INI file, etc.)
End Sub


Private Sub Form_Load()
' Put code here to get settings
' from Registry, INI File, etc.
End Sub


