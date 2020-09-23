VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Displaysettings"
   ClientHeight    =   6510
   ClientLeft      =   1650
   ClientTop       =   1545
   ClientWidth     =   5130
   LinkTopic       =   "Form1"
   ScaleHeight     =   6510
   ScaleWidth      =   5130
   Begin VB.CommandButton cmdCurrentDisplaySettings 
      Caption         =   "Get Current Display Settings"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   150
      TabIndex        =   4
      Top             =   5460
      Width           =   4905
   End
   Begin VB.CommandButton cmdChangeDisplaySettings 
      Caption         =   "Change Display Settings"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   150
      Style           =   1  'Grafisch
      TabIndex        =   2
      Top             =   90
      Width           =   4905
   End
   Begin VB.CommandButton cmdScreenDlg 
      Caption         =   "Show Screen Dialog"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   150
      TabIndex        =   1
      Top             =   5970
      Width           =   4905
   End
   Begin VB.ListBox lstDisplaySettings 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3435
      ItemData        =   "Form1.frx":0000
      Left            =   120
      List            =   "Form1.frx":0002
      TabIndex        =   0
      Top             =   990
      Width           =   4875
   End
   Begin VB.Label Label2 
      Caption         =   "Current Display Settings:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   150
      TabIndex        =   6
      Top             =   4620
      Width           =   2955
   End
   Begin VB.Label Label1 
      Caption         =   "Available Display Settings:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   150
      TabIndex        =   5
      Top             =   720
      Width           =   2355
   End
   Begin VB.Label lblCurrentDisplaySettings 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   150
      TabIndex        =   3
      Top             =   5010
      Width           =   4875
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Code by Michael Schriever
'EMail: webmaster@michael-schriever.de

Option Explicit

Private Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lpDevMode As Any) As Boolean
Private Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lpDevMode As Any, ByVal dwflags As Long) As Long

Private Const CCDEVICENAME = 32
Private Const CCFORMNAME = 32
Private Const DM_BITSPERPEL = &H40000
Private Const DM_PELSWIDTH = &H80000
Private Const DM_PELSHEIGHT = &H100000
Private Const CDS_UPDATEREGISTRY = &H1
Private Const CDS_TEST = &H4
Private Const DISP_CHANGE_SUCCESSFUL = 0
Private Const DISP_CHANGE_RESTART = 1
Private Const ENUM_CURRENT_SETTINGS = &HFFFF - 1

Private Type DEVMODE
    dmDeviceName      As String * 32
    dmSpecVersion     As Integer
    dmDriverVersion   As Integer
    dmSize            As Integer
    dmDriverExtra     As Integer
    dmFields          As Long
    dmOrientation     As Integer
    dmPaperSize       As Integer
    dmPaperLength     As Integer
    dmPaperWidth      As Integer
    dmScale           As Integer
    dmCopies          As Integer
    dmDefaultSource   As Integer
    dmPrintQuality    As Integer
    dmColor           As Integer
    dmDuplex          As Integer
    dmYResolution     As Integer
    dmTTOption        As Integer
    dmCollate         As Integer
    dmFormName        As String * 32
    dmUnusedPadding   As Integer
    dmBitsPerPel      As Integer
    dmPelsWidth       As Long
    dmPelsHeight      As Long
    dmDisplayFlags    As Long
    dmDisplayFrequency As Long
End Type

Dim DPS() As DEVMODE

Private Sub cmdChangeDisplaySettings_Click()
    Dim ret As Long
    Dim ret1 As Integer
    
    ret1 = MsgBox("Before you change the display settings make sure that your" + vbCr + _
                  "monitor supports this settings. Proceed in change settings ?", vbYesNo + vbQuestion)
    If ret1 = vbNo Then Exit Sub
    
    If lstDisplaySettings.ListIndex >= 0 Then
        ret = ChangeDisplaySettings(DPS(lstDisplaySettings.ListIndex), CDS_UPDATEREGISTRY)
        If ret <> DISP_CHANGE_SUCCESSFUL Then
            MsgBox "Display settings could not be changed !", vbInformation
        End If
        Call showCurrentDisplaySettings
    Else
        MsgBox "No items selected !"
    End If
End Sub

Private Sub cmdScreenDlg_Click()
    Shell "rundll32.exe shell32.dll, Control_RunDLL desk.cpl, ,3", 1
End Sub

Private Sub cmdCurrentDisplaySettings_Click()
    Call showCurrentDisplaySettings
End Sub

Private Sub Form_Load()
    Call showCurrentDisplaySettings
    Call showAvailableDisplaySettings
End Sub

Private Sub showCurrentDisplaySettings()
    Dim curDPS As DEVMODE
    Dim ret As Long
    Dim colors As String
    
    ret = EnumDisplaySettings(0&, ENUM_CURRENT_SETTINGS, curDPS)
    
    If ret = 0 Then
        MsgBox "Error evaluating the current display settings", vbInformation
    Else
        Select Case curDPS.dmBitsPerPel
            Case 4:      colors = "16 Color"
            Case 8:      colors = "256 Color"
            Case 16:     colors = "High Color"
            Case 24, 32: colors = "True Color"
        End Select
        lblCurrentDisplaySettings = Format(curDPS.dmPelsWidth, "@@@@") + " x " + _
                      Format(curDPS.dmPelsHeight, "@@@@") + "  " + _
                      Format(colors, "@@@@@@@@@@@@@  ") + _
                      Format(curDPS.dmDisplayFrequency, "@@@ Hz")
    End If
End Sub

Private Sub showAvailableDisplaySettings()
    Dim nr As Long
    Dim colors As String
    Dim ret As Long
    
    lstDisplaySettings.Clear
    ReDim DPS(0)
    DPS(0).dmSize = LenB(DPS(0))
    nr = 0
    ret = EnumDisplaySettings(0&, nr, DPS(nr))
    
    Do
        Select Case DPS(nr).dmBitsPerPel
            Case 4:      colors = "16 Color"
            Case 8:      colors = "256 Color"
            Case 16:     colors = "High Color"
            Case 24, 32: colors = "True Color"
        End Select
        lstDisplaySettings.AddItem Format(DPS(nr).dmPelsWidth, "@@@@") + " x " + _
                      Format(DPS(nr).dmPelsHeight, "@@@@") + "  " + _
                      Format(colors, "@@@@@@@@@@@@@  ") + _
                      Format(DPS(nr).dmDisplayFrequency, "@@@ Hz")
        nr = nr + 1
        ReDim Preserve DPS(nr)
        ret = EnumDisplaySettings(0&, nr, DPS(nr))
    Loop Until ret = 0
End Sub
