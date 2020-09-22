VERSION 5.00
Begin VB.Form fMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ucGIFViewer test"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6480
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   432
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdNextFrame 
      Caption         =   "Next frame"
      Enabled         =   0   'False
      Height          =   405
      Left            =   4215
      TabIndex        =   5
      Top             =   2610
      Width           =   2100
   End
   Begin VB.CheckBox chkAutoPlay 
      Caption         =   "AutoPlay"
      Height          =   255
      Left            =   4215
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1140
      Width           =   1410
   End
   Begin VB.CommandButton cmdLoadFromFile 
      Caption         =   "Load from file test"
      Height          =   405
      Left            =   4215
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   600
      Width           =   2100
   End
   Begin VB.CommandButton cmdLoadFromResource 
      Caption         =   "Load from resource test"
      Height          =   405
      Left            =   4215
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   135
      Width           =   2100
   End
   Begin VB.CommandButton cmdPause 
      Caption         =   "Pause"
      Enabled         =   0   'False
      Height          =   405
      Left            =   4215
      TabIndex        =   4
      Top             =   2145
      Width           =   2100
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Play"
      Enabled         =   0   'False
      Height          =   405
      Left            =   4215
      TabIndex        =   3
      Top             =   1680
      Width           =   2100
   End
   Begin GIFViewer.ucGIFViewer ucGIFViewer 
      Height          =   2880
      Left            =   135
      Top             =   135
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   5080
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub cmdLoadFromResource_Click()

    '-- Loading GIF from resource data
    Screen.MousePointer = vbArrowHourglass
    ucGIFViewer.LoadFromResource "GIF_BOUNCINGTHING", "GIF"
    Screen.MousePointer = vbDefault
    
    pvUpdateButtonsState
End Sub

Private Sub cmdLoadFromFile_Click()

    '-- Loading GIF from file
    Screen.MousePointer = vbArrowHourglass
    ucGIFViewer.LoadFromFile App.Path & "\animated.gif"
    Screen.MousePointer = vbDefault
    
    pvUpdateButtonsState
End Sub

Private Sub chkAutoPlay_Click()
    ucGIFViewer.AutoPlay = -chkAutoPlay
End Sub

Private Sub cmdPlay_Click()
    ucGIFViewer.Play
    pvUpdateButtonsState
End Sub

Private Sub cmdPause_Click()
    ucGIFViewer.Pause
    pvUpdateButtonsState
End Sub

Private Sub cmdNextFrame_Click()
    ucGIFViewer.NextFrame
    pvUpdateButtonsState
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '-- Destroy loaded GIF
    ucGIFViewer.Destroy
    '-- Free form
    Set fMain = Nothing
End Sub

'//

Private Sub pvUpdateButtonsState()
    With ucGIFViewer
        cmdPlay.Enabled = (.GIFLoaded And Not .GIFIsPlaying)
        cmdPause.Enabled = (.GIFLoaded And .GIFIsPlaying)
        cmdNextFrame.Enabled = (.GIFLoaded And Not .GIFIsPlaying)
    End With
End Sub


