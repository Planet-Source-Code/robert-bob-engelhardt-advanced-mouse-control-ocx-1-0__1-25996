VERSION 5.00
Object = "*\AAdvancedMouse.vbp"
Begin VB.Form frmMain 
   Caption         =   "Advance Mouse Control"
   ClientHeight    =   5805
   ClientLeft      =   10275
   ClientTop       =   3015
   ClientWidth     =   5115
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   387
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   341
   Begin VB.Frame fraInput 
      Caption         =   "Input"
      Height          =   735
      Left            =   0
      TabIndex        =   26
      Top             =   3360
      Width           =   1575
      Begin VB.CommandButton btnBlockInput 
         Caption         =   "Block Input"
         Height          =   375
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame fraSwapedButtons 
      Caption         =   "Swap Buttons"
      Height          =   735
      Left            =   3480
      TabIndex        =   24
      Top             =   3360
      Width           =   1575
      Begin VB.CommandButton btnSwap 
         Caption         =   "Swap"
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame fraShowCursor 
      Caption         =   "Show/Hide Cursor"
      Height          =   735
      Left            =   1740
      TabIndex        =   22
      Top             =   3360
      Width           =   1575
      Begin VB.CommandButton btnShow 
         Caption         =   "Hide"
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame fraMetrics 
      Caption         =   "System Information"
      Height          =   855
      Left            =   0
      TabIndex        =   15
      Top             =   4200
      Width           =   5055
      Begin VB.Label lblDoubleClickSize 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Double Click Size"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2760
         TabIndex        =   19
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label lblCursorSize 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Cursor Size"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2760
         TabIndex        =   18
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label lblMouseButton 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Mouse Buttons"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label lblMouseExist 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Mouse Existance"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame fraCursor 
      Caption         =   "Cursor"
      Height          =   735
      Left            =   0
      TabIndex        =   11
      Top             =   1680
      Width           =   5055
      Begin VB.CommandButton btnCursorReturn 
         Caption         =   "Return Cursor"
         Height          =   375
         Left            =   3600
         TabIndex        =   21
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton btnCustomCursor 
         Caption         =   "Custom Cursor"
         Height          =   375
         Left            =   1560
         TabIndex        =   14
         Top             =   240
         Width           =   1335
      End
      Begin VB.PictureBox picCustomCursor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   3000
         Picture         =   "Form1.frx":0442
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   13
         Top             =   160
         Width           =   495
      End
      Begin VB.CommandButton btnChangeCursor 
         Caption         =   "Change Cursor"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1335
      End
   End
   Begin AdvancedMouse.MouseControl MouseControl1 
      Left            =   120
      Top             =   5160
      _ExtentX        =   953
      _ExtentY        =   953
   End
   Begin VB.Frame fraDoubleClick 
      Caption         =   "Double Click"
      Height          =   735
      Left            =   0
      TabIndex        =   4
      Top             =   2520
      Width           =   5055
      Begin VB.CommandButton btnDBLClick 
         Caption         =   "Double Click"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   3375
      End
      Begin VB.PictureBox picDoubleClickPicture 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   4560
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   5
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblTestDoubleClick 
         Caption         =   "Test Here:"
         Height          =   255
         Left            =   3600
         TabIndex        =   6
         Top             =   280
         Width           =   855
      End
   End
   Begin VB.Frame fraCaptureMouse 
      Caption         =   "Capture"
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5055
      Begin VB.CommandButton btnRelease 
         Caption         =   "Release Mouse"
         Height          =   375
         Left            =   3480
         TabIndex        =   3
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton btnCaptureForm 
         Caption         =   "Caputre on Form"
         Height          =   375
         Left            =   1800
         TabIndex        =   2
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton btnCaptureRect 
         Caption         =   "Capture in Rec"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame fraMousePos 
      Caption         =   "Mouse Position"
      Height          =   735
      Left            =   0
      TabIndex        =   8
      Top             =   840
      Width           =   5055
      Begin VB.CommandButton btnMouseTracking 
         Caption         =   "Enable Tracking"
         Height          =   375
         Left            =   1920
         TabIndex        =   20
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton btnChangePosition 
         Caption         =   "Change Position"
         Height          =   375
         Left            =   3480
         TabIndex        =   10
         Top             =   240
         Width           =   1455
      End
      Begin VB.Timer PositionTimer 
         Enabled         =   0   'False
         Interval        =   150
         Left            =   1440
         Top             =   240
      End
      Begin VB.Label lblPosition 
         Caption         =   "Position"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Label lblStaticInfo 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Advanced Mouse Control Demo Created By Robert Engelhardt"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   28
      Top             =   5160
      Width           =   5055
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Created By Robert Engelhardt on 8/8/01

Option Explicit 'all variable are defined

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) 'API used in demo not control

Private Sub btnCaptureForm_Click()
   MouseControl1.CaptureForm Me.hWnd 'capture on form
End Sub

Private Sub btnCaptureRect_Click()
   MouseControl1.CaptureRec 100, 100, 200, 200 'capture on screen in in box of (100,100)-(200,200)
End Sub

Private Sub btnChangeCursor_Click()
   MouseControl1.Cursor = Int(Rnd * 13) + 2 'cheange to randome cursor (don't change to arrow or default)
End Sub

Private Sub btnChangePosition_Click()
   MouseControl1.mouseX = Int(Rnd * ScaleX(Screen.Width, vbTwips, vbPixels)) 'new random x position
   MouseControl1.mouseY = Int(Rnd * ScaleY(Screen.Height, vbTwips, vbPixels)) 'can't froget the y position
End Sub

Private Sub btnCursorReturn_Click()
   MouseControl1.Cursor = vbArrow 'return the cursor to the arrow
End Sub

Private Sub btnCustomCursor_Click()
   MouseControl1.SetCustomCursor picCustomCursor.Picture 'load image from picture box
   MouseControl1.Cursor = vbCustom 'set custom cursor to be displayed
End Sub

Private Sub btnDBLClick_Click()
      
   Dim result As Variant
   result = InputBox("New Double click speed.", "Enter new value", MouseControl1.DoubleClickSpeed) 'get new speed
   
   If IsNumeric(result) Then 'be sure that new speed is a number
      MouseControl1.DoubleClickSpeed = Val(result) 'set
      btnDBLClick.Caption = "Change Double Click Speed from " & MouseControl1.DoubleClickSpeed 'rename caption
   End If
      
End Sub

Private Sub btnMouseTracking_Click()
   
   If PositionTimer.Enabled Then 'if timer is on make off, if off make on
      PositionTimer.Enabled = False
      btnMouseTracking.Caption = "Enable Tracking" 'change caption
   Else
      PositionTimer.Enabled = True
      btnMouseTracking.Caption = "Disable Tracking"
   End If
      
End Sub

Private Sub btnRelease_Click()
   MouseControl1.ReleaseCaptive 'release captive cursor
End Sub

Private Sub btnBlockInput_Click()
   MouseControl1.BlockAllInput True 'block user input (mouse and cursor)
   Sleep 1000 'pause for a second
   MouseControl1.BlockAllInput False 'return usage to user
End Sub

Private Sub btnShow_Click()
   MouseControl1.ShowMouse False 'hide the mouse
   Sleep 1000 'pause for a second
   MouseControl1.ShowMouse True 'show the mouse
End Sub

Private Sub btnSwap_Click()
   MouseControl1.SwapButtons = Not MouseControl1.SwapButtons 'swap buttons to be what they are not
End Sub

Private Sub Form_Load()
   Randomize Timer 'seed random number generator
   btnDBLClick.Caption = "Change Double Click Speed from " & MouseControl1.DoubleClickSpeed 'set captions
   Call FillMetrics 'call to the filling of some info
End Sub

Private Sub picDoubleClickPicture_DblClick()
   picDoubleClickPicture.BackColor = IIf(picDoubleClickPicture.BackColor = vbBlack, vbWhite, vbBlack) 'swap color at double click
End Sub

Private Sub PositionTimer_Timer()
   lblPosition.Caption = "position (" & MouseControl1.mouseX & "," & MouseControl1.mouseY & ")" 'display info
End Sub

Private Sub FillMetrics()
   'set all relevent text to apropriate text boxes
   lblMouseExist.Caption = "Mouse Present " & IIf(MouseControl1.MouseExists, "True", "False")
   lblMouseButton.Caption = "Mouse Buttons " & MouseControl1.ButtonCount
   lblCursorSize.Caption = "Cursor size " & MouseControl1.CursorWidth & " x " & MouseControl1.CursorHeight
   lblDoubleClickSize.Caption = "Double Click size " & MouseControl1.DoubleClickWidth & " x " & MouseControl1.DoubleClickHeight
End Sub
