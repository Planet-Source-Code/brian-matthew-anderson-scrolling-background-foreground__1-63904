VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Dual scrolling backgrounds"
   ClientHeight    =   7830
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10320
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   522
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   688
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Text            =   "Press Escape key to exit."
      Top             =   960
      Width           =   4335
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   6
      TabStop         =   0   'False
      Text            =   "and the numpad to scroll the background."
      Top             =   600
      Width           =   5895
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Text            =   "Use the arrow keys to scroll the foreground"
      Top             =   240
      Width           =   5895
   End
   Begin VB.PictureBox StagingArea 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H000080FF&
      ForeColor       =   &H80000008&
      Height          =   3840
      Left            =   2880
      ScaleHeight     =   254
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   254
      TabIndex        =   4
      Top             =   5160
      Visible         =   0   'False
      Width           =   3840
   End
   Begin VB.PictureBox FoursquareMask 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7680
      Left            =   2040
      ScaleHeight     =   512
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   512
      TabIndex        =   3
      Top             =   6840
      Visible         =   0   'False
      Width           =   7680
   End
   Begin VB.PictureBox FourSquare 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7680
      Left            =   1440
      ScaleHeight     =   512
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   512
      TabIndex        =   2
      Top             =   6120
      Visible         =   0   'False
      Width           =   7680
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3840
      Left            =   6120
      Picture         =   "Scroll.frx":0000
      ScaleHeight     =   256
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   1
      Top             =   1320
      Visible         =   0   'False
      Width           =   3840
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3840
      Left            =   1800
      Picture         =   "Scroll.frx":30042
      ScaleHeight     =   256
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   0
      Top             =   1320
      Visible         =   0   'False
      Width           =   3840
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetTickCount& Lib "kernel32" ()

'A few small API calls is all we need for this example.
Private Declare Function BitBlt Lib "gdi32.dll" _
(ByVal hdcDest As Long, ByVal xDest As Long, ByVal yDest As Long, _
ByVal nWidth As Long, ByVal nHeight As Long, _
ByVal hdcSrc As Long, ByVal xSource As Long, ByVal ySource As Long, _
ByVal RasterOp As Long) As Long

Dim X1 As Long, Y1 As Long
Dim X2 As Long, Y2 As Long

Dim bRunning As Boolean

Private Sub Form_Load()
Dim OldTime As Long, CurrentTime As Long
Dim I As Long, J As Long
Dim SourceX1 As Long, SourceY1 As Long
Dim SourceX2 As Long, SourceY2 As Long

Me.Show
Me.Refresh
bRunning = True

SourceX2 = 50
SourceY2 = 100

'position the textboxes
Text1.Left = Me.ScaleWidth / 2 - (Text1.Width / 2)
Text2.Left = Text1.Left
Text3.Left = Me.ScaleWidth / 2 - (Text3.Width / 2)

'Seamless tiles are used here so the motion is continuous. The key here is to have a 2x2
'tiled area to pull graphics from. SourceX1, SourceY1, etc are really offsets in this
'2x2 grid that allows us to give the appearance of movement in center of the main form.
For I = 0 To 1
  For J = 0 To 1
    BitBlt FourSquare.hDC, I * 256, J * 256, 256, 256, Picture1.hDC, 0, 0, vbSrcCopy
    BitBlt FoursquareMask.hDC, I * 256, J * 256, 256, 256, Picture2.hDC, 0, 0, vbSrcCopy
  Next J
Next I

CurrentTime = GetTickCount

Do
  While bRunning = True
    If OldTime <> CurrentTime Then
      OldTime = CurrentTime
'No need to clear the staging area - just copy over the top of it and the last image
'will be erased. The staging area is essentially a backbuffer. Nothing is displayed
'yet...
      BitBlt StagingArea.hDC, 0, 0, 256, 256, FourSquare.hDC, SourceX1, SourceY1, vbSrcCopy
'There two blits are for the foreground. Always blit from the furthest "back" image
'to the "closest". You could easily add a third layer after these two if desired.
      BitBlt StagingArea.hDC, 0, 0, 256, 256, FoursquareMask.hDC, SourceX2, SourceY2, vbSrcAnd
      BitBlt StagingArea.hDC, 0, 0, 256, 256, FourSquare.hDC, SourceX2, SourceY2, vbSrcPaint
'One blit to copy the whole staging area to the screen. This allows for smooth
'scrolling. If you didn't use a backbuffer or staging area, this example would
'have a lot of flicker.
      BitBlt Me.hDC, (Me.ScaleWidth / 2) - (StagingArea.ScaleWidth / 2), _
      (Me.ScaleHeight / 2) - (StagingArea.ScaleHeight / 2), _
      256, 256, StagingArea.hDC, 0, 0, vbSrcCopy
'The following is used for bounds checking. We want our variables to stay within their limits.
'Try holding an arrow key down long enough and the scrolling will appear to reverse
'direction and even come to a stand-still. In real life you might see a helicopter's
'blades turning so fast that they appear to be reversing or even stopping. This is
'the same concept here.
      If SourceX1 + X1 < 0 Then
        SourceX1 = SourceX1 + 256
      ElseIf SourceX1 + X1 > 256 Then
        SourceX1 = SourceX1 - 256
      Else
        SourceX1 = SourceX1 + X1
      End If
      
      If SourceY1 + Y1 < 0 Then
        SourceY1 = SourceY1 + 256
      ElseIf SourceY1 + Y1 > 256 Then
        SourceY1 = SourceY1 - 256
      Else
        SourceY1 = SourceY1 + Y1
      End If
      
      If SourceX2 + X2 < 0 Then
        SourceX2 = SourceX2 + 256
      ElseIf SourceX2 + X2 > 256 Then
        SourceX2 = SourceX2 - 256
      Else
        SourceX2 = SourceX2 + X2
      End If

      If SourceY2 + Y2 < 0 Then
        SourceY2 = SourceY2 + 256
      ElseIf SourceY2 + Y2 > 256 Then
        SourceY2 = SourceY2 - 256
      Else
        SourceY2 = SourceY2 + Y2
      End If
      
    End If
    DoEvents
    CurrentTime = GetTickCount
  Wend
End
Loop




End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
'This section obviously changes the direction of the foreground
'or background depending on which key is pressed. "Q" or Escape exits.
Case vbKeyUp
  Y1 = Y1 - 1
Case vbKeyDown
  Y1 = Y1 + 1
Case vbKeyLeft
  X1 = X1 - 1
Case vbKeyRight
  X1 = X1 + 1
Case vbKeyNumpad8
  Y2 = Y2 - 1
Case vbKeyNumpad2
  Y2 = Y2 + 1
Case vbKeyNumpad4
  X2 = X2 - 1
Case vbKeyNumpad6
  X2 = X2 + 1
Case vbKeyEscape
  bRunning = False
Case vbKeyQ
  bRunning = False
  
End Select
End Sub



