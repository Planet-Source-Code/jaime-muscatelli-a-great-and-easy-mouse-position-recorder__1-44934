VERSION 5.00
Begin VB.Form frmMouse 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mouse Recorder"
   ClientHeight    =   1230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6480
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1230
   ScaleWidth      =   6480
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar hscrlSpeed 
      Height          =   255
      LargeChange     =   5
      Left            =   1920
      Max             =   100
      Min             =   1
      TabIndex        =   4
      Top             =   960
      Value           =   25
      Width           =   2055
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New"
      Enabled         =   0   'False
      Height          =   495
      Left            =   5280
      TabIndex        =   3
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "&Play"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3480
      TabIndex        =   2
      Top             =   360
      Width           =   1575
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "&Stop"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1920
      TabIndex        =   1
      Top             =   360
      Width           =   1455
   End
   Begin VB.CommandButton cmdRecord 
      Caption         =   "&Record"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
   Begin VB.Timer tmrMouse 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   0
      Top             =   0
   End
   Begin VB.Label lblAbout 
      AutoSize        =   -1  'True
      Caption         =   "©2003 Jaime Muscatelli"
      Height          =   195
      Left            =   600
      TabIndex        =   6
      Top             =   0
      Width           =   1695
   End
   Begin VB.Label lblSpeed 
      Caption         =   "Speed: 25"
      Height          =   255
      Left            =   4080
      TabIndex        =   5
      Top             =   960
      Width           =   1095
   End
End
Attribute VB_Name = "frmMouse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Type POINTAPI
        X As Long
        Y As Long
End Type
Private I As Long
Private sMouseArray() As String

Private Sub cmdNew_Click()
cmdStop.Enabled = False
tmrMouse.Enabled = False
cmdRecord.Enabled = True
cmdPlay.Enabled = False
I = 0
End Sub

Private Sub cmdPlay_Click()
Dim J As Long
Dim lPlay As Long
Dim lX As Long
Dim lY As Long
Dim sSplit() As String
lPlay = I

For J = 1 To lPlay
    sSplit = Split(sMouseArray(J - 1))
    lX = CLng(sSplit(0))
    lY = CLng(sSplit(1))
    Sleep hscrlSpeed.Value
    SetCursorPos lX, lY
Next J

End Sub

Private Sub cmdRecord_Click()
cmdStop.Enabled = True
tmrMouse.Enabled = True
cmdRecord.Enabled = False
cmdNew.Enabled = False
End Sub

Private Sub cmdStop_Click()
cmdStop.Enabled = False
tmrMouse.Enabled = False
cmdPlay.Enabled = True
cmdNew.Enabled = True
End Sub

Private Function RecordMouse() As String
Dim mouse As POINTAPI
GetCursorPos mouse
RecordMouse = mouse.X & " " & mouse.Y
End Function

Private Sub hscrlSpeed_Change()
lblSpeed.Caption = "Speed: " & hscrlSpeed.Value
End Sub

Private Sub hscrlSpeed_Scroll()
lblSpeed.Caption = "Speed: " & hscrlSpeed.Value
End Sub

Private Sub tmrMouse_Timer()
ReDim Preserve sMouseArray(I)
sMouseArray(I) = RecordMouse
I = I + 1
End Sub
