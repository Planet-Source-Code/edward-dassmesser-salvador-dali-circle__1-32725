VERSION 5.00
Begin VB.Form frmCircle 
   AutoRedraw      =   -1  'True
   Caption         =   "Circle"
   ClientHeight    =   9360
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   11310
   Icon            =   "circle.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9360
   ScaleWidth      =   11310
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkLock 
      Caption         =   "Lock Multipliers for Radius"
      Height          =   495
      Left            =   120
      TabIndex        =   17
      Top             =   1200
      Width           =   2535
   End
   Begin VB.HScrollBar hscMultSin 
      Height          =   255
      LargeChange     =   5
      Left            =   1200
      Max             =   50
      Min             =   1
      TabIndex        =   14
      Top             =   600
      Value           =   1
      Width           =   2055
   End
   Begin VB.HScrollBar hscMultCos 
      Height          =   255
      LargeChange     =   5
      Left            =   3360
      Max             =   50
      Min             =   1
      TabIndex        =   13
      Top             =   600
      Value           =   1
      Width           =   2055
   End
   Begin VB.HScrollBar hscAddCos 
      Height          =   255
      LargeChange     =   5
      Left            =   7680
      Max             =   50
      Min             =   -50
      TabIndex        =   11
      Top             =   600
      Width           =   2055
   End
   Begin VB.HScrollBar hscAddSin 
      Height          =   255
      LargeChange     =   5
      Left            =   5520
      Max             =   50
      Min             =   -50
      TabIndex        =   9
      Top             =   600
      Width           =   2055
   End
   Begin VB.HScrollBar hscAcc 
      Height          =   255
      LargeChange     =   10
      Left            =   7680
      Max             =   360
      Min             =   1
      TabIndex        =   4
      Top             =   0
      Value           =   360
      Width           =   2055
   End
   Begin VB.HScrollBar hscTimes 
      Height          =   255
      LargeChange     =   10
      Left            =   5520
      Max             =   360
      Min             =   1
      TabIndex        =   3
      Top             =   0
      Value           =   1
      Width           =   2055
   End
   Begin VB.HScrollBar hscThick 
      Height          =   255
      LargeChange     =   2
      Left            =   3360
      Max             =   15
      Min             =   1
      TabIndex        =   2
      Top             =   0
      Value           =   1
      Width           =   2055
   End
   Begin VB.HScrollBar hscSpeed 
      Height          =   255
      LargeChange     =   10
      Left            =   1200
      Max             =   360
      Min             =   1
      TabIndex        =   1
      Top             =   0
      Value           =   10
      Width           =   2055
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3960
      Top             =   1680
   End
   Begin VB.CommandButton cmdDraw 
      Caption         =   "Draw!"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label lblMultSin 
      BackStyle       =   0  'Transparent
      Caption         =   "Multiply Sin Value by: 1"
      Height          =   255
      Left            =   1200
      TabIndex        =   16
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label lblMultCos 
      BackStyle       =   0  'Transparent
      Caption         =   "Multiply Cos Value by: 1"
      Height          =   255
      Left            =   3360
      TabIndex        =   15
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label lblAddCos 
      BackStyle       =   0  'Transparent
      Caption         =   "Add To Cos Value: 0"
      Height          =   255
      Left            =   7680
      TabIndex        =   12
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label lblAddSin 
      BackStyle       =   0  'Transparent
      Caption         =   "Add To Sin Value: 0"
      Height          =   255
      Left            =   5520
      TabIndex        =   10
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label lblAcc 
      BackStyle       =   0  'Transparent
      Caption         =   "Accuracy: 360"
      Height          =   255
      Left            =   7680
      TabIndex        =   8
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label lblTimes 
      BackStyle       =   0  'Transparent
      Caption         =   "Times Around: 1"
      Height          =   255
      Left            =   5520
      TabIndex        =   7
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label lblWidth 
      BackStyle       =   0  'Transparent
      Caption         =   "Thickness: 1"
      Height          =   255
      Left            =   3360
      TabIndex        =   6
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label lblSpeed 
      BackStyle       =   0  'Transparent
      Caption         =   "Speed: 10"
      Height          =   255
      Left            =   1200
      TabIndex        =   5
      Top             =   360
      Width           =   1935
   End
   Begin VB.Menu mnuStuff 
      Caption         =   "&Stuff you might want to do"
      Begin VB.Menu mnuVote 
         Caption         =   "&Vote at Planet Source Code"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEmail 
         Caption         =   "E-mail Author (durnurd@hotmail.com)"
         Shortcut        =   ^E
      End
   End
End
Attribute VB_Name = "frmCircle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const Pi = 3.14159265358979
Dim A As Single, XCord As Single, YCord As Single, MidH As Integer
Dim MidV As Integer

Private Sub chkLock_Click()
    If chkLock.Value = vbChecked Then
        hscMultCos.Value = hscMultSin.Value
        hscMultCos.Enabled = False
    Else
        hscMultCos.Enabled = True
    End If
End Sub

Private Sub cmdDraw_Click()
    frmCircle.Cls
    PSet (MidH + 10000 / XCord, MidV + 10000 / YCord)
    A = 0
    If cmdDraw.Caption = "Draw!" Then cmdDraw.Caption = "Stop!" Else cmdDraw.Caption = "Draw!"
    Timer1.Enabled = Not Timer1.Enabled
End Sub

Private Sub Form_Load()
    MidH = frmCircle.Width / 2
    MidV = frmCircle.Height / 2
    A = 1
    XCord = Pi / Sin(A)
    YCord = Pi / Cos(A)
    PSet (MidH + 10000 / XCord, MidV + 10000 / YCord)
End Sub

Private Sub hscAddCos_Change()
    lblAddCos = "Add To Cos Value: " & hscAddCos.Value
End Sub

Private Sub hscAddSin_Change()
    lblAddSin = "Add To Sin Value: " & hscAddSin.Value
End Sub

Private Sub hscMultCos_Change()
    lblMultCos = "Multiply Cos Value By: " & hscMultCos.Value
End Sub

Private Sub hscMultSin_Change()
    If chkLock.Value = vbChecked Then hscMultCos.Value = hscMultSin.Value
    lblMultSin = "Multiply Sin Value By: " & hscMultSin.Value
End Sub

Private Sub hscSpeed_Change()
    lblSpeed.Caption = "Speed: " & hscSpeed.Value
End Sub

Private Sub hscThick_Change()
    lblWidth = "Thickness: " & hscThick.Value
    DrawWidth = hscThick.Value
End Sub

Private Sub hscTimes_Change()
    lblTimes = "Times Around: " & hscTimes.Value
End Sub

Private Sub hscAcc_Change()
    lblAcc = "Accuracy: " & hscAcc.Value
End Sub
Private Sub Timer1_Timer()
    MidH = frmCircle.Width / 2
    MidV = frmCircle.Height / 2
    If A > hscTimes.Value * 2 * Pi + 1 Then A = 0: Timer1.Enabled = False: Exit Sub
    For A = 1 To hscTimes.Value * 2 * Pi + 1 Step 1 / hscAcc.Value
        XCord = Pi / Sin(A) * hscMultSin.Value + hscAddSin.Value
        YCord = Pi / Cos(A) * hscMultCos.Value + hscAddCos.Value
        Line -(MidH + 10000 / XCord, MidV + 10000 / YCord), A * 1000
        If A * 360 Mod hscSpeed.Value = 0 Then Refresh: DoEvents
        If cmdDraw.Caption = "Draw!" Then A = 1: XCord = Pi / Sin(A) * hscMultSin.Value + hscAddSin.Value: YCord = Pi / Cos(A) * hscMultCos.Value + hscAddCos.Value: PSet (MidH + 10000 / XCord, MidV + 10000 / YCord): Exit Sub
    Next A
    A = 1: XCord = Pi / Sin(A) * hscMultSin.Value + hscAddSin.Value: YCord = Pi / Cos(A) * hscMultCos.Value + hscAddCos.Value: PSet (MidH + 10000 / XCord, MidV + 10000 / YCord)
    Timer1.Enabled = False
    cmdDraw.Caption = "Draw!"
End Sub

Private Sub mnuEmail_Click()
    StartURL "mailto:durnurd@hotmail.com"
End Sub
Private Sub mnuVote_Click()
    StartURL "http://www.Planet-Source-Code.com/vb/default.asp?lngCId=32725&lngWId=1"
End Sub
Private Sub StartURL(strURL As String)
    On Error Resume Next
    Shell "Explorer """ & strURL & """"
    If Err.Number <> 0 Then
        Err.Clear
        Shell "Start """ & strURL & """"
    End If
    If Err.Number <> 0 Then
        If MsgBox("Can't figure out how to navigate on this OS.  Copy the URL to the clipboard?", vbExclamation + vbYesNo) = vbYes Then
            Clipboard.SetText strURL
        End If
    End If
End Sub

