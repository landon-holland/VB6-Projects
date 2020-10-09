VERSION 5.00
Begin VB.Form frmTwoPoint 
   Caption         =   "Two Point"
   ClientHeight    =   7800
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4815
   LinkTopic       =   "Form1"
   ScaleHeight     =   7800
   ScaleWidth      =   4815
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdThreePoint 
      Caption         =   "Three-Point"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1740
      TabIndex        =   32
      Top             =   7200
      Width           =   1380
   End
   Begin VB.CommandButton cmdCalculations 
      Caption         =   "Calculations"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   28
      Top             =   7200
      Width           =   1500
   End
   Begin VB.TextBox txtY4 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3600
      TabIndex        =   19
      Top             =   6060
      Width           =   990
   End
   Begin VB.TextBox txtX4 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2460
      TabIndex        =   18
      Top             =   6060
      Width           =   990
   End
   Begin VB.TextBox txtY3 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1380
      TabIndex        =   17
      Top             =   6060
      Width           =   990
   End
   Begin VB.TextBox txtX3 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   240
      TabIndex        =   16
      Top             =   6060
      Width           =   990
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   13
      Top             =   6780
      Width           =   1500
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1740
      TabIndex        =   12
      Top             =   6780
      Width           =   1380
   End
   Begin VB.CommandButton cmdGraph 
      Caption         =   "Graph"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   6780
      Width           =   1500
   End
   Begin VB.TextBox txtY2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3600
      TabIndex        =   6
      Top             =   4980
      Width           =   990
   End
   Begin VB.TextBox txtX2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2460
      TabIndex        =   5
      Top             =   4980
      Width           =   990
   End
   Begin VB.TextBox txtY1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1380
      TabIndex        =   4
      Top             =   4980
      Width           =   990
   End
   Begin VB.TextBox txtX1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   240
      TabIndex        =   1
      Top             =   4980
      Width           =   990
   End
   Begin VB.PictureBox picGraph 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   4335
      Left            =   240
      ScaleHeight     =   4305
      ScaleWidth      =   4305
      TabIndex        =   0
      Top             =   240
      Width           =   4335
      Begin VB.Label lblMousePos 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   0
         TabIndex        =   31
         Top             =   0
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.Label lblMPoint2Pos 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   0
         TabIndex        =   30
         Top             =   0
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.Label lblMPoint1Pos 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   0
         TabIndex        =   29
         Top             =   0
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.Label lblPoint4Pos 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   325
         Left            =   0
         TabIndex        =   27
         Top             =   0
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.Label lblPoint3Pos 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   325
         Left            =   0
         TabIndex        =   26
         Top             =   0
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.Label lblPoint2Pos 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   325
         Left            =   0
         TabIndex        =   25
         Top             =   0
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.Label lblPoint1Pos 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   325
         Left            =   0
         TabIndex        =   24
         Top             =   0
         Visible         =   0   'False
         Width           =   1200
      End
   End
   Begin VB.Label lblY4Prompt 
      Alignment       =   2  'Center
      Caption         =   "Y"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3600
      TabIndex        =   23
      Top             =   6420
      Width           =   990
   End
   Begin VB.Label lblX4Prompt 
      Alignment       =   2  'Center
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2460
      TabIndex        =   22
      Top             =   6420
      Width           =   990
   End
   Begin VB.Label lblY3Prompt 
      Alignment       =   2  'Center
      Caption         =   "Y"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1380
      TabIndex        =   21
      Top             =   6420
      Width           =   990
   End
   Begin VB.Label lblX3Prompt 
      Alignment       =   2  'Center
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   240
      TabIndex        =   20
      Top             =   6420
      Width           =   990
   End
   Begin VB.Label lblPoint4Prompt 
      Alignment       =   2  'Center
      Caption         =   "Point 4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2340
      TabIndex        =   15
      Top             =   5700
      Width           =   2220
   End
   Begin VB.Label lblPoint3Prompt 
      Alignment       =   2  'Center
      Caption         =   "Point 3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   14
      Top             =   5700
      Width           =   2220
   End
   Begin VB.Label lblY2Prompt 
      Alignment       =   2  'Center
      Caption         =   "Y"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3600
      TabIndex        =   10
      Top             =   5340
      Width           =   990
   End
   Begin VB.Label lblX2Prompt 
      Alignment       =   2  'Center
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2460
      TabIndex        =   9
      Top             =   5340
      Width           =   990
   End
   Begin VB.Label lblY1Prompt 
      Alignment       =   2  'Center
      Caption         =   "Y"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1380
      TabIndex        =   8
      Top             =   5340
      Width           =   990
   End
   Begin VB.Label lblX1Prompt 
      Alignment       =   2  'Center
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   240
      TabIndex        =   7
      Top             =   5340
      Width           =   990
   End
   Begin VB.Label lblPoint2Prompt 
      Alignment       =   2  'Center
      Caption         =   "Point 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2340
      TabIndex        =   3
      Top             =   4620
      Width           =   2220
   End
   Begin VB.Label lblPoint1Prompt 
      Alignment       =   2  'Center
      Caption         =   "Point 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   2
      Top             =   4620
      Width           =   2220
   End
End
Attribute VB_Name = "frmTwoPoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public x1 As Single
Public x2 As Single
Public x3 As Single
Public x4 As Single
Public y1 As Single
Public y2 As Single
Public y3 As Single
Public y4 As Single
Public Slope1 As Single
Public Slope2 As Single
Public YInt1 As Single
Public YInt2 As Single
Public Mx1 As Single
Public My1 As Single
Public Mx2 As Single
Public My2 As Single
Public Distance1 As Single
Public Distance2 As Single
Public Ix As Single
Public Iy As Single
Public Gi As Integer

Sub ScaleDraw()

Dim i As Integer

picGraph.Scale (-10, 10)-(10, -10)
picGraph.Line (-10, 0)-(10, 0), vbRed
picGraph.Line (0, -10)-(0, 10), vbBlue

For i = -10 To 10 Step 1

    picGraph.Line (i, 0.5)-(i, -0.5), vbRed
    picGraph.Line (0.5, i)-(-0.5, i), vbBlue

Next i

End Sub

Private Sub cmdCalculations_Click()

frmCalculations.x1 = x1
frmCalculations.x2 = x2
frmCalculations.x3 = x3
frmCalculations.x4 = x4
frmCalculations.y1 = y1
frmCalculations.y2 = y2
frmCalculations.y3 = y3
frmCalculations.y4 = y4
frmCalculations.Slope1 = Slope1
frmCalculations.Slope2 = Slope2
frmCalculations.YInt1 = YInt1
frmCalculations.YInt2 = YInt2
frmCalculations.Mx1 = Mx1
frmCalculations.Mx2 = Mx2
frmCalculations.My1 = My1
frmCalculations.My2 = My2
frmCalculations.Distance1 = Distance1
frmCalculations.Distance2 = Distance2
frmCalculations.Ix = Ix
frmCalculations.Iy = Iy

frmCalculations.Show

End Sub

Private Sub cmdClear_Click()

x1 = 0
x2 = 0
x3 = 0
x4 = 0
y1 = 0
y2 = 0
y3 = 0
y4 = 0
Slope1 = 0
Slope2 = 0
YInt1 = 0
YInt2 = 0
Mx1 = 0
Mx2 = 0
My1 = 0
My2 = 0
Distance1 = 0
Distance2 = 0
Ix = 0
Iy = 0

txtX1 = ""
txtY1 = ""
txtX2 = ""
txtY2 = ""
txtX3 = ""
txtY3 = ""
txtX4 = ""
txtY4 = ""
picGraph.Cls
ScaleDraw
Gi = 0
lblPoint1Pos.Visible = False
lblPoint2Pos.Visible = False
lblPoint3Pos.Visible = False
lblPoint4Pos.Visible = False
lblMPoint1Pos.Visible = False
lblMPoint2Pos.Visible = False
lblMousePos.Visible = False
txtX1.SetFocus

End Sub

Private Sub cmdExit_Click()

End

End Sub

Private Sub cmdGraph_Click()

x1 = Val(txtX1)
y1 = Val(txtY1)
x2 = Val(txtX2)
y2 = Val(txtY2)
x3 = Val(txtX3)
y3 = Val(txtY3)
x4 = Val(txtX4)
y4 = Val(txtY4)

lblPoint1Pos.Visible = True
lblPoint1Pos = "(" + Format(x1, "fixed") + ", " + Format(y1, "fixed") + ")"
lblPoint1Pos.Top = y1 - 0.5
lblPoint1Pos.Left = x1 + 0.5

lblPoint2Pos.Visible = True
lblPoint2Pos = "(" + Format(x2, "fixed") + ", " + Format(y2, "fixed") + ")"
lblPoint2Pos.Top = y2 - 0.5
lblPoint2Pos.Left = x2 + 0.5

lblPoint3Pos.Visible = True
lblPoint3Pos = "(" + Format(x3, "fixed") + ", " + Format(y3, "fixed") + ")"
lblPoint3Pos.Top = y3 - 0.5
lblPoint3Pos.Left = x3 + 0.5

lblPoint4Pos.Visible = True
lblPoint4Pos = "(" + Format(x4, "fixed") + ", " + Format(y4, "fixed") + ")"
lblPoint4Pos.Top = y4 - 0.5
lblPoint4Pos.Left = x4 + 0.5

picGraph.Circle (x1, y1), 0.25, vbBlack
picGraph.Circle (x2, y2), 0.25, vbBlack
picGraph.Line (x1, y1)-(x2, y2), vbBlack
Mx1 = (x1 + x2) / 2
My1 = (y1 + y2) / 2
picGraph.Circle (Mx1, My1), 0.25, vbBlack

lblMPoint1Pos.Visible = True
lblMPoint1Pos = "(" + Format(Mx1, "fixed") + ", " + Format(My1, "fixed") + ")"
lblMPoint1Pos.Top = My1 - 0.5
lblMPoint1Pos.Left = Mx1 + 0.5

picGraph.Circle (x3, y3), 0.25, vbBlack
picGraph.Circle (x4, y4), 0.25, vbBlack
picGraph.Line (x3, y3)-(x4, y4), vbBlack
Mx2 = (x3 + x4) / 2
My2 = (y3 + y4) / 2
picGraph.Circle (Mx2, My2), 0.25, vbBlack

lblMPoint2Pos.Visible = True
lblMPoint2Pos = "(" + Format(Mx2, "fixed") + ", " + Format(My2, "fixed") + ")"
lblMPoint2Pos.Top = My2 - 0.5
lblMPoint2Pos.Left = Mx2 + 0.5

Gi = 5
End Sub

Private Sub cmdThreePoint_Click()

frmTwoPoint.Hide
frmTriangleChecker.Show

End Sub

Private Sub Form_Activate()

ScaleDraw

End Sub

Private Sub Form_Load()

Gi = 0

End Sub

Private Sub Form_Unload(Cancel As Integer)

End

End Sub

Private Sub picGraph_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

Gi = Gi + 1

If Gi = 1 Then

    x1 = X
    y1 = Y
    picGraph.Circle (x1, y1), 0.25, vbBlack
    txtX1 = x1
    txtY1 = y1
    
    lblPoint1Pos.Visible = True
    lblPoint1Pos = "(" + Format(x1, "fixed") + ", " + Format(y1, "fixed") + ")"
    lblPoint1Pos.Top = y1 - 0.5
    lblPoint1Pos.Left = x1 + 0.5

ElseIf Gi = 2 Then

    x2 = X
    y2 = Y
    picGraph.Line (x1, y1)-(x2, y2), vbBlack
    picGraph.Circle (x2, y2), 0.25, vbBlack
    txtX2 = x2
    txtY2 = y2
    
    lblPoint2Pos.Visible = True
    lblPoint2Pos = "(" + Format(x2, "fixed") + ", " + Format(y2, "fixed") + ")"
    lblPoint2Pos.Top = y2 - 0.5
    lblPoint2Pos.Left = x2 + 0.5
    
    Mx1 = (x1 + x2) / 2
    My1 = (y1 + y2) / 2
    picGraph.Circle (Mx1, My1), 0.25, vbBlack
    
    lblMPoint1Pos.Visible = True
    lblMPoint1Pos = "(" + Format(Mx1, "fixed") + ", " + Format(My1, "fixed") + ")"
    lblMPoint1Pos.Top = My1 - 0.5
    lblMPoint1Pos.Left = Mx1 + 0.5
    
ElseIf Gi = 3 Then

    x3 = X
    y3 = Y
    picGraph.Circle (x3, y3), 0.25, vbBlack
    txtX3 = x3
    txtY3 = y3
    
    lblPoint3Pos.Visible = True
    lblPoint3Pos = "(" + Format(x3, "fixed") + ", " + Format(y3, "fixed") + ")"
    lblPoint3Pos.Top = y3 - 0.5
    lblPoint3Pos.Left = x3 + 0.5
    
ElseIf Gi = 4 Then

    x4 = X
    y4 = Y
    picGraph.Line (x3, y3)-(x4, y4), vbBlack
    picGraph.Circle (x4, y4), 0.25, vbBlack
    txtX4 = x4
    txtY4 = y4
    
    lblPoint4Pos.Visible = True
    lblPoint4Pos = "(" + Format(x4, "fixed") + ", " + Format(y4, "fixed") + ")"
    lblPoint4Pos.Top = y4 - 0.5
    lblPoint4Pos.Left = x4 + 0.5
    
    Mx2 = (x3 + x4) / 2
    My2 = (y3 + y4) / 2
    picGraph.Circle (Mx2, My2), 0.25, vbBlack
    
    lblMPoint2Pos.Visible = True
    lblMPoint2Pos = "(" + Format(Mx2, "fixed") + ", " + Format(My2, "fixed") + ")"
    lblMPoint2Pos.Top = My2 - 0.5
    lblMPoint2Pos.Left = Mx2 + 0.5
    
End If

End Sub

Private Sub picGraph_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

lblMousePos.Visible = True
lblMousePos = "(" + Format(X, "fixed") + ", " + Format(Y, "fixed") + ")"
lblMousePos.Top = Y - 0.5
lblMousePos.Left = X + 0.5

End Sub
