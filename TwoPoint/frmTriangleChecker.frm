VERSION 5.00
Begin VB.Form frmTriangleChecker 
   Caption         =   "Three Point"
   ClientHeight    =   9075
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9405
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   18
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   9075
   ScaleWidth      =   9405
   StartUpPosition =   3  'Windows Default
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
      Left            =   7020
      TabIndex        =   36
      Top             =   6360
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
      Left            =   5880
      TabIndex        =   35
      Top             =   6360
      Width           =   990
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
      Left            =   8160
      TabIndex        =   29
      Top             =   5100
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
      Left            =   7020
      TabIndex        =   28
      Top             =   5100
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
      Left            =   5940
      TabIndex        =   27
      Top             =   5100
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
      Left            =   4800
      TabIndex        =   26
      Top             =   5100
      Width           =   990
   End
   Begin VB.PictureBox picGraph 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   4335
      Left            =   4800
      ScaleHeight     =   4305
      ScaleWidth      =   4305
      TabIndex        =   20
      Top             =   240
      Width           =   4335
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
      Height          =   312
      Left            =   3008
      TabIndex        =   5
      Top             =   960
      Width           =   1215
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
      Height          =   312
      Left            =   1620
      TabIndex        =   4
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "Calculate"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   1215
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
      Left            =   7020
      TabIndex        =   38
      Top             =   6780
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
      Left            =   5880
      TabIndex        =   37
      Top             =   6780
      Width           =   990
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
      Left            =   5880
      TabIndex        =   34
      Top             =   5880
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
      Left            =   8160
      TabIndex        =   33
      Top             =   5520
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
      Left            =   7020
      TabIndex        =   32
      Top             =   5520
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
      Left            =   5940
      TabIndex        =   31
      Top             =   5520
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
      Left            =   4800
      TabIndex        =   30
      Top             =   5520
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
      Left            =   6900
      TabIndex        =   25
      Top             =   4740
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
      Left            =   4800
      TabIndex        =   24
      Top             =   4740
      Width           =   2220
   End
   Begin VB.Label lblSideC 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   3000
      TabIndex        =   23
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label lblSideB 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   1620
      TabIndex        =   22
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label lblSideA 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   240
      TabIndex        =   21
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label lblAngleClassification 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   1020
      TabIndex        =   19
      Top             =   8400
      Width           =   2415
   End
   Begin VB.Image imgObtuse 
      Height          =   1215
      Left            =   3000
      Picture         =   "frmTriangleChecker.frx":0000
      Top             =   7020
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Image imgRight 
      Height          =   1215
      Left            =   1620
      Picture         =   "frmTriangleChecker.frx":0B13
      Top             =   7020
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Image imgAcute 
      Height          =   1215
      Left            =   240
      Picture         =   "frmTriangleChecker.frx":1790
      Top             =   7020
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblSideClassification 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   1020
      TabIndex        =   18
      Top             =   6300
      Width           =   2415
   End
   Begin VB.Image imgScalene 
      Appearance      =   0  'Flat
      Height          =   1215
      Left            =   3000
      Picture         =   "frmTriangleChecker.frx":228B
      Top             =   4860
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Image imgIsosceles 
      Height          =   1215
      Left            =   1620
      Picture         =   "frmTriangleChecker.frx":2B8B
      Top             =   4860
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Image imgEquilateral 
      Height          =   1215
      Left            =   240
      Picture         =   "frmTriangleChecker.frx":3710
      Top             =   4860
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblAngleCPrompt 
      Caption         =   "angle c"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3000
      TabIndex        =   17
      Top             =   4420
      Width           =   855
   End
   Begin VB.Label lblAngleC 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   3000
      TabIndex        =   16
      Top             =   4020
      Width           =   1215
   End
   Begin VB.Label lblAngleBPrompt 
      Caption         =   "angle b"
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
      Left            =   1620
      TabIndex        =   15
      Top             =   4420
      Width           =   1095
   End
   Begin VB.Label lblAngleB 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   1620
      TabIndex        =   14
      Top             =   4020
      Width           =   1215
   End
   Begin VB.Label lblAngleA 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   312
      Left            =   240
      TabIndex        =   13
      Top             =   4020
      Width           =   1215
   End
   Begin VB.Label lblAngleAPrompt 
      Caption         =   "angle a"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   12
      Top             =   4420
      Width           =   1215
   End
   Begin VB.Label lblArea 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   960
      TabIndex        =   11
      Top             =   3420
      Width           =   3255
   End
   Begin VB.Label lblAreaPrompt 
      Caption         =   "Area:"
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
      TabIndex        =   10
      Top             =   3480
      Width           =   675
   End
   Begin VB.Label lblPerimeter 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   1560
      TabIndex        =   9
      Top             =   2940
      Width           =   2655
   End
   Begin VB.Label lblPerimeterPrompt 
      Caption         =   "Perimeter:"
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
      TabIndex        =   8
      Top             =   3000
      Width           =   1275
   End
   Begin VB.Label lblIsNot 
      Alignment       =   2  'Center
      BackColor       =   &H00000FF0&
      Caption         =   "It is not a triangle."
      Height          =   1095
      Left            =   2280
      TabIndex        =   7
      Top             =   1560
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label lblIs 
      Alignment       =   2  'Center
      BackColor       =   &H000FF000&
      Caption         =   "It is a triangle."
      Height          =   1095
      Left            =   240
      TabIndex        =   6
      Top             =   1560
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label lblSideCPrompt 
      Caption         =   "side c"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   3008
      TabIndex        =   2
      Top             =   640
      Width           =   776
   End
   Begin VB.Label lblSideBPrompt 
      Caption         =   "side b"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   1620
      TabIndex        =   1
      Top             =   640
      Width           =   776
   End
   Begin VB.Label lblSideAPrompt 
      Caption         =   "side a"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   240
      TabIndex        =   0
      Top             =   640
      Width           =   776
   End
End
Attribute VB_Name = "frmTriangleChecker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Declare variables.
Dim sidea As Single
Dim sideb As Single
Dim sidec As Single
Dim anglea As Single
Dim angleb As Single
Dim anglec As Single
Dim lcos As Single
Dim perimeter As Single
Dim sperimeter As Single
Dim area As Single
Dim istriangle As Boolean
Dim isequilateral As Boolean
Dim isscalene As Boolean
Dim isisosceles As Boolean
Dim isright As Boolean
Dim isacute As Boolean
Dim isobtuse As Boolean
Dim x1 As Single
Dim x2 As Single
Dim x3 As Single
Dim y1 As Single
Dim y2 As Single
Dim y3 As Single
Dim Gi As Integer

Sub Clear()
lblSideA = ""
lblSideB = ""
lblSideC = ""
lblIs.Visible = False
lblIsNot.Visible = False
lblPerimeter = ""
lblArea = ""
lblAngleA = ""
lblAngleB = ""
lblAngleC = ""
imgEquilateral.Visible = False
imgIsosceles.Visible = False
imgScalene.Visible = False
lblSideClassification = ""
imgAcute.Visible = False
imgRight.Visible = False
imgObtuse.Visible = False
lblAngleClassification = ""
txtX1.SetFocus
picGraph.Cls
ScaleDraw
Gi = 0
txtX1 = ""
txtX2 = ""
txtX3 = ""
txtY1 = ""
txtY2 = ""
txtY3 = ""
isobtuse = False
isacute = False
isright = False
isequilateral = False
isscalene = False
isisosceles = False
End Sub

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


Private Sub cmdCalculate_Click()

'Check if numbers enetered are actually numbers.
If Not 1 = 1 Then
    MsgBox "One of the sides is not a number."
    Clear
Else
    'Take input of points.
    x1 = txtX1
    x2 = txtX2
    x3 = txtX3
    y1 = txtY1
    y2 = txtY2
    y3 = txtY3
    
    picGraph.Circle (x1, y1), 0.25, vbBlack
    picGraph.Circle (x2, y2), 0.25, vbBlack
    picGraph.Circle (x3, y3), 0.25, vbBlack
    
    picGraph.Line (x1, y1)-(x2, y2), vbBlack
    picGraph.Line (x2, y2)-(x3, y3), vbBlack
    picGraph.Line (x3, y3)-(x1, y1), vbBlack
    
    'Calculate side lengths.
    sidea = Sqr((x2 - x1) ^ 2 + (y2 - y1) ^ 2)
    sideb = Sqr((x3 - x2) ^ 2 + (y3 - y2) ^ 2)
    sidec = Sqr((x3 - x1) ^ 2 + (y3 - y1) ^ 2)
    
    'Clear the previous triangle.
    'Clear
    
    'Check if sides are making a triangle.
    If sidea + sideb > sidec And sidea + sidec > sideb And sideb + sidec > sidea Then
        istriangle = True
        
        'Calculate perimeter.
        perimeter = sidea + sideb + sidec
        
        'Calculate semiperimeter.
        sperimeter = perimeter / 2
        
        'Calculate area.
        area = Sqr(sperimeter * (sperimeter - sidea) * (sperimeter - sideb) * (sperimeter - sidec))
        
        'Check if sides make equilateral, scalene, or isosceles triangle.
        If sidea = sideb And sideb = sidec Then
            isequilateral = True
            isscalene = False
            isisosceles = False
            
        ElseIf Not sidea = sideb And Not sideb = sidec Then
            isscalene = True
            isequilateral = False
            isisosceles = False
            
        Else
            isisosceles = True
            isequilateral = False
            isscalene = False
            
        End If
        
        'Find measures of angle.
        'Find measure of angle A. (adj1 = side a, adj2 = side c, opp = side b).
        lcos = ((sidea ^ 2) + (sidec ^ 2) - (sideb ^ 2)) / (2 * sidea * sidec)
        anglea = Atn(-lcos / Sqr(-lcos * lcos + 1)) + 2 * Atn(1)
        anglea = anglea * 180 / 3.14159265358979

        'Find measure of Angle B. (adj1 = side a, adj = side b, opp = side c).
        lcos = ((sidea ^ 2) + (sideb ^ 2) - (sidec ^ 2)) / (2 * sidea * sideb)
        angleb = Atn(-lcos / Sqr(-lcos * lcos + 1)) + 2 * Atn(1)
        angleb = angleb * 180 / 3.14159265358979
        
        'Find measure of Angle C. (180 - (angle a + angle b).)
        anglec = 180 - (anglea + angleb)
        
        'Round angles.
        anglea = Math.Round(anglea)
        angleb = Math.Round(angleb)
        anglec = Math.Round(anglec)
        
        'Check if angles make acute, right, or obtuse triangle.
        If anglea = 90 Or angleb = 90 Or anglec = 90 Then
            isright = True
            
        ElseIf anglea > 90 Or angleb > 90 Or anglec > 90 Then
            isobtuse = True
        Else
            isacute = True
            
        End If
    
    Else
        istriangle = False
        
    End If
    
    'Check boolean variables and display all of the triangle's information.
    'Reset side lengths.
    lblSideA = sidea
    lblSideB = sideb
    lblSideC = sidec
    
    'Display "Is Triangle".
    If istriangle = True Then
        lblIs.Visible = True
        'Display perimeter.
        lblPerimeter = perimeter
    
        'Display area.
        lblArea = area
    
        'Display angle measures.
        lblAngleA = anglea
        lblAngleB = angleb
        lblAngleC = anglec
    
        'Display triangle classification.
        If isequilateral = True Then
            imgEquilateral.Visible = True
            lblSideClassification = "Equilateral"
        
        ElseIf isisosceles = True Then
            imgIsosceles.Visible = True
            lblSideClassification = "Isosceles"
            
        ElseIf isscalene = True Then
            imgScalene.Visible = True
            lblSideClassification = "Scalene"
        
        End If
    
        If isacute = True Then
            imgAcute.Visible = True
            lblAngleClassification = "Acute"
        
        ElseIf isright = True Then
            imgRight.Visible = True
            lblAngleClassification = "Right"
        
        ElseIf isobtuse = True Then
            imgObtuse.Visible = True
            lblAngleClassification = "Obtuse"
    
        End If
    ElseIf istriangle = False Then
        lblIsNot.Visible = True
        
    Else
        MsgBox "Internal error 1."
    
    End If

End If

cmdClear.SetFocus

End Sub

Private Sub cmdClear_Click()
'Clear previous triangle.
Clear
End Sub

Private Sub cmdExit_Click()
frmTriangleChecker.Hide
frmTwoPoint.Show
End Sub

Private Sub Form_Activate()

ScaleDraw
Gi = 0

End Sub

Private Sub Form_Unload(Cancel As Integer)

frmTwoPoint.Show

End Sub

Private Sub picGraph_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

Gi = Gi + 1

If Gi = 1 Then
    
    x1 = X
    y1 = Y
    txtX1 = x1
    txtY1 = y1
    picGraph.Circle (x1, y1), 0.25, vbBlack
    
ElseIf Gi = 2 Then

    x2 = X
    y2 = Y
    txtX2 = x2
    txtY2 = y2
    picGraph.Line (x1, y1)-(x2, y2), vbBlack
    picGraph.Circle (x2, y2), 0.25, vbBlack

ElseIf Gi = 3 Then

    x3 = X
    y3 = Y
    txtX3 = x3
    txtY3 = y3
    picGraph.Line (x2, y2)-(x3, y3), vbBlack
    picGraph.Line (x1, y1)-(x3, y3), vbBlack
    cmdCalculate_Click

End If

End Sub
