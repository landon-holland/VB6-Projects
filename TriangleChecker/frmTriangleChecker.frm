VERSION 5.00
Begin VB.Form frmTriangleChecker 
   Caption         =   "Triangle Checker"
   ClientHeight    =   9075
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4455
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
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
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
      TabIndex        =   8
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
      TabIndex        =   7
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
      Picture         =   "frmTriangleChecker.frx":0000
      TabIndex        =   6
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox txtSideC 
      Appearance      =   0  'Flat
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
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox txtSideB 
      Appearance      =   0  'Flat
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
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox txtSideA 
      Appearance      =   0  'Flat
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
      TabIndex        =   0
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
      TabIndex        =   22
      Top             =   8400
      Width           =   2415
   End
   Begin VB.Image imgObtuse 
      Height          =   1215
      Left            =   3000
      Picture         =   "frmTriangleChecker.frx":056B
      Top             =   7020
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Image imgRight 
      Height          =   1215
      Left            =   1620
      Picture         =   "frmTriangleChecker.frx":107E
      Top             =   7020
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Image imgAcute 
      Height          =   1215
      Left            =   240
      Picture         =   "frmTriangleChecker.frx":1CFB
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
      TabIndex        =   21
      Top             =   6300
      Width           =   2415
   End
   Begin VB.Image imgScalene 
      Appearance      =   0  'Flat
      Height          =   1215
      Left            =   3000
      Picture         =   "frmTriangleChecker.frx":27F6
      Top             =   4860
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Image imgIsosceles 
      Height          =   1215
      Left            =   1620
      Picture         =   "frmTriangleChecker.frx":30F6
      Top             =   4860
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Image imgEquilateral 
      Height          =   1215
      Left            =   240
      Picture         =   "frmTriangleChecker.frx":3C7B
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
      TabIndex        =   20
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
      TabIndex        =   19
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
      TabIndex        =   18
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
      TabIndex        =   17
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
      TabIndex        =   16
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
      TabIndex        =   15
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
      TabIndex        =   14
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
      TabIndex        =   13
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
      TabIndex        =   12
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
      TabIndex        =   11
      Top             =   3000
      Width           =   1275
   End
   Begin VB.Label lblIsNot 
      Alignment       =   2  'Center
      BackColor       =   &H00000FF0&
      Caption         =   "It is not a triangle."
      Height          =   1095
      Left            =   2280
      TabIndex        =   10
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
      TabIndex        =   9
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
      TabIndex        =   5
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
      TabIndex        =   4
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
      TabIndex        =   3
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
Sub Clear()
txtSideA = ""
txtSideB = ""
txtSideC = ""
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
txtSideA.SetFocus
End Sub


Private Sub cmdCalculate_Click()

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

'Check if numbers enetered are actually numbers.
If Not 1 = 1 Then
    MsgBox "One of the sides is not a number."
    Clear
Else
    'Take input of sides.
    sidea = Val(txtSideA)
    sideb = Val(txtSideB)
    sidec = Val(txtSideC)
    
    'Clear the previous triangle.
    Clear
    
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
            
        ElseIf anglea > 90 Or angleb > 90 Or anglec = 90 Then
            isobtuse = True
            
        Else
            isacute = True
            
        End If
    
    Else
        istriangle = False
        
    End If
    
    'Check boolean variables and display all of the triangle's information.
    'Reset side lengths.
    txtSideA = sidea
    txtSideB = sideb
    txtSideC = sidec
    
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
'Exit program.
End
End Sub

Private Sub txtSideA_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtSideB.SetFocus
End If
End Sub

Private Sub txtSideB_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtSideC.SetFocus
End If
End Sub

Private Sub txtSideC_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdCalculate_Click
End If
End Sub
