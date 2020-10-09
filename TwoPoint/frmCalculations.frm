VERSION 5.00
Begin VB.Form frmCalculations 
   Caption         =   "Calculations"
   ClientHeight    =   3990
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11790
   LinkTopic       =   "Form1"
   ScaleHeight     =   3990
   ScaleWidth      =   11790
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
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
      Left            =   10380
      TabIndex        =   30
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label lblIntersectionPoint 
      Alignment       =   2  'Center
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
      Height          =   375
      Left            =   240
      TabIndex        =   32
      Top             =   3360
      Width           =   1875
   End
   Begin VB.Label lblIntersectionPointPrompt 
      Alignment       =   2  'Center
      Caption         =   "Intersection Point"
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
      TabIndex        =   31
      Top             =   3000
      Width           =   1875
   End
   Begin VB.Label lblMidpoint2 
      Alignment       =   2  'Center
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
      Height          =   375
      Left            =   10380
      TabIndex        =   29
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label lblMidpoint2Prompt 
      Alignment       =   2  'Center
      Caption         =   "Midpoint"
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
      Left            =   10380
      TabIndex        =   28
      Top             =   1980
      Width           =   1215
   End
   Begin VB.Label lblDistance2 
      Alignment       =   2  'Center
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
      Height          =   375
      Left            =   9000
      TabIndex        =   27
      Top             =   2400
      Width           =   1275
   End
   Begin VB.Label lblDistance2Prompt 
      Alignment       =   2  'Center
      Caption         =   "Distance"
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
      Left            =   9000
      TabIndex        =   26
      Top             =   1980
      Width           =   1275
   End
   Begin VB.Label lblStandardEquation2 
      Alignment       =   2  'Center
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
      Height          =   375
      Left            =   7320
      TabIndex        =   25
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label lblStandardEquation2Prompt 
      Alignment       =   2  'Center
      Caption         =   "Standard"
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
      Left            =   7320
      TabIndex        =   24
      Top             =   1980
      Width           =   1575
   End
   Begin VB.Label lblPointSlopeEquation2 
      Alignment       =   2  'Center
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
      Height          =   375
      Left            =   4740
      TabIndex        =   23
      Top             =   2400
      Width           =   2475
   End
   Begin VB.Label lblPointSlopeEquation2Prompt 
      Alignment       =   2  'Center
      Caption         =   "Point-Slope"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4740
      TabIndex        =   22
      Top             =   2040
      Width           =   2475
   End
   Begin VB.Label lblSlopeInterceptEquation2 
      Alignment       =   2  'Center
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
      Height          =   375
      Left            =   2580
      TabIndex        =   21
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Label lblSlopeInterceptEquation2Prompt 
      Alignment       =   2  'Center
      Caption         =   "Slope-Intercept"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2580
      TabIndex        =   20
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label lblYInt2 
      Alignment       =   2  'Center
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
      Height          =   375
      Left            =   1200
      TabIndex        =   19
      Top             =   2400
      Width           =   1275
   End
   Begin VB.Label lblYInt2Prompt 
      Alignment       =   2  'Center
      Caption         =   "Y-Intercept"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   18
      Top             =   2040
      Width           =   1275
   End
   Begin VB.Label lblSlope2 
      Alignment       =   2  'Center
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
      Height          =   375
      Left            =   240
      TabIndex        =   17
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label lblSlope2Prompt 
      Alignment       =   2  'Center
      Caption         =   "Slope"
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
      TabIndex        =   16
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label lblLine2Prompt 
      Caption         =   "Line 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   15
      Top             =   1620
      Width           =   975
   End
   Begin VB.Label lblStandardEquation1 
      Alignment       =   2  'Center
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
      Height          =   375
      Left            =   7320
      TabIndex        =   14
      Top             =   1020
      Width           =   1575
   End
   Begin VB.Label lblStandardEquation1Prompt 
      Alignment       =   2  'Center
      Caption         =   "Standard"
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
      Left            =   7320
      TabIndex        =   13
      Top             =   660
      Width           =   1575
   End
   Begin VB.Label lblPointSlopeEquation1 
      Alignment       =   2  'Center
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
      Height          =   375
      Left            =   4740
      TabIndex        =   12
      Top             =   1020
      Width           =   2475
   End
   Begin VB.Label lblPointSlopeEquation1Prompt 
      Alignment       =   2  'Center
      Caption         =   "Point-Slope"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4740
      TabIndex        =   11
      Top             =   660
      Width           =   2475
   End
   Begin VB.Label lblMidpoint1 
      Alignment       =   2  'Center
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
      Height          =   375
      Left            =   10380
      TabIndex        =   10
      Top             =   1020
      Width           =   1215
   End
   Begin VB.Label lblMidpoint1Prompt 
      Alignment       =   2  'Center
      Caption         =   "Midpoint"
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
      Left            =   10380
      TabIndex        =   9
      Top             =   660
      Width           =   1215
   End
   Begin VB.Label lblDistance1 
      Alignment       =   2  'Center
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
      Height          =   375
      Left            =   9000
      TabIndex        =   8
      Top             =   1020
      Width           =   1275
   End
   Begin VB.Label lblDistance1Prompt 
      Alignment       =   2  'Center
      Caption         =   "Distance"
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
      Left            =   9000
      TabIndex        =   7
      Top             =   660
      Width           =   1275
   End
   Begin VB.Label lblSlopeInterceptEquation1 
      Alignment       =   2  'Center
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
      Height          =   375
      Left            =   2580
      TabIndex        =   6
      Top             =   1020
      Width           =   2055
   End
   Begin VB.Label lblSlopeInterceptEquation1Prompt 
      Alignment       =   2  'Center
      Caption         =   "Slope-Intercept"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2580
      TabIndex        =   5
      Top             =   660
      Width           =   2055
   End
   Begin VB.Label lblYInt1 
      Alignment       =   2  'Center
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
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Top             =   1020
      Width           =   1275
   End
   Begin VB.Label lblYInt1Prompt 
      Alignment       =   2  'Center
      Caption         =   "Y-Intercept"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   3
      Top             =   660
      Width           =   1275
   End
   Begin VB.Label lblSlope1 
      Alignment       =   2  'Center
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
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1020
      Width           =   855
   End
   Begin VB.Label lblSlope1Prompt 
      Alignment       =   2  'Center
      Caption         =   "Slope"
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
      TabIndex        =   1
      Top             =   660
      Width           =   855
   End
   Begin VB.Label lblLine1Prompt 
      Caption         =   "Line 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "frmCalculations"
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
Dim TempNum1 As Single
Dim TempNum2 As Single

Private Sub cmdClose_Click()

frmTwoPoint.x1 = x1
frmTwoPoint.x2 = x2
frmTwoPoint.x3 = x3
frmTwoPoint.x4 = x4
frmTwoPoint.y1 = y1
frmTwoPoint.y2 = y2
frmTwoPoint.y3 = y3
frmTwoPoint.y4 = y4
frmTwoPoint.Slope1 = Slope1
frmTwoPoint.Slope2 = Slope2
frmTwoPoint.YInt1 = YInt1
frmTwoPoint.YInt2 = YInt2
frmTwoPoint.Mx1 = Mx1
frmTwoPoint.Mx2 = Mx2
frmTwoPoint.My1 = My1
frmTwoPoint.My2 = My2
frmTwoPoint.Distance1 = Distance1
frmTwoPoint.Distance2 = Distance2
frmTwoPoint.Ix = Ix
frmTwoPoint.Iy = Iy

frmCalculations.Hide

End Sub

Private Sub Form_Activate()

x1 = frmTwoPoint.x1
x2 = frmTwoPoint.x2
x3 = frmTwoPoint.x3
x4 = frmTwoPoint.x4
y1 = frmTwoPoint.y1
y2 = frmTwoPoint.y2
y3 = frmTwoPoint.y3
y4 = frmTwoPoint.y4
Slope1 = frmTwoPoint.Slope1
Slope2 = frmTwoPoint.Slope2
YInt1 = frmTwoPoint.YInt1
YInt2 = frmTwoPoint.YInt2
Mx1 = frmTwoPoint.Mx1
Mx2 = frmTwoPoint.Mx2
My1 = frmTwoPoint.My1
My2 = frmTwoPoint.My2
Distance1 = frmTwoPoint.Distance1
Distance2 = frmTwoPoint.Distance2
Ix = frmTwoPoint.Ix
Iy = frmTwoPoint.Iy

    If x2 - x1 <> 0 Then
        
        Slope1 = (y2 - y1) / (x2 - x1)
        lblSlope1 = Format(Slope1, "fixed")
    
        YInt1 = y1 - Slope1 * x1
        lblYInt1 = Format(YInt1, "fixed")
    
        lblSlopeInterceptEquation1 = "y = " + Format(Slope1, "fixed") + "x + " + Format(YInt1, "fixed")
        lblPointSlopeEquation1 = "y - " + Format(y1, "fixed") + " = " + Format(Slope1, "fixed") + "(x + " + Format(x1, "fixed") + ")"
        If Slope1 > 0 Then
    
            lblStandardEquation1 = "-" + Format(Slope1, "fixed") + "x + y = " + Format(YInt1, "fixed")
        
        Else
        
            lblStandardEquation1 = Format(Slope1, "fixed") + "x + y = " + Format(YInt1, "fixed")
    
        End If

    Else

        lblSlope1 = "None"
        lblYInt1 = "None"
    
        lblSlopeInterceptEquation1 = "x = " + Format(x1, "fixed")
        lblPointSlopeEquation1 = "N/A"
        lblStandardEquation1 = "N/A"
    
    End If

    Distance1 = Sqr((x2 - x1) ^ 2 + (y2 - y1) ^ 2)
    lblDistance1 = Format(Distance1, "fixed")

    Mx1 = (x1 + x2) / 2
    My1 = (y1 + y2) / 2

    lblMidpoint1 = "(" + Format(Mx1, "fixed") + ", " + Format(My1, "fixed") + ")"
    


    If x4 - x3 <> 0 Then
    
        Slope2 = (y4 - y3) / (x4 - x3)
        lblSlope2 = Format(Slope2, "fixed")
    
        YInt2 = y3 - Slope2 * x3
        lblYInt2 = Format(YInt2, "fixed")
    
        lblSlopeInterceptEquation2 = "y = " + Format(Slope2, "fixed") + "x + " + Format(YInt2, "fixed")
        lblPointSlopeEquation2 = "y - " + Format(y3, "fixed") + " = " + Format(Slope2, "fixed") + "(x + " + Format(x3, "fixed") + ")"
        If Slope2 > 0 Then
    
            lblStandardEquation2 = "-" + Format(Slope2, "fixed") + "x + y = " + Format(YInt2, "fixed")
        
        Else
        
            lblStandardEquation2 = Format(Slope2, "fixed") + "x + y = " + Format(YInt2, "fixed")
    
        End If

    Else

        lblSlope2 = "None"
        lblYInt2 = "None"
    
        lblSlopeInterceptEquation2 = "x = " + Format(x3, "fixed")
        lblPointSlopeEquation2 = "N/A"
        lblStandardEquation2 = "N/A"
    
    End If

    Distance2 = Sqr((x4 - x3) ^ 2 + (y4 - y3) ^ 2)
    lblDistance2 = Format(Distance2, "fixed")

    Mx2 = (x3 + x4) / 2
    My2 = (y3 + y4) / 2

    lblMidpoint2 = "(" + Format(Mx2, "fixed") + ", " + Format(My2, "fixed") + ")"

    TempNum1 = Slope1 - Slope2
    TempNum2 = YInt2 - YInt1
    Ix = TempNum2 / TempNum1
    Iy = (Slope1 * Ix) + YInt1

    lblIntersectionPoint = "(" + Format(Ix, "fixed") + ", " + Format(Iy, "fixed") + ")"

End Sub

