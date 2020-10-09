VERSION 5.00
Begin VB.Form frmQuadraticGrapher 
   Caption         =   "Quadratic Grapher"
   ClientHeight    =   7365
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4260
   LinkTopic       =   "Form1"
   ScaleHeight     =   7365
   ScaleWidth      =   4260
   StartUpPosition =   3  'Windows Default
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
      Height          =   315
      Left            =   120
      TabIndex        =   9
      Top             =   5280
      Width           =   1935
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
      Height          =   315
      Left            =   2160
      TabIndex        =   8
      Top             =   5280
      Width           =   1935
   End
   Begin VB.TextBox txtC 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2940
      TabIndex        =   7
      Top             =   4860
      Width           =   1155
   End
   Begin VB.TextBox txtB 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1560
      TabIndex        =   6
      Top             =   4860
      Width           =   1155
   End
   Begin VB.TextBox txtA 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   240
      TabIndex        =   5
      Top             =   4860
      Width           =   1155
   End
   Begin VB.PictureBox picGraph 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   4000
      Left            =   120
      ScaleHeight     =   3975
      ScaleWidth      =   3975
      TabIndex        =   0
      Top             =   120
      Width           =   4000
   End
   Begin VB.Label lblVertex 
      Alignment       =   2  'Center
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
      Left            =   2160
      TabIndex        =   17
      Top             =   6900
      Width           =   1935
   End
   Begin VB.Label lblVertexPrompt 
      Alignment       =   2  'Center
      Caption         =   "Vertex"
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
      Left            =   2160
      TabIndex        =   16
      Top             =   6600
      Width           =   1875
   End
   Begin VB.Label lblYInt 
      Alignment       =   2  'Center
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
      Left            =   120
      TabIndex        =   15
      Top             =   6900
      Width           =   1935
   End
   Begin VB.Label lblYIntPrompt 
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
      Left            =   120
      TabIndex        =   14
      Top             =   6600
      Width           =   1935
   End
   Begin VB.Label lblXInt2 
      Alignment       =   2  'Center
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
      Left            =   2160
      TabIndex        =   13
      Top             =   6120
      Width           =   1935
   End
   Begin VB.Label lblXInt1 
      Alignment       =   2  'Center
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
      Left            =   120
      TabIndex        =   12
      Top             =   6120
      Width           =   1935
   End
   Begin VB.Label lblXInt2Prompt 
      Alignment       =   2  'Center
      Caption         =   "X-Intercept 2"
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
      Left            =   2160
      TabIndex        =   11
      Top             =   5820
      Width           =   1875
   End
   Begin VB.Label lblXInt1Prompt 
      Alignment       =   2  'Center
      Caption         =   "X-Intercept 1"
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
      Left            =   120
      TabIndex        =   10
      Top             =   5820
      Width           =   1935
   End
   Begin VB.Label lblCPrompt 
      Alignment       =   2  'Center
      Caption         =   "C="
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
      Left            =   2820
      TabIndex        =   4
      Top             =   4500
      Width           =   1500
   End
   Begin VB.Label lblBPrompt 
      Alignment       =   2  'Center
      Caption         =   "B="
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
      Left            =   1080
      TabIndex        =   3
      Top             =   4500
      Width           =   2220
   End
   Begin VB.Label lblAPrompt 
      Alignment       =   2  'Center
      Caption         =   "A="
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
      Left            =   120
      TabIndex        =   2
      Top             =   4500
      Width           =   1500
   End
   Begin VB.Label lblEquation 
      Alignment       =   2  'Center
      Caption         =   "y=ax^2+bx+c"
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
      Left            =   120
      TabIndex        =   1
      Top             =   4140
      Width           =   4035
   End
End
Attribute VB_Name = "frmQuadraticGrapher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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

Private Sub cmdExit_Click()

Hide

frmMainMenu.Show

End Sub

Private Sub cmdGraph_Click()

Dim i As Single

Dim a As Single
Dim b As Single
Dim c As Single

Dim vertexx As Single
Dim vertexy As Single

Dim testforxint As Single
Dim xint1 As Single
Dim xint2 As Single
Dim yint As Single
Dim noroots As Boolean

picGraph.Cls
ScaleDraw

a = txtA
b = txtB
c = txtC

'Vertex
vertexx = -b / (2 * a)
vertexy = (a * (vertexx ^ 2)) + (b * vertexx) + c

picGraph.Circle (vertexx, vertexy), 0.25, vbBlack

'xint

If b ^ 2 - 4 * a * c < 0 Then

    noroots = True
    
Else
    
    xint1 = (-b + Sqr(b ^ 2 - 4 * a * c)) / (2 * a)
    xint2 = (-b - Sqr(b ^ 2 - 4 * a * c)) / (2 * a)

    picGraph.Circle (xint1, 0), 0.25, vbBlack
    picGraph.Circle (xint2, 0), 0.25, vbBlack
    
    noroots = False

End If

'yint
yint = c

picGraph.Circle (0, yint), 0.25, vbBlack
picGraph.Circle (vertexx * 2, yint), 0.25, vbBlack

'lines

For i = -10 To 10 Step 0.0001

    If i = -10 Then
    
        picGraph.Circle (i, ((a * (i ^ 2)) + (b * i) + c)), 0.01, vbBlack

    Else
    
        picGraph.Line -(i, ((a * (i ^ 2)) + (b * i) + c)), vbBlack
    
    End If
    
Next i

'output
If noroots = False Then

    lblXInt1 = xint1
    lblXInt2 = xint2
    
Else

    lblXInt1 = "No roots."
    lblXInt2 = "No roots."
    
End If

lblYInt = yint
lblVertex = "(" + Str(vertexx) + ", " + Str(vertexy) + ")"

End Sub

Private Sub Form_Activate()

picGraph.Cls
ScaleDraw

End Sub

Private Sub Form_Unload(Cancel As Integer)

Hide

frmMainMenu.Show

End Sub
