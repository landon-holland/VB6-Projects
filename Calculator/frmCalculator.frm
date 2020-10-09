VERSION 5.00
Begin VB.Form frmCalculator 
   BackColor       =   &H00404040&
   Caption         =   "Calculator"
   ClientHeight    =   4728
   ClientLeft      =   120
   ClientTop       =   464
   ClientWidth     =   7336
   ForeColor       =   &H80000015&
   LinkTopic       =   "Form1"
   ScaleHeight     =   4728
   ScaleWidth      =   7336
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picCalc 
      Height          =   1995
      Left            =   5100
      Picture         =   "frmCalculator.frx":0000
      ScaleHeight     =   1960
      ScaleWidth      =   1960
      TabIndex        =   23
      Top             =   2280
      Width           =   1995
   End
   Begin VB.CommandButton cmdPi 
      BackColor       =   &H00404040&
      Caption         =   "Pi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4020
      TabIndex        =   16
      ToolTipText     =   "Inserts Pi into the first number box."
      Top             =   3120
      Width           =   615
   End
   Begin VB.OptionButton optDeg 
      BackColor       =   &H00404040&
      Caption         =   "DEG"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.71
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   435
      Left            =   4020
      TabIndex        =   15
      Top             =   2400
      Width           =   855
   End
   Begin VB.OptionButton optRad 
      BackColor       =   &H00404040&
      Caption         =   "RAD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.71
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   4020
      TabIndex        =   14
      Top             =   1680
      Width           =   795
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00404040&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.71
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4020
      TabIndex        =   17
      Top             =   3840
      Width           =   615
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00404040&
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.71
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2760
      TabIndex        =   13
      Top             =   3840
      Width           =   615
   End
   Begin VB.CommandButton cmdTan 
      BackColor       =   &H00404040&
      Caption         =   "tan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2760
      TabIndex        =   12
      ToolTipText     =   "Tangent"
      Top             =   3120
      Width           =   615
   End
   Begin VB.CommandButton cmdCos 
      BackColor       =   &H00404040&
      Caption         =   "cos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   11.14
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2760
      TabIndex        =   11
      ToolTipText     =   "Cosine"
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton cmdSin 
      BackColor       =   &H00404040&
      Caption         =   "sin"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2760
      TabIndex        =   10
      ToolTipText     =   "Sine"
      Top             =   1680
      Width           =   615
   End
   Begin VB.CommandButton cmdAvg 
      BackColor       =   &H00404040&
      Caption         =   "Avg"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.86
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1500
      TabIndex        =   9
      ToolTipText     =   "Average"
      Top             =   3840
      Width           =   615
   End
   Begin VB.CommandButton cmdSqrt 
      BackColor       =   &H00404040&
      Caption         =   "Sqrt"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.86
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1500
      TabIndex        =   8
      ToolTipText     =   "Square Root"
      Top             =   3120
      Width           =   615
   End
   Begin VB.CommandButton cmdSquare 
      BackColor       =   &H00404040&
      Caption         =   "x²"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1500
      TabIndex        =   7
      ToolTipText     =   "Square"
      Top             =   2400
      Width           =   615
   End
   Begin VB.TextBox txtNumber2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.86
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2760
      TabIndex        =   1
      Top             =   1020
      Width           =   1875
   End
   Begin VB.TextBox txtNumber1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.86
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   240
      TabIndex        =   0
      Top             =   1020
      Width           =   1875
   End
   Begin VB.CommandButton cmdDivide 
      BackColor       =   &H00404040&
      Caption         =   "÷"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   240
      TabIndex        =   5
      ToolTipText     =   "Division"
      Top             =   3840
      Width           =   615
   End
   Begin VB.CommandButton cmdMultiply 
      BackColor       =   &H00404040&
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   240
      TabIndex        =   4
      ToolTipText     =   "Multiplication"
      Top             =   3120
      Width           =   615
   End
   Begin VB.CommandButton cmdSubtract 
      BackColor       =   &H00404040&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   240
      TabIndex        =   3
      ToolTipText     =   "Subtraction"
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton cmdPower 
      BackColor       =   &H00404040&
      Caption         =   "^"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1500
      TabIndex        =   6
      ToolTipText     =   "Power"
      Top             =   1680
      Width           =   615
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H00404040&
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   240
      TabIndex        =   2
      ToolTipText     =   "Addition"
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label lblAnswerPrompt 
      BackColor       =   &H00404040&
      Caption         =   "Answer:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5220
      TabIndex        =   22
      Top             =   660
      Width           =   1875
   End
   Begin VB.Label lblNumber2Prompt 
      BackColor       =   &H00404040&
      Caption         =   "Number 2:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2760
      TabIndex        =   21
      Top             =   660
      Width           =   1875
   End
   Begin VB.Label lblNumber1Prompt 
      BackColor       =   &H00404040&
      Caption         =   "Number 1:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   240
      TabIndex        =   20
      Top             =   660
      Width           =   1875
   End
   Begin VB.Label lblAnswer 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.86
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5220
      TabIndex        =   19
      ToolTipText     =   "Click me to copy to clipboard."
      Top             =   1020
      Width           =   1815
   End
   Begin VB.Label lblCalculatorTitle 
      BackColor       =   &H00404040&
      Caption         =   "Calculator"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   16.71
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      TabIndex        =   18
      Top             =   120
      Width           =   1875
   End
End
Attribute VB_Name = "frmCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
Dim Num1 As Single
Dim Num2 As Single
Dim Ans As Single

If Not IsNumeric(txtNumber1.Text) Or Not IsNumeric(txtNumber2.Text) Then
    MsgBox "A number you entered is actually not a number, you moron."
    txtNumber1.SetFocus
Else
    Num1 = txtNumber1.Text
    Num2 = txtNumber2.Text
    
    Ans = Num1 + Num2
    
    lblAnswer = Ans
End If
End Sub

Private Sub cmdAvg_Click()
Dim Num1 As Single
Dim Num2 As Single
Dim Ans As Single

If Not IsNumeric(txtNumber1.Text) Or Not IsNumeric(txtNumber2.Text) Then
    MsgBox "A number you entered is actually not a number, you moron."
    txtNumber1.SetFocus
Else
    Num1 = txtNumber1.Text
    Num2 = txtNumber2.Text
    
    Ans = Num1 + Num2
    Ans = Ans / 2
    
    lblAnswer = Ans
End If
End Sub

Private Sub cmdClear_Click()
txtNumber1 = ""
txtNumber2 = ""
lblAnswer = ""
txtNumber1.SetFocus
End Sub

Private Sub cmdCos_Click()
Dim Num1 As Double
Dim Ans As Double

If Not IsNumeric(txtNumber1.Text) Then
    MsgBox "The number you entered is actually not a number, you moron."
    txtNumber1.SetFocus
ElseIf Not txtNumber2.Text = "" Then
    MsgBox "You can't have anything in the second number text box if you are using the 'sin' function, you moron."
    txtNumber2.SetFocus
Else
    Num1 = txtNumber1.Text
    If optDeg = False And optRad = False Then
        MsgBox "Please select RAD or DEG, you moron."
        optRad.SetFocus
    ElseIf optRad = True Then
        Ans = Cos(Num1)
        
        lblAnswer = Ans
    ElseIf optDeg = True Then
        Num1 = Num1 * 3.14159265358979 / 180
        
        Ans = Cos(Num1)
        
        lblAnswer = Ans
    Else
        MsgBox "Internal error. (I'm the moron...)"
        txtNumber1.SetFocus
    End If
End If
End Sub

Private Sub cmdDivide_Click()
Dim Num1 As Single
Dim Num2 As Single
Dim Ans As Single

If Not IsNumeric(txtNumber1.Text) Or Not IsNumeric(txtNumber2.Text) Then
    MsgBox "A number you entered is actually not a number, you moron."
    txtNumber1.SetFocus
ElseIf txtNumber2.Text = 0 Then
    MsgBox "You can't divide by zero, you moron."
Else
    Num1 = txtNumber1.Text
    Num2 = txtNumber2.Text
    
    Ans = Num1 / Num2
    
    lblAnswer = Ans
End If
End Sub

Private Sub cmdExit_Click()
End
End Sub

Private Sub cmdMultiply_Click()
Dim Num1 As Single
Dim Num2 As Single
Dim Ans As Single

If Not IsNumeric(txtNumber1.Text) Or Not IsNumeric(txtNumber2.Text) Then
    MsgBox "A number you entered is actually not a number, you moron."
    txtNumber1.SetFocus
Else
    Num1 = txtNumber1.Text
    Num2 = txtNumber2.Text
    
    Ans = Num1 * Num2
    
    lblAnswer = Ans
End If
End Sub

Private Sub cmdPi_Click()
txtNumber1 = 3.14159265358979
End Sub

Private Sub cmdPower_Click()
Dim Num1 As Single
Dim Num2 As Single
Dim Ans As Single

If Not IsNumeric(txtNumber1.Text) Or Not IsNumeric(txtNumber2.Text) Then
    MsgBox "A number you entered is actually not a number, you moron."
    txtNumber1.SetFocus
ElseIf Num2 < 1 Then
    MsgBox "If you are using square root, use that function not this one."
    txtNumber1.SetFocus
Else
    Num1 = txtNumber1.Text
    Num2 = txtNumber2.Text
    
    Ans = Num1 ^ Num2
    
    lblAnswer = Ans
End If
End Sub

Private Sub cmdSin_Click()
Dim Num1 As Double
Dim Ans As Double

If Not IsNumeric(txtNumber1.Text) Then
    MsgBox "The number you entered is actually not a number, you moron."
    txtNumber1.SetFocus
ElseIf Not txtNumber2.Text = "" Then
    MsgBox "You can't have anything in the second number text box if you are using the 'sin' function, you moron."
    txtNumber2.SetFocus
Else
    Num1 = txtNumber1.Text
    If optDeg = False And optRad = False Then
        MsgBox "Please select RAD or DEG, you moron."
        optRad.SetFocus
    ElseIf optRad = True Then
        Ans = Sin(Num1)
        
        lblAnswer = Ans
    ElseIf optDeg = True Then
        Num1 = Num1 * 3.14159265358979 / 180
        
        Ans = Sin(Num1)
        
        lblAnswer = Ans
    Else
        MsgBox "Internal error. (I'm the moron...)"
        txtNumber1.SetFocus
    End If
End If
End Sub

Private Sub cmdSqrt_Click()
Dim Num1 As Single
Dim Ans As Single

If Not IsNumeric(txtNumber1.Text) Then
    MsgBox "The number you entered is actually not a number, you moron."
    txtNumber1.SetFocus
ElseIf Not txtNumber2.Text = "" Then
    MsgBox "You can't have anything in the second number text box if you are square rooting something, you moron."
    txtNumber2.SetFocus
ElseIf txtNumber1.Text < 0 Then
    Num1 = txtNumber1.Text * -1
    Ans = Num1 ^ 0.5
    
    lblAnswer = Str(Ans) + "i"
Else
    Num1 = txtNumber1.Text

    Ans = Num1 ^ 0.5
    
    lblAnswer = Ans
End If
End Sub

Private Sub cmdSquare_Click()
Dim Num1 As Single
Dim Ans As Single

If Not IsNumeric(txtNumber1.Text) Then
    MsgBox "The number you entered is actually not a number, you moron."
    txtNumber1.SetFocus
ElseIf Not txtNumber2.Text = "" Then
    MsgBox "You can't have anything in the second number text box if you are squaring something, you moron."
    txtNumber2.SetFocus
Else
    Num1 = txtNumber1.Text
    
    Ans = Num1 ^ 2
    
    lblAnswer = Ans
End If
End Sub

Private Sub cmdSubtract_Click()
Dim Num1 As Single
Dim Num2 As Single
Dim Ans As Single

If Not IsNumeric(txtNumber1.Text) Or Not IsNumeric(txtNumber2.Text) Then
    MsgBox "A number you entered is actually not a number, you moron."
    txtNumber1.SetFocus
Else
    Num1 = txtNumber1.Text
    Num2 = txtNumber2.Text
    
    Ans = Num1 - Num2
    
    lblAnswer = Ans
End If
End Sub

Private Sub cmdTan_Click()
Dim Num1 As Double
Dim Ans As Double

If Not IsNumeric(txtNumber1.Text) Then
    MsgBox "The number you entered is actually not a number, you moron."
    txtNumber1.SetFocus
ElseIf Not txtNumber2.Text = "" Then
    MsgBox "You can't have anything in the second number text box if you are using the 'sin' function, you moron."
    txtNumber2.SetFocus
Else
    Num1 = txtNumber1.Text
    If optDeg = False And optRad = False Then
        MsgBox "Please select RAD or DEG, you moron."
        optRad.SetFocus
    ElseIf optRad = True Then
        Ans = Tan(Num1)
        
        lblAnswer = Ans
    ElseIf optDeg = True Then
        Num1 = Num1 * 3.14159265358979 / 180
        
        Ans = Tan(Num1)
        
        lblAnswer = Ans
    Else
        MsgBox "Internal error. (I'm the moron...)"
        txtNumber1.SetFocus
    End If
End If
End Sub

Private Sub lblAnswer_Click()
Clipboard.Clear
Clipboard.SetText lblAnswer.Caption, vbCFText
End Sub

