VERSION 5.00
Begin VB.Form frmPrimeNumbers 
   Caption         =   "Prime Numbers"
   ClientHeight    =   5550
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8115
   LinkTopic       =   "Form1"
   ScaleHeight     =   5550
   ScaleWidth      =   8115
   StartUpPosition =   3  'Windows Default
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
      Height          =   315
      Left            =   5940
      TabIndex        =   27
      Top             =   4980
      Width           =   915
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
      Left            =   6960
      TabIndex        =   26
      Top             =   4980
      Width           =   915
   End
   Begin VB.ListBox lstFactors2 
      Appearance      =   0  'Flat
      Height          =   1200
      ItemData        =   "frmPrimeNumbers.frx":0000
      Left            =   6120
      List            =   "frmPrimeNumbers.frx":0002
      TabIndex        =   17
      Top             =   1800
      Width           =   1755
   End
   Begin VB.TextBox txtFactors2 
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
      Height          =   1230
      Left            =   4260
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   16
      Top             =   1800
      Width           =   1755
   End
   Begin VB.TextBox txtNumber2 
      Alignment       =   1  'Right Justify
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
      Height          =   345
      Left            =   5580
      TabIndex        =   2
      Top             =   240
      Width           =   2295
   End
   Begin VB.ListBox lstFactors1 
      Appearance      =   0  'Flat
      Height          =   1200
      ItemData        =   "frmPrimeNumbers.frx":0004
      Left            =   2100
      List            =   "frmPrimeNumbers.frx":0006
      TabIndex        =   7
      Top             =   1800
      Width           =   1755
   End
   Begin VB.TextBox txtFactors1 
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
      Height          =   1230
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   1800
      Width           =   1755
   End
   Begin VB.TextBox txtNumber1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1560
      TabIndex        =   1
      Top             =   240
      Width           =   2295
   End
   Begin VB.Label lblLCM 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   5040
      TabIndex        =   25
      Top             =   4500
      Width           =   2835
   End
   Begin VB.Label lblLCMPrompt 
      Caption         =   "LCM:"
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
      Left            =   4260
      TabIndex        =   24
      Top             =   4500
      Width           =   975
   End
   Begin VB.Label lblGCF 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1020
      TabIndex        =   23
      Top             =   4500
      Width           =   2835
   End
   Begin VB.Label lblGCFPrompt 
      Caption         =   "GCF:"
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
      Left            =   240
      TabIndex        =   22
      Top             =   4500
      Width           =   795
   End
   Begin VB.Label lblNumberOfFactors2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   6120
      TabIndex        =   21
      Top             =   3960
      Width           =   1755
   End
   Begin VB.Label lblSum2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   4260
      TabIndex        =   20
      Top             =   3960
      Width           =   1755
   End
   Begin VB.Label lblNumberOfFactorsPrompt2 
      Caption         =   "Number of Factors"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6120
      TabIndex        =   19
      Top             =   3180
      Width           =   1755
   End
   Begin VB.Label lblSumPrompt2 
      Alignment       =   2  'Center
      Caption         =   "Sum"
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
      Left            =   4260
      TabIndex        =   18
      Top             =   3120
      Width           =   1755
   End
   Begin VB.Label lblFactorsPrompt2 
      Alignment       =   2  'Center
      Caption         =   "Factors"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4260
      TabIndex        =   15
      Top             =   1320
      Width           =   3615
   End
   Begin VB.Label lblNotPrime2 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "Not Prime"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   435
      Left            =   6060
      TabIndex        =   14
      Top             =   780
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label lblPrime2 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   "Prime"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   435
      Left            =   4260
      TabIndex        =   13
      Top             =   780
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label lblNumber2Prompt 
      Caption         =   "Number 2: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4260
      TabIndex        =   12
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label lblNumberOfFactors1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   2100
      TabIndex        =   11
      Top             =   3960
      Width           =   1755
   End
   Begin VB.Label lblSum1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   240
      TabIndex        =   10
      Top             =   3960
      Width           =   1755
   End
   Begin VB.Label lblNumberOfFactorsPrompt1 
      Caption         =   "Number of Factors"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2100
      TabIndex        =   9
      Top             =   3120
      Width           =   1755
   End
   Begin VB.Label lblSumPrompt1 
      Alignment       =   2  'Center
      Caption         =   "Sum"
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
      Left            =   240
      TabIndex        =   8
      Top             =   3120
      Width           =   1755
   End
   Begin VB.Label lblFactorsPrompt1 
      Alignment       =   2  'Center
      Caption         =   "Factors"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   1320
      Width           =   3615
   End
   Begin VB.Label lblNotPrime1 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "Not Prime"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   435
      Left            =   2040
      TabIndex        =   4
      Top             =   780
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label lblPrime1 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   "Prime"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   435
      Left            =   240
      TabIndex        =   3
      Top             =   780
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label lblNumber1Prompt 
      Caption         =   "Number 1: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "frmPrimeNumbers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Number1 As Long
Dim Number2 As Long
Dim Sum1 As Long
Dim Sum2 As Long
Dim GCF As Long
Dim LCM As Long
Dim NumberOfFactors1 As Integer
Dim NumberOfFactors2 As Integer
Dim FactorList1 As String
Dim FactorList2 As String
Dim Number1Complete As Boolean
Dim Number2Complete As Boolean
Dim LCMFound As Boolean

Sub Clear()
txtNumber1 = ""
txtNumber2 = ""
lblPrime1.Visible = False
lblPrime2.Visible = False
lblNotPrime1.Visible = False
lblNotPrime2.Visible = False
txtFactors1 = ""
txtFactors2 = ""
lstFactors1.Clear
lstFactors2.Clear
lblSum1 = ""
lblSum2 = ""
lblNumberOfFactors1 = ""
lblNumberOfFactors2 = ""
lblGCF = ""
lblLCM = ""
Number1Complete = False
Number2Complete = False
End Sub

Sub Clear1()
txtNumber1 = ""
lblPrime1.Visible = False
lblNotPrime1.Visible = False
txtFactors1 = ""
lstFactors1.Clear
lblSum1 = ""
lblNumberOfFactors1 = ""
Number1Complete = False
End Sub

Sub Clear2()
txtNumber2 = ""
lblPrime2.Visible = False
lblNotPrime2.Visible = False
txtFactors2 = ""
lstFactors2.Clear
lblSum2 = ""
lblNumberOfFactors2 = ""
Number2Complete = False
End Sub

Sub Check()
Dim i As Integer
LCMFound = False
If Number1Complete = True And Number2Complete = True Then
    For i = 1 To Number1
        If Number1 Mod i = 0 And Number2 Mod i = 0 Then
            GCF = i
        End If
    Next i
    lblGCF = GCF
    
    LCM = (Number1 / GCF) * Number2
    lblLCM = LCM
End If
End Sub

Private Sub cmdClear_Click()
Clear
End Sub

Private Sub cmdExit_Click()
End
End Sub

Private Sub txtNumber1_KeyPress(KeyAscii As Integer)
Dim i As Integer

If KeyAscii = 13 Then
    If Not IsNumeric(txtNumber1) Then
        MsgBox "This is not a number."
    Else
        Sum1 = 0
        NumberOfFactors1 = 0
        FactorList1 = ""
        Number1 = Val(txtNumber1)
        Clear1
        
        For i = 1 To Number1
            If Number1 Mod i = 0 Then
                Sum1 = Sum1 + i
                FactorList1 = FactorList1 + Str(i) + " "
                lstFactors1.AddItem (Str(i))
                NumberOfFactors1 = NumberOfFactors1 + 1
            End If
        Next i
        
        txtFactors1 = FactorList1
        txtNumber1 = Number1
        lblNumberOfFactors1 = NumberOfFactors1
        lblSum1 = Sum1
        If Sum1 = Number1 + 1 Then
            lblPrime1.Visible = True
        Else
            lblNotPrime1.Visible = True
        End If
        Number1Complete = True
        txtNumber2.SetFocus
        Check
    End If
End If
End Sub

Private Sub txtNumber2_KeyPress(KeyAscii As Integer)
Dim i As Integer

If KeyAscii = 13 Then
    If Not IsNumeric(txtNumber1) Then
        MsgBox "This is not a number."
    Else
        Sum2 = 0
        NumberOfFactors2 = 0
        FactorList2 = ""
        Number2 = Val(txtNumber2)
        Clear2
        
        For i = 1 To Number2
            If Number2 Mod i = 0 Then
                Sum2 = Sum2 + i
                FactorList2 = FactorList2 + Str(i) + " "
                lstFactors2.AddItem (Str(i))
                NumberOfFactors2 = NumberOfFactors2 + 1
            End If
        Next i
        
        txtFactors2 = FactorList2
        txtNumber2 = Number2
        lblNumberOfFactors2 = NumberOfFactors2
        lblSum2 = Sum2
        If Sum2 = Number2 + 1 Then
            lblPrime2.Visible = True
        Else
            lblNotPrime2.Visible = True
        End If
        Number2Complete = True
        cmdClear.SetFocus
        Check
    End If
End If
End Sub
