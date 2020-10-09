VERSION 5.00
Begin VB.Form frmDiveJudge 
   Caption         =   "Dive Judge"
   ClientHeight    =   4170
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3915
   LinkTopic       =   "Form1"
   ScaleHeight     =   4170
   ScaleWidth      =   3915
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Appearance      =   0  'Flat
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
      Left            =   0
      TabIndex        =   12
      Top             =   3720
      Width           =   3885
   End
   Begin VB.CommandButton cmdRNG 
      Appearance      =   0  'Flat
      Caption         =   "RNG"
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
      Left            =   1980
      TabIndex        =   4
      Top             =   1560
      Width           =   1935
   End
   Begin VB.CommandButton cmdUserInput 
      Appearance      =   0  'Flat
      Caption         =   "User Input"
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
      Left            =   0
      TabIndex        =   3
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox txtDifficulty 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   1500
      TabIndex        =   0
      Top             =   480
      Width           =   840
   End
   Begin VB.Label lblTotalScore 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1560
      TabIndex        =   11
      Top             =   3120
      Width           =   795
   End
   Begin VB.Label lblTotalScorePrompt 
      Alignment       =   2  'Center
      Caption         =   "Total Score"
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
      Left            =   -60
      TabIndex        =   10
      Top             =   2700
      Width           =   4005
   End
   Begin VB.Label lblScore 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   555
      Index           =   4
      Left            =   3120
      TabIndex        =   9
      Top             =   2040
      Width           =   795
   End
   Begin VB.Label lblScore 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   555
      Index           =   3
      Left            =   2340
      TabIndex        =   8
      Top             =   2040
      Width           =   795
   End
   Begin VB.Label lblScore 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   555
      Index           =   2
      Left            =   1560
      TabIndex        =   7
      Top             =   2040
      Width           =   795
   End
   Begin VB.Label lblScore 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   555
      Index           =   1
      Left            =   780
      TabIndex        =   6
      Top             =   2040
      Width           =   795
   End
   Begin VB.Label lblScore 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   555
      Index           =   0
      Left            =   0
      TabIndex        =   5
      Top             =   2040
      Width           =   795
   End
   Begin VB.Label lblJudgeScoresPrompt 
      Alignment       =   2  'Center
      Caption         =   "Judge Scores"
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
      Left            =   -60
      TabIndex        =   2
      Top             =   1140
      Width           =   4005
   End
   Begin VB.Label lblDifficultyPrompt 
      Alignment       =   2  'Center
      Caption         =   "Difficulty"
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
      Left            =   -60
      TabIndex        =   1
      Top             =   60
      Width           =   4005
   End
End
Attribute VB_Name = "frmDiveJudge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim arrscores(1 To 5) As Integer
Dim difficulty As Single
Dim finalscore As Single

Sub CalculateScore()

Dim i As Integer
Dim big As Integer
Dim low As Integer

difficulty = txtDifficulty

finalscore = 0

big = 0
low = 10

For i = 1 To 5

    If arrscores(i) > big Then
    
        big = arrscores(i)
        
    End If
    
    If arrscores(i) < low Then
    
        low = arrscores(i)
        
    End If

Next i

For i = 1 To 5

    If arrscores(i) <> big And arrscores(i) <> low Then
    
        finalscore = finalscore + arrscores(i)
    
    End If

Next i

finalscore = finalscore * difficulty

lblTotalScore = finalscore

End Sub

Private Sub cmdExit_Click()

Hide

frmMainMenu.Show

End Sub

Private Sub cmdRNG_Click()

Dim i As Integer

For i = 1 To 5

    arrscores(i) = Int((Rnd * 10) + 1)
    lblScore(i - 1) = arrscores(i)
    
Next i

CalculateScore

End Sub

Private Sub cmdUserInput_Click()

Dim i As Integer
Dim answer As String

For i = 1 To 5

    arrscores(i) = InputBox("Enter Judge " + Str(i) + "'s Score", "Input")
    lblScore(i - 1) = arrscores(i)

Next i

CalculateScore

End Sub

Private Sub Form_Unload(Cancel As Integer)

Hide

frmMainMenu.Show

End Sub

Private Sub txtDifficulty_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then

    cmdUserInput.SetFocus

End If

End Sub
