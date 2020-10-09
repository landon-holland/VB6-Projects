VERSION 5.00
Begin VB.Form frmMillionaire 
   BackColor       =   &H00000000&
   Caption         =   "Who Wants to Be a Millionaire?"
   ClientHeight    =   6570
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13695
   LinkTopic       =   "Form1"
   ScaleHeight     =   6570
   ScaleWidth      =   13695
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrGame 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton cmdWalkAway 
      Caption         =   "Walk Away"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Segoe UI Symbol"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   9840
      TabIndex        =   12
      Top             =   2520
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Segoe UI Symbol"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   9840
      TabIndex        =   7
      Top             =   600
      Width           =   3855
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Play"
      BeginProperty Font 
         Name            =   "Segoe UI Symbol"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   9840
      TabIndex        =   6
      Top             =   60
      Width           =   3855
   End
   Begin VB.Label lblTimer 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Poplar Std"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   8940
      TabIndex        =   13
      Top             =   0
      Width           =   795
   End
   Begin VB.Label lblPayout 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11220
      TabIndex        =   11
      Top             =   1860
      Width           =   2295
   End
   Begin VB.Label lblPayoutPrompt 
      BackColor       =   &H00000000&
      Caption         =   "Payout:"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   9840
      TabIndex        =   10
      Top             =   1860
      Width           =   1395
   End
   Begin VB.Label lblQuestionNumber 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11220
      TabIndex        =   9
      Top             =   1260
      Width           =   375
   End
   Begin VB.Label lblQuestionNumberPrompt 
      BackColor       =   &H00000000&
      Caption         =   "Question:"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   9840
      TabIndex        =   8
      Top             =   1260
      Width           =   1395
   End
   Begin VB.Line lneDivider2 
      BorderColor     =   &H00FFFFFF&
      X1              =   9780
      X2              =   13680
      Y1              =   1140
      Y2              =   1140
   End
   Begin VB.Label lblQuestion 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1815
      Left            =   -60
      TabIndex        =   5
      Top             =   1680
      Width           =   9765
   End
   Begin VB.Line lneDivider 
      BorderColor     =   &H00FFFFFF&
      X1              =   9780
      X2              =   9780
      Y1              =   0
      Y2              =   7620
   End
   Begin VB.Label lblAns 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Nirmala UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1155
      Index           =   3
      Left            =   4860
      TabIndex        =   4
      Top             =   5160
      Width           =   4995
   End
   Begin VB.Label lblAns 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Nirmala UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1155
      Index           =   2
      Left            =   0
      TabIndex        =   3
      Top             =   5160
      Width           =   4995
   End
   Begin VB.Label lblAns 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Nirmala UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1155
      Index           =   1
      Left            =   4860
      TabIndex        =   2
      Top             =   3900
      Width           =   4995
   End
   Begin VB.Label lblAns 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Nirmala UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1155
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   3900
      Width           =   4995
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Who Wants to Be a Millionaire?"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1815
      Left            =   -60
      TabIndex        =   0
      Top             =   -60
      Width           =   9765
   End
End
Attribute VB_Name = "frmMillionaire"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim arrq(1 To 14) As String
Dim arra(1 To 56) As String
Dim question As Integer
Dim timer As Integer
Dim payout As Currency

Public Function RandomNumber(ByVal MaxValue As Long, Optional _
ByVal MinValue As Long = 0)

  On Error Resume Next
  Randomize
  RandomNumber = Int((MaxValue - MinValue + 1) * Rnd) + MinValue

End Function

Sub ShuffleAnswers()

Dim i As Integer
Dim rng As Integer
Dim temp As String

For i = 1 To 4

    rng = RandomNumber(4, 1)
    
    temp = lblAns(i - 1)
    lblAns(i - 1) = lblAns(rng - 1)
    lblAns(rng - 1) = temp

Next i

End Sub

Sub GameOver()

MsgBox "You won " + Format(payout, "Currency") + "."

tmrGame.Enabled = False

End Sub

Sub SpawnQuestion()

Dim i As Integer

question = question + 1
lblQuestionNumber = question
lblQuestion = arrq(question)
timer = 60
lblTimer = 60

If question = 1 Then

    payout = 500
    
    For i = 1 To 4
    
        lblAns(i - 1) = arra(i)
    
    Next i
    
    ShuffleAnswers
    
ElseIf question = 2 Then

    payout = 1000
    
    For i = 5 To 8
    
        lblAns(i - 5) = arra(i)
    
    Next i
    
    ShuffleAnswers
    
ElseIf question = 3 Then

    payout = 2000
    
    For i = 9 To 12
    
        lblAns(i - 9) = arra(i)
    
    Next i
    
    ShuffleAnswers
    
ElseIf question = 4 Then

    payout = 3000
    
    For i = 13 To 16
    
        lblAns(i - 13) = arra(i)
    
    Next i
    
    ShuffleAnswers
    
ElseIf question = 5 Then

    payout = 5000
    
    For i = 17 To 20
    
        lblAns(i - 17) = arra(i)
    
    Next i
    
    ShuffleAnswers
    
ElseIf question = 6 Then

    payout = 7000
    
    For i = 21 To 24
    
        lblAns(i - 21) = arra(i)
    
    Next i
    
    ShuffleAnswers
    
ElseIf question = 7 Then

    payout = 10000
    
    For i = 25 To 28
    
        lblAns(i - 25) = arra(i)
    
    Next i
    
    ShuffleAnswers
    
ElseIf question = 8 Then

    payout = 20000
    
    For i = 29 To 32
    
        lblAns(i - 29) = arra(i)
    
    Next i
    
    ShuffleAnswers

ElseIf question = 9 Then

    payout = 30000
    
    For i = 33 To 36
    
        lblAns(i - 33) = arra(i)
    
    Next i
    
    ShuffleAnswers
    
ElseIf question = 10 Then

    payout = 50000
    
    For i = 37 To 40
    
        lblAns(i - 5) = arra(i)
    
    Next i
    
    ShuffleAnswers
    
ElseIf question = 11 Then

    payout = 100000
    
    For i = 41 To 44
    
        lblAns(i - 5) = arra(i)
    
    Next i
    
    ShuffleAnswers
    
ElseIf question = 12 Then

    payout = 250000
    
    For i = 45 To 48
    
        lblAns(i - 45) = arra(i)
    
    Next i
    
    ShuffleAnswers
    
ElseIf question = 13 Then

    payout = 500000
    
    For i = 49 To 52
    
        lblAns(i - 49) = arra(i)
    
    Next i
    
    ShuffleAnswers
    
ElseIf question = 14 Then

    payout = 1000000
    
    For i = 53 To 56
    
        lblAns(i - 53) = arra(i)
    
    Next i
    
    ShuffleAnswers

End If

lblPayout = Format(payout, "Currency")

End Sub

Sub CorrectAnswer()

MsgBox "Correct!"

If question = 14 Then

    GameOver
    
Else

    SpawnQuestion
    
End If

End Sub

Sub IncorrectAnswer()

Dim thecorrectanswer As String

If question = 1 Then

    thecorrectanswer = arra(1)
    
ElseIf question = 2 Then
    
    thecorrectanswer = arra(5)
    
ElseIf question = 3 Then

    thecorrectanswer = arra(9)
    
ElseIf question = 4 Then

    thecorrectanswer = arra(13)

ElseIf question = 5 Then

    thecorrectanswer = arra(17)

ElseIf question = 6 Then
    
    thecorrectanswer = arra(21)

ElseIf question = 7 Then

    thecorrectanswer = arra(25)
    
ElseIf question = 8 Then

    thecorrectanswer = arra(29)
    
ElseIf question = 9 Then
    
    thecorrectanswer = arra(33)
    
ElseIf question = 10 Then
    
    thecorrectanswer = arra(37)
    
ElseIf question = 11 Then
    
        thecorrectanswer = arra(41)
    
ElseIf question = 12 Then

    thecorrectanswer = arra(45)
    
ElseIf question = 13 Then
    
    thecorrectanswer = arra(49)
    
ElseIf question = 14 Then

    thecorrectanswer = arra(53)

End If

MsgBox "Incorrect. The correct answer was: " + thecorrectanswer

payout = 0

GameOver

End Sub

Private Sub cmdExit_Click()

End

End Sub

Private Sub cmdPlay_Click()

Dim line As String
Dim i As Integer

cmdPlay.Enabled = False

tmrGame.Enabled = True
lblTimer = 60

cmdWalkAway.Enabled = True
cmdWalkAway.Visible = True

Open "D:\CP1\VB6 Project Files\GameShow\resources\questions.txt" For Input As #1

    For i = 1 To 14
    
        If EOF(1) Then
        
            Exit For
        
        End If
        
        Line Input #1, line
        
        arrq(i) = line
    
    Next i

Close #1

Open "D:\CP1\VB6 Project Files\GameShow\resources\answers.txt" For Input As #1

    For i = 1 To 56
    
        If EOF(1) Then
        
            Exit For
        
        End If
        
        Line Input #1, line
        
        arra(i) = line
    
    Next i

Close #1

SpawnQuestion

End Sub

Private Sub cmdWalkAway_Click()

GameOver

End Sub

Private Sub Form_Load()

timer = 60

question = 0

End Sub

Private Sub lblAns_Click(Index As Integer)

If question = 1 Then

    If lblAns(Index) = arra(1) Then
    
        CorrectAnswer
        
    Else
    
        IncorrectAnswer
        
    End If
    
ElseIf question = 2 Then

    If lblAns(Index) = arra(5) Then
    
        CorrectAnswer
        
    Else
    
        IncorrectAnswer
        
    End If
    
ElseIf question = 3 Then

    If lblAns(Index) = arra(9) Then
    
        CorrectAnswer
        
    Else
    
        IncorrectAnswer
        
    End If
    
ElseIf question = 4 Then

    If lblAns(Index) = arra(13) Then
    
        CorrectAnswer
        
    Else
    
        IncorrectAnswer
        
    End If
    
ElseIf question = 5 Then

    If lblAns(Index) = arra(17) Then
    
        CorrectAnswer
        
    Else
    
        IncorrectAnswer
        
    End If
    
ElseIf question = 6 Then

    If lblAns(Index) = arra(21) Then
    
        CorrectAnswer
        
    Else
    
        IncorrectAnswer
        
    End If

ElseIf question = 7 Then

    If lblAns(Index) = arra(25) Then
    
        CorrectAnswer
        
    Else
    
        IncorrectAnswer
        
    End If
    
ElseIf question = 8 Then

    If lblAns(Index) = arra(29) Then
    
        CorrectAnswer
        
    Else
    
        IncorrectAnswer
        
    End If
    
ElseIf question = 9 Then

    If lblAns(Index) = arra(33) Then
    
        CorrectAnswer
        
    Else
    
        IncorrectAnswer
        
    End If
    
ElseIf question = 10 Then

    If lblAns(Index) = arra(37) Then
    
        CorrectAnswer
        
    Else
    
        IncorrectAnswer
        
    End If
    
ElseIf question = 11 Then

    If lblAns(Index) = arra(41) Then
    
        CorrectAnswer
        
    Else
    
        IncorrectAnswer
        
    End If
    
ElseIf question = 12 Then

    If lblAns(Index) = arra(45) Then
    
        CorrectAnswer
        
    Else
    
        IncorrectAnswer
        
    End If
    
ElseIf question = 13 Then

    If lblAns(Index) = arra(49) Then
    
        CorrectAnswer
        
    Else
    
        IncorrectAnswer
        
    End If
    
ElseIf question = 14 Then

    If lblAns(Index) = arra(53) Then
    
        CorrectAnswer
        
    Else
    
        IncorrectAnswer
        
    End If

End If

End Sub

Private Sub tmrGame_Timer()

timer = timer - 1
lblTimer = timer

If timer = 0 Then

    IncorrectAnswer
    
End If

End Sub
