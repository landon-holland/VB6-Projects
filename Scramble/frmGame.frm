VERSION 5.00
Begin VB.Form frmGame 
   Caption         =   "-Letter Word"
   ClientHeight    =   4485
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7095
   LinkTopic       =   "Form1"
   ScaleHeight     =   4485
   ScaleWidth      =   7095
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Appearance      =   0  'Flat
      Caption         =   "Exit"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   -60
      TabIndex        =   11
      Top             =   2940
      Visible         =   0   'False
      Width           =   2000
   End
   Begin VB.Timer tmrTimeLeft 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin VB.Label lblGameOver 
      Alignment       =   2  'Center
      Caption         =   "Game Over"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   0
      TabIndex        =   10
      Top             =   1500
      Visible         =   0   'False
      Width           =   5835
   End
   Begin VB.Label lblScore 
      Caption         =   "Score: 0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   4080
      Width           =   3435
   End
   Begin VB.Label lblGuessWord 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   0
      TabIndex        =   8
      Top             =   2820
      Width           =   3255
   End
   Begin VB.Label lblTimer 
      Caption         =   "Time Left: "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   4455
   End
   Begin VB.Label lblLetter 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Gadugi"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Index           =   6
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   1005
   End
   Begin VB.Label lblLetter 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Gadugi"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Index           =   5
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   1005
   End
   Begin VB.Label lblLetter 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Gadugi"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Index           =   4
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   1005
   End
   Begin VB.Label lblLetter 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Gadugi"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Index           =   3
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   1005
   End
   Begin VB.Label lblLetter 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Gadugi"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Index           =   2
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   1005
   End
   Begin VB.Label lblLetter 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Gadugi"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Index           =   1
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1005
   End
   Begin VB.Label lblLetter 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Gadugi"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1005
   End
End
Attribute VB_Name = "frmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim chosenword As String
Dim guessword As String
Dim lettersclicked As Integer
Dim score As Long
Dim timeleft As Integer
Dim arrletters(1 To 100) As String

Public Function RandomNumber(ByVal MaxValue As Long, Optional _
ByVal MinValue As Long = 0)

  On Error Resume Next
  Randomize
  RandomNumber = Int((MaxValue - MinValue + 1) * Rnd) + MinValue

End Function

Sub GenerateWord(numberofletters As Integer)

Dim i As Long
Dim j As Long
Dim randomnumbergenerated As Long
Dim line As String

j = 0

Open "D:\CP1\VB6 Project Files\Scramble\resources\dictionary.txt" For Input As #1

    For i = 1 To 58112
    
        If EOF(1) Then
        
            Exit For
            
        End If
    
        Line Input #1, line
        
        line = Trim(line)
        
        If Len(line) <> 0 And Len(line) = numberofletters Then
        
            j = j + 1
        
            worddictionary(j) = line
        
        End If
        
    Next i

Close #1

Do While Len(chosenword) <> numberofletters

    randomnumbergenerated = RandomNumber(58112, 1)
    
    chosenword = worddictionary(randomnumbergenerated)

Loop

End Sub

Sub Shuffle(numberofletters As Integer)

Dim i As Integer
Dim temp As String
Dim rng As Integer

For i = 1 To numberofletters

    rng = RandomNumber(numberofletters, 1)
    
    temp = arrletters(i)
    arrletters(i) = arrletters(rng)
    arrletters(rng) = temp

Next i

End Sub

Sub CheckWin()

Dim i As Long
Dim j As Integer
Dim donotcheck As Boolean

For i = 1 To 58112

    donotcheck = False

    If guessword = worddictionary(i) Then
    
        MsgBox "Correct!"
        
        chosenword = ""
        guessword = ""
        lblGuessWord = ""
        lettersclicked = 0
        
        score = score + 100
        lblScore = "Score: " + Str(score)
        
        If gametype = 1 Then
        
            GenerateWord (3)
            
            For j = 1 To 3
    
                arrletters(j) = Mid(chosenword, j, 1)
    
            Next j
            
            Shuffle (3)
            
            For j = 1 To 3
            
                lblLetter(j - 1).Enabled = True
                lblLetter(j - 1).Visible = True
    
                lblLetter(j - 1) = arrletters(j)
    
            Next j
            
        ElseIf gametype = 2 Then
        
            GenerateWord (5)
            
            For j = 1 To 5
    
                arrletters(j) = Mid(chosenword, j, 1)
    
            Next j
            
            Shuffle (5)
            
            For j = 1 To 5
            
                lblLetter(j - 1).Enabled = True
                lblLetter(j - 1).Visible = True
                
                lblLetter(j - 1) = arrletters(j)
    
            Next j
            
        ElseIf gametype = 3 Then
        
            GenerateWord (7)
            
            For j = 1 To 7
    
                arrletters(j) = Mid(chosenword, j, 1)
    
            Next j
            
            Shuffle (7)
            
            For j = 1 To 7
            
                lblLetter(j - 1).Enabled = True
                lblLetter(j - 1).Visible = True
                
                lblLetter(j - 1) = arrletters(j)
    
            Next j
            
        End If
        
        donotcheck = True
        Exit For
    
    End If

Next i

If donotcheck = False Then

    MsgBox "Incorrect."
    
    guessword = ""
    lblGuessWord = ""
    
    lettersclicked = 0
    
    score = score - 100
    lblScore = "Score: " + Str(score)
    
    If gametype = 1 Then
    
        For i = 1 To 3
        
            lblLetter(i - 1).Enabled = True
            lblLetter(i - 1).Visible = True
        
        Next i

    ElseIf gametype = 2 Then
    
        For i = 1 To 5
        
            lblLetter(i - 1).Enabled = True
            lblLetter(i - 1).Visible = True
            
        Next i
        
    ElseIf gametype = 3 Then
    
        For i = 1 To 7
        
            lblLetter(i - 1).Enabled = True
            lblLetter(i - 1).Visible = True
        
        Next i
    
    End If
    
End If

End Sub

Private Sub cmdExit_Click()

globalscore = score

If gametype = 1 Then

    Hide
    
    frm3HighScore.Show

ElseIf gametype = 2 Then

    Hide
    
    frm5HighScore.Show
    
ElseIf gametype = 3 Then
    
    Hide
    
    frm7HighScore.Show
    
End If

End Sub

Private Sub Form_Activate()

Dim i As Integer

lettersclicked = 0
score = 0
chosenword = ""
guessword = ""
lblGuessWord = ""

If gametype = 1 Then

    Caption = "3-Letter Word"
    
    Width = 5000
    lblGuessWord.Width = 5000
    lblGameOver.Width = 5000
    lblGameOver.FontSize = 36
    cmdExit.Width = 5000
    
    lblLetter(0).Move 1000, 1000
    lblLetter(1).Move 2000, 1000
    lblLetter(2).Move 3000, 1000
    
    lblLetter(3).Visible = False
    lblLetter(4).Visible = False
    lblLetter(5).Visible = False
    lblLetter(6).Visible = False
    
    timeleft = 60
    lblTimer = "Time Left: 60"
    
    GenerateWord (3)
    
    For i = 1 To 3
    
        arrletters(i) = Mid(chosenword, i, 1)
    
    Next i
    
    Shuffle (3)
    
    For i = 1 To 3
    
        lblLetter(i - 1) = arrletters(i)
    
    Next i

ElseIf gametype = 2 Then

    Caption = "5-Letter Word"
    
    Width = 7000
    lblGuessWord.Width = 7000
    lblGameOver.Width = 7000
    cmdExit.Width = 7000
    
    lblLetter(0).Move 1000, 1000
    lblLetter(1).Move 2000, 1000
    lblLetter(2).Move 3000, 1000
    lblLetter(3).Move 4000, 1000
    lblLetter(4).Move 5000, 1000
    
    lblLetter(5).Visible = False
    lblLetter(6).Visible = False
    
    timeleft = 120
    lblTimer = "Time Left: 120"
    
    GenerateWord (5)
    
    For i = 1 To 5
    
        arrletters(i) = Mid(chosenword, i, 1)
    
    Next i

    Shuffle (5)
    
    For i = 1 To 5
    
        lblLetter(i - 1) = arrletters(i)
    
    Next i
    
ElseIf gametype = 3 Then

    Caption = "7-Letter Word"
    
    Width = 9000
    lblGuessWord.Width = 9000
    lblGameOver.Width = 9000
    cmdExit.Width = 9000
    
    lblLetter(0).Move 1000, 1000
    lblLetter(1).Move 2000, 1000
    lblLetter(2).Move 3000, 1000
    lblLetter(3).Move 4000, 1000
    lblLetter(4).Move 5000, 1000
    lblLetter(5).Move 6000, 1000
    lblLetter(6).Move 7000, 1000
    
    timeleft = 180
    lblTimer = "Time Left: 180"
    
    GenerateWord (7)
    
    For i = 1 To 7
    
        arrletters(i) = Mid(chosenword, i, 1)
    
    Next i
    
    Shuffle (7)
    
    For i = 1 To 7
    
        lblLetter(i - 1) = arrletters(i)
    
    Next i
    
End If

tmrTimeLeft.Enabled = True

End Sub

Private Sub Form_Unload(Cancel As Integer)

Hide

frmScrambleMenu.Show

End Sub

Private Sub lblLetter_Click(Index As Integer)

lblLetter(Index).Enabled = False
lblLetter(Index).Visible = False

lettersclicked = lettersclicked + 1

guessword = guessword + lblLetter(Index)
lblGuessWord = guessword

If gametype = 1 And lettersclicked = 3 Then

    CheckWin
    
ElseIf gametype = 2 And lettersclicked = 5 Then

    CheckWin

ElseIf gametype = 3 And lettersclicked = 7 Then

    CheckWin

End If

End Sub

Private Sub tmrTimeLeft_Timer()

timeleft = timeleft - 1
lblTimer = "Time Left: " + Str(timeleft)

If timeleft = 0 Then

    Dim i As Integer
    
    If gametype = 1 Then
    
        For i = 1 To 3
        
            lblLetter(i - 1).Enabled = False
            lblLetter(i - 1).Visible = False
            lblGuessWord.Visible = False
        
        Next i
        
    ElseIf gametype = 2 Then
    
        For i = 1 To 5
        
            lblLetter(i - 1).Enabled = False
            lblLetter(i - 1).Visible = False
            lblGuessWord.Visible = False
        
        Next i
        
    ElseIf gametype = 3 Then
    
        For i = 1 To 7
        
            lblLetter(i - 1).Enabled = False
            lblLetter(i - 1).Visible = False
            lblGuessWord.Visible = False
        
        Next i
    
    End If
    
    lblGameOver.Visible = True
    cmdExit.Enabled = True
    cmdExit.Visible = True
    cmdExit.SetFocus
    tmrTimeLeft.Enabled = False

End If

End Sub
