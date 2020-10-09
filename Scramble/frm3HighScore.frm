VERSION 5.00
Begin VB.Form frm3HighScore 
   Caption         =   "3-Letter Highscores"
   ClientHeight    =   3300
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4875
   LinkTopic       =   "Form1"
   ScaleHeight     =   3300
   ScaleWidth      =   4875
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstScores 
      Appearance      =   0  'Flat
      Height          =   1980
      ItemData        =   "frm3HighScore.frx":0000
      Left            =   1200
      List            =   "frm3HighScore.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   480
      Visible         =   0   'False
      Width           =   2000
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   120
      TabIndex        =   3
      Top             =   2760
      Width           =   4635
   End
   Begin VB.ListBox lstScoresReal 
      Appearance      =   0  'Flat
      Height          =   1980
      ItemData        =   "frm3HighScore.frx":0004
      Left            =   2760
      List            =   "frm3HighScore.frx":0006
      TabIndex        =   2
      Top             =   600
      Width           =   2000
   End
   Begin VB.ListBox lstNames 
      Appearance      =   0  'Flat
      Height          =   1980
      ItemData        =   "frm3HighScore.frx":0008
      Left            =   100
      List            =   "frm3HighScore.frx":000A
      TabIndex        =   1
      Top             =   600
      Width           =   2000
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "3-Letter Highscores"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   -60
      TabIndex        =   0
      Top             =   120
      Width           =   4995
   End
End
Attribute VB_Name = "frm3HighScore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim arrscores(1 To 10) As String
Dim arrnames(1 To 10) As String

Sub ReverseListBox()

lstScoresReal.Clear

lstScoresReal.AddItem lstScores.List(9)
lstScoresReal.AddItem lstScores.List(8)
lstScoresReal.AddItem lstScores.List(7)
lstScoresReal.AddItem lstScores.List(6)
lstScoresReal.AddItem lstScores.List(5)
lstScoresReal.AddItem lstScores.List(4)
lstScoresReal.AddItem lstScores.List(3)
lstScoresReal.AddItem lstScores.List(2)
lstScoresReal.AddItem lstScores.List(1)
lstScoresReal.AddItem lstScores.List(0)

End Sub

Private Sub cmdBack_Click()

Hide

frmHighScores.Show

End Sub

Private Sub Form_Load()

Dim i As Integer
Dim line As String
Dim donelooping As Boolean
Dim answer As String

lstNames.Clear

Open "D:\CP1\VB6 Project Files\Scramble\resources\highscores\3scores.txt" For Input As #1

    For i = 1 To 10
    
        If EOF(1) Then
        
            Exit For
            
        End If
        
        Line Input #1, line
        
        line = Trim(line)
        
        If Len(line) <> 0 Then
        
            arrscores(i) = line
            lstScores.AddItem Format(line, "@@@@@")
            
        End If
        
    Next i
    
Close #1

ReverseListBox

Open "D:\CP1\VB6 Project Files\Scramble\resources\highscores\3names.txt" For Input As #1

    For i = 1 To 10
    
        If EOF(1) Then
        
            Exit For
            
        End If
        
        Line Input #1, line
        
        line = Trim(line)
        
        If Len(line) <> 0 Then
        
            arrnames(i) = line
            lstNames.AddItem line
            
        End If
        
    Next i
    
Close #1

If Not globalscore = -1 Then

    answer = InputBox("Please enter your name:", "You Got A Highscore", "CoolGuy")

    If Format(globalscore, "@@@@@") >= lstScoresReal.List(9) Then
    
        arrscores(10) = Format(globalscore, "@@@@@")

        lstScores.Clear
        For i = 1 To 10
            
            lstScores.AddItem Format(arrscores(i), "@@@@@")

        Next i

        ReverseListBox

        For i = 1 To 10

            arrscores(i) = lstScoresReal.List(i - 1)

        Next i
        
        For i = 1 To 10
        
            If lstScoresReal.List(i - 1) = Format(globalscore, "@@@@@") And donelooping = False Then
            
                If i = 1 Then
                
                    arrnames(10) = arrnames(9)
                    arrnames(9) = arrnames(8)
                    arrnames(8) = arrnames(7)
                    arrnames(7) = arrnames(6)
                    arrnames(6) = arrnames(5)
                    arrnames(5) = arrnames(4)
                    arrnames(4) = arrnames(3)
                    arrnames(3) = arrnames(2)
                    arrnames(2) = arrnames(1)
                    arrnames(1) = answer
                    donelooping = True
                
                ElseIf i = 2 Then
                
                    arrnames(10) = arrnames(9)
                    arrnames(9) = arrnames(8)
                    arrnames(8) = arrnames(7)
                    arrnames(7) = arrnames(6)
                    arrnames(6) = arrnames(5)
                    arrnames(5) = arrnames(4)
                    arrnames(4) = arrnames(3)
                    arrnames(3) = arrnames(2)
                    arrnames(2) = answer
                    donelooping = True
                
                ElseIf i = 3 Then
                
                    arrnames(10) = arrnames(9)
                    arrnames(9) = arrnames(8)
                    arrnames(8) = arrnames(7)
                    arrnames(7) = arrnames(6)
                    arrnames(6) = arrnames(5)
                    arrnames(5) = arrnames(4)
                    arrnames(4) = arrnames(3)
                    arrnames(3) = answer
                    donelooping = True
                
                ElseIf i = 4 Then
                
                    arrnames(10) = arrnames(9)
                    arrnames(9) = arrnames(8)
                    arrnames(8) = arrnames(7)
                    arrnames(7) = arrnames(6)
                    arrnames(6) = arrnames(5)
                    arrnames(5) = arrnames(4)
                    arrnames(4) = answer
                    donelooping = True
                
                ElseIf i = 5 Then
                
                    arrnames(10) = arrnames(9)
                    arrnames(9) = arrnames(8)
                    arrnames(8) = arrnames(7)
                    arrnames(7) = arrnames(6)
                    arrnames(6) = arrnames(5)
                    arrnames(5) = answer
                    donelooping = True
                
                ElseIf i = 6 Then
                
                    arrnames(10) = arrnames(9)
                    arrnames(9) = arrnames(8)
                    arrnames(8) = arrnames(7)
                    arrnames(7) = arrnames(6)
                    arrnames(6) = answer
                    donelooping = True
                
                ElseIf i = 7 Then
                
                    arrnames(10) = arrnames(9)
                    arrnames(9) = arrnames(8)
                    arrnames(8) = arrnames(7)
                    arrnames(7) = answer
                    donelooping = True
                
                ElseIf i = 8 Then
                
                    arrnames(10) = arrnames(9)
                    arrnames(9) = arrnames(8)
                    arrnames(8) = answer
                    donelooping = True
                
                ElseIf i = 9 Then
                
                    arrnames(10) = arrnames(9)
                    arrnames(9) = answer
                    donelooping = True
                
                ElseIf i = 10 Then
                
                    arrnames(10) = answer
                    donelooping = True
                
                End If
                
                lstNames.Clear
                
            End If
            
        Next i
        
        For i = 1 To 10
                
            lstNames.AddItem arrnames(i)
                
        Next i
        
        'Write file
        
        Open "D:\CP1\VB6 Project Files\Scramble\resources\highscores\3scores.txt" For Output As #1
        
            For i = 1 To 10
            
                Print #1, arrscores(i)
            
            Next i
            
        Close #1
        
        Open "D:\CP1\VB6 Project Files\Scramble\resources\highscores\3names.txt" For Output As #1
        
            For i = 1 To 10
            
                Print #1, arrnames(i)
                
            Next i
            
        Close #1
        
    End If

End If

End Sub
