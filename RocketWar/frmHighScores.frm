VERSION 5.00
Begin VB.Form frmHighScores 
   Caption         =   "High Scores"
   ClientHeight    =   4245
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4605
   LinkTopic       =   "Form1"
   ScaleHeight     =   4245
   ScaleWidth      =   4605
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRefresh 
      Height          =   264
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   328
   End
   Begin VB.ListBox lstScoresReal 
      Appearance      =   0  'Flat
      Height          =   1980
      ItemData        =   "frmHighScores.frx":0000
      Left            =   2820
      List            =   "frmHighScores.frx":0002
      TabIndex        =   4
      Top             =   1200
      Width           =   1695
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   675
      Left            =   60
      TabIndex        =   3
      Top             =   3360
      Width           =   4455
   End
   Begin VB.ListBox lstScores 
      Appearance      =   0  'Flat
      Height          =   1785
      ItemData        =   "frmHighScores.frx":0004
      Left            =   1560
      List            =   "frmHighScores.frx":0006
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   1140
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.ListBox lstNames 
      Appearance      =   0  'Flat
      Height          =   1980
      Left            =   60
      TabIndex        =   1
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label lblHighScoresTitle 
      Alignment       =   2  'Center
      Caption         =   "High Scores"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   -60
      TabIndex        =   0
      Top             =   180
      Width           =   4725
   End
End
Attribute VB_Name = "frmHighScores"
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

Private Sub cmdExit_Click()

lstScores.Clear
lstScoresReal.Clear
lstNames.Clear

Hide

frmMenu.Show

End Sub

Private Sub cmdRefresh_Click()

Form_Load

End Sub

Private Sub Form_Load()

Dim answer As String
Dim i As Integer
Dim line As String
Dim donelooping As Boolean

Open "D:\CP1\VB6 Project Files\RocketWar\resources\highscores\scores.txt" For Input As #1

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

Open "D:\CP1\VB6 Project Files\RocketWar\resources\highscores\names.txt" For Input As #1

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
        
        Open "D:\CP1\VB6 Project Files\RocketWar\resources\highscores\scores.txt" For Output As #1
        
            For i = 1 To 10
            
                Print #1, arrscores(i)
            
            Next i
            
        Close #1
        
        Open "D:\CP1\VB6 Project Files\RocketWar\resources\highscores\names.txt" For Output As #1
        
            For i = 1 To 10
            
                Print #1, arrnames(i)
                
            Next i
            
        Close #1
        
    End If

End If

End Sub

Private Sub lbl1Score_Click()

End Sub

