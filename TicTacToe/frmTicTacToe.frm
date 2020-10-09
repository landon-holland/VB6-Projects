VERSION 5.00
Begin VB.Form frmTicTacToe 
   Caption         =   "Tic Tac Toe"
   ClientHeight    =   4800
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9330
   LinkTopic       =   "Form1"
   ScaleHeight     =   4800
   ScaleWidth      =   9330
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrAI 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5040
      Top             =   3300
   End
   Begin VB.Timer tmrWin 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   8700
      Top             =   4320
   End
   Begin VB.Timer tmrWhosTurn 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   240
      Top             =   120
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2640
      TabIndex        =   22
      Top             =   3840
      Width           =   2400
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   240
      TabIndex        =   21
      Top             =   3840
      Width           =   2400
   End
   Begin VB.CommandButton cmdAI 
      Caption         =   "AI"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2640
      TabIndex        =   20
      Top             =   3300
      Width           =   2400
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Play"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   240
      TabIndex        =   19
      Top             =   3300
      Width           =   2400
   End
   Begin VB.TextBox txtPlayer2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2640
      TabIndex        =   14
      Top             =   1680
      Width           =   2400
   End
   Begin VB.TextBox txtPlayer1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   240
      TabIndex        =   13
      Top             =   1680
      Width           =   2400
   End
   Begin VB.Line lin8 
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   8760
      X2              =   5520
      Y1              =   900
      Y2              =   4080
   End
   Begin VB.Line lin7 
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   5520
      X2              =   8760
      Y1              =   900
      Y2              =   4080
   End
   Begin VB.Line lin6 
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   5460
      X2              =   8820
      Y1              =   3780
      Y2              =   3780
   End
   Begin VB.Line lin5 
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   5460
      X2              =   8820
      Y1              =   2460
      Y2              =   2460
   End
   Begin VB.Line lin4 
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   5460
      X2              =   8820
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line lin3 
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   8460
      X2              =   8460
      Y1              =   720
      Y2              =   4200
   End
   Begin VB.Line lin2 
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   7140
      X2              =   7140
      Y1              =   720
      Y2              =   4200
   End
   Begin VB.Line lin1 
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   5820
      X2              =   5820
      Y1              =   720
      Y2              =   4200
   End
   Begin VB.Label lblWin 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   240
      TabIndex        =   23
      Top             =   4320
      Width           =   8955
   End
   Begin VB.Label lblPlayer2Wins 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
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
      Height          =   435
      Left            =   2640
      TabIndex        =   18
      Top             =   2700
      Width           =   2400
   End
   Begin VB.Label lblPlayer1Wins 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
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
      Height          =   435
      Left            =   240
      TabIndex        =   17
      Top             =   2700
      Width           =   2400
   End
   Begin VB.Label lblWinsPlayer2Prompt 
      Alignment       =   2  'Center
      Caption         =   "Wins"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2640
      TabIndex        =   16
      Top             =   2220
      Width           =   2400
   End
   Begin VB.Label lblWinsPlayer1Prompt 
      Alignment       =   2  'Center
      Caption         =   "Wins"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   240
      TabIndex        =   15
      Top             =   2220
      Width           =   2400
   End
   Begin VB.Label lblPlayer2Prompt 
      Alignment       =   2  'Center
      Caption         =   "Player 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2640
      TabIndex        =   12
      Top             =   1200
      Width           =   2400
   End
   Begin VB.Label lblPlayer1Prompt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "Player 1"
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
      Height          =   435
      Left            =   240
      TabIndex        =   11
      Top             =   1200
      Width           =   2400
   End
   Begin VB.Label lblWhosTurn 
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
      Height          =   435
      Left            =   240
      TabIndex        =   10
      Top             =   660
      Width           =   4815
   End
   Begin VB.Label lblWhosTurnPrompt 
      Alignment       =   2  'Center
      Caption         =   "Who's Turn?"
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
      TabIndex        =   9
      Top             =   240
      Width           =   4815
   End
   Begin VB.Line lin1H 
      BorderWidth     =   4
      X1              =   5280
      X2              =   9060
      Y1              =   3180
      Y2              =   3180
   End
   Begin VB.Line lin2H 
      BorderWidth     =   4
      X1              =   5280
      X2              =   9000
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line lin2V 
      BorderWidth     =   4
      X1              =   7800
      X2              =   7800
      Y1              =   600
      Y2              =   4440
   End
   Begin VB.Line lin1V 
      BorderWidth     =   4
      X1              =   6480
      X2              =   6480
      Y1              =   600
      Y2              =   4440
   End
   Begin VB.Label lbl9 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   48.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   7800
      TabIndex        =   8
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label lbl8 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   48.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   6540
      TabIndex        =   7
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label lbl7 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   48.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   5280
      TabIndex        =   6
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label lbl6 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   48.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   7800
      TabIndex        =   5
      Top             =   1860
      Width           =   1215
   End
   Begin VB.Label lbl5 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   48.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   6540
      TabIndex        =   4
      Top             =   1860
      Width           =   1215
   End
   Begin VB.Label lbl4 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   48.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   5280
      TabIndex        =   3
      Top             =   1860
      Width           =   1215
   End
   Begin VB.Label lbl3 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   48.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   7800
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label lbl2 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   48.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   6540
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   48.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   5280
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "frmTicTacToe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Flash As Integer
Dim Player1Wins As Integer
Dim Player2Wins As Integer
Dim Player1Name As String
Dim Player2Name As String
Dim GamePlaying As Boolean
Dim Player1Turn As Boolean
Dim lbl1Filled As Boolean
Dim lbl2Filled As Boolean
Dim lbl3Filled As Boolean
Dim lbl4Filled As Boolean
Dim lbl5Filled As Boolean
Dim lbl6Filled As Boolean
Dim lbl7Filled As Boolean
Dim lbl8Filled As Boolean
Dim lbl9Filled As Boolean
Dim LastWinner As Integer
Dim PlayingAI As Boolean

Sub Enable()
lbl1.Enabled = True
lbl2.Enabled = True
lbl3.Enabled = True
lbl4.Enabled = True
lbl5.Enabled = True
lbl6.Enabled = True
lbl7.Enabled = True
lbl8.Enabled = True
lbl9.Enabled = True
End Sub

Sub Disable()
lbl1.Enabled = False
lbl2.Enabled = False
lbl3.Enabled = False
lbl4.Enabled = False
lbl5.Enabled = False
lbl6.Enabled = False
lbl7.Enabled = False
lbl8.Enabled = False
lbl9.Enabled = False
End Sub

Sub CheckWin()
If lbl1.Caption = "X" And lbl2.Caption = "X" And lbl3.Caption = "X" Then
    Player1Win
    lin4.Visible = True
ElseIf lbl1.Caption = "O" And lbl2.Caption = "O" And lbl3.Caption = "O" Then
    Player2Win
    lin4.Visible = True
ElseIf lbl1.Caption = "X" And lbl4.Caption = "X" And lbl7.Caption = "X" Then
    Player1Win
    lin1.Visible = True
ElseIf lbl1.Caption = "O" And lbl4.Caption = "O" And lbl7.Caption = "O" Then
    Player2Win
    lin1.Visible = True
ElseIf lbl1.Caption = "X" And lbl5.Caption = "X" And lbl9.Caption = "X" Then
    Player1Win
    lin7.Visible = True
ElseIf lbl1.Caption = "O" And lbl5.Caption = "O" And lbl9.Caption = "O" Then
    Player2Win
    lin7.Visible = True
ElseIf lbl2.Caption = "X" And lbl5.Caption = "X" And lbl8.Caption = "X" Then
    Player1Win
    lin2.Visible = True
ElseIf lbl2.Caption = "O" And lbl5.Caption = "O" And lbl8.Caption = "O" Then
    Player2Win
    lin2.Visible = True
ElseIf lbl3.Caption = "X" And lbl6.Caption = "X" And lbl9.Caption = "X" Then
    Player1Win
    lin3.Visible = True
ElseIf lbl3.Caption = "O" And lbl6.Caption = "O" And lbl9.Caption = "O" Then
    Player2Win
    lin3.Visible = True
ElseIf lbl4.Caption = "X" And lbl5.Caption = "X" And lbl6.Caption = "X" Then
    Player1Win
    lin5.Visible = True
ElseIf lbl4.Caption = "O" And lbl5.Caption = "O" And lbl6.Caption = "O" Then
    Player2Win
    lin5.Visible = True
ElseIf lbl7.Caption = "X" And lbl8.Caption = "X" And lbl9.Caption = "X" Then
    Player1Win
    lin6.Visible = True
ElseIf lbl7.Caption = "O" And lbl8.Caption = "O" And lbl9.Caption = "O" Then
    Player2Win
    lin6.Visible = True
ElseIf lbl3.Caption = "X" And lbl5.Caption = "X" And lbl7.Caption = "X" Then
    Player1Win
    lin8.Visible = True
ElseIf lbl3.Caption = "O" And lbl5.Caption = "O" And lbl7.Caption = "O" Then
    Player2Win
    lin8.Visible = True
ElseIf lbl1Filled = True And lbl2Filled = True And lbl3Filled = True And lbl4Filled = True And lbl5Filled = True And lbl6Filled = True And lbl7Filled = True And lbl8Filled = True And lbl9Filled = True Then
    TieGame
End If
End Sub

Sub Player1Win()
Disable
tmrWhosTurn.Enabled = False
lblWhosTurn.BackColor = &H80000005
lblWhosTurn.ForeColor = &H80000008
lblWin.Caption = Player1Name + " wins!"
tmrWin.Enabled = True
GamePlaying = False
Player1Wins = Player1Wins + 1
lblPlayer1Wins.Caption = Player1Wins
LastWinner = 1
End Sub

Sub Player2Win()
Disable
tmrWhosTurn.Enabled = False
lblWhosTurn.BackColor = &H80000005
lblWhosTurn.ForeColor = &H80000008
lblWin.Caption = Player2Name + " wins!"
tmrWin.Enabled = True
GamePlaying = False
Player2Wins = Player2Wins + 1
lblPlayer2Wins.Caption = Player1Wins
LastWinner = 2
End Sub

Sub TieGame()
Disable
tmrWhosTurn.Enabled = False
lblWhosTurn.BackColor = &H80000005
lblWhosTurn.ForeColor = &H80000008
lblWin.Caption = "Tie Game!"
tmrWin.Enabled = True
GamePlaying = False
End Sub

Sub GameClear()
lbl1.Caption = ""
lbl2.Caption = ""
lbl3.Caption = ""
lbl4.Caption = ""
lbl5.Caption = ""
lbl6.Caption = ""
lbl7.Caption = ""
lbl8.Caption = ""
lbl9.Caption = ""
lbl1Filled = False
lbl2Filled = False
lbl3Filled = False
lbl4Filled = False
lbl5Filled = False
lbl6Filled = False
lbl7Filled = False
lbl8Filled = False
lbl9Filled = False
lin1.Visible = False
lin2.Visible = False
lin3.Visible = False
lin4.Visible = False
lin5.Visible = False
lin6.Visible = False
lin7.Visible = False
lin8.Visible = False
tmrWin.Enabled = False
tmrWhosTurn.Enabled = False
tmrAI.Enabled = False
GamePlaying = False
PlayingAI = False
lblWin.BackColor = &H8000000F
lblWin.ForeColor = &H80000008
lblWin.Caption = ""
lblWhosTurn.Caption = ""
End Sub

Sub MatchClear()
GameClear
Disable
Player1Name = ""
Player2Name = ""
Player1Wins = 0
Player2Wins = 0
lblPlayer1Wins = "0"
lblPlayer2Wins = "0"
txtPlayer1.Text = ""
txtPlayer2.Text = ""
txtPlayer1.SetFocus
LastWinner = 2
End Sub

Private Sub cmdAI_Click()
If GamePlaying = True Then
    MsgBox "We're already playing a game!"
Else
    Enable
    GameClear
    GamePlaying = True
    tmrAI.Enabled = True
    PlayingAI = True
    txtPlayer2.Text = "AI"
    Player2Name = "AI"
    If txtPlayer1.Text = "" Then
        Player1Name = "X"
        txtPlayer1.Text = "X"
    Else
        Player1Name = txtPlayer1.Text
    End If

    If txtPlayer2.Text = "" Then
        Player2Name = "O"
        txtPlayer2.Text = "O"
    Else
        Player2Name = txtPlayer2.Text
    End If
    
    If LastWinner = 2 Then
        lblWhosTurn.Caption = Player1Name
        Player1Turn = True
        tmrWhosTurn.Enabled = True
    ElseIf LastWinner = 1 Then
        lblWhosTurn.Caption = Player2Name
        Player1Turn = False
        tmrWhosTurn.Enabled = True
    End If
End If
End Sub

Private Sub cmdClear_Click()
MatchClear
End Sub

Private Sub cmdExit_Click()
End
End Sub

Private Sub cmdPlay_Click()
If GamePlaying = True Then
    MsgBox "We're already playing a game!"
Else
    tmrAI.Enabled = False
    PlayingAI = False
    Enable
    GameClear
    GamePlaying = True
    If txtPlayer1.Text = "" Then
        Player1Name = "X"
        txtPlayer1.Text = "X"
    Else
        Player1Name = txtPlayer1.Text
    End If

    If txtPlayer2.Text = "" Then
        Player2Name = "O"
        txtPlayer2.Text = "O"
    Else
        Player2Name = txtPlayer2.Text
    End If
    
    If LastWinner = 2 Then
        lblWhosTurn.Caption = Player1Name
        Player1Turn = True
        tmrWhosTurn.Enabled = True
    ElseIf LastWinner = 1 Then
        lblWhosTurn.Caption = Player2Name
        Player1Turn = False
        tmrWhosTurn.Enabled = True
    End If
End If
End Sub

Private Sub Form_Load()
Player1Wins = 0
Player2Wins = 0
lblPlayer1Wins.Caption = Player1Wins
lblPlayer2Wins.Caption = Player2Wins
LastWinner = 2
Randomize
End Sub

Private Sub lbl1_Click()
If lbl1Filled = False Then
    lbl1Filled = True
    If Player1Turn = True Then
        lbl1.Caption = "X"
        Player1Turn = False
        lblWhosTurn.Caption = Player2Name
        If PlayingAI = True Then
            tmrAI.Enabled = True
        End If
    Else
        lbl1.Caption = "O"
        Player1Turn = True
        lblWhosTurn.Caption = Player1Name
    End If
    CheckWin
Else
    MsgBox "This box is already filled."
End If
End Sub

Private Sub lbl2_Click()
If lbl2Filled = False Then
    lbl2Filled = True
    If Player1Turn = True Then
        lbl2.Caption = "X"
        Player1Turn = False
        lblWhosTurn.Caption = Player2Name
        If PlayingAI = True Then
            tmrAI.Enabled = True
        End If
    Else
        lbl2.Caption = "O"
        Player1Turn = True
        lblWhosTurn.Caption = Player1Name
    End If
    CheckWin
Else
    MsgBox "This box is already filled."
End If
End Sub

Private Sub lbl3_Click()
If lbl3Filled = False Then
    lbl3Filled = True
    If Player1Turn = True Then
        lbl3.Caption = "X"
        Player1Turn = False
        lblWhosTurn.Caption = Player2Name
        If PlayingAI = True Then
            tmrAI.Enabled = True
        End If
    Else
        lbl3.Caption = "O"
        Player1Turn = True
        lblWhosTurn.Caption = Player1Name
    End If
    CheckWin
Else
    MsgBox "This box is already filled."
End If
End Sub

Private Sub lbl4_Click()
If lbl4Filled = False Then
    lbl4Filled = True
    If Player1Turn = True Then
        lbl4.Caption = "X"
        Player1Turn = False
        lblWhosTurn.Caption = Player2Name
        If PlayingAI = True Then
            tmrAI.Enabled = True
        End If
    Else
        lbl4.Caption = "O"
        Player1Turn = True
        lblWhosTurn.Caption = Player1Name
    End If
    CheckWin
Else
    MsgBox "This box is already filled."
End If
End Sub

Private Sub lbl5_Click()
If lbl5Filled = False Then
    lbl5Filled = True
    If Player1Turn = True Then
        lbl5.Caption = "X"
        Player1Turn = False
        lblWhosTurn.Caption = Player2Name
        If PlayingAI = True Then
            tmrAI.Enabled = True
        End If
    Else
        lbl5.Caption = "O"
        Player1Turn = True
        lblWhosTurn.Caption = Player1Name
    End If
    CheckWin
Else
    MsgBox "This box is already filled."
End If
End Sub

Private Sub lbl6_Click()
If lbl6Filled = False Then
    lbl6Filled = True
    If Player1Turn = True Then
        lbl6.Caption = "X"
        Player1Turn = False
        lblWhosTurn.Caption = Player2Name
        If PlayingAI = True Then
            tmrAI.Enabled = True
        End If
    Else
        lbl6.Caption = "O"
        Player1Turn = True
        lblWhosTurn.Caption = Player1Name
    End If
    CheckWin
Else
    MsgBox "This box is already filled."
End If
End Sub

Private Sub lbl7_Click()
If lbl7Filled = False Then
    lbl7Filled = True
    If Player1Turn = True Then
        lbl7.Caption = "X"
        Player1Turn = False
        lblWhosTurn.Caption = Player2Name
        If PlayingAI = True Then
            tmrAI.Enabled = True
        End If
    Else
        lbl7.Caption = "O"
        Player1Turn = True
        lblWhosTurn.Caption = Player1Name
    End If
    CheckWin
Else
    MsgBox "This box is already filled."
End If
End Sub

Private Sub lbl8_Click()
If lbl8Filled = False Then
    lbl8Filled = True
    If Player1Turn = True Then
        lbl8.Caption = "X"
        Player1Turn = False
        lblWhosTurn.Caption = Player2Name
        If PlayingAI = True Then
            tmrAI.Enabled = True
        End If
    Else
        lbl8.Caption = "O"
        Player1Turn = True
        lblWhosTurn.Caption = Player1Name
    End If
    CheckWin
Else
    MsgBox "This box is already filled."
End If
End Sub

Private Sub lbl9_Click()
If lbl9Filled = False Then
    lbl9Filled = True
    If Player1Turn = True Then
        lbl9.Caption = "X"
        Player1Turn = False
        lblWhosTurn.Caption = Player2Name
        If PlayingAI = True Then
            tmrAI.Enabled = True
        End If
    Else
        lbl9.Caption = "O"
        Player1Turn = True
        lblWhosTurn.Caption = Player1Name
    End If
    CheckWin
Else
    MsgBox "This box is already filled."
End If
End Sub

Private Sub tmrAI_Timer()
Dim RNG As Integer
If Player1Turn = False And GamePlaying = True Then
    RNG = Int((9 - 1 + 1) * Rnd + 1)
    If RNG = 1 Then
        If lbl1Filled = False Then
            lbl1_Click
            tmrAI.Enabled = False
        End If
    ElseIf RNG = 2 Then
        If lbl2Filled = False Then
            lbl2_Click
            tmrAI.Enabled = False
        End If
    ElseIf RNG = 3 Then
        If lbl3Filled = False Then
            lbl3_Click
            tmrAI.Enabled = False
        End If
    ElseIf RNG = 4 Then
        If lbl4Filled = False Then
            lbl4_Click
            tmrAI.Enabled = False
        End If
    ElseIf RNG = 5 Then
        If lbl5Filled = False Then
            lbl5_Click
            tmrAI.Enabled = False
        End If
    ElseIf RNG = 6 Then
        If lbl6Filled = False Then
            lbl6_Click
            tmrAI.Enabled = False
        End If
    ElseIf RNG = 7 Then
        If lbl7Filled = False Then
            lbl7_Click
            tmrAI.Enabled = False
        End If
    ElseIf RNG = 8 Then
        If lbl8Filled = False Then
            lbl8_Click
            tmrAI.Enabled = False
        End If
    ElseIf RNG = 9 Then
        If lbl9Filled = False Then
            lbl9_Click
            tmrAI.Enabled = False
        End If
    End If
End If
End Sub

Private Sub tmrWhosTurn_Timer()
Dim Flash As Integer
Flash = Second(Now) Mod 2
If Flash = 0 Then
    lblWhosTurn.BackColor = &H80000008
    lblWhosTurn.ForeColor = &H80000005
Else
    lblWhosTurn.BackColor = &H80000005
    lblWhosTurn.ForeColor = &H80000008
End If
End Sub

Private Sub tmrWin_Timer()
Dim Flash As Integer
Flash = Second(Now) Mod 2
If Flash = 0 Then
    lblWin.BackColor = &H80000008
    lblWin.ForeColor = &H80000005
Else
    lblWin.BackColor = &H8000000F
    lblWin.ForeColor = &H80000008
End If
End Sub

Private Sub txtPlayer1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtPlayer2.SetFocus
End If
End Sub

Private Sub txtPlayer2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdPlay_Click
End If
End Sub
