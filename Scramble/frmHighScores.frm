VERSION 5.00
Begin VB.Form frmHighScores 
   Caption         =   "High Scores"
   ClientHeight    =   4065
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5880
   LinkTopic       =   "Form1"
   ScaleHeight     =   4065
   ScaleWidth      =   5880
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Appearance      =   0  'Flat
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   1980
      TabIndex        =   5
      Top             =   3000
      Width           =   2000
   End
   Begin VB.CommandButton cmd7Letter 
      Appearance      =   0  'Flat
      Caption         =   "7-Letter"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   3960
      TabIndex        =   4
      Top             =   2340
      Width           =   2000
   End
   Begin VB.CommandButton cmd5Letter 
      Appearance      =   0  'Flat
      Caption         =   "5-Letter"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   1980
      TabIndex        =   3
      Top             =   2340
      Width           =   2000
   End
   Begin VB.CommandButton cmd3Letter 
      Appearance      =   0  'Flat
      Caption         =   "3-Letter"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   0
      TabIndex        =   2
      Top             =   2340
      Width           =   2000
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "S C O R E S"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1395
      Left            =   -60
      TabIndex        =   1
      Top             =   1020
      Width           =   6000
   End
   Begin VB.Label lblHighScoresTitle1 
      Alignment       =   2  'Center
      Caption         =   "H I G H"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1395
      Left            =   -60
      TabIndex        =   0
      Top             =   0
      Width           =   6000
   End
End
Attribute VB_Name = "frmHighScores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd3Letter_Click()

Hide

frm3HighScore.Show

End Sub

Private Sub cmd5Letter_Click()

Hide

frm5HighScore.Show

End Sub

Private Sub cmd7Letter_Click()

Hide

frm7HighScore.Show

End Sub

Private Sub cmdExit_Click()

Hide

frmScrambleMenu.Show

End Sub

Private Sub Form_Unload(Cancel As Integer)

Hide

frmScrambleMenu.Show

End Sub
