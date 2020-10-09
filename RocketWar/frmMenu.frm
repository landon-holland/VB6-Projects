VERSION 5.00
Begin VB.Form frmMenu 
   Caption         =   "Main Menu"
   ClientHeight    =   7590
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7875
   LinkTopic       =   "Form1"
   ScaleHeight     =   7590
   ScaleWidth      =   7875
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOptions 
      Appearance      =   0  'Flat
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   2460
      TabIndex        =   4
      Top             =   5040
      Width           =   2895
   End
   Begin VB.CommandButton cmdQuit 
      Appearance      =   0  'Flat
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   2460
      TabIndex        =   3
      Top             =   6360
      Width           =   2895
   End
   Begin VB.CommandButton cmdHighScores 
      Appearance      =   0  'Flat
      Caption         =   "High Scores"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   2460
      TabIndex        =   2
      Top             =   5700
      Width           =   2895
   End
   Begin VB.CommandButton cmdPlay 
      Appearance      =   0  'Flat
      Caption         =   "Play"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   2460
      TabIndex        =   1
      Top             =   4380
      Width           =   2895
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "Rocket War"
      BeginProperty Font 
         Name            =   "Poplar Std"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   45
      TabIndex        =   0
      Top             =   600
      Width           =   7755
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdHighScores_Click()

Hide
frmHighScores.Show

End Sub

Private Sub cmdOptions_Click()

Hide
frmOptions.Show

End Sub

Private Sub cmdPlay_Click()

Hide
frmGame.Show

End Sub

Private Sub cmdQuit_Click()

End

End Sub

Private Sub Form_Load()

MouseControl = True

globalscore = -1

End Sub

