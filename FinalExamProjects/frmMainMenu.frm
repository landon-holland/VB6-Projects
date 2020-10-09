VERSION 5.00
Begin VB.Form frmMainMenu 
   Caption         =   "Final Exam Projects"
   ClientHeight    =   3165
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3165
   ScaleWidth      =   9000
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Appearance      =   0  'Flat
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3000
      TabIndex        =   7
      Top             =   2340
      Width           =   3000
   End
   Begin VB.CommandButton cmd6 
      Appearance      =   0  'Flat
      Caption         =   "6. Palindrome Checker"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6000
      TabIndex        =   6
      Top             =   1740
      Width           =   3000
   End
   Begin VB.CommandButton cmd5 
      Appearance      =   0  'Flat
      Caption         =   "5. Shooter"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3000
      TabIndex        =   5
      Top             =   1740
      Width           =   3000
   End
   Begin VB.CommandButton cmd4 
      Appearance      =   0  'Flat
      Caption         =   "4. Quadratic Grapher"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   4
      Top             =   1740
      Width           =   3000
   End
   Begin VB.CommandButton cmd3 
      Appearance      =   0  'Flat
      Caption         =   "3. Letter Wizard"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6000
      TabIndex        =   3
      Top             =   1140
      Width           =   3000
   End
   Begin VB.CommandButton cmd2 
      Appearance      =   0  'Flat
      Caption         =   "2. Dive Judge"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3000
      TabIndex        =   2
      Top             =   1140
      Width           =   3000
   End
   Begin VB.CommandButton cmd1 
      Appearance      =   0  'Flat
      Caption         =   "1. Class Fees"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Top             =   1140
      Width           =   3000
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "Main Menu"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   26.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   -60
      TabIndex        =   0
      Top             =   300
      Width           =   9000
   End
End
Attribute VB_Name = "frmMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd1_Click()

Hide

frmClassFees.Show

End Sub

Private Sub cmd2_Click()

Hide

frmDiveJudge.Show

End Sub

Private Sub cmd3_Click()

Hide

frmLetterWizard.Show

End Sub

Private Sub cmd4_Click()

Hide

frmQuadraticGrapher.Show

End Sub

Private Sub cmd5_Click()

Hide

frmShooter.Show

End Sub

Private Sub cmd6_Click()

Hide

frmPalindromeChecker.Show

End Sub

Private Sub cmdExit_Click()

End

End Sub

