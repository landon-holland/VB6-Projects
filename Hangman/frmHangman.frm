VERSION 5.00
Begin VB.Form frmHangman 
   AutoRedraw      =   -1  'True
   Caption         =   "Hangman"
   ClientHeight    =   4230
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6630
   LinkTopic       =   "Form1"
   ScaleHeight     =   4230
   ScaleWidth      =   6630
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrAnimation 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   6180
      Top             =   3300
   End
   Begin VB.Timer tmrAI 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6180
      Top             =   3780
   End
   Begin VB.TextBox txtInput 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   0
      TabIndex        =   31
      Top             =   3900
      Width           =   315
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Play"
      Height          =   255
      Left            =   4860
      TabIndex        =   3
      Top             =   60
      Width           =   855
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   255
      Left            =   5760
      TabIndex        =   2
      Top             =   60
      Width           =   855
   End
   Begin VB.Image imgWinner 
      Height          =   1470
      Left            =   1920
      Picture         =   "frmHangman.frx":0000
      Stretch         =   -1  'True
      Top             =   420
      Visible         =   0   'False
      Width           =   2760
   End
   Begin VB.Image imgFace 
      Height          =   660
      Left            =   960
      Picture         =   "frmHangman.frx":0C7E
      Stretch         =   -1  'True
      Top             =   660
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Line lnBody4 
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   1320
      X2              =   1020
      Y1              =   1500
      Y2              =   1440
   End
   Begin VB.Line lnBody5 
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   1320
      X2              =   1620
      Y1              =   1500
      Y2              =   1440
   End
   Begin VB.Line lnBody3 
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   1320
      X2              =   1440
      Y1              =   1860
      Y2              =   1980
   End
   Begin VB.Line lnBody2 
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   1320
      X2              =   1200
      Y1              =   1860
      Y2              =   1980
   End
   Begin VB.Line lnBody1 
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   1320
      X2              =   1320
      Y1              =   1260
      Y2              =   1860
   End
   Begin VB.Shape shpCircle 
      BorderWidth     =   2
      Height          =   675
      Left            =   1020
      Shape           =   3  'Circle
      Top             =   660
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblKeyboard 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "m"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   25
      Left            =   4500
      TabIndex        =   30
      Tag             =   "m"
      Top             =   3660
      Width           =   405
   End
   Begin VB.Label lblKeyboard 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "n"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   24
      Left            =   4020
      TabIndex        =   29
      Tag             =   "n"
      Top             =   3660
      Width           =   405
   End
   Begin VB.Label lblKeyboard 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "b"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   23
      Left            =   3540
      TabIndex        =   28
      Tag             =   "b"
      Top             =   3660
      Width           =   405
   End
   Begin VB.Label lblKeyboard 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "v"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   22
      Left            =   3060
      TabIndex        =   27
      Tag             =   "v"
      Top             =   3660
      Width           =   405
   End
   Begin VB.Label lblKeyboard 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "c"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   21
      Left            =   2580
      TabIndex        =   26
      Tag             =   "c"
      Top             =   3660
      Width           =   405
   End
   Begin VB.Label lblKeyboard 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "x"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   20
      Left            =   2100
      TabIndex        =   25
      Tag             =   "x"
      Top             =   3660
      Width           =   405
   End
   Begin VB.Label lblKeyboard 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "z"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   19
      Left            =   1620
      TabIndex        =   24
      Tag             =   "z"
      Top             =   3660
      Width           =   405
   End
   Begin VB.Label lblKeyboard 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "l"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   18
      Left            =   4980
      TabIndex        =   23
      Tag             =   "l"
      Top             =   3240
      Width           =   405
   End
   Begin VB.Label lblKeyboard 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "k"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   17
      Left            =   4500
      TabIndex        =   22
      Tag             =   "k"
      Top             =   3240
      Width           =   405
   End
   Begin VB.Label lblKeyboard 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "j"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   16
      Left            =   4020
      TabIndex        =   21
      Tag             =   "j"
      Top             =   3240
      Width           =   405
   End
   Begin VB.Label lblKeyboard 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "h"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   15
      Left            =   3540
      TabIndex        =   20
      Tag             =   "h"
      Top             =   3240
      Width           =   405
   End
   Begin VB.Label lblKeyboard 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "g"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   14
      Left            =   3060
      TabIndex        =   19
      Tag             =   "g"
      Top             =   3240
      Width           =   405
   End
   Begin VB.Label lblKeyboard 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "f"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   13
      Left            =   2580
      TabIndex        =   18
      Tag             =   "f"
      Top             =   3240
      Width           =   405
   End
   Begin VB.Label lblKeyboard 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "d"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   12
      Left            =   2100
      TabIndex        =   17
      Tag             =   "d"
      Top             =   3240
      Width           =   405
   End
   Begin VB.Label lblKeyboard 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "s"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   11
      Left            =   1620
      TabIndex        =   16
      Tag             =   "s"
      Top             =   3240
      Width           =   405
   End
   Begin VB.Label lblKeyboard 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "a"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   10
      Left            =   1140
      TabIndex        =   15
      Tag             =   "a"
      Top             =   3240
      Width           =   405
   End
   Begin VB.Label lblKeyboard 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "p"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   9
      Left            =   5280
      TabIndex        =   14
      Tag             =   "p"
      Top             =   2820
      Width           =   405
   End
   Begin VB.Label lblKeyboard 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "o"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   8
      Left            =   4800
      TabIndex        =   13
      Tag             =   "o"
      Top             =   2820
      Width           =   405
   End
   Begin VB.Label lblKeyboard 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "i"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   7
      Left            =   4320
      TabIndex        =   12
      Tag             =   "i"
      Top             =   2820
      Width           =   405
   End
   Begin VB.Label lblKeyboard 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "u"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   6
      Left            =   3840
      TabIndex        =   11
      Tag             =   "u"
      Top             =   2820
      Width           =   405
   End
   Begin VB.Label lblKeyboard 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "y"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   5
      Left            =   3360
      TabIndex        =   10
      Tag             =   "y"
      Top             =   2820
      Width           =   405
   End
   Begin VB.Label lblKeyboard 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "t"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   4
      Left            =   2880
      TabIndex        =   9
      Tag             =   "t"
      Top             =   2820
      Width           =   405
   End
   Begin VB.Label lblKeyboard 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "r"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   3
      Left            =   2400
      TabIndex        =   8
      Tag             =   "r"
      Top             =   2820
      Width           =   405
   End
   Begin VB.Label lblKeyboard 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "e"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   2
      Left            =   1920
      TabIndex        =   7
      Tag             =   "e"
      Top             =   2820
      Width           =   405
   End
   Begin VB.Label lblKeyboard 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "w"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   1
      Left            =   1440
      TabIndex        =   6
      Tag             =   "w"
      Top             =   2820
      Width           =   405
   End
   Begin VB.Label lblKeyboard 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "q"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   0
      Left            =   960
      TabIndex        =   5
      Tag             =   "q"
      Top             =   2820
      Width           =   405
   End
   Begin VB.Label lblLoading 
      Alignment       =   2  'Center
      Caption         =   "Loading..."
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
      Left            =   4860
      TabIndex        =   4
      Top             =   360
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.Label lblWord 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   1440
      TabIndex        =   1
      Top             =   1380
      Width           =   5115
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   1320
      X2              =   1320
      Y1              =   660
      Y2              =   480
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   1320
      X2              =   660
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   660
      X2              =   660
      Y1              =   2160
      Y2              =   480
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   0
      X2              =   1380
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "Hangman"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6615
   End
   Begin VB.Menu mnuSettings 
      Caption         =   "Settings"
      Begin VB.Menu mnuShort 
         Caption         =   "Short Words (3 Letters)"
      End
      Begin VB.Menu mnuLong 
         Caption         =   "Long Words (7 and up)"
      End
      Begin VB.Menu mnuAI 
         Caption         =   "AI"
      End
      Begin VB.Menu mnuRegular 
         Caption         =   "Regular"
      End
   End
End
Attribute VB_Name = "frmHangman"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim arrwords(1 To 58112) As String
Dim arrshort(1 To 637) As String
Dim arrlong(1 To 43978) As String
Dim arrletters(1 To 22) As String
Dim arrwsf(1 To 22) As String
Dim arrchosen(0 To 25) As Boolean
Dim arrai(0 To 25) As Boolean
Dim lives As Integer
Dim length As Integer
Dim wsf As String
Dim wsf2 As String
Dim word As String
Dim enabledk As Boolean
Dim gametype As Integer
Dim userword As String
Dim aicounter As Long
Dim gotvowel As Boolean
Dim tai As Long
     
Private Const SND_APPLICATION As Long = &H80
Private Const SND_ALIAS As Long = &H10000
Private Const SND_ALIAS_ID As Long = &H110000
Private Const SND_ASYNC As Long = &H1
Private Const SND_FILENAME As Long = &H20000
Private Const SND_LOOP As Long = &H8
Private Const SND_MEMORY As Long = &H4
Private Const SND_NODEFAULT As Long = &H2
Private Const SND_NOSTOP As Long = &H10
Private Const SND_NOWAIT As Long = &H2000
Private Const SND_PURGE As Long = &H40
Private Const SND_RESOURCE As Long = &H40004
Private Const SND_SYNC As Long = &H0

Sub displayWord()

Dim i As Integer

wsf = ""
For i = 1 To length
    wsf = wsf + arrwsf(i) + " "
Next i
lblWord = wsf

wsf2 = ""
For i = 1 To length
    wsf2 = wsf2 + arrwsf(i)
Next i

End Sub

Sub Enable()

Dim i As Integer

For i = 0 To 25
    lblKeyboard(i).Enabled = True
Next i

enabledk = True

End Sub

Sub Disable()

Dim i As Integer

For i = 0 To 25
    lblKeyboard(i).Enabled = False
Next i

enabledk = False

End Sub

Sub Reset()

Dim i As Integer

gotvowel = False

For i = 0 To 25
    lblKeyboard(i) = lblKeyboard(i).Tag
Next i

For i = 0 To 25
    arrchosen(i) = False
Next i

tmrAnimation.Enabled = False

End Sub

Private Sub cmdExit_Click()

End

End Sub

Private Sub cmdPlay_Click()

Dim i As Long
Dim rng As Long

shpCircle.Visible = False
imgFace.Visible = False
lnBody1.Visible = False
lnBody2.Visible = False
lnBody3.Visible = False
lnBody4.Visible = False
lnBody5.Visible = False
imgWinner.Visible = False

lives = 6

Reset

Randomize
If gametype = 0 Then
    rng = Int(58112 * Rnd + 1)
    word = arrwords(rng)
ElseIf gametype = 1 Then
    rng = Int(637 * Rnd + 1)
    word = arrshort(rng)
ElseIf gametype = 2 Then
    rng = Int(43978 * Rnd + 1)
    word = arrlong(rng)
ElseIf gametype = 3 Then
    word = userword
    For i = 0 To 25
        arrchosen(i) = True
        lblKeyboard(i).Enabled = False
    Next i
    tmrAI.Enabled = True
End If
length = Len(word)
For i = 1 To length
    arrletters(i) = Mid(word, i, 1)
    arrwsf(i) = "-"
Next i

Enable

displayWord

End Sub

Private Sub Form_Load()

Dim path As String
Dim i As Long
Dim j As Long
Dim k As Long
Dim line As String

lblLoading.Visible = True

gametype = 0

j = 0
k = 0

aicounter = 0

path = "D:\CP2\VB6 Project Files\Hangman\resources\dictionary.txt"

Open path For Input As #1
    For i = 1 To 58112
        Line Input #1, line
        line = Trim(line)
        arrwords(i) = line
        If Len(line) = 3 Then
            j = j + 1
            arrshort(j) = line
        ElseIf Len(line) >= 7 Then
            k = k + 1
            arrlong(k) = line
        End If
    Next i
Close #1

lblLoading.Visible = False

End Sub

Private Sub lblKeyboard_Click(Index As Integer)

Dim i As Integer
Dim letterchosen As String
Dim found As Boolean

letterchosen = lblKeyboard(Index)
found = False

If arrchosen(Index) = False Or gametype = 3 Then
    
    arrchosen(Index) = True
    
    For i = 1 To length
        If letterchosen = arrletters(i) Then
            arrwsf(i) = letterchosen
            lblKeyboard(Index) = "-"
            lblKeyboard(Index).Enabled = False
            found = True
        
            displayWord
        
            If wsf2 = word Then
                PlaySound "D:\CP2\VB6 Project Files\Hangman\resources\victory.wav", 0, SND_FILENAME Or SND_ASYNC
                If gametype = 3 Then
                    MsgBox "The AI got the word. You lose!"
                Else
                    MsgBox "You win!"
                    imgWinner.Visible = True
                    tmrAnimation.Enabled = True
                End If
                Disable
            End If
        End If
    Next i

    If found = False Then
        lives = lives - 1
        lblKeyboard(Index) = "-"
        lblKeyboard(Index).Enabled = False
    
        If lives = 5 Then
            shpCircle.Visible = True
            imgFace.Visible = True
        ElseIf lives = 4 Then
            lnBody1.Visible = True
        ElseIf lives = 3 Then
            lnBody2.Visible = True
        ElseIf lives = 2 Then
            lnBody3.Visible = True
        ElseIf lives = 1 Then
            lnBody4.Visible = True
        ElseIf lives = 0 Then
            lnBody5.Visible = True
            PlaySound "D:\CP2\VB6 Project Files\Hangman\resources\oof.wav", 0, SND_FILENAME Or SND_ASYNC
            If gametype = 3 Then
                MsgBox "The AI failed! You win!"
            Else
                MsgBox "You lose! The correct word was " + word + "."
            End If
            Disable
        End If
    End If
End If

End Sub

Private Sub mnuAI_Click()

gametype = 3
userword = InputBox("Enter a word for the AI to guess:", "AI")

End Sub

Private Sub mnuLong_Click()

gametype = 2

End Sub

Private Sub mnuRegular_Click()

gametype = 0

End Sub

Private Sub mnuShort_Click()

gametype = 1

End Sub

Private Sub tmrAI_Timer()

Dim rng As Integer

aicounter = aicounter + 1
If aicounter = 1 And gotvowel = False Then
    lblKeyboard_Click (2) 'e
    gotvowel = True
ElseIf aicounter = 2 And gotvowel = False Then
    lblKeyboard_Click (10) 'a
    gotvowel = True
ElseIf aicounter = 3 And gotvowel = False Then
    lblKeyboard_Click (8) 'o
    gotvowel = True
ElseIf aicounter = 4 And gotvowel = False Then
    lblKeyboard_Click (7) 'i
    gotvowel = True
ElseIf aicounter = 5 And gotvowel = False Then
    lblKeyboard_Click (6) 'u
    gotvowel = True
ElseIf aicounter > 5 And gotvowel = True Then
    rng = Int(25 * Rnd)
    lblKeyboard_Click (rng)
End If

End Sub

Private Sub tmrAnimation_Timer()

If imgWinner.Visible = False Then
    imgWinner.Visible = True
Else
    imgWinner.Visible = False
End If

End Sub

Private Sub txtInput_Change()

If enabledk = True Then
    If txtInput = "q" Then
        lblKeyboard_Click (0)
    ElseIf txtInput = "w" Then
        lblKeyboard_Click (1)
    ElseIf txtInput = "e" Then
        lblKeyboard_Click (2)
    ElseIf txtInput = "r" Then
        lblKeyboard_Click (3)
    ElseIf txtInput = "t" Then
        lblKeyboard_Click (4)
    ElseIf txtInput = "y" Then
        lblKeyboard_Click (5)
    ElseIf txtInput = "u" Then
        lblKeyboard_Click (6)
    ElseIf txtInput = "i" Then
        lblKeyboard_Click (7)
    ElseIf txtInput = "o" Then
        lblKeyboard_Click (8)
    ElseIf txtInput = "p" Then
        lblKeyboard_Click (9)
    ElseIf txtInput = "a" Then
        lblKeyboard_Click (10)
    ElseIf txtInput = "s" Then
        lblKeyboard_Click (11)
    ElseIf txtInput = "d" Then
        lblKeyboard_Click (12)
    ElseIf txtInput = "f" Then
        lblKeyboard_Click (13)
    ElseIf txtInput = "g" Then
        lblKeyboard_Click (14)
    ElseIf txtInput = "h" Then
        lblKeyboard_Click (15)
    ElseIf txtInput = "j" Then
        lblKeyboard_Click (16)
    ElseIf txtInput = "k" Then
        lblKeyboard_Click (17)
    ElseIf txtInput = "l" Then
        lblKeyboard_Click (18)
    ElseIf txtInput = "z" Then
        lblKeyboard_Click (19)
    ElseIf txtInput = "x" Then
        lblKeyboard_Click (20)
    ElseIf txtInput = "c" Then
        lblKeyboard_Click (21)
    ElseIf txtInput = "v" Then
        lblKeyboard_Click (22)
    ElseIf txtInput = "b" Then
        lblKeyboard_Click (23)
    ElseIf txtInput = "n" Then
        lblKeyboard_Click (24)
    ElseIf txtInput = "m" Then
        lblKeyboard_Click (25)
    End If
End If
txtInput = ""

End Sub
