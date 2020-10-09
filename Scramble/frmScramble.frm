VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form frmScrambleMenu 
   Caption         =   "Scramble"
   ClientHeight    =   3465
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12600
   LinkTopic       =   "Form1"
   ScaleHeight     =   3465
   ScaleWidth      =   12600
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrMusic 
      Enabled         =   0   'False
      Interval        =   469
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton cmdHelp 
      Appearance      =   0  'Flat
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   4200
      TabIndex        =   4
      Top             =   2700
      Width           =   4200
   End
   Begin VB.CommandButton cmdHighScores 
      Appearance      =   0  'Flat
      Caption         =   "High Scores"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   0
      TabIndex        =   3
      Top             =   2700
      Width           =   4200
   End
   Begin VB.CommandButton cmdExit 
      Appearance      =   0  'Flat
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   8400
      TabIndex        =   5
      Top             =   2700
      Width           =   4200
   End
   Begin VB.CommandButton cmd3Letter 
      Appearance      =   0  'Flat
      Caption         =   "3-Letter Word"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   0
      TabIndex        =   0
      Top             =   2040
      Width           =   4200
   End
   Begin VB.CommandButton cmd5Letter 
      Appearance      =   0  'Flat
      Caption         =   "5-Letter Word"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   4200
      TabIndex        =   1
      Top             =   2040
      Width           =   4200
   End
   Begin VB.CommandButton cmd7Letter 
      Appearance      =   0  'Flat
      Caption         =   "7-Letter Word"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   8400
      TabIndex        =   2
      Top             =   2040
      Width           =   4200
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmpMusic 
      Height          =   435
      Left            =   0
      TabIndex        =   7
      Top             =   420
      Visible         =   0   'False
      Width           =   435
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   767
      _cy             =   767
   End
   Begin VB.Label lblScrambleTitle 
      Alignment       =   2  'Center
      Caption         =   "S C R A M B L E"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1515
      Left            =   -60
      TabIndex        =   6
      Top             =   0
      Width           =   12840
   End
End
Attribute VB_Name = "frmScrambleMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim beat As Long

Private Sub cmd3Letter_Click()

gametype = 1

Hide

frmGame.Show

End Sub

Private Sub cmd5Letter_Click()

gametype = 2

Hide

frmGame.Show

End Sub

Private Sub cmd7Letter_Click()

gametype = 3

Hide

frmGame.Show

End Sub

Private Sub cmdHelp_Click()

Hide

frmHelp.Show

End Sub

Private Sub cmdExit_Click()

End

End Sub

Private Sub cmdHighScores_Click()

globalscore = -1

Hide

frmHighScores.Show

End Sub

Private Sub Form_Load()

beat = 0

tmrMusic.Enabled = True

End Sub

Private Sub tmrMusic_Timer()

beat = beat + 1

If beat = 1 Then

    wmpMusic.URL = "E:\CP1\VB6 Project Files\Scramble\resources\mainmenumusic.wav"
    
End If
    
If beat Mod 2 = 0 Then

    'BackColor = RGB(0, 0, 0)

Else

    'BackColor = RGB(255, 255, 255)
    
End If

End Sub
