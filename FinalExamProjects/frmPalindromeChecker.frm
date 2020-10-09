VERSION 5.00
Begin VB.Form frmPalindromeChecker 
   Caption         =   "Palindrome Checker"
   ClientHeight    =   1590
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3885
   LinkTopic       =   "Form1"
   ScaleHeight     =   1590
   ScaleWidth      =   3885
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Appearance      =   0  'Flat
      Caption         =   "Exit"
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
      Left            =   0
      TabIndex        =   4
      Top             =   1260
      Width           =   3900
   End
   Begin VB.TextBox txtWord 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Top             =   420
      Width           =   3885
   End
   Begin VB.Label lblNo 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "No"
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
      Left            =   1980
      TabIndex        =   3
      Top             =   780
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.Label lblYes 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      Caption         =   "Palindrome"
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
      Left            =   0
      TabIndex        =   2
      Top             =   780
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.Label lblWordPrompt 
      Alignment       =   2  'Center
      Caption         =   "Word:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   -60
      TabIndex        =   0
      Top             =   120
      Width           =   4005
   End
End
Attribute VB_Name = "frmPalindromeChecker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()

Hide

frmMainMenu.Show

End Sub

Private Sub Form_Unload(Cancel As Integer)

Hide

frmMainMenu.Show

End Sub

Private Sub txtWord_Change()

Dim body As String

body = txtWord

body = Replace(body, " ", "")
body = Replace(body, ",", "")
body = Replace(body, ".", "")
body = Replace(body, "!", "")
body = Replace(body, "?", "")
body = Replace(body, """", "")
body = Replace(body, "''", "")

body = UCase(body)

If body = StrReverse(body) Then

    lblYes.Visible = True
    lblNo.Visible = False
    
Else

    lblNo.Visible = True
    lblYes.Visible = False
    
End If

End Sub
