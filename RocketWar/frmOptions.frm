VERSION 5.00
Begin VB.Form frmOptions 
   Caption         =   "Options"
   ClientHeight    =   1035
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2370
   LinkTopic       =   "Form1"
   ScaleHeight     =   1035
   ScaleWidth      =   2370
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optKeyboardControl 
      Caption         =   "Keyboard"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1140
      TabIndex        =   2
      Top             =   540
      Width           =   1215
   End
   Begin VB.OptionButton optMouseControl 
      Caption         =   "Mouse"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   180
      TabIndex        =   1
      Top             =   540
      Width           =   855
   End
   Begin VB.Label lblControlPrompt 
      Caption         =   "Control Method:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   2055
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Unload(Cancel As Integer)

Hide
frmMenu.Show

End Sub

Private Sub optKeyboardControl_Click()

MouseControl = False

End Sub

Private Sub optMouseControl_Click()

MouseControl = True

End Sub
