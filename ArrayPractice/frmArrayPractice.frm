VERSION 5.00
Begin VB.Form frmArrayPractice 
   Caption         =   "Array Practice"
   ClientHeight    =   5055
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6300
   LinkTopic       =   "Form1"
   ScaleHeight     =   5055
   ScaleWidth      =   6300
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClear 
      Appearance      =   0  'Flat
      Caption         =   "Clear"
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
      Left            =   2220
      TabIndex        =   7
      Top             =   4620
      Width           =   1875
   End
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
      Left            =   4260
      TabIndex        =   6
      Top             =   4620
      Width           =   1875
   End
   Begin VB.ListBox lstFile 
      Appearance      =   0  'Flat
      Height          =   3930
      ItemData        =   "frmArrayPractice.frx":0000
      Left            =   4260
      List            =   "frmArrayPractice.frx":0002
      TabIndex        =   5
      Top             =   600
      Width           =   1875
   End
   Begin VB.CommandButton cmdFile 
      Appearance      =   0  'Flat
      Caption         =   "File"
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
      Left            =   4260
      TabIndex        =   4
      Top             =   180
      Width           =   1875
   End
   Begin VB.ListBox lstRNG 
      Appearance      =   0  'Flat
      Height          =   3930
      ItemData        =   "frmArrayPractice.frx":0004
      Left            =   2220
      List            =   "frmArrayPractice.frx":0006
      TabIndex        =   3
      Top             =   600
      Width           =   1875
   End
   Begin VB.CommandButton cmdRNG 
      Appearance      =   0  'Flat
      Caption         =   "RNG"
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
      Left            =   2220
      TabIndex        =   2
      Top             =   180
      Width           =   1875
   End
   Begin VB.ListBox lstKeyboard 
      Appearance      =   0  'Flat
      Height          =   3930
      ItemData        =   "frmArrayPractice.frx":0008
      Left            =   180
      List            =   "frmArrayPractice.frx":000A
      TabIndex        =   1
      Top             =   600
      Width           =   1875
   End
   Begin VB.CommandButton cmdKeyboard 
      Appearance      =   0  'Flat
      Caption         =   "Keyboard Entry"
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
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   1875
   End
End
Attribute VB_Name = "frmArrayPractice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClear_Click()

lstKeyboard.Clear
lstRNG.Clear
lstFile.Clear

End Sub

Private Sub cmdExit_Click()

End

End Sub

Private Sub cmdFile_Click()

Dim arrints(20) As Integer
Dim i As Integer
Dim path As String
Dim line As String

path = "D:\CP1\VB6 Project Files\ArrayPractice\array.txt"

Open path For Input As #1

    For i = 1 To 20
    
        If EOF(1) Then
    
            Exit For
        
        End If
    
        Line Input #1, line
    
        arrints(i) = Val(line)
        lstFile.AddItem arrints(i)
        
    Next i
    
Close #1


End Sub

Private Sub cmdKeyboard_Click()

Dim arrints(20) As Integer
Dim i As Integer

For i = 1 To 20

    arrints(i) = Val(InputBox("Enter Number:", "Enter Number"))
    
    lstKeyboard.AddItem arrints(i)
    
Next i

End Sub

Private Sub cmdRNG_Click()

Dim arrints(20) As Integer
Dim i As Integer

Randomize

For i = 1 To 20

    arrints(i) = Int((100 - 1 + 1) * Rnd + 1)
    
    lstRNG.AddItem arrints(i)
    
Next i

End Sub

