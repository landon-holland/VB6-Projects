VERSION 5.00
Begin VB.Form frmMultiples 
   Caption         =   "Multiples"
   ClientHeight    =   2895
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3705
   LinkTopic       =   "Form1"
   ScaleHeight     =   2895
   ScaleWidth      =   3705
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstMultiples 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   1860
      TabIndex        =   10
      Top             =   1920
      Width           =   1575
   End
   Begin VB.TextBox txtMultiplesOutput 
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
      Height          =   840
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   1920
      Width           =   1575
   End
   Begin VB.TextBox txtMultiples 
      Appearance      =   0  'Flat
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
      Left            =   1980
      TabIndex        =   1
      Top             =   660
      Width           =   1455
   End
   Begin VB.CommandButton cmdExit 
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
      Left            =   2340
      TabIndex        =   5
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton cmdClear 
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
      Left            =   1320
      TabIndex        =   4
      Top             =   1080
      Width           =   1035
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "Calculate"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox txtNumber 
      Appearance      =   0  'Flat
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
      Left            =   1260
      TabIndex        =   0
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label lblSum 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   900
      TabIndex        =   8
      Top             =   1500
      Width           =   2535
   End
   Begin VB.Label lblSumPrompt 
      Caption         =   "Sum:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   7
      Top             =   1500
      Width           =   675
   End
   Begin VB.Label lblMultiplesPrompt 
      Caption         =   "# Of Multiples:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   6
      Top             =   660
      Width           =   1755
   End
   Begin VB.Label lblNumberPrompt 
      Caption         =   "Number:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   1035
   End
End
Attribute VB_Name = "frmMultiples"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCalc_Click()
Dim i As Integer
Dim Num As Long
Dim Multiples As Integer
Dim Sum As Long

If Not IsNumeric(txtNumber) Or Not IsNumeric(txtMultiples) Then
    MsgBox "Those are not numbers!"
ElseIf txtNumber = 0 Or txtMultiples = 0 Then
    MsgBox "Please don't enter zero."
Else
    Num = Val(txtNumber)
    Multiples = Val(txtMultiples)
    txtMultiplesOutput = ""
    lstMultiples.Clear
    For i = Num To Num * Multiples Step Num
        lstMultiples.AddItem (i)
        txtMultiplesOutput = txtMultiplesOutput + Str(i) + " "
        Sum = Sum + i
    Next i
    lblSum = Sum
    cmdClear.SetFocus
End If
End Sub

Private Sub cmdClear_Click()
txtNumber = ""
txtMultiples = ""
txtMultiplesOutput = ""
lblSum = ""
lstMultiples.Clear
txtNumber.SetFocus
End Sub

Private Sub cmdExit_Click()
End
End Sub

Private Sub txtMultiples_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdCalc_Click
End If
End Sub

Private Sub txtNumber_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtMultiples.SetFocus
End If
End Sub
