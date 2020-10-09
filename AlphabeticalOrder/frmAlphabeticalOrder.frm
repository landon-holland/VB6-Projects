VERSION 5.00
Begin VB.Form frmAlphabeticalOrder 
   Caption         =   "Alphabetical Order"
   ClientHeight    =   3750
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3660
   LinkTopic       =   "Form1"
   ScaleHeight     =   3750
   ScaleWidth      =   3660
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtWord4 
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
      Height          =   315
      Left            =   1080
      TabIndex        =   6
      Top             =   1320
      Width           =   2355
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
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
      Left            =   2400
      TabIndex        =   9
      Top             =   1800
      Width           =   1035
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
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
      Left            =   1260
      TabIndex        =   8
      Top             =   1800
      Width           =   1155
   End
   Begin VB.CommandButton cmdCompare 
      Caption         =   "Compare"
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
      TabIndex        =   7
      Top             =   1800
      Width           =   1035
   End
   Begin VB.TextBox txtWord3 
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
      Height          =   315
      Left            =   1080
      TabIndex        =   5
      Top             =   960
      Width           =   2355
   End
   Begin VB.TextBox txtWord2 
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
      Height          =   315
      Left            =   1080
      TabIndex        =   4
      Top             =   600
      Width           =   2355
   End
   Begin VB.TextBox txtWord1 
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
      Height          =   315
      Left            =   1080
      TabIndex        =   3
      Top             =   240
      Width           =   2355
   End
   Begin VB.Label lblA4 
      Alignment       =   2  'Center
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
      Left            =   240
      TabIndex        =   14
      Top             =   3360
      Width           =   3195
   End
   Begin VB.Label lblA3 
      Alignment       =   2  'Center
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
      Left            =   240
      TabIndex        =   13
      Top             =   3000
      Width           =   3195
   End
   Begin VB.Label lblA2 
      Alignment       =   2  'Center
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
      Left            =   240
      TabIndex        =   12
      Top             =   2640
      Width           =   3195
   End
   Begin VB.Label lblA1 
      Alignment       =   2  'Center
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
      Left            =   240
      TabIndex        =   11
      Top             =   2280
      Width           =   3195
   End
   Begin VB.Label lblWord4Prompt 
      Caption         =   "Word 4:"
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
      Left            =   240
      TabIndex        =   10
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label lblWord3Prompt 
      Caption         =   "Word 3:"
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
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   855
   End
   Begin VB.Label lblWord2Prompt 
      Caption         =   "Word 2:"
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
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   855
   End
   Begin VB.Label lblWord1Prompt 
      Caption         =   "Word 1:"
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
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "frmAlphabeticalOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClear_Click()
txtWord1 = ""
txtWord2 = ""
txtWord3 = ""
txtWord4 = ""
lblA1 = ""
lblA2 = ""
lblA3 = ""
lblA4 = ""
txtWord1.SetFocus
End Sub

Private Sub cmdCompare_Click()
Dim W1 As String
Dim W2 As String
Dim W3 As String
Dim W4 As String
Dim A1 As String
Dim A2 As String
Dim A3 As String
Dim A4 As String

W1 = txtWord1
W2 = txtWord2
W3 = txtWord3
W4 = txtWord4

If StrComp(W1, W2, vbTextCompare) = 0 Or StrComp(W1, W2, vbTextCompare) = -1 Then
    A1 = W1
    A2 = W2
Else
    A1 = W2
    A2 = W1
End If

If StrComp(W3, A1, vbTextCompare) = 0 Or StrComp(W3, A1, vbTextCompare) = -1 Then
    A3 = A2
    A2 = A1
    A1 = W3
ElseIf StrComp(W3, A2, vbTextCompare) = -1 Or StrComp(W3, A2, vbTextCompare) = 0 Then
    A3 = A2
    A2 = W3
ElseIf StrComp(W3, A2, vbTextCompare) = 1 Then
    A3 = W3
End If

If StrComp(W4, A1, vbTextCompare) = 0 Or StrComp(W4, A1, vbTextCompare) = -1 Then
    A4 = A3
    A3 = A2
    A2 = A1
    A1 = W4
ElseIf StrComp(W4, A2, vbTextCompare) = -1 Or StrComp(W4, A2, vbTextCompare) = 0 Then
    A4 = A3
    A3 = A2
    A2 = W4
ElseIf StrComp(W4, A3, vbTextCompare) = -1 Or StrComp(W4, A3, vbTextCompare) = 0 Then
    A4 = A3
    A3 = W4
ElseIf StrComp(W4, A3, vbTextCompare) = 1 Then
    A4 = W4
End If

lblA1 = A1
lblA2 = A2
lblA3 = A3
lblA4 = A4
cmdClear.SetFocus
End Sub

Private Sub cmdExit_Click()
End
End Sub

Private Sub txtWord1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtWord2.SetFocus
End If
End Sub

Private Sub txtWord2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtWord3.SetFocus
End If
End Sub

Private Sub txtWord3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtWord4.SetFocus
End If
End Sub

Private Sub txtWord4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdCompare_Click
End If
End Sub
