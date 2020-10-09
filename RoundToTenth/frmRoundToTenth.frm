VERSION 5.00
Begin VB.Form frmRoundToTenth 
   Caption         =   "Round to the Nearest Tenth"
   ClientHeight    =   1680
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   2850
   LinkTopic       =   "Form1"
   ScaleHeight     =   1680
   ScaleWidth      =   2850
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   315
      Left            =   1800
      TabIndex        =   4
      Top             =   660
      Width           =   855
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   315
      Left            =   1080
      TabIndex        =   3
      Top             =   660
      Width           =   735
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "Calculate"
      Height          =   315
      Left            =   240
      TabIndex        =   2
      Top             =   660
      Width           =   855
   End
   Begin VB.TextBox txtNumber 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   180
      Width           =   1455
   End
   Begin VB.Label lblRounded 
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
      Height          =   375
      Left            =   1320
      TabIndex        =   6
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label lblRoundedPrompt 
      Caption         =   "Rounded:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   1140
      Width           =   1155
   End
   Begin VB.Label lblNumberPrompt 
      Caption         =   "Number:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "frmRoundToTenth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCalculate_Click()
Dim Num As Double
Dim TempNum As Double
Dim TempNum2 As Double
If Not IsNumeric(txtNumber) Then
    MsgBox "That's not a number."
Else
    Num = Val(txtNumber)
    TempNum = Num * 100
    TempNum = Int(TempNum)
    TempNum = TempNum Mod 10
    If Num >= 0 Then
    
        If TempNum >= 5 Then
            TempNum2 = Num * 100
            TempNum2 = Int(TempNum2)
            TempNum2 = TempNum2 / 10
            TempNum2 = Int(TempNum2)
            TempNum2 = TempNum2 + 1
            TempNum2 = TempNum2 / 10
            Num = TempNum2
        Else
            TempNum2 = Num * 100
            TempNum2 = Int(TempNum2)
            TempNum2 = TempNum2 / 10
            TempNum2 = Int(TempNum2)
            TempNum2 = TempNum2 / 10
            Num = TempNum2
        End If
    Else
            If TempNum >= -5 Then
            TempNum2 = Num * 100
            TempNum2 = Int(TempNum2)
            TempNum2 = TempNum2 / 10
            TempNum2 = Int(TempNum2)
            TempNum2 = TempNum2 + 1
            TempNum2 = TempNum2 / 10
            Num = TempNum2
        Else
            TempNum2 = Num * 100
            TempNum2 = Int(TempNum2)
            TempNum2 = TempNum2 / 10
            TempNum2 = Int(TempNum2)
            TempNum2 = TempNum2
            TempNum2 = TempNum2 / 10
            Num = TempNum2
        End If
    End If
    cmdClear.SetFocus
    lblRounded = Num
End If
End Sub

Private Sub cmdClear_Click()
txtNumber = ""
lblRounded = ""
txtNumber.SetFocus

End Sub

Private Sub cmdExit_Click()
End
End Sub

