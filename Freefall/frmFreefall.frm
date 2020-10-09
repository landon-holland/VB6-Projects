VERSION 5.00
Begin VB.Form frmFreefall 
   Caption         =   "Freefall"
   ClientHeight    =   2344
   ClientLeft      =   120
   ClientTop       =   464
   ClientWidth     =   3792
   LinkTopic       =   "Form1"
   ScaleHeight     =   2344
   ScaleWidth      =   3792
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optMeters 
      Caption         =   "Meters"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.86
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1920
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
   Begin VB.OptionButton optFeet 
      Caption         =   "Feet"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.86
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   840
      TabIndex        =   1
      Top             =   600
      Width           =   915
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   315
      Left            =   2460
      TabIndex        =   5
      Top             =   1020
      Width           =   1095
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   315
      Left            =   1380
      TabIndex        =   4
      Top             =   1020
      Width           =   1095
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "Find"
      Height          =   315
      Left            =   240
      TabIndex        =   3
      Top             =   1020
      Width           =   1155
   End
   Begin VB.TextBox txtTime 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.86
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2040
      TabIndex        =   0
      Top             =   180
      Width           =   1515
   End
   Begin VB.Label lblDistanceUnit 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   376
      Left            =   1320
      TabIndex        =   9
      Top             =   1920
      Width           =   1512
   End
   Begin VB.Label lblDistance 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1320
      TabIndex        =   8
      Top             =   1500
      Width           =   2295
   End
   Begin VB.Label lblDistanceOutput 
      Caption         =   "Distance:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.86
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label lblTimeInput 
      Caption         =   "Time (Seconds):"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.86
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   240
      Width           =   1755
   End
End
Attribute VB_Name = "frmFreefall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClear_Click()
txtTime = ""
lblDistance = ""
lblDistanceUnit = ""
txtTime.SetFocus
End Sub

Private Sub cmdExit_Click()
End
End Sub

Private Sub cmdFind_Click()
Dim FallTime As Single
Dim MGrav As Single
Dim FGrav As Single
Dim Distance As Single
FGrav = 32.2
MGrav = 9.81

If Not IsNumeric(txtTime.Text) Then
    MsgBox "The value you entered is not a number."
    txtTime = ""
    lblDistance = ""
    lblDistanceUnit = ""
    txtTime.SetFocus
ElseIf optFeet = False And optMeters = False Then
    MsgBox "Please select Meters or Feet."
    optFeet.SetFocus
Else
    FallTime = txtTime.Text
    FallTime = Val(FallTime)
    
    If optFeet = True Then
        Distance = 0.5 * FGrav * FallTime ^ 2
        lblDistance = Distance
        lblDistanceUnit = "feet"
    ElseIf optMeters = True Then
        Distance = 0.5 * MGrav * FallTime ^ 2
        lblDistance = Distance
        lblDistanceUnit = "meters"
    Else
        MsgBox "Internal Error."
        txtTime.SetFocus
    End If
End If
End Sub
