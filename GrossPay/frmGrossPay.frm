VERSION 5.00
Begin VB.Form frmGrossPay 
   Caption         =   "Gross Pay Calculator"
   ClientHeight    =   3150
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3705
   LinkTopic       =   "Form1"
   ScaleHeight     =   3150
   ScaleWidth      =   3705
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   315
      Left            =   1500
      TabIndex        =   3
      Top             =   1800
      Width           =   795
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   315
      Left            =   2760
      TabIndex        =   4
      Top             =   1800
      Width           =   855
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "Find"
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   855
   End
   Begin VB.TextBox txtHours 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   1
      Top             =   900
      Width           =   2715
   End
   Begin VB.TextBox txtWage 
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
      Left            =   840
      TabIndex        =   0
      Top             =   360
      Width           =   2715
   End
   Begin VB.Label lblGrossPay 
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
      Height          =   435
      Left            =   1260
      TabIndex        =   9
      Top             =   2460
      Width           =   1635
   End
   Begin VB.Label lblGrossPayOutput 
      Caption         =   "Gross Pay:"
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
      Left            =   60
      TabIndex        =   8
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label lblHoursInput 
      Caption         =   "Hours"
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
      Left            =   60
      TabIndex        =   7
      Top             =   1020
      Width           =   615
   End
   Begin VB.Label lblWageInput 
      Caption         =   "Wage"
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
      Left            =   60
      TabIndex        =   6
      Top             =   420
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   75
      Left            =   6540
      TabIndex        =   5
      Top             =   900
      Width           =   15
   End
End
Attribute VB_Name = "frmGrossPay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClear_Click()
txtWage = ""
txtHours = ""
lblGrossPay = ""
txtWage.SetFocus
End Sub

Private Sub cmdFind_Click()
Dim hours As Single
Dim wage As Single
Dim gp As Single

If Not IsNumeric(txtWage.Text) Or Not IsNumeric(txtHours.Text) Then
    MsgBox "The value you entered is not a number."
    txtWage = ""
    txtHours = ""
    lblGrossPay = ""
    txtWage.SetFocus
Else
    hours = txtHours.Text
    wage = txtWage.Text
    
    hours = Val(hours)
    wage = Val(wage)
    
    gp = hours * wage

    lblGrossPay = gp
End If
End Sub

Private Sub cmdQuit_Click()
End
End Sub

