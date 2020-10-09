VERSION 5.00
Begin VB.Form frmClassFees 
   Caption         =   "Class Fees"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5565
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   5565
   StartUpPosition =   3  'Windows Default
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
      Height          =   375
      Left            =   2100
      TabIndex        =   16
      Top             =   2580
      Width           =   1395
   End
   Begin VB.HScrollBar hsbTable 
      Height          =   255
      LargeChange     =   4
      Left            =   120
      Max             =   10
      Min             =   1
      TabIndex        =   15
      Top             =   1320
      Value           =   1
      Width           =   5295
   End
   Begin VB.TextBox txtClassFees 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   960
      Width           =   5295
   End
   Begin VB.CommandButton cmdInput 
      Caption         =   "Input"
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
      Left            =   3840
      TabIndex        =   1
      Top             =   60
      Width           =   1575
   End
   Begin VB.Label lblG12 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   4440
      TabIndex        =   14
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label lblG11 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   4440
      TabIndex        =   13
      Top             =   1860
      Width           =   1095
   End
   Begin VB.Label lblG10 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   120
      TabIndex        =   12
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label lblG9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   120
      TabIndex        =   11
      Top             =   1860
      Width           =   1095
   End
   Begin VB.Label lblG12Prompt 
      Caption         =   "Grade 12"
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
      Left            =   4440
      TabIndex        =   10
      Top             =   2340
      Width           =   975
   End
   Begin VB.Label lblG11Prompt 
      Caption         =   "Grade 11"
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
      Left            =   4440
      TabIndex        =   9
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label lblG10Prompt 
      Caption         =   "Grade 10"
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
      Left            =   120
      TabIndex        =   8
      Top             =   2340
      Width           =   1035
   End
   Begin VB.Label lblG9Prompt 
      Caption         =   "Grade 9"
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
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   915
   End
   Begin VB.Label lblFee 
      Caption         =   "Fee"
      Height          =   315
      Left            =   2400
      TabIndex        =   6
      Top             =   660
      Width           =   1215
   End
   Begin VB.Label lblGrade 
      Caption         =   "Grade"
      Height          =   315
      Left            =   1500
      TabIndex        =   5
      Top             =   660
      Width           =   1215
   End
   Begin VB.Label lblName 
      Caption         =   "Name"
      Height          =   315
      Left            =   720
      TabIndex        =   4
      Top             =   660
      Width           =   1215
   End
   Begin VB.Label lblNumber 
      Caption         =   "#"
      Height          =   315
      Left            =   240
      TabIndex        =   3
      Top             =   660
      Width           =   255
   End
   Begin VB.Label lblInputButtonPrompt 
      Caption         =   "Click here to input information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4005
   End
End
Attribute VB_Name = "frmClassFees"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim arrnames(1 To 10) As String
Dim arrgrade(1 To 10) As Integer
Dim arrfee(1 To 10) As Currency
Dim table(1 To 10) As String

Private Sub cmdExit_Click()

Hide

frmMainMenu.Show

End Sub

Private Sub cmdInput_Click()

Dim i As Integer
Dim displine As String
Dim total9 As Currency
Dim total10 As Currency
Dim total11 As Currency
Dim total12 As Currency

For i = 1 To 2

    arrnames(i) = InputBox("Enter Student " + Str(i) + "'s Names:", "Input")
    arrgrade(i) = InputBox("Enter Student " + Str(i) + "'s Grade:", "Input")
    arrfee(i) = InputBox("Enter Student " + Str(i) + "'s Fee:", "Input")

Next i

For i = 1 To 10

    displine = ""
    displine = displine + Str(i)
    displine = displine + vbTab + arrnames(i)
    displine = displine + vbTab + Str(arrgrade(i))
    displine = displine + vbTab + Format(arrfee(i), "Currency")
    
    table(i) = displine

Next i

txtClassFees = table(1)

For i = 1 To 10

    If arrgrade(i) = 9 Then
    
        total9 = total9 + arrfee(i)
        
    ElseIf arrgrade(i) = 10 Then
    
        total10 = total10 + arrfee(i)
        
    ElseIf arrgrade(i) = 11 Then
    
        total11 = total11 + arrfee(i)
        
    ElseIf arrgrade(i) = 12 Then
    
        total12 = total12 + arrfee(i)
        
    End If
    
Next i

lblG9 = Format(total9, "Currency")
lblG10 = Format(total10, "Currency")
lblG11 = Format(total11, "Currency")
lblG12 = Format(total12, "Currency")

End Sub

Private Sub Form_Unload(Cancel As Integer)

Hide

frmMainMenu.Show

End Sub

Private Sub hsbTable_Change()

txtClassFees = table(hsbTable.Value)

End Sub
