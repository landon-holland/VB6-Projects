VERSION 5.00
Begin VB.Form frmOneHundredStudents 
   Caption         =   "One Hundred Students"
   ClientHeight    =   975
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5835
   LinkTopic       =   "Form1"
   ScaleHeight     =   975
   ScaleWidth      =   5835
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open"
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
      Left            =   4380
      TabIndex        =   4
      Top             =   660
      Width           =   1425
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
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
      Left            =   2940
      TabIndex        =   3
      Top             =   660
      Width           =   1425
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "Find"
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
      TabIndex        =   2
      Top             =   660
      Width           =   2918
   End
   Begin VB.HScrollBar hsbStudents 
      Height          =   255
      LargeChange     =   10
      Left            =   0
      Max             =   100
      Min             =   1
      TabIndex        =   1
      Top             =   360
      Value           =   1
      Width           =   5835
   End
   Begin VB.TextBox txtStudents 
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
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   5835
   End
End
Attribute VB_Name = "frmOneHundredStudents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim arrnames(1 To 88799) As String
Dim textline(1 To 100) As String

Sub outputtext()

Dim i As Integer

For i = 1 To 100
    With arrstudents(i)
        textline(i) = .last + vbTab + Str(.id) + Str(.gpa)
    End With
Next i
txtStudents = textline(1)

End Sub

Private Sub cmdFind_Click()

Dim i As Long
Randomize

Dim line As String
Open "D:\CP2\VB6 Project Files\OneHundredStudents\names.txt" For Input As #1
    For i = 1 To 88799
        Line Input #1, line
        line = Trim(line)
        arrnames(i) = line
    Next i
Close #1

'Fill
Dim currentid As Long
currentid = 100000
For i = 1 To 100
    With arrstudents(i)
        .gpa = (Int(Rnd * 300) + 100) / 100
        currentid = currentid + 1
        .id = currentid
        .last = arrnames(Int(Rnd * 88799) + 1)
    End With
Next i

'Output
outputtext

End Sub

Private Sub cmdOpen_Click()

Dim i As Integer

Open "D:\CP2\VB6 Project Files\OneHundredStudents\students.dat" For Random Access Read As #1
    i = 1
    Do While Not EOF(1)
        Get #1, , arrstudents(i)
        i = i + 1
    Loop
Close #1
outputtext

End Sub

Private Sub cmdSave_Click()

Dim i As Integer

Open "D:\CP2\VB6 Project Files\OneHundredStudents\students.dat" For Random Access Write As #1
    For i = 1 To 100
        Put #1, i, arrstudents(i)
    Next i
Close #1

End Sub

Private Sub hsbStudents_Change()

txtStudents = textline(hsbStudents)

End Sub
