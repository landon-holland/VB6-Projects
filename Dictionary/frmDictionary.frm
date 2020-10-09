VERSION 5.00
Begin VB.Form frmDictionary 
   Caption         =   "Dictionary"
   ClientHeight    =   3915
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   3600
   LinkTopic       =   "Form1"
   ScaleHeight     =   3915
   ScaleWidth      =   3600
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFind 
      Caption         =   "Find"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   60
      TabIndex        =   6
      Top             =   3600
      Width           =   3495
   End
   Begin VB.CommandButton cmdFill 
      Appearance      =   0  'Flat
      Caption         =   "Fill"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2565
      Left            =   1800
      TabIndex        =   1
      Top             =   0
      Width           =   1755
   End
   Begin VB.ListBox lstDictionary 
      Appearance      =   0  'Flat
      Height          =   2370
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1755
   End
   Begin VB.Label lblWordCount 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   0
      TabIndex        =   9
      Top             =   2400
      Width           =   1755
   End
   Begin VB.Label lblRndWord 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1260
      TabIndex        =   8
      Top             =   3240
      Width           =   2295
   End
   Begin VB.Label lblRndWordPrompt 
      Caption         =   "Random Word:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   3240
      Width           =   1875
   End
   Begin VB.Label lblLength 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   660
      TabIndex        =   5
      Top             =   2940
      Width           =   2895
   End
   Begin VB.Label lblLengthPrompt 
      Caption         =   "Length:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   2940
      Width           =   1875
   End
   Begin VB.Label lblLongestWord 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1260
      TabIndex        =   3
      Top             =   2640
      Width           =   2295
   End
   Begin VB.Label lblLongestWordPrompt 
      Caption         =   "Longest Word:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   2640
      Width           =   1875
   End
   Begin VB.Menu mnuSettings 
      Caption         =   "Settings"
      Begin VB.Menu mnuShort 
         Caption         =   "Short"
      End
      Begin VB.Menu mnuLong 
         Caption         =   "Long"
      End
      Begin VB.Menu mnuDisable 
         Caption         =   "Disable"
      End
   End
End
Attribute VB_Name = "frmDictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim arrwords(1 To 58112) As String
Dim arr3(1 To 637) As String
Dim arr7(1 To 43978) As String
Dim wordlength As Integer

Private Sub cmdFill_Click()

Dim path As String
Dim line As String
Dim i As Long
Dim j As Long
Dim k As Long

j = 0
k = 0

path = "D:\CP2\VB6 Project Files\Dictionary\resources\dictionary.txt"

Open path For Input As #1

    For i = 1 To 58112
    
        Line Input #1, line
        
        line = Trim(line)
        
        arrwords(i) = line
        
        lstDictionary.AddItem line
        
        If Len(line) = 3 Then
            j = j + 1
            arr3(j) = line
        ElseIf Len(line) >= 7 Then
            k = k + 1
            arr7(k) = line
        End If
        
    Next i
    
Close #1

End Sub

Private Sub cmdFind_Click()

Dim i As Long
Dim testword As String
Dim longword As String
Dim rng As Long
Dim done As Boolean

longword = "a"

For i = 1 To 58112

    testword = arrwords(i)
    
    If Len(testword) > Len(longword) Then
    
        longword = testword
        
    End If
    
Next i

lblLongestWord = longword
lblLength = Str(Len(longword))

done = False

Do While done = False

    Randomize
    
    If wordlength = 0 Then
        rng = Int(58112 * Rnd + 1)
        lblRndWord = arrwords(rng)
        done = True
    Else
        If wordlength = 1 Then
        rng = Int(637 * Rnd + 1)
        lblRndWord = arr3(rng)
        done = True
        ElseIf wordlength = 2 Then
        rng = Int(43977 * Rnd + 1)
        lblRndWord = arr7(rng)
        done = True
        End If
    End If
Loop

lblWordCount = "58112"

End Sub

Private Sub Form_Load()

wordlength = 0

End Sub

Private Sub mnuDisable_Click()

wordlength = 0

End Sub

Private Sub mnuLong_Click()

wordlength = 2

End Sub

Private Sub mnuShort_Click()

wordlength = 1

End Sub
