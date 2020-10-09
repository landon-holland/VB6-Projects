VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Employee"
   ClientHeight    =   6272
   ClientLeft      =   152
   ClientTop       =   744
   ClientWidth     =   6400
   LinkTopic       =   "Form1"
   ScaleHeight     =   6272
   ScaleWidth      =   6400
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cndOpen 
      Left            =   5940
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Open"
      InitDir         =   "D:\CP2\VB6 Project Files\Employee\saves"
   End
   Begin VB.ListBox lstEmployees 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.71
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6160
      ItemData        =   "frmMain.frx":0000
      Left            =   0
      List            =   "frmMain.frx":0002
      TabIndex        =   0
      Top             =   0
      Width           =   6400
   End
   Begin MSComDlg.CommonDialog cndSaveAs 
      Left            =   5940
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Save as"
      InitDir         =   "D:\CP2\VB6 Project Files\Employee\saves"
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuNew 
         Caption         =   "New"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Open.."
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save As..."
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuAdd 
         Caption         =   "Add..."
      End
      Begin VB.Menu mnuChange 
         Caption         =   "Change..."
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete..."
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "View"
      Begin VB.Menu mnuRoledex 
         Caption         =   "Roledex"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim path As String

Sub RefreshList()

Dim i As Integer

lstEmployees.Clear

For i = 1 To entries
    With arremployees(i)
        lstEmployees.AddItem Str(.id) + " - " + .lastname + ", " + .firstname + " - " + Str(.age) + " - " + .paytype + " - " + Format(.wage, "Currency") + " - " + .phonenumber
    End With
Next i
End Sub

Private Sub Form_Activate()

RefreshList

End Sub

Private Sub Form_Load()

entries = 0
currententry = 0
entrytochange = 1
entrytodelete = 1
path = ""

End Sub

Private Sub Form_Unload(Cancel As Integer)

End

End Sub

Private Sub lstEmployees_DblClick()

entrytochange = lstEmployees.ListIndex + 1

Hide
With frmDialogue
    .Show
    .cmdChange.Visible = True
    .cmdAdd.Visible = False
    .cmdDelete.Visible = False
    .cmdLeft.Visible = False
    .cmdRight.Visible = False
    
    .txtFirstName = arremployees(entrytochange).firstname
    .txtLastName = arremployees(entrytochange).lastname
    .txtAge = Str(arremployees(entrytochange).age)
    .txtID = Str(arremployees(entrytochange).id)
    .txtPayType = arremployees(entrytochange).paytype
    .txtWage = Str(arremployees(entrytochange).wage)
    .txtPhoneNumber = arremployees(entrytochange).phonenumber
End With

End Sub

Private Sub mnuAdd_Click()

Hide
With frmDialogue
    .Show
    
    .cmdAdd.Visible = True
    .cmdChange.Visible = False
    .cmdDelete.Visible = False
    .cmdLeft.Visible = False
    .cmdRight.Visible = False
    
    .txtFirstName = ""
    .txtLastName = ""
    .txtAge = ""
    .txtID = ""
    .txtPayType = ""
    .txtWage = ""
    .txtPhoneNumber = ""
End With

End Sub

Private Sub mnuChange_Click()

Dim answer As String
Dim gtg As Boolean
Dim i As Integer

gtg = False

answer = InputBox("Enter an employee's ID or last name.", "Change")

If IsNumeric(answer) Then
    If answer > 0 And answer <= entries Then
        entrytochange = answer
        gtg = True
    Else
        gtg = False
    End If
Else
    For i = 1 To entries
        If StrComp(arremployees(i).lastname, answer, vbTextCompare) = 0 Then
            entrytochange = i
            gtg = True
            Exit For
        End If
    gtg = False
    Next i
End If

If gtg = True Then
    Hide
    With frmDialogue
        .Show
        .cmdChange.Visible = True
        .cmdAdd.Visible = False
        .cmdDelete.Visible = False
        .cmdLeft.Visible = False
        .cmdRight.Visible = False

        .txtFirstName = arremployees(entrytochange).firstname
        .txtLastName = arremployees(entrytochange).lastname
        .txtAge = Str(arremployees(entrytochange).age)
        .txtID = Str(arremployees(entrytochange).id)
        .txtPayType = arremployees(entrytochange).paytype
        .txtWage = Str(arremployees(entrytochange).wage)
        .txtPhoneNumber = arremployees(entrytochange).phonenumber
    End With
Else
    MsgBox "Can't find employee."
End If

End Sub

Private Sub mnuDelete_Click()

Dim answer As String
Dim gtg As Boolean
Dim i As Integer

gtg = False

answer = InputBox("Enter an employee's ID or last name.", "Delete")

If IsNumeric(answer) Then
    If answer > 0 And answer <= entries Then
        entrytodelete = answer
        gtg = True
    Else
        gtg = False
    End If
Else
    For i = 1 To entries
        If StrComp(arremployees(i).lastname, answer, vbTextCompare) = 0 Then
            entrytodelete = i
            gtg = True
            Exit For
        End If
    gtg = False
    Next i
End If

If gtg = True Then
    Hide
    With frmDialogue
        .Show
        .cmdChange.Visible = False
        .cmdAdd.Visible = False
        .cmdDelete.Visible = True
        .cmdLeft.Visible = False
        .cmdRight.Visible = False

        .txtFirstName = arremployees(entrytodelete).firstname
        .txtLastName = arremployees(entrytodelete).lastname
        .txtAge = Str(arremployees(entrytodelete).age)
        .txtID = Str(arremployees(entrytodelete).id)
        .txtPayType = arremployees(entrytodelete).paytype
        .txtWage = Str(arremployees(entrytodelete).wage)
        .txtPhoneNumber = arremployees(entrytodelete).phonenumber
    End With
Else
    MsgBox "Can't find employee."
End If

End Sub

Private Sub mnuExit_Click()

End

End Sub

Private Sub mnuOpen_Click()

Dim done As Boolean
Dim answer As String
Dim i As Integer

i = 1
done = False
Do While done = False
    cndOpen.Filter = "Data (*.dat)|*.dat|All Files (*.*)|*.*"
    cndOpen.ShowOpen
    path = cndOpen.FileName
    cndOpen.FileName = ""
    
    If path = "" Then
        Exit Do
    Else
        answer = MsgBox(path, vbYesNo, "Is this the file?")
        If answer = vbYes Then
            Open path For Random Access Read As #1
                Do While Not EOF(1)
                    Get #1, , arremployees(i)
                    i = i + 1
                Loop
                entries = i - 1
            Close #1
            
            done = True
            Caption = "Employee - " + cndOpen.FileTitle
            RefreshList
        End If
    End If
Loop

End Sub

Private Sub mnuRoledex_Click()

Hide
With frmDialogue
    .Show
    
    .cmdLeft.Visible = True
    .cmdRight.Visible = True
    .cmdAdd.Visible = False
    .cmdChange.Visible = False
    .cmdDelete.Visible = False
    .lblTitle = "Dialogue"
    
    .txtFirstName = arremployees(1).firstname
    .txtLastName = arremployees(1).lastname
    .txtAge = Str(arremployees(1).age)
    .txtID = Str(arremployees(1).id)
    .txtPayType = arremployees(1).paytype
    .txtWage = Format(arremployees(1).wage, "Currency")
    .txtPhoneNumber = arremployees(1).phonenumber
End With

End Sub

Private Sub mnuSave_Click()

Dim i As Integer

If path = "" Then
    mnuSaveAs_Click
Else
    Open path For Random Access Write As #1
        For i = 1 To entries
            Put #1, i, arremployees(i)
        Next i
    Close #1
End If

End Sub

Private Sub mnuSaveAs_Click()

Dim done As Boolean
Dim answer As String
Dim i As Integer

done = False
Do While done = False
    cndSaveAs.Filter = "Data (*.dat)|*.dat|All Files (*.*)|*.*"
    cndSaveAs.ShowSave
    path = cndSaveAs.FileName
    cndSaveAs.FileName = ""
    
    If path = "" Then
        Exit Do
    Else
        answer = MsgBox(path, vbYesNo, "Is this the file?")
        If answer = vbYes Then
            Open path For Random Access Write As #1
                For i = 1 To entries
                    Put #1, i, arremployees(i)
                Next i
            Close #1
            
            done = True
            Caption = "Employee - " + cndSaveAs.FileTitle
        End If
    End If
Loop

End Sub
