VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmLetterWizard 
   Caption         =   "Letter Wizard"
   ClientHeight    =   6600
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6885
   LinkTopic       =   "Form1"
   ScaleHeight     =   6600
   ScaleWidth      =   6885
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cndSaveAs 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Text (*.txt)|*.txt|All Files (*.*)|*.*"
   End
   Begin VB.TextBox txtLetterWizard 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   7000
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   7000
   End
   Begin MSComDlg.CommonDialog cndOpen 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Text (*.txt)|*.txt|All Files (*.*)|*.*"
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuOpen 
         Caption         =   "Open..."
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save As..."
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu mnuGreetings 
         Caption         =   "Greetings"
         Begin VB.Menu mnuToWhom 
            Caption         =   "To whom..."
         End
         Begin VB.Menu mnuSir 
            Caption         =   "Dear sir..."
         End
         Begin VB.Menu mnuMadam 
            Caption         =   "Dear madam..."
         End
      End
      Begin VB.Menu mnuClosings 
         Caption         =   "Closings"
         Begin VB.Menu mnuKind 
            Caption         =   "Kind regards..."
         End
         Begin VB.Menu mnuSincerely 
            Caption         =   "Sincerely..."
         End
         Begin VB.Menu mnuLove 
            Caption         =   "Love you..."
         End
      End
   End
End
Attribute VB_Name = "frmLetterWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim closingname As String

Sub GetName()

closingname = InputBox("Please enter your name:", "Input")

End Sub

Private Sub Form_Unload(Cancel As Integer)

Hide

frmMainMenu.Show

End Sub

Private Sub mnuExit_Click()

Hide

frmMainMenu.Show

End Sub

Private Sub mnuKind_Click()

GetName

txtLetterWizard = txtLetterWizard + vbCrLf _
+ vbCrLf _
+ "Kind regards," + vbCrLf _
+ closingname

End Sub

Private Sub mnuLove_Click()

GetName

txtLetterWizard = txtLetterWizard + vbCrLf _
+ vbCrLf _
+ "Love you," + vbCrLf _
+ closingname

End Sub

Private Sub mnuMadam_Click()

txtLetterWizard = "Dear madam, " + vbCrLf _
+ vbCrLf _
+ txtLetterWizard

End Sub

Private Sub mnuOpen_Click()

Dim done As Boolean
Dim answer As String
Dim path As String

done = False
Do While done = False

    cndOpen.ShowOpen
    path = cndOpen.FileName
    cndOpen.FileName = ""
    
    If path = "" Then
            
        Exit Do
        txtLetterWizard.SetFocus
        
    Else
        
        answer = MsgBox(path, vbYesNo, "Is this the file?")
        
        If answer = vbYes Then
            Open path For Input As #1

                Dim FileSize As Long
                FileSize = LOF(1)
                txtLetterWizard = Input(FileSize, 1)

            Close #1

            done = True
            frmLetterWizard.Caption = "LetterWizard - " + cndOpen.FileTitle
            
        End If
        
    End If

Loop

End Sub

Private Sub mnuSaveAs_Click()

Dim path As String
Dim done As Boolean
Dim answer As String

done = False
Do While done = False

    cndSaveAs.ShowSave
    path = cndSaveAs.FileName
    cndSaveAs.FileName = ""
    
    If path = "" Then
    
        txtLetterWizard.SetFocus
        Exit Do
        
    Else
    
        answer = MsgBox(path, vbYesNo, "Is this the file?")
        If answer = vbYes Then

            Open path For Output As #1

                Print #1, txtLetterWizard

            Close #1
            
            done = True
            frmLetterWizard.Caption = "Letter Wizard - " + cndSaveAs.FileTitle
            
        End If
        
    End If

Loop

End Sub

Private Sub mnuSincerely_Click()

GetName

txtLetterWizard = txtLetterWizard + vbCrLf _
+ vbCrLf _
+ "Sincerely," + vbCrLf _
+ closingname

End Sub

Private Sub mnuSir_Click()

txtLetterWizard = "Dear sir, " + vbCrLf _
+ vbCrLf _
+ txtLetterWizard

End Sub

Private Sub mnuToWhom_Click()

txtLetterWizard = "To whom it may concern, " + vbCrLf _
+ vbCrLf _
+ txtLetterWizard

End Sub
