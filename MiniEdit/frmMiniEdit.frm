VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMiniEdit 
   Caption         =   "MiniEdit"
   ClientHeight    =   7395
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7395
   ScaleWidth      =   9000
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cndSaveAs 
      Left            =   8280
      Top             =   900
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Save As"
      InitDir         =   "D:\CP1\VB6 Project Files\MiniEdit\saves"
   End
   Begin MSComDlg.CommonDialog cndOpen 
      Left            =   8280
      Top             =   300
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Open"
      InitDir         =   "D:\CP1\VB6 Project Files\MiniEdit\saves"
   End
   Begin VB.TextBox txtEdit 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8295
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   9240
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuNew 
         Caption         =   "New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
         Shortcut        =   ^S
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
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuCut 
         Caption         =   "Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "Paste"
         Enabled         =   0   'False
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuUndoText 
         Caption         =   "Undo Text"
      End
      Begin VB.Menu mnuSelectAll 
         Caption         =   "Select All"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu mnuFont 
         Caption         =   "Font"
         Begin VB.Menu mnuArial 
            Caption         =   "Arial"
         End
         Begin VB.Menu mnuComicSans 
            Caption         =   "Comic Sans"
         End
         Begin VB.Menu mnuMSSansSerif 
            Caption         =   "MS Sans Serif"
         End
      End
      Begin VB.Menu mnuFontSize 
         Caption         =   "Font Size..."
      End
      Begin VB.Menu mnuBackgroundColor 
         Caption         =   "Background Color"
         Begin VB.Menu mnuBackRed 
            Caption         =   "Red"
         End
         Begin VB.Menu mnuBackGreen 
            Caption         =   "Green"
         End
         Begin VB.Menu mnuBackBlue 
            Caption         =   "Blue"
         End
         Begin VB.Menu mnuBackWhite 
            Caption         =   "White"
         End
         Begin VB.Menu mnuBackBlack 
            Caption         =   "Black"
         End
      End
      Begin VB.Menu mnuFontColor 
         Caption         =   "Font Color"
         Begin VB.Menu mnuForeWhite 
            Caption         =   "White"
         End
         Begin VB.Menu mnuForeBlack 
            Caption         =   "Black"
         End
         Begin VB.Menu mnuForeRed 
            Caption         =   "Red"
         End
         Begin VB.Menu mnuForeGreen 
            Caption         =   "Green"
         End
         Begin VB.Menu mnuForeBlue 
            Caption         =   "Blue"
         End
      End
   End
End
Attribute VB_Name = "frmMiniEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Path As String
Dim FormatPath As String
Dim UserFont As String
Dim UserFontSize As String
Dim BackgroundColor As String
Dim FontColor As String
Dim PreviousFormatType As Integer
Dim PreviousFormat As String
Dim CurrentFormatType As String
Dim OldBody As String
Dim NewBody As String

Sub OpenFile()

Dim Done As Boolean
Dim Answer As String
Dim FileName As String

Done = False
Do While Done = False

    cndOpen.Filter = "Text (*.txt)|*.txt|All Files (*.*)|*.*"
    cndOpen.ShowOpen
    Path = cndOpen.FileName
    cndOpen.FileName = ""
    
    If Path = "" Then
            
        Exit Do
        txtEdit.SetFocus
        
    Else
        
        Answer = MsgBox(Path, vbYesNo, "Is this the file?")
        
        If Answer = vbYes Then
            Open Path For Input As #1

                Dim FileSize As Long
                FileSize = LOF(1)
                txtEdit = Input(FileSize, 1)

            Close #1

            OldBody = txtEdit
            Done = True
            frmMiniEdit.Caption = "MiniEdit - " + cndOpen.FileTitle
            FormatLoad
            
        End If
        
    End If

Loop

End Sub

Sub SaveFile()

If Path = "" Then

    SaveAsFile
    
Else

    Open Path For Output As #1

        Print #1, txtEdit

    Close #1
    
    FormatSave
    
End If

End Sub

Sub SaveAsFile()

Dim Done As Boolean
Dim FileName As String
Dim Answer As String

Done = False
Do While Done = False

    cndSaveAs.Filter = "Text (*.txt)|*.txt|All Files (*.*)|*.*"
    cndSaveAs.ShowSave
    Path = cndSaveAs.FileName
    cndSaveAs.FileName = ""
    
    If Path = "" Then
    
        txtEdit.SetFocus
        Exit Do
        
    Else
    
        Answer = MsgBox(Path, vbYesNo, "Is this the file?")
        If Answer = vbYes Then

            Open Path For Output As #1

                Print #1, txtEdit

            Close #1
            
            Done = True
            frmMiniEdit.Caption = "MiniEdit - " + cndSaveAs.FileTitle
            FormatSave
            
        End If

        OldBody = txtEdit
        
    End If

Loop

End Sub

Sub FormatSave()

Dim FormatFileSave As String

FormatFileSave = UserFont + vbCrLf _
+ UserFontSize + vbCrLf _
+ BackgroundColor + vbCrLf _
+ FontColor

FormatPath = Path + ".format"

Open FormatPath For Output As #1

    Print #1, FormatFileSave

Close #1

End Sub

Sub FormatLoad()

Dim CurrentLine As String
Dim i As Integer

FormatClear

i = 1
FormatPath = Path + ".format"

Open FormatPath For Input As #1

    Do While Not EOF(1)
        
        Line Input #1, CurrentLine
        
        If i = 1 Then
            
            txtEdit.Font = CurrentLine
            i = i + 1
        
        ElseIf i = 2 Then
        
            txtEdit.FontSize = CurrentLine
            i = i + 1
        
        ElseIf i = 3 Then
        
            If CurrentLine = 1 Then
                
                txtEdit.BackColor = vbWhite
                
            ElseIf CurrentLine = 2 Then
            
                txtEdit.BackColor = vbBlack
                
            ElseIf CurrentLine = 3 Then
            
                txtEdit.BackColor = vbRed
                
            ElseIf CurrentLine = 4 Then
            
                txtEdit.BackColor = vbGreen
                
            ElseIf CurrentLine = 5 Then
            
                txtEdit.BackColor = vbBlue
                
            End If
            i = i + 1
        
        ElseIf i = 4 Then
            
            If CurrentLine = 1 Then
                
                txtEdit.ForeColor = vbWhite
                
            ElseIf CurrentLine = 2 Then
            
                txtEdit.ForeColor = vbBlack
                
            ElseIf CurrentLine = 3 Then
            
                txtEdit.ForeColor = vbRed
                
            ElseIf CurrentLine = 4 Then
            
                txtEdit.ForeColor = vbGreen
                
            ElseIf CurrentLine = 5 Then
            
                txtEdit.ForeColor = vbBlue
                
            End If
            
            i = 1 + 1
        
        End If
    
    Loop
    
Close #1

End Sub

Sub FormatClear()

UserFont = ""
UserFontSize = ""
BackgroundColor = ""
FontColor = ""
txtEdit.Font = "MS Sans Serif"
txtEdit.FontSize = "10"
txtEdit.BackColor = vbWhite
txtEdit.ForeColor = vbBlack
PreviousFormatType = 0
PreviousFormat = ""
CurrentFormatType = 0

End Sub

Private Sub Form_Load()

OldBody = ""
NewBody = ""
UserFont = "MS Sans Serif"
UserFontSize = "10"
BackgroundColor = "1"
FontColor = "2"
PreviousFormatType = 0
CurrentFormatType = 0

End Sub

Private Sub Form_Resize()

txtEdit.Width = frmMiniEdit.ScaleWidth
txtEdit.Height = frmMiniEdit.ScaleHeight

End Sub

Private Sub Form_Unload(Cancel As Integer)

Dim Answer As String

If OldBody <> NewBody Then

    Answer = MsgBox("Do you want to save your changes?", vbYesNo, "Exit")
    
    If Answer = vbNo Then
    
        End
    
    Else
        
        SaveFile
        
    End If
    
End If

End Sub

Private Sub mnuArial_Click()

If CurrentFormatType = 1 Then
    
    PreviousFormat = txtEdit.Font
    PreviousFormatType = 1

ElseIf CurrentFormatType = 2 Then

    PreviousFormat = txtEdit.FontSize
    PreviousFormatType = 2
    
ElseIf CurrentFormatType = 3 Then

    PreviousFormat = txtEdit.BackColor
    PreviousFormatType = 3
    
ElseIf CurrentFormatType = 4 Then

    PreviousFormat = txtEdit.ForeColor
    PreviousFormatType = 4
    
End If

txtEdit.Font = "Arial"
UserFont = "Arial"
CurrentFormatType = 1

End Sub

Private Sub mnuBackBlack_Click()

If CurrentFormatType = 1 Then
    
    PreviousFormat = txtEdit.Font
    PreviousFormatType = 1

ElseIf CurrentFormatType = 2 Then

    PreviousFormat = txtEdit.FontSize
    PreviousFormatType = 2
    
ElseIf CurrentFormatType = 3 Then

    PreviousFormat = txtEdit.BackColor
    PreviousFormatType = 3
    
ElseIf CurrentFormatType = 4 Then

    PreviousFormat = txtEdit.ForeColor
    PreviousFormatType = 4
    
End If

txtEdit.BackColor = vbBlack
BackgroundColor = "2"
CurrentFormatType = 3

End Sub

Private Sub mnuBackBlue_Click()

If CurrentFormatType = 1 Then
    
    PreviousFormat = txtEdit.Font
    PreviousFormatType = 1

ElseIf CurrentFormatType = 2 Then

    PreviousFormat = txtEdit.FontSize
    PreviousFormatType = 2
    
ElseIf CurrentFormatType = 3 Then

    PreviousFormat = txtEdit.BackColor
    PreviousFormatType = 3
    
ElseIf CurrentFormatType = 4 Then

    PreviousFormat = txtEdit.ForeColor
    PreviousFormatType = 4
    
End If

txtEdit.BackColor = vbBlue
BackgroundColor = "5"
CurrentFormatType = 3

End Sub

Private Sub mnuBackGreen_Click()

If CurrentFormatType = 1 Then
    
    PreviousFormat = txtEdit.Font
    PreviousFormatType = 1

ElseIf CurrentFormatType = 2 Then

    PreviousFormat = txtEdit.FontSize
    PreviousFormatType = 2
    
ElseIf CurrentFormatType = 3 Then

    PreviousFormat = txtEdit.BackColor
    PreviousFormatType = 3
    
ElseIf CurrentFormatType = 4 Then

    PreviousFormat = txtEdit.ForeColor
    PreviousFormatType = 4
    
End If

txtEdit.BackColor = vbGreen
BackgroundColor = "4"
CurrentFormatType = 3

End Sub

Private Sub mnuBackRed_Click()

If CurrentFormatType = 1 Then
    
    PreviousFormat = txtEdit.Font
    PreviousFormatType = 1

ElseIf CurrentFormatType = 2 Then

    PreviousFormat = txtEdit.FontSize
    PreviousFormatType = 2
    
ElseIf CurrentFormatType = 3 Then

    PreviousFormat = txtEdit.BackColor
    PreviousFormatType = 3
    
ElseIf CurrentFormatType = 4 Then

    PreviousFormat = txtEdit.ForeColor
    PreviousFormatType = 4
    
End If

txtEdit.BackColor = vbRed
BackgroundColor = "3"
CurrentFormatType = 3

End Sub

Private Sub mnuBackWhite_Click()

If CurrentFormatType = 1 Then
    
    PreviousFormat = txtEdit.Font
    PreviousFormatType = 1

ElseIf CurrentFormatType = 2 Then

    PreviousFormat = txtEdit.FontSize
    PreviousFormatType = 2
    
ElseIf CurrentFormatType = 3 Then

    PreviousFormat = txtEdit.BackColor
    PreviousFormatType = 3
    
ElseIf CurrentFormatType = 4 Then

    PreviousFormat = txtEdit.ForeColor
    PreviousFormatType = 4
    
End If

txtEdit.BackColor = vbWhite
BackgroundColor = "1"
CurrentFormatType = 3

End Sub

Private Sub mnuExit_Click()
Dim Answer As String

If OldBody <> NewBody Then

    Answer = MsgBox("Do you want to save your changes?", vbYesNo, "Exit")
    
    If Answer = vbNo Then
    
        End
    
    Else
        
        SaveFile
        End
        
    End If
    
Else

    End
    
End If

End Sub

Private Sub mnuComicSans_Click()

If CurrentFormatType = 1 Then
    
    PreviousFormat = txtEdit.Font
    PreviousFormatType = 1

ElseIf CurrentFormatType = 2 Then

    PreviousFormat = txtEdit.FontSize
    PreviousFormatType = 2
    
ElseIf CurrentFormatType = 3 Then

    PreviousFormat = txtEdit.BackColor
    PreviousFormatType = 3
    
ElseIf CurrentFormatType = 4 Then

    PreviousFormat = txtEdit.ForeColor
    PreviousFormatType = 4
    
End If

txtEdit.Font = "Comic Sans MS"
UserFont = "Comic Sans MS"
CurrentFormatType = 1


End Sub

Private Sub mnuCopy_Click()

Clipboard.SetText txtEdit.SelText
mnuPaste.Enabled = True

End Sub

Private Sub mnuCut_Click()

Clipboard.SetText txtEdit.SelText
txtEdit.SelText = ""
mnuPaste.Enabled = True

End Sub

Private Sub mnuFontSize_Click()

Dim Answer As String

If CurrentFormatType = 1 Then
    
    PreviousFormat = txtEdit.Font
    PreviousFormatType = 1

ElseIf CurrentFormatType = 2 Then

    PreviousFormat = txtEdit.FontSize
    PreviousFormatType = 2
    
ElseIf CurrentFormatType = 3 Then

    PreviousFormat = txtEdit.BackColor
    PreviousFormatType = 3
    
ElseIf CurrentFormatType = 4 Then

    PreviousFormat = txtEdit.ForeColor
    PreviousFormatType = 4
    
End If

Answer = InputBox("Enter font size:", "Font Size", "10")
txtEdit.FontSize = Answer
UserFontSize = Answer
CurrentFormatType = 2

End Sub

Private Sub mnuForeBlack_Click()

If CurrentFormatType = 1 Then
    
    PreviousFormat = txtEdit.Font
    PreviousFormatType = 1

ElseIf CurrentFormatType = 2 Then

    PreviousFormat = txtEdit.FontSize
    PreviousFormatType = 2
    
ElseIf CurrentFormatType = 3 Then

    PreviousFormat = txtEdit.BackColor
    PreviousFormatType = 3
    
ElseIf CurrentFormatType = 4 Then

    PreviousFormat = txtEdit.ForeColor
    PreviousFormatType = 4
    
End If

txtEdit.ForeColor = vbBlack
FontColor = "2"
CurrentFormatType = 4

End Sub

Private Sub mnuForeBlue_Click()

If CurrentFormatType = 1 Then
    
    PreviousFormat = txtEdit.Font
    PreviousFormatType = 1

ElseIf CurrentFormatType = 2 Then

    PreviousFormat = txtEdit.FontSize
    PreviousFormatType = 2
    
ElseIf CurrentFormatType = 3 Then

    PreviousFormat = txtEdit.BackColor
    PreviousFormatType = 3
    
ElseIf CurrentFormatType = 4 Then

    PreviousFormat = txtEdit.ForeColor
    PreviousFormatType = 4
    
End If

txtEdit.ForeColor = vbBlue
FontColor = "5"
CurrentFormatType = 4

End Sub

Private Sub mnuForeGreen_Click()

If CurrentFormatType = 1 Then
    
    PreviousFormat = txtEdit.Font
    PreviousFormatType = 1

ElseIf CurrentFormatType = 2 Then

    PreviousFormat = txtEdit.FontSize
    PreviousFormatType = 2
    
ElseIf CurrentFormatType = 3 Then

    PreviousFormat = txtEdit.BackColor
    PreviousFormatType = 3
    
ElseIf CurrentFormatType = 4 Then

    PreviousFormat = txtEdit.ForeColor
    PreviousFormatType = 4
    
End If

txtEdit.ForeColor = vbGreen
FontColor = "4"
CurrentFormatType = 4

End Sub

Private Sub mnuForeRed_Click()

If CurrentFormatType = 1 Then
    
    PreviousFormat = txtEdit.Font
    PreviousFormatType = 1

ElseIf CurrentFormatType = 2 Then

    PreviousFormat = txtEdit.FontSize
    PreviousFormatType = 2
    
ElseIf CurrentFormatType = 3 Then

    PreviousFormat = txtEdit.BackColor
    PreviousFormatType = 3
    
ElseIf CurrentFormatType = 4 Then

    PreviousFormat = txtEdit.ForeColor
    PreviousFormatType = 4
    
End If

txtEdit.ForeColor = vbRed
FontColor = "3"
CurrentFormatType = 4

End Sub

Private Sub mnuForeWhite_Click()

If CurrentFormatType = 1 Then
    
    PreviousFormat = txtEdit.Font
    PreviousFormatType = 1

ElseIf CurrentFormatType = 2 Then

    PreviousFormat = txtEdit.FontSize
    PreviousFormatType = 2
    
ElseIf CurrentFormatType = 3 Then

    PreviousFormat = txtEdit.BackColor
    PreviousFormatType = 3
    
ElseIf CurrentFormatType = 4 Then

    PreviousFormat = txtEdit.ForeColor
    PreviousFormatType = 4
    
End If

txtEdit.ForeColor = vbWhite
FontColor = "1"
CurrentFormatType = 4

End Sub

Private Sub mnuMSSansSerif_Click()

If CurrentFormatType = 1 Then
    
    PreviousFormat = txtEdit.Font
    PreviousFormatType = 1

ElseIf CurrentFormatType = 2 Then

    PreviousFormat = txtEdit.FontSize
    PreviousFormatType = 2
    
ElseIf CurrentFormatType = 3 Then

    PreviousFormat = txtEdit.BackColor
    PreviousFormatType = 3
    
ElseIf CurrentFormatType = 4 Then

    PreviousFormat = txtEdit.ForeColor
    PreviousFormatType = 4
    
End If

txtEdit.Font = "MS Sans Serif"
UserFont = "MS Sans Serif"
CurrentFormatType = 1
End Sub

Private Sub mnuNew_Click()

Dim Answer As String

If OldBody <> NewBody Then
    
    Answer = MsgBox("Do you want to save your changes?", vbYesNo, "New")
    
    If Answer = vbNo Then
    
        txtEdit = ""
        Path = ""
        OldBody = ""
        FormatClear
        frmMiniEdit.Caption = "MiniEdit"
        
    Else
    
        SaveFile
        txtEdit = ""
        Path = ""
        OldBody = ""
        FormatClear
        frmMiniEdit.Caption = "MiniEdit"
        
    End If

Else

    txtEdit = ""
    Path = ""
    OldBody = ""
    FormatClear
    frmMiniEdit.Caption = "MiniEdit"
    
End If

End Sub

Private Sub mnuOpen_Click()

Dim Answer As String

If OldBody <> NewBody Then

    Answer = MsgBox("Do you want to save your changes?", vbYesNo, "Open")
    
    If Answer = vbNo Then
    
        OpenFile
    
    Else
    
        SaveFile
        OpenFile
    
    End If

Else

    OpenFile

End If

End Sub

Private Sub mnuPaste_Click()

txtEdit.SelText = Clipboard.GetText()

End Sub

Private Sub mnuSave_Click()

SaveFile

End Sub

Private Sub mnuSaveAs_Click()

SaveAsFile

End Sub

Private Sub mnuSelectAll_Click()

txtEdit.SelStart = 0
txtEdit.SelLength = Len(txtEdit)

End Sub

Private Sub mnuUndoText_Click()

SendKeys ("^Z")

End Sub

Private Sub txtEdit_Change()

NewBody = txtEdit

End Sub

