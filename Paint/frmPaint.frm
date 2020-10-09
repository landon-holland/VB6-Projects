VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPaint 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "Paint"
   ClientHeight    =   8085
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8385
   LinkTopic       =   "Form1"
   ScaleHeight     =   8085
   ScaleWidth      =   8385
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cndSaveAs 
      Left            =   7800
      Top             =   540
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Save As"
      Filter          =   "Data (*.dat)|*.dat|All Files (*.*)|*.*"
      InitDir         =   "D:\CP1\VB6 Project Files\Paint\saves"
   End
   Begin MSComDlg.CommonDialog cndOpen 
      Left            =   7800
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Open"
      Filter          =   "Data (*.dat)|*.dat|All Files (*.*)|*.*"
      InitDir         =   "D:\CP1\VB6 Project Files\Paint\saves"
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
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
   End
End
Attribute VB_Name = "frmPaint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim gfdrawing As Integer
Dim npoints As Long
Dim gsngx(1 To 10000) As Single
Dim ogsngx(1 To 10000) As Single
Dim gsngy(1 To 10000) As Single
Dim ogsngy(1 To 10000) As Single
Dim gintr(1 To 10000) As Integer
Dim gintg(1 To 10000) As Integer
Dim gintb(1 To 10000) As Integer
Dim gints(1 To 10000) As Integer
Dim path As String

Sub OpenFile()

Dim i As Integer
Dim done As Boolean
Dim answer As String
Dim filename As String

done = False
Do While done = False

    cndOpen.ShowOpen
    path = cndOpen.filename
    cndOpen.filename = ""

    If path = "" Then
    
        Exit Do
        
    Else
    
        answer = MsgBox(path, vbYesNo, "Is this the file?")

        If answer = vbYes Then
        
            
            Open path For Binary Access Read As #1

                Get #1, , npoints
    
                For i = 1 To npoints
    
                    Get #1, , gsngx(i)
                    Get #1, , gsngy(i)
                    Get #1, , gintr(i)
                    Get #1, , gintg(i)
                    Get #1, , gintb(i)
                    Get #1, , gints(i)

                Next i
    
            Close #1
            
            For i = 1 To 10000
            
                ogsngx(i) = gsngx(i)
                ogsngy(i) = gsngy(i)
                
            Next i
            
            frmPaint.Caption = "Paint - " + cndOpen.FileTitle
            frmPaint.Cls
            DrawLines
            done = True
            
        End If
        
    End If
    
Loop

End Sub

Sub SaveFileAs()

Dim i As Integer
Dim done As Boolean
Dim filename As String
Dim answer As String

done = False
Do While done = False

    cndSaveAs.ShowSave
    path = cndSaveAs.filename
    cndSaveAs.filename = ""

    If path = "" Then
    
        Exit Do
        
    Else
    
        answer = MsgBox(path, vbYesNo, "Is this the file?")
        If answer = vbYes Then
        
            Open path For Binary Access Write As #1

                Put #1, , npoints
    
                For i = 1 To npoints
    
                    Put #1, , gsngx(i)
                    Put #1, , gsngy(i)
                    Put #1, , gintr(i)
                    Put #1, , gintg(i)
                    Put #1, , gintb(i)
                    Put #1, , gints(i)
        
                Next i
    
            Close #1
        
        For i = 1 To 10000
        
            ogsngx(i) = gsngx(i)
            ogsngy(i) = gsngy(i)
        
        Next i
        
        frmPaint.Caption = "Paint - " + cndSaveAs.FileTitle
        done = True
        
        End If
        
    End If
    
Loop

End Sub

Sub SaveFile()

Dim i As Integer

If path = "" Then

    SaveFileAs
    
Else

    Open path For Binary Access Write As #1

        Put #1, , npoints
    
        For i = 1 To npoints
    
            Put #1, , gsngx(i)
            Put #1, , gsngy(i)
            Put #1, , gintr(i)
            Put #1, , gintg(i)
            Put #1, , gintb(i)
            Put #1, , gints(i)
        
        Next i
        
    Close #1
    
    For i = 1 To 10000
        
        ogsngx(i) = gsngx(i)
        ogsngy(i) = gsngy(i)
        
    Next i
    
End If

End Sub

Sub DrawCircle(X As Single, Y As Single, R As Integer, G As Integer, B As Integer, S As Integer)

Circle (X, Y), 1, RGB(R, G, B)

If npoints < 10000 Then

    npoints = npoints + 1
    
    gsngx(npoints) = X
    gsngy(npoints) = Y
    gintr(npoints) = R
    gintg(npoints) = G
    gintb(npoints) = B
    gints(npoints) = S
End If

End Sub

Sub DrawLines()

Dim i As Integer
Dim loopagain As Integer

Circle (gsngx(1), gsngy(1)), 1, RGB(gintr(1), gintg(1), gintb(1))

For i = 2 To npoints

    If gsngx(i) = -1 Or gsngy(i) = -1 Or gintr(i) = -1 Or gintg(i) = -1 Or gintb(i) = -1 Or gints(i) = -1 Then
    
        loopagain = 1
        
        Circle (0, 0), 1, RGB(255, 255, 255)
        
    ElseIf loopagain = 1 Then
        
        loopagain = 2
        
        Circle (0, 0), 1, RGB(255, 255, 255)
    
    Else
    
        If loopagain = 2 Then
        
            frmPaint.DrawWidth = gints(i)
            
            Circle (gsngx(i), gsngy(i)), 1, RGB(gintr(i), gintg(i), gintb(i))
            loopagain = 0
        
        Else
        
            frmPaint.DrawWidth = gints(i)
            Line -(gsngx(i), gsngy(i)), RGB(gintr(i), gintg(i), gintb(i))
            Circle (gsngx(i), gsngy(i)), 1, RGB(gintr(i), gintg(i), gintb(i))
            
        End If
        
    End If
    
Next i

frmPaint.DrawWidth = ds

End Sub

Private Sub Form_Load()

gfdrawing = 0

dr = 0
dg = 0
db = 0
ds = 5
frmPaint.DrawWidth = ds

npoints = 0

path = ""

drawtype = 1
'1 paintbrush
'2 line
'3 box

linepoint = 1

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If drawtype = 1 Then

    gfdrawing = 1

    DrawCircle X, Y, dr, dg, db, ds
    
ElseIf drawtype = 2 Then

    gfdrawing = 1
    
    If linepoint = 1 Then
    
        DrawCircle X, Y, dr, dg, db, ds
    
        linepoint = 2
        
    ElseIf linepoint = 2 Then
    
        frmPaint.DrawWidth = ds
    
        Line -(X, Y), RGB(dr, dg, db)
        DrawCircle X, Y, dr, dg, db, ds
        
        linepoint = 1
        
    End If
    
ElseIf drawtype = 3 Then

    gfdrawing = 1
    
    If linepoint = 1 Then
    
        DrawCircle X, Y, dr, dg, db, ds
    
        linepoint = 2
        
    ElseIf linepoint = 2 Then
    
        frmPaint.DrawWidth = ds
    
        Line -(X, Y), RGB(dr, dg, db)
        
        linepoint = 1
        
    End If
    
End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If gfdrawing = 1 Then

    If drawtype = 1 Then

        frmPaint.DrawWidth = ds
        Line -(X, Y), RGB(dr, dg, db)
        DrawCircle X, Y, dr, dg, db, ds
        
    End If
    
End If

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

gfdrawing = 0

If drawtype = 1 Then

    If npoints < 10000 Then

        npoints = npoints + 1
    
        gsngx(npoints) = -1
        gsngy(npoints) = -1
        gintr(npoints) = -1
        gintg(npoints) = -1
        gintb(npoints) = -1
        gints(npoints) = -1
        
    End If
    
ElseIf drawtype = 2 Or drawtype = 3 Then

    If linepoint = 1 Then
        
        If npoints < 10000 Then

            npoints = npoints + 1
    
            gsngx(npoints) = -1
            gsngy(npoints) = -1
            gintr(npoints) = -1
            gintg(npoints) = -1
            gintb(npoints) = -1
            gints(npoints) = -1
        
        End If
        
        DrawCircle X, Y, dr, dg, db, ds
        
    End If
    
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

Dim i As Integer
Dim answer As String

For i = 1 To 10000

    If (ogsngx(i) <> gsngx(i)) Or (ogsngy(i) <> gsngy(i)) Then
    
        answer = MsgBox("Do you want to save your changes?", vbYesNo, "Exit")
        
        If answer = vbNo Then
        
            End
            
        Else
        
            SaveFile
            End
            
        End If
        
    ElseIf i = 10000 Then
    
        End
        
    End If
    
Next i

End Sub

Private Sub mnuExit_Click()

Dim i As Integer
Dim answer As String

For i = 1 To 10000

    If (ogsngx(i) <> gsngx(i)) Or (ogsngy(i) <> gsngy(i)) Then
    
        answer = MsgBox("Do you want to save your changes?", vbYesNo, "Exit")
        
        If answer = vbNo Then
        
            End
            
        Else
        
            SaveFile
            End
            
        End If
        
    ElseIf i = 10000 Then
    
        End
        
    End If
    
Next i

End Sub

Private Sub mnuNew_Click()

Dim i As Integer
Dim j As Integer
Dim answer As String

For i = 1 To 10000

    If (ogsngx(i) <> gsngx(i)) Or (ogsngy(i) <> gsngy(i)) Then
    
        answer = MsgBox("Do you want to save your changes?", vbYesNo, "New")
        
        If answer = vbNo Then
        
            frmPaint.Cls
            frmPaint.Caption = "Paint"
            
            For j = 1 To 10000
            
                ogsngx(j) = 0
                ogsngy(j) = 0
                
            Next j
            
        Else
        
            SaveFile
        
            frmPaint.Cls
            frmPaint.Caption = "Paint"
            
            For j = 1 To 10000
            
                ogsngx(j) = 0
                ogsngy(j) = 0
                
            Next j
            
        End If
        
    ElseIf i = 10000 Then
    
        frmPaint.Cls
        frmPaint.Caption = "Paint"
            
        For j = 1 To 10000
            
            ogsngx(j) = 0
            ogsngy(j) = 0
                
        Next j
        
    End If
    
Next i

End Sub

Private Sub mnuOpen_Click()

Dim i As Integer
Dim answer As String

For i = 1 To 10000

    If (ogsngx(i) <> gsngx(i)) Or (ogsngy(i) <> gsngy(i)) Then
    
        answer = MsgBox("Do you want to save your changes?", vbYesNo, "Open")
        
        If answer = vbNo Then
        
            OpenFile
            
        Else
        
            SaveFile
            OpenFile
            
        End If
        
    ElseIf i = 10000 Then
    
        OpenFile
    
    End If
    
Next i

End Sub

Private Sub mnuOptions_Click()

frmOptions.Show

End Sub

Private Sub mnuSave_Click()

SaveFile

End Sub

Private Sub mnuSaveAs_Click()

SaveFileAs

End Sub
