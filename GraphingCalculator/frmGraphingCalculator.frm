VERSION 5.00
Begin VB.Form frmGraphingCalculator 
   Caption         =   "Graphing Calculator"
   ClientHeight    =   6045
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4470
   LinkTopic       =   "Form1"
   ScaleHeight     =   6045
   ScaleWidth      =   4470
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   315
      Left            =   2220
      TabIndex        =   6
      Top             =   5460
      Width           =   2175
   End
   Begin VB.CommandButton cmdGraph 
      Caption         =   "Graph"
      Height          =   315
      Left            =   60
      TabIndex        =   5
      Top             =   5460
      Width           =   2175
   End
   Begin VB.TextBox txt12 
      Appearance      =   0  'Flat
      Height          =   435
      Left            =   2220
      TabIndex        =   2
      Top             =   4860
      Width           =   2175
   End
   Begin VB.TextBox txt11 
      Appearance      =   0  'Flat
      Height          =   435
      Left            =   60
      TabIndex        =   1
      Top             =   4860
      Width           =   2175
   End
   Begin VB.PictureBox picGraph 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4335
      Left            =   60
      ScaleHeight     =   4305
      ScaleWidth      =   4305
      TabIndex        =   0
      Top             =   60
      Width           =   4335
   End
   Begin VB.Label lblY 
      Alignment       =   2  'Center
      Caption         =   "Y"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2220
      TabIndex        =   4
      Top             =   4500
      Width           =   2175
   End
   Begin VB.Label lblX1 
      Alignment       =   2  'Center
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   60
      TabIndex        =   3
      Top             =   4500
      Width           =   2175
   End
   Begin VB.Menu mnuDrawtype 
      Caption         =   "Drawtype"
      Begin VB.Menu mnuLine 
         Caption         =   "Line"
      End
      Begin VB.Menu mnuCircle 
         Caption         =   "Circle"
      End
      Begin VB.Menu mnuParabola 
         Caption         =   "Parabola"
      End
      Begin VB.Menu mnuBox 
         Caption         =   "Box"
      End
      Begin VB.Menu mnuAbs 
         Caption         =   "Absolute Value"
      End
   End
End
Attribute VB_Name = "frmGraphingCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim p As Integer
Dim dt As Integer
Dim g As Integer
Dim lx1 As Double
Dim ly1 As Double
Dim ph As Double
Dim pk As Double
Dim ch As Double
Dim ck As Double
Dim bx1 As Double
Dim by1 As Double
Dim ah As Double
Dim ak As Double

Sub Clear()

p = 0
g = 0
lx1 = 0
ly1 = 0
ph = 0
pk = 0
ch = 0
ck = 0

picGraph.Cls
DrawScale

txt11 = ""
txt12 = ""

End Sub

Sub DrawScale()

Dim i As Integer

picGraph.Scale (-10, 10)-(10, -10)
picGraph.Line (-10, 0)-(10, 0), vbRed
picGraph.Line (0, -10)-(0, 10), vbBlue

For i = -10 To 10 Step 1

    picGraph.Line (i, 0.5)-(i, -0.5), vbRed
    picGraph.Line (0.5, i)-(-0.5, i), vbBlue

Next i

End Sub

Private Sub cmdClear_Click()

Clear

End Sub

Private Sub cmdExit_Click()

End

End Sub

Private Sub cmdGraph_Click()

Dim X As Single
Dim Y As Single

X = txt11
Y = txt12

p = p + 1

If dt = 1 And g <> 2 Then
    
    If p = 1 Then
        lx1 = X
        ly1 = Y
        picGraph.Circle (lx1, ly1), 0.25
    ElseIf p = 2 Then
        Dim lm As Double
        Dim lb As Double
        Dim lx2 As Double
        Dim ly2 As Double
        Dim nly1 As Double
        Dim nly2 As Double
        
        p = 0
        
        lx2 = X
        ly2 = Y
        picGraph.Circle (lx2, ly2), 0.25
        
        lm = (ly1 - ly2) / (lx1 - lx2)
        lb = ly1 - (lm * lx1)
        
        nly1 = -10 * lm + lb
        nly2 = 10 * lm + lb
        
        picGraph.Line (-10, nly1)-(10, nly2), vbBlack
        
        g = g + 1
        
        If g = 1 Then
            'txt11 = Str(lm)
            'txt12 = Str(lb)
        ElseIf g = 2 Then
            'txt21 = Str(lm)
            'txt22 = Str(lb)
        End If
    End If
    
ElseIf dt = 2 And g <> 2 Then
    
    If p = 1 Then
        ch = X
        ck = Y
        picGraph.Circle (ch, ck), 0.25, vbBlack
    ElseIf p = 2 Then
        Dim cx As Double
        Dim cy As Double
        Dim cr As Double
        
        p = 0
        
        cx = X
        cy = Y
        picGraph.Circle (cx, cy), 0.25, vbBlack
        
        cr = (((ch - cx) ^ 2) + ((ck - cy) ^ 2)) ^ (1 / 2)
        picGraph.Circle (ch, ck), cr, vbBlack
        picGraph.Line (ch, ck)-(cx, cy), vbBlack
        
        g = g + 1
        
        If g = 1 Then
            'txt11 = Str(ch)
            'txt12 = Str(ck)
            'txt13 = Str(cr)
        ElseIf g = 2 Then
            'txt21 = Str(ch)
            'txt22 = Str(ck)
            'txt23 = Str(cr)
        End If
    End If
        
    
ElseIf dt = 3 And g <> 2 Then

    If p = 1 Then
        ph = X
        pk = Y
        picGraph.Circle (ph, pk), 0.25, vbBlack
    ElseIf p = 2 Then
        Dim px As Double
        Dim py As Double
        Dim a As Double
        Dim xi As Double
        Dim yi As Double
        Dim yi2 As Double
        Dim nx As Double
        Dim ny As Double
        
        p = 0
        
        px = X
        py = Y
        
        a = (py - pk) / (px - ph) ^ 2
        xi = -10
        yi = (a * ((xi - ph) ^ 2)) + pk
        xi = -9.99
        yi2 = (a * ((xi - ph) ^ 2)) + pk
        picGraph.Line (-10, yi)-(-9.99, yi2), vbBlack
        For nx = -9.98 To 10 Step 0.01
            ny = (a * ((nx - ph) ^ 2)) + pk
            picGraph.Line -(nx, ny)
        Next nx
        
        g = g + 1
        
        If g = 1 Then
            'txt11 = Str(a)
            'txt12 = Str(ph)
            'txt13 = Str(pk)
        ElseIf g = 2 Then
            'txt21 = Str(a)
            'txt22 = Str(ph)
            'txt23 = Str(pk)
        End If
    End If
    
ElseIf dt = 4 And g <> 2 Then
    
    If p = 1 Then
        bx1 = X
        by1 = Y
        picGraph.Circle (bx1, by1), 0.25, vbBlack
    ElseIf p = 2 Then
        Dim bx2 As Double
        Dim by2 As Double
        Dim bm As Double
        Dim bb As Double
        
        p = 0
        
        bx2 = X
        by2 = Y
        
        bm = (by1 - by2) / (bx1 - bx2)
        bb = by1 - (bm * bx1)
        
        picGraph.Circle (bx2, by2), 0.25, vbBlack
        picGraph.Line (bx1, by1)-(bx2, by2), vbBlack, B
        
        g = g + 1
        
        If g = 1 Then
            'txt11 = Str(bm)
            'txt12 = Str(bb)
        ElseIf g = 2 Then
            'txt21 = Str(bm)
            'txt22 = Str(bb)
        End If
    End If

ElseIf dt = 5 And g <> 2 Then

    If p = 1 Then
        ah = X
        ak = Y
        picGraph.Circle (ah, ak), 0.25, vbBlack
    ElseIf p = 2 Then
        Dim ax As Double
        Dim ay As Double
        Dim m1 As Double
        Dim m2 As Double
        Dim b1 As Double
        Dim b2 As Double
        Dim nay1 As Double
        Dim nay2 As Double
        
        p = 0
        
        ax = X
        ay = Y
        picGraph.Circle (ax, ay), 0.25, vbBlack
        
        m1 = (ak - ay) / (ah - ax)
        m2 = m1 * -1
        b1 = ak - (m1 * ah)
        b2 = ak - (m2 * ah)
        
        nay1 = -10 * m2 + b2
        nay2 = 10 * m1 + b1
        
        picGraph.Line (-10, nay1)-(ah, ak), vbBlack
        picGraph.Line (ah, ak)-(10, nay2), vbBlack
        
        g = g + 1
        
        If g = 1 Then
            'txt11 = Str(m1)
            'txt12 = Str(ah)
            'txt13 = Str(ak)
        ElseIf g = 2 Then
            'txt21 = Str(m1)
            'txt22 = Str(ah)
            'txt23 = Str(ak)
        End If
    End If
End If

End Sub

Private Sub Form_Activate()

dt = 1
p = 0
g = 0
DrawScale

End Sub

Private Sub mnuAbs_Click()

If p = 1 Then
    MsgBox "You can't change draw type while you are building a graph!"
Else
    dt = 5
End If

End Sub

Private Sub mnuBox_Click()

If p = 1 Then
    MsgBox "You can't change draw type while you are building a graph!"
Else
    dt = 4
End If

End Sub

Private Sub mnuCircle_Click()

If p = 1 Then
    MsgBox "You can't change draw type while you are building a graph!"
Else
    dt = 2
End If

End Sub

Private Sub mnuLine_Click()

If p = 1 Then
    MsgBox "You can't change draw type while you are building a graph!"
Else
    dt = 1
End If

End Sub

Private Sub mnuParabola_Click()

If p = 1 Then
    MsgBox "You can't change draw type while you are building a graph!"
Else
    dt = 3
End If

End Sub

Private Sub picGraph_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

p = p + 1

txt11 = X
txt12 = Y

If dt = 1 And g <> 2 Then
    
    If p = 1 Then
        lx1 = X
        ly1 = Y
        picGraph.Circle (lx1, ly1), 0.25
    ElseIf p = 2 Then
        Dim lm As Double
        Dim lb As Double
        Dim lx2 As Double
        Dim ly2 As Double
        Dim nly1 As Double
        Dim nly2 As Double
        
        p = 0
        
        lx2 = X
        ly2 = Y
        picGraph.Circle (lx2, ly2), 0.25
        
        lm = (ly1 - ly2) / (lx1 - lx2)
        lb = ly1 - (lm * lx1)
        
        nly1 = -10 * lm + lb
        nly2 = 10 * lm + lb
        
        picGraph.Line (-10, nly1)-(10, nly2), vbBlack
        
        g = g + 1
        
        If g = 1 Then
            'txt11 = Str(lm)
            'txt12 = Str(lb)
        ElseIf g = 2 Then
            'txt21 = Str(lm)
            'txt22 = Str(lb)
        End If
    End If
    
ElseIf dt = 2 And g <> 2 Then
    
    If p = 1 Then
        ch = X
        ck = Y
        picGraph.Circle (ch, ck), 0.25, vbBlack
    ElseIf p = 2 Then
        Dim cx As Double
        Dim cy As Double
        Dim cr As Double
        
        p = 0
        
        cx = X
        cy = Y
        picGraph.Circle (cx, cy), 0.25, vbBlack
        
        cr = (((ch - cx) ^ 2) + ((ck - cy) ^ 2)) ^ (1 / 2)
        picGraph.Circle (ch, ck), cr, vbBlack
        picGraph.Line (ch, ck)-(cx, cy), vbBlack
        
        g = g + 1
        
        If g = 1 Then
            'txt11 = Str(ch)
            'txt12 = Str(ck)
            'txt13 = Str(cr)
        ElseIf g = 2 Then
            'txt21 = Str(ch)
            'txt22 = Str(ck)
            'txt23 = Str(cr)
        End If
    End If
        
    
ElseIf dt = 3 And g <> 2 Then

    If p = 1 Then
        ph = X
        pk = Y
        picGraph.Circle (ph, pk), 0.25, vbBlack
    ElseIf p = 2 Then
        Dim px As Double
        Dim py As Double
        Dim a As Double
        Dim xi As Double
        Dim yi As Double
        Dim yi2 As Double
        Dim nx As Double
        Dim ny As Double
        
        p = 0
        
        px = X
        py = Y
        
        a = (py - pk) / (px - ph) ^ 2
        xi = -10
        yi = (a * ((xi - ph) ^ 2)) + pk
        xi = -9.99
        yi2 = (a * ((xi - ph) ^ 2)) + pk
        picGraph.Line (-10, yi)-(-9.99, yi2), vbBlack
        For nx = -9.98 To 10 Step 0.01
            ny = (a * ((nx - ph) ^ 2)) + pk
            picGraph.Line -(nx, ny)
        Next nx
        
        g = g + 1
        
        If g = 1 Then
            'txt11 = Str(a)
            'txt12 = Str(ph)
            'txt13 = Str(pk)
        ElseIf g = 2 Then
            'txt21 = Str(a)
            'txt22 = Str(ph)
            'txt23 = Str(pk)
        End If
    End If
    
ElseIf dt = 4 And g <> 2 Then
    
    If p = 1 Then
        bx1 = X
        by1 = Y
        picGraph.Circle (bx1, by1), 0.25, vbBlack
    ElseIf p = 2 Then
        Dim bx2 As Double
        Dim by2 As Double
        Dim bm As Double
        Dim bb As Double
        
        p = 0
        
        bx2 = X
        by2 = Y
        
        bm = (by1 - by2) / (bx1 - bx2)
        bb = by1 - (bm * bx1)
        
        picGraph.Circle (bx2, by2), 0.25, vbBlack
        picGraph.Line (bx1, by1)-(bx2, by2), vbBlack, B
        
        g = g + 1
        
        If g = 1 Then
            'txt11 = Str(bm)
            'txt12 = Str(bb)
        ElseIf g = 2 Then
            'txt21 = Str(bm)
            'txt22 = Str(bb)
        End If
    End If

ElseIf dt = 5 And g <> 2 Then

    If p = 1 Then
        ah = X
        ak = Y
        picGraph.Circle (ah, ak), 0.25, vbBlack
    ElseIf p = 2 Then
        Dim ax As Double
        Dim ay As Double
        Dim m1 As Double
        Dim m2 As Double
        Dim b1 As Double
        Dim b2 As Double
        Dim nay1 As Double
        Dim nay2 As Double
        
        p = 0
        
        ax = X
        ay = Y
        picGraph.Circle (ax, ay), 0.25, vbBlack
        
        m1 = (ak - ay) / (ah - ax)
        m2 = m1 * -1
        b1 = ak - (m1 * ah)
        b2 = ak - (m2 * ah)
        
        nay1 = -10 * m2 + b2
        nay2 = 10 * m1 + b1
        
        picGraph.Line (-10, nay1)-(ah, ak), vbBlack
        picGraph.Line (ah, ak)-(10, nay2), vbBlack
        
        g = g + 1
        
        If g = 1 Then
            'txt11 = Str(m1)
            'txt12 = Str(ah)
            'txt13 = Str(ak)
        ElseIf g = 2 Then
            'txt21 = Str(m1)
            'txt22 = Str(ah)
            'txt23 = Str(ak)
        End If
    End If
End If

End Sub
