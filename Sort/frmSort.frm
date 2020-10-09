VERSION 5.00
Begin VB.Form frmSort 
   Caption         =   "Sort"
   ClientHeight    =   6720
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9105
   LinkTopic       =   "Form1"
   ScaleHeight     =   6720
   ScaleWidth      =   9105
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstComb 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5295
      Left            =   7020
      TabIndex        =   8
      Top             =   780
      Width           =   2000
   End
   Begin VB.ListBox lstBubble 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5295
      Left            =   4860
      TabIndex        =   6
      Top             =   780
      Width           =   2000
   End
   Begin VB.ListBox lstExchange 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5295
      Left            =   2700
      TabIndex        =   4
      Top             =   780
      Width           =   2000
   End
   Begin VB.ListBox lstOriginal 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5295
      Left            =   240
      TabIndex        =   2
      Top             =   780
      Width           =   2000
   End
   Begin VB.Label lblTime 
      Caption         =   "Time (in ms)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   12
      Top             =   6300
      Width           =   1995
   End
   Begin VB.Label lblCombTime 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   7020
      TabIndex        =   11
      Top             =   6300
      Width           =   1995
   End
   Begin VB.Label lblBubbleTime 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   4860
      TabIndex        =   10
      Top             =   6300
      Width           =   1995
   End
   Begin VB.Label lblExchangeTime 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   2700
      TabIndex        =   9
      Top             =   6300
      Width           =   1995
   End
   Begin VB.Label lblComb 
      Alignment       =   2  'Center
      Caption         =   "Comb"
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
      Left            =   7020
      TabIndex        =   7
      Top             =   480
      Width           =   1995
   End
   Begin VB.Label lblBubble 
      Alignment       =   2  'Center
      Caption         =   "Bubble"
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
      Left            =   4860
      TabIndex        =   5
      Top             =   480
      Width           =   1995
   End
   Begin VB.Label lblExchange 
      Alignment       =   2  'Center
      Caption         =   "Exchange"
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
      Left            =   2700
      TabIndex        =   3
      Top             =   480
      Width           =   1995
   End
   Begin VB.Line lnDivider2 
      BorderWidth     =   2
      X1              =   60
      X2              =   9960
      Y1              =   420
      Y2              =   420
   End
   Begin VB.Line lnDivider 
      BorderWidth     =   2
      X1              =   2460
      X2              =   2460
      Y1              =   0
      Y2              =   6660
   End
   Begin VB.Label lblOriginal 
      Alignment       =   2  'Center
      Caption         =   "Original"
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
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   1995
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "Sort"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   16.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   -60
      Width           =   9075
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuClear 
         Caption         =   "Clear"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuGenerate 
      Caption         =   "Generate"
      Begin VB.Menu mnuNumbers 
         Caption         =   "Numbers"
      End
      Begin VB.Menu mnuWords 
         Caption         =   "Words"
      End
   End
   Begin VB.Menu mnuSort 
      Caption         =   "Sort"
      Begin VB.Menu mnuExchange 
         Caption         =   "Exchange"
      End
      Begin VB.Menu mnuBubble 
         Caption         =   "Bubble"
      End
      Begin VB.Menu mnuComb 
         Caption         =   "Comb"
      End
   End
   Begin VB.Menu mnuSearch 
      Caption         =   "Search"
   End
End
Attribute VB_Name = "frmSort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim arrnumbers(1 To 5000) As Integer
Dim arrdictionary(1 To 58112) As String
Dim arrwords(1 To 5000) As String
Dim arrnexchange(1 To 5000) As Integer
Dim arrwexchange(1 To 5000) As String
Dim arrnbubble(1 To 5000) As Integer
Dim arrwbubble(1 To 5000) As String
Dim arrncomb(1 To 5000) As Integer
Dim arrwcomb(1 To 5000) As String
Dim arrnautosort(1 To 5000) As Integer
Dim arrwautosort(1 To 5000) As String
Dim sorttype As Integer
Dim sortm As Integer
Dim amount As Integer
Dim arrposition(1 To 5000) As Integer
Dim searchresults As Integer

Private Declare Function GetTickCount Lib "kernel32" () As Long

Sub AutoNSort()

Dim i As Integer
Dim gap As Single
Dim temp As Integer
Dim swapped As Boolean
Const shrink = 1.3

If sorttype = 1 Then
    
    For i = 1 To amount
        arrnautosort(i) = arrnumbers(i)
    Next i
    
    gap = amount
    Do
        gap = Int(gap / shrink)
        If gap < 1 Then gap = 1
            swapped = False
        For i = 1 To amount - gap
            If arrnautosort(i) > arrnautosort(i + gap) Then
                temp = arrnautosort(i)
                arrnautosort(i) = arrnautosort(i + gap)
                arrnautosort(i + gap) = temp
                swapped = True
            End If
        Next i
    Loop Until Not swapped And gap = 1
End If

End Sub

Sub AutoWSort()

Dim i As Integer
Dim gap As Single
Dim temp As String
Dim swapped As Boolean
Const shrink = 1.3

If sorttype = 2 Then
    
    For i = 1 To amount
        arrwautosort(i) = arrwords(i)
    Next i
    
    gap = amount
    Do
        gap = Int(gap / shrink)
        If gap < 1 Then gap = 1
            swapped = False
        For i = 1 To amount - gap
            If arrwautosort(i) > arrwautosort(i + gap) Then
                temp = arrwautosort(i)
                arrwautosort(i) = arrwautosort(i + gap)
                arrwautosort(i + gap) = temp
                swapped = True
            End If
        Next i
    Loop Until Not swapped And gap = 1
End If

End Sub

Sub NBinSearch(arr, searchitem, maxitems As Integer)

Dim i As Integer
Dim j As Integer
Dim low As Integer
Dim high As Integer
Dim md As Integer

For i = 1 To amount
    arrposition(i) = 0
Next i
low = 1
high = maxitems

Do While low <= high
    md = (low + high) / 2
    
    If searchitem = arr(md) Then
        arrposition(1) = md
        Exit Do
    ElseIf searchitem < arr(md) Then
        high = md - 1
    ElseIf searchitem > arr(md) Then
        low = md + 1
    End If
Loop

If arrposition(1) <> 0 Then
    i = 1
    j = 2
    Do While searchitem = arr(md + i)
        If searchitem <> arr(md + i) Then
            Exit Do
        Else
            arrposition(j) = md + i
            j = j + 1
        End If
        i = i + 1
    Loop
    
    If md <> 1 Then
        i = 1
        Do While searchitem = arr(md - i)
            If searchitem <> arr(md - i) Or (md - i) - 1 = 0 Then
                Exit Do
            Else
                arrposition(j) = md - i
                j = j + 1
            End If
            i = i + 1
        Loop
    Else
        arrposition(1) = 1
    End If
    searchresults = j
End If

If arrposition(1) <> 0 Then
    
    Dim gap As Single
    Dim swapped As Boolean
    Dim tempi As Integer
    Const shrink = 1.3
    
    gap = searchresults
    Do
        gap = Int(gap / shrink)
        If gap < 1 Then gap = 1
            swapped = False
        For i = 1 To searchresults - gap
            If arrposition(i) > arrposition(i + gap) Then
                tempi = arrposition(i)
                arrposition(i) = arrposition(i + gap)
                arrposition(i + gap) = tempi
                swapped = True
            End If
        Next i
    Loop Until Not swapped And gap = 1
ElseIf md = 1 Then
    arrposition(1) = searchitem
Else
    arrposition(1) = -1
End If

End Sub

Sub WBinSearch(arr, searchitem, maxitems As Integer)

Dim i As Integer
Dim j As Integer
Dim low As Integer
Dim high As Integer
Dim md As Integer

For i = 1 To amount
    arrposition(i) = 0
Next i
low = 1
high = maxitems

Do While low <= high
    md = (low + high) / 2
    
    If searchitem = arr(md) Then
        arrposition(1) = md
        Exit Do
    ElseIf searchitem < arr(md) Then
        high = md - 1
    ElseIf searchitem > arr(md) Then
        low = md + 1
    End If
Loop

If arrposition(1) <> 0 Then
    i = 1
    j = 2
    Do While searchitem = arr(md + i)
        If searchitem <> arr(md + i) Then
            Exit Do
        Else
            arrposition(j) = md + i
            j = j + 1
        End If
        i = i + 1
    Loop
    
    If md <> 1 Then
        i = 1
        Do While searchitem = arr(md - i)
            If searchitem <> arr(md - i) Or (md - i) - 1 = 0 Then
                Exit Do
            Else
                arrposition(j) = md - i
                j = j + 1
            End If
            i = i + 1
        Loop
    Else
        arrposition(j) = 1
    End If
    searchresults = j
End If

If arrposition(1) <> 0 Then
    
    Dim gap As Single
    Dim swapped As Boolean
    Dim tempi As Integer
    Const shrink = 1.3
    
    gap = searchresults
    Do
        gap = Int(gap / shrink)
        If gap < 1 Then gap = 1
            swapped = False
        For i = 1 To searchresults - gap
            If arrposition(i) > arrposition(i + gap) Then
                tempi = arrposition(i)
                arrposition(i) = arrposition(i + gap)
                arrposition(i + gap) = tempi
                swapped = True
            End If
        Next i
    Loop Until Not swapped And gap = 1
Else
    arrposition(1) = -1
End If

End Sub

Private Sub Form_Load()

sorttype = 0
sortm = 0
amount = 0

End Sub

Private Sub mnuBubble_Click()

Dim i As Integer
Dim j As Integer
Dim tempi As Integer
Dim temps As String
Dim swapped As Boolean

Dim startt As Long
Dim endt As Long

If sorttype = 0 Then
    MsgBox "You need to generate an array of words or numbers first."
ElseIf sorttype = 1 Then
    
    sortm = 2
    
    For i = 1 To amount
        arrnbubble(i) = arrnumbers(i)
    Next i
    
    startt = GetTickCount()
    i = amount
    Do
        swapped = False
        For j = 1 To i - 1
            If arrnbubble(j) > arrnbubble(j + 1) Then
                tempi = arrnbubble(j)
                arrnbubble(j) = arrnbubble(j + 1)
                arrnbubble(j + 1) = tempi
                swapped = True
            End If
        Next j
        i = i - 1
    Loop Until Not swapped
    endt = GetTickCount()
    
    For i = 1 To amount
        lstBubble.AddItem arrnbubble(i)
    Next i
    lblBubbleTime = (endt - startt)
    
ElseIf sorttype = 2 Then
    
    sortm = 2
    
    For i = 1 To amount
        arrwbubble(i) = arrwords(i)
    Next i
    
    startt = GetTickCount()
    i = amount
    Do
        swapped = False
        For j = 1 To i - 1
            If arrwbubble(j) > arrwbubble(j + 1) Then
                temps = arrwbubble(j)
                arrwbubble(j) = arrwbubble(j + 1)
                arrwbubble(j + 1) = temps
                swapped = True
            End If
        Next j
        i = i - 1
    Loop Until Not swapped
    endt = GetTickCount()
    
    For i = 1 To amount
        lstBubble.AddItem arrwbubble(i)
    Next i
    lblBubbleTime = (endt - startt)
End If

End Sub

Private Sub mnuClear_Click()

Dim i As Integer

sorttype = 0
sortm = 0
amount = 0

lstOriginal.Clear
lstExchange.Clear
lstBubble.Clear
lstComb.Clear

lblExchangeTime = ""
lblBubbleTime = ""
lblCombTime = ""

For i = 1 To 5000
    arrnumbers(i) = 0
    arrnexchange(i) = 0
    arrnbubble(i) = 0
    arrncomb(i) = 0
    arrnautosort(i) = 0
    
    arrwords(i) = ""
    arrwexchange(i) = ""
    arrwbubble(i) = ""
    arrwcomb(i) = ""
    arrwautosort(i) = ""
    
    arrposition(i) = 0
Next i

End Sub

Private Sub mnuComb_Click()

Dim i As Integer
Dim gap As Single
Dim tempi As Integer
Dim temps As String
Dim swapped As Boolean
Const shrink = 1.3

Dim startt As Long
Dim endt As Long

If sorttype = 0 Then
    MsgBox "You need to generate an array of words or numbers first."
ElseIf sorttype = 1 Then
    
    sortm = 3
    
    For i = 1 To amount
        arrncomb(i) = arrnumbers(i)
    Next i
    
    startt = GetTickCount()
    gap = amount
    Do
        gap = Int(gap / shrink)
        If gap < 1 Then gap = 1
            swapped = False
        For i = 1 To amount - gap
            If arrncomb(i) > arrncomb(i + gap) Then
                tempi = arrncomb(i)
                arrncomb(i) = arrncomb(i + gap)
                arrncomb(i + gap) = tempi
                swapped = True
            End If
        Next i
    Loop Until Not swapped And gap = 1
    endt = GetTickCount()
    
    For i = 1 To amount
        lstComb.AddItem arrncomb(i)
    Next i
    lblCombTime = Str(endt - startt)
ElseIf sorttype = 2 Then
    
    sortm = 3
    
    For i = 1 To amount
        arrwcomb(i) = arrwords(i)
    Next i
    
    startt = GetTickCount()
    gap = amount
    Do
        gap = Int(gap / shrink)
        If gap < 1 Then gap = 1
            swapped = False
        For i = 1 To amount - gap
            If arrwcomb(i) > arrwcomb(i + gap) Then
                temps = arrwcomb(i)
                arrwcomb(i) = arrwcomb(i + gap)
                arrwcomb(i + gap) = temps
                swapped = True
            End If
        Next i
    Loop Until Not swapped And gap = 1
    endt = GetTickCount()
    
    For i = 1 To amount
        lstComb.AddItem arrwcomb(i)
    Next i
    lblCombTime = Str(endt - startt)
End If

End Sub

Private Sub mnuExchange_Click()

Dim i As Integer
Dim f As Integer
Dim b As Integer
Dim tempi As Integer
Dim temps As String

Dim startt As Long
Dim endt As Long

If sorttype = 0 Then
    MsgBox "You need to generate an array of words or numbers first."
ElseIf sorttype = 1 Then
    
    sortm = 1
    
    For i = 1 To amount
        arrnexchange(i) = arrnumbers(i)
    Next i

    startt = GetTickCount()
    For f = 1 To amount - 1
        For b = f + 1 To amount
            If arrnexchange(f) > arrnexchange(b) Then
                tempi = arrnexchange(f)
                arrnexchange(f) = arrnexchange(b)
                arrnexchange(b) = tempi
            End If
        Next b
    Next f
    endt = GetTickCount()
    
    For i = 1 To amount
        lstExchange.AddItem arrnexchange(i)
    Next i
    lblExchangeTime = Str(endt - startt)
    
ElseIf sorttype = 2 Then

    sortm = 1
    
    For i = 1 To amount
        arrwexchange(i) = arrwords(i)
    Next i

    startt = GetTickCount()
    For f = 1 To amount - 1
        For b = f + 1 To amount
            If arrwexchange(f) > arrwexchange(b) Then
                temps = arrwexchange(f)
                arrwexchange(f) = arrwexchange(b)
                arrwexchange(b) = temps
            End If
        Next b
    Next f
    endt = GetTickCount()
    
    For i = 1 To amount
        lstExchange.AddItem arrwexchange(i)
    Next i
    lblExchangeTime = Str(endt - startt)
End If

End Sub

Private Sub mnuNumbers_Click()

Dim minn As Integer
Dim maxn As Integer
Dim i As Integer

Randomize
lstOriginal.Clear
For i = 1 To 5000
    arrnumbers(i) = 0
Next i

minn = Int(InputBox("Enter the lowest value desired in your array.", "Generate Numbers", "1"))
maxn = Int(InputBox("Enter the highest value desired in your array.", "Generate Numbers", "1000"))
amount = Int(InputBox("Enter the amount of numbers desired in your array.", "Generate Numbers", "5000"))

For i = 1 To amount
    arrnumbers(i) = Int((maxn - minn + 1) * Rnd) + minn
    lstOriginal.AddItem arrnumbers(i)
Next i

sorttype = 1

End Sub

Private Sub mnuSearch_Click()

Dim i As Integer
Dim searchingi As Integer
Dim searchings As String

If sorttype = 0 Then
    MsgBox "You need to generate an array of words or numbers first."
ElseIf sorttype = 1 Then
    AutoNSort
    searchingi = InputBox("Enter a number to search for:", "Search")
    Call NBinSearch(arrnautosort(), searchingi, amount)
    
    If arrposition(1) = -1 Then
        MsgBox "The number does not exist in the array."
    Else
        For i = 1 To searchresults
            If arrposition(i) <> 0 Then
                MsgBox "The number" + Str(searchingi) + " was found at position" + Str(arrposition(i)) + "."
            End If
        Next i
    End If
ElseIf sorttype = 2 Then
    AutoWSort
    searchings = InputBox("Enter a word to search for:", "Search")
    Call WBinSearch(arrwautosort(), searchings, amount)
    
    If arrposition(1) = -1 Then
        MsgBox "The word does not exist in the array."
    Else
        For i = 1 To searchresults
            If arrposition(i) <> 0 Then
                MsgBox "The word " + searchings + " was found at position" + Str(arrposition(i)) + "."
            End If
        Next i
    End If
End If

End Sub

Private Sub mnuWords_Click()

Dim line As String
Dim i As Long

Randomize
lstOriginal.Clear
For i = 1 To 5000
    arrwords(i) = ""
Next i
amount = Int(InputBox("Enter the amount of words desired in your array.", "Generate Words", "5000"))

Open "D:\CP2\VB6 Project Files\Sort\resources\dictionary.txt" For Input As #1
    For i = 1 To 58112
        Line Input #1, line
        line = Trim(line)
        arrdictionary(i) = line
    Next i
Close #1

For i = 1 To amount
    arrwords(i) = arrdictionary(Int((Rnd * 58112) + 1))
    lstOriginal.AddItem arrwords(i)
Next i

sorttype = 2

End Sub
