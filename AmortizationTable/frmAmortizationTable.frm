VERSION 5.00
Begin VB.Form frmAmortizationTable 
   Caption         =   "Amortization Table"
   ClientHeight    =   4560
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8085
   LinkTopic       =   "Form1"
   ScaleHeight     =   4560
   ScaleWidth      =   8085
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optOneTime 
      Caption         =   "One Time"
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
      Left            =   5520
      TabIndex        =   10
      Top             =   2100
      Width           =   1155
   End
   Begin VB.OptionButton optYearly 
      Caption         =   "Yearly"
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
      Left            =   5520
      TabIndex        =   9
      Top             =   1800
      Width           =   975
   End
   Begin VB.OptionButton optMonthly 
      Caption         =   "Monthly"
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
      Left            =   5520
      TabIndex        =   8
      Top             =   1500
      Width           =   1035
   End
   Begin VB.Frame freTable 
      Caption         =   "Table"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   180
      TabIndex        =   16
      Top             =   3060
      Width           =   7755
      Begin VB.HScrollBar hsbTable 
         Height          =   255
         LargeChange     =   12
         Left            =   120
         Max             =   360
         Min             =   1
         TabIndex        =   18
         Top             =   900
         Value           =   1
         Width           =   7515
      End
      Begin VB.TextBox txtTable 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   17
         Top             =   540
         Width           =   7515
      End
      Begin VB.Label lblMonthlyPrincipalLabel 
         Caption         =   "Monthly Prin."
         Height          =   255
         Left            =   6180
         TabIndex        =   24
         Top             =   300
         Width           =   1455
      End
      Begin VB.Label lblMonthlyInterestLabel 
         Caption         =   "Monthly Int."
         Height          =   315
         Left            =   4980
         TabIndex        =   23
         Top             =   300
         Width           =   1215
      End
      Begin VB.Label lblTotalInterestLabel 
         Caption         =   "Total Interest"
         Height          =   195
         Left            =   3720
         TabIndex        =   22
         Top             =   300
         Width           =   1215
      End
      Begin VB.Label lblCurrentBalanceLabel 
         Caption         =   "Current Balance"
         Height          =   255
         Left            =   2280
         TabIndex        =   21
         Top             =   300
         Width           =   1215
      End
      Begin VB.Label lblYearLabel 
         Caption         =   "Year"
         Height          =   195
         Left            =   1620
         TabIndex        =   20
         Top             =   300
         Width           =   915
      End
      Begin VB.Label lblPaymentNumberLabel 
         Caption         =   "Payment Number"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   300
         Width           =   1515
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6240
      TabIndex        =   15
      Top             =   2460
      Width           =   1035
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "C&lear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5160
      TabIndex        =   14
      Top             =   2460
      Width           =   1035
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "&Calculate"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4080
      TabIndex        =   13
      Top             =   2460
      Width           =   1035
   End
   Begin VB.Frame freMonthlyPayment 
      Caption         =   "Monthly Payment"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Left            =   180
      TabIndex        =   11
      Top             =   1620
      Width           =   2175
      Begin VB.Label lblMonthlyPayment 
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
         Left            =   360
         TabIndex        =   12
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.Frame freInput 
      Caption         =   "Enter Values"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   7095
      Begin VB.TextBox txtExtraPayment 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5340
         TabIndex        =   7
         Top             =   540
         Width           =   1575
      End
      Begin VB.TextBox txtYears 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3900
         TabIndex        =   6
         Top             =   540
         Width           =   1335
      End
      Begin VB.TextBox txtYearlyInterest 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2460
         TabIndex        =   4
         Top             =   540
         Width           =   1335
      End
      Begin VB.TextBox txtLoanAmount 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   180
         TabIndex        =   2
         Top             =   540
         Width           =   2175
      End
      Begin VB.Label lblExtraPaymentPrompt 
         Alignment       =   2  'Center
         Caption         =   "Extra Payment"
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
         Left            =   5340
         TabIndex        =   25
         Top             =   300
         Width           =   1575
      End
      Begin VB.Label lblYearsPrompt 
         Alignment       =   2  'Center
         Caption         =   "Years"
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
         Left            =   3900
         TabIndex        =   5
         Top             =   300
         Width           =   1335
      End
      Begin VB.Label lblYearlyInterestPrompt 
         Alignment       =   2  'Center
         Caption         =   "Yearly Interest"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2460
         TabIndex        =   3
         Top             =   300
         Width           =   1335
      End
      Begin VB.Label lblLoanAmountPrompt 
         Alignment       =   2  'Center
         Caption         =   "Loan Amount"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   180
         TabIndex        =   1
         Top             =   300
         Width           =   2175
      End
   End
End
Attribute VB_Name = "frmAmortizationTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim amorttable(360) As String

Private Sub cmdCalculate_Click()

Dim loanamount As Currency
Dim monthlypayment As Currency
Dim yearlyrate As Single
Dim monthlyrate As Single
Dim years As Integer
Dim extrapayment As Currency
Dim payments As Integer

Dim paymentnumber As Integer
Dim monthlyint As Currency
Dim monthlyprincipal As Currency
Dim totalint As Currency
Dim currentamt As Currency
Dim yearnumber As Integer
Dim displine As String
Dim scrollmax As Integer

hsbTable.Value = 1

totalint = 0
displine = ""
scrollmax = 0

loanamount = Val(txtLoanAmount)
yearlyrate = Val(txtYearlyInterest)
years = Val(txtYears)
extrapayment = Val(txtExtraPayment)

monthlyrate = yearlyrate / 1200
payments = years * 12

If optMonthly = True Then

    monthlypayment = (loanamount * monthlyrate / (1 - (1 + monthlyrate) ^ (-payments))) + extrapayment
    lblMonthlyPayment = Format(monthlypayment, "Currency")
    
ElseIf optYearly = True Then

    extrapayment = extrapayment / 12
    monthlypayment = (loanamount * monthlyrate / (1 - (1 + monthlyrate) ^ (-payments))) + extrapayment
    lblMonthlyPayment = Format(monthlypayment, "Currency")
    
ElseIf optOneTime = True Then

    loanamount = loanamount - extrapayment
    monthlypayment = (loanamount * monthlyrate / (1 - (1 + monthlyrate) ^ (-payments)))
    lblMonthlyPayment = Format(monthlypayment, "Currency")

Else

    monthlypayment = (loanamount * monthlyrate / (1 - (1 + monthlyrate) ^ (-payments)))
    lblMonthlyPayment = Format(monthlypayment, "Currency")
    
End If

currentamt = loanamount

For paymentnumber = 1 To payments

    scrollmax = scrollmax + 1
    
    If currentamt < monthlypayment Then
    
        monthlypayment = currentamt
    
        monthlyint = currentamt * monthlyrate
        monthlyprincipal = monthlypayment - monthlyint
        totalint = totalint + monthlyint
        currentamt = currentamt + monthlyint - monthlypayment
        yearnumber = Int(paymentnumber / 12) + 1
        
        If paymentnumber Mod 12 = 0 Then
        
            yearnumber = yearnumber - 1
            
        End If
            
        displine = "             " + Format(paymentnumber, "####")
        displine = displine + "                  " + Format(yearnumber, "#0")
        displine = displine + "            " + "$0.00"
        displine = displine + "            " + Format(totalint, "Currency")
        displine = displine + "           " + Format(monthlyint, "Currency")
        displine = displine + "            " + Format(monthlyprincipal, "Currency")
    
        amorttable(paymentnumber) = displine
        
        Exit For
        
    Else
    
        monthlyint = currentamt * monthlyrate
        monthlyprincipal = monthlypayment - monthlyint
        totalint = totalint + monthlyint
        currentamt = currentamt + monthlyint - monthlypayment
        yearnumber = Int(paymentnumber / 12) + 1
    
        If paymentnumber Mod 12 = 0 Then
        
            yearnumber = yearnumber - 1
            
        End If
        
        displine = "             " + Format(paymentnumber, "####")
        displine = displine + "                  " + Format(yearnumber, "#0")
        displine = displine + "            " + Format(currentamt, "Currency")
        displine = displine + "            " + Format(totalint, "Currency")
        displine = displine + "           " + Format(monthlyint, "Currency")
        displine = displine + "            " + Format(monthlyprincipal, "Currency")
    
        amorttable(paymentnumber) = displine
        
    End If
    
Next paymentnumber

hsbTable.Max = scrollmax

txtTable = amorttable(1)

End Sub

Private Sub cmdClear_Click()

txtLoanAmount = ""
txtYearlyInterest = ""
txtYears = ""
txtTable = ""
txtExtraPayment = ""
optMonthly = False
optYearly = False
optOneTime = False

lblMonthlyPayment = ""

hsbTable.Value = 1

End Sub

Private Sub cmdExit_Click()

End

End Sub

Private Sub hsbTable_Change()

txtTable = amorttable(hsbTable.Value)

End Sub

