VERSION 5.00
Begin VB.Form frmCarLoan 
   Caption         =   "Car Loan"
   ClientHeight    =   3375
   ClientLeft      =   0
   ClientTop       =   240
   ClientWidth     =   5340
   LinkTopic       =   "Form1"
   ScaleHeight     =   3375
   ScaleWidth      =   5340
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClear 
      Appearance      =   0  'Flat
      Caption         =   "&Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   16
      Top             =   2940
      Width           =   1515
   End
   Begin VB.CommandButton cmdExit 
      Appearance      =   0  'Flat
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
      Height          =   375
      Left            =   3660
      TabIndex        =   15
      Top             =   2940
      Width           =   1515
   End
   Begin VB.CommandButton cmdFind 
      Appearance      =   0  'Flat
      Caption         =   "Find &Payment"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   180
      TabIndex        =   14
      Top             =   2940
      Width           =   1515
   End
   Begin VB.Frame freOutput 
      Caption         =   "Calculated Values"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2715
      Left            =   2750
      TabIndex        =   7
      Top             =   120
      Width           =   2415
      Begin VB.Label lblTotalPayback 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   180
         TabIndex        =   13
         Top             =   2160
         Width           =   2055
      End
      Begin VB.Label lblTotalPaybackPrompt 
         Alignment       =   2  'Center
         Caption         =   "Total Payback"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   60
         TabIndex        =   12
         Top             =   1860
         Width           =   2295
      End
      Begin VB.Label lblTotalInterestPaid 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   180
         TabIndex        =   11
         Top             =   1380
         Width           =   2055
      End
      Begin VB.Label lblTotalInterestPaidPrompt 
         Alignment       =   2  'Center
         Caption         =   "Total Interest Paid"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   60
         TabIndex        =   10
         Top             =   1080
         Width           =   2295
      End
      Begin VB.Label lblMonthlyPayment 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   312
         Left            =   192
         TabIndex        =   9
         Top             =   576
         Width           =   2056
      End
      Begin VB.Label lblMonthlyPaymentPrompt 
         Alignment       =   2  'Center
         Caption         =   "Monthly Payment"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   200
         Left            =   64
         TabIndex        =   8
         Top             =   320
         Width           =   2295
      End
   End
   Begin VB.TextBox txtYears 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   360
      TabIndex        =   6
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Frame freValues 
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
      Height          =   2715
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   2415
      Begin VB.TextBox txtInterestRate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   180
         TabIndex        =   4
         Top             =   1380
         Width           =   2055
      End
      Begin VB.TextBox txtLoanAmount 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   180
         TabIndex        =   2
         Top             =   600
         Width           =   2055
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
         Height          =   315
         Left            =   60
         TabIndex        =   5
         Top             =   1860
         Width           =   2295
      End
      Begin VB.Label lblInterestRatePrompt 
         Alignment       =   2  'Center
         Caption         =   "Yearly Interest Rate"
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
         Left            =   60
         TabIndex        =   3
         Top             =   1080
         Width           =   2295
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
         Height          =   255
         Left            =   60
         TabIndex        =   1
         Top             =   300
         Width           =   2295
      End
   End
End
Attribute VB_Name = "frmCarLoan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClear_Click()

txtLoanAmount = ""
txtInterestRate = ""
txtYears = ""
lblMonthlyPayment = ""
lblTotalInterestPaid = ""
lblTotalPayback = ""

txtLoanAmount.SetFocus

End Sub

Private Sub cmdExit_Click()

End

End Sub

Private Sub cmdFind_Click()

Dim loanamount As Currency
Dim monthlypayment As Currency
Dim totalinterest As Currency
Dim totalpayback As Currency
Dim yearlyrate As Single
Dim monthlyrate As Single
Dim years As Integer
Dim payments As Integer

loanamount = Val(txtLoanAmount)
yearlyrate = Val(txtInterestRate)
years = Val(txtYears)

monthlyrate = yearlyrate / 1200
payments = years * 12

monthlypayment = loanamount * monthlyrate / (1 - (1 + monthlyrate) ^ (-payments))

totalpayback = monthlypayment * payments

totalinterest = totalpayback - loanamount

lblMonthlyPayment = Format(monthlypayment, "Currency")
lblTotalInterestPaid = Format(totalinterest, "Currency")
lblTotalPayback = Format(totalpayback, "Currency")

End Sub

Private Sub cmdSwitch_Click()

frmCarLoan.Hide
frmReverseCarLoan.Show

End Sub
