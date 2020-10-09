VERSION 5.00
Begin VB.Form frmReverseCarLoan 
   Caption         =   "Car Loan"
   ClientHeight    =   3855
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5340
   LinkTopic       =   "Form1"
   ScaleHeight     =   3855
   ScaleWidth      =   5340
   StartUpPosition =   3  'Windows Default
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
      Height          =   2775
      Left            =   2760
      TabIndex        =   7
      Top             =   180
      Width           =   2415
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   180
         TabIndex        =   10
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   180
         TabIndex        =   9
         Top             =   1380
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   180
         TabIndex        =   8
         Top             =   2220
         Width           =   2055
      End
      Begin VB.Label htrh 
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
         Height          =   195
         Left            =   60
         TabIndex        =   13
         Top             =   300
         Width           =   2295
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Total Budget"
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
         Top             =   1080
         Width           =   2295
      End
      Begin VB.Label Label1 
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
         Height          =   255
         Left            =   60
         TabIndex        =   11
         Top             =   1920
         Width           =   2295
      End
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
      Height          =   2775
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   2415
      Begin VB.TextBox txtYearlyInterestRate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   180
         TabIndex        =   6
         Top             =   2220
         Width           =   2055
      End
      Begin VB.TextBox txtTotalBudget 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   180
         TabIndex        =   4
         Top             =   1380
         Width           =   2055
      End
      Begin VB.TextBox txtMonthlyPayment 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   180
         TabIndex        =   2
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label lblYears 
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
         Left            =   60
         TabIndex        =   5
         Top             =   1920
         Width           =   2295
      End
      Begin VB.Label lbld 
         Alignment       =   2  'Center
         Caption         =   "Total Budget"
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
         TabIndex        =   3
         Top             =   1080
         Width           =   2295
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
         Height          =   195
         Left            =   60
         TabIndex        =   1
         Top             =   300
         Width           =   2295
      End
   End
End
Attribute VB_Name = "frmReverseCarLoan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

