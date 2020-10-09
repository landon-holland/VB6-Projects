VERSION 5.00
Begin VB.Form frmCafeteria 
   Caption         =   "Cafeteria"
   ClientHeight    =   3555
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3885
   LinkTopic       =   "Form1"
   ScaleHeight     =   3555
   ScaleWidth      =   3885
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPie 
      Caption         =   "Pie Chart"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   60
      TabIndex        =   9
      Top             =   3000
      Width           =   3795
   End
   Begin VB.CommandButton cmdVertical 
      Caption         =   "Vertical Bar Chat"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   60
      TabIndex        =   8
      Top             =   2400
      Width           =   3795
   End
   Begin VB.CommandButton cmdHorizontal 
      Caption         =   "Horizontal Bar Chat"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   60
      TabIndex        =   7
      Top             =   1800
      Width           =   3795
   End
   Begin VB.HScrollBar hsbChoices 
      Height          =   315
      LargeChange     =   2
      Left            =   60
      Max             =   10
      Min             =   1
      TabIndex        =   5
      Top             =   1380
      Value           =   10
      Width           =   3075
   End
   Begin VB.TextBox txtChoices 
      Alignment       =   1  'Right Justify
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
      Height          =   315
      Left            =   3180
      TabIndex        =   4
      Text            =   "10"
      Top             =   1380
      Width           =   675
   End
   Begin VB.TextBox txtPeople 
      Alignment       =   1  'Right Justify
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
      Height          =   315
      Left            =   3180
      TabIndex        =   3
      Text            =   "1000"
      Top             =   660
      Width           =   675
   End
   Begin VB.HScrollBar hsbPeople 
      Height          =   315
      LargeChange     =   100
      Left            =   60
      Max             =   1000
      Min             =   1
      TabIndex        =   2
      Top             =   660
      Value           =   1000
      Width           =   3075
   End
   Begin VB.Label lblChoices 
      Caption         =   "Choices:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   60
      TabIndex        =   6
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label lblPeople 
      Caption         =   "People:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   60
      TabIndex        =   1
      Top             =   360
      Width           =   855
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "Cafeteria"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3855
   End
End
Attribute VB_Name = "frmCafeteria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdHorizontal_Click()

Hide
frmHorizontal.Show

End Sub

Private Sub cmdPie_Click()

Hide
frmPie.Show

End Sub

Private Sub cmdVertical_Click()

Hide
frmVertical.Show

End Sub

Private Sub Form_Load()

people = 1000
choices = 10

End Sub

Private Sub hsbChoices_Change()

txtChoices = hsbChoices
choices = hsbChoices

End Sub

Private Sub hsbPeople_Change()

txtPeople = hsbPeople
people = hsbPeople

End Sub

Private Sub txtChoices_Change()

hsbChoices = txtChoices
choices = txtChoices

End Sub

Private Sub txtPeople_Change()

hsbPeople = txtPeople
people = txtPeople

End Sub
