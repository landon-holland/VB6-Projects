VERSION 5.00
Begin VB.Form frmPie 
   Caption         =   "Pie Chart"
   ClientHeight    =   5055
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8295
   LinkTopic       =   "Form1"
   ScaleHeight     =   5055
   ScaleWidth      =   8295
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmr 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5820
      Top             =   4260
   End
   Begin VB.CommandButton cmdGraph 
      Caption         =   "Graph"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1020
      TabIndex        =   0
      Top             =   4320
      Width           =   2355
   End
   Begin VB.Label lbl1 
      Caption         =   "If you can't dazzle them with your brilliance, baffle them with your BS."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   4680
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   2475
   End
   Begin VB.Image img 
      Height          =   4170
      Left            =   0
      Picture         =   "frmPie.frx":0000
      Top             =   0
      Visible         =   0   'False
      Width           =   4110
   End
End
Attribute VB_Name = "frmPie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim c As Integer

Private Sub cmdGraph_Click()

img.Visible = True
lbl1.Visible = True
tmr.Enabled = True

End Sub

Private Sub Form_Load()
c = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)

Hide
frmCafeteria.Show

End Sub

Private Sub tmr_Timer()

c = c + 1
If c = 1 Then
    img.Visible = False
ElseIf c = 2 Then
    img.Visible = True
    c = 0
End If

End Sub
