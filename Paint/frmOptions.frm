VERSION 5.00
Begin VB.Form frmOptions 
   Caption         =   "Options"
   ClientHeight    =   2580
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4785
   LinkTopic       =   "Form1"
   ScaleHeight     =   2580
   ScaleWidth      =   4785
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBox 
      Caption         =   "Box"
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
      Left            =   2460
      TabIndex        =   17
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton cmdLine 
      Caption         =   "Line"
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
      Left            =   1260
      TabIndex        =   16
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton cmdPencil 
      Caption         =   "Pencil"
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
      Left            =   180
      TabIndex        =   15
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton cmdEraser 
      Caption         =   "Eraser"
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
      Left            =   3540
      TabIndex        =   14
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox txtS 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2880
      TabIndex        =   13
      Text            =   "5"
      Top             =   1740
      Width           =   495
   End
   Begin VB.HScrollBar hsbS 
      Height          =   255
      LargeChange     =   10
      Left            =   180
      Max             =   100
      Min             =   1
      TabIndex        =   11
      Top             =   1800
      Value           =   5
      Width           =   2355
   End
   Begin VB.TextBox txtB 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2880
      TabIndex        =   10
      Text            =   "0"
      Top             =   1320
      Width           =   495
   End
   Begin VB.TextBox txtG 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2880
      TabIndex        =   9
      Text            =   "0"
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox txtR 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2880
      TabIndex        =   8
      Text            =   "0"
      Top             =   540
      Width           =   495
   End
   Begin VB.HScrollBar hsbB 
      Height          =   255
      LargeChange     =   50
      Left            =   180
      Max             =   255
      TabIndex        =   3
      Top             =   1380
      Width           =   2355
   End
   Begin VB.HScrollBar hsbG 
      Height          =   255
      LargeChange     =   50
      Left            =   180
      Max             =   255
      TabIndex        =   2
      Top             =   960
      Width           =   2355
   End
   Begin VB.HScrollBar hsbR 
      Height          =   255
      LargeChange     =   50
      Left            =   180
      Max             =   255
      TabIndex        =   1
      Top             =   540
      Width           =   2355
   End
   Begin VB.Label lblSPrompt 
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   256
      Left            =   2580
      TabIndex        =   12
      Top             =   1800
      Width           =   256
   End
   Begin VB.Label lblColorPicker 
      BackColor       =   &H00000000&
      Height          =   1515
      Left            =   3540
      TabIndex        =   7
      Top             =   540
      Width           =   1095
   End
   Begin VB.Label lblBPrompt 
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2580
      TabIndex        =   6
      Top             =   1380
      Width           =   255
   End
   Begin VB.Label lblGPrompt 
      Caption         =   "G"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2580
      TabIndex        =   5
      Top             =   960
      Width           =   255
   End
   Begin VB.Label lblRPrompt 
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2580
      TabIndex        =   4
      Top             =   540
      Width           =   255
   End
   Begin VB.Label lblColorPrompt 
      Caption         =   "Color:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   1455
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bdr As Integer
Dim bdg As Integer
Dim bdb As Integer

Private Sub cmdBox_Click()

drawtype = 3

linepoint = 1

End Sub

Private Sub cmdEraser_Click()

dr = 255
dg = 255
db = 255

bdr = txtR
bdg = txtG
bdb = txtB

drawtype = 1

End Sub

Private Sub cmdLine_Click()

linepoint = 1

drawtype = 2

End Sub

Private Sub cmdPencil_Click()

dr = bdr
dg = bdg
db = bdb

drawtype = 1

End Sub

Private Sub Form_Load()

bdr = 0
bdg = 0
bdb = 0

End Sub

Private Sub hsbB_Change()

db = hsbB.Value
txtB = db
lblColorPicker.BackColor = RGB(dr, dg, db)

End Sub

Private Sub hsbG_Change()

dg = hsbG.Value
txtG = dg
lblColorPicker.BackColor = RGB(dr, dg, db)

End Sub

Private Sub hsbR_Change()

dr = hsbR.Value
txtR = dr
lblColorPicker.BackColor = RGB(dr, dg, db)

End Sub

Private Sub hsbS_Change()

ds = hsbS.Value
txtS = ds
frmPaint.DrawWidth = ds

End Sub

Private Sub txtB_Change()

db = txtB
hsbB.Value = db
lblColorPicker.BackColor = RGB(dr, dg, db)

End Sub

Private Sub txtG_Change()

dg = txtG
hsbG.Value = dg
lblColorPicker.BackColor = RGB(dr, dg, db)

End Sub

Private Sub txtR_Change()

dr = txtR
hsbR.Value = dr
lblColorPicker.BackColor = RGB(dr, dg, db)

End Sub

Private Sub txtS_Change()

ds = txtS
hsbS.Value = ds
frmPaint.DrawWidth = ds

End Sub
