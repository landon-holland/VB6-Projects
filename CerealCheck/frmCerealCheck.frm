VERSION 5.00
Begin VB.Form frmCerealCheck 
   Caption         =   "Cereal Checker"
   ClientHeight    =   3200
   ClientLeft      =   120
   ClientTop       =   448
   ClientWidth     =   3824
   LinkTopic       =   "Form1"
   ScaleHeight     =   3200
   ScaleWidth      =   3824
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.86
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2460
      TabIndex        =   6
      Top             =   1380
      Width           =   1095
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.86
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1320
      TabIndex        =   5
      Top             =   1380
      Width           =   1155
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "Check"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.86
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   4
      Top             =   1380
      Width           =   1095
   End
   Begin VB.TextBox txtBowls 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.86
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1500
      TabIndex        =   3
      Top             =   780
      Width           =   2055
   End
   Begin VB.TextBox txtBoxes 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.86
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1500
      TabIndex        =   1
      Top             =   180
      Width           =   2055
   End
   Begin VB.Label lblGood 
      BackColor       =   &H0000C000&
      Caption         =   "No need to buy more cereal! :)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   1920
      TabIndex        =   8
      Top             =   1980
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.Label lblMoreCereal 
      BackColor       =   &H000000FF&
      Caption         =   "GET MORE CEREAL!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   240
      TabIndex        =   7
      Top             =   1980
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.Label lblBowls 
      Caption         =   "Bowls Eaten Weekly:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.71
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label lblBoxesPrompt 
      Caption         =   "Boxes Left:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.86
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1155
   End
End
Attribute VB_Name = "frmCerealCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCheck_Click()
Dim Boxes As Integer
Dim Bowls As Integer
Dim Servings As Integer

If Not IsNumeric(txtBoxes) Or Not IsNumeric(txtBowls) Then
    MsgBox "Boxes and bowls are described in numbers, you moron."
    txtBoxes.SetFocus
Else
    Boxes = Val(txtBoxes)
    Bowls = Val(txtBowls)
    
    Servings = Boxes * 12
    
    If Servings >= Bowls * 2 Then
        lblMoreCereal.Visible = False
        lblGood.Visible = True
    Else
        lblGood.Visible = False
        lblMoreCereal.Visible = True
    End If
    cmdClear.SetFocus
End If
End Sub

Private Sub cmdClear_Click()
txtBoxes = ""
txtBowls = ""
lblMoreCereal.Visible = False
lblGood.Visible = False
txtBoxes.SetFocus
End Sub

Private Sub cmdExit_Click()
End
End Sub

Private Sub txtBowls_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdCheck.SetFocus
End If
End Sub

Private Sub txtBoxes_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtBowls.SetFocus
End If
End Sub
