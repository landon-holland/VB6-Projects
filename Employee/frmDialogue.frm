VERSION 5.00
Begin VB.Form frmDialogue 
   Caption         =   "Dialogue"
   ClientHeight    =   5192
   ClientLeft      =   64
   ClientTop       =   344
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   5192
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back to Main"
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
      Left            =   60
      TabIndex        =   20
      Top             =   4740
      Width           =   4575
   End
   Begin VB.CommandButton cmdRight 
      Caption         =   "-->"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.86
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4200
      TabIndex        =   19
      Top             =   4260
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.CommandButton cmdLeft 
      Caption         =   "<--"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.86
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   60
      TabIndex        =   18
      Top             =   4260
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.86
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3000
      TabIndex        =   17
      Top             =   4260
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "Change"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.86
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1740
      TabIndex        =   16
      Top             =   4260
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.86
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   540
      TabIndex        =   15
      Top             =   4260
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.TextBox txtPhoneNumber 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.71
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   60
      TabIndex        =   14
      Top             =   3840
      Width           =   4575
   End
   Begin VB.TextBox txtWage 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.71
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   60
      TabIndex        =   12
      Top             =   3300
      Width           =   4575
   End
   Begin VB.TextBox txtID 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.71
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   60
      TabIndex        =   9
      Top             =   2220
      Width           =   4575
   End
   Begin VB.TextBox txtPayType 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.71
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   60
      TabIndex        =   11
      Top             =   2760
      Width           =   4575
   End
   Begin VB.TextBox txtAge 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.71
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   60
      TabIndex        =   6
      Top             =   1680
      Width           =   4575
   End
   Begin VB.TextBox txtLastName 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.71
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   60
      TabIndex        =   4
      Top             =   1140
      Width           =   4575
   End
   Begin VB.TextBox txtFirstName 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.71
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   60
      TabIndex        =   2
      Top             =   600
      Width           =   4575
   End
   Begin VB.Label lblPhoneNumber 
      Caption         =   "Phone Number:"
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
      Left            =   60
      TabIndex        =   13
      Top             =   3600
      Width           =   1935
   End
   Begin VB.Label lblWage 
      Caption         =   "Wage:"
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
      Left            =   60
      TabIndex        =   10
      Top             =   3060
      Width           =   1335
   End
   Begin VB.Label lblID 
      Caption         =   "ID:"
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
      Left            =   60
      TabIndex        =   8
      Top             =   1980
      Width           =   1335
   End
   Begin VB.Label lblPayType 
      Caption         =   "Pay type:"
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
      Left            =   60
      TabIndex        =   7
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label lblAge 
      Caption         =   "Age:"
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
      Left            =   60
      TabIndex        =   5
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label lblLastName 
      Caption         =   "Last Name:"
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
      Left            =   60
      TabIndex        =   3
      Top             =   900
      Width           =   1335
   End
   Begin VB.Label lblFirstName 
      Caption         =   "First Name:"
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
      Left            =   60
      TabIndex        =   1
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "Dialogue"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.72
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
      Width           =   4635
   End
End
Attribute VB_Name = "frmDialogue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim roledexc As Integer

Sub RoledexChange()

With arremployees(roledexc)
txtFirstName = .firstname
txtLastName = .lastname
txtAge = Str(.age)
txtID = Str(.id)
txtPayType = .paytype
txtWage = Format(.wage, "Currency")
txtPhoneNumber = .phonenumber
End With

End Sub

Private Sub cmdAdd_Click()

entries = entries + 1

With arremployees(entries)
    .firstname = txtFirstName
    .lastname = txtLastName
    .age = txtAge
    .id = txtID
    .paytype = txtPayType
    .wage = txtWage
    .phonenumber = txtPhoneNumber
End With

txtFirstName = ""
txtLastName = ""
txtAge = ""
txtID = ""
txtPayType = ""
txtWage = ""
txtPhoneNumber = ""
txtFirstName.SetFocus

End Sub

Private Sub cmdBack_Click()

Hide
frmMain.Show

End Sub

Private Sub cmdChange_Click()

With arremployees(entrytochange)
    .firstname = txtFirstName
    .lastname = txtLastName
    .age = txtAge
    .id = txtID
    .paytype = txtPayType
    .wage = txtWage
    .phonenumber = txtPhoneNumber
End With

Hide
frmMain.Show

End Sub

Private Sub cmdDelete_Click()

Dim i As Integer

For i = entrytodelete To entries
    arremployees(i).age = arremployees(i + 1).age
    arremployees(i).firstname = arremployees(i + 1).firstname
    arremployees(i).id = arremployees(i + 1).id
    arremployees(i).lastname = arremployees(i + 1).lastname
    arremployees(i).paytype = arremployees(i + 1).paytype
    arremployees(i).phonenumber = arremployees(i + 1).phonenumber
    arremployees(i).wage = arremployees(i + 1).wage
Next i

entries = entries - 1

Hide
frmMain.Show

End Sub

Private Sub cmdLeft_Click()

If roledexc = 1 Then
    roledexc = entries
Else
    roledexc = roledexc - 1
End If
RoledexChange

End Sub

Private Sub cmdRight_Click()

If roledexc = entries Then
    roledexc = 1
Else
    roledexc = roledexc + 1
End If
RoledexChange

End Sub

Private Sub Form_Load()

roledexc = 1

End Sub

Private Sub Form_Unload(Cancel As Integer)

Hide
frmMain.Show

End Sub
