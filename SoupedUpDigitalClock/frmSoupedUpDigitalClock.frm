VERSION 5.00
Begin VB.Form frmSoupedUpDigitalClock 
   Caption         =   "Souped Up Digital Clock"
   ClientHeight    =   3456
   ClientLeft      =   96
   ClientTop       =   416
   ClientWidth     =   3736
   LinkTopic       =   "Form1"
   ScaleHeight     =   3456
   ScaleWidth      =   3736
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   648
      Left            =   1728
      TabIndex        =   1
      Top             =   1792
      Width           =   1352
   End
   Begin VB.Timer Timer 
      Interval        =   1000
      Left            =   832
      Top             =   2112
   End
   Begin VB.Label lblTime 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "00:00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   49.29
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1160
      Left            =   640
      TabIndex        =   0
      Top             =   320
      Width           =   2440
   End
End
Attribute VB_Name = "frmSoupedUpDigitalClock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOk_Click()
End
End Sub

Sub Timer_Timer()
Dim nHeight As Integer
Dim n As Integer
n = Second(Now) Mod 10
If n = 0 Then
    nHeight = 8.25
ElseIf n = 1 Or n = 9 Then
    nHeight = 9.75
ElseIf n = 2 Or n = 8 Then
    nHeight = 12
ElseIf n = 3 Or n = 7 Then
    nHeight = 13.5
ElseIf n = 4 Or n = 6 Then
    nHeight = 18
Else
    nHeight = 24
End If
lblTime.FontSize = nHeight
lblTime.Caption = Time
End Sub
