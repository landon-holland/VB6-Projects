VERSION 5.00
Begin VB.Form frmPong 
   Caption         =   "Pong"
   ClientHeight    =   7590
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7875
   LinkTopic       =   "Form1"
   ScaleHeight     =   7590
   ScaleWidth      =   7875
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrBall 
      Interval        =   33
      Left            =   60
      Top             =   60
   End
   Begin VB.Label lblPadel 
      BackColor       =   &H00000000&
      Height          =   255
      Left            =   3360
      TabIndex        =   0
      Top             =   6840
      Width           =   915
   End
   Begin VB.Image imgBall 
      Height          =   180
      Left            =   1860
      Picture         =   "frmPong.frx":0000
      Top             =   6540
      Width           =   180
   End
End
Attribute VB_Name = "frmPong"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bsx As Integer
Dim bsy As Integer

Private Sub Form_Load()

bsx = 1
bsy = 1

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

lblPadel.Left = X - 457

End Sub

Private Sub tmrBall_Timer()

'Animation
imgBall.Move imgBall.Left + bsx * 80, imgBall.Top + bsy * 80

'Check for borders
If imgBall.Top + imgBall.Height + 450 > Height Then

    bsy = bsy * -1
    
End If

If imgBall.Left + imgBall.Width + 190 > Width Then

    bsx = bsx * -1
    
End If

If imgBall.Top < 0 Then

    bsy = bsy * -1
    
End If

If imgBall.Left - 100 < 0 Then

    bsx = bsx * -1
    
End If


'Padel checking.
If imgBall.Left > lblPadel.Left And imgBall.Left + imgBall.Width < lblPadel.Left + lblPadel.Width And imgBall.Top + imgBall.Height > lblPadel.Top Then

    bsy = bsy * -1
    
End If
End Sub
