VERSION 5.00
Begin VB.Form frmShooter 
   Caption         =   "Shooter"
   ClientHeight    =   2595
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2850
   LinkTopic       =   "Form1"
   ScaleHeight     =   2595
   ScaleWidth      =   2850
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrEnemy 
      Enabled         =   0   'False
      Interval        =   25
      Left            =   60
      Top             =   660
   End
   Begin VB.Timer tmrBullet 
      Enabled         =   0   'False
      Interval        =   25
      Left            =   60
      Top             =   60
   End
   Begin VB.Image imgObject4 
      Height          =   4500
      Left            =   -240
      Picture         =   "frmShooter.frx":0000
      Top             =   -960
      Visible         =   0   'False
      Width           =   4500
   End
   Begin VB.Label lblBullet 
      BackColor       =   &H00000000&
      Height          =   255
      Left            =   1320
      TabIndex        =   0
      Top             =   1860
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Image imgObject3 
      Height          =   240
      Left            =   0
      Picture         =   "frmShooter.frx":8FB8
      Top             =   360
      Width           =   240
   End
   Begin VB.Image imgObject1 
      Height          =   240
      Left            =   1260
      Picture         =   "frmShooter.frx":934A
      Top             =   2160
      Width           =   240
   End
End
Attribute VB_Name = "frmShooter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim directionchange As Integer
Dim oby As Integer
Dim bully As Integer

Private Sub Form_Activate()

tmrEnemy.Enabled = True

directionchange = 0

oby = 1

bully = -1

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

lblBullet.Visible = True

tmrBullet.Enabled = True

End Sub

Private Sub Form_Unload(Cancel As Integer)

Hide

frmMainMenu.Show

End Sub

Private Sub tmrBullet_Timer()

lblBullet.Move lblBullet.Left, lblBullet.Top + bully * 150

If lblBullet.Left + lblBullet.Width >= imgObject3.Left And lblBullet.Left <= imgObject3.Left + imgObject3.Width And lblBullet.Top <= imgObject3.Top + imgObject3.Height And lblBullet.Top > imgObject3.Top Then

    tmrEnemy.Enabled = False
    
    imgObject3.Visible = False

    imgObject4.Visible = True
    
End If

End Sub

Private Sub tmrEnemy_Timer()

directionchange = directionchange + 1

If directionchange = 60 Then

    oby = -1
    
ElseIf directionchange = 120 Then

    oby = 1
    directionchange = 0
    
End If

imgObject3.Move imgObject3.Left + oby * 50

End Sub
