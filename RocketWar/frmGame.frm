VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form frmGame 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Rocket War"
   ClientHeight    =   10635
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10875
   LinkTopic       =   "Form1"
   ScaleHeight     =   10635
   ScaleWidth      =   10875
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit Game"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1395
      Left            =   3000
      TabIndex        =   59
      Top             =   5400
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Timer tmrEnemyAnimation 
      Enabled         =   0   'False
      Interval        =   33
      Left            =   600
      Top             =   1740
   End
   Begin VB.Timer tmrBulletAnimation 
      Interval        =   33
      Left            =   1080
      Top             =   1260
   End
   Begin VB.Timer tmrShoot 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   600
      Top             =   1260
   End
   Begin VB.Timer tmrDownMove 
      Enabled         =   0   'False
      Interval        =   33
      Left            =   600
      Top             =   600
   End
   Begin VB.Timer tmrUpMove 
      Enabled         =   0   'False
      Interval        =   33
      Left            =   600
      Top             =   120
   End
   Begin VB.Timer tmrGlobal 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   120
      Top             =   1260
   End
   Begin VB.Timer tmrRightMove 
      Enabled         =   0   'False
      Interval        =   33
      Left            =   1080
      Top             =   600
   End
   Begin VB.Timer tmrLeftMove 
      Enabled         =   0   'False
      Interval        =   33
      Left            =   120
      Top             =   600
   End
   Begin VB.Label lblGameOver 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Game Over"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1995
      Left            =   0
      TabIndex        =   58
      Top             =   4020
      Visible         =   0   'False
      Width           =   10995
   End
   Begin VB.Label lblScore 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Score: 0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   57
      Top             =   0
      Width           =   5235
   End
   Begin VB.Label lblPhase 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Phase 1"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9600
      TabIndex        =   56
      Top             =   0
      Width           =   1275
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmpMusic 
      Height          =   555
      Left            =   10320
      TabIndex        =   55
      Top             =   1200
      Visible         =   0   'False
      Width           =   555
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   979
      _cy             =   979
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmpGo 
      Height          =   555
      Left            =   10320
      TabIndex        =   54
      Top             =   600
      Visible         =   0   'False
      Width           =   555
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   979
      _cy             =   979
   End
   Begin VB.Label lblGo 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Go!"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1995
      Left            =   -60
      TabIndex        =   53
      Top             =   4020
      Visible         =   0   'False
      Width           =   10995
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmpReady 
      Height          =   555
      Left            =   10320
      TabIndex        =   52
      Top             =   0
      Visible         =   0   'False
      Width           =   555
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   979
      _cy             =   979
   End
   Begin VB.Label lblReady 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Get Ready..."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1995
      Left            =   -60
      TabIndex        =   51
      Top             =   4020
      Visible         =   0   'False
      Width           =   10995
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   254
      Left            =   0
      Picture         =   "frmGame.frx":0000
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   253
      Left            =   0
      Picture         =   "frmGame.frx":0392
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   252
      Left            =   0
      Picture         =   "frmGame.frx":0724
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   251
      Left            =   0
      Picture         =   "frmGame.frx":0AB6
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   250
      Left            =   0
      Picture         =   "frmGame.frx":0E48
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   249
      Left            =   0
      Picture         =   "frmGame.frx":11DA
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   248
      Left            =   0
      Picture         =   "frmGame.frx":156C
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   247
      Left            =   0
      Picture         =   "frmGame.frx":18FE
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   246
      Left            =   0
      Picture         =   "frmGame.frx":1C90
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   245
      Left            =   0
      Picture         =   "frmGame.frx":2022
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   244
      Left            =   0
      Picture         =   "frmGame.frx":23B4
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   243
      Left            =   0
      Picture         =   "frmGame.frx":2746
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   242
      Left            =   0
      Picture         =   "frmGame.frx":2AD8
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   241
      Left            =   0
      Picture         =   "frmGame.frx":2E6A
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   240
      Left            =   0
      Picture         =   "frmGame.frx":31FC
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   239
      Left            =   0
      Picture         =   "frmGame.frx":358E
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   238
      Left            =   0
      Picture         =   "frmGame.frx":3920
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   237
      Left            =   0
      Picture         =   "frmGame.frx":3CB2
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   236
      Left            =   0
      Picture         =   "frmGame.frx":4044
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   235
      Left            =   0
      Picture         =   "frmGame.frx":43D6
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   234
      Left            =   0
      Picture         =   "frmGame.frx":4768
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   233
      Left            =   0
      Picture         =   "frmGame.frx":4AFA
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   232
      Left            =   0
      Picture         =   "frmGame.frx":4E8C
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   231
      Left            =   0
      Picture         =   "frmGame.frx":521E
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   230
      Left            =   0
      Picture         =   "frmGame.frx":55B0
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   229
      Left            =   0
      Picture         =   "frmGame.frx":5942
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   228
      Left            =   0
      Picture         =   "frmGame.frx":5CD4
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   227
      Left            =   0
      Picture         =   "frmGame.frx":6066
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   226
      Left            =   0
      Picture         =   "frmGame.frx":63F8
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   225
      Left            =   0
      Picture         =   "frmGame.frx":678A
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   224
      Left            =   0
      Picture         =   "frmGame.frx":6B1C
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   223
      Left            =   0
      Picture         =   "frmGame.frx":6EAE
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   222
      Left            =   0
      Picture         =   "frmGame.frx":7240
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   221
      Left            =   0
      Picture         =   "frmGame.frx":75D2
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   220
      Left            =   0
      Picture         =   "frmGame.frx":7964
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   219
      Left            =   0
      Picture         =   "frmGame.frx":7CF6
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   218
      Left            =   0
      Picture         =   "frmGame.frx":8088
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   217
      Left            =   0
      Picture         =   "frmGame.frx":841A
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   216
      Left            =   0
      Picture         =   "frmGame.frx":87AC
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   215
      Left            =   0
      Picture         =   "frmGame.frx":8B3E
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   214
      Left            =   0
      Picture         =   "frmGame.frx":8ED0
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   213
      Left            =   0
      Picture         =   "frmGame.frx":9262
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   212
      Left            =   0
      Picture         =   "frmGame.frx":95F4
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   211
      Left            =   0
      Picture         =   "frmGame.frx":9986
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   210
      Left            =   0
      Picture         =   "frmGame.frx":9D18
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   209
      Left            =   0
      Picture         =   "frmGame.frx":A0AA
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   208
      Left            =   0
      Picture         =   "frmGame.frx":A43C
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   207
      Left            =   0
      Picture         =   "frmGame.frx":A7CE
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   206
      Left            =   0
      Picture         =   "frmGame.frx":AB60
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   205
      Left            =   0
      Picture         =   "frmGame.frx":AEF2
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   204
      Left            =   0
      Picture         =   "frmGame.frx":B284
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   203
      Left            =   0
      Picture         =   "frmGame.frx":B616
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   202
      Left            =   0
      Picture         =   "frmGame.frx":B9A8
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   201
      Left            =   0
      Picture         =   "frmGame.frx":BD3A
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   200
      Left            =   0
      Picture         =   "frmGame.frx":C0CC
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   199
      Left            =   0
      Picture         =   "frmGame.frx":C45E
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   198
      Left            =   0
      Picture         =   "frmGame.frx":C7F0
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   197
      Left            =   0
      Picture         =   "frmGame.frx":CB82
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   196
      Left            =   0
      Picture         =   "frmGame.frx":CF14
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   195
      Left            =   0
      Picture         =   "frmGame.frx":D2A6
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   194
      Left            =   0
      Picture         =   "frmGame.frx":D638
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   193
      Left            =   0
      Picture         =   "frmGame.frx":D9CA
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   192
      Left            =   0
      Picture         =   "frmGame.frx":DD5C
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   191
      Left            =   0
      Picture         =   "frmGame.frx":E0EE
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   190
      Left            =   0
      Picture         =   "frmGame.frx":E480
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   189
      Left            =   0
      Picture         =   "frmGame.frx":E812
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   188
      Left            =   0
      Picture         =   "frmGame.frx":EBA4
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   187
      Left            =   0
      Picture         =   "frmGame.frx":EF36
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   186
      Left            =   0
      Picture         =   "frmGame.frx":F2C8
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   185
      Left            =   0
      Picture         =   "frmGame.frx":F65A
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   184
      Left            =   0
      Picture         =   "frmGame.frx":F9EC
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   183
      Left            =   0
      Picture         =   "frmGame.frx":FD7E
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   182
      Left            =   0
      Picture         =   "frmGame.frx":10110
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   181
      Left            =   0
      Picture         =   "frmGame.frx":104A2
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   180
      Left            =   0
      Picture         =   "frmGame.frx":10834
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   179
      Left            =   0
      Picture         =   "frmGame.frx":10BC6
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   178
      Left            =   0
      Picture         =   "frmGame.frx":10F58
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   177
      Left            =   0
      Picture         =   "frmGame.frx":112EA
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   176
      Left            =   0
      Picture         =   "frmGame.frx":1167C
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   175
      Left            =   0
      Picture         =   "frmGame.frx":11A0E
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   174
      Left            =   0
      Picture         =   "frmGame.frx":11DA0
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   173
      Left            =   0
      Picture         =   "frmGame.frx":12132
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   172
      Left            =   0
      Picture         =   "frmGame.frx":124C4
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   171
      Left            =   0
      Picture         =   "frmGame.frx":12856
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   170
      Left            =   0
      Picture         =   "frmGame.frx":12BE8
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   169
      Left            =   0
      Picture         =   "frmGame.frx":12F7A
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   168
      Left            =   0
      Picture         =   "frmGame.frx":1330C
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   167
      Left            =   0
      Picture         =   "frmGame.frx":1369E
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   166
      Left            =   0
      Picture         =   "frmGame.frx":13A30
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   165
      Left            =   0
      Picture         =   "frmGame.frx":13DC2
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   164
      Left            =   0
      Picture         =   "frmGame.frx":14154
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   163
      Left            =   0
      Picture         =   "frmGame.frx":144E6
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   162
      Left            =   0
      Picture         =   "frmGame.frx":14878
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   161
      Left            =   0
      Picture         =   "frmGame.frx":14C0A
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   160
      Left            =   0
      Picture         =   "frmGame.frx":14F9C
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   159
      Left            =   0
      Picture         =   "frmGame.frx":1532E
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   158
      Left            =   0
      Picture         =   "frmGame.frx":156C0
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   157
      Left            =   0
      Picture         =   "frmGame.frx":15A52
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   156
      Left            =   0
      Picture         =   "frmGame.frx":15DE4
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   155
      Left            =   0
      Picture         =   "frmGame.frx":16176
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   154
      Left            =   0
      Picture         =   "frmGame.frx":16508
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   153
      Left            =   0
      Picture         =   "frmGame.frx":1689A
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   152
      Left            =   0
      Picture         =   "frmGame.frx":16C2C
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   151
      Left            =   0
      Picture         =   "frmGame.frx":16FBE
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   150
      Left            =   0
      Picture         =   "frmGame.frx":17350
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   149
      Left            =   0
      Picture         =   "frmGame.frx":176E2
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   148
      Left            =   0
      Picture         =   "frmGame.frx":17A74
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   147
      Left            =   0
      Picture         =   "frmGame.frx":17E06
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   146
      Left            =   0
      Picture         =   "frmGame.frx":18198
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   145
      Left            =   0
      Picture         =   "frmGame.frx":1852A
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   144
      Left            =   0
      Picture         =   "frmGame.frx":188BC
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   143
      Left            =   0
      Picture         =   "frmGame.frx":18C4E
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   142
      Left            =   0
      Picture         =   "frmGame.frx":18FE0
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   141
      Left            =   0
      Picture         =   "frmGame.frx":19372
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   140
      Left            =   0
      Picture         =   "frmGame.frx":19704
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   139
      Left            =   0
      Picture         =   "frmGame.frx":19A96
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   138
      Left            =   0
      Picture         =   "frmGame.frx":19E28
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   137
      Left            =   0
      Picture         =   "frmGame.frx":1A1BA
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   136
      Left            =   0
      Picture         =   "frmGame.frx":1A54C
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   135
      Left            =   0
      Picture         =   "frmGame.frx":1A8DE
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   134
      Left            =   0
      Picture         =   "frmGame.frx":1AC70
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   133
      Left            =   0
      Picture         =   "frmGame.frx":1B002
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   132
      Left            =   0
      Picture         =   "frmGame.frx":1B394
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   131
      Left            =   0
      Picture         =   "frmGame.frx":1B726
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   130
      Left            =   0
      Picture         =   "frmGame.frx":1BAB8
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   129
      Left            =   0
      Picture         =   "frmGame.frx":1BE4A
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   128
      Left            =   0
      Picture         =   "frmGame.frx":1C1DC
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   127
      Left            =   0
      Picture         =   "frmGame.frx":1C56E
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   126
      Left            =   0
      Picture         =   "frmGame.frx":1C900
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   125
      Left            =   0
      Picture         =   "frmGame.frx":1CC92
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   124
      Left            =   0
      Picture         =   "frmGame.frx":1D024
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   123
      Left            =   0
      Picture         =   "frmGame.frx":1D3B6
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   122
      Left            =   0
      Picture         =   "frmGame.frx":1D748
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   121
      Left            =   0
      Picture         =   "frmGame.frx":1DADA
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   120
      Left            =   0
      Picture         =   "frmGame.frx":1DE6C
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   119
      Left            =   0
      Picture         =   "frmGame.frx":1E1FE
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   118
      Left            =   0
      Picture         =   "frmGame.frx":1E590
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   117
      Left            =   0
      Picture         =   "frmGame.frx":1E922
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   116
      Left            =   0
      Picture         =   "frmGame.frx":1ECB4
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   115
      Left            =   0
      Picture         =   "frmGame.frx":1F046
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   114
      Left            =   0
      Picture         =   "frmGame.frx":1F3D8
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   113
      Left            =   0
      Picture         =   "frmGame.frx":1F76A
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   112
      Left            =   0
      Picture         =   "frmGame.frx":1FAFC
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   111
      Left            =   0
      Picture         =   "frmGame.frx":1FE8E
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   110
      Left            =   0
      Picture         =   "frmGame.frx":20220
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   109
      Left            =   0
      Picture         =   "frmGame.frx":205B2
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   108
      Left            =   0
      Picture         =   "frmGame.frx":20944
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   107
      Left            =   0
      Picture         =   "frmGame.frx":20CD6
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   106
      Left            =   0
      Picture         =   "frmGame.frx":21068
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   105
      Left            =   0
      Picture         =   "frmGame.frx":213FA
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   104
      Left            =   0
      Picture         =   "frmGame.frx":2178C
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   103
      Left            =   0
      Picture         =   "frmGame.frx":21B1E
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   102
      Left            =   0
      Picture         =   "frmGame.frx":21EB0
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   101
      Left            =   0
      Picture         =   "frmGame.frx":22242
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   100
      Left            =   0
      Picture         =   "frmGame.frx":225D4
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   99
      Left            =   0
      Picture         =   "frmGame.frx":22966
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   98
      Left            =   0
      Picture         =   "frmGame.frx":22CF8
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   97
      Left            =   0
      Picture         =   "frmGame.frx":2308A
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   96
      Left            =   0
      Picture         =   "frmGame.frx":2341C
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   95
      Left            =   0
      Picture         =   "frmGame.frx":237AE
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   94
      Left            =   0
      Picture         =   "frmGame.frx":23B40
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   93
      Left            =   0
      Picture         =   "frmGame.frx":23ED2
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   92
      Left            =   0
      Picture         =   "frmGame.frx":24264
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   91
      Left            =   0
      Picture         =   "frmGame.frx":245F6
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   90
      Left            =   0
      Picture         =   "frmGame.frx":24988
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   89
      Left            =   0
      Picture         =   "frmGame.frx":24D1A
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   88
      Left            =   0
      Picture         =   "frmGame.frx":250AC
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   87
      Left            =   0
      Picture         =   "frmGame.frx":2543E
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   86
      Left            =   0
      Picture         =   "frmGame.frx":257D0
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   85
      Left            =   0
      Picture         =   "frmGame.frx":25B62
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   84
      Left            =   0
      Picture         =   "frmGame.frx":25EF4
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   83
      Left            =   0
      Picture         =   "frmGame.frx":26286
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   82
      Left            =   0
      Picture         =   "frmGame.frx":26618
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   81
      Left            =   0
      Picture         =   "frmGame.frx":269AA
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   80
      Left            =   0
      Picture         =   "frmGame.frx":26D3C
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   79
      Left            =   0
      Picture         =   "frmGame.frx":270CE
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   78
      Left            =   0
      Picture         =   "frmGame.frx":27460
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   77
      Left            =   0
      Picture         =   "frmGame.frx":277F2
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   76
      Left            =   0
      Picture         =   "frmGame.frx":27B84
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   75
      Left            =   0
      Picture         =   "frmGame.frx":27F16
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   74
      Left            =   0
      Picture         =   "frmGame.frx":282A8
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   73
      Left            =   0
      Picture         =   "frmGame.frx":2863A
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   72
      Left            =   0
      Picture         =   "frmGame.frx":289CC
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   71
      Left            =   0
      Picture         =   "frmGame.frx":28D5E
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   70
      Left            =   0
      Picture         =   "frmGame.frx":290F0
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   69
      Left            =   0
      Picture         =   "frmGame.frx":29482
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   68
      Left            =   0
      Picture         =   "frmGame.frx":29814
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   67
      Left            =   0
      Picture         =   "frmGame.frx":29BA6
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   66
      Left            =   0
      Picture         =   "frmGame.frx":29F38
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   65
      Left            =   0
      Picture         =   "frmGame.frx":2A2CA
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   64
      Left            =   0
      Picture         =   "frmGame.frx":2A65C
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   63
      Left            =   0
      Picture         =   "frmGame.frx":2A9EE
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   62
      Left            =   0
      Picture         =   "frmGame.frx":2AD80
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   61
      Left            =   0
      Picture         =   "frmGame.frx":2B112
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   60
      Left            =   0
      Picture         =   "frmGame.frx":2B4A4
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   59
      Left            =   0
      Picture         =   "frmGame.frx":2B836
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   58
      Left            =   0
      Picture         =   "frmGame.frx":2BBC8
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   57
      Left            =   0
      Picture         =   "frmGame.frx":2BF5A
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   56
      Left            =   0
      Picture         =   "frmGame.frx":2C2EC
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   55
      Left            =   0
      Picture         =   "frmGame.frx":2C67E
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   54
      Left            =   0
      Picture         =   "frmGame.frx":2CA10
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   53
      Left            =   0
      Picture         =   "frmGame.frx":2CDA2
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   52
      Left            =   0
      Picture         =   "frmGame.frx":2D134
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   51
      Left            =   0
      Picture         =   "frmGame.frx":2D4C6
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   50
      Left            =   0
      Picture         =   "frmGame.frx":2D858
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   49
      Left            =   0
      Picture         =   "frmGame.frx":2DBEA
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   48
      Left            =   0
      Picture         =   "frmGame.frx":2DF7C
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   47
      Left            =   0
      Picture         =   "frmGame.frx":2E30E
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   46
      Left            =   0
      Picture         =   "frmGame.frx":2E6A0
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   45
      Left            =   0
      Picture         =   "frmGame.frx":2EA32
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   44
      Left            =   0
      Picture         =   "frmGame.frx":2EDC4
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   43
      Left            =   0
      Picture         =   "frmGame.frx":2F156
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   42
      Left            =   0
      Picture         =   "frmGame.frx":2F4E8
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   41
      Left            =   0
      Picture         =   "frmGame.frx":2F87A
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   40
      Left            =   0
      Picture         =   "frmGame.frx":2FC0C
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   39
      Left            =   0
      Picture         =   "frmGame.frx":2FF9E
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   38
      Left            =   0
      Picture         =   "frmGame.frx":30330
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   37
      Left            =   0
      Picture         =   "frmGame.frx":306C2
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   36
      Left            =   0
      Picture         =   "frmGame.frx":30A54
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   35
      Left            =   0
      Picture         =   "frmGame.frx":30DE6
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   34
      Left            =   0
      Picture         =   "frmGame.frx":31178
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   33
      Left            =   0
      Picture         =   "frmGame.frx":3150A
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   32
      Left            =   0
      Picture         =   "frmGame.frx":3189C
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   31
      Left            =   0
      Picture         =   "frmGame.frx":31C2E
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   30
      Left            =   0
      Picture         =   "frmGame.frx":31FC0
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   29
      Left            =   0
      Picture         =   "frmGame.frx":32352
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   28
      Left            =   0
      Picture         =   "frmGame.frx":326E4
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   27
      Left            =   0
      Picture         =   "frmGame.frx":32A76
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   26
      Left            =   0
      Picture         =   "frmGame.frx":32E08
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   25
      Left            =   0
      Picture         =   "frmGame.frx":3319A
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   24
      Left            =   0
      Picture         =   "frmGame.frx":3352C
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   23
      Left            =   0
      Picture         =   "frmGame.frx":338BE
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   22
      Left            =   0
      Picture         =   "frmGame.frx":33C50
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   21
      Left            =   0
      Picture         =   "frmGame.frx":33FE2
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   20
      Left            =   0
      Picture         =   "frmGame.frx":34374
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   19
      Left            =   0
      Picture         =   "frmGame.frx":34706
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   18
      Left            =   0
      Picture         =   "frmGame.frx":34A98
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   17
      Left            =   0
      Picture         =   "frmGame.frx":34E2A
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   16
      Left            =   0
      Picture         =   "frmGame.frx":351BC
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   15
      Left            =   0
      Picture         =   "frmGame.frx":3554E
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   14
      Left            =   0
      Picture         =   "frmGame.frx":358E0
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   13
      Left            =   0
      Picture         =   "frmGame.frx":35C72
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   12
      Left            =   0
      Picture         =   "frmGame.frx":36004
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   11
      Left            =   0
      Picture         =   "frmGame.frx":36396
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   10
      Left            =   0
      Picture         =   "frmGame.frx":36728
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   9
      Left            =   0
      Picture         =   "frmGame.frx":36ABA
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   8
      Left            =   0
      Picture         =   "frmGame.frx":36E4C
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   7
      Left            =   0
      Picture         =   "frmGame.frx":371DE
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   6
      Left            =   0
      Picture         =   "frmGame.frx":37570
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   5
      Left            =   0
      Picture         =   "frmGame.frx":37902
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   4
      Left            =   0
      Picture         =   "frmGame.frx":37C94
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   3
      Left            =   0
      Picture         =   "frmGame.frx":38026
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   2
      Left            =   0
      Picture         =   "frmGame.frx":383B8
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   1
      Left            =   0
      Picture         =   "frmGame.frx":3874A
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEnemy 
      Height          =   240
      Index           =   0
      Left            =   0
      Picture         =   "frmGame.frx":38ADC
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label lblAmmo 
      BackColor       =   &H00FFFFFF&
      Caption         =   "50/50"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   10140
      TabIndex        =   50
      Top             =   10320
      Width           =   735
   End
   Begin VB.Label lblBullet 
      BackColor       =   &H00000000&
      Height          =   75
      Index           =   49
      Left            =   0
      TabIndex        =   49
      Top             =   0
      Visible         =   0   'False
      Width           =   50
   End
   Begin VB.Label lblBullet 
      BackColor       =   &H00000000&
      Height          =   75
      Index           =   48
      Left            =   0
      TabIndex        =   48
      Top             =   0
      Visible         =   0   'False
      Width           =   50
   End
   Begin VB.Label lblBullet 
      BackColor       =   &H00000000&
      Height          =   75
      Index           =   47
      Left            =   0
      TabIndex        =   47
      Top             =   0
      Visible         =   0   'False
      Width           =   50
   End
   Begin VB.Label lblBullet 
      BackColor       =   &H00000000&
      Height          =   75
      Index           =   46
      Left            =   0
      TabIndex        =   46
      Top             =   0
      Visible         =   0   'False
      Width           =   50
   End
   Begin VB.Label lblBullet 
      BackColor       =   &H00000000&
      Height          =   75
      Index           =   45
      Left            =   0
      TabIndex        =   45
      Top             =   0
      Visible         =   0   'False
      Width           =   50
   End
   Begin VB.Label lblBullet 
      BackColor       =   &H00000000&
      Height          =   75
      Index           =   44
      Left            =   0
      TabIndex        =   44
      Top             =   0
      Visible         =   0   'False
      Width           =   50
   End
   Begin VB.Label lblBullet 
      BackColor       =   &H00000000&
      Height          =   75
      Index           =   43
      Left            =   0
      TabIndex        =   43
      Top             =   0
      Visible         =   0   'False
      Width           =   50
   End
   Begin VB.Label lblBullet 
      BackColor       =   &H00000000&
      Height          =   75
      Index           =   42
      Left            =   0
      TabIndex        =   42
      Top             =   0
      Visible         =   0   'False
      Width           =   50
   End
   Begin VB.Label lblBullet 
      BackColor       =   &H00000000&
      Height          =   75
      Index           =   41
      Left            =   0
      TabIndex        =   41
      Top             =   0
      Visible         =   0   'False
      Width           =   50
   End
   Begin VB.Label lblBullet 
      BackColor       =   &H00000000&
      Height          =   75
      Index           =   40
      Left            =   0
      TabIndex        =   40
      Top             =   0
      Visible         =   0   'False
      Width           =   50
   End
   Begin VB.Label lblBullet 
      BackColor       =   &H00000000&
      Height          =   75
      Index           =   39
      Left            =   0
      TabIndex        =   39
      Top             =   0
      Visible         =   0   'False
      Width           =   50
   End
   Begin VB.Label lblBullet 
      BackColor       =   &H00000000&
      Height          =   75
      Index           =   38
      Left            =   0
      TabIndex        =   38
      Top             =   0
      Visible         =   0   'False
      Width           =   50
   End
   Begin VB.Label lblBullet 
      BackColor       =   &H00000000&
      Height          =   75
      Index           =   37
      Left            =   0
      TabIndex        =   37
      Top             =   0
      Visible         =   0   'False
      Width           =   50
   End
   Begin VB.Label lblBullet 
      BackColor       =   &H00000000&
      Height          =   75
      Index           =   36
      Left            =   0
      TabIndex        =   36
      Top             =   0
      Visible         =   0   'False
      Width           =   50
   End
   Begin VB.Label lblBullet 
      BackColor       =   &H00000000&
      Height          =   75
      Index           =   35
      Left            =   0
      TabIndex        =   35
      Top             =   0
      Visible         =   0   'False
      Width           =   50
   End
   Begin VB.Label lblBullet 
      BackColor       =   &H00000000&
      Height          =   75
      Index           =   34
      Left            =   0
      TabIndex        =   34
      Top             =   0
      Visible         =   0   'False
      Width           =   50
   End
   Begin VB.Label lblBullet 
      BackColor       =   &H00000000&
      Height          =   75
      Index           =   33
      Left            =   0
      TabIndex        =   33
      Top             =   0
      Visible         =   0   'False
      Width           =   50
   End
   Begin VB.Label lblBullet 
      BackColor       =   &H00000000&
      Height          =   75
      Index           =   32
      Left            =   0
      TabIndex        =   32
      Top             =   0
      Visible         =   0   'False
      Width           =   50
   End
   Begin VB.Label lblBullet 
      BackColor       =   &H00000000&
      Height          =   75
      Index           =   31
      Left            =   0
      TabIndex        =   31
      Top             =   0
      Visible         =   0   'False
      Width           =   50
   End
   Begin VB.Label lblBullet 
      BackColor       =   &H00000000&
      Height          =   75
      Index           =   30
      Left            =   0
      TabIndex        =   30
      Top             =   0
      Visible         =   0   'False
      Width           =   50
   End
   Begin VB.Label lblBullet 
      BackColor       =   &H00000000&
      Height          =   75
      Index           =   29
      Left            =   0
      TabIndex        =   29
      Top             =   0
      Visible         =   0   'False
      Width           =   50
   End
   Begin VB.Label lblBullet 
      BackColor       =   &H00000000&
      Height          =   75
      Index           =   28
      Left            =   0
      TabIndex        =   28
      Top             =   0
      Visible         =   0   'False
      Width           =   50
   End
   Begin VB.Label lblBullet 
      BackColor       =   &H00000000&
      Height          =   75
      Index           =   27
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Visible         =   0   'False
      Width           =   50
   End
   Begin VB.Label lblBullet 
      BackColor       =   &H00000000&
      Height          =   75
      Index           =   26
      Left            =   0
      TabIndex        =   26
      Top             =   0
      Visible         =   0   'False
      Width           =   50
   End
   Begin VB.Label lblBullet 
      BackColor       =   &H00000000&
      Height          =   75
      Index           =   25
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Visible         =   0   'False
      Width           =   50
   End
   Begin VB.Label lblBullet 
      BackColor       =   &H00000000&
      Height          =   75
      Index           =   24
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Visible         =   0   'False
      Width           =   50
   End
   Begin VB.Label lblBullet 
      BackColor       =   &H00000000&
      Height          =   75
      Index           =   23
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Visible         =   0   'False
      Width           =   50
   End
   Begin VB.Label lblBullet 
      BackColor       =   &H00000000&
      Height          =   75
      Index           =   22
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Visible         =   0   'False
      Width           =   50
   End
   Begin VB.Label lblBullet 
      BackColor       =   &H00000000&
      Height          =   75
      Index           =   21
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Visible         =   0   'False
      Width           =   50
   End
   Begin VB.Label lblBullet 
      BackColor       =   &H00000000&
      Height          =   75
      Index           =   20
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Visible         =   0   'False
      Width           =   50
   End
   Begin VB.Label lblBullet 
      BackColor       =   &H00000000&
      Height          =   75
      Index           =   19
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Visible         =   0   'False
      Width           =   50
   End
   Begin VB.Label lblBullet 
      BackColor       =   &H00000000&
      Height          =   75
      Index           =   18
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Visible         =   0   'False
      Width           =   50
   End
   Begin VB.Label lblBullet 
      BackColor       =   &H00000000&
      Height          =   75
      Index           =   17
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Visible         =   0   'False
      Width           =   50
   End
   Begin VB.Label lblBullet 
      BackColor       =   &H00000000&
      Height          =   75
      Index           =   16
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Visible         =   0   'False
      Width           =   50
   End
   Begin VB.Label lblBullet 
      BackColor       =   &H00000000&
      Height          =   75
      Index           =   15
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Visible         =   0   'False
      Width           =   50
   End
   Begin VB.Label lblBullet 
      BackColor       =   &H00000000&
      Height          =   75
      Index           =   14
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Visible         =   0   'False
      Width           =   50
   End
   Begin VB.Label lblBullet 
      BackColor       =   &H00000000&
      Height          =   75
      Index           =   13
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   50
   End
   Begin VB.Label lblBullet 
      BackColor       =   &H00000000&
      Height          =   75
      Index           =   12
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   50
   End
   Begin VB.Label lblBullet 
      BackColor       =   &H00000000&
      Height          =   75
      Index           =   11
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   50
   End
   Begin VB.Label lblBullet 
      BackColor       =   &H00000000&
      Height          =   75
      Index           =   10
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   50
   End
   Begin VB.Label lblBullet 
      BackColor       =   &H00000000&
      Height          =   75
      Index           =   9
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   50
   End
   Begin VB.Label lblBullet 
      BackColor       =   &H00000000&
      Height          =   75
      Index           =   8
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   50
   End
   Begin VB.Label lblBullet 
      BackColor       =   &H00000000&
      Height          =   75
      Index           =   7
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   50
   End
   Begin VB.Label lblBullet 
      BackColor       =   &H00000000&
      Height          =   75
      Index           =   6
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   50
   End
   Begin VB.Label lblBullet 
      BackColor       =   &H00000000&
      Height          =   75
      Index           =   5
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   50
   End
   Begin VB.Label lblBullet 
      BackColor       =   &H00000000&
      Height          =   75
      Index           =   4
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   50
   End
   Begin VB.Label lblBullet 
      BackColor       =   &H00000000&
      Height          =   75
      Index           =   3
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   50
   End
   Begin VB.Label lblBullet 
      BackColor       =   &H00000000&
      Height          =   75
      Index           =   2
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   50
   End
   Begin VB.Label lblBullet 
      BackColor       =   &H00000000&
      Height          =   75
      Index           =   1
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   50
   End
   Begin VB.Label lblBullet 
      BackColor       =   &H00000000&
      Height          =   75
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   50
   End
   Begin VB.Image imgShip 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   5260
      Picture         =   "frmGame.frx":38E6E
      Top             =   9840
      Width           =   240
   End
End
Attribute VB_Name = "frmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ssx As Integer
Dim ssy As Integer
Dim esx(1 To 255) As Integer
Dim numberofbullets As Integer
Dim numberofenemies As Integer
Dim bulletnumber As Integer
Dim reload As Boolean
Dim ammocounter As Integer
Dim gametime As Long
Dim phase As Integer
Dim score As Long
Dim gameovertest As Boolean

Sub SpawnEnemy()

numberofenemies = numberofenemies + 1

imgEnemy(numberofenemies - 1).Move Int(10995 * Rnd), 0
imgEnemy(numberofenemies - 1).Visible = True

End Sub

Sub KillEnemy(EnemyKilled As Integer)

imgEnemy(EnemyKilled).Visible = False

score = score + 100

lblScore = "Score: " + Str(score)

End Sub

Sub gameover()

Dim i As Integer

gameovertest = True
tmrGlobal.Enabled = False
imgShip.Visible = False

For i = 0 To 254

    imgEnemy(i).Visible = False

Next i

For i = 0 To 49

    lblBullet(i).Visible = False

Next i

lblGameOver.Visible = True
cmdExit.Visible = True
cmdExit.Enabled = True

globalscore = score

End Sub

Private Sub cmdExit_Click()

Hide

frmHighScores.Show

End Sub

Private Sub Form_Activate()

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If MouseControl = False And gameovertest = False Then

    If KeyCode = 37 Then
    
        tmrLeftMove.Enabled = True
        
    ElseIf KeyCode = 39 Then
    
        tmrRightMove.Enabled = True
        
    ElseIf KeyCode = 38 Then
    
        tmrUpMove.Enabled = True
        
    ElseIf KeyCode = 40 Then
    
        tmrDownMove.Enabled = True
        
    End If
    
End If

If KeyCode = 90 And reload = False And gameovertest = False Then

    tmrShoot.Enabled = True
    
End If

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

If MouseControl = False Then

    If KeyCode = 37 Then
    
        tmrLeftMove.Enabled = False
        
    ElseIf KeyCode = 39 Then
    
        tmrRightMove.Enabled = False
        
    ElseIf KeyCode = 38 Then
    
        tmrUpMove.Enabled = False
        
    ElseIf KeyCode = 40 Then
    
        tmrDownMove.Enabled = False
        
    End If
    
End If

If KeyCode = 90 Then

    tmrShoot.Enabled = False
    
End If

End Sub

Private Sub Form_Load()

Dim i As Integer
Dim tempnumber As Integer

ssx = 1
ssy = 1

For i = 1 To 255

    tempnumber = Int((2 * Rnd) + 1)
    
    If tempnumber = 1 Then
    
        esx(i) = 1
        
    ElseIf tempnumber = 2 Then
    
        esx(i) = -1
        
    End If

Next i

gametime = 0

tmrGlobal.Enabled = True

numberofbullets = 0
bulletnumber = 0
ammocounter = 50

Randomize

score = 0

gameovertest = False

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If MouseControl = True And gameovertest = False Then

    imgShip.Left = X - 120
    imgShip.Top = Y - 120
    
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

tmrGlobal.Enabled = False

Hide
frmMenu.Show

End Sub

Private Sub tmrBulletAnimation_Timer()

Dim i As Integer
Dim j As Integer

If reload = True And lblBullet(49).Visible = False Then

    numberofbullets = 0
    ammocounter = 50
    tmrShoot.Enabled = True
    reload = False
    
End If

For i = 1 To numberofbullets
    
    If lblBullet(i - 1).Visible = True Then
    
        lblBullet(i - 1).Move lblBullet(i - 1).Left, lblBullet(i - 1).Top - 300
        
        If lblBullet(49).Visible = True Then
            
            tmrShoot.Enabled = False
            reload = True
            
        End If
        
        'Collision
        If lblBullet(i - 1).Top <= 0 Then
        
            lblBullet(i - 1).Visible = False
            lblBullet(i - 1).Top = 0
            lblBullet(i - 1).Left = 0
            
        End If
        
        For j = 0 To 254
        
            If lblBullet(i - 1).Visible = True And imgEnemy(j).Visible = True And lblBullet(i - 1).Left + 75 >= imgEnemy(j).Left And lblBullet(i - 1).Left <= imgEnemy(j).Left + 240 And lblBullet(i - 1).Top + 75 >= imgEnemy(j).Top And lblBullet(i - 1).Top + 75 <= imgEnemy(j).Top + 240 Then
            
                KillEnemy (j)
                
            ElseIf lblBullet(i - 1).Visible = True And imgEnemy(j).Visible = True And lblBullet(i - 1).Top + 75 >= imgEnemy(j).Top And lblBullet(i - 1).Top <= imgEnemy(j).Top + 240 And lblBullet(i - 1).Left <= imgEnemy(j).Left + 240 And lblBullet(i - 1).Left >= imgEnemy(j).Left Then
            
                KillEnemy (j)
                
            ElseIf lblBullet(i - 1).Visible = True And imgEnemy(j).Visible = True And lblBullet(i - 1).Top + 75 >= imgEnemy(j).Top And lblBullet(i - 1).Top <= imgEnemy(j).Top + 240 And lblBullet(i - 1).Left + 75 >= imgEnemy(j).Left And lblBullet(i - 1).Left <= imgEnemy(j).Left + 240 Then
                
                KillEnemy (j)
                
            End If
        
        Next j
    
    End If
    
Next i

End Sub

Private Sub tmrDownMove_Timer()

ssy = 1

imgShip.Move imgShip.Left, imgShip.Top + ssy * 100

End Sub

Private Sub tmrEnemyAnimation_Timer()

Dim i As Integer

For i = 1 To numberofenemies

    If imgEnemy(i - 1).Visible = True Then
            
        If imgEnemy(i - 1).Left < 0 Then
        
            esx(i) = 1
            
        ElseIf imgEnemy(i - 1).Left > frmGame.Width - 250 Then
        
            esx(i) = -1
            
        End If
        
        If phase = 1 Then
        
            imgEnemy(i - 1).Move imgEnemy(i - 1).Left + esx(i) * 150, imgEnemy(i - 1).Top + 75
            
        ElseIf phase = 2 Then
            
            imgEnemy(i - 1).Move imgEnemy(i - 1).Left + esx(i) * 200, imgEnemy(i - 1).Top + 150
            
        ElseIf phase = 3 Then
        
            imgEnemy(i - 1).Move imgEnemy(i - 1).Left + esx(i) * 200, imgEnemy(i - 1).Top + 200
            
        ElseIf phase = 4 Then
        
            imgEnemy(i - 1).Move imgEnemy(i - 1).Left + esx(i) * 250, imgEnemy(i - 1).Top + 300
        
        End If
        
        'Collision
        If imgEnemy(i - 1).Visible = True And imgEnemy(i - 1).Left + 240 >= imgShip.Left And imgEnemy(i - 1).Left <= imgShip.Left + 240 And imgEnemy(i - 1).Top + 240 >= imgShip.Top And imgEnemy(i - 1).Top + 240 <= imgShip.Top + 240 Then
        
            gameover
            
        End If
        
    End If

Next i

End Sub

Private Sub tmrGlobal_Timer()

gametime = gametime + 1

If gametime = 1 Then

    lblReady.Visible = True
    wmpReady.URL = "D:\CP1\VB6 Project Files\RocketWar\resources\sound\ready.wav"
    
ElseIf gametime = 30 Then

    lblReady.Visible = False
    lblGo.Visible = True
    wmpGo.URL = "D:\CP1\VB6 Project Files\RocketWar\resources\sound\go.wav"
    
ElseIf gametime = 36 Then

    lblGo.Visible = False
    
'Phase 1
ElseIf gametime = 40 Then

    phase = 1
    
    wmpMusic.settings.setMode "loop", True
    wmpMusic.settings.volume = 10
    wmpMusic.URL = "D:\CP1\VB6 Project Files\RocketWar\resources\sound\music.wav"
    
    '1
    SpawnEnemy
    
    tmrEnemyAnimation.Enabled = True

ElseIf gametime = 50 Then

    '2
    SpawnEnemy
    
ElseIf gametime = 60 Then

    '3
    SpawnEnemy
    
ElseIf gametime = 70 Then

    '4
    SpawnEnemy
    
ElseIf gametime = 80 Then

    '5
    SpawnEnemy
    
ElseIf gametime = 90 Then

    '6
    SpawnEnemy
    
ElseIf gametime = 100 Then

    '7
    SpawnEnemy
    
ElseIf gametime = 110 Then

    '8
    SpawnEnemy
    
ElseIf gametime = 120 Then

    '9
    SpawnEnemy
    
ElseIf gametime = 130 Then

    '10
    SpawnEnemy
    
ElseIf gametime = 140 Then

    '11
    SpawnEnemy
    
ElseIf gametime = 150 Then

    '12
    SpawnEnemy
    
ElseIf gametime = 160 Then

    '13
    SpawnEnemy
    
ElseIf gametime = 170 Then

    '14
    SpawnEnemy

ElseIf gametime = 180 Then

    '15
    SpawnEnemy

'Phase 2
ElseIf gametime = 190 Then

    phase = 2
    lblPhase = "Phase 2"
    
    '16
    SpawnEnemy
    
ElseIf gametime = 195 Then

    '17
    SpawnEnemy
    
ElseIf gametime = 200 Then

    '18
    SpawnEnemy

ElseIf gametime = 205 Then

    '19
    SpawnEnemy
    
ElseIf gametime = 210 Then

    '20
    SpawnEnemy
    
ElseIf gametime = 215 Then

    '21
    SpawnEnemy
    
ElseIf gametime = 220 Then

    '22
    SpawnEnemy
    
ElseIf gametime = 225 Then

    '23
    SpawnEnemy
    
ElseIf gametime = 230 Then

    '24
    SpawnEnemy
    
ElseIf gametime = 235 Then

    '25
    SpawnEnemy
    
ElseIf gametime = 240 Then

    '26
    SpawnEnemy
    
ElseIf gametime = 245 Then

    '27
    SpawnEnemy

ElseIf gametime = 250 Then

    '28
    SpawnEnemy
    
ElseIf gametime = 255 Then

    '29
    SpawnEnemy
    
ElseIf gametime = 260 Then

    '30
    SpawnEnemy
    
ElseIf gametime = 265 Then

    '31
    SpawnEnemy
    
ElseIf gametime = 270 Then

    '32
    SpawnEnemy
    
ElseIf gametime = 275 Then

    '33
    SpawnEnemy
    
ElseIf gametime = 280 Then

    '34
    SpawnEnemy
    
ElseIf gametime = 285 Then

    '35
    SpawnEnemy
    
ElseIf gametime = 290 Then
    
    '36
    SpawnEnemy
    
ElseIf gametime = 295 Then

    '37
    SpawnEnemy
    
ElseIf gametime = 300 Then
    
    '38
    SpawnEnemy
    
ElseIf gametime = 305 Then
    
    '39
    SpawnEnemy
    
ElseIf gametime = 310 Then

    '40
    SpawnEnemy
    
ElseIf gametime = 315 Then

    '41
    SpawnEnemy
    
ElseIf gametime = 320 Then

    '42
    SpawnEnemy
    
ElseIf gametime = 325 Then

    '43
    SpawnEnemy
    
ElseIf gametime = 330 Then

    '44
    SpawnEnemy

ElseIf gametime = 335 Then
    
    '45
    SpawnEnemy
    
'Phase 3
ElseIf gametime = 340 Then
    
    phase = 3
    lblPhase = "Phase 3"
    
    '46
    SpawnEnemy
    
ElseIf gametime = 343 Then

    '47
    SpawnEnemy

ElseIf gametime = 346 Then

    '48
    SpawnEnemy
    
ElseIf gametime = 349 Then

    '49
    SpawnEnemy
    
ElseIf gametime = 352 Then

    '50
    SpawnEnemy
    
ElseIf gametime = 355 Then

    '51
    SpawnEnemy
    
ElseIf gametime = 358 Then

    '52
    SpawnEnemy
    
ElseIf gametime = 361 Then

    '52
    SpawnEnemy
    
ElseIf gametime = 364 Then

    '53
    SpawnEnemy

ElseIf gametime = 367 Then

    '54
    SpawnEnemy
    
ElseIf gametime = 370 Then

    '55
    SpawnEnemy
    
ElseIf gametime = 373 Then

    '56
    SpawnEnemy
    
ElseIf gametime = 376 Then

    '57
    SpawnEnemy
    
ElseIf gametime = 379 Then

    '58
    SpawnEnemy
    
ElseIf gametime = 382 Then

    '59
    SpawnEnemy
    
ElseIf gametime = 385 Then

    '60
    SpawnEnemy
    
ElseIf gametime = 388 Then

    '61
    SpawnEnemy
    
ElseIf gametime = 391 Then

    '62
    SpawnEnemy
    
ElseIf gametime = 394 Then

    '63
    SpawnEnemy
    
ElseIf gametime = 397 Then

    '64
    SpawnEnemy
    
ElseIf gametime = 400 Then

    '65
    SpawnEnemy
    
ElseIf gametime = 403 Then

    '66
    SpawnEnemy
    
ElseIf gametime = 406 Then

    '67
    SpawnEnemy
    
ElseIf gametime = 409 Then

    '68
    SpawnEnemy
    
ElseIf gametime = 412 Then

    '69
    SpawnEnemy
    
ElseIf gametime = 415 Then

    '70
    SpawnEnemy
    
ElseIf gametime = 418 Then

    '71
    SpawnEnemy
    
ElseIf gametime = 421 Then

    '72
    SpawnEnemy
    
ElseIf gametime = 424 Then

    '73
    SpawnEnemy
    
ElseIf gametime = 427 Then

    '74
    SpawnEnemy
    
ElseIf gametime = 430 Then

    '75
    SpawnEnemy
    
ElseIf gametime = 433 Then

    '76
    SpawnEnemy
    
ElseIf gametime = 436 Then

    '77
    SpawnEnemy
    
ElseIf gametime = 439 Then

    '78
    SpawnEnemy
    
ElseIf gametime = 442 Then

    '79
    SpawnEnemy
    
ElseIf gametime = 445 Then

    '80
    SpawnEnemy
    
ElseIf gametime = 448 Then

    '81
    SpawnEnemy
    
ElseIf gametime = 451 Then

    '82
    SpawnEnemy
    
ElseIf gametime = 454 Then

    '83
    SpawnEnemy
    
ElseIf gametime = 457 Then

    '84
    SpawnEnemy
    
ElseIf gametime = 460 Then

    '85
    SpawnEnemy
    
ElseIf gametime = 463 Then

    '86
    SpawnEnemy

ElseIf gametime = 466 Then

    '87
    SpawnEnemy
    
ElseIf gametime = 469 Then

    '88
    SpawnEnemy
    
ElseIf gametime = 472 Then

    '89
    SpawnEnemy
    
ElseIf gametime = 475 Then

    '90
    SpawnEnemy
    
ElseIf gametime = 478 Then

    '91
    SpawnEnemy
    
ElseIf gametime = 481 Then

    '92
    SpawnEnemy
    
ElseIf gametime = 484 Then

    '93
    SpawnEnemy
    
ElseIf gametime = 487 Then

    '94
    SpawnEnemy
    
ElseIf gametime = 490 Then

    '95
    SpawnEnemy
    
ElseIf gametime = 493 Then

    '96
    SpawnEnemy
    
ElseIf gametime = 496 Then

    '97
    SpawnEnemy
    
ElseIf gametime = 499 Then

    '98
    SpawnEnemy
    
ElseIf gametime = 502 Then

    '99
    SpawnEnemy
    
ElseIf gametime = 505 Then

    '100
    SpawnEnemy
    
ElseIf gametime = 508 Then

    '101
    SpawnEnemy
    
ElseIf gametime = 511 Then

    '102
    SpawnEnemy
    
ElseIf gametime = 514 Then

    '103
    SpawnEnemy
    
ElseIf gametime = 517 Then

    '104
    SpawnEnemy
    
ElseIf gametime = 520 Then

    '105
    SpawnEnemy
    
'Phase 4
ElseIf gametime = 523 Then

    phase = 4
    lblPhase = "Phase 4"
    '106
    SpawnEnemy
    

ElseIf gametime = 524 Then
'107
SpawnEnemy
ElseIf gametime = 525 Then
'108
SpawnEnemy
ElseIf gametime = 526 Then
'109
SpawnEnemy
ElseIf gametime = 527 Then
'110
SpawnEnemy
ElseIf gametime = 528 Then
'111
SpawnEnemy
ElseIf gametime = 529 Then
'112
SpawnEnemy
ElseIf gametime = 530 Then
'113
SpawnEnemy
ElseIf gametime = 531 Then
'114
SpawnEnemy
ElseIf gametime = 532 Then
'115
SpawnEnemy
ElseIf gametime = 533 Then
'116
SpawnEnemy
ElseIf gametime = 534 Then
'117
SpawnEnemy
ElseIf gametime = 535 Then
'118
SpawnEnemy
ElseIf gametime = 536 Then
'119
SpawnEnemy
ElseIf gametime = 537 Then
'120
SpawnEnemy
ElseIf gametime = 538 Then
'121
SpawnEnemy
ElseIf gametime = 539 Then
'122
SpawnEnemy
ElseIf gametime = 540 Then
'123
SpawnEnemy
ElseIf gametime = 541 Then
'124
SpawnEnemy
ElseIf gametime = 542 Then
'125
SpawnEnemy
ElseIf gametime = 543 Then
'126
SpawnEnemy
ElseIf gametime = 544 Then
'127
SpawnEnemy
ElseIf gametime = 545 Then
'128
SpawnEnemy
ElseIf gametime = 546 Then
'129
SpawnEnemy
ElseIf gametime = 547 Then
'130
SpawnEnemy
ElseIf gametime = 548 Then
'131
SpawnEnemy
ElseIf gametime = 549 Then
'132
SpawnEnemy
ElseIf gametime = 550 Then
'133
SpawnEnemy
ElseIf gametime = 551 Then
'134
SpawnEnemy
ElseIf gametime = 552 Then
'135
SpawnEnemy
ElseIf gametime = 553 Then
'136
SpawnEnemy
ElseIf gametime = 554 Then
'137
SpawnEnemy
ElseIf gametime = 555 Then
'138
SpawnEnemy
ElseIf gametime = 556 Then
'139
SpawnEnemy
ElseIf gametime = 557 Then
'140
SpawnEnemy
ElseIf gametime = 558 Then
'141
SpawnEnemy
ElseIf gametime = 559 Then
'142
SpawnEnemy
ElseIf gametime = 560 Then
'143
SpawnEnemy
ElseIf gametime = 561 Then
'144
SpawnEnemy
ElseIf gametime = 562 Then
'145
SpawnEnemy
ElseIf gametime = 563 Then
'146
SpawnEnemy
ElseIf gametime = 564 Then
'147
SpawnEnemy
ElseIf gametime = 565 Then
'148
SpawnEnemy
ElseIf gametime = 566 Then
'149
SpawnEnemy
ElseIf gametime = 567 Then
'150
SpawnEnemy
ElseIf gametime = 568 Then
'151
SpawnEnemy
ElseIf gametime = 569 Then
'152
SpawnEnemy
ElseIf gametime = 570 Then
'153
SpawnEnemy
ElseIf gametime = 571 Then
'154
SpawnEnemy
ElseIf gametime = 572 Then
'155
SpawnEnemy
ElseIf gametime = 573 Then
'156
SpawnEnemy
ElseIf gametime = 574 Then
'157
SpawnEnemy
ElseIf gametime = 575 Then
'158
SpawnEnemy
ElseIf gametime = 576 Then
'159
SpawnEnemy
ElseIf gametime = 577 Then
'160
SpawnEnemy
ElseIf gametime = 578 Then
'161
SpawnEnemy
ElseIf gametime = 579 Then
'162
SpawnEnemy
ElseIf gametime = 580 Then
'163
SpawnEnemy
ElseIf gametime = 581 Then
'164
SpawnEnemy
ElseIf gametime = 582 Then
'165
SpawnEnemy
ElseIf gametime = 583 Then
'166
SpawnEnemy
ElseIf gametime = 584 Then
'167
SpawnEnemy
ElseIf gametime = 585 Then
'168
SpawnEnemy
ElseIf gametime = 586 Then
'169
SpawnEnemy
ElseIf gametime = 587 Then
'170
SpawnEnemy
ElseIf gametime = 588 Then
'171
SpawnEnemy
ElseIf gametime = 589 Then
'172
SpawnEnemy
ElseIf gametime = 590 Then
'173
SpawnEnemy
ElseIf gametime = 591 Then
'174
SpawnEnemy
ElseIf gametime = 592 Then
'175
SpawnEnemy
ElseIf gametime = 593 Then
'176
SpawnEnemy
ElseIf gametime = 594 Then
'177
SpawnEnemy
ElseIf gametime = 595 Then
'178
SpawnEnemy
ElseIf gametime = 596 Then
'179
SpawnEnemy
ElseIf gametime = 597 Then
'180
SpawnEnemy
ElseIf gametime = 598 Then
'181
SpawnEnemy
ElseIf gametime = 599 Then
'182
SpawnEnemy
ElseIf gametime = 600 Then
'183
SpawnEnemy
ElseIf gametime = 601 Then
'184
SpawnEnemy
ElseIf gametime = 602 Then
'185
SpawnEnemy
ElseIf gametime = 603 Then
'186
SpawnEnemy
ElseIf gametime = 604 Then
'187
SpawnEnemy
ElseIf gametime = 605 Then
'188
SpawnEnemy
ElseIf gametime = 606 Then
'189
SpawnEnemy
ElseIf gametime = 607 Then
'190
SpawnEnemy
ElseIf gametime = 608 Then
'191
SpawnEnemy
ElseIf gametime = 609 Then
'192
SpawnEnemy
ElseIf gametime = 610 Then
'193
SpawnEnemy
ElseIf gametime = 611 Then
'194
SpawnEnemy
ElseIf gametime = 612 Then
'195
SpawnEnemy
ElseIf gametime = 613 Then
'196
SpawnEnemy
ElseIf gametime = 614 Then
'197
SpawnEnemy
ElseIf gametime = 615 Then
'198
SpawnEnemy
ElseIf gametime = 616 Then
'199
SpawnEnemy
ElseIf gametime = 617 Then
'200
SpawnEnemy
ElseIf gametime = 618 Then
'201
SpawnEnemy
ElseIf gametime = 619 Then
'202
SpawnEnemy
ElseIf gametime = 620 Then
'203
SpawnEnemy
ElseIf gametime = 621 Then
'204
SpawnEnemy
ElseIf gametime = 622 Then
'205
SpawnEnemy
ElseIf gametime = 623 Then
'206
SpawnEnemy
ElseIf gametime = 624 Then
'207
SpawnEnemy
ElseIf gametime = 625 Then
'208
SpawnEnemy
ElseIf gametime = 626 Then
'209
SpawnEnemy
ElseIf gametime = 627 Then
'210
SpawnEnemy
ElseIf gametime = 628 Then
'211
SpawnEnemy
ElseIf gametime = 629 Then
'212
SpawnEnemy
ElseIf gametime = 630 Then
'213
SpawnEnemy
ElseIf gametime = 631 Then
'214
SpawnEnemy
ElseIf gametime = 632 Then
'215
SpawnEnemy
ElseIf gametime = 633 Then
'216
SpawnEnemy
ElseIf gametime = 634 Then
'217
SpawnEnemy
ElseIf gametime = 635 Then
'218
SpawnEnemy
ElseIf gametime = 636 Then
'219
SpawnEnemy
ElseIf gametime = 637 Then
'220
SpawnEnemy
ElseIf gametime = 638 Then
'221
SpawnEnemy
ElseIf gametime = 639 Then
'222
SpawnEnemy
ElseIf gametime = 640 Then
'223
SpawnEnemy
ElseIf gametime = 641 Then
'224
SpawnEnemy
ElseIf gametime = 642 Then
'225
SpawnEnemy
ElseIf gametime = 643 Then
'226
SpawnEnemy
ElseIf gametime = 644 Then
'227
SpawnEnemy
ElseIf gametime = 645 Then
'228
SpawnEnemy
ElseIf gametime = 646 Then
'229
SpawnEnemy
ElseIf gametime = 647 Then
'230
SpawnEnemy
ElseIf gametime = 648 Then
'231
SpawnEnemy
ElseIf gametime = 649 Then
'232
SpawnEnemy
ElseIf gametime = 650 Then
'233
SpawnEnemy
ElseIf gametime = 651 Then
'234
SpawnEnemy
ElseIf gametime = 652 Then
'235
SpawnEnemy
ElseIf gametime = 653 Then
'236
SpawnEnemy
ElseIf gametime = 654 Then
'237
SpawnEnemy
ElseIf gametime = 655 Then
'238
SpawnEnemy
ElseIf gametime = 656 Then
'239
SpawnEnemy
ElseIf gametime = 657 Then
'240
SpawnEnemy
ElseIf gametime = 658 Then
'241
SpawnEnemy
ElseIf gametime = 659 Then
'242
SpawnEnemy
ElseIf gametime = 660 Then
'243
SpawnEnemy
ElseIf gametime = 661 Then
'244
SpawnEnemy
ElseIf gametime = 662 Then
'245
SpawnEnemy
ElseIf gametime = 663 Then
'246
SpawnEnemy
ElseIf gametime = 664 Then
'247
SpawnEnemy
ElseIf gametime = 665 Then
'248
SpawnEnemy
ElseIf gametime = 666 Then
'249
SpawnEnemy
ElseIf gametime = 667 Then
'250
SpawnEnemy
ElseIf gametime = 668 Then
'251
SpawnEnemy
ElseIf gametime = 669 Then
'252
SpawnEnemy
ElseIf gametime = 670 Then
'253
SpawnEnemy
ElseIf gametime = 671 Then
'254
SpawnEnemy
    
ElseIf gametime = 690 Then
gameover

End If

End Sub

Private Sub tmrLeftMove_Timer()

ssx = -1

imgShip.Move imgShip.Left + ssx * 100

End Sub

Private Sub tmrRightMove_Timer()

ssx = 1

imgShip.Move imgShip.Left + ssx * 100

End Sub

Private Sub tmrShoot_Timer()

numberofbullets = numberofbullets + 1

lblBullet(numberofbullets - 1).Visible = True
    
lblBullet(numberofbullets - 1).Move imgShip.Left + 100, imgShip.Top

tmrBulletAnimation.Enabled = True
    
If numberofbullets = 50 Then
        
    tmrShoot.Enabled = False
        
End If

ammocounter = ammocounter - 1

lblAmmo = Str(ammocounter) + "/50"

End Sub

Private Sub tmrUpMove_Timer()

ssy = -1

imgShip.Move imgShip.Left, imgShip.Top + ssy * 100

End Sub
