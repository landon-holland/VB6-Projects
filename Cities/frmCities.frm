VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form frmCities 
   Caption         =   "Cities"
   ClientHeight    =   5490
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   17565
   LinkTopic       =   "Form1"
   ScaleHeight     =   5490
   ScaleWidth      =   17565
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstMerge123 
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
      Height          =   1590
      ItemData        =   "frmCities.frx":0000
      Left            =   15300
      List            =   "frmCities.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   34
      Top             =   3120
      Width           =   1935
   End
   Begin VB.ListBox lstMerge23 
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
      Height          =   1590
      ItemData        =   "frmCities.frx":0004
      Left            =   15300
      List            =   "frmCities.frx":0006
      Sorted          =   -1  'True
      TabIndex        =   31
      Top             =   600
      Width           =   1935
   End
   Begin VB.ListBox lstMerge13 
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
      Height          =   1590
      ItemData        =   "frmCities.frx":0008
      Left            =   13260
      List            =   "frmCities.frx":000A
      Sorted          =   -1  'True
      TabIndex        =   28
      Top             =   3120
      Width           =   1935
   End
   Begin VB.ListBox lstCommon123 
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
      Height          =   1590
      ItemData        =   "frmCities.frx":000C
      Left            =   10980
      List            =   "frmCities.frx":000E
      Sorted          =   -1  'True
      TabIndex        =   25
      Top             =   3120
      Width           =   1935
   End
   Begin VB.ListBox lstCommon23 
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
      Height          =   1590
      ItemData        =   "frmCities.frx":0010
      Left            =   10980
      List            =   "frmCities.frx":0012
      Sorted          =   -1  'True
      TabIndex        =   22
      Top             =   600
      Width           =   1935
   End
   Begin VB.ListBox lstCommon13 
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
      Height          =   1590
      ItemData        =   "frmCities.frx":0014
      Left            =   8940
      List            =   "frmCities.frx":0016
      Sorted          =   -1  'True
      TabIndex        =   19
      Top             =   3120
      Width           =   1935
   End
   Begin MSComDlg.CommonDialog cndSaveAs 
      Left            =   1740
      Top             =   1140
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Save As"
      Filter          =   "Text (*.txt)|*.txt|All Files (*.*)|*.*"
      InitDir         =   "D:\CP1\VB6 Project Files\Cities\lists"
   End
   Begin MSComDlg.CommonDialog cndOpen 
      Left            =   1740
      Top             =   660
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Open"
      Filter          =   "Text (*.txt)|*.txt|All Files (*.*)|*.*"
      InitDir         =   "D:\CP1\VB6 Project Files\Cities\lists"
   End
   Begin VB.ListBox lstMerge12 
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
      Height          =   1590
      ItemData        =   "frmCities.frx":0018
      Left            =   13260
      List            =   "frmCities.frx":001A
      Sorted          =   -1  'True
      TabIndex        =   16
      Top             =   600
      Width           =   1935
   End
   Begin VB.ListBox lstCommon12 
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
      Height          =   1590
      ItemData        =   "frmCities.frx":001C
      Left            =   8940
      List            =   "frmCities.frx":001E
      Sorted          =   -1  'True
      TabIndex        =   13
      Top             =   600
      Width           =   1935
   End
   Begin VB.ListBox lstCity3 
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
      Height          =   4125
      ItemData        =   "frmCities.frx":0020
      Left            =   6660
      List            =   "frmCities.frx":0022
      Sorted          =   -1  'True
      TabIndex        =   7
      Top             =   600
      Width           =   1935
   End
   Begin VB.ListBox lstCity2 
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
      Height          =   4125
      ItemData        =   "frmCities.frx":0024
      Left            =   4620
      List            =   "frmCities.frx":0026
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   600
      Width           =   1935
   End
   Begin VB.ListBox lstCity1 
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
      Height          =   4125
      ItemData        =   "frmCities.frx":0028
      Left            =   2580
      List            =   "frmCities.frx":002A
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   600
      Width           =   1935
   End
   Begin VB.TextBox txtList 
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
      Height          =   4155
      Left            =   300
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label lblListMerge123Number 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   15780
      TabIndex        =   35
      Top             =   4920
      Width           =   975
   End
   Begin VB.Label lbllstMerge123Title 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Merge 1-2-3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   15300
      TabIndex        =   33
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Label lblListMerge23Number 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   15780
      TabIndex        =   32
      Top             =   2340
      Width           =   975
   End
   Begin VB.Label lbllstMerge23Title 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Merge 2-3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   15300
      TabIndex        =   30
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label lblListMerge13Number 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   13740
      TabIndex        =   29
      Top             =   4920
      Width           =   975
   End
   Begin VB.Label lbllstMerge13Title 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Merge 1-3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   13260
      TabIndex        =   27
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Label lblListCommon123Number 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   11520
      TabIndex        =   26
      Top             =   4920
      Width           =   975
   End
   Begin VB.Label lblListCommon23Number 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   11460
      TabIndex        =   24
      Top             =   2340
      Width           =   975
   End
   Begin VB.Label lbllstCommon123Title 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Common 1-2-3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   10980
      TabIndex        =   23
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Label lbllstCommon23Title 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Common 2-3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   10980
      TabIndex        =   21
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label lblListCommon13Number 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   9420
      TabIndex        =   20
      Top             =   4920
      Width           =   975
   End
   Begin VB.Label lbllstCommon13Title 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Common 1-3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   8940
      TabIndex        =   18
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Label lblListMerge12Number 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   13740
      TabIndex        =   17
      Top             =   2340
      Width           =   975
   End
   Begin VB.Label lbllstMerge12Title 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Merge 1-2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   13260
      TabIndex        =   15
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label lblListCommon12Number 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   9420
      TabIndex        =   14
      Top             =   2340
      Width           =   975
   End
   Begin VB.Label lbllstCommon12Title 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Common 1-2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   8940
      TabIndex        =   12
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label lblList3Number 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7140
      TabIndex        =   11
      Top             =   4920
      Width           =   975
   End
   Begin VB.Label lblList2Number 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5100
      TabIndex        =   10
      Top             =   4920
      Width           =   975
   End
   Begin VB.Label lblList1Number 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3060
      TabIndex        =   9
      Top             =   4920
      Width           =   975
   End
   Begin VB.Label lblNumberPrompt 
      Caption         =   "Number Of Workers:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   300
      TabIndex        =   8
      Top             =   4860
      Width           =   1935
   End
   Begin VB.Label lbllstCity3Title 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   6660
      TabIndex        =   6
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label lbllstCity2Title 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   4620
      TabIndex        =   5
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label lbllstCity1Title 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   2580
      TabIndex        =   3
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label lbltxtListTitle 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   300
      TabIndex        =   1
      Top             =   240
      Width           =   1935
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuNew 
         Caption         =   "New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save As..."
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuCut 
         Caption         =   "Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "Paste"
         Enabled         =   0   'False
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuUndo 
         Caption         =   "Undo"
      End
      Begin VB.Menu mnuSelectAll 
         Caption         =   "Select All"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu mnuList1 
         Caption         =   "Open to List 1"
      End
      Begin VB.Menu mnuList2 
         Caption         =   "Open to List 2"
      End
      Begin VB.Menu mnuList3 
         Caption         =   "Open to List 3"
      End
      Begin VB.Menu mnuCommon 
         Caption         =   "Common"
      End
      Begin VB.Menu mnuMerge 
         Caption         =   "Merge"
      End
      Begin VB.Menu mnuClear 
         Caption         =   "Clear"
      End
   End
End
Attribute VB_Name = "frmCities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Path As String
Dim OldBody As String
Dim NewBody As String

Sub OpenFile()

Dim Done As Boolean
Dim Answer As String
Dim FileName As String

Done = False
Do While Done = False

    cndOpen.ShowOpen
    Path = cndOpen.FileName
    cndOpen.FileName = ""
    
    If Path = "" Then
            
        Exit Do
        txtList.SetFocus
        
    Else
        
        Answer = MsgBox(Path, vbYesNo, "Is this the file?")
        
        If Answer = vbYes Then
            Open Path For Input As #1

                Dim FileSize As Long
                FileSize = LOF(1)
                txtList = Input(FileSize, 1)

            Close #1

            OldBody = txtList
            Done = True
            lbltxtListTitle = cndOpen.FileTitle
            
        End If
        
    End If

Loop

End Sub

Sub SaveFile()

If Path = "" Then

    SaveAsFile
    
Else

    Open Path For Output As #1

        Print #1, txtList

    Close #1
    
    OldBody = txtList
    
End If

End Sub

Sub SaveAsFile()

Dim Done As Boolean
Dim FileName As String
Dim Answer As String

Done = False
Do While Done = False

    cndSaveAs.Filter = "Text (*.txt)|*.txt|All Files (*.*)|*.*"
    cndSaveAs.ShowSave
    Path = cndSaveAs.FileName
    cndSaveAs.FileName = ""
    
    If Path = "" Then
    
        txtList.SetFocus
        Exit Do
        
    Else
    
        Answer = MsgBox(Path, vbYesNo, "Is this the file?")
        If Answer = vbYes Then

            Open Path For Output As #1

                Print #1, txtList

            Close #1
            
            Done = True
            
        End If
        
        lbltxtListTitle = cndSaveAs.FileTitle

        OldBody = txtList
        
    End If

Loop

End Sub

Private Sub Form_Load()

NewBody = ""
OldBody = ""

End Sub

Private Sub Form_Unload(Cancel As Integer)

Dim Answer As String

If OldBody <> NewBody Then

    Answer = MsgBox("Do you want to save your changes?", vbYesNo, "Exit")
    
    If Answer = vbNo Then
    
        End
    
    Else
        
        SaveFile
        End
        
    End If
    
Else

    End
    
End If

End Sub

Private Sub mnuClear_Click()

lstCity1.Clear
lstCity2.Clear
lstCity3.Clear
lstCommon12.Clear
lstCommon23.Clear
lstCommon13.Clear
lstCommon123.Clear
lstMerge12.Clear
lstMerge23.Clear
lstMerge13.Clear
lstMerge123.Clear

lblList1Number = ""
lblList2Number = ""
lblList3Number = ""
lblListCommon12Number = ""
lblListCommon23Number = ""
lblListCommon13Number = ""
lblListCommon123Number = ""
lblListMerge12Number = ""
lblListMerge23Number = ""
lblListMerge13Number = ""
lblListMerge123Number = ""

lbllstCity1Title = ""
lbllstCity2Title = ""
lbllstCity3Title = ""

End Sub

Private Sub mnuCommon_Click()

Dim i As Integer
Dim j As Integer
Dim k As Integer

'1-2
For i = 0 To lstCity1.ListCount - 1

    For j = 0 To lstCity2.ListCount - 1
    
        If StrComp(lstCity1.List(i), lstCity2.List(j)) = 0 Then
        
            lstCommon12.AddItem lstCity1.List(i)
            
        End If
        
    Next j
    
Next i

lblListCommon12Number = lstCommon12.ListCount

'2-3
For i = 0 To lstCity2.ListCount - 1

    For j = 0 To lstCity3.ListCount - 1
    
        If StrComp(lstCity2.List(i), lstCity3.List(j)) = 0 Then
        
            lstCommon23.AddItem lstCity2.List(i)
            
        End If
        
    Next j
    
Next i

lblListCommon23Number = lstCommon23.ListCount

'1-3
For i = 0 To lstCity1.ListCount - 1

    For j = 0 To lstCity3.ListCount - 1
    
        If StrComp(lstCity1.List(i), lstCity3.List(j)) = 0 Then
        
            lstCommon13.AddItem lstCity1.List(i)
            
        End If
        
    Next j
    
Next i

lblListCommon13Number = lstCommon13.ListCount

'1-2-3
For i = 0 To lstCity1.ListCount - 1

    For j = 0 To lstCity2.ListCount - 1
    
        For k = 0 To lstCity3.ListCount - 1
        
            If StrComp(lstCity1.List(i), lstCity2.List(j)) = 0 And StrComp(lstCity2.List(j), lstCity3.List(k)) = 0 And StrComp(lstCity1.List(i), lstCity3.List(k)) = 0 Then
            
                lstCommon123.AddItem lstCity1.List(i)
                
            End If
            
        Next k
        
    Next j
    
Next i

lblListCommon123Number = lstCommon123.ListCount

End Sub

Private Sub mnuCopy_Click()

Clipboard.SetText txtList.SelText
mnuPaste.Enabled = True

End Sub

Private Sub mnuCut_Click()

Clipboard.SetText txtList.SelText
txtList.SelText = ""
mnuPaste.Enabled = True

End Sub

Private Sub mnuExit_Click()

Dim Answer As String

If OldBody <> NewBody Then

    Answer = MsgBox("Do you want to save your changes?", vbYesNo, "Exit")
    
    If Answer = vbNo Then
    
        End
    
    Else
        
        SaveFile
        End
        
    End If
    
Else

    End
    
End If

End Sub

Private Sub mnuList1_Click()

Dim Done As Boolean
Dim Answer As String
Dim FileName As String

Done = False
Do While Done = False

    cndOpen.ShowOpen
    Path = cndOpen.FileName
    cndOpen.FileName = ""
    
    If Path = "" Then
            
        Exit Do
        txtList.SetFocus
        
    Else
        
        Answer = MsgBox(Path, vbYesNo, "Is this the file?")
        
        If Answer = vbYes Then
        
            Dim Line As String
            Dim i As Integer
            
            Open Path For Input As #1

                For i = 1 To 20
                
                    If EOF(1) Then
                        
                        Exit For
                        
                    End If
                    
                    Line Input #1, Line
                    
                    Line = Trim(Line)
                    
                    If Len(Line) <> 0 Then
                    
                        lstCity1.AddItem Line
                        
                    End If
                    
                Next i

            Close #1
            
            lbllstCity1Title = cndOpen.FileTitle
            lblList1Number = lstCity1.ListCount
            
            Done = True
            
        End If
        
    End If

Loop

End Sub

Private Sub mnuList2_Click()

Dim Done As Boolean
Dim Answer As String
Dim FileName As String

Done = False
Do While Done = False

    cndOpen.ShowOpen
    Path = cndOpen.FileName
    cndOpen.FileName = ""
    
    If Path = "" Then
            
        Exit Do
        txtList.SetFocus
        
    Else
        
        Answer = MsgBox(Path, vbYesNo, "Is this the file?")
        
        If Answer = vbYes Then
        
            Dim Line As String
            Dim i As Integer
            
            Open Path For Input As #1

                For i = 1 To 20
                
                    If EOF(1) Then
                        
                        Exit For
                        
                    End If
                    
                    Line Input #1, Line
                    
                    Line = Trim(Line)
                    
                    If Len(Line) <> 0 Then
                    
                        lstCity2.AddItem Line
                        
                    End If
                    
                Next i

            Close #1

            lbllstCity2Title = cndOpen.FileTitle
            lblList2Number = lstCity2.ListCount
            
            Done = True
            
        End If
        
    End If

Loop

End Sub

Private Sub mnuList3_Click()

Dim Done As Boolean
Dim Answer As String
Dim FileName As String

Done = False
Do While Done = False

    cndOpen.ShowOpen
    Path = cndOpen.FileName
    cndOpen.FileName = ""
    
    If Path = "" Then
            
        Exit Do
        txtList.SetFocus
        
    Else
        
        Answer = MsgBox(Path, vbYesNo, "Is this the file?")
        
        If Answer = vbYes Then
        
            Dim Line As String
            Dim i As Integer
            
            Open Path For Input As #1

                For i = 1 To 20
                
                    If EOF(1) Then
                        
                        Exit For
                        
                    End If
                    
                    Line Input #1, Line
                    
                    Line = Trim(Line)
                    
                    If Len(Line) <> 0 Then
                    
                        lstCity3.AddItem Line
                        
                    End If
                    
                Next i

            Close #1
            
            lbllstCity3Title = cndOpen.FileTitle
            lblList3Number = lstCity3.ListCount

            Done = True
            
        End If
        
    End If

Loop

End Sub

Private Sub mnuMerge_Click()

Dim i As Integer
Dim j As Integer
Dim k As Integer

'1-2
For i = 0 To lstCity1.ListCount - 1

    lstMerge12.AddItem lstCity1.List(i)
    
Next i

For i = 0 To lstCity2.ListCount - 1

    lstMerge12.AddItem lstCity2.List(i)
    
Next i

For i = 0 To lstMerge12.ListCount - 1

    For j = i + 1 To lstMerge12.ListCount - 1
    
        If StrComp(lstMerge12.List(i), lstMerge12.List(j)) = 0 Then
        
            lstMerge12.RemoveItem j
            
        End If
        
    Next j
    
Next i

lblListMerge12Number = lstMerge12.ListCount

'2-3
For i = 0 To lstCity2.ListCount - 1

    lstMerge23.AddItem lstCity2.List(i)
    
Next i

For i = 0 To lstCity3.ListCount - 1

    lstMerge23.AddItem lstCity3.List(i)
    
Next i

For i = 0 To lstMerge23.ListCount - 1

    For j = i + 1 To lstMerge23.ListCount - 1
    
        If StrComp(lstMerge23.List(i), lstMerge23.List(j)) = 0 Then
        
            lstMerge23.RemoveItem j
            
        End If
        
    Next j
    
Next i

lblListMerge23Number = lstMerge23.ListCount

'1-3
For i = 0 To lstCity1.ListCount - 1

    lstMerge13.AddItem lstCity1.List(i)
    
Next i

For i = 0 To lstCity3.ListCount - 1

    lstMerge13.AddItem lstCity3.List(i)
    
Next i

For i = 0 To lstMerge13.ListCount - 1

    For j = i + 1 To lstMerge13.ListCount - 1
    
        If StrComp(lstMerge13.List(i), lstMerge13.List(j)) = 0 Then
        
            lstMerge13.RemoveItem j
            
        End If
        
    Next j
    
Next i

lblListMerge13Number = lstMerge13.ListCount

'1-2-3
For i = 0 To lstCity1.ListCount - 1

    lstMerge123.AddItem lstCity1.List(i)
    
Next i

For i = 0 To lstCity2.ListCount - 1

    lstMerge123.AddItem lstCity2.List(i)
    
Next i

For i = 0 To lstCity3.ListCount - 1

    lstMerge123.AddItem lstCity3.List(i)
    
Next i

For i = 0 To lstMerge123.ListCount - 1

    For j = i + 1 To lstMerge123.ListCount - 1
    
        If StrComp(lstMerge123.List(i), lstMerge123.List(j)) = 0 Then
        
            lstMerge123.RemoveItem j
            
        End If
        
    Next j
    
Next i

For i = 0 To lstMerge123.ListCount - 1

    For j = i + 1 To lstMerge123.ListCount - 1
    
        If StrComp(lstMerge123.List(i), lstMerge123.List(j)) = 0 Then
        
            lstMerge123.RemoveItem j
            
        End If
        
    Next j
    
Next i

lblListMerge123Number = lstMerge123.ListCount

End Sub

Private Sub mnuNew_Click()

Dim Answer As String

If OldBody <> NewBody Then
    
    Answer = MsgBox("Do you want to save your changes?", vbYesNo, "New")
    
    If Answer = vbNo Then
    
        txtList = ""
        Path = ""
        OldBody = ""
        lbltxtListTitle = ""
        
    Else
    
        SaveFile
        txtList = ""
        Path = ""
        OldBody = ""
        lbltxtListTitle = ""
        
    End If

Else

    txtList = ""
    Path = ""
    OldBody = ""
    lbltxtListTitle = ""
    
End If

End Sub

Private Sub mnuOpen_Click()

Dim Answer As String

If OldBody <> NewBody Then

    Answer = MsgBox("Do you want to save your changes?", vbYesNo, "Open")
    
    If Answer = vbNo Then
    
        OpenFile
    
    Else
    
        SaveFile
        OpenFile
    
    End If

Else

    OpenFile

End If

End Sub

Private Sub mnuPaste_Click()

txtList.SelText = Clipboard.GetText()

End Sub

Private Sub mnuSave_Click()

SaveFile

End Sub

Private Sub mnuSaveAs_Click()

SaveAsFile

End Sub

Private Sub mnuSelectAll_Click()

txtList.SelStart = 0
txtList.SelLength = Len(txtList)

End Sub

Private Sub mnuUndo_Click()

SendKeys ("^Z")

End Sub

Private Sub txtList_Change()

NewBody = txtList

End Sub
