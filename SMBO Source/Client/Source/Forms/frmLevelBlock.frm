VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmLevelBlock 
   Caption         =   "Level Block"
   ClientHeight    =   1935
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3615
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   1935
   ScaleWidth      =   3615
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3345
      _ExtentX        =   5900
      _ExtentY        =   2990
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   353
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Level Block"
      TabPicture(0)   =   "frmLevelBlock.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Level"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "LevelNum"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "scrlLvlBlock"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Submit"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Cancel"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      Begin VB.CommandButton Cancel 
         Caption         =   "Cancel"
         Height          =   255
         Left            =   1920
         TabIndex        =   5
         Top             =   1250
         Width           =   1095
      End
      Begin VB.CommandButton Submit 
         Caption         =   "Ok"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1250
         Width           =   1095
      End
      Begin VB.HScrollBar scrlLvlBlock 
         Height          =   255
         Left            =   240
         Max             =   150
         Min             =   1
         TabIndex        =   1
         Top             =   700
         Value           =   1
         Width           =   2775
      End
      Begin VB.Label LevelNum 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   720
         TabIndex        =   3
         Top             =   460
         Width           =   495
      End
      Begin VB.Label Level 
         BackStyle       =   0  'Transparent
         Caption         =   "Level:"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   460
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmLevelBlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    If LevelNum.Caption = vbNullString Then
        LevelNum.Caption = "1"
    End If
End Sub

Private Sub scrlLvlBlock_Change()
    LevelNum.Caption = scrlLvlBlock.Value
End Sub

Private Sub Submit_Click()
    LevelToBlock = Int(scrlLvlBlock.Value)
    frmLevelBlock.Visible = False
End Sub
