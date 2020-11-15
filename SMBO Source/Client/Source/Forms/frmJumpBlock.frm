VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmJumpBlock 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Jump Block"
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4515
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   4515
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton cmdSubmit 
      Caption         =   "Submit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   3
      Top             =   2880
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2535
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   4471
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   6
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Info"
      TabPicture(0)   =   "frmJumpBlock.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblJumpBlockHeight"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblJumpBlockDecrease"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "chkUp"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "scrlHeight"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "chkDown"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "chkLeft"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "chkRight"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "chkAddUp"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "chkAddDown"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "chkAddLeft"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "chkAddRight"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "scrlDecrease"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).ControlCount=   13
      Begin VB.HScrollBar scrlDecrease 
         Height          =   255
         Left            =   2510
         Max             =   10
         TabIndex        =   15
         Top             =   720
         Width           =   1455
      End
      Begin VB.CheckBox chkAddRight 
         Caption         =   "Adds Height?"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2800
         TabIndex        =   13
         Top             =   2040
         Width           =   1215
      End
      Begin VB.CheckBox chkAddLeft 
         Caption         =   "Adds Height?"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   2040
         Width           =   1215
      End
      Begin VB.CheckBox chkAddDown 
         Caption         =   "Adds Height?"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2800
         TabIndex        =   11
         Top             =   1560
         Width           =   1215
      End
      Begin VB.CheckBox chkAddUp 
         Caption         =   "Adds Height?"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1560
         Width           =   1215
      End
      Begin VB.CheckBox chkRight 
         Caption         =   "Right"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2050
         TabIndex        =   9
         Top             =   2040
         Width           =   735
      End
      Begin VB.CheckBox chkLeft 
         Caption         =   "Left"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1450
         TabIndex        =   8
         Top             =   2040
         Width           =   615
      End
      Begin VB.CheckBox chkDown 
         Caption         =   "Down"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2050
         TabIndex        =   7
         Top             =   1560
         Width           =   735
      End
      Begin VB.HScrollBar scrlHeight 
         Height          =   255
         Left            =   240
         Max             =   10
         TabIndex        =   2
         Top             =   720
         Width           =   2055
      End
      Begin VB.CheckBox chkUp 
         Caption         =   "Up"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1450
         TabIndex        =   6
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label lblJumpBlockDecrease 
         Caption         =   "Decreases Height By: 0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2505
         TabIndex        =   14
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Direction able to jump from:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1200
         Width           =   3015
      End
      Begin VB.Label lblJumpBlockHeight 
         Caption         =   "Height Required to Jump Over: 0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   2175
      End
   End
End
Attribute VB_Name = "frmJumpBlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Dim i As Byte
    
    JumpHeight = 0
    JumpDecrease = 0
    
    For i = 1 To 4
        JumpDir(i) = 0
        JumpDirAddHeight(i) = 0
    Next i
    
    Unload Me
End Sub

Private Sub cmdSubmit_Click()
    JumpHeight = scrlHeight.Value
    JumpDecrease = scrlDecrease.Value
    JumpDir(1) = chkUp.Value
    JumpDir(2) = chkDown.Value
    JumpDir(3) = chkLeft.Value
    JumpDir(4) = chkRight.Value
    JumpDirAddHeight(1) = chkAddUp.Value
    JumpDirAddHeight(2) = chkAddDown.Value
    JumpDirAddHeight(3) = chkAddLeft.Value
    JumpDirAddHeight(4) = chkAddRight.Value
    Me.Visible = False
End Sub

Private Sub Form_Load()
    lblJumpBlockHeight.Caption = "Height Required to Jump Over: " & scrlHeight.Value
    lblJumpBlockDecrease.Caption = "Decreases Height By: " & scrlDecrease.Value
End Sub

Private Sub scrlDecrease_Change()
    lblJumpBlockDecrease.Caption = "Decreases Height By: " & scrlDecrease.Value
End Sub

Private Sub scrlHeight_Change()
    lblJumpBlockHeight.Caption = "Height Required to Jump Over: " & scrlHeight.Value
End Sub
