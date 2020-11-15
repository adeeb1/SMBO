VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmSwitch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Switch Attribute"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5085
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   5085
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   5530
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   353
      TabMaxWidth     =   1764
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Switch"
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblX"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblY"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdOk"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdCancel"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtMap"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "scrlX"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "scrlY"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      Begin VB.HScrollBar scrlY 
         Height          =   255
         Left            =   720
         Max             =   29
         TabIndex        =   9
         Top             =   1395
         Width           =   3375
      End
      Begin VB.HScrollBar scrlX 
         Height          =   255
         Left            =   720
         Max             =   29
         TabIndex        =   8
         Top             =   1050
         Width           =   3375
      End
      Begin VB.TextBox txtMap 
         Height          =   285
         Left            =   960
         TabIndex        =   5
         Text            =   "0"
         Top             =   720
         Width           =   3735
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
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
         Left            =   2640
         TabIndex        =   2
         Top             =   2760
         Width           =   1935
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "Ok"
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
         Left            =   480
         TabIndex        =   1
         Top             =   2760
         Width           =   1935
      End
      Begin VB.Frame Frame1 
         Caption         =   "Switch can be activated with:"
         Height          =   735
         Left            =   480
         TabIndex        =   10
         Top             =   1920
         Width           =   3975
         Begin VB.CheckBox chkMelee 
            Caption         =   "Melee Weapon"
            Height          =   195
            Left            =   240
            TabIndex        =   12
            Top             =   360
            Width           =   1575
         End
         Begin VB.CheckBox chkRanged 
            Caption         =   "Ranged Weapon"
            Height          =   195
            Left            =   2160
            TabIndex        =   11
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.Label lblY 
         Alignment       =   2  'Center
         Caption         =   "30"
         Height          =   255
         Left            =   4080
         TabIndex        =   14
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label lblX 
         Alignment       =   2  'Center
         Caption         =   "30"
         Height          =   255
         Left            =   4080
         TabIndex        =   13
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Y"
         Height          =   195
         Left            =   480
         TabIndex        =   7
         Top             =   1440
         Width           =   105
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "X"
         Height          =   195
         Left            =   480
         TabIndex        =   6
         Top             =   1095
         Width           =   105
      End
      Begin VB.Label Label2 
         Caption         =   "Map:"
         Height          =   255
         Left            =   480
         TabIndex        =   4
         Top             =   750
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Warp player to ..."
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
         Left            =   360
         TabIndex        =   3
         Top             =   480
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmSwitch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' This will be tricky, I need to store 5 values with only 3 variables.

Private Sub cmdCancel_Click()
    frmSwitch.Visible = False
End Sub

Private Sub cmdOk_Click()
    SwitchWarpMap = Val(txtMap.Text)
    SwitchWarpPos = scrlX.Value * 256 + scrlY.Value
    SwitchWarpFlags = (chkMelee.Value + chkRanged.Value * 2)
    frmSwitch.Visible = False
End Sub

Private Sub Form_Load()
    scrlX.Max = MAX_MAPX
    scrlY.Max = MAX_MAPY
    
    txtMap.Text = SwitchWarpMap
    scrlX.Value = SwitchWarpPos \ 256               ' Kinda advanced math to store the X and Y in one variable.
    scrlY.Value = SwitchWarpPos Mod 256
    chkMelee.Value = (SwitchWarpFlags And 1)        ' Bit Flags, Ask me if you need help with these. They're kinda advanced.
    chkMelee.Value = (SwitchWarpFlags And 2) \ 2
End Sub

Private Sub scrlX_Change()
    lblX.Caption = scrlX.Value
End Sub

Private Sub scrlY_Change()
    lblY.Caption = scrlY.Value
End Sub
