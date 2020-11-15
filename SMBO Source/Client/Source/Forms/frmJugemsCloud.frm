VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmJugemsCloud 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Jugem's Cloud"
   ClientHeight    =   2355
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3975
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   3975
   StartUpPosition =   3  'Windows Default
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
      Left            =   360
      TabIndex        =   5
      Top             =   1800
      Width           =   1215
   End
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
      Left            =   2400
      TabIndex        =   4
      Top             =   1800
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3690
      _ExtentX        =   6509
      _ExtentY        =   2566
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
      TabPicture(0)   =   "frmJugemsCloud.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "chkLeft"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "chkRight"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "chkUp"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "chkDown"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      Begin VB.CheckBox chkDown 
         Caption         =   "Down"
         Height          =   255
         Left            =   2520
         TabIndex        =   7
         Top             =   960
         Width           =   855
      End
      Begin VB.CheckBox chkUp 
         Caption         =   "Up"
         Height          =   255
         Left            =   1680
         TabIndex        =   6
         Top             =   960
         Width           =   615
      End
      Begin VB.CheckBox chkRight 
         Caption         =   "Right"
         Height          =   255
         Left            =   2520
         TabIndex        =   3
         Top             =   600
         Width           =   855
      End
      Begin VB.CheckBox chkLeft 
         Caption         =   "Left"
         Height          =   255
         Left            =   1680
         TabIndex        =   2
         Top             =   600
         Value           =   1  'Checked
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Direction to face:"
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
         TabIndex        =   1
         Top             =   600
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmJugemsCloud"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkLeft_Click()
    If chkLeft.Value = Checked Then
        chkRight.Value = Unchecked
        chkUp.Value = Unchecked
        chkDown.Value = Unchecked
    End If
End Sub

Private Sub chkRight_Click()
    If chkRight.Value = Checked Then
        chkLeft.Value = Unchecked
        chkUp.Value = Unchecked
        chkDown.Value = Unchecked
    End If
End Sub

Private Sub chkUp_Click()
    If chkUp.Value = Checked Then
        chkLeft.Value = Unchecked
        chkRight.Value = Unchecked
        chkDown.Value = Unchecked
    End If
End Sub

Private Sub chkDown_Click()
    If chkDown.Value = Checked Then
        chkLeft.Value = Unchecked
        chkRight.Value = Unchecked
        chkUp.Value = Unchecked
    End If
End Sub

Private Sub cmdSubmit_Click()
    Select Case Checked
        Case chkLeft.Value
            CloudDir = DIR_LEFT
        Case chkRight.Value
            CloudDir = DIR_RIGHT
        Case chkUp.Value
            CloudDir = DIR_UP
        Case chkDown.Value
            CloudDir = DIR_DOWN
        Case Else
            MsgBox "You must choose a direction to face!", vbOKOnly
            Exit Sub
    End Select
    
    Me.Visible = False
End Sub

Private Sub cmdCancel_Click()
    Me.Visible = False
End Sub
