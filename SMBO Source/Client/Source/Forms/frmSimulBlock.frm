VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmSimulBlock 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Simultaneous Blocks"
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4095
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   4095
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
      Left            =   2400
      TabIndex        =   2
      Top             =   4200
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
      Left            =   600
      TabIndex        =   1
      Top             =   4200
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3855
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3690
      _ExtentX        =   6509
      _ExtentY        =   6800
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
      TabPicture(0)   =   "frmSimulBlock.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblSimulBlock(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblXCoor(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblXCoorValue(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblYCoor(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblYCoorValue(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblYCoorValue(1)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblYCoor(1)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblXCoorValue(1)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblXCoor(1)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblSimulBlock(1)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lblYCoorValue(2)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lblYCoor(2)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "lblXCoorValue(2)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "lblXCoor(2)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "lblSimulBlock(2)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "lblYCoorValue(3)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "lblYCoor(3)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "lblXCoorValue(3)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "lblXCoor(3)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "lblSimulBlock(3)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "lblWarpTo"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "lblYWarpCoorValue"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "lblYWarpCoor"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "lblXWarpCoorValue"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "lblXWarpCoor"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "scrlXCoor(0)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "scrlYCoor(0)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "scrlYCoor(1)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "scrlXCoor(1)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "scrlYCoor(2)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "scrlXCoor(2)"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "scrlYCoor(3)"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "scrlXCoor(3)"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "scrlYWarpCoor"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "scrlXWarpCoor"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).ControlCount=   35
      Begin VB.HScrollBar scrlXWarpCoor 
         Height          =   255
         Left            =   480
         Max             =   30
         TabIndex        =   33
         Top             =   3360
         Width           =   735
      End
      Begin VB.HScrollBar scrlYWarpCoor 
         Height          =   255
         Left            =   2160
         Max             =   30
         TabIndex        =   32
         Top             =   3360
         Width           =   735
      End
      Begin VB.HScrollBar scrlXCoor 
         Height          =   255
         Index           =   3
         Left            =   480
         Max             =   30
         TabIndex        =   25
         Top             =   2640
         Width           =   735
      End
      Begin VB.HScrollBar scrlYCoor 
         Height          =   255
         Index           =   3
         Left            =   2160
         Max             =   30
         TabIndex        =   24
         Top             =   2640
         Width           =   735
      End
      Begin VB.HScrollBar scrlXCoor 
         Height          =   255
         Index           =   2
         Left            =   480
         Max             =   30
         TabIndex        =   18
         Top             =   2040
         Width           =   735
      End
      Begin VB.HScrollBar scrlYCoor 
         Height          =   255
         Index           =   2
         Left            =   2160
         Max             =   30
         TabIndex        =   17
         Top             =   2040
         Width           =   735
      End
      Begin VB.HScrollBar scrlXCoor 
         Height          =   255
         Index           =   1
         Left            =   480
         Max             =   30
         TabIndex        =   11
         Top             =   1440
         Width           =   735
      End
      Begin VB.HScrollBar scrlYCoor 
         Height          =   255
         Index           =   1
         Left            =   2160
         Max             =   30
         TabIndex        =   10
         Top             =   1440
         Width           =   735
      End
      Begin VB.HScrollBar scrlYCoor 
         Height          =   255
         Index           =   0
         Left            =   2160
         Max             =   30
         TabIndex        =   8
         Top             =   840
         Width           =   735
      End
      Begin VB.HScrollBar scrlXCoor 
         Height          =   255
         Index           =   0
         Left            =   480
         Max             =   30
         TabIndex        =   5
         Top             =   840
         Width           =   735
      End
      Begin VB.Label lblXWarpCoor 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Left            =   240
         TabIndex        =   37
         Top             =   3405
         Width           =   135
      End
      Begin VB.Label lblXWarpCoorValue 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Left            =   1320
         TabIndex        =   36
         Top             =   3405
         Width           =   150
      End
      Begin VB.Label lblYWarpCoor 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Y"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Left            =   1920
         TabIndex        =   35
         Top             =   3405
         Width           =   135
      End
      Begin VB.Label lblYWarpCoorValue 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Left            =   3000
         TabIndex        =   34
         Top             =   3405
         Width           =   150
      End
      Begin VB.Label lblWarpTo 
         BackStyle       =   0  'Transparent
         Caption         =   "Warp Players To:"
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
         TabIndex        =   31
         Top             =   3120
         Width           =   1095
      End
      Begin VB.Label lblSimulBlock 
         BackStyle       =   0  'Transparent
         Caption         =   "Simultaneous Block #4"
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
         Index           =   3
         Left            =   240
         TabIndex        =   30
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Label lblXCoor 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   3
         Left            =   240
         TabIndex        =   29
         Top             =   2685
         Width           =   135
      End
      Begin VB.Label lblXCoorValue 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   3
         Left            =   1320
         TabIndex        =   28
         Top             =   2685
         Width           =   150
      End
      Begin VB.Label lblYCoor 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Y"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   3
         Left            =   1920
         TabIndex        =   27
         Top             =   2685
         Width           =   135
      End
      Begin VB.Label lblYCoorValue 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   3
         Left            =   3000
         TabIndex        =   26
         Top             =   2685
         Width           =   150
      End
      Begin VB.Label lblSimulBlock 
         BackStyle       =   0  'Transparent
         Caption         =   "Simultaneous Block #3"
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
         Index           =   2
         Left            =   240
         TabIndex        =   23
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label lblXCoor 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   2
         Left            =   240
         TabIndex        =   22
         Top             =   2085
         Width           =   135
      End
      Begin VB.Label lblXCoorValue 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   2
         Left            =   1320
         TabIndex        =   21
         Top             =   2085
         Width           =   150
      End
      Begin VB.Label lblYCoor 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Y"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   2
         Left            =   1920
         TabIndex        =   20
         Top             =   2085
         Width           =   135
      End
      Begin VB.Label lblYCoorValue 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   2
         Left            =   3000
         TabIndex        =   19
         Top             =   2085
         Width           =   150
      End
      Begin VB.Label lblSimulBlock 
         BackStyle       =   0  'Transparent
         Caption         =   "Simultaneous Block #2"
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
         Index           =   1
         Left            =   240
         TabIndex        =   16
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label lblXCoor 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   1
         Left            =   240
         TabIndex        =   15
         Top             =   1485
         Width           =   135
      End
      Begin VB.Label lblXCoorValue 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   1
         Left            =   1320
         TabIndex        =   14
         Top             =   1485
         Width           =   150
      End
      Begin VB.Label lblYCoor 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Y"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   1
         Left            =   1920
         TabIndex        =   13
         Top             =   1485
         Width           =   135
      End
      Begin VB.Label lblYCoorValue 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   1
         Left            =   3000
         TabIndex        =   12
         Top             =   1485
         Width           =   150
      End
      Begin VB.Label lblYCoorValue 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   0
         Left            =   3000
         TabIndex        =   9
         Top             =   885
         Width           =   150
      End
      Begin VB.Label lblYCoor 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Y"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   0
         Left            =   1920
         TabIndex        =   7
         Top             =   885
         Width           =   135
      End
      Begin VB.Label lblXCoorValue 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   0
         Left            =   1320
         TabIndex        =   6
         Top             =   885
         Width           =   150
      End
      Begin VB.Label lblXCoor 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   890
         Width           =   135
      End
      Begin VB.Label lblSimulBlock 
         BackStyle       =   0  'Transparent
         Caption         =   "Simultaneous Block #1"
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
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmSimulBlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Dim i As Byte
    
    For i = 1 To 4
        SimulBlockCoords(i) = vbNullString
    Next i
    
    SimulBlockWarpCoords(1) = 0
    SimulBlockWarpCoords(2) = 0
    
    Me.Visible = False
End Sub

Private Sub cmdSubmit_Click()
    Dim i As Byte, Index As Byte
    
    For i = 1 To 4
        Index = (i - 1)
        
        SimulBlockCoords(i) = scrlXCoor(Index).Value & "," & scrlYCoor(Index).Value
    Next i
    
    SimulBlockWarpCoords(1) = scrlXWarpCoor.Value
    SimulBlockWarpCoords(2) = scrlYWarpCoor.Value
    
    Me.Visible = False
End Sub

Private Sub Form_Load()
    Dim i As Byte
    
    For i = 1 To 4
        SimulBlockCoords(i) = vbNullString
    Next i
    
    SimulBlockWarpCoords(1) = 0
    SimulBlockWarpCoords(2) = 0
End Sub

Private Sub scrlXCoor_Change(Index As Integer)
    lblXCoorValue(Index).Caption = scrlXCoor(Index).Value
End Sub

Private Sub scrlXWarpCoor_Change()
    lblXWarpCoorValue.Caption = scrlXWarpCoor.Value
End Sub

Private Sub scrlYCoor_Change(Index As Integer)
    lblYCoorValue(Index).Caption = scrlYCoor(Index).Value
End Sub

Private Sub scrlYWarpCoor_Change()
    lblYWarpCoorValue.Caption = scrlYWarpCoor.Value
End Sub
