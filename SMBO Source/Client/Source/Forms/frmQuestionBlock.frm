VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmQuestionBlock 
   Caption         =   "Question Block"
   ClientHeight    =   7815
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7290
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7815
   ScaleWidth      =   7290
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   7575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7065
      _ExtentX        =   12462
      _ExtentY        =   13361
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
      TabCaption(0)   =   "? Block"
      TabPicture(0)   =   "frmQuestionBlock.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Item1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Item2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Item1Chance"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label4"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Item2Chance"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label5"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Item3"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label6"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label7"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label8"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label9"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label10"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label11"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label12"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Item3Chance"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Item4Chance"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Item5Chance"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Item6Chance"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Item4"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Item5"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Item6"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Label13"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Value1"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Label14"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Value2"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Label16"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Value3"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Label15"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Value4"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Label17"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Value5"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "Label18"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "Value6"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "scrlItem1"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "scrlItem2"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "scrlChance1"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "scrlChance2"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "scrlItem3"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "scrlItem4"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "scrlItem5"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "scrlItem6"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "scrlChance3"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "scrlChance4"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "scrlChance5"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "scrlChance6"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "Ok"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "Cancel"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "scrlItem1Val"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "scrlItem2Val"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "scrlItem3Val"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "scrlItem4Val"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).Control(54)=   "scrlItem5Val"
      Tab(0).Control(54).Enabled=   0   'False
      Tab(0).Control(55)=   "scrlItem6Val"
      Tab(0).Control(55).Enabled=   0   'False
      Tab(0).ControlCount=   56
      Begin VB.HScrollBar scrlItem6Val 
         Height          =   255
         Left            =   3720
         Max             =   1000
         TabIndex        =   54
         Top             =   6000
         Width           =   3015
      End
      Begin VB.HScrollBar scrlItem5Val 
         Height          =   255
         Left            =   3720
         Max             =   1000
         TabIndex        =   51
         Top             =   4920
         Width           =   3015
      End
      Begin VB.HScrollBar scrlItem4Val 
         Height          =   255
         Left            =   3720
         Max             =   1000
         TabIndex        =   48
         Top             =   3840
         Width           =   3015
      End
      Begin VB.HScrollBar scrlItem3Val 
         Height          =   255
         Left            =   3720
         Max             =   1000
         TabIndex        =   45
         Top             =   2760
         Width           =   3015
      End
      Begin VB.HScrollBar scrlItem2Val 
         Height          =   255
         Left            =   3720
         Max             =   1000
         TabIndex        =   42
         Top             =   1680
         Width           =   3015
      End
      Begin VB.HScrollBar scrlItem1Val 
         Height          =   255
         Left            =   3720
         Max             =   1000
         TabIndex        =   39
         Top             =   600
         Width           =   3015
      End
      Begin VB.CommandButton Cancel 
         Caption         =   "Cancel"
         Height          =   435
         Left            =   4440
         TabIndex        =   35
         Top             =   6960
         Width           =   1695
      End
      Begin VB.CommandButton Ok 
         Caption         =   "Ok"
         Height          =   435
         Left            =   960
         TabIndex        =   34
         Top             =   6960
         Width           =   1695
      End
      Begin VB.HScrollBar scrlChance6 
         Height          =   255
         Left            =   3720
         Max             =   100
         TabIndex        =   25
         Top             =   6480
         Width           =   3015
      End
      Begin VB.HScrollBar scrlChance5 
         Height          =   255
         Left            =   3720
         Max             =   100
         TabIndex        =   24
         Top             =   5400
         Width           =   3015
      End
      Begin VB.HScrollBar scrlChance4 
         Height          =   255
         Left            =   3720
         Max             =   100
         TabIndex        =   23
         Top             =   4320
         Width           =   3015
      End
      Begin VB.HScrollBar scrlChance3 
         Height          =   255
         Left            =   3720
         Max             =   100
         TabIndex        =   22
         Top             =   3240
         Width           =   3015
      End
      Begin VB.HScrollBar scrlItem6 
         Height          =   255
         Left            =   240
         Max             =   1000
         TabIndex        =   18
         Top             =   6000
         Width           =   3015
      End
      Begin VB.HScrollBar scrlItem5 
         Height          =   255
         Left            =   240
         Max             =   1000
         TabIndex        =   17
         Top             =   4920
         Width           =   3015
      End
      Begin VB.HScrollBar scrlItem4 
         Height          =   255
         Left            =   240
         Max             =   1000
         TabIndex        =   16
         Top             =   3840
         Width           =   3015
      End
      Begin VB.HScrollBar scrlItem3 
         Height          =   255
         Left            =   240
         Max             =   1000
         TabIndex        =   13
         Top             =   2760
         Width           =   3015
      End
      Begin VB.HScrollBar scrlChance2 
         Height          =   255
         Left            =   3720
         Max             =   100
         TabIndex        =   10
         Top             =   2160
         Width           =   3015
      End
      Begin VB.HScrollBar scrlChance1 
         Height          =   255
         Left            =   3720
         Max             =   100
         TabIndex        =   7
         Top             =   1080
         Width           =   3015
      End
      Begin VB.HScrollBar scrlItem2 
         Height          =   255
         Left            =   240
         Max             =   1000
         TabIndex        =   4
         Top             =   1680
         Width           =   3015
      End
      Begin VB.HScrollBar scrlItem1 
         Height          =   255
         Left            =   240
         Max             =   1000
         TabIndex        =   2
         Top             =   600
         Width           =   3015
      End
      Begin VB.Label Value6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Value6"
         Height          =   255
         Left            =   4320
         TabIndex        =   56
         Top             =   5760
         Width           =   2415
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Value:"
         Height          =   255
         Left            =   3720
         TabIndex        =   55
         Top             =   5760
         Width           =   615
      End
      Begin VB.Label Value5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Value5"
         Height          =   255
         Left            =   4320
         TabIndex        =   53
         Top             =   4680
         Width           =   2415
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Value:"
         Height          =   255
         Left            =   3720
         TabIndex        =   52
         Top             =   4680
         Width           =   615
      End
      Begin VB.Label Value4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Value4"
         Height          =   255
         Left            =   4320
         TabIndex        =   50
         Top             =   3600
         Width           =   2415
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Value:"
         Height          =   255
         Left            =   3720
         TabIndex        =   49
         Top             =   3600
         Width           =   615
      End
      Begin VB.Label Value3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Value3"
         Height          =   255
         Left            =   4320
         TabIndex        =   47
         Top             =   2520
         Width           =   2415
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Value:"
         Height          =   255
         Left            =   3720
         TabIndex        =   46
         Top             =   2520
         Width           =   615
      End
      Begin VB.Label Value2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Value2"
         Height          =   255
         Left            =   4320
         TabIndex        =   44
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Value:"
         Height          =   255
         Left            =   3720
         TabIndex        =   43
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Value1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Value1"
         Height          =   255
         Left            =   4320
         TabIndex        =   41
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Value:"
         Height          =   255
         Left            =   3720
         TabIndex        =   40
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Item6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Item6"
         Height          =   255
         Left            =   840
         TabIndex        =   38
         Top             =   5760
         Width           =   2415
      End
      Begin VB.Label Item5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Item5"
         Height          =   255
         Left            =   840
         TabIndex        =   37
         Top             =   4680
         Width           =   2415
      End
      Begin VB.Label Item4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Item4"
         Height          =   255
         Left            =   840
         TabIndex        =   36
         Top             =   3600
         Width           =   2415
      End
      Begin VB.Label Item6Chance 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Item6 Chance"
         Height          =   255
         Left            =   5640
         TabIndex        =   33
         Top             =   6240
         Width           =   1095
      End
      Begin VB.Label Item5Chance 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Item5 Chance"
         Height          =   255
         Left            =   5640
         TabIndex        =   32
         Top             =   5160
         Width           =   1095
      End
      Begin VB.Label Item4Chance 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Item4 Chance"
         Height          =   255
         Left            =   5640
         TabIndex        =   31
         Top             =   4080
         Width           =   1095
      End
      Begin VB.Label Item3Chance 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Item3 Chance"
         Height          =   255
         Left            =   5640
         TabIndex        =   30
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Chance:"
         Height          =   255
         Left            =   3720
         TabIndex        =   29
         Top             =   6240
         Width           =   615
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Chance:"
         Height          =   255
         Left            =   3720
         TabIndex        =   28
         Top             =   5160
         Width           =   615
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Chance:"
         Height          =   255
         Left            =   3720
         TabIndex        =   27
         Top             =   4080
         Width           =   615
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Chance:"
         Height          =   255
         Left            =   3720
         TabIndex        =   26
         Top             =   3000
         Width           =   615
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Item 6:"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   5760
         Width           =   615
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Item 5:"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   4680
         Width           =   615
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Item 4:"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   3600
         Width           =   615
      End
      Begin VB.Label Item3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Item3"
         Height          =   255
         Left            =   840
         TabIndex        =   15
         Top             =   2520
         Width           =   2415
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Item 3:"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   2520
         Width           =   615
      End
      Begin VB.Label Item2Chance 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Item2 Chance"
         Height          =   255
         Left            =   5640
         TabIndex        =   12
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Chance:"
         Height          =   255
         Left            =   3720
         TabIndex        =   11
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label Item1Chance 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Item1 Chance"
         Height          =   255
         Left            =   5640
         TabIndex        =   9
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Chance:"
         Height          =   255
         Left            =   3720
         TabIndex        =   8
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Item2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Item2"
         Height          =   255
         Left            =   840
         TabIndex        =   6
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Item 2:"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Item1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Item1"
         Height          =   255
         Left            =   840
         TabIndex        =   3
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Item 1:"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmQuestionBlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cancel_Click()
    frmQuestionBlock.Visible = False
End Sub

Private Sub Form_Load()
  
  scrlItem1.Max = MAX_ITEMS
  scrlItem2.Max = MAX_ITEMS
  scrlItem3.Max = MAX_ITEMS
  scrlItem4.Max = MAX_ITEMS
  scrlItem5.Max = MAX_ITEMS
  scrlItem6.Max = MAX_ITEMS
  
  If Int(scrlItem1.Value) < 1 Then
    Item1.Caption = "None"
  Else
    Item1.Caption = Trim$(Item(Int(scrlItem1.Value)).Name)
  End If
  If Int(scrlItem2.Value) < 1 Then
    Item2.Caption = "None"
  Else
    Item2.Caption = Trim$(Item(Int(scrlItem2.Value)).Name)
  End If
  If Int(scrlItem3.Value) < 1 Then
    Item3.Caption = "None"
  Else
    Item3.Caption = Trim$(Item(Int(scrlItem3.Value)).Name)
  End If
  If Int(scrlItem4.Value) < 1 Then
    Item4.Caption = "None"
  Else
    Item4.Caption = Trim$(Item(Int(scrlItem4.Value)).Name)
  End If
  If Int(scrlItem5.Value) < 1 Then
    Item5.Caption = "None"
  Else
    Item5.Caption = Trim$(Item(Int(scrlItem5.Value)).Name)
  End If
  If Int(scrlItem6.Value) < 1 Then
    Item6.Caption = "None"
  Else
    Item6.Caption = Trim$(Item(Int(scrlItem6.Value)).Name)
  End If
  
    Item1Chance.Caption = Int(scrlChance1.Value) & " out of 100"
    Item2Chance.Caption = Int(scrlChance2.Value) & " out of 100"
    Item3Chance.Caption = Int(scrlChance3.Value) & " out of 100"
    Item4Chance.Caption = Int(scrlChance4.Value) & " out of 100"
    Item5Chance.Caption = Int(scrlChance5.Value) & " out of 100"
    Item6Chance.Caption = Int(scrlChance6.Value) & " out of 100"
    
    Value1.Caption = Int(scrlItem1Val.Value)
    Value2.Caption = Int(scrlItem2Val.Value)
    Value3.Caption = Int(scrlItem3Val.Value)
    Value4.Caption = Int(scrlItem4Val.Value)
    Value5.Caption = Int(scrlItem5Val.Value)
    Value6.Caption = Int(scrlItem6Val.Value)

End Sub

Private Sub Ok_Click()

    ItemThing1 = Int(scrlItem1.Value)
    ItemThing2 = Int(scrlItem2.Value)
    ItemThing3 = Int(scrlItem3.Value)
    ItemThing4 = Int(scrlItem4.Value)
    ItemThing5 = Int(scrlItem5.Value)
    ItemThing6 = Int(scrlItem6.Value)
    ChanceThing1 = Int(scrlChance1.Value)
    ChanceThing2 = Int(scrlChance2.Value)
    ChanceThing3 = Int(scrlChance3.Value)
    ChanceThing4 = Int(scrlChance4.Value)
    ChanceThing5 = Int(scrlChance5.Value)
    ChanceThing6 = Int(scrlChance6.Value)
    ValueThing1 = Int(scrlItem1Val.Value)
    ValueThing2 = Int(scrlItem2Val.Value)
    ValueThing3 = Int(scrlItem3Val.Value)
    ValueThing4 = Int(scrlItem4Val.Value)
    ValueThing5 = Int(scrlItem5Val.Value)
    ValueThing6 = Int(scrlItem6Val.Value)
    
    frmQuestionBlock.Visible = False
End Sub

Private Sub scrlItem1_Change()
  If Int(scrlItem1.Value) < 1 Then
    Item1.Caption = "None"
  Else
    Item1.Caption = Trim$(Item(Int(scrlItem1.Value)).Name)
  End If
End Sub

Private Sub scrlItem2_Change()
  If Int(scrlItem2.Value) < 1 Then
    Item2.Caption = "None"
  Else
    Item2.Caption = Trim$(Item(Int(scrlItem2.Value)).Name)
  End If
End Sub

Private Sub scrlItem3_Change()
  If Int(scrlItem3.Value) < 1 Then
    Item3.Caption = "None"
  Else
    Item3.Caption = Trim$(Item(Int(scrlItem3.Value)).Name)
  End If
End Sub

Private Sub scrlItem4_Change()
  If Int(scrlItem4.Value) < 1 Then
    Item4.Caption = "None"
  Else
    Item4.Caption = Trim$(Item(Int(scrlItem4.Value)).Name)
  End If
End Sub

Private Sub scrlItem5_Change()
  If Int(scrlItem5.Value) < 1 Then
    Item5.Caption = "None"
  Else
    Item5.Caption = Trim$(Item(Int(scrlItem5.Value)).Name)
  End If
End Sub

Private Sub scrlItem6_Change()
  If Int(scrlItem6.Value) < 1 Then
    Item6.Caption = "None"
  Else
    Item6.Caption = Trim$(Item(Int(scrlItem6.Value)).Name)
  End If
End Sub

Private Sub scrlChance1_Change()
    Item1Chance.Caption = scrlChance1.Value & " out of 100"
End Sub

Private Sub scrlChance2_Change()
    Item2Chance.Caption = scrlChance2.Value & " out of 100"
End Sub

Private Sub scrlChance3_Change()
    Item3Chance.Caption = scrlChance3.Value & " out of 100"
End Sub

Private Sub scrlChance4_Change()
    Item4Chance.Caption = scrlChance4.Value & " out of 100"
End Sub

Private Sub scrlChance5_Change()
    Item5Chance.Caption = scrlChance5.Value & " out of 100"
End Sub

Private Sub scrlChance6_Change()
    Item6Chance.Caption = scrlChance6.Value & " out of 100"
End Sub

Private Sub scrlItem1Val_Change()
    Value1.Caption = Int(scrlItem1Val.Value)
End Sub

Private Sub scrlItem2Val_Change()
    Value2.Caption = Int(scrlItem2Val.Value)
End Sub

Private Sub scrlItem3Val_Change()
    Value3.Caption = Int(scrlItem3Val.Value)
End Sub

Private Sub scrlItem4Val_Change()
    Value4.Caption = Int(scrlItem4Val.Value)
End Sub

Private Sub scrlItem5Val_Change()
    Value5.Caption = Int(scrlItem5Val.Value)
End Sub

Private Sub scrlItem6Val_Change()
    Value6.Caption = Int(scrlItem6Val.Value)
End Sub
