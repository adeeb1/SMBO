VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmBeanTile 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bean Tile"
   ClientHeight    =   2700
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4125
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   4125
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
      Left            =   600
      TabIndex        =   6
      Top             =   2160
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
      TabIndex        =   5
      Top             =   2160
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   1815
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3690
      _ExtentX        =   6509
      _ExtentY        =   3201
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
      TabPicture(0)   =   "frmBeanTile.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblQuantity"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblItem"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "scrlQuantity"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "scrlItem"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      Begin VB.HScrollBar scrlItem 
         Height          =   255
         Left            =   240
         Max             =   1000
         TabIndex        =   2
         Top             =   720
         Width           =   3135
      End
      Begin VB.HScrollBar scrlQuantity 
         Height          =   255
         Left            =   240
         Max             =   1000
         TabIndex        =   1
         Top             =   1320
         Width           =   3135
      End
      Begin VB.Label lblItem 
         BackStyle       =   0  'Transparent
         Caption         =   "Item:"
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
         TabIndex        =   4
         Top             =   480
         Width           =   3135
      End
      Begin VB.Label lblQuantity 
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity:"
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
         TabIndex        =   3
         Top             =   1080
         Width           =   3135
      End
   End
End
Attribute VB_Name = "frmBeanTile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub cmdSubmit_Click()
    BeanItemNum = scrlItem.Value
    BeanItemQuantity = scrlQuantity.Value
    
    Me.Hide
End Sub

Private Sub Form_Load()
    scrlItem.Max = MAX_ITEMS
End Sub

Private Sub scrlItem_Change()
    lblItem.Caption = "Item: " & Trim$(Item(scrlItem.Value).Name) & " (" & scrlItem.Value & ")"
End Sub

Private Sub scrlQuantity_Change()
    lblQuantity.Caption = "Quantity: " & scrlQuantity.Value
End Sub
