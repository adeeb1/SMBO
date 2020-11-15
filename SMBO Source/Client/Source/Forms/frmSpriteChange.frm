VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmSpriteChange 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sprite Change Attribute"
   ClientHeight    =   2610
   ClientLeft      =   60
   ClientTop       =   480
   ClientWidth     =   3885
   ControlBox      =   0   'False
   Icon            =   "frmSpriteChange.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   3885
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   4048
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
      TabCaption(0)   =   "Set Sprite"
      TabPicture(0)   =   "frmSpriteChange.frx":0FC2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblSprite"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "scrlSprite"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdCancel"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdOk"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "picSprite"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      Begin VB.PictureBox picSprite 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   2700
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   6
         Top             =   600
         Width           =   480
         Begin VB.PictureBox picSprites 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   480
            Left            =   0
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   7
            Top             =   0
            Width           =   480
         End
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
         TabIndex        =   3
         Top             =   1800
         Width           =   855
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
         Left            =   2160
         TabIndex        =   2
         Top             =   1800
         Width           =   855
      End
      Begin VB.HScrollBar scrlSprite 
         Height          =   255
         Left            =   360
         Max             =   500
         TabIndex        =   1
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Sprite:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   585
         TabIndex        =   5
         Top             =   720
         Width           =   405
      End
      Begin VB.Label lblSprite 
         AutoSize        =   -1  'True
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
         Height          =   165
         Left            =   1080
         TabIndex        =   4
         Top             =   720
         Width           =   75
      End
   End
End
Attribute VB_Name = "frmSpriteChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Me.Visible = False
End Sub

Private Sub cmdOk_Click()
    SpritePic = scrlSprite.Value
    scrlSprite.Value = 0
    Me.Visible = False
End Sub

Private Sub Form_Load()
    If SpritePic < scrlSprite.Min Then
        SpritePic = scrlSprite.Min
    End If
    scrlSprite.Value = SpritePic

    picSprites.Left = (3 * PIC_X) * -1
    picSprites.Top = (scrlSprite.Value * 64) * -1
    Call BitBlt(picSprite.hDC, 0, 0, PIC_X, 64, picSprites.hDC, 3 * PIC_X, scrlSprite.Value * 64, SRCCOPY)
End Sub

Private Sub scrlSprite_Change()
    lblSprite.Caption = scrlSprite.Value
    
    picSprites.Left = (3 * PIC_X) * -1
    picSprites.Top = (scrlSprite.Value * 64) * -1
    Call BitBlt(picSprite.hDC, 0, 0, PIC_X, 64, picSprites.hDC, 3 * PIC_X, scrlSprite.Value * 64, SRCCOPY)
End Sub

Private Sub Timer1_Timer()
    picSprites.Left = (3 * PIC_X) * -1
    picSprites.Top = (scrlSprite.Value * 64) * -1
    Call BitBlt(picSprite.hDC, 0, 0, PIC_X, 64, picSprites.hDC, 3 * PIC_X, scrlSprite.Value * 64, SRCCOPY)
End Sub
