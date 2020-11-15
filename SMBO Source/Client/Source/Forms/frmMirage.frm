VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{96366485-4AD2-4BC8-AFBF-B1FC132616A5}#2.0#0"; "VBMP.ocx"
Begin VB.Form frmMirage 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Eclipse Evolution"
   ClientHeight    =   8865
   ClientLeft      =   555
   ClientTop       =   780
   ClientWidth     =   12000
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   FontTransparent =   0   'False
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmMirage.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmMirage.frx":08CA
   ScaleHeight     =   591
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   Visible         =   0   'False
   Begin VB.PictureBox itmDesc 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      FillColor       =   &H00828B82&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5775
      Left            =   9360
      ScaleHeight     =   383
      ScaleMode       =   0  'User
      ScaleWidth      =   175
      TabIndex        =   35
      Top             =   420
      Visible         =   0   'False
      Width           =   2655
      Begin VB.Label descCritBlock 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Crit: XXXX Dodge: XXXX"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   0
         TabIndex        =   131
         Top             =   3600
         Width           =   2655
      End
      Begin VB.Label descHP 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "HP"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   0
         TabIndex        =   106
         Top             =   720
         Width           =   2655
      End
      Begin VB.Label descFP 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "FP"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   0
         TabIndex        =   105
         Top             =   960
         Width           =   2655
      End
      Begin VB.Label descLevel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Level"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   0
         TabIndex        =   104
         Top             =   2400
         Width           =   2655
      End
      Begin VB.Label descClass 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Class"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   0
         TabIndex        =   103
         Top             =   2160
         Width           =   2655
      End
      Begin VB.Label descMagic 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Stache"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   0
         TabIndex        =   81
         Top             =   1680
         Width           =   2655
      End
      Begin VB.Label descName 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   570
         Left            =   0
         TabIndex        =   46
         Top             =   0
         Width           =   2655
      End
      Begin VB.Label lblRequirements 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Requirements"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   0
         TabIndex        =   45
         Top             =   480
         Width           =   2655
      End
      Begin VB.Label descStr 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Strength"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   0
         TabIndex        =   44
         Top             =   1200
         Width           =   2655
      End
      Begin VB.Label descDef 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Defense"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   0
         TabIndex        =   43
         Top             =   1440
         Width           =   2655
      End
      Begin VB.Label descSpeed 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Speed"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   0
         TabIndex        =   42
         Top             =   1920
         Width           =   2655
      End
      Begin VB.Label lblAdditions 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Equipment Stats"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   0
         TabIndex        =   41
         Top             =   2640
         Width           =   2655
      End
      Begin VB.Label descHpMp 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "HP: XXXX FP: XXXX SP: XXXX"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   0
         TabIndex        =   40
         Top             =   2880
         Width           =   2655
      End
      Begin VB.Label descSD 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Str: XXXX Def: XXXXX"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   120
         TabIndex        =   39
         Top             =   3120
         Width           =   2655
      End
      Begin VB.Label desc 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   1215
         Left            =   120
         TabIndex        =   38
         Top             =   4320
         Width           =   2415
      End
      Begin VB.Label lblDescription 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   0
         TabIndex        =   37
         Top             =   3960
         Width           =   2655
      End
      Begin VB.Label descMS 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Speed: XXXX Stache: XXXX"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   0
         TabIndex        =   36
         Top             =   3360
         Width           =   2655
      End
   End
   Begin VB.PictureBox picInventory 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3375
      Left            =   9570
      ScaleHeight     =   225
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   160
      TabIndex        =   141
      Top             =   5460
      Visible         =   0   'False
      Width           =   2400
      Begin VB.PictureBox picDragBox 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   0
         ScaleHeight     =   450
         ScaleWidth      =   450
         TabIndex        =   148
         Top             =   0
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.PictureBox picUp 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   840
         Picture         =   "frmMirage.frx":16020C
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   142
         Top             =   3000
         Width           =   270
      End
      Begin VB.PictureBox picDown 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1230
         Picture         =   "frmMirage.frx":1604A4
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   143
         Top             =   3000
         Width           =   270
      End
      Begin VB.PictureBox picInventory3 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2655
         Left            =   15
         ScaleHeight     =   177
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   158
         TabIndex        =   147
         Top             =   30
         Width           =   2370
      End
      Begin VB.VScrollBar scrlInventory 
         Height          =   330
         Left            =   2640
         Max             =   3
         TabIndex        =   144
         Top             =   2400
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblUseItem 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Use Item"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   0
         TabIndex        =   146
         Top             =   3000
         Width           =   690
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   40
         X2              =   128
         Y1              =   192
         Y2              =   192
      End
      Begin VB.Label lblDropItem 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Drop Item"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   1560
         TabIndex        =   145
         Top             =   3000
         Width           =   795
      End
   End
   Begin RichTextLib.RichTextBox txtChat 
      Height          =   1260
      Left            =   30
      TabIndex        =   0
      Top             =   7560
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   2223
      _Version        =   393217
      BackColor       =   0
      Enabled         =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmMirage.frx":16072F
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picOptions 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5610
      Left            =   6360
      Picture         =   "frmMirage.frx":1607AC
      ScaleHeight     =   372
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   175
      TabIndex        =   47
      Top             =   600
      Visible         =   0   'False
      Width           =   2655
      Begin VB.CheckBox chkAutoRun 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Enable Auto-Run"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   120
         TabIndex        =   140
         Top             =   1080
         Value           =   1  'Checked
         Width           =   2325
      End
      Begin VB.CheckBox chkTurnBased 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Turn-based Battles"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   120
         TabIndex        =   122
         Top             =   1320
         Value           =   1  'Checked
         Width           =   2325
      End
      Begin VB.CheckBox chkPlayerBar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Mini HP Bar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   120
         TabIndex        =   59
         Top             =   840
         Value           =   1  'Checked
         Width           =   2325
      End
      Begin VB.CheckBox chkPlayerName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Names"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   120
         TabIndex        =   58
         Top             =   360
         Value           =   1  'Checked
         Width           =   2325
      End
      Begin VB.CheckBox chkNpcName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Names"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   120
         TabIndex        =   57
         Top             =   1905
         Value           =   1  'Checked
         Width           =   2325
      End
      Begin VB.CheckBox chkBubbleBar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Display Map Text"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   120
         TabIndex        =   56
         Top             =   3825
         Value           =   1  'Checked
         Width           =   2325
      End
      Begin VB.CheckBox chkNpcBar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "NPC HP Bars"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   120
         TabIndex        =   55
         Top             =   2385
         Value           =   1  'Checked
         Width           =   2325
      End
      Begin VB.CheckBox chkPlayerDamage 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Damage Above Head"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   120
         TabIndex        =   54
         Top             =   600
         Value           =   1  'Checked
         Width           =   2325
      End
      Begin VB.CheckBox chkNpcDamage 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Damage Above Heads"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   120
         TabIndex        =   53
         Top             =   2145
         Value           =   1  'Checked
         Width           =   2325
      End
      Begin VB.CheckBox chkMusic 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Music"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   120
         TabIndex        =   52
         Top             =   2985
         Value           =   1  'Checked
         Width           =   2325
      End
      Begin VB.CheckBox chkSound 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Sound"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   120
         TabIndex        =   51
         Top             =   3225
         Value           =   1  'Checked
         Width           =   2325
      End
      Begin VB.HScrollBar scrlBltText 
         Height          =   255
         Left            =   120
         Max             =   20
         Min             =   4
         TabIndex        =   50
         Top             =   4665
         Value           =   6
         Width           =   2295
      End
      Begin VB.CommandButton cmdSaveConfig 
         Caption         =   "Save Settings"
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
         Left            =   120
         TabIndex        =   49
         Top             =   5145
         Width           =   2325
      End
      Begin VB.CheckBox chkAutoScroll 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Auto Scroll"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   120
         TabIndex        =   48
         Top             =   4065
         Value           =   1  'Checked
         Width           =   2325
      End
      Begin VB.Label lblNPCData 
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "NPC Data"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   120
         TabIndex        =   82
         Top             =   1665
         Width           =   855
      End
      Begin VB.Label lblLines 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "On Screen Text Line Amount: 6"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Left            =   120
         TabIndex        =   63
         Top             =   4425
         Width           =   2295
      End
      Begin VB.Label lblPlayerData 
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Player Data"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   120
         TabIndex        =   62
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label lblSoundData 
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Sound Data"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   120
         TabIndex        =   61
         Top             =   2745
         Width           =   975
      End
      Begin VB.Label lblChatData 
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Chat Data"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   120
         TabIndex        =   60
         Top             =   3585
         Width           =   855
      End
   End
   Begin VB.PictureBox picItems 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   9480
      ScaleHeight     =   477.09
      ScaleMode       =   0  'User
      ScaleWidth      =   477.091
      TabIndex        =   33
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Timer tmrSnowDrop 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   10920
      Top             =   120
   End
   Begin VB.Timer tmrRainDrop 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   10440
      Top             =   120
   End
   Begin VB.PictureBox ScreenShot 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   10920
      ScaleHeight     =   495
      ScaleWidth      =   525
      TabIndex        =   24
      Top             =   240
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.PictureBox picPlayerSpells 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00789298&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3375
      Left            =   9570
      ScaleHeight     =   225
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   160
      TabIndex        =   1
      Top             =   5460
      Visible         =   0   'False
      Width           =   2400
      Begin VB.ListBox lstSpells 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   2835
         ItemData        =   "frmMirage.frx":191126
         Left            =   45
         List            =   "frmMirage.frx":191128
         TabIndex        =   2
         Top             =   60
         Width           =   2310
      End
      Begin VB.Label lblDetails 
         BackStyle       =   0  'Transparent
         Caption         =   "Info"
         BeginProperty Font 
            Name            =   "Porky's"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1800
         TabIndex        =   123
         Top             =   3000
         Width           =   375
      End
      Begin VB.Label lblForgetSpell 
         BackStyle       =   0  'Transparent
         Caption         =   "Forget"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   840
         TabIndex        =   31
         Top             =   3120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label lblCast 
         BackStyle       =   0  'Transparent
         Caption         =   "Use"
         BeginProperty Font 
            Name            =   "Porky's"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   240
         TabIndex        =   3
         Top             =   3000
         Width           =   375
      End
   End
   Begin VB.PictureBox picWhosOnline 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00789298&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3375
      Left            =   9570
      ScaleHeight     =   225
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   160
      TabIndex        =   8
      Top             =   5460
      Visible         =   0   'False
      Width           =   2400
      Begin VB.ListBox lstOnline 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2505
         ItemData        =   "frmMirage.frx":19112A
         Left            =   45
         List            =   "frmMirage.frx":19112C
         TabIndex        =   9
         Top             =   60
         Width           =   2310
      End
      Begin VB.Label lblFriendList 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Friends List"
         BeginProperty Font 
            Name            =   "Porky's"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   93
         Top             =   3000
         Width           =   2175
      End
      Begin VB.Label PrivateMsg 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Private Message"
         BeginProperty Font 
            Name            =   "Porky's"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   92
         Top             =   2640
         Width           =   2175
      End
   End
   Begin VB.PictureBox picGuildAdmin 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00789298&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3375
      Left            =   9570
      ScaleHeight     =   3375
      ScaleWidth      =   2400
      TabIndex        =   11
      Top             =   5460
      Visible         =   0   'False
      Width           =   2400
      Begin VB.TextBox txtAccess 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   750
         TabIndex        =   13
         Top             =   585
         Width           =   1575
      End
      Begin VB.TextBox txtName 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   750
         TabIndex        =   12
         Top             =   345
         Width           =   1575
      End
      Begin VB.Label cmdTrainee 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Recruit Player"
         BeginProperty Font 
            Name            =   "Porky's"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   300
         TabIndex        =   121
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label cmdMember 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Recruit as a higher rank"
         BeginProperty Font 
            Name            =   "Porky's"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   300
         TabIndex        =   120
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label cmdDisown 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Kick from Group"
         BeginProperty Font 
            Name            =   "Porky's"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   300
         TabIndex        =   119
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label cmdAccess 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Change Rank"
         BeginProperty Font 
            Name            =   "Porky's"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   300
         TabIndex        =   118
         Top             =   2160
         Width           =   1815
      End
      Begin VB.Label lblGroupMenu 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Back to Group menu"
         BeginProperty Font 
            Name            =   "Porky's"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   300
         TabIndex        =   107
         Top             =   2790
         Width           =   1815
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rank:"
         BeginProperty Font 
            Name            =   "Porky's"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   210
         TabIndex        =   15
         Top             =   615
         Width           =   330
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         BeginProperty Font 
            Name            =   "Porky's"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   210
         TabIndex        =   14
         Top             =   360
         Width           =   390
      End
   End
   Begin VB.PictureBox picGuildMember 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00789298&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3375
      Left            =   9570
      ScaleHeight     =   3375
      ScaleWidth      =   2400
      TabIndex        =   17
      Top             =   5460
      Visible         =   0   'False
      Width           =   2400
      Begin VB.Label lblGroupMembers 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Online Members List"
         BeginProperty Font 
            Name            =   "Porky's"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   315
         TabIndex        =   109
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Label lblGroupChat 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Group Chat"
         BeginProperty Font 
            Name            =   "Porky's"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   550
         TabIndex        =   108
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label AdminOptions 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Admin Options"
         BeginProperty Font 
            Name            =   "Porky's"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   600
         TabIndex        =   91
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Group"
         BeginProperty Font 
            Name            =   "Porky's"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   855
         TabIndex        =   83
         Top             =   120
         Width           =   690
      End
      Begin VB.Label cmdLeave 
         BackStyle       =   0  'Transparent
         Caption         =   "Leave Group"
         BeginProperty Font 
            Name            =   "Porky's"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   660
         TabIndex        =   22
         Top             =   2760
         Width           =   1155
      End
      Begin VB.Label lblGuildRank 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rank"
         BeginProperty Font 
            Name            =   "Porky's"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   675
         TabIndex        =   21
         Top             =   1005
         Width           =   375
      End
      Begin VB.Label lblGuildName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Group"
         BeginProperty Font 
            Name            =   "Porky's"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   780
         TabIndex        =   20
         Top             =   630
         Width           =   450
      End
      Begin VB.Label Label17 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Rank:"
         BeginProperty Font 
            Name            =   "Porky's"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   120
         TabIndex        =   19
         Top             =   960
         Width           =   465
      End
      Begin VB.Label Label16 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         BeginProperty Font 
            Name            =   "Porky's"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   555
      End
   End
   Begin VB.PictureBox picEquipment 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00789298&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3375
      Left            =   9570
      ScaleHeight     =   3375
      ScaleWidth      =   2400
      TabIndex        =   23
      Top             =   5460
      Width           =   2400
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   555
         Left            =   930
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   35
         TabIndex        =   32
         Top             =   2380
         Width           =   555
         Begin VB.PictureBox EquipImage 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   4
            Left            =   15
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   137
            Top             =   15
            Width           =   495
         End
      End
      Begin VB.PictureBox AmuletImage2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   555
         Left            =   200
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   35
         TabIndex        =   30
         Top             =   120
         Width           =   555
         Begin VB.PictureBox EquipImage 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   6
            Left            =   15
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   139
            Top             =   15
            Width           =   495
         End
      End
      Begin VB.PictureBox Picture12 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   555
         Left            =   1660
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   35
         TabIndex        =   29
         Top             =   120
         Width           =   555
         Begin VB.PictureBox EquipImage 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   5
            Left            =   15
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   138
            Top             =   15
            Width           =   495
         End
      End
      Begin VB.PictureBox Picture6 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   555
         Left            =   1660
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   35
         TabIndex        =   28
         Top             =   1660
         Width           =   555
         Begin VB.PictureBox EquipImage 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   0
            Left            =   15
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   133
            Top             =   15
            Width           =   495
         End
      End
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   555
         Left            =   930
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   35
         TabIndex        =   27
         Top             =   1660
         Width           =   555
         Begin VB.PictureBox EquipImage 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   1
            Left            =   15
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   134
            Top             =   15
            Width           =   495
         End
      End
      Begin VB.PictureBox Picture7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   555
         Left            =   200
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   35
         TabIndex        =   26
         Top             =   1660
         Width           =   555
         Begin VB.PictureBox EquipImage 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   3
            Left            =   15
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   136
            Top             =   15
            Width           =   495
         End
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   555
         Left            =   930
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   35
         TabIndex        =   25
         Top             =   940
         Width           =   555
         Begin VB.PictureBox EquipImage 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   2
            Left            =   15
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   135
            Top             =   15
            Width           =   495
         End
      End
   End
   Begin VB.TextBox txtMyTextBox 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   30
      MaxLength       =   200
      TabIndex        =   6
      Top             =   7230
      Width           =   9495
   End
   Begin MSWinsockLib.Winsock Socket 
      Left            =   11520
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.PictureBox picCharStatus 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   3375
      Left            =   9570
      ScaleHeight     =   225
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   160
      TabIndex        =   67
      Top             =   5460
      Width           =   2400
      Begin VB.PictureBox lblLevelUp 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   405
         Picture         =   "frmMirage.frx":19112E
         ScaleHeight     =   23
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   103
         TabIndex        =   132
         Top             =   2805
         Width           =   1575
      End
      Begin VB.Label lblDEF 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "DEFENCE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   300
         Left            =   960
         TabIndex        =   80
         Top             =   1560
         Width           =   930
      End
      Begin VB.Label lblSTR 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "STRENGTH"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   300
         Left            =   840
         TabIndex        =   79
         Top             =   1200
         Width           =   1050
      End
      Begin VB.Label lblSPEED 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "SPEED"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   300
         Left            =   840
         TabIndex        =   78
         Top             =   1920
         Width           =   1050
      End
      Begin VB.Label lblMAGI 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "MAGIC"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   300
         Left            =   840
         TabIndex        =   77
         Top             =   2280
         Width           =   1050
      End
      Begin VB.Label lblLevel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "LEVEL"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   300
         Left            =   840
         TabIndex        =   76
         Top             =   480
         Width           =   1050
      End
      Begin VB.Label lblSTATWINDOW 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "CHARACTER"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   75
         Top             =   120
         Width           =   2175
      End
      Begin VB.Label Label22 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Level :  "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   74
         Top             =   480
         Width           =   1125
      End
      Begin VB.Label Label23 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Stache :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   73
         ToolTipText     =   "Affects purchase and sell prices at shops as well as Critical Hit chance."
         Top             =   2280
         Width           =   1125
      End
      Begin VB.Label Label24 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Speed :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   72
         ToolTipText     =   "Affects Block Chance, Attack Speed, Run Energy, and run chance in battle."
         Top             =   1920
         Width           =   1125
      End
      Begin VB.Label Label25 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Attack :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   71
         ToolTipText     =   "How much damage you can inflict upon enemies or other players."
         Top             =   1200
         Width           =   1125
      End
      Begin VB.Label Label26 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Defense :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   70
         ToolTipText     =   "How much damage you can resist from enemies or other players."
         Top             =   1560
         Width           =   1125
      End
      Begin VB.Label Label28 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Run Energy :  "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   69
         Top             =   840
         Width           =   1125
      End
      Begin VB.Label lblSP 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "SP"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   300
         Left            =   1230
         TabIndex        =   68
         Top             =   840
         Width           =   660
      End
   End
   Begin VB.PictureBox picScreen 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7200
      Left            =   0
      ScaleHeight     =   480
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   635
      TabIndex        =   34
      Top             =   0
      Width           =   9525
      Begin VB.PictureBox picFriendList 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3630
         Left            =   120
         Picture         =   "frmMirage.frx":192D78
         ScaleHeight     =   3600
         ScaleWidth      =   2970
         TabIndex        =   84
         Top             =   3600
         Width           =   3000
         Begin VB.TextBox txtFriend 
            Height          =   285
            Left            =   1170
            TabIndex        =   89
            Top             =   2760
            Width           =   1695
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "Delete"
            Height          =   375
            Left            =   1980
            TabIndex        =   88
            Top             =   3120
            Width           =   855
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "Add"
            Height          =   375
            Left            =   1020
            TabIndex        =   87
            Top             =   3120
            Width           =   855
         End
         Begin VB.CommandButton cmdHide 
            Caption         =   "Close"
            Height          =   375
            Left            =   90
            TabIndex        =   86
            Top             =   3120
            Width           =   855
         End
         Begin VB.Label lblFriend 
            BackStyle       =   0  'Transparent
            Height          =   255
            Index           =   9
            Left            =   75
            TabIndex        =   102
            Top             =   2400
            Width           =   2805
         End
         Begin VB.Label lblFriend 
            BackStyle       =   0  'Transparent
            Height          =   255
            Index           =   8
            Left            =   75
            TabIndex        =   101
            Top             =   2160
            Width           =   2805
         End
         Begin VB.Label lblFriend 
            BackStyle       =   0  'Transparent
            Height          =   255
            Index           =   7
            Left            =   75
            TabIndex        =   100
            Top             =   1920
            Width           =   2805
         End
         Begin VB.Label lblFriend 
            BackStyle       =   0  'Transparent
            Height          =   255
            Index           =   6
            Left            =   75
            TabIndex        =   99
            Top             =   1680
            Width           =   2805
         End
         Begin VB.Label lblFriend 
            BackStyle       =   0  'Transparent
            Height          =   255
            Index           =   5
            Left            =   75
            TabIndex        =   98
            Top             =   1440
            Width           =   2805
         End
         Begin VB.Label lblFriend 
            BackStyle       =   0  'Transparent
            Height          =   255
            Index           =   4
            Left            =   75
            TabIndex        =   97
            Top             =   1200
            Width           =   2805
         End
         Begin VB.Label lblFriend 
            BackStyle       =   0  'Transparent
            Height          =   255
            Index           =   3
            Left            =   75
            TabIndex        =   96
            Top             =   960
            Width           =   2805
         End
         Begin VB.Label lblFriend 
            BackStyle       =   0  'Transparent
            Height          =   255
            Index           =   2
            Left            =   75
            TabIndex        =   95
            Top             =   720
            Width           =   2805
         End
         Begin VB.Label lblFriend 
            BackStyle       =   0  'Transparent
            Height          =   255
            Index           =   1
            Left            =   75
            TabIndex        =   94
            Top             =   480
            Width           =   2805
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Friend Name:"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   90
            Top             =   2760
            Width           =   975
         End
         Begin VB.Label lblFriend 
            BackStyle       =   0  'Transparent
            Height          =   255
            Index           =   0
            Left            =   75
            TabIndex        =   85
            Top             =   240
            Width           =   2805
         End
      End
      Begin VB.PictureBox picGroupMsg 
         Appearance      =   0  'Flat
         BackColor       =   &H00789298&
         ForeColor       =   &H80000008&
         Height          =   2415
         Left            =   7200
         ScaleHeight     =   2385
         ScaleWidth      =   2265
         TabIndex        =   113
         Top             =   4680
         Visible         =   0   'False
         Width           =   2295
         Begin VB.TextBox GroupMsgText 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Porky's"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1215
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   115
            Top             =   480
            Width           =   2055
         End
         Begin VB.Label lblGroupMsgCancel 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Cancel"
            BeginProperty Font 
               Name            =   "Porky's"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1200
            TabIndex        =   117
            Top             =   1920
            Width           =   855
         End
         Begin VB.Label lblGroupMsgOk 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Send"
            BeginProperty Font 
               Name            =   "Porky's"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   116
            Top             =   1920
            Width           =   735
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Group Message"
            BeginProperty Font 
               Name            =   "Porky's"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   0
            TabIndex        =   114
            Top             =   0
            Width           =   2295
         End
      End
      Begin VB.PictureBox picGroupMembers 
         Appearance      =   0  'Flat
         BackColor       =   &H00789298&
         ForeColor       =   &H80000008&
         Height          =   2145
         Left            =   7680
         ScaleHeight     =   2115
         ScaleWidth      =   1785
         TabIndex        =   110
         Top             =   4950
         Visible         =   0   'False
         Width           =   1815
         Begin VB.ListBox lstGroupMembers 
            Appearance      =   0  'Flat
            BackColor       =   &H00789298&
            Height          =   1590
            ItemData        =   "frmMirage.frx":1B64EA
            Left            =   -15
            List            =   "frmMirage.frx":1B64F1
            TabIndex        =   111
            Top             =   -15
            Width           =   1815
         End
         Begin VB.Label lblGroupMemBack 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Back"
            BeginProperty Font 
               Name            =   "Porky's"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   420
            TabIndex        =   112
            Top             =   1680
            Width           =   975
         End
      End
      Begin VB.PictureBox picSpecialAtkDetails 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         ForeColor       =   &H80000008&
         Height          =   4215
         Left            =   7080
         ScaleHeight     =   4185
         ScaleWidth      =   2265
         TabIndex        =   124
         Top             =   840
         Visible         =   0   'False
         Width           =   2295
         Begin VB.Label lblSpecialAtkClose 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Close"
            BeginProperty Font 
               Name            =   "Porky's"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   130
            Top             =   3720
            Width           =   1455
         End
         Begin VB.Label lblSpecialAtkDesc 
            BackStyle       =   0  'Transparent
            Caption         =   "Description:"
            BeginProperty Font 
               Name            =   "Porky's"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1455
            Left            =   120
            TabIndex        =   129
            Top             =   2040
            Width           =   1935
         End
         Begin VB.Label lblSpecialAtkRange 
            BackStyle       =   0  'Transparent
            Caption         =   "Range:"
            BeginProperty Font 
               Name            =   "Porky's"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   128
            Top             =   1200
            Width           =   1815
         End
         Begin VB.Label lblSpecialAtkFPCost 
            BackStyle       =   0  'Transparent
            Caption         =   "FP Cost:"
            BeginProperty Font 
               Name            =   "Porky's"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   127
            Top             =   1560
            Width           =   1815
         End
         Begin VB.Label lblSpecialAtkDmg 
            BackStyle       =   0  'Transparent
            Caption         =   "Damage:"
            BeginProperty Font 
               Name            =   "Porky's"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   126
            Top             =   840
            Width           =   1815
         End
         Begin VB.Label lblSpecialAtkName 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Mushroom Barrier"
            BeginProperty Font 
               Name            =   "Porky's"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   125
            Top             =   120
            Width           =   2040
         End
      End
   End
   Begin VBMP.VBMPlayer BGSPlayer 
      Height          =   1095
      Left            =   4560
      Top             =   120
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1931
   End
   Begin VBMP.VBMPlayer MusicPlayer 
      Height          =   1095
      Left            =   6600
      Top             =   120
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1931
   End
   Begin VBMP.VBMPlayer SoundPlayer 
      Height          =   1095
      Left            =   2520
      Top             =   120
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1931
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1290
      Left            =   30
      Picture         =   "frmMirage.frx":1B6506
      Top             =   7560
      Width           =   9525
   End
   Begin VB.Label lblEquipment 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "  "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   10185
      TabIndex        =   66
      Top             =   2955
      Width           =   1260
   End
   Begin VB.Label lblCharStats 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "  "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   270
      Left            =   10185
      TabIndex        =   65
      Top             =   2610
      Width           =   1260
   End
   Begin VB.Label lblMenuQuit 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "  "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   270
      Left            =   10185
      TabIndex        =   64
      Top             =   5025
      Width           =   1260
   End
   Begin VB.Label lblGuild 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "  "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   270
      Left            =   10185
      TabIndex        =   16
      Top             =   4320
      Width           =   1260
   End
   Begin VB.Label lblWhosOnline 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "  "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   270
      Left            =   10185
      TabIndex        =   10
      Top             =   3630
      Width           =   1260
   End
   Begin VB.Label lblOptions 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "  "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   270
      Left            =   10185
      TabIndex        =   7
      Top             =   4680
      Width           =   1260
   End
   Begin VB.Label lblSpells 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "  "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   270
      Left            =   10185
      TabIndex        =   5
      Top             =   3285
      Width           =   1260
   End
   Begin VB.Label lblInventory 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "  "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   270
      Left            =   10185
      TabIndex        =   4
      Top             =   3975
      Width           =   1260
   End
   Begin VB.Menu mnuOnline 
      Caption         =   "Online List Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuGetInfo 
         Caption         =   "Get Info"
      End
      Begin VB.Menu mnuSendMessage 
         Caption         =   "Send Message"
      End
   End
End
Attribute VB_Name = "frmMirage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const TextThing = (-20)
Private Const TransparentChatBox = &H20&
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const EM_CHARFROMPOS& = &HD7
Private Type POINTAPI
    x As Long
    y As Long
End Type
Private FirstItemSlot As Long, SecondItemSlot As Long

Private Sub AdminOptions_Click()
    If LenB(GetPlayerGuild(MyIndex)) <> 0 Then
        picGuildAdmin.Visible = True
    End If
End Sub

Private Sub chkAutoRun_Click()
    Call WriteINI("CONFIG", "AutoRun", CStr(chkAutoRun.Value), App.Path & "\Config.ini")
End Sub

Private Sub chkSound_Click()
    Call WriteINI("CONFIG", "Sound", CStr(chkSound.Value), App.Path & "\Config.ini")
End Sub

Private Sub chkBubbleBar_Click()
    Call WriteINI("CONFIG", "SpeechBubbles", CStr(chkBubbleBar.Value), App.Path & "\Config.ini")
End Sub

Private Sub chkNpcBar_Click()
    Call WriteINI("CONFIG", "NPCBar", CStr(chkNpcBar.Value), App.Path & "\Config.ini")
End Sub

Private Sub chkNpcDamage_Click()
    Call WriteINI("CONFIG", "NPCDamage", CStr(chkNpcDamage.Value), App.Path & "\Config.ini")
End Sub

Private Sub chkNpcName_Click()
    Call WriteINI("CONFIG", "NPCName", CStr(chkNpcName.Value), App.Path & "\Config.ini")
End Sub

Private Sub chkPlayerBar_Click()
    Call WriteINI("CONFIG", "PlayerBar", CStr(chkPlayerBar.Value), App.Path & "\Config.ini")
End Sub

Private Sub chkPlayerDamage_Click()
    Call WriteINI("CONFIG", "PlayerDamage", CStr(chkPlayerDamage.Value), App.Path & "\Config.ini")
End Sub

Private Sub chkAutoScroll_Click()
    Call WriteINI("CONFIG", "AutoScroll", CStr(chkAutoScroll.Value), App.Path & "\Config.ini")
End Sub

Private Sub chkPlayerName_Click()
    Call WriteINI("CONFIG", "PlayerName", CStr(chkPlayerName.Value), App.Path & "\Config.ini")
End Sub

Private Sub chkMusic_Click()
    If chkMusic = Checked Then
        Call WriteINI("CONFIG", "Music", "1", App.Path & "\Config.ini")
      
        If Not Player(MyIndex).InBattle Then
            Call PlayBGM(Trim$(Map(GetPlayerMap(MyIndex)).music))
        End If
    Else
        Call WriteINI("CONFIG", "Music", "0", App.Path & "\Config.ini")
        Call StopBGM
    End If
End Sub

Private Sub chkTurnBased_Click()
    Call WriteINI("CONFIG", GetPlayerName(MyIndex) & " - TurnBased", CStr(chkTurnBased.Value), App.Path & "\Config.ini")
    Call SendData(CPackets.Chasturnbased & SEP_CHAR & chkTurnBased.Value & END_CHAR)
End Sub

Private Sub cmdAdd_Click()
    Dim i As Integer

    For i = 0 To 8
       If LenB(lblFriend(i).Caption) < 2 Then
          lblFriend(i).Caption = txtFriend.Text
          txtFriend.Text = vbNullString
          Call SendData(CPackets.Ccaption & SEP_CHAR & Trim$(lblFriend(i).Caption) & SEP_CHAR & i & END_CHAR)
       ElseIf LenB(lblFriend(i).Caption) >= 2 And LenB(lblFriend(i + 1).Caption) < 2 Then
          lblFriend(i + 1).Caption = txtFriend.Text
          txtFriend.Text = vbNullString
          Call SendData(CPackets.Ccaption & SEP_CHAR & Trim$(lblFriend(i + 1).Caption) & SEP_CHAR & (i + 1) & END_CHAR)
       End If
    Next i
End Sub

Private Sub cmdDelete_Click()
    Dim i As Integer

    For i = 0 To 9
        If lblFriend(i).Caption = txtFriend.Text Then
            lblFriend(i).Caption = vbNullString
            Call SendData(CPackets.Ccaption & SEP_CHAR & lblFriend(i).Caption & SEP_CHAR & i & END_CHAR)
            txtFriend.Text = vbNullString
        End If
    Next i
End Sub

Private Sub cmdHide_Click()
    picFriendList.Visible = False
End Sub

Private Sub cmdLeave_Click()
    Call SendGuildLeave
    
    If GetPlayerGuild(MyIndex) <> vbNullString Then
        picGuildMember.Visible = False
        picInventory.Visible = True
    End If
End Sub

Private Sub cmdMember_Click()
    Call SendGuildMemberRequest(txtName.Text, 0)
End Sub

Private Sub cmdSaveConfig_Click()
    picOptions.Visible = False
End Sub

Private Sub EquipImage_Click(Index As Integer)
    If Player(MyIndex).InBattle = True Then
        Exit Sub
    End If
    
    If GetPlayerEquipSlotNum(MyIndex, Index + 1) > 0 Then
        Call SendData(CPackets.Cunequip & SEP_CHAR & Index + 1 & END_CHAR)
    End If
End Sub

Private Sub Form_Load()
    SetWindowLong txtChat.hWnd, TextThing, TransparentChatBox
    
    picFriendList.Visible = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    itmDesc.Visible = False
End Sub

Private Sub Form_Terminate()
    Call GameDestroy
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call GameDestroy
End Sub

Private Sub lblDetails_Click()
    If Player(MyIndex).InBattle = True Then
        If IsPlayerTurn = False Then
            Exit Sub
        Else
            If CanUseSpecial = False Then
                Exit Sub
            End If
        End If
    End If
    
    If Player(MyIndex).Spell(lstSpells.ListIndex + 1) > 0 Then
        Call SendData(CPackets.Cspecialattackdetails & SEP_CHAR & Player(MyIndex).Spell(lstSpells.ListIndex + 1) & END_CHAR)
    End If
End Sub

Private Sub lblFriend_Click(Index As Integer)
    If lblFriend(Index).Caption <> vbNullString Then
        If lblFriend(Index).ForeColor <> &HFF& Then
            KeepUsername = True
            frmPersonalMsg.EnterUsername.Text = lblFriend(Index).Caption
            frmPersonalMsg.Visible = True
            frmPersonalMsg.EnterMsg.SetFocus
        Else
            Call PlayerMsg(lblFriend(Index).Caption & " is currently offline.", lblFriend(Index).Caption)
        End If
    End If

End Sub

Private Sub lblFriendList_Click()
    Dim i As Integer
    picFriendList.Visible = True
    
    ' Finds out if players are online or not
    For i = 0 To 9
        Call SendData(CPackets.Cfriendonlineoffline & SEP_CHAR & lblFriend(i).Caption & SEP_CHAR & i & END_CHAR)
    Next i
End Sub

Private Sub lblGroupChat_Click()
    If LenB(GetPlayerGuild(MyIndex)) > 0 Then
        picGroupMembers.Visible = False
        picGroupMsg.Visible = True
    End If
End Sub

Private Sub lblGroupMemBack_Click()
    picGroupMembers.Visible = False
End Sub

Private Sub lblGroupMembers_Click()
    If LenB(GetPlayerGuild(MyIndex)) > 0 Then
        Call SendData(CPackets.Cgroupmemberlist & END_CHAR)
        picGroupMsg.Visible = False
    End If
End Sub

Private Sub lblGroupMsgCancel_Click()
    picGroupMsg.Visible = False
    GroupMsgText.Text = vbNullString
End Sub

Private Sub lblGroupMsgOk_Click()
    If Len(GroupMsgText.Text) > 0 Then
        Call GroupMsg(Trim$(GroupMsgText.Text))
    End If
    GroupMsgText.Text = vbNullString
End Sub

Private Sub lblLevelUp_Click()
    Call frmLevelUp.Show(vbModeless, frmMirage)
End Sub

Private Sub lblOptions_Click()
    chkPlayerName.Value = CInt(ReadINI("CONFIG", "PlayerName", App.Path & "\Config.ini"))
    chkPlayerDamage.Value = CInt(ReadINI("CONFIG", "PlayerDamage", App.Path & "\Config.ini"))
    chkPlayerBar.Value = CInt(ReadINI("CONFIG", "PlayerBar", App.Path & "\Config.ini"))
    chkNpcName.Value = CInt(ReadINI("CONFIG", "NPCName", App.Path & "\Config.ini"))
    chkNpcDamage.Value = CInt(ReadINI("CONFIG", "NPCDamage", App.Path & "\Config.ini"))
    chkNpcBar.Value = CInt(ReadINI("CONFIG", "NPCBar", App.Path & "\Config.ini"))
    chkMusic.Value = CInt(ReadINI("CONFIG", "Music", App.Path & "\Config.ini"))
    chkSound.Value = CInt(ReadINI("CONFIG", "Sound", App.Path & "\Config.ini"))
    chkBubbleBar.Value = CInt(ReadINI("CONFIG", "SpeechBubbles", App.Path & "\Config.ini"))
    chkAutoScroll.Value = CInt(ReadINI("CONFIG", "AutoScroll", App.Path & "\Config.ini"))
    chkTurnBased.Value = CInt(ReadINI("CONFIG", GetPlayerName(MyIndex) & " - TurnBased", App.Path & "\Config.ini"))
    chkAutoRun.Value = CInt(ReadINI("CONFIG", "AutoRun", App.Path & "\Config.ini"))
    
    picOptions.Visible = True
End Sub

Private Sub lblGuild_Click()
    If LenB(GetPlayerGuild(MyIndex)) <> 0 Then
        lblGuildName.Caption = GetPlayerGuild(MyIndex)
        lblGuildRank.Caption = GetPlayerGuildAccess(MyIndex)
    Else
        lblGuildName.Caption = "None"
        lblGuildRank.Caption = "None"
    End If
    
    picInventory.Visible = False
    picPlayerSpells.Visible = False
    picEquipment.Visible = False
    picWhosOnline.Visible = False
    picCharStatus.Visible = False
    picGuildAdmin.Visible = False
    picGuildMember.Visible = True
    
    If GetPlayerGuildAccess(MyIndex) > 2 Then
        AdminOptions.Visible = True
    Else
        AdminOptions.Visible = False
    End If
End Sub

Private Sub lblGroupMenu_Click()
    picGuildAdmin.Visible = False
    picGuildMember.Visible = True
End Sub

Private Sub lblEquipment_Click()
    Call UpdateVisInv
    
    picInventory.Visible = False
    picPlayerSpells.Visible = False
    picWhosOnline.Visible = False
    picGuildMember.Visible = False
    picGuildAdmin.Visible = False
    picCharStatus.Visible = False
    picEquipment.Visible = True
End Sub

Private Sub lblInventory_Click()
    Call UpdateVisInv
    
    picGuildMember.Visible = False
    picGuildAdmin.Visible = False
    picEquipment.Visible = False
    picPlayerSpells.Visible = False
    picWhosOnline.Visible = False
    picCharStatus.Visible = False
    picInventory.Visible = True
End Sub

Private Sub lblSpecialAtkClose_Click()
    picSpecialAtkDetails.Visible = False
End Sub

Private Sub lblSpells_Click()
    Call SendRequestSpells
    
    picInventory.Visible = False
    picGuildAdmin.Visible = False
    picEquipment.Visible = False
    picGuildMember.Visible = False
    picWhosOnline.Visible = False
    picCharStatus.Visible = False
    picPlayerSpells.Visible = True
End Sub

Private Sub lblCharStats_Click()
    picWhosOnline.Visible = False
    picInventory.Visible = False
    picEquipment.Visible = False
    picGuildMember.Visible = False
    picGuildAdmin.Visible = False
    picPlayerSpells.Visible = False
    picCharStatus.Visible = True
End Sub

Private Sub lblForgetSpell_Click()
    Call SendForgetSpell(lstSpells.ListIndex + 1)
End Sub

Private Sub lblMenuQuit_Click()
    If IsCooking = True Or IsTrading = True Or IsShopping = True Or IsBanking = True Or IsHiderFrozen = True Then
        Exit Sub
    End If

    frmLogout.Visible = False
    frmLogout.Visible = True
End Sub

Private Sub lblSTATWINDOW_Click()
    Call SendRequestMyStats
End Sub

Private Sub lblWhosOnline_Click()
    Call SendOnlineList
    
    picInventory.Visible = False
    picEquipment.Visible = False
    picGuildMember.Visible = False
    picGuildAdmin.Visible = False
    picPlayerSpells.Visible = False
    picCharStatus.Visible = False
    picWhosOnline.Visible = True
End Sub

Private Sub lstOnline_DblClick()
    Call SendPlayerChat(Trim$(lstOnline.Text))
End Sub

Private Sub lstOnline_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' Show the player menu with commands - Get Info and Send Message
    If Button = 2 Then
        ' Make sure the user is selected on a player before displaying the menu
        
        If lstOnline.ListIndex > -1 Then
            If lstOnline.List(lstOnline.ListIndex) <> GetPlayerName(MyIndex) Then
                PopupMenu mnuOnline
            End If
        End If
    End If
End Sub

Private Sub mnuGetInfo_Click()
    ' Make sure the user is selected on another player
    If lstOnline.ListIndex > -1 Then
        Dim PlayerName As String
        
        PlayerName = lstOnline.List(lstOnline.ListIndex)
    
        If PlayerName <> GetPlayerName(MyIndex) Then
            Call SendGetPlayerInfo(PlayerName)
        End If
    End If
End Sub

Private Sub mnuSendMessage_Click()
    ' Make sure the user is selected on another player
    If lstOnline.ListIndex > -1 Then
        Dim PlayerName As String
        
        PlayerName = lstOnline.List(lstOnline.ListIndex)
    
        If PlayerName <> GetPlayerName(MyIndex) Then
            ' Display the Personal Message screen and fill in the user's name
            KeepUsername = True
            frmPersonalMsg.EnterUsername.Text = PlayerName
            frmPersonalMsg.Visible = True
            frmPersonalMsg.EnterMsg.SetFocus
        End If
    End If
End Sub

Private Sub picEquipment_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    itmDesc.Visible = False
End Sub

Private Sub picFriendList_Click()
    picFriendList.Visible = True
End Sub

Private Sub picInventory3_DblClick()
    If Inventory < 1 Then
        Exit Sub
    End If

    Dim ItemNum As Long
    
    ItemNum = GetPlayerInvItemNum(MyIndex, Inventory)
    
    If Player(MyIndex).InBattle = True And IsPlayerTurn = True Then
        If CanUseItem = False Then
            Exit Sub
        Else
            If ItemNum > 0 Then
                Call SendUseTurnBasedItem(Inventory, ItemNum)
            End If
        End If
    End If
    
    If ItemNum < 1 Then
        Exit Sub
    End If
    
    If Player(MyIndex).InBattle = False Then
        Call SendUseItem(Inventory)
    End If
    
    If Item(ItemNum).Type >= ITEM_TYPE_WEAPON And Item(ItemNum).Type <= ITEM_TYPE_MUSHROOMBADGE Then
        Call UpdateVisInv
    End If
End Sub

Private Sub picInventory3_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Player(MyIndex).InBattle = True Then
        Exit Sub
    End If
    
    ' Stop drawing drag and drop
    picDragBox.Visible = False
    
    ' Stop players from moving around their inventory while trading
    If IsTrading = True Then
        FirstItemSlot = 0
        SecondItemSlot = 0
        
        Exit Sub
    End If
    
    SecondItemSlot = GetInvSlot(x, y)
    
    If FirstItemSlot <= 0 Or SecondItemSlot <= 0 Then
        FirstItemSlot = 0
        SecondItemSlot = 0
        
        Exit Sub
    End If
    
    If Player(MyIndex).NewInv(FirstItemSlot).num > 0 And FirstItemSlot <> SecondItemSlot Then
        Call SendData(CPackets.Cdragdropinv & SEP_CHAR & FirstItemSlot & SEP_CHAR & SecondItemSlot & END_CHAR)
    End If

    FirstItemSlot = 0
    SecondItemSlot = 0
End Sub

Private Sub picInventory3_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Player(MyIndex).InBattle = True And CanUseItem = False Then
        Exit Sub
    ElseIf Player(MyIndex).InBattle = True And IsPlayerTurn = False Then
        Exit Sub
    End If
    
    Inventory = GetInvSlot(x, y)
    
    If Inventory < 1 Then
        Exit Sub
    End If
    
    Dim ItemNum As Long
    
    ItemNum = GetPlayerInvItemNum(MyIndex, Inventory)
    
    If Player(MyIndex).InBattle = False Then
        If Button = 1 Then
            FirstItemSlot = Inventory
        End If
    End If

    ' Right-click
    If Button = 2 Then
        If Player(MyIndex).InBattle = False Then
            Call DropItem
        End If
    End If
End Sub

Private Sub DrawDragDrop(ItemNum As Long, x As Single, y As Single)
    If ItemNum < 1 Then
        Exit Sub
    End If
    
    Dim srec As RECT, drec As RECT
    
    With srec
        .Left = (Item(ItemNum).Pic - (Item(ItemNum).Pic \ 6) * 6) * PIC_X
        .Right = srec.Left + PIC_X
        .Top = (Item(ItemNum).Pic \ 6) * PIC_Y
        .Bottom = srec.Top + PIC_Y
    End With
    
    With drec
        .Left = 0
        .Right = PIC_X
        .Top = 0
        .Bottom = PIC_Y
    End With
    
    picDragBox.Left = x
    picDragBox.Top = y
    
    Call DD_ItemSurf.BltToDC(picDragBox.hDC, srec, drec)
    
    If picDragBox.Visible = False Then
        picDragBox.Visible = True
    End If
End Sub

Private Sub EquipImage_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim d As Long, ItemNum As Long
    Dim Msg As String
    
    If Player(MyIndex).InBattle = True Then
        Exit Sub
    End If
    
    d = Index
    ItemNum = GetPlayerEquipSlotNum(MyIndex, d + 1)
    
    If ItemNum > 0 Then
        If Trim$(Item(ItemNum).desc) = vbNullString Then
            itmDesc.Top = 96
            itmDesc.Height = 209
        Else
            itmDesc.Top = 8
            itmDesc.Height = 369
        End If
        
        If Item(ItemNum).Ammo > -1 Then
            descName.Caption = Trim$(Item(ItemNum).Name) & " (Ammo: " & GetPlayerEquipSlotAmmo(MyIndex, d + 1) & ")"
        Else
            descName.Caption = Trim$(Item(ItemNum).Name)
        End If
        
        descHP.Caption = "HP: " & Item(ItemNum).HPReq
        descFP.Caption = "FP: " & Item(ItemNum).FPReq
        descStr.Caption = "Attack: " & Item(ItemNum).StrReq
        descDef.Caption = "Defense: " & Item(ItemNum).DefReq
        descSpeed.Caption = "Speed: " & Item(ItemNum).SpeedReq
        descMagic.Caption = "Stache: " & Item(ItemNum).MagicReq
        
        If Item(ItemNum).ClassReq > -1 Then
            descClass.Caption = "Character: " & Trim$(Class(Item(ItemNum).ClassReq).Name)
        Else
            descClass.Caption = "Character: None"
        End If
        
        descLevel.Caption = "Level: " & Item(ItemNum).LevelReq
        descHpMp.Caption = "HP: " & Item(ItemNum).AddHP & " FP: " & Item(ItemNum).AddMP & " SP: " & Item(ItemNum).AddSP
        descSD.Caption = "Attack: " & Item(ItemNum).AddSTR & " Defense: " & Item(ItemNum).AddDef
        descMS.Caption = "Speed: " & Item(ItemNum).AddSpeed & " Stache: " & Item(ItemNum).AddMAGI
        
        ' Determine message for Critical and Block chances
        If Item(ItemNum).AddCritChance > 0 Then
            Msg = "Crit: +" & Item(ItemNum).AddCritChance & "% "
        ElseIf Item(ItemNum).AddCritChance < 0 Then
            Msg = "Crit: " & Item(ItemNum).AddCritChance & "% "
        Else
            Msg = "Crit: " & Item(ItemNum).AddCritChance & "% "
        End If
        If Item(ItemNum).AddBlockChance > 0 Then
            Msg = Msg & "Block: +" & Item(ItemNum).AddBlockChance & "% "
        ElseIf Item(ItemNum).AddBlockChance < 0 Then
            Msg = Msg & "Block: " & Item(ItemNum).AddBlockChance & "% "
        Else
            Msg = Msg & "Block: " & Item(ItemNum).AddBlockChance & "% "
        End If
        descCritBlock.Caption = Msg
        desc.Caption = Trim$(Item(ItemNum).desc)
        
        itmDesc.Visible = True
    Else
        itmDesc.Visible = False
    End If
End Sub

Function GetInvSlot(ByVal x As Single, y As Single) As Long
    Dim i As Integer, LeftInvSlot As Integer, TopInvSlot As Integer
    Dim rec As RECT
    
    TopInvSlot = -1
    
    For i = InventorySlotsIndex To Player(MyIndex).MaxInv
        If LeftInvSlot > 3 Then
            LeftInvSlot = 0
        End If
                        
        If LeftInvSlot = 0 Then
            TopInvSlot = TopInvSlot + 1
        End If
        
        With rec
            .Left = 6 + (LeftInvSlot * 38)
            .Right = .Left + PIC_X
            .Top = 1 + (TopInvSlot * 35)
            .Bottom = .Top + PIC_Y
            
            ' Check if the user's mouse is within the range of an inventory slot
            If (x >= .Left And x <= .Right) And (y >= .Top And y <= .Bottom) Then
                GetInvSlot = i
                
                Exit Function
            End If
        End With
        
        LeftInvSlot = LeftInvSlot + 1
    Next
End Function

Private Sub picInventory3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Player(MyIndex).InBattle = True And IsPlayerTurn = True And CanUseItem = False Then
        Exit Sub
    End If
    
    Dim d As Long, ItemNum As Long
    Dim Msg As String
    
    d = GetInvSlot(x, y)
    
    If d < 1 Then
        Exit Sub
    End If
    
    ' Check if the user is dragging
    If FirstItemSlot > 0 And Button = 1 Then
        Call DrawDragDrop(GetPlayerInvItemNum(MyIndex, FirstItemSlot), x, y)
        
        Exit Sub
    End If
    
    ItemNum = GetPlayerInvItemNum(MyIndex, d)
    
    If ItemNum > 0 Then
        If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Or Item(ItemNum).Stackable = 1 Then
            If Trim$(Item(ItemNum).desc) = vbNullString Then
                itmDesc.Top = 224
                itmDesc.Height = 17
            Else
                itmDesc.Top = 8
                itmDesc.Height = 369
            End If
        Else
            If Trim$(Item(ItemNum).desc) = vbNullString Then
                itmDesc.Top = 96
                itmDesc.Height = 209
            Else
                itmDesc.Top = 8
                itmDesc.Height = 369
            End If
        End If
        If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Or Item(ItemNum).Stackable = 1 Then
            descName.Caption = Trim$(Item(ItemNum).Name) & " (" & GetPlayerInvItemValue(MyIndex, d) & ")"
        Else
            If Item(ItemNum).Ammo > -1 Then
                descName.Caption = Trim$(Item(ItemNum).Name) & " (Ammo: " & GetPlayerInvItemAmmo(MyIndex, d) & ")"
            Else
                descName.Caption = Trim$(Item(ItemNum).Name)
            End If
        End If
        
        descHP.Caption = "HP: " & Item(ItemNum).HPReq
        descFP.Caption = "FP: " & Item(ItemNum).FPReq
        descStr.Caption = "Attack: " & Item(ItemNum).StrReq
        descDef.Caption = "Defense: " & Item(ItemNum).DefReq
        descSpeed.Caption = "Speed: " & Item(ItemNum).SpeedReq
        descMagic.Caption = "Stache: " & Item(ItemNum).MagicReq
        
        If Item(ItemNum).Type >= ITEM_TYPE_CHANGEHPFPSP And Item(ItemNum).Type <= ITEM_TYPE_SCRIPTED Then
            descClass.Caption = "Character: " & "None"
        ElseIf Item(ItemNum).Type = ITEM_TYPE_NONE Or Item(ItemNum).ClassReq = -1 Then
            descClass.Caption = "Character: " & "None"
        Else
            descClass.Caption = "Character: " & Trim$(Class(Item(ItemNum).ClassReq).Name)
        End If
        
        descLevel.Caption = "Level: " & Item(ItemNum).LevelReq
        descHpMp.Caption = "HP: " & Item(ItemNum).AddHP & " FP: " & Item(ItemNum).AddMP & " SP: " & Item(ItemNum).AddSP
        descSD.Caption = "Attack: " & Item(ItemNum).AddSTR & " Defense: " & Item(ItemNum).AddDef
        descMS.Caption = "Speed: " & Item(ItemNum).AddSpeed & " Stache: " & Item(ItemNum).AddMAGI
        ' Determine message for Critical and Block chances
        If Item(ItemNum).AddCritChance > 0 Then
            Msg = "Crit: +" & Item(ItemNum).AddCritChance & "% "
        ElseIf Item(ItemNum).AddCritChance < 0 Then
            Msg = "Crit: " & Item(ItemNum).AddCritChance & "% "
        Else
            Msg = "Crit: " & Item(ItemNum).AddCritChance & "% "
        End If
        If Item(ItemNum).AddBlockChance > 0 Then
            Msg = Msg & "Block: +" & Item(ItemNum).AddBlockChance & "% "
        ElseIf Item(ItemNum).AddBlockChance < 0 Then
            Msg = Msg & "Block: " & Item(ItemNum).AddBlockChance & "% "
        Else
            Msg = Msg & "Block: " & Item(ItemNum).AddBlockChance & "% "
        End If
        descCritBlock.Caption = Msg
        desc.Caption = Trim$(Item(ItemNum).desc)
        
        itmDesc.Visible = True
    Else
        itmDesc.Visible = False
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim i As Long, ScreenID As Long, ItemNum As Long

    Call CheckInput(0, KeyCode, Shift)

    If KeyCode = vbKeyF1 Then
        If GetPlayerAccess(MyIndex) > 1 Then
            frmAdmin.Visible = False
            frmAdmin.Visible = True
        End If
    End If

    If Player(MyIndex).InBattle = False Then
        If KeyCode = vbKeyF2 Then
            For i = 1 To Player(MyIndex).MaxInv
              ItemNum = GetPlayerInvItemNum(MyIndex, i)
                If ItemNum > 0 And ItemNum <= MAX_ITEMS Then
                    If Item(ItemNum).Type = ITEM_TYPE_CHANGEHPFPSP Then
                        If Item(ItemNum).Data1 <> 0 And Item(ItemNum).Data2 = 0 And Item(ItemNum).Data3 = 0 Then
                            Call PlaySound("spm_get_health.wav")
                            Call SendUseItem(i)
                            Exit Sub
                        End If
                    End If
                Else
                    If i = Player(MyIndex).MaxInv Then
                        Call AddText("You don't have any HP healing items!", BRIGHTRED)
                    End If
                End If
            Next i
        End If
    End If
    
    If Player(MyIndex).InBattle = False Then
        If KeyCode = vbKeyF3 Then
            For i = 1 To Player(MyIndex).MaxInv
              ItemNum = GetPlayerInvItemNum(MyIndex, i)
                If ItemNum > 0 And ItemNum <= MAX_ITEMS Then
                    If Item(ItemNum).Type = ITEM_TYPE_CHANGEHPFPSP Then
                        If Item(ItemNum).Data1 = 0 And Item(ItemNum).Data2 <> 0 And Item(ItemNum).Data3 = 0 Then
                            Call PlaySound("spm_get_health.wav")
                            Call SendUseItem(i)
                            Exit Sub
                        End If
                    End If
                Else
                    If i = Player(MyIndex).MaxInv Then
                        Call AddText("You don't have any FP healing items!", BRIGHTRED)
                    End If
                End If
            Next i
        End If
    End If

    If KeyCode = vbKeyF4 Then
        If GetPlayerAccess(MyIndex) = 5 Then
            frmGuild.Show vbModeless, frmMirage
        End If
    End If
    
    If GetPlayerGuildAccess(MyIndex) > 2 Then
        If KeyCode = vbKeyF5 Then
            picInventory.Visible = False
            picGuildMember.Visible = False
            picEquipment.Visible = False
            picPlayerSpells.Visible = False
            picWhosOnline.Visible = False
            picGuildAdmin.Visible = True
        End If
    End If
    
    If GetPlayerAccess(MyIndex) >= 1 Then
        If KeyCode = vbKeyF6 Then
            frmWarn.Show vbModeless, frmMirage
        End If
    End If
    
    If Player(MyIndex).InBattle = False And Player(MyIndex).Jumping = False Then
        If KeyCode = vbKeyInsert Then
            If (lstSpells.ListIndex + 1) > 0 Then
                If GetTickCount > Player(MyIndex).AttackTimer + 1000 Then
                    Call SendData(CPackets.Ccast & SEP_CHAR & (lstSpells.ListIndex + 1) & END_CHAR)

                    Player(MyIndex).Attacking = 1
                    Player(MyIndex).AttackTimer = GetTickCount
                    Player(MyIndex).CastedSpell = Yes
                End If
            Else
                Call AddText("There's no special attack here!", BRIGHTRED)
            End If
        End If
    End If

    If KeyCode = vbKeyF11 Then
        ScreenShot.Picture = CaptureForm(frmMirage)

        Do
            If FileExists("Screenshot_" & ScreenID & ".bmp") Then
                ScreenID = ScreenID + 1
            Else
                Call SavePicture(ScreenShot.Picture, App.Path & "\Screenshot_" & ScreenID & ".bmp")
                Exit Do
            End If
        Loop
    End If

    If KeyCode = vbKeyF12 Then
        ScreenShot.Picture = CaptureArea(frmMirage, picScreen.Left, picScreen.Top, picScreen.Width, picScreen.Height)

        Do
            If FileExists("Screenshot_" & ScreenID & ".bmp") Then
                ScreenID = ScreenID + 1
            Else
                Call SavePicture(ScreenShot.Picture, App.Path & "\Screenshot_" & ScreenID & ".bmp")
                Exit Do
            End If
        Loop
    End If

    If KeyCode = vbKeyHome Then
        If Player(MyIndex).Moving = No Then
            If Player(MyIndex).Dir = DIR_DOWN Then
                Call SetPlayerDir(MyIndex, DIR_LEFT)
            ElseIf Player(MyIndex).Dir = DIR_LEFT Then
                Call SetPlayerDir(MyIndex, DIR_UP)
            ElseIf Player(MyIndex).Dir = DIR_UP Then
                Call SetPlayerDir(MyIndex, DIR_RIGHT)
            ElseIf Player(MyIndex).Dir = DIR_RIGHT Then
                Call SetPlayerDir(MyIndex, DIR_DOWN)
            End If

            Call SendPlayerDir
        End If
    End If
    
    ' Jumping
    If Player(MyIndex).InBattle = False Then
        ' Disable jumping in STS
        If GetPlayerMap(MyIndex) <> 33 Then
            If KeyCode = vbKeyDelete Then
                If Player(MyIndex).Jumping = False Then
                    Player(MyIndex).JumpDir = 0
                    Player(MyIndex).TempJumpAnim = 0
                    Player(MyIndex).JumpAnim = 0
                    Player(MyIndex).Jumping = True
                    
                    ' Send to everyone that the player is jumping
                    Call SendData(CPackets.Cjumping & SEP_CHAR & Player(MyIndex).JumpDir & SEP_CHAR & Player(MyIndex).TempJumpAnim & SEP_CHAR & Player(MyIndex).JumpAnim & END_CHAR)
                    Call SendHotScript(2)
                End If
            End If
        End If
    End If
    
    ' Special Badges
    If Player(MyIndex).InBattle = False And Player(MyIndex).Jumping = False Then
        If KeyCode = vbKeyEnd Then
            If GetPlayerEquipSlotNum(MyIndex, 4) > 0 Then
                If Player(MyIndex).Attacking = 0 Then
                    If SpecialBadgeTime + 1000 < GetTickCount Then
                        Call SendUseSpecialBadge(GetPlayerEquipSlotNum(MyIndex, 4))
                                    
                        ' Set player to attacking
                        Player(MyIndex).Attacking = 1
                        Player(MyIndex).AttackTimer = GetTickCount
                                        
                        ' Reset timer for using special badges
                        If GetPlayerEquipSlotNum(MyIndex, 4) = 262 Then
                            SpecialBadgeTime = GetTickCount + 3000
                        Else
                            SpecialBadgeTime = GetTickCount
                        End If
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub picOptions_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    SOffsetX = x
    SOffsetY = y
End Sub

Private Sub picOptions_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call MovePicture(frmMirage.picOptions, Button, Shift, x, y)
End Sub

Private Sub picScreen_GotFocus()
    On Error Resume Next

    txtMyTextBox.SetFocus
End Sub

Private Sub picScreen_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim i As Long

    If (Button = 1 Or Button = 2) And InEditor Then
        Call EditorMouseDown(Button, Shift, CurX, CurY)
    End If

    If Button = 1 And Not InEditor Then
        Call PlayerSearch(Button, Shift, CurX, CurY)
    End If
    
    If Shift = 1 And Not InEditor Then
        If GetPlayerAccess(MyIndex) > 1 Then
            Call LocalWarp(CurX, CurY)
        End If
    End If
End Sub

Private Sub picScreen_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    CurX = ((x + (NewPlayerX * PIC_X)) \ PIC_X)
    CurY = ((y + (NewPlayerY * PIC_Y)) \ PIC_Y)

    If (Button = 1 Or Button = 2) And InEditor Then
        Call EditorMouseDown(Button, Shift, CurX, CurY)
    End If

    frmMapEditor.Caption = "Map Editor - " & "X: " & CurX & " Y: " & CurY
End Sub

Private Sub picInventory_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    itmDesc.Visible = False
End Sub

Private Sub PrivateMsg_Click()
    KeepUsername = False
    frmPersonalMsg.EnterUsername.Text = vbNullString
    frmPersonalMsg.EnterMsg.Text = vbNullString
    frmPersonalMsg.Visible = True
End Sub

Private Sub scrlBltText_Change()
    Dim i As Long

    For i = 1 To MAX_BLT_LINE
        BattlePMsg(i).Index = 1
        BattlePMsg(i).Time = i
        BattleMMsg(i).Index = 1
        BattleMMsg(i).Time = i
    Next i

    MAX_BLT_LINE = scrlBltText.Value

    ReDim BattlePMsg(1 To MAX_BLT_LINE) As BattleMsgRec
    ReDim BattleMMsg(1 To MAX_BLT_LINE) As BattleMsgRec

    lblLines.Caption = "On Screen Text Line Amount: " & scrlBltText.Value
End Sub

Private Sub Socket_DataArrival(ByVal bytesTotal As Long)
    If IsConnected Then
        Call IncomingData(bytesTotal)
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Call HandleKeypresses(KeyAscii)

    If (KeyAscii = vbKeyReturn) Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call CheckInput(1, KeyCode, Shift)
End Sub

Private Sub tmrRainDrop_Timer()
    If BLT_RAIN_DROPS > RainIntensity Then
        tmrRainDrop.Enabled = False
        Exit Sub
    End If

    If BLT_RAIN_DROPS > 0 Then
        If DropRain(BLT_RAIN_DROPS).Randomized = False Then
            Call RNDRainDrop(BLT_RAIN_DROPS)
        End If
    End If

    BLT_RAIN_DROPS = BLT_RAIN_DROPS + 1

    If tmrRainDrop.Interval > 30 Then
        tmrRainDrop.Interval = tmrRainDrop.Interval - 10
    End If
End Sub

Private Sub tmrSnowDrop_Timer()
    If BLT_SNOW_DROPS > RainIntensity Then
        tmrSnowDrop.Enabled = False
        Exit Sub
    End If

    If BLT_SNOW_DROPS > 0 Then
        If DropSnow(BLT_SNOW_DROPS).Randomized = False Then
            Call RNDSnowDrop(BLT_SNOW_DROPS)
        End If
    End If

    BLT_SNOW_DROPS = BLT_SNOW_DROPS + 1

    If tmrSnowDrop.Interval > 30 Then
        tmrSnowDrop.Interval = tmrSnowDrop.Interval - 10
    End If
End Sub

Private Sub txtChat_GotFocus()
    On Error Resume Next

    frmMirage.txtMyTextBox.SetFocus
End Sub

Private Sub txtChat_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim i As Long
    Dim TheSite As String
    Dim pos As Long, Start As Long, finish As Long, OpenSite As Long
    Dim pointer As POINTAPI
    
    If Button = 1 Then
        pointer.x = x \ Screen.TwipsPerPixelX
        pointer.y = y \ Screen.TwipsPerPixelY
        Start = 1
        finish = Len(txtChat.Text) + 1
        pos = SendMessage(txtChat.hWnd, EM_CHARFROMPOS, 0&, pointer)

        For i = pos To 1 Step -1
            If Asc(Mid$(txtChat.Text, i, 1)) < 33 Then
                Start = i + 1
                Exit For
            End If
        Next i
    
        For i = Start To Len(txtChat.Text)
            If Asc(Mid$(txtChat.Text, i, 1)) < 33 Then
                finish = i
                Exit For
            End If
        Next i
 
        TheSite = Mid$(txtChat.Text, Start, finish - Start)
        
        If Mid$(TheSite, 1, 7) = "http://" Or Mid$(TheSite, 1, 4) = "www." Then
            OpenSite = ShellExecute(0, "open", TheSite, 0, 0, 1)
            Exit Sub
        End If
    End If
End Sub

Private Sub picInventory_Click()
    picInventory.Visible = True
End Sub

Private Sub picInventory3_LostFocus()
    itmDesc.Visible = False
End Sub

Private Sub lblUseItem_Click()
    Call UseItem
End Sub

Private Sub lblDropItem_Click()
    Call DropItem
End Sub

Private Sub lblCast_Click()
    Dim SpellSlot As Integer, SpellNum As Integer
    
    SpellSlot = lstSpells.ListIndex + 1
    SpellNum = Player(MyIndex).Spell(SpellSlot)

    If Player(MyIndex).InBattle = True And IsPlayerTurn = True Then
        If CanUseSpecial = False Then
            Exit Sub
        Else
            If SpellNum > 0 Then
                If SpellNum <> 43 And SpellNum <> 45 Then
                    Call SendUseTurnBasedSpecial(SpellSlot)
                    Exit Sub
                Else
                    Exit Sub
                End If
            End If
        End If
    End If
    
    ' Can't use special attacks while jumping
    If Player(MyIndex).Jumping = True Then
        Exit Sub
    End If
    
    If SpellNum > 0 Then
        If GetTickCount > Player(MyIndex).AttackTimer + 1000 Then
            Call SendData(CPackets.Ccast & SEP_CHAR & SpellSlot & END_CHAR)
            Player(MyIndex).Attacking = 1
            Player(MyIndex).AttackTimer = GetTickCount
            Player(MyIndex).CastedSpell = Yes
        End If
    Else
        Call AddText("There's no special attack here!", WHITE)
    End If
End Sub

Private Sub cmdAccess_Click()
    Call SendChangeGuildAccess(txtName.Text, Val(txtAccess.Text))
End Sub

Private Sub cmdDisown_Click()
    Call SendGuildDisown(txtName.Text)
End Sub

Private Sub cmdTrainee_Click()
    Call SendGuildMemberRequest(txtName.Text, 1)
End Sub

Private Sub picUp_Click()
    If InventorySlotsIndex >= 5 Then
        InventorySlotsIndex = InventorySlotsIndex - 4
    End If
End Sub

Private Sub picDown_Click()
    If (InventorySlotsIndex + 19) < Player(MyIndex).MaxInv Then
        InventorySlotsIndex = InventorySlotsIndex + 4
    End If
End Sub

Private Sub lstSpells_GotFocus()
    On Error Resume Next
    picScreen.SetFocus
End Sub
