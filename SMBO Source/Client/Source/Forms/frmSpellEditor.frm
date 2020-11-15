VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmSpellEditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Spell Editor"
   ClientHeight    =   5655
   ClientLeft      =   195
   ClientTop       =   525
   ClientWidth     =   6375
   ControlBox      =   0   'False
   DrawMode        =   14  'Copy Pen
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSpellEditor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   377
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Height          =   300
      Left            =   4080
      TabIndex        =   45
      Top             =   5280
      Width           =   2190
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Save Spell"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   44
      Top             =   5280
      Width           =   2175
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   120
      Top             =   120
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   8916
      _Version        =   393216
      TabHeight       =   397
      TabMaxWidth     =   3545
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Information"
      TabPicture(0)   =   "frmSpellEditor.frx":0FC2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblSound"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label8"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblVitalMod"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblRange"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblElement"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label9"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "info"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "scrlVitalMod"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "scrlRange"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "scrlElement"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "chkArea"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cmdSound"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "chkSelfSpell"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "fraChooseStat"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).ControlCount=   16
      TabCaption(1)   =   "Animation"
      TabPicture(1)   =   "frmSpellEditor.frx":0FDE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblSpellDone"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lblSpellTime"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lblSpellAnim"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Picture1"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Frame3"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "chkBig"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "picSpell"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Command1"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "scrlSpellDone"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "scrlSpellTime"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "scrlSpellAnim"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).ControlCount=   11
      TabCaption(2)   =   "Requirements"
      TabPicture(2)   =   "frmSpellEditor.frx":0FFA
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame1"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Frame2"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      Begin VB.Frame fraChooseStat 
         Caption         =   "Choose Stat"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   1680
         TabIndex        =   47
         Top             =   2040
         Visible         =   0   'False
         Width           =   2535
         Begin VB.HScrollBar scrlTime 
            Height          =   255
            Left            =   155
            Max             =   100
            Min             =   1
            TabIndex        =   57
            Top             =   1740
            Value           =   1
            Width           =   2175
         End
         Begin VB.HScrollBar scrlMult2 
            Height          =   255
            Left            =   1320
            Max             =   100
            TabIndex        =   54
            Top             =   1080
            Value           =   1
            Width           =   975
         End
         Begin VB.HScrollBar scrlMult1 
            Height          =   255
            Left            =   120
            Max             =   100
            TabIndex        =   53
            Top             =   1080
            Value           =   1
            Width           =   975
         End
         Begin VB.CommandButton cmdStatCancel 
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
            Height          =   285
            Left            =   1440
            TabIndex        =   51
            Top             =   2160
            Width           =   615
         End
         Begin VB.CommandButton cmdStatOk 
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
            Height          =   285
            Left            =   360
            TabIndex        =   50
            Top             =   2160
            Width           =   615
         End
         Begin VB.ComboBox cmbStat 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            ItemData        =   "frmSpellEditor.frx":1016
            Left            =   120
            List            =   "frmSpellEditor.frx":1026
            Style           =   2  'Dropdown List
            TabIndex        =   48
            Top             =   480
            Width           =   2295
         End
         Begin VB.Label lblStatTime 
            BackStyle       =   0  'Transparent
            Caption         =   "Time:"
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
            Left            =   1260
            TabIndex        =   59
            Top             =   1440
            Width           =   1065
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Time:"
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
            Left            =   150
            TabIndex        =   58
            Top             =   1440
            Width           =   465
         End
         Begin VB.Label lblMultiplier 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Multiplier"
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
            Left            =   1560
            TabIndex        =   56
            Top             =   840
            Width           =   720
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Multiplier:"
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
            TabIndex        =   55
            Top             =   840
            Width           =   720
         End
         Begin VB.Label lblStat 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Stat"
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
            Left            =   960
            TabIndex        =   52
            Top             =   240
            Width           =   1440
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Stat:"
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
            Top             =   240
            Width           =   360
         End
      End
      Begin VB.CheckBox chkSelfSpell 
         Caption         =   "Self-Spell"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2640
         TabIndex        =   67
         Top             =   3550
         Width           =   1575
      End
      Begin VB.Frame Frame2 
         Caption         =   "Passive Effects"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   -74760
         TabIndex        =   60
         Top             =   2880
         Width           =   5175
         Begin VB.CheckBox chkPassive 
            Caption         =   "Enable Passive Effect"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   120
            TabIndex        =   66
            Top             =   1560
            Width           =   1695
         End
         Begin VB.HScrollBar scrlPassiveStat 
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            Max             =   500
            Min             =   -500
            TabIndex        =   63
            Top             =   1200
            Value           =   1
            Width           =   4815
         End
         Begin VB.ComboBox cmbPassiveStat 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            ItemData        =   "frmSpellEditor.frx":104A
            Left            =   120
            List            =   "frmSpellEditor.frx":1060
            Style           =   2  'Dropdown List
            TabIndex        =   61
            Top             =   480
            Width           =   4905
         End
         Begin VB.Label lblPassiveStatChange 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Stat"
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
            Left            =   3600
            TabIndex        =   65
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Increase/Decrease Stat By:"
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
            Left            =   120
            TabIndex        =   64
            Top             =   960
            Width           =   1710
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Stat to Modify:"
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
            Left            =   120
            TabIndex        =   62
            Top             =   240
            Width           =   945
         End
      End
      Begin VB.CommandButton cmdSound 
         Caption         =   "Choose Sound"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   46
         Top             =   3120
         Width           =   855
      End
      Begin VB.HScrollBar scrlSpellAnim 
         Height          =   270
         Left            =   -74760
         Max             =   2000
         TabIndex        =   40
         Top             =   720
         Value           =   1
         Width           =   3495
      End
      Begin VB.HScrollBar scrlSpellTime 
         Height          =   270
         Left            =   -74760
         Max             =   500
         Min             =   40
         TabIndex        =   39
         Top             =   3120
         Value           =   40
         Width           =   5655
      End
      Begin VB.HScrollBar scrlSpellDone 
         Height          =   270
         Left            =   -74760
         Max             =   10
         Min             =   1
         TabIndex        =   38
         Top             =   3840
         Value           =   1
         Width           =   5655
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Refresh"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   220
         Left            =   -70680
         TabIndex        =   37
         Top             =   2400
         Width           =   1605
      End
      Begin VB.PictureBox picSpell 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   -70120
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   35
         Top             =   1275
         Width           =   480
      End
      Begin VB.CheckBox chkBig 
         Caption         =   "Big Spell"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   -74760
         TabIndex        =   34
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Frame Frame3 
         Caption         =   "Icon"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   -74760
         TabIndex        =   27
         Top             =   1440
         Width           =   3735
         Begin VB.HScrollBar HScroll1 
            Height          =   255
            Left            =   840
            Max             =   100
            TabIndex        =   31
            Top             =   600
            Width           =   2775
         End
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   540
            Left            =   120
            ScaleHeight     =   34
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   34
            TabIndex        =   28
            Top             =   360
            Width           =   540
            Begin VB.PictureBox picEmoticon 
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   480
               Left            =   15
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   29
               Top             =   15
               Width           =   480
               Begin VB.PictureBox iconn 
                  AutoRedraw      =   -1  'True
                  AutoSize        =   -1  'True
                  BackColor       =   &H00000000&
                  BorderStyle     =   0  'None
                  Height          =   480
                  Left            =   0
                  ScaleHeight     =   32
                  ScaleMode       =   3  'Pixel
                  ScaleWidth      =   32
                  TabIndex        =   30
                  Top             =   0
                  Width           =   480
               End
            End
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Spell ID:"
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
            Left            =   840
            TabIndex        =   33
            Top             =   360
            Width           =   555
         End
         Begin VB.Label Label13 
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
            Left            =   1440
            TabIndex        =   32
            Top             =   360
            Width           =   315
         End
      End
      Begin VB.CheckBox chkArea 
         Caption         =   "Area Effect"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1320
         TabIndex        =   26
         Top             =   3550
         Width           =   1215
      End
      Begin VB.Frame Frame1 
         Caption         =   "Spell Requirements"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   -74760
         TabIndex        =   15
         Top             =   600
         Width           =   5175
         Begin VB.ComboBox cmbClassReq 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            ItemData        =   "frmSpellEditor.frx":108C
            Left            =   120
            List            =   "frmSpellEditor.frx":108E
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   480
            Width           =   4905
         End
         Begin VB.HScrollBar scrlCost 
            Height          =   270
            Left            =   120
            Max             =   1000
            TabIndex        =   17
            Top             =   1680
            Width           =   4935
         End
         Begin VB.HScrollBar scrlLevelReq 
            Height          =   270
            Left            =   120
            Max             =   500
            TabIndex        =   16
            Top             =   1080
            Value           =   1
            Width           =   4935
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Class Required"
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
            Left            =   120
            TabIndex        =   23
            Top             =   240
            Width           =   915
         End
         Begin VB.Label lblCost 
            AutoSize        =   -1  'True
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
            Height          =   165
            Left            =   840
            TabIndex        =   21
            Top             =   1440
            Width           =   75
         End
         Begin VB.Label lblLevelReq 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "God's Only Spell"
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
            Left            =   1200
            TabIndex        =   20
            Top             =   840
            Width           =   1050
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "MP Cost:"
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
            Left            =   120
            TabIndex        =   19
            Top             =   1440
            Width           =   585
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Level Required:"
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
            Left            =   120
            TabIndex        =   18
            Top             =   840
            Width           =   960
         End
      End
      Begin VB.HScrollBar scrlElement 
         Height          =   255
         Left            =   240
         Max             =   1000
         TabIndex        =   12
         Top             =   4560
         Value           =   1
         Width           =   5655
      End
      Begin VB.HScrollBar scrlRange 
         Height          =   270
         Left            =   240
         Max             =   30
         Min             =   1
         TabIndex        =   11
         Top             =   3840
         Value           =   1
         Width           =   5655
      End
      Begin VB.HScrollBar scrlVitalMod 
         Height          =   270
         Left            =   240
         Max             =   1000
         TabIndex        =   4
         Top             =   2400
         Width           =   5655
      End
      Begin VB.Frame info 
         Caption         =   "Information"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   240
         TabIndex        =   1
         Top             =   510
         Width           =   5655
         Begin VB.ComboBox cmbType 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            ItemData        =   "frmSpellEditor.frx":1090
            Left            =   120
            List            =   "frmSpellEditor.frx":10AC
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   1080
            Width           =   5355
         End
         Begin VB.TextBox txtName 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   2
            Top             =   480
            Width           =   5355
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Spell Type"
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
            TabIndex        =   24
            Top             =   840
            Width           =   825
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Spell Name"
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
            TabIndex        =   3
            Top             =   240
            Width           =   960
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1605
         Left            =   -70680
         ScaleHeight     =   105
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   105
         TabIndex        =   36
         Top             =   720
         Width           =   1605
      End
      Begin VB.Label lblSpellAnim 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Anim: 0"
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
         Left            =   -74760
         TabIndex        =   43
         Top             =   480
         Width           =   495
      End
      Begin VB.Label lblSpellTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Time: 40"
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
         Left            =   -74760
         TabIndex        =   42
         Top             =   2880
         Width           =   555
      End
      Begin VB.Label lblSpellDone 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cycle Animation 1 Time"
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
         Left            =   -74760
         TabIndex        =   41
         Top             =   3600
         Width           =   1515
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Element:"
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
         Left            =   255
         TabIndex        =   14
         Top             =   4320
         Width           =   555
      End
      Begin VB.Label lblElement 
         AutoSize        =   -1  'True
         Caption         =   "None"
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
         Left            =   840
         TabIndex        =   13
         Top             =   4320
         Width           =   1410
      End
      Begin VB.Label lblRange 
         AutoSize        =   -1  'True
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
         Height          =   165
         Left            =   720
         TabIndex        =   10
         Top             =   3600
         Width           =   75
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Range:"
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
         TabIndex        =   9
         Top             =   3600
         Width           =   780
      End
      Begin VB.Label lblVitalMod 
         AutoSize        =   -1  'True
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
         Height          =   165
         Left            =   960
         TabIndex        =   8
         Top             =   2160
         Width           =   75
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Vital Mod:"
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
         TabIndex        =   7
         Top             =   2160
         Width           =   660
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Sound:"
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
         TabIndex        =   6
         Top             =   2880
         Width           =   540
      End
      Begin VB.Label lblSound 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Sound"
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
         Left            =   720
         TabIndex        =   5
         Top             =   2880
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmSpellEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Done As Long
Private Time As Long
Private SpellVar As Long

Private Sub chkArea_Click()
    If chkArea.Value = Checked Then
        If chkSelfSpell.Value = Checked Then
            chkSelfSpell.Value = Unchecked
        End If
    End If
End Sub

Private Sub chkBig_Click()
    frmSpellEditor.ScaleMode = 3
    Done = 0
    SpellVar = 0
    picSpell.Refresh
    If chkBig.Value = Checked Then
        picSpell.Width = 1440
        picSpell.Height = 1440
        picSpell.Top = 800
        picSpell.Left = 4400
    Else
        picSpell.Width = 480
        picSpell.Height = 480
        picSpell.Top = 1275
        picSpell.Left = 4880
    End If
End Sub

Private Sub chkPassive_Click()
    If chkPassive.Value = Unchecked Then
        cmbPassiveStat.Enabled = False
        scrlPassiveStat.Enabled = False
    Else
        cmbPassiveStat.Enabled = True
        scrlPassiveStat.Enabled = True
    End If
End Sub

Private Sub chkSelfSpell_Click()
    If chkSelfSpell.Value = Checked Then
        If chkArea.Value = Checked Then
            chkArea.Value = Unchecked
        End If
    End If
End Sub

Private Sub cmbStat_Click()
    lblStat.Caption = cmbStat.List(cmbStat.ListIndex)
End Sub

Private Sub cmdSound_Click()
    frmSpellSound.Show vbModal
End Sub

Private Sub cmdStatCancel_Click()
    fraChooseStat.Visible = False
    scrlMult1.Value = Int(1)
    scrlMult2.Value = Int(0)
End Sub

Private Sub cmdStatOk_Click()
    fraChooseStat.Visible = False
End Sub

Private Sub Command1_Click()
    Done = 0
End Sub

Private Sub Form_Load()
    scrlElement.Max = MAX_ELEMENTS
    
    If scrlTime.Value <> 1 Then
        lblStatTime.Caption = Trim$(STR(scrlTime.Value) & " Seconds")
    Else
        lblStatTime.Caption = Trim$(STR(scrlTime.Value) & " Second")
    End If
    
    lblMultiplier.Caption = Trim$(STR(scrlMult1.Value) & "." & STR(scrlMult2.Value))
End Sub

Private Sub HScroll1_Change()
    Label13.Caption = STR(HScroll1.Value)
    frmSpellEditor.iconn.Top = (HScroll1.Value * 32) * -1
End Sub

Private Sub scrlCost_Change()
    lblCost.Caption = STR(scrlCost.Value)
End Sub

Private Sub scrlElement_Change()
    lblElement.Caption = Element(scrlElement.Value).Name
End Sub

Private Sub scrlLevelReq_Change()
    If STR(scrlLevelReq.Value) = 0 Then
        lblLevelReq.Caption = "God's Only Spell"
    Else
        lblLevelReq.Caption = STR(scrlLevelReq.Value)
    End If
End Sub

Private Sub scrlMult1_Change()
    lblMultiplier.Caption = Trim$(STR(scrlMult1.Value) & "." & STR(scrlMult2.Value))
End Sub

Private Sub scrlMult2_Change()
    lblMultiplier.Caption = Trim$(STR(scrlMult1.Value) & "." & STR(scrlMult2.Value))
End Sub

Private Sub scrlPassiveStat_Change()
    lblPassiveStatChange.Caption = scrlPassiveStat.Value
End Sub

Private Sub scrlRange_Change()
    lblRange.Caption = STR(scrlRange.Value)
End Sub

Private Sub scrlSpellAnim_Change()
    lblSpellAnim.Caption = "Anim: " & scrlSpellAnim.Value
    Done = 0
End Sub

Private Sub scrlSpellDone_Change()
    Dim String2 As String
    String2 = "Times"
    If scrlSpellDone.Value = 1 Then
        String2 = "Time"
    End If
    lblSpellDone.Caption = "Cycle Animation " & scrlSpellDone.Value & " " & String2
    Done = 0
End Sub

Private Sub scrlSpellTime_Change()
    lblSpellTime.Caption = "Time: " & scrlSpellTime.Value
    Done = 0
End Sub

Private Sub scrlTime_Change()
    If Int(scrlTime.Value) <> 1 Then
        lblStatTime.Caption = STR(scrlTime.Value) & " Seconds"
    Else
        lblStatTime.Caption = STR(scrlTime.Value) & " Second"
    End If
End Sub

Private Sub scrlVitalMod_Change()
    lblVitalMod.Caption = STR(scrlVitalMod.Value)
End Sub

Private Sub cmdOk_Click()
    Call SpellEditorOk
End Sub

Private Sub cmdCancel_Click()
    Call SpellEditorCancel
End Sub

Private Sub Timer1_Timer()
    Dim sRECT As RECT
    Dim dRECT As RECT
    Dim SpellDone As Long
    Dim SpellAnim As Long
    Dim SpellTime As Long

    SpellDone = scrlSpellDone.Value
    SpellAnim = scrlSpellAnim.Value
    SpellTime = scrlSpellTime.Value

    If chkBig.Value = Checked Then
        SpellAnim = scrlSpellAnim.Value * 3
    End If

    If SpellAnim <= 0 Then
        Exit Sub
    End If
    If Done = SpellDone Then
        Exit Sub
    End If
    If chkBig = Checked Then
        With dRECT
            .Top = 0
            .Bottom = PIC_Y + 64
            .Left = 0
            .Right = PIC_X + 64
        End With
    Else
        With dRECT
            .Top = 0
            .Bottom = PIC_Y
            .Left = 0
            .Right = PIC_X
        End With
    End If
    If chkBig.Value = Checked Then
        If SpellVar > 32 Then
            Done = Done + 1
            SpellVar = 0
        End If
        If GetTickCount > Time + SpellTime Then
            Time = GetTickCount
            SpellVar = SpellVar + 3
        End If
    Else
        If SpellVar > 10 Then
            Done = Done + 1
            SpellVar = 0
        End If
        If GetTickCount > Time + SpellTime Then
            Time = GetTickCount
            SpellVar = SpellVar + 1
        End If
    End If
    If chkBig = Checked Then
        If DD_BigSpellAnim Is Nothing Then
        Else
            With sRECT
                .Top = SpellAnim * PIC_Y
                .Bottom = .Top + (PIC_Y * 3)
                .Left = SpellVar * PIC_X
                .Right = .Left + (PIC_X * 3)
            End With

            Call DD_BigSpellAnim.BltToDC(picSpell.hDC, sRECT, dRECT)
            picSpell.Refresh
        End If
    Else
        If DD_SpellAnim Is Nothing Then
        Else
            With sRECT
                .Top = SpellAnim * PIC_Y
                .Bottom = .Top + PIC_Y
                .Left = SpellVar * PIC_X
                .Right = .Left + PIC_X
            End With

            Call DD_SpellAnim.BltToDC(picSpell.hDC, sRECT, dRECT)
            picSpell.Refresh
        End If
    End If
End Sub
Private Sub cmbType_Click()
    If (cmbType.ListIndex = SPELL_TYPE_SCRIPTED) Then
        Label4.Caption = "Script"
        Label14.Visible = True
        Label8.Visible = True
        lblSound.Visible = True
        cmdSound.Visible = True
        Label2.Visible = True
        lblRange.Visible = True
        scrlRange.Visible = True
        lblSpellAnim.Visible = True
        scrlSpellAnim.Visible = True
        lblSpellTime.Visible = True
        scrlSpellTime.Visible = True
        lblSpellDone.Visible = True
        scrlSpellDone.Visible = True
        chkArea.Visible = True
        Command1.Visible = True
        picSpell.Visible = True
        fraChooseStat.Visible = False
    ElseIf (cmbType.ListIndex = SPELL_TYPE_STATCHANGE) Then
        Label9.Visible = False
        Label4.Visible = False
        lblVitalMod.Visible = False
        scrlVitalMod.Visible = False
        Label8.Visible = True
        lblSound.Visible = True
        cmdSound.Visible = True
        Label2.Visible = True
        lblRange.Visible = True
        scrlRange.Visible = True
        lblSpellAnim.Visible = True
        scrlSpellAnim.Visible = True
        lblSpellTime.Visible = True
        scrlSpellTime.Visible = True
        lblSpellDone.Visible = True
        scrlSpellDone.Visible = True
        chkArea.Visible = False
        Command1.Visible = True
        picSpell.Visible = True
        scrlElement.Visible = False
        lblElement.Visible = False
        fraChooseStat.Visible = True
    Else
        lblElement.Visible = True
        Label4.Caption = "Vital Mod"
        Label4.Visible = True
        lblVitalMod.Visible = True
        Label14.Visible = True
        Label9.Visible = True
        scrlVitalMod.Visible = True
        Label8.Visible = True
        lblSound.Visible = True
        cmdSound.Visible = True
        Label2.Visible = True
        lblRange.Visible = True
        scrlRange.Visible = True
        lblSpellAnim.Visible = True
        scrlSpellAnim.Visible = True
        lblSpellTime.Visible = True
        scrlSpellTime.Visible = True
        lblSpellDone.Visible = True
        scrlSpellDone.Visible = True
        chkArea.Visible = True
        Command1.Visible = True
        picSpell.Visible = True
        fraChooseStat.Visible = False
        scrlElement.Visible = True
    End If
End Sub
