VERSION 5.00
Begin VB.Form frmChars 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Character Selection"
   ClientHeight    =   5985
   ClientLeft      =   165
   ClientTop       =   330
   ClientWidth     =   5985
   ControlBox      =   0   'False
   Icon            =   "frmChars.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmChars.frx":0FC2
   ScaleHeight     =   399
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   399
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstChars 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1830
      ItemData        =   "frmChars.frx":75E54
      Left            =   1620
      List            =   "frmChars.frx":75E56
      TabIndex        =   0
      Top             =   1320
      Width           =   2670
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      Height          =   375
      Left            =   1395
      TabIndex        =   5
      Top             =   4980
      Visible         =   0   'False
      Width           =   3180
   End
   Begin VB.Label picUseChar 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   1815
      TabIndex        =   4
      Top             =   3600
      Width           =   2280
   End
   Begin VB.Label picNewChar 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   1815
      TabIndex        =   3
      Top             =   4035
      Width           =   2280
   End
   Begin VB.Label picCancel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1335
      TabIndex        =   2
      Top             =   5460
      Width           =   3330
   End
   Begin VB.Label picDelChar 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   1815
      TabIndex        =   1
      Top             =   4470
      Width           =   2280
   End
End
Attribute VB_Name = "frmChars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Label1_Click()
   
    AutoLogin = ReadINI("CONFIG", "Auto", (App.Path & "\config.ini"))
    If AutoLogin = 1 Then
        Call WriteINI("CONFIG", "Auto", 0, (App.Path & "\config.ini"))
        Call TcpDestroy
        frmLogin.Check2 = 0
        Me.Visible = False
        frmMainMenu.Visible = True
    End If
End Sub

Private Sub picCancel_Click()
    Call TcpDestroy

    Me.Visible = False
    frmLogin.Visible = True
End Sub

Private Sub picNewChar_Click()
    If lstChars.List(lstChars.ListIndex) <> "Free Character Slot" Then
        MsgBox "There is already a character in this slot!"
        Exit Sub
    End If
    
    Call InitLoadPicture(App.Path & "\GFX\Sprites.smbo", frmNewChar.picSprites)

    Call MenuState(MENU_STATE_NEWCHAR)
End Sub

Private Sub picUseChar_Click()
    If lstChars.List(lstChars.ListIndex) = "Free Character Slot" Then
        MsgBox "There is no character in this slot!"
        Exit Sub
    End If
    
    Call InitLoadPicture(App.Path & "\GFX\Items.smbo", frmMirage.picItems)
    Call InitLoadPicture(App.Path & "\GFX\Sprites.smbo", frmSpriteChange.picSprites)

    Call MenuState(MENU_STATE_USECHAR)
End Sub

Private Sub picDelChar_Click()
    Dim Value As Integer

    If lstChars.List(lstChars.ListIndex) = "Free Character Slot" Then
        MsgBox "There is no character in this slot!"
        Exit Sub
    End If

    Value = MsgBox("Are you sure you wish to delete this character?", vbYesNo, "Super Mario Bros. Online")
    If Value = vbYes Then
        Call MenuState(MENU_STATE_DELCHAR)
    End If
End Sub
