VERSION 5.00
Object = "{96366485-4AD2-4BC8-AFBF-B1FC132616A5}#2.0#0"; "VBMP.ocx"
Begin VB.Form frmMainMenu 
   Appearance      =   0  'Flat
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Main Menu"
   ClientHeight    =   5985
   ClientLeft      =   225
   ClientTop       =   435
   ClientWidth     =   8940
   ControlBox      =   0   'False
   FillColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMainMenu.frx":0000
   ScaleHeight     =   5985
   ScaleWidth      =   8940
   StartUpPosition =   2  'CenterScreen
   Begin VBMP.VBMPlayer MenuMusic 
      Height          =   1095
      Left            =   6840
      Top             =   120
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1931
   End
   Begin VB.Timer Status 
      Interval        =   2000
      Left            =   6360
      Top             =   120
   End
   Begin VB.Label lblPlayers 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Left            =   240
      TabIndex        =   10
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label lblOnline 
      BackStyle       =   0  'Transparent
      Caption         =   "Checking..."
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
      Left            =   4560
      TabIndex        =   9
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label lblVersion 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label picNews 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Receiving News..."
      BeginProperty Font 
         Name            =   "Porky's"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1815
      Left            =   3330
      TabIndex        =   7
      Top             =   2520
      Width           =   2295
   End
   Begin VB.Label picAutoLogin 
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
      ForeColor       =   &H00000000&
      Height          =   975
      Left            =   2160
      TabIndex        =   6
      Top             =   4800
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Server Status:"
      BeginProperty Font 
         Name            =   "Porky's"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3240
      TabIndex        =   5
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label picIpConfig 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      Height          =   510
      Left            =   5980
      TabIndex        =   4
      Top             =   2750
      Width           =   1420
   End
   Begin VB.Label picLogin 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      Height          =   510
      Left            =   1500
      TabIndex        =   3
      Top             =   2750
      Width           =   1420
   End
   Begin VB.Label picNewAccount 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      Height          =   510
      Left            =   1500
      TabIndex        =   2
      Top             =   3430
      Width           =   1420
   End
   Begin VB.Label picDeleteAccount 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      Height          =   510
      Left            =   5980
      TabIndex        =   1
      Top             =   3430
      Width           =   1420
   End
   Begin VB.Label picQuit 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      Height          =   510
      Left            =   7250
      TabIndex        =   0
      Top             =   5320
      Width           =   1420
   End
End
Attribute VB_Name = "frmMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim Ending As String

    Ending = ReadINI("CONFIG", "MenuMusic", App.Path & "\config.ini")
    If LenB(Ending) <> 0 Then
        Call MenuMusic.PlayMedia(App.Path & "\Music\" & Ending, True)
    End If

    Call MainMenuInit
End Sub

Private Sub Form_GotFocus()
    If frmMirage.Socket.State = 0 Then
        frmMirage.Socket.Connect
    End If
End Sub

Private Sub picAutoLogin_Click()
    If ConnectToServer = False Or (ConnectToServer = True And AutoLogin = 1 And AllDataReceived) Then
        Call MenuState(MENU_STATE_AUTO_LOGIN)
    End If
End Sub

'Private Sub picIpConfig_Click()
    'Me.Visible = False
    'frmIpconfig.Visible = True
'End Sub

Private Sub picNewAccount_Click()
    Me.Visible = False
    frmNewAccount.Visible = True
End Sub

Private Sub picDeleteAccount_Click()
    frmDeleteAccount.Visible = True
    Me.Visible = False
End Sub

Private Sub picLogin_Click()
    If ReadINI("CONFIG", "Auto", App.Path & "\config.ini") = 0 Then
        If LenB(frmLogin.txtPassword.Text) <> 0 Then
            frmLogin.Check1.Value = Checked
        Else
            frmLogin.Check1.Value = Unchecked
        End If
        frmLogin.Visible = True
        Me.Visible = False
    Else
        If AllDataReceived Then
            If LenB(frmLogin.txtName.Text) < 6 Then
                Call MsgBox("Your username must be at least three characters in length.")
                Exit Sub
            End If
    
            If LenB(frmLogin.txtPassword.Text) < 6 Then
                Call MsgBox("Your password must be at least three characters in length.")
                Exit Sub
            End If

            Call WriteINI("CONFIG", "Account", frmLogin.txtName.Text, (App.Path & "\config.ini"))
            Call MenuState(MENU_STATE_LOGIN)
            Me.Visible = False
        End If
    End If
End Sub

Private Sub picQuit_Click()
    Call GameDestroy
End Sub

Private Sub Status_Timer()
    If ConnectToServer = True Then
        If Not AllDataReceived Then
            Call SendData(CPackets.Cgivemethemax & END_CHAR)
        Else
            lblOnline.Caption = "Online"
            lblOnline.ForeColor = vbGreen
        End If
    
        Call SendData(CPackets.Cgetonline & END_CHAR)
    Else
        picNews.Caption = "Could not connect. The server may be down."

        lblOnline.Caption = "Offline"
        lblOnline.ForeColor = vbRed
    End If
End Sub
