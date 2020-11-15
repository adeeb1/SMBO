VERSION 5.00
Begin VB.Form frmIpconfig 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configure Server IP"
   ClientHeight    =   5985
   ClientLeft      =   165
   ClientTop       =   405
   ClientWidth     =   5985
   ControlBox      =   0   'False
   Icon            =   "frmIpconfig.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmIpconfig.frx":0FC2
   ScaleHeight     =   399
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   399
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtPort 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   3840
      TabIndex        =   1
      Top             =   2040
      Width           =   1515
   End
   Begin VB.TextBox TxtIP 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   540
      TabIndex        =   0
      Top             =   2040
      Width           =   1515
   End
   Begin VB.Label PicCancel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   1800
      TabIndex        =   3
      Top             =   4935
      Width           =   2655
   End
   Begin VB.Label PicConfirm 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   2760
      TabIndex        =   2
      Top             =   4005
      Width           =   780
   End
End
Attribute VB_Name = "frmIpconfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    TxtIP = ReadINI("IPCONFIG", "IP", App.Path & "\config.ini")
    TxtPort = ReadINI("IPCONFIG", "PORT", App.Path & "\config.ini")
    TxtIP.Text = ReadINI("IPCONFIG", "IP", App.Path & "\config.ini")
    TxtPort.Text = ReadINI("IPCONFIG", "PORT", App.Path & "\config.ini")
End Sub

Private Sub picCancel_Click()
    frmMainMenu.Visible = True
    frmIpconfig.Visible = False
End Sub

Private Sub picConfirm_Click()
    Dim IP As String, Port As String

    IP = Trim$(TxtIP)
    Port = Val(TxtPort)

    If Len(IP) = 0 Then
        Call MsgBox("You've entered an invalid IP Address!", vbCritical, "Super Mario Bros. Online")
        Exit Sub
    End If
    If Port <= 0 Then
        Call MsgBox("You've entered an invalid Port number!", vbCritical, "Super Mario Bros. Online")
        Exit Sub
    End If
    
    Call WriteINI("IPCONFIG", "IP", TxtIP.Text, (App.Path & "\config.ini"))
    Call WriteINI("IPCONFIG", "PORT", TxtPort.Text, (App.Path & "\config.ini"))
    
    ' Call MenuState(MENU_STATE_IPCONFIG)
    Call TcpDestroy
    frmMirage.Socket.RemoteHost = TxtIP.Text
    frmMirage.Socket.RemotePort = TxtPort.Text
    frmMainMenu.Visible = True
    frmIpconfig.Visible = False
End Sub

