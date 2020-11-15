VERSION 5.00
Begin VB.Form frmWelcome 
   BorderStyle     =   0  'None
   ClientHeight    =   3000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   Picture         =   "frmWelcome.frx":0000
   ScaleHeight     =   3000
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label Okbutton 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Porky's"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   3650
      TabIndex        =   2
      Top             =   1500
      Width           =   735
   End
   Begin VB.Label Welcometxt 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmWelcome.frx":2DB82
      BeginProperty Font 
         Name            =   "Porky's"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   240
      TabIndex        =   1
      Top             =   1150
      Width           =   3015
   End
   Begin VB.Label Titletxt 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome to Super Mario Bros. Online!"
      BeginProperty Font 
         Name            =   "Porky's"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "frmWelcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private IsStartGame As Boolean

Public Sub ShowWelcomeMsg(ByVal TitleText As String, ByVal MessageText As String)
    Titletxt.Caption = TitleText
    Welcometxt.Caption = MessageText
    
    Call Me.Show(vbModeless, frmMirage)
End Sub

Private Sub Okbutton_Click()
    Call PlaySound("yi_messageBoxDisappear.wav")
    Unload Me
End Sub
