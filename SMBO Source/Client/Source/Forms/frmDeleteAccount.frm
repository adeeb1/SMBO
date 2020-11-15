VERSION 5.00
Begin VB.Form frmDeleteAccount 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Delete Account"
   ClientHeight    =   5985
   ClientLeft      =   195
   ClientTop       =   345
   ClientWidth     =   5985
   ControlBox      =   0   'False
   Icon            =   "frmDeleteAccount.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmDeleteAccount.frx":0FC2
   ScaleHeight     =   399
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   399
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPassword 
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
      Height          =   270
      IMEMode         =   3  'DISABLE
      Left            =   3645
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   3420
      Width           =   2115
   End
   Begin VB.TextBox txtName 
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
      Height          =   270
      Left            =   3660
      MaxLength       =   20
      TabIndex        =   0
      Top             =   1530
      Width           =   2100
   End
   Begin VB.Label picCancel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   120
      TabIndex        =   3
      Top             =   5550
      Width           =   975
   End
   Begin VB.Label picConnect 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2910
      TabIndex        =   2
      Top             =   5115
      Width           =   2250
   End
End
Attribute VB_Name = "frmDeleteAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub picCancel_Click()
    frmDeleteAccount.Visible = False
    frmMainMenu.Visible = True
End Sub

Private Sub picConnect_Click()
    Dim Answer As Long

    If LenB(txtName.Text) < 6 Then
        MsgBox ("Your username must be at least three characters in length.")
        Exit Sub
    End If

    If LenB(txtPassword.Text) < 6 Then
        MsgBox ("Your password must be at least three characters in length.")
        Exit Sub
    End If

    Answer = MsgBox("Are you sure you want to delete your account?", vbYesNo, "Super Mario Bros. Online")
    If Answer = vbYes Then
        Call MenuState(MENU_STATE_DELACCOUNT)
    End If
End Sub
