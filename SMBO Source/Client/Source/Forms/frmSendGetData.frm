VERSION 5.00
Begin VB.Form frmSendGetData 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   ClientHeight    =   2490
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   6570
   ControlBox      =   0   'False
   Icon            =   "frmSendGetData.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSendGetData.frx":0FC2
   ScaleHeight     =   166
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   438
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblQuit 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6240
      TabIndex        =   1
      Top             =   0
      Width           =   135
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Loading Game..."
      BeginProperty Font 
         Name            =   "Porky's"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   1320
      TabIndex        =   0
      Top             =   1200
      Width           =   3945
   End
End
Attribute VB_Name = "frmSendGetData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyEscape) Then
        Call GameDestroy
    End If
End Sub

Private Sub lblQuit_Click()
    Call GameDestroy
End Sub

