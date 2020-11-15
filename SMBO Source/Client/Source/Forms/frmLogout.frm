VERSION 5.00
Begin VB.Form frmLogout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Logout?"
   ClientHeight    =   2235
   ClientLeft      =   4380
   ClientTop       =   4140
   ClientWidth     =   2925
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLogout.frx":0000
   ScaleHeight     =   2235
   ScaleWidth      =   2925
   Begin VB.CommandButton Command3 
      Caption         =   "No"
      Height          =   495
      Left            =   1800
      TabIndex        =   2
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Yes"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Are you sure you want to log off?"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   675
      Width           =   2415
   End
End
Attribute VB_Name = "frmLogout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
    InGame = False
End Sub

Private Sub Command3_Click()
    Unload Me
End Sub
