VERSION 5.00
Begin VB.Form frmLevelUp 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5400
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   3750
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLevelUp.frx":0000
   ScaleHeight     =   5400
   ScaleWidth      =   3750
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label lblStatPoint 
      BackStyle       =   0  'Transparent
      Height          =   725
      Index           =   5
      Left            =   160
      TabIndex        =   5
      Top             =   1570
      Width           =   705
   End
   Begin VB.Label lblStatPoint 
      BackStyle       =   0  'Transparent
      Height          =   725
      Index           =   4
      Left            =   160
      TabIndex        =   4
      Top             =   840
      Width           =   705
   End
   Begin VB.Label lblStatPoint 
      BackStyle       =   0  'Transparent
      Height          =   725
      Index           =   3
      Left            =   160
      TabIndex        =   3
      Top             =   3780
      Width           =   705
   End
   Begin VB.Label lblStatPoint 
      BackStyle       =   0  'Transparent
      Height          =   725
      Index           =   2
      Left            =   160
      TabIndex        =   2
      Top             =   4520
      Width           =   705
   End
   Begin VB.Label lblStatPoint 
      BackStyle       =   0  'Transparent
      Height          =   725
      Index           =   1
      Left            =   160
      TabIndex        =   1
      Top             =   3040
      Width           =   705
   End
   Begin VB.Label lblStatPoint 
      BackStyle       =   0  'Transparent
      Height          =   725
      Index           =   0
      Left            =   160
      TabIndex        =   0
      Top             =   2310
      Width           =   705
   End
End
Attribute VB_Name = "frmLevelUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub lblStatPoint_Click(Index As Integer)
    Call SendData(CPackets.Cusestatpoint & SEP_CHAR & Index & END_CHAR)
End Sub
