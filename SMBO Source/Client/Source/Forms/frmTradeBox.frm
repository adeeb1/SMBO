VERSION 5.00
Begin VB.Form frmTradeBox 
   Caption         =   "Trade Offer"
   ClientHeight    =   1860
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2985
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Picture         =   "frmTradeBox.frx":0000
   ScaleHeight     =   1860
   ScaleWidth      =   2985
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton DeclineBox 
      Caption         =   "Decline"
      Height          =   615
      Left            =   1680
      TabIndex        =   1
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton AcceptBox 
      Caption         =   "Accept"
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   240
      TabIndex        =   2
      Top             =   350
      Width           =   1695
   End
End
Attribute VB_Name = "frmTradeBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AcceptBox_Click()
    Call SendAcceptTrade
End Sub

Private Sub DeclineBox_Click()
    Call SendDeclineTrade
End Sub
