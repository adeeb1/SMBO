VERSION 5.00
Begin VB.Form frmGroupMember 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Group Member Request"
   ClientHeight    =   2100
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3315
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmGroupMember.frx":0000
   ScaleHeight     =   140
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   221
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label lblDecline 
      BackStyle       =   0  'Transparent
      Height          =   510
      Left            =   1935
      TabIndex        =   2
      Top             =   1485
      Width           =   975
   End
   Begin VB.Label lblAccept 
      BackStyle       =   0  'Transparent
      Height          =   510
      Left            =   420
      TabIndex        =   1
      Top             =   1470
      Width           =   975
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Porky's"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   705
      Width           =   2895
   End
End
Attribute VB_Name = "frmGroupMember"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private PlayerName As String
Private Trainee As Integer

Public Sub SetGuildRequesterName(ByVal Name As String, ByVal IsTrainee As Integer)
    PlayerName = Name
    Trainee = IsTrainee
End Sub

Private Sub lblAccept_Click()
    If Trainee = 1 Then
        Call SendSetTrainee(PlayerName)
    Else
        Call SendGuildMember(PlayerName)
    End If
    Unload Me
End Sub

Private Sub lblDecline_Click()
    Call SendGuildMemberDecline(PlayerName)
    Call AddText("You have declined " & PlayerName & "'s Group invitation.", BRIGHTRED)
    Unload Me
End Sub
