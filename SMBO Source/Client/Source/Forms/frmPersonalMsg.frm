VERSION 5.00
Begin VB.Form frmPersonalMsg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Send Someone A Message"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4605
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmPersonalMsg.frx":0000
   ScaleHeight     =   3210
   ScaleWidth      =   4605
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox EnterMsg 
      Height          =   1215
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   1440
      Width           =   3855
   End
   Begin VB.TextBox EnterUsername 
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   240
      Width           =   2295
   End
   Begin VB.Label Cancel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Porky's"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   3000
      TabIndex        =   4
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label SendMsg 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Send Message"
      BeginProperty Font 
         Name            =   "Porky's"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label Message 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Message"
      BeginProperty Font 
         Name            =   "Porky's"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   1050
      Width           =   1575
   End
   Begin VB.Label Username 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Player Username"
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
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   300
      Width           =   1575
   End
End
Attribute VB_Name = "frmPersonalMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cancel_Click()
    Unload Me
End Sub

Private Sub SendMsg_Click()
Dim Name As String
Dim Message As String

Name = EnterUsername.Text
Message = EnterMsg.Text

' Sends message if there is one there.

If Message <> vbNullString And Name <> vbNullString Then

Call OtherMsg(Message, Name)

    If KeepUsername = False Then
        EnterUsername.Text = vbNullString
        EnterUsername.SetFocus
    Else
        EnterMsg.SetFocus
    End If

    EnterMsg.Text = vbNullString

ElseIf Message = vbNullString Then
   Call AddText("You must enter a message!", BRIGHTRED)

ElseIf Name = vbNullString Then
   Call AddText("You must enter a username!", BRIGHTRED)
End If

End Sub
