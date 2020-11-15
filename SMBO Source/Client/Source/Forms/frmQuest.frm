VERSION 5.00
Begin VB.Form frmQuest 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2895
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4035
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmQuest.frx":0000
   ScaleHeight     =   2895
   ScaleWidth      =   4035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Continue 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Porky's"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   3
      Top             =   1920
      Width           =   500
   End
   Begin VB.Label QuestMsg 
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
      ForeColor       =   &H00C00000&
      Height          =   1455
      Left            =   195
      TabIndex        =   2
      Top             =   1320
      Width           =   2535
   End
   Begin VB.Label QuestProgress 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Porky's"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   3495
   End
   Begin VB.Label QuestTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Porky's"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "frmQuest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private NextText As String
Private TalkDone As Boolean

Public Sub FavorStart(ByVal Title As String, ByVal Progress As String, ByVal Msg As String, ByVal Msg2 As String)
    Call frmQuest.Show(vbModeless, frmMirage)
    QuestTitle.Caption = Title
    QuestProgress.Caption = Progress
    QuestMsg.Caption = Msg
    
    If Msg2 = vbNullString Then
        TalkDone = True
    Else
        TalkDone = False
        NextText = Msg2
    End If
End Sub

Private Sub Continue_Click()
    If TalkDone = True Then
        Unload Me
    Else
        QuestMsg.Caption = NextText
        TalkDone = True
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call Continue_Click
    End If
End Sub
