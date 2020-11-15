VERSION 5.00
Begin VB.Form frmWarn 
   Caption         =   "Warn Player"
   ClientHeight    =   2895
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4320
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Picture         =   "frmWarn.frx":0000
   ScaleHeight     =   2895
   ScaleWidth      =   4320
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox ReasonTxt 
      Height          =   855
      Left            =   600
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   1240
      Width           =   3135
   End
   Begin VB.TextBox Username 
      Height          =   285
      Left            =   600
      MaxLength       =   20
      TabIndex        =   0
      Top             =   560
      Width           =   3135
   End
   Begin VB.Label Cancel 
      BackStyle       =   0  'Transparent
      Height          =   360
      Left            =   2860
      TabIndex        =   4
      Top             =   2400
      Width           =   790
   End
   Begin VB.Label RemoveWarn 
      BackStyle       =   0  'Transparent
      Height          =   550
      Left            =   1800
      TabIndex        =   3
      Top             =   2280
      Width           =   840
   End
   Begin VB.Label WarnPlayer 
      BackStyle       =   0  'Transparent
      Height          =   550
      Left            =   830
      TabIndex        =   2
      Top             =   2280
      Width           =   690
   End
End
Attribute VB_Name = "frmWarn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cancel_Click()
    Unload Me
End Sub

Private Sub RemoveWarn_Click()
    Dim PlayerName As String
    
    PlayerName = Username.Text
    
    Call SendData(CPackets.Cremovewarn & SEP_CHAR & Trim$(PlayerName) & END_CHAR)
    
    Username.Text = vbNullString
    ReasonTxt.Text = vbNullString
End Sub

Private Sub WarnPlayer_Click()
    Dim PlayerName As String
    Dim Reason As String
    
    PlayerName = Username.Text
      
    If ReasonTxt.Text <> vbNullString Then
        Reason = ReasonTxt.Text
    Else
        Reason = "No reason given."
    End If
    
    Call SendData(CPackets.Cwarn & SEP_CHAR & Trim$(PlayerName) & SEP_CHAR & Trim$(Reason) & END_CHAR)
    
    Username.Text = vbNullString
    ReasonTxt.Text = vbNullString
End Sub
