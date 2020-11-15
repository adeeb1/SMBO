VERSION 5.00
Begin VB.Form frmNpcTalk 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3000
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   3300
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmNpcTalk.frx":0000
   ScaleHeight     =   3000
   ScaleWidth      =   3300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label OkNext 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   2680
      TabIndex        =   2
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label NpcText 
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
      Height          =   1450
      Left            =   240
      TabIndex        =   1
      Top             =   1040
      Width           =   2275
   End
   Begin VB.Label NpcName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Porky's"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   135
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "frmNpcTalk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private TextNext2 As String
Private TalkDone As Boolean, WillAct As Boolean
Private RealNpcNum As Long
Private Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean

Public Sub NpcTalk(ByVal NpcNum As Long, ByVal Text1 As String, ByVal Text2 As String, Optional ByVal ShouldAct As Boolean = False)
    Call Me.Show(vbModeless, frmMirage)
    
    RealNpcNum = NpcNum
    WillAct = ShouldAct
    
    NpcName.Caption = Trim$(Npc(NpcNum).Name)
    NpcText.Caption = Text1
    
    If Text2 = vbNullString Then
        TalkDone = True
    Else
        TalkDone = False
        TextNext2 = Text2
    End If
    
    ' Blit NPC graphics
    If Npc(NpcNum).Big = 0 Then ' 32 x 64
        Call TransparentBlt(Me.hDC, 183, 108, PIC_X, 64, frmNpcEditor.picSprites.hDC, 3 * PIC_X, Npc(NpcNum).Sprite * 64, PIC_X, 64, RGB(0, 0, 0))
    Else ' 64 x 64
        Call TransparentBlt(Me.hDC, 183, 108, 64, 64, frmNpcEditor.picSprites.hDC, 3 * PIC_X, Npc(NpcNum).Sprite * 64, 64, 64, RGB(0, 0, 0))
    End If
End Sub

Private Sub OkNext_Click()
    If TalkDone = True Then
        If IsChefBeanB = True Then
            IsChefBeanB = False
            Unload Me
            
            Call ShowCookForm(224)
        Else
            Unload Me
            
            If WillAct = True Then
                Call EndNpcTalkConditions(RealNpcNum)
            End If
        End If
    Else
        NpcText.Caption = TextNext2
        TalkDone = True
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call OkNext_Click
    End If
End Sub
