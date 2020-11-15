VERSION 5.00
Begin VB.Form frmNpcTalkYesNo 
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
   Picture         =   "frmNpcTalkYesNo.frx":0000
   ScaleHeight     =   3000
   ScaleWidth      =   3300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label lblNo 
      BackStyle       =   0  'Transparent
      Height          =   330
      Left            =   1440
      TabIndex        =   3
      Top             =   2390
      Width           =   600
   End
   Begin VB.Label lblYes 
      BackStyle       =   0  'Transparent
      Height          =   335
      Left            =   640
      TabIndex        =   2
      Top             =   2390
      Width           =   605
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
      Height          =   1420
      Left            =   270
      TabIndex        =   1
      Top             =   720
      Width           =   2200
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
      Height          =   495
      Left            =   135
      TabIndex        =   0
      Top             =   92
      Width           =   3015
   End
End
Attribute VB_Name = "frmNpcTalkYesNo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private StoredText(1 To 2) As String
Private RealNpcNum As Long
Private Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean

Public Sub NpcTalk(ByVal NpcNum As Long, ByVal Text As String, ByVal YesText As String, ByVal NoText As String)
    Call Me.Show(vbModeless, frmMirage)
    
    RealNpcNum = NpcNum
    
    NpcName.Caption = Trim$(Npc(NpcNum).Name)
    NpcText.Caption = Text
    
    StoredText(1) = YesText
    StoredText(2) = NoText
    
    ' Blit NPC graphics
    If Npc(NpcNum).Big = 0 Then ' 32 x 64
        Call TransparentBlt(Me.hDC, 183, 108, PIC_X, 64, frmNpcEditor.picSprites.hDC, 3 * PIC_X, Npc(NpcNum).Sprite * 64, PIC_X, 64, RGB(0, 0, 0))
    Else ' 64 x 64
        Call TransparentBlt(Me.hDC, 183, 108, 64, 64, frmNpcEditor.picSprites.hDC, 3 * PIC_X, Npc(NpcNum).Sprite * 64, 64, 64, RGB(0, 0, 0))
    End If
End Sub

Private Sub lblNo_Click()
    Unload Me
    
    If StoredText(2) <> vbNullString Then
        Call frmNpcTalk.NpcTalk(RealNpcNum, StoredText(2), vbNullString, False)
    Else
        Call EndNpcTalkToConditions(RealNpcNum, False)
    End If
End Sub

Private Sub lblYes_Click()
    Unload Me

    If StoredText(1) <> vbNullString Then
        Call frmNpcTalk.NpcTalk(RealNpcNum, StoredText(1), vbNullString, True)
    Else
        Unload Me
        Call EndNpcTalkToConditions(RealNpcNum, True)
    End If
End Sub


