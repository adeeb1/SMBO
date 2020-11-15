VERSION 5.00
Begin VB.Form frmCard 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6900
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   8175
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCard.frx":0000
   ScaleHeight     =   460
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   545
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picCard1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   1200
      ScaleHeight     =   71
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   82
      TabIndex        =   1
      Top             =   1080
      Width           =   1260
      Begin VB.PictureBox picCard 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   375
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   8
         Top             =   270
         Width           =   480
      End
   End
   Begin VB.ListBox lstCards 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   4350
      ItemData        =   "frmCard.frx":B7BF2
      Left            =   5280
      List            =   "frmCard.frx":B7BF9
      TabIndex        =   0
      Top             =   1080
      Width           =   2415
   End
   Begin VB.Label lblSpeed 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Speed:"
      BeginProperty Font 
         Name            =   "Porky's"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   600
      TabIndex        =   9
      Top             =   3720
      Width           =   2775
   End
   Begin VB.Label lblDesc 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Porky's"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   600
      TabIndex        =   7
      Top             =   4800
      Width           =   2655
   End
   Begin VB.Label lblAttack 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Attack:"
      BeginProperty Font 
         Name            =   "Porky's"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   600
      TabIndex        =   6
      Top             =   3000
      Width           =   2775
   End
   Begin VB.Label lblDefense 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Defense:"
      BeginProperty Font 
         Name            =   "Porky's"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   600
      TabIndex        =   5
      Top             =   3360
      Width           =   2775
   End
   Begin VB.Label lblExp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Exp earned:"
      BeginProperty Font 
         Name            =   "Porky's"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   600
      TabIndex        =   4
      Top             =   4080
      Width           =   2775
   End
   Begin VB.Label lblHP 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "HP:"
      BeginProperty Font 
         Name            =   "Porky's"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   600
      TabIndex        =   3
      Top             =   2640
      Width           =   2775
   End
   Begin VB.Label ReturnButton 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   5280
      TabIndex        =   2
      Top             =   6000
      Width           =   2415
   End
End
Attribute VB_Name = "frmCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub UpdateInfo(ByVal HP As Integer, ByVal Attack As Integer, ByVal Defense As Integer, ByVal speed As Integer, ByVal Exp As Integer, ByVal Description As String, ByVal ItemNum As Long)
    lblHP.Caption = "HP: " & HP
    lblAttack.Caption = "Attack: " & Attack
    lblDefense.Caption = "Defense: " & Defense
    lblSpeed.Caption = "Speed: " & speed
    lblExp.Caption = "Exp earned: " & Exp
    lblDesc.Caption = Description
    
    Call DrawCard(ItemNum)
End Sub

Private Sub DrawCard(ItemNum As Long)
    Dim srec As RECT, drec As RECT
    
    drec.Top = 0
    drec.Bottom = PIC_X
    drec.Left = 0
    drec.Right = PIC_Y
    srec.Top = (Item(ItemNum).Pic \ 6) * PIC_Y
    srec.Bottom = srec.Top + PIC_X
    srec.Left = (Item(ItemNum).Pic - (Item(ItemNum).Pic \ 6) * 6) * PIC_X
    srec.Right = srec.Left + PIC_Y
    
    Call DD_ItemSurf.BltToDC(picCard.hDC, srec, drec)
End Sub

Private Sub lstCards_Click()
    Dim i As Long
    
    If Trim$(lstCards.List(lstCards.ListIndex)) = "<Empty Card Slot>" Then
        lblHP.Caption = "HP: "
        lblAttack.Caption = "Attack: "
        lblDefense.Caption = "Defense: "
        lblSpeed.Caption = "Speed: "
        lblExp.Caption = "Exp earned: "
        lblDesc.Caption = vbNullString
        
        picCard.Picture = LoadPicture()
        Exit Sub
    End If
    
    For i = 94 To MAX_ITEMS
        If Trim$(Item(i).Name) = Trim$(lstCards.List(lstCards.ListIndex)) Then
            Call SendData(CPackets.Ccardshop & SEP_CHAR & i & END_CHAR)
            Exit Sub
        End If
    Next i
End Sub

Private Sub ReturnButton_Click()
    Unload Me
End Sub
