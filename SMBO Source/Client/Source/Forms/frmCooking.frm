VERSION 5.00
Begin VB.Form frmCooking 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cooking"
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4395
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCooking.frx":0000
   ScaleHeight     =   226
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   293
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picSecondItem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1845
      Left            =   1080
      Picture         =   "frmCooking.frx":31002
      ScaleHeight     =   121
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   143
      TabIndex        =   3
      Top             =   600
      Visible         =   0   'False
      Width           =   2175
      Begin VB.Label No 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Porky's"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   1215
         TabIndex        =   6
         Top             =   1200
         Width           =   630
      End
      Begin VB.Label Yes 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Porky's"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   195
         TabIndex        =   5
         Top             =   1200
         Width           =   660
      End
      Begin VB.Label lblSecondItem 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Would you like to cook a second item?"
         BeginProperty Font 
            Name            =   "Porky's"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   150
         TabIndex        =   4
         Top             =   135
         Width           =   1875
      End
   End
   Begin VB.ListBox lstInventory 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      ForeColor       =   &H00FF0000&
      Height          =   1785
      ItemData        =   "frmCooking.frx":3DC74
      Left            =   2160
      List            =   "frmCooking.frx":3DC76
      TabIndex        =   0
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label Cancel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Porky's"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   390
      TabIndex        =   2
      Top             =   2640
      Width           =   1275
   End
   Begin VB.Label Submit 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Porky's"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   390
      TabIndex        =   1
      Top             =   1875
      Width           =   1275
   End
End
Attribute VB_Name = "frmCooking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private FirstItemSlot As Long, ItemNum As Long
Private SecondItem As Boolean, HasOffered As Boolean

Private Sub Cancel_Click()
    If picSecondItem.Visible = True Then
        Exit Sub
    End If
    
    Unload Me
    
    Call frmNpcTalk.NpcTalk(CookNpcNum, "Oh, you don't want to cook anything? Come back next time!", vbNullString)
End Sub

Private Sub Form_Load()
    SecondItem = False
    HasOffered = False
End Sub

Private Sub No_Click()
    SecondItem = False
    picSecondItem.Visible = False
End Sub

Private Sub Submit_Click()
    If picSecondItem.Visible = True Then
        Exit Sub
    End If
    If lstInventory.ListIndex + 1 <= 0 Then
        Exit Sub
    End If
    If HasOffered = True Then
        If SecondItem = True Then
            If lstInventory.ListIndex + 1 = FirstItemSlot Then
                Call AddText("Your second item must be in a different inventory slot than the first item!", WHITE)
                Exit Sub
            End If
            ItemNum = GetPlayerInvItemNum(MyIndex, lstInventory.ListIndex + 1)
            If ItemNum = 0 Then
                Call AddText("There is no item in that slot to cook!", WHITE)
                Exit Sub
            ElseIf ItemNum = GetPlayerInvItemNum(MyIndex, FirstItemSlot) Then
                Call AddText("You cannot cook two of the same item!", BRIGHTRED)
                Exit Sub
            End If
            If Item(ItemNum).Cookable = False Then
                Call AddText("That's not an item you can cook!", BRIGHTRED)
                Exit Sub
            End If
            Call CookItem(FirstItemSlot, lstInventory.ListIndex + 1)
            Unload Me
        Else
            Call CookItem(FirstItemSlot)
            Unload Me
        End If
    Else
        FirstItemSlot = lstInventory.ListIndex + 1
        ItemNum = GetPlayerInvItemNum(MyIndex, FirstItemSlot)
        If ItemNum < 1 Then
            Call AddText("There is no item in that slot to cook!", WHITE)
            Exit Sub
        End If
        If Item(ItemNum).Cookable = True Then
            picSecondItem.Visible = True
            HasOffered = True
        Else
            Call AddText("That's not an item you can cook!", BRIGHTRED)
            Exit Sub
        End If
    End If
End Sub

Private Sub Yes_Click()
    SecondItem = True
    picSecondItem.Visible = False
End Sub
