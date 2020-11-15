VERSION 5.00
Begin VB.Form frmRecipeEditor 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recipe Editor"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6975
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   314
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   465
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Cancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      TabIndex        =   9
      Top             =   3720
      Width           =   975
   End
   Begin VB.CommandButton SaveRecipe 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   8
      Top             =   3720
      Width           =   975
   End
   Begin VB.Frame fraResult 
      Caption         =   "Resulting Item"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2228
      Left            =   3600
      TabIndex        =   5
      Top             =   120
      Width           =   3255
      Begin VB.ListBox lstResultItem 
         Appearance      =   0  'Flat
         Height          =   1590
         ItemData        =   "frmRecipeEditor.frx":0000
         Left            =   1200
         List            =   "frmRecipeEditor.frx":0002
         TabIndex        =   7
         Top             =   360
         Width           =   1815
      End
      Begin VB.PictureBox picResultItem 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   240
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   6
         Top             =   840
         Width           =   480
      End
   End
   Begin VB.Frame fraIngredient 
      Caption         =   "Ingredients"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3255
      Begin VB.PictureBox picIngredient2 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   240
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   4
         Top             =   3000
         Width           =   480
      End
      Begin VB.ListBox lstIngredient2 
         Appearance      =   0  'Flat
         Height          =   1590
         ItemData        =   "frmRecipeEditor.frx":0004
         Left            =   1200
         List            =   "frmRecipeEditor.frx":0006
         TabIndex        =   3
         Top             =   2520
         Width           =   1815
      End
      Begin VB.PictureBox picIngredient1 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   240
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   2
         Top             =   840
         Width           =   480
      End
      Begin VB.ListBox lstIngredient1 
         Appearance      =   0  'Flat
         Height          =   1590
         ItemData        =   "frmRecipeEditor.frx":0008
         Left            =   1200
         List            =   "frmRecipeEditor.frx":000A
         TabIndex        =   1
         Top             =   360
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmRecipeEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cancel_Click()
    Call RecipeEditorCancel
End Sub

Private Sub DrawIngredient1(ByVal ItemNum As Long)
    Dim srec As RECT, drec As RECT
    
    If ItemNum > 0 Then
        With drec
            .Top = 0
            .Bottom = PIC_X
            .Left = 0
            .Right = PIC_Y
        End With
        With srec
            .Top = (Item(ItemNum).Pic \ 6) * PIC_Y ' BitBlt ySrc - top of rectangle
            .Bottom = .Top + 32 ' height of rectangle
            .Left = (Item(ItemNum).Pic - (Item(ItemNum).Pic \ 6) * 6) * PIC_X ' BitBlt xSrc - left of rectangle
            .Right = .Left + 32 ' width of rectangle
        End With
        
        Call DD_ItemSurf.BltToDC(picIngredient1.hDC, srec, drec)
        picIngredient1.Refresh
    End If
End Sub

Private Sub DrawIngredient2(ByVal ItemNum As Long)
    Dim srec As RECT, drec As RECT
    
    If ItemNum > 0 Then
        With drec
            .Top = 0
            .Bottom = PIC_X
            .Left = 0
            .Right = PIC_Y
        End With
        With srec
            .Top = (Item(ItemNum).Pic \ 6) * PIC_Y
            .Bottom = .Top + 32
            .Left = (Item(ItemNum).Pic - (Item(ItemNum).Pic \ 6) * 6) * PIC_X
            .Right = .Left + 32
        End With
        
        Call DD_ItemSurf.BltToDC(picIngredient2.hDC, srec, drec)
        picIngredient2.Refresh
    End If
End Sub

Private Sub DrawResultItem(ByVal ItemNum As Long)
    Dim srec As RECT, drec As RECT
    
    If ItemNum > 0 Then
        With drec
            .Top = 0
            .Bottom = PIC_X
            .Left = 0
            .Right = PIC_Y
        End With
        With srec
            .Top = (Item(ItemNum).Pic \ 6) * PIC_Y
            .Bottom = .Top + 32
            .Left = (Item(ItemNum).Pic - (Item(ItemNum).Pic \ 6) * 6) * PIC_X
            .Right = .Left + 32
        End With
        
        Call DD_ItemSurf.BltToDC(picResultItem.hDC, srec, drec)
        picResultItem.Refresh
    End If
End Sub

Private Sub Form_Load()
    If lstIngredient1.ListIndex > 0 Then
        If LenB(Trim$(Item(lstIngredient1.ListIndex).Name)) <> 100 Then
            Call DrawIngredient1(lstIngredient1.ListIndex)
        Else
            picIngredient1.Picture = LoadPicture()
        End If
    End If
    If lstIngredient2.ListIndex > 0 Then
        If LenB(Trim$(Item(lstIngredient2.ListIndex).Name)) <> 100 Then
            Call DrawIngredient2(lstIngredient2.ListIndex)
        Else
            picIngredient2.Picture = LoadPicture()
        End If
    End If
    If lstResultItem.ListIndex > 0 Then
        If LenB(Trim$(Item(lstResultItem.ListIndex).Name)) <> 100 Then
            Call DrawResultItem(lstResultItem.ListIndex)
        Else
            picResultItem.Picture = LoadPicture()
        End If
    End If
End Sub

Private Sub lstIngredient1_Click()
    Dim ItemNum As Long
    
    ItemNum = lstIngredient1.ListIndex
    
    If ItemNum > 0 Then
        If LenB(Trim$(Item(ItemNum).Name)) <> 100 Then
            Call DrawIngredient1(ItemNum)
        Else
            picIngredient1.Picture = LoadPicture()
        End If
    Else
        picIngredient1.Picture = LoadPicture()
    End If
End Sub

Private Sub lstIngredient2_Click()
    Dim ItemNum As Long
    
    ItemNum = lstIngredient2.ListIndex
    
    If ItemNum > 0 Then
        If LenB(Trim$(Item(ItemNum).Name)) <> 100 Then
            Call DrawIngredient2(ItemNum)
        Else
            picIngredient2.Picture = LoadPicture()
        End If
    Else
        picIngredient2.Picture = LoadPicture()
    End If
End Sub

Private Sub lstResultItem_Click()
    Dim ItemNum As Long
    
    ItemNum = lstResultItem.ListIndex
    
    If ItemNum > 0 Then
        If LenB(Trim$(Item(ItemNum).Name)) <> 100 Then
            Call DrawResultItem(ItemNum)
        Else
            picResultItem.Picture = LoadPicture()
        End If
    Else
        picResultItem.Picture = LoadPicture()
    End If
End Sub

Private Sub SaveRecipe_Click()
    Call RecipeEditorOk
End Sub
