VERSION 5.00
Begin VB.Form frmRecipe 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4800
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   6300
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmRecipe.frx":0000
   ScaleHeight     =   320
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   420
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lstRecipes 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Height          =   2430
      ItemData        =   "frmRecipe.frx":62742
      Left            =   360
      List            =   "frmRecipe.frx":62744
      TabIndex        =   6
      Top             =   1200
      Width           =   1815
   End
   Begin VB.PictureBox picRecipe1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   4470
      ScaleHeight     =   71
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   80
      TabIndex        =   3
      Top             =   1245
      Width           =   1230
      Begin VB.PictureBox picRecipe 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   375
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   4
         Top             =   270
         Width           =   480
      End
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
      Height          =   1290
      Left            =   4215
      TabIndex        =   5
      Top             =   2820
      Width           =   1770
   End
   Begin VB.Label ReturnButton 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   375
      TabIndex        =   2
      Top             =   4035
      Width           =   2055
   End
   Begin VB.Label lblIngredient2 
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
      Height          =   300
      Left            =   2460
      TabIndex        =   1
      Top             =   2040
      Width           =   1740
   End
   Begin VB.Label lblIngredient1 
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
      Height          =   300
      Left            =   2460
      TabIndex        =   0
      Top             =   1350
      Width           =   1740
   End
End
Attribute VB_Name = "frmRecipe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub DrawRecipe(ItemNum As Long)
    Dim srec As RECT, drec As RECT
    
    drec.Top = 0
    drec.Bottom = PIC_X
    drec.Left = 0
    drec.Right = PIC_Y
    srec.Top = (Item(ItemNum).Pic \ 6) * PIC_Y
    srec.Bottom = srec.Top + PIC_X
    srec.Left = (Item(ItemNum).Pic - (Item(ItemNum).Pic \ 6) * 6) * PIC_X
    srec.Right = srec.Left + PIC_Y
    
    Call DD_ItemSurf.BltToDC(picRecipe.hDC, srec, drec)
End Sub

Private Sub lstRecipes_Click()
    Dim RecipeNum As Integer
    
    If Trim$(lstRecipes.List(lstRecipes.ListIndex)) = "<Empty Slot>" Then
        lblIngredient1.Caption = vbNullString
        lblIngredient2.Caption = vbNullString
        lblDesc.Caption = vbNullString
        
        picRecipe.Picture = LoadPicture()
        Exit Sub
    End If
    
    RecipeNum = lstRecipes.ListIndex + 1
    
    lblIngredient1.Caption = Trim$(Item(Recipe(RecipeNum).Ingredient1).Name)
    
    If Recipe(RecipeNum).Ingredient2 > 0 Then
        lblIngredient2.Caption = Trim$(Item(Recipe(RecipeNum).Ingredient2).Name)
    Else
        lblIngredient2.Caption = "None"
    End If
    
    lblDesc.Caption = Trim$(Item(Recipe(RecipeNum).ResultItem).desc)
    
    Call DrawRecipe(Recipe(RecipeNum).ResultItem)
End Sub

Private Sub ReturnButton_Click()
    Unload Me
End Sub
