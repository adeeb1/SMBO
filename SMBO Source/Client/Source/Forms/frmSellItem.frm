VERSION 5.00
Begin VB.Form frmSellItem 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sell Item"
   ClientHeight    =   6120
   ClientLeft      =   465
   ClientTop       =   660
   ClientWidth     =   3285
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSellItem.frx":0000
   ScaleHeight     =   6120
   ScaleWidth      =   3285
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   0
      Top             =   0
   End
   Begin VB.ListBox lstSellItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   4905
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   3255
   End
   Begin VB.Timer tmrClear 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3360
      Top             =   240
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Refresh"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   1320
      TabIndex        =   6
      Top             =   5160
      Width           =   855
   End
   Begin VB.Label CloseSell 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   2640
      TabIndex        =   5
      Top             =   5160
      Width           =   615
   End
   Begin VB.Label lblSellItem 
      BackStyle       =   0  'Transparent
      Caption         =   "Sell Item"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   0
      TabIndex        =   4
      Top             =   5160
      Width           =   855
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sell Items"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   3240
   End
   Begin VB.Label lblPrice 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   0
      TabIndex        =   2
      Top             =   5400
      Width           =   3255
   End
   Begin VB.Label lblSold 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   0
      TabIndex        =   1
      Top             =   5760
      Width           =   3255
   End
End
Attribute VB_Name = "frmSellItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Label1_Click()
    Dim i As Long, ItemNum As Long

    frmBank.lstInventory.Clear
    frmBank.lstBank.Clear
    
    For i = 1 To Player(MyIndex).MaxInv
        ItemNum = GetPlayerInvItemNum(MyIndex, i)
        
        If ItemNum > 0 Then
            If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Then
                frmBank.lstInventory.addItem i & "> " & Trim$(Item(ItemNum).Name) & " (" & GetPlayerInvItemValue(MyIndex, i) & ")"
            Else
                frmBank.lstInventory.addItem i & "> " & Trim$(Item(ItemNum).Name)
            End If
        Else
            frmBank.lstInventory.addItem i & "> Empty"
        End If
    Next i
    
    frmSellItem.lstSellItem.Clear
    
    For i = 1 To Player(MyIndex).MaxInv
        ItemNum = GetPlayerInvItemNum(MyIndex, i)
        
        If ItemNum > 0 Then
            If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Or Item(ItemNum).Stackable = 1 Then
                frmSellItem.lstSellItem.addItem i & "> " & Trim$(Item(ItemNum).Name) & " (" & GetPlayerInvItemValue(MyIndex, i) & ")"
            Else
                frmSellItem.lstSellItem.addItem i & "> " & Trim$(Item(ItemNum).Name)
            End If
        Else
            frmSellItem.lstSellItem.addItem i & "> Empty"
        End If
    Next i
    
    frmSellItem.lstSellItem.ListIndex = 0
End Sub

Private Sub lblSellItem_Click()
    Dim packet As String
    Dim ItemNum As Long, AMT As Long
    Dim ItemSlot As Integer

    ItemSlot = lstSellItem.ListIndex + 1
    ItemNum = GetPlayerInvItemNum(MyIndex, ItemSlot)
    
    If ItemNum > 0 Then
        If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Then
            Exit Sub
        Else
            If Item(ItemNum).Price > 0 Then
                Dim ItemName As String
                
                ItemName = Trim$(Item(ItemNum).Name)
                
                If Right$(ItemName, 1) = "s" Then
                    ItemName = Mid$(ItemName, 1, Len(ItemName) - 1)
                End If
                
                AMT = Val(InputBox("How many " & ItemName & "s would you like to sell?", "Sell " & ItemName, 0))
                
                If IsNumeric(AMT) Then
                    If AMT > 0 Then
                        packet = CPackets.Csellitem & SEP_CHAR & snumber & SEP_CHAR & ItemNum & SEP_CHAR & ItemSlot & SEP_CHAR & AMT & END_CHAR
                        Call SendData(packet)
                        
                        lblSold.Caption = "You sold " & AMT & " " & ItemName & "s ."
                    Else
                        Call MsgBox("You cannot sell fewer than 1 of an item!", 0, "Cannot Sell Fewer Than 1!")
                    End If
                End If
                
                tmrClear.Enabled = True
            Else
                Exit Sub
            End If
        End If
    Else
        Exit Sub
    End If
    
    Timer1.Enabled = True
End Sub

Private Sub lstSellItem_Click()
    Dim Q As Integer, i As Integer, p As Integer, ShopNum As Integer
    Dim ItemPrice As Currency
    Dim Stache As Long, ItemNum As Long
  
    ItemNum = GetPlayerInvItemNum(MyIndex, (lstSellItem.ListIndex + 1))
    
    If ItemNum > 0 Then
        ItemPrice = Item(ItemNum).Price
    End If
    
    For Q = 1 To MAX_SHOPS
        For i = 1 To MAX_SHOP_ITEMS
            If Shop(Q).ShopItem(i).ItemNum = ItemNum Then
                p = Shop(Q).ShopItem(i).currencyItem
                ShopNum = Q
                Exit For
                Exit For
            End If
        Next i
    Next Q
    
    If ShopNum <> 14 And ShopNum <> 19 And ShopNum <> 21 And ShopNum <> 23 Then
        Stache = GetPlayerStache(MyIndex)
    
        If Stache <= 30 Then
            ItemPrice = Int(ItemPrice * ((100 + Stache) / 100))
        Else
            ItemPrice = Int(ItemPrice * 1.3)
        End If
    End If
    
    If p <= 0 Then
        ' Give Coins for Mushroom Kingdom items and Beanbean Coins for Beanbean Kingdom items
        If ItemNum < 272 Then
            p = 1
        Else
            p = 271
        End If
    End If
    
    If ItemNum > 0 Then
        If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Then
            lblPrice.Caption = "Please select an item to sell."
        Else
            If Item(ItemNum).Price > 1 Then
                lblPrice.Caption = "Price: " & ItemPrice & " " & Trim$(Item(p).Name) & "(s)"
            ElseIf Item(ItemNum).Price = 1 Then
                lblPrice.Caption = "Price: 1 " & Trim$(Item(p).Name)
            Else
                lblPrice.Caption = "This item cannot be sold."
            End If
        End If
    Else
        lblPrice.Caption = "Please select an item to sell."
    End If
End Sub

Private Sub Form_Load()
    Dim i As Long, ItemNum As Long
        
    lblSold.Caption = vbNullString
    lblPrice.Caption = vbNullString
    frmSellItem.lstSellItem.Clear
    
    For i = 1 To Player(MyIndex).MaxInv
      ItemNum = GetPlayerInvItemNum(MyIndex, i)
        If ItemNum > 0 Then
            If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Or Item(ItemNum).Stackable = 1 Then
                frmSellItem.lstSellItem.addItem i & "> " & Trim$(Item(ItemNum).Name) & " (" & GetPlayerInvItemValue(MyIndex, i) & ")"
            Else
                frmSellItem.lstSellItem.addItem i & "> " & Trim$(Item(ItemNum).Name)
            End If
        Else
            frmSellItem.lstSellItem.addItem i & "> Empty"
        End If
    Next i
    
    frmSellItem.lstSellItem.ListIndex = 0
End Sub

Private Sub Timer1_Timer()
    Call Label1_Click
    Timer1.Enabled = False
End Sub

Private Sub tmrClear_Timer()
    lblSold.Caption = vbNullString
End Sub

Private Sub CloseSell_Click()
    Unload Me
End Sub
