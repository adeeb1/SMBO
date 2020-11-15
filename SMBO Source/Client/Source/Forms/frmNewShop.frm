VERSION 5.00
Begin VB.Form frmNewShop 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Shop"
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   4290
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmNewShop.frx":0000
   ScaleHeight     =   315
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   286
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picItemInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      HasDC           =   0   'False
      Height          =   4200
      Left            =   1680
      ScaleHeight     =   4170
      ScaleWidth      =   2385
      TabIndex        =   22
      Top             =   15
      Width           =   2415
      Begin VB.Label lblCritBlockBonus 
         BackStyle       =   0  'Transparent
         Caption         =   "Crit/Dodge Bonus:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   39
         Top             =   2520
         Width           =   2055
      End
      Begin VB.Line Line2 
         X1              =   240
         X2              =   2160
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Label lblHPFPReq 
         BackStyle       =   0  'Transparent
         Caption         =   "HP/FP Req:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   38
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label lblLevelReq 
         BackStyle       =   0  'Transparent
         Caption         =   "Level Req:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   37
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label lblClassReq 
         BackStyle       =   0  'Transparent
         Caption         =   "Character Req:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   36
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label lblDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Desc"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   990
         Left            =   240
         TabIndex        =   29
         Top             =   3000
         Width           =   1935
      End
      Begin VB.Label lblSpdBonus 
         BackStyle       =   0  'Transparent
         Caption         =   "Spd/Stache Bonus:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   2280
         Width           =   2055
      End
      Begin VB.Label lblAddStr 
         BackStyle       =   0  'Transparent
         Caption         =   "Atk/Def Bonus:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   2040
         Width           =   2055
      End
      Begin VB.Line Line1 
         X1              =   240
         X2              =   2160
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Label lblVital 
         BackStyle       =   0  'Transparent
         Caption         =   "HP/FP/SP Bonus:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   1800
         Width           =   2055
      End
      Begin VB.Label lblSpdReq 
         BackStyle       =   0  'Transparent
         Caption         =   "Spd/Stache Req:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label lblStrReq 
         BackStyle       =   0  'Transparent
         Caption         =   "Atk/Def Req:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "-Item Info-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   120
         Width           =   2055
      End
   End
   Begin VB.PictureBox ConfirmBuy 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      ForeColor       =   &H80000008&
      Height          =   1530
      Left            =   1560
      Picture         =   "frmNewShop.frx":42276
      ScaleHeight     =   1500
      ScaleWidth      =   1755
      TabIndex        =   32
      Top             =   1440
      Width           =   1785
      Begin VB.Label NoBuy 
         BackStyle       =   0  'Transparent
         Height          =   290
         Left            =   1090
         TabIndex        =   35
         Top             =   1135
         Width           =   500
      End
      Begin VB.Label YesBuy 
         BackStyle       =   0  'Transparent
         Height          =   290
         Left            =   202
         TabIndex        =   34
         Top             =   1135
         Width           =   500
      End
      Begin VB.Label ConfirmTxt 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Are you sure you want to buy this item?"
         BeginProperty Font 
            Name            =   "Porky's"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   120
         TabIndex        =   33
         Top             =   260
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "<"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   4320
      Width           =   255
   End
   Begin VB.PictureBox imgBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   540
      Index           =   4
      Left            =   120
      ScaleHeight     =   34
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   34
      TabIndex        =   10
      Top             =   3480
      Width           =   540
      Begin VB.PictureBox picEmoticon 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   480
         Index           =   4
         Left            =   15
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   11
         Top             =   15
         Width           =   480
         Begin VB.PictureBox iconn 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   480
            Index           =   4
            Left            =   0
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   12
            Top             =   0
            Width           =   480
         End
      End
   End
   Begin VB.PictureBox imgBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   540
      Index           =   3
      Left            =   120
      ScaleHeight     =   34
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   34
      TabIndex        =   7
      Top             =   2640
      Width           =   540
      Begin VB.PictureBox picEmoticon 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   480
         Index           =   3
         Left            =   15
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   8
         Top             =   15
         Width           =   480
         Begin VB.PictureBox iconn 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   480
            Index           =   3
            Left            =   0
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   9
            Top             =   0
            Width           =   480
         End
      End
   End
   Begin VB.PictureBox imgBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   540
      Index           =   2
      Left            =   120
      ScaleHeight     =   34
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   34
      TabIndex        =   4
      Top             =   1800
      Width           =   540
      Begin VB.PictureBox picEmoticon 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   480
         Index           =   2
         Left            =   15
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   5
         Top             =   15
         Width           =   480
         Begin VB.PictureBox iconn 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   480
            Index           =   2
            Left            =   0
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   6
            Top             =   0
            Width           =   480
         End
      End
   End
   Begin VB.PictureBox imgBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   540
      Index           =   0
      Left            =   120
      ScaleHeight     =   34
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   34
      TabIndex        =   2
      Top             =   120
      Width           =   540
      Begin VB.PictureBox picEmoticon 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   480
         Index           =   1
         Left            =   15
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   3
         Top             =   15
         Width           =   480
         Begin VB.PictureBox iconn 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   480
            Index           =   0
            Left            =   0
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   21
            Top             =   0
            Width           =   480
         End
      End
   End
   Begin VB.PictureBox imgBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   540
      Index           =   1
      Left            =   120
      ScaleHeight     =   34
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   34
      TabIndex        =   0
      Top             =   960
      Width           =   540
      Begin VB.PictureBox picEmoticon 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   480
         Index           =   0
         Left            =   15
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   1
         Top             =   15
         Width           =   480
         Begin VB.PictureBox iconn 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   480
            Index           =   1
            Left            =   0
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   20
            Top             =   0
            Width           =   480
         End
      End
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   ">"
      Height          =   255
      Left            =   3840
      TabIndex        =   14
      Top             =   4320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblSell 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Sell Items"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   480
      TabIndex        =   31
      Top             =   4320
      Width           =   975
   End
   Begin VB.Label lblPage 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Page: X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1560
      TabIndex        =   30
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Label lblItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   615
      Index           =   4
      Left            =   840
      TabIndex        =   19
      Top             =   3600
      Width           =   3255
   End
   Begin VB.Label lblItem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   615
      Index           =   3
      Left            =   840
      TabIndex        =   18
      Top             =   2760
      Width           =   3255
   End
   Begin VB.Label lblItem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   615
      Index           =   2
      Left            =   840
      TabIndex        =   17
      Top             =   1920
      Width           =   3255
   End
   Begin VB.Label lblItem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   615
      Index           =   1
      Left            =   840
      TabIndex        =   16
      Top             =   1080
      Width           =   3255
   End
   Begin VB.Label lblItem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   615
      Index           =   0
      Left            =   840
      TabIndex        =   15
      Top             =   240
      Width           =   3255
   End
End
Attribute VB_Name = "frmNewShop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private numItems As Integer, pageIndex As Integer, Q As Integer
Public ShopNum As Integer
Public SellItems As Boolean 'Is the shop selling items?
Private NextPage As Double
Private BuyAmt As Long
Private BuyItem As Boolean

' Loads shop data into the form for the first time.
Public Sub loadShop(ByVal sNum As Integer)
    Dim i As Integer
    
    numItems = 0
    pageIndex = 0
    ShopNum = sNum
    
    cmdBack.Visible = False
    Me.Caption = Shop(sNum).Name

    ' Check to see if there are more pages
    For i = 1 To MAX_SHOP_ITEMS
        If Shop(ShopNum).ShopItem(i).ItemNum > 0 Then
            numItems = numItems + 1
        End If
    Next i

    NextPage = numItems / 5

    If numItems > 5 Then
        cmdNext.Visible = True
    Else
        cmdNext.Visible = False
    End If

    ' Check if this shop buys back items
    If Shop(sNum).BuysItems = Yes Then
        lblSell.Visible = True
    Else
        lblSell.Visible = False
    End If
End Sub

' Shows the specified page
Public Sub showPage(ByVal page As Integer)
    Dim i As Integer
    Dim ItemName As String, shopCurrency As String
    Dim ItemPrice As Currency
    Dim Stache As Long
    
    On Error GoTo showPage_Error

    lblPage.Caption = "Page: " & (page + 1)
    
    Stache = GetPlayerStache(MyIndex)

    For i = 1 To 5
        If Shop(ShopNum).ShopItem(page * 5 + i).ItemNum = 0 Then
            imgBox(i - 1).Visible = False
            lblItem(i - 1).Visible = False
        Else
            imgBox(i - 1).Visible = True
            lblItem(i - 1).Visible = True
            
            ItemPrice = Shop(ShopNum).ShopItem(pageIndex * 5 + i).Price
            
            If ShopNum <> 14 And ShopNum <> 19 And ShopNum <> 21 And ShopNum <> 23 Then
                If Stache <= 30 Then
                    ItemPrice = Int(ItemPrice * ((100 - Stache) / 100))
                Else
                    ItemPrice = Int(ItemPrice * 0.7)
                End If
            End If
            
            ItemName = Trim$(Item(Shop(ShopNum).ShopItem(pageIndex * 5 + i).ItemNum).Name)
            shopCurrency = Trim$(Item(Shop(ShopNum).ShopItem(pageIndex * 5 + i).currencyItem).Name)
            lblItem(i - 1).Caption = ItemName & vbNewLine & "Original Price: " & STR(Shop(ShopNum).ShopItem(pageIndex * 5 + i).Price) & " " & shopCurrency & "(s)" & vbNewLine & "Your Price: " & STR(ItemPrice) & " " & shopCurrency & "(s)"
            
            Me.iconn(i - 1).Cls

            Call BltIcon(i - 1, Shop(ShopNum).ShopItem(pageIndex * 5 + i).ItemNum)
        End If
    Next i

    ' If numItems / 5 - (pageIndex * 5) > 1 Then
    If (page + 1) < NextPage And NextPage <= (MAX_SHOP_ITEMS / 5) Then
        cmdNext.Visible = True
    Else
        cmdNext.Visible = False
    End If

    If pageIndex > 0 Then
        cmdBack.Visible = True
    Else
        cmdBack.Visible = False
    End If

    On Error GoTo 0
    Exit Sub

showPage_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure showPage of Form frmNewShop"
    
    If MsgBox("Could not show page.", vbRetryCancel) = vbRetry Then
        Call showPage(page)
    Else
        frmNewShop.Visible = False
    End If
    
    Exit Sub
End Sub

Public Sub Buy(shopItemIndex As Integer, ItemAmount As Long)
    ' Send buy request to server
    Call SendData(CPackets.Cbuy & SEP_CHAR & ShopNum & SEP_CHAR & shopItemIndex & SEP_CHAR & ItemAmount & END_CHAR)
End Sub

Public Sub Buyback(ByVal Item As Integer, ByVal slot As Integer, Optional ByVal AMT As Integer = 1)
    Call SendData(CPackets.Csellitem & SEP_CHAR & ShopNum & SEP_CHAR & Item & SEP_CHAR & slot & SEP_CHAR & AMT & END_CHAR)
End Sub

' Draws icons to teh boxx0r
Private Sub BltIcon(ByVal iconNum As Integer, ByVal ItemNum As Integer)
    On Error Resume Next
    
    Dim itemX As Integer, itemY As Integer
    Dim srec As RECT, drec As RECT

    ItemNum = Shop(ShopNum).ShopItem(pageIndex * 5 + iconNum + 1).ItemNum
    itemX = (Item(ItemNum).Pic - (Item(ItemNum).Pic \ 6) * 6) * PIC_X
    itemY = (Item(ItemNum).Pic \ 6) * PIC_Y
    
    drec.Top = 0
    drec.Bottom = 32
    drec.Left = 0
    drec.Right = 32
    srec.Top = itemY
    srec.Bottom = srec.Top + 32
    srec.Left = itemX
    srec.Right = srec.Left + 32
    
    Call DD_ItemSurf.BltToDC(iconn(iconNum).hDC, srec, drec)

    ' Clear any errors
    Err.Clear
End Sub

Private Sub ShowItemInfo(ByVal itemN As Integer)
    Dim HPReq As String, FPReq As String, SPBonus As String

    picItemInfo.Visible = True
    
    ' Descriptions for Atk/Defense Requirements
    lblStrReq.Caption = "Atk/Def Req: " & Item(itemN).StrReq & "/" & Item(itemN).DefReq
    
    ' Descriptions for Speed/Stache Requirements
    lblSpdReq.Caption = "Spd/Stache Req: " & Item(itemN).SpeedReq & "/" & Item(itemN).MagicReq
    
    ' Descriptions for HP/FP Requirements
    lblHPFPReq.Caption = "HP/FP Req: " & Item(itemN).HPReq & "/" & Item(itemN).FPReq
    
    ' Descriptions for Character Requirements
    If Item(itemN).Type >= ITEM_TYPE_CHANGEHPFPSP And Item(itemN).Type <= ITEM_TYPE_SCRIPTED Then
        lblClassReq.Caption = "Character Req: " & "None"
    ElseIf Item(itemN).Type = ITEM_TYPE_NONE Or Item(itemN).ClassReq = -1 Then
        lblClassReq.Caption = "Character Req: " & "None"
    Else
        lblClassReq.Caption = "Character Req: " & Trim$(Class(Item(itemN).ClassReq).Name)
    End If
    
    ' Descriptions for Level Requirements
    If Item(itemN).LevelReq > 0 Then
        lblLevelReq.Caption = "Level Req: " & Item(itemN).LevelReq
    Else
        lblLevelReq.Caption = "Level Req: None"
    End If
    
    ' Descriptions for HP/FP/SP Bonuses
    If Item(itemN).Type >= ITEM_TYPE_WEAPON And Item(itemN).Type <= ITEM_TYPE_MUSHROOMBADGE Then
        HPReq = CStr(Item(itemN).AddHP)
        FPReq = CStr(Item(itemN).AddMP)
        SPBonus = CStr(Item(itemN).AddSP)
    ElseIf Item(itemN).Type = ITEM_TYPE_CHANGEHPFPSP Then
        HPReq = CStr(Item(itemN).Data1)
        FPReq = CStr(Item(itemN).Data2)
        SPBonus = CStr(Item(itemN).Data3)
    Else
        HPReq = CStr(0)
        FPReq = CStr(0)
        SPBonus = CStr(0)
    End If
    
    ' Descriptions for HP/FP/SP Bonuses
    lblVital.Caption = "HP/FP/SP Bonus: " & HPReq & "/" & FPReq & "/" & SPBonus
    
    ' Descriptions for Atk/Def Bonuses
    lblAddStr.Caption = "Atk/Def Bonus: " & Item(itemN).AddSTR & "/" & Item(itemN).AddDef
    
    ' Descriptions for Speed/Stache Bonuses
    lblSpdBonus.Caption = "Spd/Stache Bonus: " & Item(itemN).AddSpeed & "/" & Item(itemN).AddMAGI
    
    ' Descriptions for Critical Hit/Block Chance Bonuses
    lblCritBlockBonus.Caption = "Crit/Block Bonus: " & Item(itemN).AddCritChance & "/" & Item(itemN).AddBlockChance
    
    ' Item Description
    lblDesc.Caption = Item(itemN).desc
End Sub

Private Sub HideItemInfo()
    picItemInfo.Visible = False
End Sub

Private Sub lblSell_Click()
    frmSellItem.Show vbModeless, frmNewShop
End Sub

Private Sub cmdBack_Click()
    pageIndex = pageIndex - 1
    Call showPage(pageIndex)
End Sub

Private Sub cmdNext_Click()
    pageIndex = pageIndex + 1
    Call showPage(pageIndex)
End Sub

Private Sub Form_Load()
    ConfirmBuy.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    IsShopping = False
    frmMirage.SetFocus
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Shop(ShopNum).ShowInfo = 1 Then
        Call HideItemInfo
    End If
End Sub

' Buy item
Private Sub imgBox_Click(Index As Integer)
    Q = Index
 
    ConfirmBuy.Visible = True
    
    If BuyItem = True Then
        Dim ItemName As String
                
        ItemName = Trim$(Item(Shop(ShopNum).ShopItem(pageIndex * 5 + Index + 1).ItemNum).Name)
                    
        If Right$(ItemName, 1) = "s" Then
            ItemName = Mid$(ItemName, 1, Len(ItemName) - 1)
        End If
        
        BuyAmt = Val(InputBox("How many " & ItemName & "s would you like to buy?"))
        
        If IsNumeric(BuyAmt) And BuyAmt > 0 Then
            Buy pageIndex * 5 + Index + 1, BuyAmt
        Else
            Call MsgBox("You cannot buy fewer than 1 of an item!", 0, "Cannot Buy Fewer Than 1!")
        End If
        
        BuyItem = False
     End If
End Sub

' Buy item
Private Sub iconn_Click(Index As Integer)
    Q = Index
 
    ConfirmBuy.Visible = True
    
    If BuyItem = True Then
        Dim ItemName As String
                
        ItemName = Trim$(Item(Shop(ShopNum).ShopItem(pageIndex * 5 + Index + 1).ItemNum).Name)
                    
        If Right$(ItemName, 1) = "s" Then
            ItemName = Mid$(ItemName, 1, Len(ItemName) - 1)
        End If
        
        BuyAmt = Val(InputBox("How many " & ItemName & "s would you like to buy?"))
        
        If IsNumeric(BuyAmt) And BuyAmt > 0 Then
            Buy pageIndex * 5 + Index + 1, BuyAmt
        Else
            Call MsgBox("You cannot buy fewer than 1 of an item!", 0, "Cannot Buy Fewer Than 1!")
        End If
        
        BuyItem = False
     End If
End Sub

Private Sub iconn_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Shop(ShopNum).ShowInfo = 1 Then
        Call ShowItemInfo(Shop(ShopNum).ShopItem(pageIndex * 5 + Index + 1).ItemNum)
    End If
End Sub

Private Sub NoBuy_Click()
    ConfirmBuy.Visible = False
End Sub

Private Sub YesBuy_Click()
    BuyItem = True
  
    Call iconn_Click(Q)
  
    ConfirmBuy.Visible = False
End Sub
