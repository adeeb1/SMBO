VERSION 5.00
Begin VB.Form frmPlayerTrade 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Trading"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5985
   ControlBox      =   0   'False
   Icon            =   "frmPlayerTrade.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmPlayerTrade.frx":0FC2
   ScaleHeight     =   399
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   399
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox Items2 
      Appearance      =   0  'Flat
      BackColor       =   &H00008080&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1590
      Left            =   3240
      TabIndex        =   2
      Top             =   2160
      Width           =   1905
   End
   Begin VB.ListBox Items1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1590
      Left            =   795
      TabIndex        =   1
      Top             =   3735
      Width           =   1905
   End
   Begin VB.ListBox PlayerInv1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1590
      Left            =   795
      TabIndex        =   0
      Top             =   1530
      Width           =   1905
   End
   Begin VB.Label Accepted 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   795
      TabIndex        =   11
      Top             =   5340
      Width           =   2775
   End
   Begin VB.Label TradingWith 
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
      ForeColor       =   &H00400040&
      Height          =   825
      Left            =   615
      TabIndex        =   10
      Top             =   360
      Width           =   1365
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   900
      TabIndex        =   9
      Top             =   3405
      Width           =   1725
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   885
      TabIndex        =   8
      Top             =   1200
      Width           =   1725
   End
   Begin VB.Label Command3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   4155
      TabIndex        =   7
      Top             =   4395
      Width           =   885
   End
   Begin VB.Label Command4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   4155
      TabIndex        =   6
      Top             =   4815
      Width           =   885
   End
   Begin VB.Label Command5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   3855
      TabIndex        =   5
      Top             =   5295
      Width           =   1230
   End
   Begin VB.Label Command2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   2880
      TabIndex        =   4
      Top             =   1320
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.Label Command1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   4140
      TabIndex        =   3
      Top             =   3960
      Width           =   915
   End
End
Attribute VB_Name = "frmPlayerTrade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Call SendData(CPackets.Ccompletetrade & END_CHAR)
End Sub

Private Sub Command3_Click()
    Dim n As Long, InvNum As Long, ItemNum As Long, Amount As Long
    Dim itemName As String
    
    InvNum = PlayerInv1.ListIndex + 1
    ItemNum = GetPlayerInvItemNum(MyIndex, InvNum)
    
    If ItemNum < 1 Or ItemNum > MAX_ITEMS Then
        Exit Sub
    End If
    
    If Item(ItemNum).Bound = 1 Then
        Call AddText("This item is untradeable.", BRIGHTRED)
        Exit Sub
    End If
        
    itemName = Trim$(Item(ItemNum).Name)
    
    ' Loop through all current trades
    For n = 1 To MAX_PLAYER_TRADES
        ' Make sure we don't offer something that we've already put up for offer
        If PlayerTrading(n).InvNum = InvNum Then
            MsgBox "You can only trade that item once!"
            Exit Sub
        End If
        
        ' Make sure we don't overwrite a current trade offer
        If PlayerTrading(n).InvNum <= 0 Then
            If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Or Item(ItemNum).Stackable = 1 Then
                Amount = Val(InputBox("How many " & itemName & " (" & GetPlayerInvItemValue(MyIndex, InvNum) & ") would you like to trade?", "Trade " & itemName, vbNullString, frmPlayerTrade.Left, frmPlayerTrade.Top))
                    
                ' Make sure the player trades at least 1 of the item
                If Amount <= 0 Then
                    Call MsgBox("You must trade at least 1 of an item!", 0, "Must Trade At Least 1!")
                    Exit Sub
                End If
                ' Make sure the player has enough of that item in his/her inventory
                If Amount > GetPlayerInvItemValue(MyIndex, InvNum) Then
                    Call AddText("You don't have that many to trade!", BRIGHTRED)
                    Exit Sub
                End If
                
                ' Update the listboxes
                PlayerInv1.List(InvNum - 1) = PlayerInv1.Text & " **"
                Items1.List(n - 1) = n & ": " & itemName & " (" & Amount & ")"
                
                ' Store the offer details
                PlayerTrading(n).InvName = itemName
                PlayerTrading(n).InvNum = InvNum
                PlayerTrading(n).InvVal = Amount
                
                ' Send the information to the server
                Call SendData(CPackets.Cupdatetradeoffers & SEP_CHAR & n & SEP_CHAR & itemName & SEP_CHAR & InvNum & SEP_CHAR & Amount & END_CHAR)
                Exit Sub
            Else
                ' Update the listboxes
                PlayerInv1.List(InvNum - 1) = PlayerInv1.Text & " **"
                Items1.List(n - 1) = n & ": " & itemName
                
                ' Store the offer details
                PlayerTrading(n).InvName = itemName
                PlayerTrading(n).InvNum = InvNum
                PlayerTrading(n).InvVal = 1
                
                ' Send the information to the server
                Call SendData(CPackets.Cupdatetradeoffers & SEP_CHAR & n & SEP_CHAR & itemName & SEP_CHAR & InvNum & SEP_CHAR & 1 & END_CHAR)
                Exit Sub
            End If
        End If
    Next n
End Sub

Private Sub Command4_Click()
    Dim n As Long, TradeOfferNum As Long
    
    TradeOfferNum = Items1.ListIndex + 1

    If PlayerTrading(TradeOfferNum).InvNum <= 0 Then
        MsgBox "No item to remove!"
        Exit Sub
    End If

    ' Update the listboxes
    PlayerInv1.List(PlayerTrading(TradeOfferNum).InvNum - 1) = Mid$(Trim$(PlayerInv1.List(PlayerTrading(TradeOfferNum).InvNum - 1)), 1, Len(PlayerInv1.List(PlayerTrading(TradeOfferNum).InvNum - 1)) - 3)
    Items1.List(TradeOfferNum - 1) = TradeOfferNum & ": <Nothing>"
    
    ' Store the offer details
    PlayerTrading(TradeOfferNum).InvName = vbNullString
    PlayerTrading(TradeOfferNum).InvNum = 0
    PlayerTrading(TradeOfferNum).InvVal = 0
    
    ' Send the information to the server
    Call SendData(CPackets.Cupdatetradeoffers & SEP_CHAR & TradeOfferNum & SEP_CHAR & vbNullString & SEP_CHAR & 0 & SEP_CHAR & 0 & END_CHAR)
End Sub

Private Sub Command5_Click()
    Call SendData(CPackets.Cstoptrading & END_CHAR)
End Sub
