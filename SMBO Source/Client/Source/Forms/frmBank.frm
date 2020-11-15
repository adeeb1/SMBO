VERSION 5.00
Begin VB.Form frmBank 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bank"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5985
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmBank.frx":0000
   ScaleHeight     =   399
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   399
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox ThrowAwayItem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   2160
      Picture         =   "frmBank.frx":74E92
      ScaleHeight     =   1665
      ScaleWidth      =   1785
      TabIndex        =   12
      Top             =   2160
      Visible         =   0   'False
      Width           =   1815
      Begin VB.Label NoThrowAway 
         BackStyle       =   0  'Transparent
         Height          =   275
         Left            =   1020
         TabIndex        =   15
         Top             =   1190
         Width           =   455
      End
      Begin VB.Label YesThrowAway 
         BackStyle       =   0  'Transparent
         Height          =   275
         Left            =   320
         TabIndex        =   14
         Top             =   1190
         Width           =   455
      End
      Begin VB.Label ThrowAwayMsg 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Are you sure you want to throw this item away?"
         BeginProperty Font 
            Name            =   "Porky's"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   13
         Top             =   125
         Width           =   1575
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   540
      Index           =   1
      Left            =   7320
      ScaleHeight     =   34
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   34
      TabIndex        =   8
      Top             =   2400
      Width           =   540
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Height          =   480
         Index           =   1
         Left            =   15
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   9
         Top             =   15
         Width           =   480
         Begin VB.PictureBox PicBank 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   480
            Index           =   1
            Left            =   15
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   128
            TabIndex        =   10
            Top             =   15
            Width           =   1920
         End
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   540
      Index           =   0
      Left            =   7320
      ScaleHeight     =   34
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   34
      TabIndex        =   5
      Top             =   1680
      Width           =   540
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Height          =   480
         Index           =   0
         Left            =   15
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   6
         Top             =   15
         Width           =   480
         Begin VB.PictureBox PicBank 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   480
            Index           =   0
            Left            =   15
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   128
            TabIndex        =   7
            Top             =   15
            Width           =   1920
         End
      End
   End
   Begin VB.ListBox lstBank 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   1290
      IntegralHeight  =   0   'False
      Left            =   225
      TabIndex        =   1
      Top             =   3420
      Width           =   2655
   End
   Begin VB.ListBox lstInventory 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   1290
      IntegralHeight  =   0   'False
      Left            =   225
      TabIndex        =   0
      Top             =   1215
      Width           =   2655
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   1590
      TabIndex        =   11
      Top             =   4935
      Width           =   1110
   End
   Begin VB.Label lblMsg 
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
      Left            =   5235
      TabIndex        =   4
      Top             =   360
      Width           =   525
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   390
      TabIndex        =   3
      Top             =   4935
      Width           =   1110
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   405
      TabIndex        =   2
      Top             =   2730
      Width           =   2295
   End
End
Attribute VB_Name = "frmBank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Label1_Click()
    Call BankItems
End Sub

Private Sub Label2_Click()
Dim BankNum As Long

    BankNum = lstBank.ListIndex + 1
    If GetPlayerBankItemNum(MyIndex, BankNum) > 0 And GetPlayerBankItemNum(MyIndex, BankNum) <= MAX_ITEMS Then
        ThrowAwayItem.Visible = True
    End If
    
End Sub

Private Sub Label3_Click()
    Call InvItems
End Sub

Sub BankItems()
    Dim InvNum As Long
    Dim GoldAmount As String
    Dim ItemNum As Long
    On Error GoTo Done

    InvNum = lstInventory.ListIndex + 1
    ItemNum = GetPlayerInvItemNum(MyIndex, InvNum)
    If ItemNum > 0 And ItemNum <= MAX_ITEMS Then
        If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Then
            GoldAmount = InputBox("How much " & Trim$(Item(ItemNum).Name) & "(" & GetPlayerInvItemValue(MyIndex, InvNum) & ") would you like to deposit?", "Deposit " & Trim$(Item(ItemNum).Name), 0, frmBank.Left, frmBank.Top)
            If IsNumeric(GoldAmount) Then
              If GoldAmount > 0 Then
                Call SendData(CPackets.Cbankdeposit & SEP_CHAR & lstInventory.ListIndex + 1 & SEP_CHAR & Val(GoldAmount) & END_CHAR)
              Else
                Call MsgBox("You must deposit at least 1 of an item!", 0, "Must Deposit At Least 1!")
              End If
            End If
        Else
            Call SendData(CPackets.Cbankdeposit & SEP_CHAR & lstInventory.ListIndex + 1 & SEP_CHAR & GetPlayerInvItemValue(MyIndex, InvNum) & END_CHAR)
        End If
    End If
    Exit Sub
Done:
    If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Then
    ' MsgBox "The variable cant handle that amount!"
    End If
End Sub

Sub InvItems()
    Dim BankNum As Long
    Dim GoldAmount As String
    On Error GoTo Done

    BankNum = lstBank.ListIndex + 1
    If GetPlayerBankItemNum(MyIndex, BankNum) > 0 And GetPlayerBankItemNum(MyIndex, BankNum) <= MAX_ITEMS Then
        If Item(GetPlayerBankItemNum(MyIndex, BankNum)).Type = ITEM_TYPE_CURRENCY Then
            GoldAmount = InputBox("How much " & Trim$(Item(GetPlayerBankItemNum(MyIndex, BankNum)).Name) & "(" & GetPlayerBankItemValue(MyIndex, BankNum) & ") would you like to withdraw?", "Withdraw " & Trim$(Item(GetPlayerBankItemNum(MyIndex, BankNum)).Name), 0, frmBank.Left, frmBank.Top)
            If IsNumeric(GoldAmount) Then
              If GoldAmount > 0 Then
                Call SendData(CPackets.Cbankwithdraw & SEP_CHAR & lstBank.ListIndex + 1 & SEP_CHAR & Val(GoldAmount) & END_CHAR)
              Else
                Call MsgBox("You must withdraw at least 1 of an item!", 0, "Must Withdraw At Least 1!")
              End If
            End If
        Else
            Call SendData(CPackets.Cbankwithdraw & SEP_CHAR & lstBank.ListIndex + 1 & SEP_CHAR & GetPlayerBankItemValue(MyIndex, BankNum) & END_CHAR)
        End If
    End If
    Exit Sub
Done:
    If Item(GetPlayerBankItemNum(MyIndex, BankNum)).Type = ITEM_TYPE_CURRENCY Then
        MsgBox "The variable cant handle that amount!"
    End If
End Sub

Private Sub lblMsg_Click()
    IsBanking = False
    Unload Me
End Sub

Private Sub NoThrowAway_Click()
    ThrowAwayItem.Visible = False
End Sub

Private Sub YesThrowAway_Click()
    Dim BankNum As Long
    Dim GoldAmount As String

    BankNum = lstBank.ListIndex + 1
    If GetPlayerBankItemNum(MyIndex, BankNum) > 0 And GetPlayerBankItemNum(MyIndex, BankNum) <= MAX_ITEMS Then
        If Item(GetPlayerBankItemNum(MyIndex, BankNum)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerBankItemNum(MyIndex, BankNum)).Stackable = 1 Then
            GoldAmount = InputBox("How many " & Trim$(Item(GetPlayerBankItemNum(MyIndex, BankNum)).Name) & "(" & GetPlayerBankItemValue(MyIndex, BankNum) & ") would you like to throw away?", "Throw Away " & Trim$(Item(GetPlayerBankItemNum(MyIndex, BankNum)).Name), 0, frmBank.Left, frmBank.Top)
            If IsNumeric(GoldAmount) Then
              If GoldAmount > 0 Then
                Call SendData(CPackets.Cbankdestroy & SEP_CHAR & lstBank.ListIndex + 1 & SEP_CHAR & Val(GoldAmount) & END_CHAR)
              Else
                Call MsgBox("You must throw away at least 1 of an item!", 0, "Must Throw Away At Least 1!")
              End If
            End If
        Else
            Call SendData(CPackets.Cbankdestroy & SEP_CHAR & lstBank.ListIndex + 1 & SEP_CHAR & GetPlayerBankItemValue(MyIndex, BankNum) & END_CHAR)
        End If
    End If
    
    ThrowAwayItem.Visible = False
End Sub

Public Sub OpenBank()
    Dim i As Long, n As Long
    Dim ItemName As String
    
    frmBank.lstInventory.Clear
    frmBank.lstBank.Clear
    
    For i = 1 To Player(MyIndex).MaxInv
        n = GetPlayerInvItemNum(MyIndex, i)
        
        If n > 0 Then
            ItemName = Trim$(Item(n).Name)
            
            If Item(n).Type = ITEM_TYPE_CURRENCY Or Item(n).Stackable = 1 Then
                frmBank.lstInventory.addItem i & "> " & ItemName & " (" & GetPlayerInvItemValue(MyIndex, i) & ")"
            Else
                frmBank.lstInventory.addItem i & "> " & ItemName
            End If
        Else
            frmBank.lstInventory.addItem i & "> Empty"
        End If
    Next i

    For i = 1 To MAX_BANK
        n = GetPlayerBankItemNum(MyIndex, i)
        
        If n > 0 Then
            ItemName = Trim$(Item(n).Name)
            
            If Item(n).Type = ITEM_TYPE_CURRENCY Or Item(n).Stackable = 1 Then
                frmBank.lstBank.addItem i & "> " & ItemName & " (" & GetPlayerBankItemValue(MyIndex, i) & ")"
            Else
                frmBank.lstBank.addItem i & "> " & ItemName
            End If
        Else
            frmBank.lstBank.addItem i & "> Empty"
        End If
    Next i
    
    frmBank.lstBank.ListIndex = 0
    frmBank.lstInventory.ListIndex = 0
    
    IsBanking = True
    
    Me.Show vbModeless, frmMirage
End Sub
