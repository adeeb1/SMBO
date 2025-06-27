VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Character Finder"
   ClientHeight    =   2130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4560
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   142
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   304
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   2760
      TabIndex        =   5
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton btnFindChar 
      Caption         =   "Find Character!"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   1560
      Width           =   1815
   End
   Begin VB.TextBox txtAccountName 
      Height          =   375
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   720
      Width           =   2415
   End
   Begin VB.TextBox txtCharName 
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label lblAccountName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Character's acc name:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1590
   End
   Begin VB.Label lblCharName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Character name to find:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1650
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type NewPlayerInvRec
    num As Integer
    Value As Long
    Dur As Integer
    Ammo As Integer
End Type

Private Type PlayerInvRec
    num As Integer
    Value As Long
    Ammo As Integer
End Type

Private Type NewBankRec
    num As Integer
    Value As Long
    Dur As Integer
    Ammo As Integer
End Type

Private Type PlayerRec
    ' General
 '090829 Scorpious2k
    Vflag As Byte       ' version flag - always > 127
    Ver As Byte
    SubVer As Byte
    Rel As Byte
 '090829 End
    Name As String * 20
    Guild As String
    GuildAccess As Byte
    Sex As Byte
    Class As Integer
    Sprite As Long
    LEVEL As Integer
    Exp As Long
    Access As Byte
    PK As Byte

    ' Vitals
    HP As Long
    MP As Long
    SP As Long

    ' Stats
    STR As Long
    DEF As Long
    Speed As Long
    Magi As Long
    POINTS As Long

    ' Worn equipment
    ArmorSlot As Integer
    WeaponSlot As Integer
    HelmetSlot As Integer
    ShieldSlot As Integer
    LegsSlot As Integer
    RingSlot As Integer
    NecklaceSlot As Integer

    ' Inventory
    Inv(1 To 24) As NewPlayerInvRec
    Spell(1 To 20) As Integer
    Bank(1 To 50) As NewBankRec

    ' Position and movement
    Map As Integer
 '090829    X As Byte
 '090829    Y As Byte
    X As Integer
    Y As Integer
    Dir As Byte

    TargetNPC As Integer

    Head As Integer
    Body As Integer
    Leg As Integer

    PAPERDOLL As Byte

    MAXHP As Long
    MAXMP As Long
    MAXSP As Long
    InBattle As Boolean
    Turn As Boolean
    OldX As Integer
    OldY As Integer
    HasTurnBased As Boolean
    RecoverTime As Long
    CritHitChance As Double
    BlockChance As Double
    ' Max Stats
    MAXSTR As Long
    MAXDEF As Long
    MAXSpeed As Long
    MAXStache As Long
    Equipment(1 To 7) As PlayerInvRec
    PartyNum As Long
    PartyInvitedBy As Long
    Height As Integer
End Type

Private Player As PlayerRec

Public Sub LoadChar(ByVal Name As String, ByVal FolderDir As String)
    Dim Directory As New FileSystemObject
    Dim folder As folder
    Dim subfolder As folder
    Dim f As Long
    Dim i As Integer
    Dim FileName As String

    Name = LCase$(Name)
    
    Set folder = Directory.GetFolder(FolderDir)
        
    For Each subfolder In folder.SubFolders
        For i = 1 To 3
            On Error Resume Next
        
            FileName = subfolder.Path & "\Char" & i & ".dat"
    
            f = FreeFile
            
            Open FileName For Binary As #f
                Get #f, , Player
            Close #f
            
            If LCase$(Trim$(Player.Name)) = Name Then
                txtAccountName.Text = subfolder.Name
                Exit Sub
            End If
        Next i
    Next
    
    txtAccountName.Text = "None found"
End Sub

Private Sub btnClose_Click()
    End
End Sub

Private Sub btnFindChar_Click()
    ' Make sure the inputted name is larger than 3 characters
    If Len(txtCharName.Text) < 3 Then
        MsgBox "You've entered an invalid character name!", 0, "Character Finder"
        Exit Sub
    End If
    
    Dim shell As Shell32.shell
    Set shell = New Shell32.shell
    
    Dim shellfolder As Shell32.folder
    Set shellfolder = shell.BrowseForFolder(Me.hWnd, "Select a Folder", BIF_RETURNONLYFSDIRS)
    
    Call LoadChar(Trim$(txtCharName.Text), shellfolder.Items.Item.Path)
End Sub
