VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCommands"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Function GetVar(file As String, Header As String, Var As String) As String
    Dim sSpaces As String   ' Max string length
    Dim szReturn As String  ' Return default value if not found

    szReturn = vbNullString

    sSpaces = Space(5000)

    file = App.Path & "\" & file

    Call GetPrivateProfileString(Header, Var, szReturn, sSpaces, Len(sSpaces), file)

    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
End Function

Public Sub PutVar(file As String, Header As String, Var As String, Value As String)
    On Error GoTo PutVar_Error
    Dim fpath As String
    fpath = App.Path & "\" & file
    Call WritePrivateProfileString(Header, Var, Value, fpath)

    On Error GoTo 0
    Exit Sub

PutVar_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure PutVar of Module modDatabase"
End Sub

Public Sub AdminMsg(ByVal Msg As String, ByVal Color As Byte)
    Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            If GetPlayerAccess(i) > 1 Then
                Call SendDataTo(i, "ADMINMSG" & SEP_CHAR & Msg & SEP_CHAR & Color & END_CHAR)
            End If
        End If
    Next i
End Sub

Public Sub GlobalMsg(ByVal Msg As String, ByVal Color As Byte)
    Call SendDataToAll("GLOBALMSG" & SEP_CHAR & Msg & SEP_CHAR & Color & END_CHAR)
End Sub

Public Sub PlayerMsg(ByVal Index As Long, ByVal Msg As String, ByVal Color As Byte)
    Call SendDataTo(Index, "PLAYERMSG" & SEP_CHAR & Msg & SEP_CHAR & Color & END_CHAR)
End Sub
Public Sub OtherMsg(ByVal Index As Long, ByVal Msg As String, ByVal Color As Byte)
    Call SendDataTo(Index, "OTHERMSG" & SEP_CHAR & Msg & SEP_CHAR & Color & END_CHAR)
End Sub
Public Sub MapMsg(ByVal MapNum As Long, ByVal Msg As String, ByVal Color As Byte)
    Call SendDataToMap(MapNum, "MAPMSG" & SEP_CHAR & Msg & SEP_CHAR & Color & END_CHAR)
End Sub

Public Sub AlertMsg(ByVal Index As Long, ByVal Msg As String)
    Call SendDataTo(Index, "ALERTMSG" & SEP_CHAR & Msg & END_CHAR)
    Call CloseSocket(Index)
End Sub

Function GetPlayerLogin(ByVal Index As Long) As String
    GetPlayerLogin = Trim$(Player(Index).Login)
End Function

Function GetPlayerName(ByVal Index As Long) As String
    GetPlayerName = Trim$(Player(Index).Char(Player(Index).CharNum).Name)
End Function

Function GetPlayerGuild(ByVal Index As Long) As String
    GetPlayerGuild = Trim$(Player(Index).Char(Player(Index).CharNum).Guild)
End Function

Function GetPlayerGuildAccess(ByVal Index As Long) As Long
    GetPlayerGuildAccess = Player(Index).Char(Player(Index).CharNum).GuildAccess
End Function

Sub SetPlayerGuildAccess(ByVal Index As Long, ByVal GuildAccess As Long)
    Player(Index).Char(Player(Index).CharNum).GuildAccess = GuildAccess
End Sub

Sub SetPlayerGuild(ByVal Index As Long, ByVal Guild As String)
    Player(Index).Char(Player(Index).CharNum).Guild = Guild
End Sub

Function GetPlayerClass(ByVal Index As Long) As Long
    GetPlayerClass = Player(Index).Char(Player(Index).CharNum).Class
End Function

Sub SetPlayerClass(ByVal Index As Long, ByVal ClassNum As Long)
    Player(Index).Char(Player(Index).CharNum).Class = ClassNum
End Sub

Function GetPlayerClassName(ByVal Index As Long) As String
    GetPlayerClassName = ClassData(GetPlayerClass(Index)).Name
End Function

Function GetPlayerSprite(ByVal Index As Long) As Long
    GetPlayerSprite = Player(Index).Char(Player(Index).CharNum).Sprite
End Function

Sub SetPlayerSprite(ByVal Index As Long, ByVal Sprite As Long)
    Player(Index).Char(Player(Index).CharNum).Sprite = Sprite
End Sub

Function GetPlayerLevel(ByVal Index As Long) As Long
    GetPlayerLevel = Player(Index).Char(Player(Index).CharNum).LEVEL
End Function

Sub SetPlayerLevel(ByVal Index As Long, ByVal LEVEL As Long)
    Player(Index).Char(Player(Index).CharNum).LEVEL = LEVEL
End Sub

Function GetPlayerNextLevel(ByVal Index As Long) As Long
    If GetPlayerLevel(Index) <= MAX_LEVEL Then
        GetPlayerNextLevel = Experience(GetPlayerLevel(Index))
    End If
End Function

Function GetPlayerExp(ByVal Index As Long) As Long
    GetPlayerExp = Player(Index).Char(Player(Index).CharNum).Exp
End Function

Sub SetPlayerExp(ByVal Index As Long, ByVal Exp As Long)
    Player(Index).Char(Player(Index).CharNum).Exp = Exp
End Sub

Function GetPlayerAccess(ByVal Index As Long) As Long
    GetPlayerAccess = Player(Index).Char(Player(Index).CharNum).Access
End Function

Sub SetPlayerAccess(ByVal Index As Long, ByVal Access As Long)
    Player(Index).Char(Player(Index).CharNum).Access = Access
End Sub

Function GetPlayerPK(ByVal Index As Long) As Long
    GetPlayerPK = Player(Index).Char(Player(Index).CharNum).PK
End Function

Sub SetPlayerPK(ByVal Index As Long, ByVal PK As Long)
    Player(Index).Char(Player(Index).CharNum).PK = PK
End Sub

Function GetPlayerHP(ByVal Index As Long) As Long
    GetPlayerHP = Player(Index).Char(Player(Index).CharNum).HP
    If GetPlayerHP >= GetPlayerMaxHP(Index) Then
        GetPlayerHP = GetPlayerMaxHP(Index)
    End If
End Function

Sub SetPlayerHP(ByVal Index As Long, ByVal HP As Long)
    Player(Index).Char(Player(Index).CharNum).HP = HP

    If Player(Index).Char(Player(Index).CharNum).HP > GetPlayerMaxHP(Index) Then
        Player(Index).Char(Player(Index).CharNum).HP = GetPlayerMaxHP(Index)
    End If
    If Player(Index).Char(Player(Index).CharNum).HP < 0 Then
        Player(Index).Char(Player(Index).CharNum).HP = 0
    End If
End Sub

Function GetPlayerMP(ByVal Index As Long) As Long
    GetPlayerMP = Player(Index).Char(Player(Index).CharNum).MP
    If GetPlayerMP >= GetPlayerMaxMP(Index) Then
        GetPlayerMP = GetPlayerMaxMP(Index)
    End If
End Function

Sub SetPlayerMP(ByVal Index As Long, ByVal MP As Long)
    Player(Index).Char(Player(Index).CharNum).MP = MP

    If Player(Index).Char(Player(Index).CharNum).MP > GetPlayerMaxMP(Index) Then
        Player(Index).Char(Player(Index).CharNum).MP = GetPlayerMaxMP(Index)
    End If
    If Player(Index).Char(Player(Index).CharNum).MP < 0 Then
        Player(Index).Char(Player(Index).CharNum).MP = 0
    End If
End Sub

Function GetPlayerSP(ByVal Index As Long) As Long
    GetPlayerSP = Player(Index).Char(Player(Index).CharNum).SP
End Function

Sub SetPlayerSP(ByVal Index As Long, ByVal SP As Long)
    Player(Index).Char(Player(Index).CharNum).SP = SP

    If GetPlayerSP(Index) < 0 Then
        Player(Index).Char(Player(Index).CharNum).SP = 0
    End If

    If GetPlayerSP(Index) > GetPlayerMaxSP(Index) Then
        Player(Index).Char(Player(Index).CharNum).SP = GetPlayerMaxSP(Index)
    End If
End Sub

Function GetPlayerMaxHP(ByVal Index As Long) As Long
    Dim i As Long, CharNum As Long, Add As Long, LvlUp As Long
    
    LvlUp = Val(GetVar("Level Up.ini", GetPlayerName(Index), "HP")) * 5
    Add = 0
    
    For i = 1 To 7
        If GetPlayerEquipSlotNum(Index, i) > 0 Then
            Add = Add + Item(GetPlayerEquipSlotNum(Index, i)).addHP
        End If
    Next i

    CharNum = Player(Index).CharNum
    
    If GetPlayerClass(Index) = 1 Or GetPlayerClass(Index) = 3 Or GetPlayerClass(Index) = 5 Then
        GetPlayerMaxHP = 15 + Add + LvlUp
    Else
        GetPlayerMaxHP = 10 + Add + LvlUp
    End If
End Function

Function GetPlayerMaxMP(ByVal Index As Long) As Long
    Dim i As Long, CharNum As Long, Add As Long, LvlUp As Long
    
    LvlUp = Val(GetVar("Level Up.ini", GetPlayerName(Index), "FP")) * 5
    Add = 0
    
    For i = 1 To 7
        If GetPlayerEquipSlotNum(Index, i) > 0 Then
            Add = Add + Item(GetPlayerEquipSlotNum(Index, i)).addMP
        End If
    Next i

    CharNum = Player(Index).CharNum
    
    If GetPlayerClass(Index) = 4 Then
        GetPlayerMaxMP = 10 + Add + LvlUp
    Else
        GetPlayerMaxMP = 5 + Add + LvlUp
    End If
End Function

Function GetPlayerMaxSP(ByVal Index As Long) As Long
    Dim i As Long, Add As Long, CharNum As Long, ItemNum As Long
    
    Add = 0
    
    For i = 1 To 7
        ItemNum = GetPlayerEquipSlotNum(Index, i)
        If ItemNum > 0 Then
            Add = Add + Item(ItemNum).addSP
        End If
    Next i

    CharNum = Player(Index).CharNum
    
    If GetPlayerClass(Index) = 0 Then
        GetPlayerMaxSP = 35 + (GetPlayerLevel(Index) * addSP.LEVEL) + (GetPlayerSPEED(Index) * addSP.Speed) + Add
    Else
        GetPlayerMaxSP = 30 + (GetPlayerLevel(Index) * addSP.LEVEL) + (GetPlayerSPEED(Index) * addSP.Speed) + Add
    End If
End Function

Function GetClassMaxHP(ByVal ClassNum As Long) As Long
    GetClassMaxHP = (1 + (ClassData(ClassNum).STR \ 2) + ClassData(ClassNum).STR) * 2
End Function

Function GetClassMaxMP(ByVal ClassNum As Long) As Long
    GetClassMaxMP = (1 + (ClassData(ClassNum).Magi \ 2) + ClassData(ClassNum).Magi) * 2
End Function

Function GetClassMaxSP(ByVal ClassNum As Long) As Long
    GetClassMaxSP = (1 + (ClassData(ClassNum).Speed \ 2) + ClassData(ClassNum).Speed) * 2
End Function

Function GetClassSTR(ByVal ClassNum As Long) As Long
    GetClassSTR = ClassData(ClassNum).STR
End Function

Function GetClassDEF(ByVal ClassNum As Long) As Long
    GetClassDEF = ClassData(ClassNum).DEF
End Function

Function GetClassSPEED(ByVal ClassNum As Long) As Long
    GetClassSPEED = ClassData(ClassNum).Speed
End Function

Function GetClassMAGI(ByVal ClassNum As Long) As Long
    GetClassMAGI = ClassData(ClassNum).Magi
End Function

Function GetPlayerSTR(ByVal Index As Long) As Long
    GetPlayerSTR = Player(Index).Char(Player(Index).CharNum).STR
End Function

Sub SetPlayerSTR(ByVal Index As Long, ByVal STR As Long)
    Player(Index).Char(Player(Index).CharNum).STR = STR
End Sub

Function GetPlayerDEF(ByVal Index As Long) As Long
    GetPlayerDEF = Player(Index).Char(Player(Index).CharNum).DEF
End Function

Sub SetPlayerDEF(ByVal Index As Long, ByVal DEF As Long)
    Player(Index).Char(Player(Index).CharNum).DEF = DEF
End Sub

Function GetPlayerSPEED(ByVal Index As Long) As Long
    GetPlayerSPEED = Player(Index).Char(Player(Index).CharNum).Speed
End Function

Sub SetPlayerSPEED(ByVal Index As Long, ByVal Speed As Long)
    Player(Index).Char(Player(Index).CharNum).Speed = Speed
End Sub

Function GetPlayerStache(ByVal Index As Long) As Long
    GetPlayerStache = Player(Index).Char(Player(Index).CharNum).Magi
End Function

Sub SetPlayerStache(ByVal Index As Long, ByVal Stache As Long)
    Player(Index).Char(Player(Index).CharNum).Magi = Stache
End Sub

Function GetPlayerPOINTS(ByVal Index As Long) As Long
    GetPlayerPOINTS = Player(Index).Char(Player(Index).CharNum).POINTS
End Function

Sub SetPlayerPOINTS(ByVal Index As Long, ByVal POINTS As Long)
    Player(Index).Char(Player(Index).CharNum).POINTS = POINTS
End Sub

Function GetPlayerMap(ByVal Index As Long) As Long
    GetPlayerMap = Player(Index).Char(Player(Index).CharNum).Map
End Function

Sub SetPlayerMap(ByVal Index As Long, ByVal MapNum As Long)
    If MapNum > 0 And MapNum <= MAX_MAPS Then
        Player(Index).Char(Player(Index).CharNum).Map = MapNum
    End If
End Sub

Function GetPlayerX(ByVal Index As Long) As Long
    GetPlayerX = Player(Index).Char(Player(Index).CharNum).X
End Function

Sub SetPlayerX(ByVal Index As Long, ByVal X As Long)
    Player(Index).Char(Player(Index).CharNum).X = X
End Sub

Function GetPlayerY(ByVal Index As Long) As Long
    GetPlayerY = Player(Index).Char(Player(Index).CharNum).Y
End Function

Sub SetPlayerY(ByVal Index As Long, ByVal Y As Long)
    Player(Index).Char(Player(Index).CharNum).Y = Y
End Sub

Function GetPlayerDir(ByVal Index As Long) As Long
    GetPlayerDir = Player(Index).Char(Player(Index).CharNum).Dir
End Function

Sub SetPlayerDir(ByVal Index As Long, ByVal Dir As Long)
    Player(Index).Char(Player(Index).CharNum).Dir = Dir
End Sub

Function GetPlayerIP(ByVal Index As Long) As String
    GetPlayerIP = frmServer.Socket(Index).RemoteHostIP
End Function

Function GetPlayerInvItemNum(ByVal Index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemNum = Player(Index).Char(Player(Index).CharNum).Inv(InvSlot).num
End Function

Sub SetPlayerInvItemNum(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemNum As Long)
    Player(Index).Char(Player(Index).CharNum).Inv(InvSlot).num = ItemNum
End Sub

Function GetPlayerInvItemValue(ByVal Index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemValue = Player(Index).Char(Player(Index).CharNum).Inv(InvSlot).Value
End Function

Sub SetPlayerInvItemValue(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemValue As Long)
    Player(Index).Char(Player(Index).CharNum).Inv(InvSlot).Value = ItemValue
End Sub

Function GetPlayerSpell(ByVal Index As Long, ByVal SpellSlot As Long) As Long
    GetPlayerSpell = Player(Index).Char(Player(Index).CharNum).Spell(SpellSlot)
End Function

Sub SetPlayerSpell(ByVal Index As Long, ByVal SpellSlot As Long, ByVal SpellNum As Long)
    Player(Index).Char(Player(Index).CharNum).Spell(SpellSlot) = SpellNum
End Sub

Function GetMapBootMap(ByVal Index As Long) As Long
    GetMapBootMap = Map(GetPlayerMap(Index)).BootMap
End Function

Function GetMapBootX(ByVal Index As Long) As Long
    GetMapBootX = Map(GetPlayerMap(Index)).BootX
End Function

Function GetMapBootY(ByVal Index As Long) As Long
    GetMapBootY = Map(GetPlayerMap(Index)).BootY
End Function

Sub MutePlayer(ByVal Index As Integer)
    Player(Index).Mute = True
End Sub

Sub UnMutePlayer(ByVal Index As Integer)
    Player(Index).Mute = False
End Sub

Public Sub MakeDay()
    GameTime = TIME_DAY

    Hours = 7
    Minutes = 0
    Seconds = 0

    Call SendTimeToAll
End Sub

Public Sub MakeNight()
    GameTime = TIME_NIGHT

    Hours = 21
    Minutes = 0
    Seconds = 0

    Call SendTimeToAll
End Sub

Function IsScrolling() As Boolean
    If IS_SCROLLING = 0 Then
        IsScrolling = False
    Else
        IsScrolling = True
    End If
End Function

Function GetMaxPlayers() As Long
    GetMaxPlayers = MAX_PLAYERS
End Function

Sub SpawnItemSlot(ByVal MapItemSlot As Long, ByVal ItemNum As Long, ByVal ItemVal As Long, ByVal ItemAmmo As Long, ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long)
    Dim packet As String
    Dim i As Long

    ' Check for subscript out of range
    If MapItemSlot <= 0 Or MapItemSlot > MAX_MAP_ITEMS Or ItemNum < 0 Or ItemNum > MAX_ITEMS Or MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If

    i = MapItemSlot

    If i <> 0 Then
        If ItemNum >= 0 Then
            If ItemNum <= MAX_ITEMS Then
                MapItem(MapNum, i).num = ItemNum
                MapItem(MapNum, i).Value = ItemVal
        
                If ItemNum <> 0 Then
                    If (Item(ItemNum).Type >= ITEM_TYPE_WEAPON) And (Item(ItemNum).Type <= ITEM_TYPE_MUSHROOMBADGE) Then
                        MapItem(MapNum, i).Ammo = ItemAmmo
                    Else
                        MapItem(MapNum, i).Ammo = -1
                    End If
                Else
                    MapItem(MapNum, i).Ammo = -1
                End If
        
                MapItem(MapNum, i).X = X
                MapItem(MapNum, i).Y = Y
        
                packet = "SPAWNITEM" & SEP_CHAR & i & SEP_CHAR & ItemNum & SEP_CHAR & ItemVal & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & MapItem(MapNum, i).Ammo & END_CHAR
                Call SendDataToMap(MapNum, packet)
            End If
        End If
    End If
End Sub

Function IsConnected(ByVal Index As Long) As Boolean
    If frmServer.Socket(Index).State = sckConnected Then
        IsConnected = True
    Else
        IsConnected = False
    End If
End Function

Function IsPlaying(ByVal Index As Long) As Boolean
    If IsConnected(Index) Then
        If Player(Index).InGame Then
            IsPlaying = True
        Else
            IsPlaying = False
        End If
    Else
        IsPlaying = False
    End If
End Function

Sub SendInventory(ByVal Index As Long)
    Dim packet As String
    Dim i As Long

    packet = "PLAYERINV" & SEP_CHAR & Index & SEP_CHAR
    For i = 1 To MAX_INV
        packet = packet & GetPlayerInvItemNum(Index, i) & SEP_CHAR & GetPlayerInvItemValue(Index, i) & SEP_CHAR & GetPlayerInvItemAmmo(Index, i) & SEP_CHAR
    Next i
    packet = packet & END_CHAR

    Call SendDataToMap(GetPlayerMap(Index), packet)
End Sub

Sub SendInventoryUpdate(ByVal Index As Long, ByVal InvSlot As Long)
    Call SendDataToMap(GetPlayerMap(Index), "PLAYERINVUPDATE" & SEP_CHAR & InvSlot & SEP_CHAR & Index & SEP_CHAR & GetPlayerInvItemNum(Index, InvSlot) & SEP_CHAR & GetPlayerInvItemValue(Index, InvSlot) & SEP_CHAR & GetPlayerInvItemAmmo(Index, InvSlot) & END_CHAR)
End Sub

Sub SendIndexInventoryFromMap(ByVal Index As Long)
    Dim packet As String
    Dim n As Long
    Dim i As Long
    
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            If GetPlayerMap(i) = GetPlayerMap(Index) Then
                packet = "PLAYERINV" & SEP_CHAR & i & SEP_CHAR
                For n = 1 To MAX_INV
                    packet = packet & GetPlayerInvItemNum(i, n) & SEP_CHAR & GetPlayerInvItemValue(i, n) & SEP_CHAR & GetPlayerInvItemAmmo(i, n) & SEP_CHAR
                Next n
                packet = packet & END_CHAR

                Call SendDataTo(Index, packet)
            End If
        End If
    Next i
End Sub

Sub SendWornEquipment(ByVal Index As Long)
    Dim i As Long
    Dim packet As String, Equipment As String
    
    If IsPlaying(Index) Then
        packet = "PLAYERWORNEQ" & SEP_CHAR & Index
        For i = 1 To 7
            packet = packet & SEP_CHAR & GetPlayerEquipSlotNum(Index, i)
        Next i
        packet = packet & END_CHAR
        Call SendDataToMap(GetPlayerMap(Index), packet)
    End If
End Sub

Sub SendHP(ByVal Index As Long)
    Dim Exp As Long, i As Long
  
    Call SendDataTo(Index, "playerhp" & SEP_CHAR & Index & SEP_CHAR & GetPlayerMaxHP(Index) & SEP_CHAR & GetPlayerHP(Index) & END_CHAR)
    If GetPlayerGuild(Index) <> vbNullString Then
        Call SendGuildMemberHP(Index)
    End If

    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type = TILE_TYPE_LOWER_STAT Or Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type = TILE_TYPE_KILL Or Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type = TILE_TYPE_HEAL Then
      ' Kill player and restore stats
        If GetPlayerHP(Index) < 1 Then
            ' Triggers Life Shroom
            If HasItem(Index, 43) = 1 Then
                Call LifeShroom(Index)
                Exit Sub
            End If
            Call PlayerDeath(Index)
        End If
    End If
    
    ' Handle sounds
    If GetPlayerHP(Index) > 0 And GetPlayerHP(Index) <= 5 And Val(GetVar(App.Path & "\Sounds.ini", GetPlayerName(Index), "Sound")) = 0 Then
        Call PutVar(App.Path & "\Sounds.ini", GetPlayerName(Index), "Sound", "1")
        Call SendDataTo(Index, "sound" & SEP_CHAR & "soundattribute" & SEP_CHAR & "sms_lowhealth.wav" & END_CHAR)
    ElseIf GetPlayerHP(Index) > 5 And Val(GetVar(App.Path & "\Sounds.ini", GetPlayerName(Index), "Sound")) = 1 Then
        Call PutVar(App.Path & "\Sounds.ini", GetPlayerName(Index), "Sound", "0")
    End If
End Sub

Sub SendMP(ByVal Index As Long)
    Call SendDataTo(Index, "PLAYERMP" & SEP_CHAR & GetPlayerMaxMP(Index) & SEP_CHAR & GetPlayerMP(Index) & END_CHAR)
End Sub

Sub SendSP(ByVal Index As Long)
    Call SendDataTo(Index, "PLAYERSP" & SEP_CHAR & GetPlayerMaxSP(Index) & SEP_CHAR & GetPlayerSP(Index) & END_CHAR)
End Sub

Sub SendPTS(ByVal Index As Long)
    Call SendDataTo(Index, "PLAYERPOINTS" & SEP_CHAR & GetPlayerPOINTS(Index) & END_CHAR)
End Sub

Sub SendStats(ByVal Index As Long)
    Call modServerTCP.SendStats(Index)
End Sub

Sub PlaySound(ByVal Index As Long, ByVal Sound As String)
    Call SendDataTo(Index, "sound" & SEP_CHAR & "soundattribute" & SEP_CHAR & Sound & END_CHAR)
End Sub

Sub PlaySoundToMap(ByVal Index As Long, ByVal Sound As String)
    Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "soundattribute" & SEP_CHAR & Sound & END_CHAR)
End Sub

Sub SendPlayerData(ByVal Index As Long)
    Dim packet As String

    ' Send index's player data to everyone including himself on the map
    packet = "PLAYERDATA" & SEP_CHAR
    packet = packet & Index & SEP_CHAR
    packet = packet & GetPlayerName(Index) & SEP_CHAR
    packet = packet & GetPlayerSprite(Index) & SEP_CHAR
    packet = packet & GetPlayerMap(Index) & SEP_CHAR
    packet = packet & GetPlayerX(Index) & SEP_CHAR
    packet = packet & GetPlayerY(Index) & SEP_CHAR
    packet = packet & GetPlayerDir(Index) & SEP_CHAR
    packet = packet & GetPlayerAccess(Index) & SEP_CHAR
    packet = packet & GetPlayerPK(Index) & SEP_CHAR
    packet = packet & GetPlayerGuild(Index) & SEP_CHAR
    packet = packet & GetPlayerGuildAccess(Index) & SEP_CHAR
    packet = packet & GetPlayerClass(Index) & SEP_CHAR
    packet = packet & GetPlayerLevel(Index) & END_CHAR
    
    Call SendDataToMap(GetPlayerMap(Index), packet)
    Call SendGuildMemberHP(Index)
End Sub

Sub SendDataTo(ByVal Index As Long, ByVal Data As String)
    Dim dbytes() As Byte

    dbytes = StrConv(Data, vbFromUnicode)
    If IsConnected(Index) Then
        frmServer.Socket(Index).SendData dbytes
    End If
End Sub

Sub SendDataToAll(ByVal Data As String)
    Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            Call SendDataTo(i, Data)
        End If
    Next i
End Sub

Sub SendDataToAllBut(ByVal Index As Long, ByVal Data As String)
    Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            If i <> Index Then
                Call SendDataTo(i, Data)
            End If
        End If
    Next i
End Sub

Sub SendDataToMap(ByVal MapNum As Long, ByVal Data As String)
    Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            If GetPlayerMap(i) = MapNum Then
                Call SendDataTo(i, Data)
            End If
        End If
    Next i
End Sub

Sub SendDataToMapBut(ByVal Index As Long, ByVal MapNum As Long, ByVal Data As String)
    Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            If GetPlayerMap(i) = MapNum And i <> Index Then
                Call SendDataTo(i, Data)
            End If
        End If
    Next i
End Sub

Sub SetPlayerName(ByVal Index As Long, ByVal Name As String)
    Player(Index).Char(Player(Index).CharNum).Name = Name
End Sub

Function GetPlayerCharNum(ByVal Index As Long) As Long
    GetPlayerCharNum = Player(Index).CharNum
End Function

Function FindPlayer(ByVal Name As String) As Long
    Dim i As Long

    Name = LCase$(Name)

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            If Len(GetPlayerName(i)) >= Len(Name) Then
                If LCase$(GetPlayerName(i)) = Name Then
                    FindPlayer = i
                    Exit Function
                End If
            End If
        End If
    Next i
End Function

Public Sub PlayerWarp(ByVal Index As Long, ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long)
    Dim OldMap As Long

    On Error GoTo WarpErr

    If Not IsPlaying(Index) Then
        Exit Sub
    End If

    If MapNum < 1 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If

    ' Save old map to send erase player data to
    OldMap = GetPlayerMap(Index)

    If Not OldMap = MapNum Then
        Call SendLeaveMap(Index, OldMap)
    End If

    Call SetPlayerMap(Index, MapNum)
    Call SetPlayerX(Index, X)
    Call SetPlayerY(Index, Y)

    If GetTotalMapPlayers(OldMap) = 0 Then
        PlayersOnMap(OldMap) = NO
    End If

    PlayersOnMap(MapNum) = YES

    Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "warp" & END_CHAR)

    Player(Index).GettingMap = YES

    Call SendDataTo(Index, "CHECKFORMAP" & SEP_CHAR & MapNum & SEP_CHAR & Map(MapNum).Revision & END_CHAR)

    Call SendInventory(Index)
    Call SendIndexInventoryFromMap(Index)
    Call SendIndexWornEquipmentFromMap(Index)
    Call SendMapPlayersInBattle(Index, MapNum)
    
    Exit Sub

WarpErr:
    Call AddLog("PlayerWarp error for player index " & Index & " on map " & GetPlayerMap(Index) & ".", "logs\ErrorLog.txt")
End Sub

Sub JoinWarp(ByVal Index As Long, ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long)
    Dim OldMap As Long

    If Not IsPlaying(Index) Then
        Exit Sub
    End If

    If MapNum < 1 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If

    ' Save old map to send erase player data to
    OldMap = GetPlayerMap(Index)

    Call SendLeaveMap(Index, OldMap)

    Call SetPlayerMap(Index, MapNum)
    Call SetPlayerX(Index, X)
    Call SetPlayerY(Index, Y)

    If GetTotalMapPlayers(OldMap) = 0 Then
        PlayersOnMap(OldMap) = NO
    End If

    PlayersOnMap(MapNum) = YES

    Player(Index).GettingMap = YES

    Call SendDataTo(Index, "CHECKFORMAP" & SEP_CHAR & MapNum & SEP_CHAR & Map(MapNum).Revision & END_CHAR)

    Call SendInventory(Index)
    Call SendIndexWornEquipmentFromMap(Index)
End Sub

Public Sub AddLog(ByVal Text As String, ByVal FN As String)
    Dim FileName As String
    Dim FileID As Long

    If ServerLog Then
        FileName = App.Path & "\" & FN

        If FileExists(FN) Then
            FileID = FreeFile
            Open FileName For Output As #FileID
            Print #FileID, Time & ": " & Text
            Close #FileID
        End If
    End If
End Sub

Sub HackingAttempt(ByVal Index As Long, ByVal Reason As String)
    If Index > 0 Then
        If IsPlaying(Index) Then
            Call AdminMsg(GetPlayerLogin(Index) & "/" & GetPlayerName(Index) & " has been booted for (" & Reason & ")", WHITE)
            Call AddLog(GetPlayerName(Index) & " was kicked for '" & Reason & "'.", "Logs\HackAttempt.txt")
        End If

        Call AlertMsg(Index, "You have lost your connection with " & GAME_NAME & ".")
    End If
End Sub

Sub BattleMsg(ByVal Index As Long, ByVal Msg As String, ByVal Color As Byte, ByVal Side As Long)
    Call SendDataTo(Index, "damagedisplay" & SEP_CHAR & Side & SEP_CHAR & Msg & SEP_CHAR & Color & END_CHAR)
End Sub

Public Function Rand(ByVal Low As Long, ByVal High As Long) As Long
    Rand = Int((High - Low + 1) * Rnd) + Low
End Function
Function GetPlayerBankItemNum(ByVal Index As Long, ByVal BankSlot As Long) As Long
    GetPlayerBankItemNum = Player(Index).Char(Player(Index).CharNum).Bank(BankSlot).num
End Function

Sub SetPlayerBankItemNum(ByVal Index As Long, ByVal BankSlot As Long, ByVal ItemNum As Long)
    Player(Index).Char(Player(Index).CharNum).Bank(BankSlot).num = ItemNum
    Call SendBankUpdate(Index, BankSlot)
End Sub

Function GetPlayerBankItemValue(ByVal Index As Long, ByVal BankSlot As Long) As Long
    GetPlayerBankItemValue = Player(Index).Char(Player(Index).CharNum).Bank(BankSlot).Value
End Function

Sub SetPlayerBankItemValue(ByVal Index As Long, ByVal BankSlot As Long, ByVal ItemValue As Long)
    Player(Index).Char(Player(Index).CharNum).Bank(BankSlot).Value = ItemValue
    Call SendBankUpdate(Index, BankSlot)
End Sub

Function GetPlayerTarget(ByVal Index As Long)
    If Player(Index).TargetType = TARGET_TYPE_PLAYER Then
        GetPlayerTarget = Player(Index).Target
    Else
        GetPlayerTarget = -1
    End If
End Function

Sub SetTile(ByVal mapper As Long, ByVal X As Long, ByVal Y As Long, ByVal xset As Long, ByVal yset As Long, ByVal tileset As Long, ByVal layer As Long)
    Call ScriptSetTile(mapper, X, Y, xset, yset, tileset, layer)
End Sub

Function GetTileX(ByVal mapper As Long, ByVal X As Long, ByVal Y As Long, ByVal layer As Long)
    Select Case layer
        Case 0
            GetTileX = Map(mapper).Tile(X, Y).Ground - (Map(mapper).Tile(X, Y).Ground \ 14) * 14
        Case 1
            GetTileX = Map(mapper).Tile(X, Y).Mask - (Map(mapper).Tile(X, Y).Mask \ 14) * 14
        Case 2
            GetTileX = Map(mapper).Tile(X, Y).Anim - (Map(mapper).Tile(X, Y).Anim \ 14) * 14
        Case 3
            GetTileX = Map(mapper).Tile(X, Y).Mask2 - (Map(mapper).Tile(X, Y).Mask2 \ 14) * 14
        Case 4
            GetTileX = Map(mapper).Tile(X, Y).M2Anim - (Map(mapper).Tile(X, Y).M2Anim \ 14) * 14
        Case 5
            GetTileX = Map(mapper).Tile(X, Y).Fringe - (Map(mapper).Tile(X, Y).Fringe \ 14) * 14
        Case 6
            GetTileX = Map(mapper).Tile(X, Y).FAnim - (Map(mapper).Tile(X, Y).FAnim \ 14) * 14
        Case 7
            GetTileX = Map(mapper).Tile(X, Y).Fringe2 - (Map(mapper).Tile(X, Y).Fringe2 \ 14) * 14
        Case 8
            GetTileX = Map(mapper).Tile(X, Y).F2Anim - (Map(mapper).Tile(X, Y).F2Anim \ 14) * 14
    End Select
End Function
Function GetTileY(ByVal mapper As Long, ByVal X As Long, ByVal Y As Long, ByVal layer As Long)
    Select Case layer
        Case 0
            GetTileY = (Map(mapper).Tile(X, Y).Ground \ 14)
        Case 1
            GetTileY = (Map(mapper).Tile(X, Y).Mask \ 14)
        Case 2
            GetTileY = (Map(mapper).Tile(X, Y).Anim \ 14)
        Case 3
            GetTileY = (Map(mapper).Tile(X, Y).Mask2 \ 14)
        Case 4
            GetTileY = (Map(mapper).Tile(X, Y).M2Anim \ 14)
        Case 5
            GetTileY = (Map(mapper).Tile(X, Y).Fringe \ 14)
        Case 6
            GetTileY = (Map(mapper).Tile(X, Y).FAnim \ 14)
        Case 7
            GetTileY = (Map(mapper).Tile(X, Y).Fringe2 \ 14)
        Case 8
            GetTileY = (Map(mapper).Tile(X, Y).F2Anim \ 14)
    End Select
End Function
Function GetTileSet(ByVal mapper As Long, ByVal X As Long, ByVal Y As Long, ByVal layer As Long)
    Select Case layer
        Case 0
            GetTileSet = Map(mapper).Tile(X, Y).GroundSet
        Case 1
            GetTileSet = Map(mapper).Tile(X, Y).MaskSet
        Case 2
            GetTileSet = Map(mapper).Tile(X, Y).AnimSet
        Case 3
            GetTileSet = Map(mapper).Tile(X, Y).Mask2Set
        Case 4
            GetTileSet = Map(mapper).Tile(X, Y).M2AnimSet
        Case 5
            GetTileSet = Map(mapper).Tile(X, Y).FringeSet
        Case 6
            GetTileSet = Map(mapper).Tile(X, Y).FAnimSet
        Case 7
            GetTileSet = Map(mapper).Tile(X, Y).Fringe2Set
        Case 8
            GetTileSet = Map(mapper).Tile(X, Y).F2AnimSet
    End Select
End Function

Sub SendMap(ByVal mapper As Long)
    Call SendDataToMap(mapper, "CHECKFORMAP" & SEP_CHAR & mapper & SEP_CHAR & (Map(mapper).Revision + 1) & END_CHAR)
End Sub

Sub SpellAnim(ByVal SpellNum As Long, ByVal mapper As Long, ByVal X As Long, ByVal Y As Long)
    Call SendDataToMap(mapper, "scriptspellanim" & SEP_CHAR & SpellNum & SEP_CHAR & Spell(SpellNum).SpellAnim & SEP_CHAR & Spell(SpellNum).SpellTime & SEP_CHAR & Spell(SpellNum).SpellDone & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & Spell(SpellNum).Big & END_CHAR)
End Sub
Sub SetAttribute(ByVal mapper As Long, ByVal X As Long, ByVal Y As Long, ByVal Attrib As Long, ByVal Data1 As Long, ByVal Data2 As Long, ByVal Data3 As Long, ByVal String1 As String, ByVal String2 As String, ByVal String3 As String)
    Call ScriptSetAttribute(mapper, X, Y, Attrib, Data1, Data2, Data3, String1, String2, String3)
End Sub
Function GetAttribute(ByVal mapper As Long, ByVal X As Long, ByVal Y As Long)
    GetAttribute = Map(mapper).Tile(X, Y).Type
End Function
Function GetTileData1(ByVal mapper As Long, ByVal X As Long, ByVal Y As Long)
    GetTileData1 = Map(mapper).Tile(X, Y).Data1
End Function
Function GetTileData2(ByVal mapper As Long, ByVal X As Long, ByVal Y As Long)
    GetTileData2 = Map(mapper).Tile(X, Y).Data2
End Function
Function GetTileData3(ByVal mapper As Long, ByVal X As Long, ByVal Y As Long)
    GetTileData3 = Map(mapper).Tile(X, Y).Data3
End Function
Function GetTileString1(ByVal mapper As Long, ByVal X As Long, ByVal Y As Long)
    GetTileString1 = Map(mapper).Tile(X, Y).String1
End Function
Function GetTileString2(ByVal mapper As Long, ByVal X As Long, ByVal Y As Long)
    GetTileString2 = Map(mapper).Tile(X, Y).String2
End Function
Function GetTileString3(ByVal mapper As Long, ByVal X As Long, ByVal Y As Long)
    GetTileString3 = Map(mapper).Tile(X, Y).String3
End Function

Sub SetMapNpcNumber(ByVal MapNum As Long, ByVal Index As Long, ByVal Number As Long)
    If Number >= 1 Then
        MapNPC(MapNum, Index).num = Number
    End If
End Sub

Sub SetMapNpcTarget(ByVal MapNum As Long, ByVal Index As Long, ByVal Target As Long)
    MapNPC(MapNum, Index).Target = Target
End Sub

Public Sub SetMapName(ByVal MapNum As Long, ByVal MapName As String)
    Dim i As Long

    Map(MapNum).Name = MapName
    Map(MapNum).Revision = Map(MapNum).Revision + 1

    If GetTotalMapPlayers(MapNum) > 0 Then
        For i = 1 To MAX_PLAYERS
            If IsPlaying(i) Then
                If GetPlayerMap(i) = MapNum Then
                    Call SendMap(MapNum)
                End If
            End If
        Next i
    End If
End Sub

Sub SetMapNpcDir(ByVal MapNum As Long, ByVal Index As Long, ByVal Direction As Long)
    MapNPC(MapNum, Index).Dir = Direction
End Sub

Sub SetMapNpcY(ByVal MapNum As Long, ByVal Index As Long, ByVal NPC_Y As Long)
    If NPC_Y <= MAX_MAPY And NPC_Y >= 0 Then
        MapNPC(MapNum, Index).Y = NPC_Y
    End If
End Sub

Sub SetMapNpcX(ByVal MapNum As Long, ByVal Index As Long, ByVal NPC_X As Long)
    If NPC_X <= MAX_MAPX And NPC_X >= 0 Then
        MapNPC(MapNum, Index).X = NPC_X
    End If
End Sub

Sub SetMapNpcHP(ByVal MapNum As Long, ByVal Index As Long, ByVal HitPoints As Long)
    MapNPC(MapNum, Index).HP = HitPoints
End Sub

Sub SendNPC(ByVal MapNum As Long, ByVal MapNpcNum As Long)
    Call SendDataToMap(MapNum, "SPAWNNPC" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNPC(MapNum, MapNpcNum).num & SEP_CHAR & MapNPC(MapNum, MapNpcNum).X & SEP_CHAR & MapNPC(MapNum, MapNpcNum).Y & SEP_CHAR & MapNPC(MapNum, MapNpcNum).Dir & SEP_CHAR & NPC(MapNPC(MapNum, MapNpcNum).num).Big & END_CHAR)
End Sub

Function GetNpcMaxHP(ByVal Index As Long) As Long
    GetNpcMaxHP = NPC(Index).MAXHP
End Function

Function GetMapNpcNumber(ByVal MapNum As Long, ByVal Index As Long) As Long
    GetMapNpcNumber = MapNPC(MapNum, Index).num
End Function

Function GetMapNpcHP(ByVal MapNum As Long, ByVal Index As Long) As Long
    GetMapNpcHP = MapNPC(MapNum, Index).HP
End Function

Function GetNpcName(ByVal Number As Long) As String
    GetNpcName = Trim$(NPC(Number).Name)
End Function

Function GetNpcBehavior(ByVal Number As Long) As Long
    GetNpcBehavior = NPC(Number).Behavior
End Function

Function GetNpcExp(ByVal Number As Long) As Long
    GetNpcExp = NPC(Number).Exp
End Function

Function GetNpcDefense(ByVal Number As Long) As Long
    GetNpcDefense = NPC(Number).DEF
End Function

Function GetNpcStrength(ByVal Number As Long) As Long
    GetNpcStrength = NPC(Number).STR
End Function

Sub SendIndexWornEquipment(ByVal Index As Long)
    Dim packet As String
    Dim i As Long

    packet = "itemworn" & SEP_CHAR & Index

    For i = 1 To 7
        If GetPlayerEquipSlotNum(Index, i) > 0 Then
            packet = packet & SEP_CHAR & GetPlayerEquipSlotNum(Index, i)
        Else
            packet = packet & SEP_CHAR & 0
        End If
    Next i

    packet = packet & END_CHAR
    Call SendDataToMap(GetPlayerMap(Index), packet)
End Sub

Sub SendIndexWornEquipmentFromMap(ByVal Index As Long)
    Dim packet As String
    Dim i As Long, X As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) = True Then
            If GetPlayerMap(Index) = GetPlayerMap(i) Then
                packet = "itemworn" & SEP_CHAR & i

                For X = 1 To 7
                    packet = packet & SEP_CHAR & GetPlayerEquipSlotNum(i, X)
                Next X
                    
                Call SendDataTo(Index, packet)
            End If
        End If
    Next i
End Sub

Function GetPlayersOnMap(ByVal MapNum As Long) As Long
    Dim i As Long
    Dim Count As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            If GetPlayerMap(i) = MapNum Then
                Count = Count + 1
            End If
        End If
    Next i

    GetPlayersOnMap = Count
End Function

Sub ScriptSpawnNpc(ByVal MapNpcNum As Long, ByVal MapNum As Long, ByVal spawn_X As Long, ByVal spawn_Y As Long, ByVal NPCNum As Long)
    ' NPC_index               map_number          X spawn          Y spawn            NPC_number
    Dim packet As String
    Dim i As Long

    ' Check for subscript out of range
    If MapNpcNum < 0 Or MapNpcNum > MAX_MAP_NPCS Or MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If

    If NPCNum = 0 Then
        Map(MapNum).Revision = Map(MapNum).Revision + 1
        MapNPC(MapNum, MapNpcNum).num = 0
        Map(MapNum).NPC(MapNpcNum) = 0
        MapNPC(MapNum, MapNpcNum).Target = 0
        MapNPC(MapNum, MapNpcNum).HP = 0
        MapNPC(MapNum, MapNpcNum).MP = 0
        MapNPC(MapNum, MapNpcNum).SP = 0
        MapNPC(MapNum, MapNpcNum).Dir = 0
        MapNPC(MapNum, MapNpcNum).X = 0
        MapNPC(MapNum, MapNpcNum).Y = 0

        ' Packet = "SPAWNNPC" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(mapnum, MapNpcNum).num & SEP_CHAR & MapNpc(mapnum, MapNpcNum).X & SEP_CHAR & MapNpc(mapnum, MapNpcNum).Y & SEP_CHAR & MapNpc(mapnum, MapNpcNum).Dir & SEP_CHAR & Npc(MapNpc(mapnum, MapNpcNum).num).Big & END_CHAR
        ' Call SendDataToMap(mapnum, Packet)
        Call SaveMap(MapNum)
    End If

' MapNpc(mapnum, MapNpcNum).num = 0
' MapNpc(mapnum, MapNpcNum).SpawnWait = GetTickCount
' MapNpc(mapnum, MapNpcNum).HP = 0
' Call SendDataToMap(mapnum, "NPCDEAD" & SEP_CHAR & MapNpcNum & END_CHAR)


    Map(MapNum).Revision = Map(MapNum).Revision + 1

    MapNPC(MapNum, MapNpcNum).num = NPCNum
    Map(MapNum).NPC(MapNpcNum) = NPCNum

    MapNPC(MapNum, MapNpcNum).Target = 0

    MapNPC(MapNum, MapNpcNum).HP = GetNpcMaxHP(NPCNum)
    MapNPC(MapNum, MapNpcNum).MP = GetNpcMaxMP(NPCNum)
    MapNPC(MapNum, MapNpcNum).SP = GetNpcMaxSP(NPCNum)

    MapNPC(MapNum, MapNpcNum).Dir = Int(Rnd * 4)

    MapNPC(MapNum, MapNpcNum).X = spawn_X
    MapNPC(MapNum, MapNpcNum).Y = spawn_Y

    packet = "SPAWNNPC" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNPC(MapNum, MapNpcNum).num & SEP_CHAR & MapNPC(MapNum, MapNpcNum).X & SEP_CHAR & MapNPC(MapNum, MapNpcNum).Y & SEP_CHAR & MapNPC(MapNum, MapNpcNum).Dir & SEP_CHAR & NPC(MapNPC(MapNum, MapNpcNum).num).Big & END_CHAR
    Call SendDataToMap(MapNum, packet)

    'Call SaveMap(MapNum)

    'For i = 1 To MAX_PLAYERS
    '    If IsPlaying(i) And GetPlayerMap(i) = MapNum Then
    '        Call SendDataTo(i, "CHECKFORMAP" & SEP_CHAR & GetPlayerMap(i) & SEP_CHAR & Map(GetPlayerMap(i)).Revision & END_CHAR)
    '    End If
    'Next i

End Sub

Sub SpawnItemOnMap(ByVal Map As Long, ByVal X As Long, ByVal Y As Long, ByVal ItemNum As Long, ByVal Amount As Long, ByVal Ammo As Long)
    Dim i As Long

    If ItemNum > MAX_ITEMS Or ItemNum = 0 Then
        Exit Sub
    End If

    If Map < 0 Or Map > MAX_MAPS Then
        Exit Sub
    End If

    If Amount < 0 Then
        Exit Sub
    End If

    i = FindOpenMapItemSlot(Map)

    If i <> 0 Then

        Call SpawnItemSlot(i, ItemNum, Amount, Ammo, Map, X, Y)
    Else
        i = MAX_MAP_ITEMS
        Call SpawnItemSlot(i, ItemNum, Amount, Ammo, Map, X, Y)
    End If
End Sub

Function GetItemName(ByVal Number As Long)
    GetItemName = Item(Number).Name
End Function

Sub ClearItemSlot(ByVal Map As Long, ByVal ItemIndex As Long)
    Call SpawnItemSlot(ItemIndex, 0, 0, -1, Map, MapItem(Map, ItemIndex).X, MapItem(Map, ItemIndex).Y)
End Sub

Sub SendWhosOnline(ByVal Index As Integer)
    Dim S As String
    Dim n As Long, i As Long

    S = vbNullString
    n = 0
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) And i <> Index Then
            S = S & GetPlayerName(i) & ", "
            n = n + 1
        End If
    Next i

    If n = 0 Then
        S = "There are no other players online."
    Else
        S = Mid$(S, 1, Len(S) - 2)
        S = "There are " & n & " other players online: " & S & "."
    End If

    Call PlayerMsg(Index, S, WhoColor)
End Sub

Sub LockPlayer(ByVal Index As Long, ByVal Locked As Long)
    If Locked = 0 Then
        Player(Index).Locked = False
        Call SendDataTo(Index, "CHECKFORMAP" & SEP_CHAR & GetPlayerMap(Index) & SEP_CHAR & Map(GetPlayerMap(Index)).Revision & END_CHAR)
    End If

    If Locked = 1 Then
        Player(Index).Locked = True
        Call SendDataTo(Index, "CHECKFORMAP" & SEP_CHAR & GetPlayerMap(Index) & SEP_CHAR & Map(GetPlayerMap(Index)).Revision & END_CHAR)
    End If
End Sub

Function IsPlayerLocked(ByVal Index As Long) As Boolean
    IsPlayerLocked = Player(Index).Locked
End Function

Public Sub SetServerTime(ByVal Time_Hour As Long, ByVal Time_Minute As Long, ByVal Time_Second As Long)
    If Time_Hour <> -1 Then
        If Time_Hour > -1 And Time_Hour < 24 Then
            Hours = Time_Hour
        End If
    End If

    If Time_Minute <> -1 Then
        If Time_Minute > -1 And Time_Minute < 60 Then
            Minutes = Time_Minute
        End If
    End If

    If Time_Second <> -1 Then
        If Time_Second > -1 And Time_Second < 60 Then
            Seconds = Time_Second
        End If
    End If
End Sub

Function GetServerTime(ByVal TimeType As Long) As Long
    Select Case TimeType
        Case 0
            GetServerTime = Hours

        Case 1
            GetServerTime = Minutes

        Case 2
            GetServerTime = Seconds
    End Select
End Function

Function GetMapName(F_map)
    GetMapName = Map(F_map).Name
End Function

Function GetMapUp(F_map)
    GetMapUp = Map(F_map).Up
End Function

Function GetMapDown(F_map)
    GetMapDown = Map(F_map).Down
End Function

Function GetMapLeft(F_map)
    GetMapLeft = Map(F_map).Left
End Function

Function GetMapRight(F_map)
    GetMapRight = Map(F_map).Right
End Function

Sub npcmoving(ByVal MapNum As Long, ByVal MapNpcNum As Long, ByVal Direction As Long, ByVal Speed As Long)
    If CanNpcMove(MapNum, MapNpcNum, Direction) Then
        Call NpcMove(MapNum, MapNpcNum, Direction, Speed)
    End If
End Sub

Function GetPlayerDirX(Index)
    Select Case GetPlayerDir(Index)
        Case 0
            GetPlayerDirX = GetPlayerX(Index)
        Case 1
            GetPlayerDirX = GetPlayerX(Index)
        Case 2
            GetPlayerDirX = GetPlayerX(Index) - 1
        Case 3
            GetPlayerDirX = GetPlayerX(Index) + 1
    End Select
End Function

Function GetPlayerDirY(Index)
    Select Case GetPlayerDir(Index)
        Case 0
            GetPlayerDirY = GetPlayerY(Index) - 1
        Case 1
            GetPlayerDirY = GetPlayerY(Index) + 1
        Case 2
            GetPlayerDirY = GetPlayerY(Index)
        Case 3
            GetPlayerDirY = GetPlayerY(Index)
    End Select
End Function

Function GetNpcX(ByVal MapNum As Long, ByVal MapNpcNum As Long)
    If MapNpcNum < 1 Or MapNpcNum > 25 Then
        Exit Function
    End If
    
    GetNpcX = MapNPC(MapNum, MapNpcNum).X
End Function

Function GetPlayerTargetNpc(ByVal Index As Long)
    If Index > 0 Then
        If Player(Index).TargetType = TARGET_TYPE_NPC Then
            GetPlayerTargetNpc = Player(Index).TargetNPC
        Else
            GetPlayerTargetNpc = -1
        End If
    End If
End Function

Function GetNpcY(ByVal MapNum As Long, ByVal MapNpcNum As Long)
    If MapNpcNum > 0 Then
        GetNpcY = MapNPC(MapNum, MapNpcNum).Y
    End If
End Function

Sub SetWeather(ByVal MapNum As Long, ByVal Weather As Long, ByVal Interval As Long)
    Call SendDataToMap(MapNum, "mapweather" & SEP_CHAR & MapNum & SEP_CHAR & Weather & SEP_CHAR & Interval & END_CHAR)
End Sub

Function GetWeather(ByVal MapNum As Long)
    GetWeather = Map(MapNum).Weather
End Function

Sub LockSpells(ByVal Index As Long, ByVal Locked As Long)
    If Locked = 0 Then
        Player(Index).LockedSpells = False
        Call SendDataTo(Index, "CHECKFORMAP" & SEP_CHAR & GetPlayerMap(Index) & SEP_CHAR & Map(GetPlayerMap(Index)).Revision & END_CHAR)
    End If

    If Locked = 1 Then
        Player(Index).LockedSpells = True
        Call SendDataTo(Index, "CHECKFORMAP" & SEP_CHAR & GetPlayerMap(Index) & SEP_CHAR & Map(GetPlayerMap(Index)).Revision & END_CHAR)
    End If
End Sub

Sub LockItems(ByVal Index As Long, ByVal Locked As Long)
    If Locked = 0 Then
        Player(Index).LockedItems = False
        Call SendDataTo(Index, "CHECKFORMAP" & SEP_CHAR & GetPlayerMap(Index) & SEP_CHAR & Map(GetPlayerMap(Index)).Revision & END_CHAR)
    End If

    If Locked = 1 Then
        Player(Index).LockedItems = True
        Call SendDataTo(Index, "CHECKFORMAP" & SEP_CHAR & GetPlayerMap(Index) & SEP_CHAR & Map(GetPlayerMap(Index)).Revision & END_CHAR)
    End If
End Sub

Sub LockAttack(ByVal Index As Long, ByVal Locked As Long)
    If Locked = 0 Then
        Player(Index).LockedAttack = False
        Call SendDataTo(Index, "CHECKFORMAP" & SEP_CHAR & GetPlayerMap(Index) & SEP_CHAR & Map(GetPlayerMap(Index)).Revision & END_CHAR)
    End If

    If Locked = 1 Then
        Player(Index).LockedAttack = True
        Call SendDataTo(Index, "CHECKFORMAP" & SEP_CHAR & GetPlayerMap(Index) & SEP_CHAR & Map(GetPlayerMap(Index)).Revision & END_CHAR)
    End If
End Sub

Function GetIndexPlayerOnMap(ByVal Map As Long, ByVal X As Long, ByVal Y As Long)
    Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) And GetPlayerMap(i) = Map Then
            If GetPlayerX(i) = X And GetPlayerY(i) = Y Then
                GetIndexPlayerOnMap = i
                Exit Function
            End If
        End If
    Next i
End Function

Function GetPlayerHead(ByVal Index As Long)
    GetPlayerHead = 0
End Function

Function GetPlayerBody(ByVal Index As Long)
    GetPlayerBody = 0
End Function

Function GetPlayerleg(ByVal Index As Long)
    GetPlayerleg = 0
End Function

Public Sub MovePlayer(ByVal Index As Long, ByVal Direction As Long, ByVal Movement As Long)
    Dim X As Integer
    Dim Y As Integer

    Y = GetPlayerY(Index)
    X = GetPlayerX(Index)

    Select Case Direction
        Case DIR_UP
            Y = Y - 1
        Case DIR_DOWN
           Y = Y + 1
        Case DIR_LEFT
            X = X - 1
        Case DIR_RIGHT
            X = X + 1
    End Select

    Call PlayerMove(Index, Direction, Movement, X, Y)
    Call SendPlayerNewXY(Index)
End Sub

Public Sub SavePlayer(ByVal Index As Long)
    Call SavePlayer(Index)
End Sub

Function GetPlayerGender(ByVal Index As Long) As Long
    GetPlayerGender = Player(Index).Char(Player(Index).CharNum).Sex
End Function

Public Sub ChangeMapName(ByVal Index As Long, ByVal Name As String)
    Map(Index).Name = Name
    Call SaveMap(Index)
    Call SendMap(Index)
End Sub

' The subs below are added by DFA
Public Sub PlayerMapDropItem(ByVal Index As Long, ByVal InvNum As Long, ByVal Amount As Long)
    Call modGameLogic.PlayerMapDropItem(Index, InvNum, Amount)
End Sub

Public Function ScriptGetTickCount() As Long
    ScriptGetTickCount = GetTickCount
End Function


' Eclipse 2.7 Scripting Commands

Public Function GetMaxMapX() As Byte
    GetMaxMapX = MAX_MAPX
End Function

Public Function GetMaxMapY() As Byte
    GetMaxMapY = MAX_MAPY
End Function

Public Sub SetMapMoral(ByVal MapNum As Long, ByVal Moral As Long)
    Map(MapNum).Moral = Moral
End Sub

Public Function GetMapMoral(ByVal MapNum As Long) As Byte
    GetMapMoral = Map(MapNum).Moral
End Function

Public Sub ShopReload(ByVal ShopNum As Long)
    Dim FileID As Integer
    Dim FileName As String
    
    FileName = App.Path & "\Shops\Shop" & ShopNum & ".dat"

    FileID = FreeFile

    Open FileName For Binary As #FileID
        Get #FileID, , Shop(ShopNum)
    Close #FileID
End Sub

Sub DamageNPC(ByVal Index As Long, ByVal NPCNum As Long, ByVal Damage As Long)
    Call AttackNpc(Index, NPCNum, Damage)
    Call SendDataTo(Index, "BLITPLAYERDMG" & SEP_CHAR & Damage & SEP_CHAR & NPCNum & END_CHAR)
End Sub

Sub DamagePlayer(ByVal Index As Long, ByVal PIndex As Long, ByVal Damage As Long)
    Call AttackPlayer(Index, PIndex, Damage)
End Sub

Sub NPCAttack(ByVal X As Long, ByVal Target As Long, ByVal Damage As Long)
    Call NpcAttackPlayer(X, Target, Damage)
End Sub

Function GetPlayerInvItemAmmo(ByVal Index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemAmmo = Player(Index).Char(Player(Index).CharNum).Inv(InvSlot).Ammo
End Function

Sub SetPlayerInvItemAmmo(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemAmmo As Long)
    Player(Index).Char(Player(Index).CharNum).Inv(InvSlot).Ammo = ItemAmmo
End Sub

Function GetPlayerBankItemAmmo(ByVal Index As Long, ByVal BankSlot As Long) As Long
    GetPlayerBankItemAmmo = Player(Index).Char(Player(Index).CharNum).Bank(BankSlot).Ammo
End Function

Sub SetPlayerBankItemAmmo(ByVal Index As Long, ByVal BankSlot As Long, ByVal ItemAmmo As Long)
    Player(Index).Char(Player(Index).CharNum).Bank(BankSlot).Ammo = ItemAmmo
End Sub

Sub RemovePartyMember(ByVal Index As Long)
    Call RemovePMember(Index)
End Sub

Sub SetTimer(ByVal Index As Long, ByVal Number As Long, ByVal Interval As Long, Optional ByVal Parameter1 As Long = 0, Optional ByVal Parameter2 As Long = 0, Optional ByVal Parameter3 As Long = 0)
    Call AddNewTimer(Index, Number, Interval, Parameter1, Parameter2, Parameter3)
End Sub

Function GetPlayerEquipSlotNum(ByVal Index As Long, ByVal EquipSlot As Long) As Long
    GetPlayerEquipSlotNum = Player(Index).Char(Player(Index).CharNum).Equipment(EquipSlot).num
End Function

Function GetPlayerEquipSlotValue(ByVal Index As Long, ByVal EquipSlot As Long) As Long
    GetPlayerEquipSlotValue = Player(Index).Char(Player(Index).CharNum).Equipment(EquipSlot).Value
End Function

Function GetPlayerEquipSlotAmmo(ByVal Index As Long, ByVal EquipSlot As Long) As Long
    GetPlayerEquipSlotAmmo = Player(Index).Char(Player(Index).CharNum).Equipment(EquipSlot).Ammo
End Function

Sub SetPlayerEquipSlotNum(ByVal Index As Long, ByVal EquipSlot As Long, ByVal ItemNum As Long)
    Player(Index).Char(Player(Index).CharNum).Equipment(EquipSlot).num = ItemNum
End Sub

Sub SetPlayerEquipSlotValue(ByVal Index As Long, ByVal EquipSlot As Long, ByVal Value As Long)
    Player(Index).Char(Player(Index).CharNum).Equipment(EquipSlot).Value = Value
End Sub

Sub SetPlayerEquipSlotAmmo(ByVal Index As Long, ByVal EquipSlot As Long, ByVal Ammo As Long)
    Player(Index).Char(Player(Index).CharNum).Equipment(EquipSlot).Ammo = Ammo
End Sub

Sub SendEquipmentUpdate(ByVal Index As Long, ByVal EquipSlot As Long)
    Call SendDataToMap(GetPlayerMap(Index), "PLAYEREQUIPUPDATE" & SEP_CHAR & EquipSlot & SEP_CHAR & Index & SEP_CHAR & GetPlayerEquipSlotNum(Index, EquipSlot) & SEP_CHAR & GetPlayerEquipSlotValue(Index, EquipSlot) & SEP_CHAR & GetPlayerEquipSlotAmmo(Index, EquipSlot) & END_CHAR)
End Sub
