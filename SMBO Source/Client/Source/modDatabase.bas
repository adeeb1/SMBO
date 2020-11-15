Attribute VB_Name = "modDatabase"
Option Explicit

Sub SaveLocalMap(ByVal MapNum As Long)
    Dim FileName As String
    Dim F As Long

    FileName = App.Path & "\maps\map" & MapNum & ".dat"
    
    F = FreeFile
    Open FileName For Binary As #F
    Put #F, , Map(MapNum)
    Close #F
End Sub

Sub LoadMap(ByVal MapNum As Long)
    Dim FileName As String
    Dim F As Long

    FileName = App.Path & "\maps\map" & MapNum & ".dat"

    If FileExists("maps\map" & MapNum & ".dat") = False Then
        Exit Sub
    End If
    F = FreeFile
    Open FileName For Binary As #F
    Get #F, , Map(MapNum)
    Close #F
End Sub

Function GetMapRevision(ByVal MapNum As Long) As Long
    GetMapRevision = Map(MapNum).Revision
End Function

Sub ClearTempTile()
    Dim X As Long, Y As Long

    For Y = 0 To MAX_MAPY
        For X = 0 To MAX_MAPX
            TempTile(X, Y).DoorOpen = No
        Next X
    Next Y
End Sub

Sub ClearPlayer(ByVal Index As Long)
    Dim i As Long, n As Long
    
    Player(Index).Name = vbNullString
    Player(Index).Guild = vbNullString
    Player(Index).Guildaccess = 0
    Player(Index).Class = 0
    Player(Index).Level = 0
    Player(Index).Sprite = 0
    Player(Index).Exp = 0
    Player(Index).Access = 0
    Player(Index).PK = No

    Player(Index).HP = 0
    Player(Index).MP = 0
    Player(Index).SP = 0

    Player(Index).STR = 0
    Player(Index).DEF = 0
    Player(Index).speed = 0
    Player(Index).MAGI = 0
    
    For n = 1 To MAX_INV
        Player(Index).Inv(n).num = 0
        Player(Index).Inv(n).Value = 0
        Player(Index).Inv(n).Ammo = -1
    Next n

    For n = 1 To MAX_BANK
        Player(Index).Bank(n).num = 0
        Player(Index).Bank(n).Value = 0
        Player(Index).Bank(n).Ammo = -1
    Next n
    
    For n = 1 To 7
        Call SetPlayerEquipSlotNum(Index, n, 0)
        Call SetPlayerEquipSlotValue(Index, n, 0)
        Call SetPlayerEquipSlotAmmo(Index, n, -1)
    Next n
    
    If Player(Index).MaxInv < 1 Then
        Player(Index).MaxInv = 24
    End If
    
    ReDim Player(Index).NewInv(1 To Player(Index).MaxInv) As PlayerInvRec
    
    For n = 1 To Player(Index).MaxInv
        Player(Index).NewInv(n).num = 0
        Player(Index).NewInv(n).Value = 0
        Player(Index).NewInv(n).Ammo = -1
    Next n
    
    Player(Index).Map = 0
    Player(Index).X = 0
    Player(Index).Y = 0
    Player(Index).Dir = 0

    ' Client use only
    Player(Index).MaxHp = 0
    Player(Index).MaxMP = 0
    Player(Index).MaxSP = 0
    Player(Index).NextLvlExp = 0
    Player(Index).xOffset = 0
    Player(Index).yOffset = 0
    Player(Index).MovingH = 0
    Player(Index).MovingV = 0
    Player(Index).Moving = 0
    Player(Index).Attacking = 0
    Player(Index).AttackTimer = 0
    Player(Index).MapGetTimer = 0
    Player(Index).CastedSpell = No
    Player(Index).EmoticonNum = -1
    Player(Index).EmoticonTime = 0
    Player(Index).EmoticonVar = 0
    Player(Index).JumpDir = 0
    Player(Index).TempJumpAnim = 0
    Player(Index).JumpAnim = 0
    Player(Index).Jumping = False
    Player(Index).BattleVictory = False

    For i = 1 To MAX_SPELL_ANIM
        Player(Index).SpellAnim(i).CastedSpell = No
        Player(Index).SpellAnim(i).SpellTime = 0
        Player(Index).SpellAnim(i).SpellVar = 0
        Player(Index).SpellAnim(i).SpellDone = 0

        Player(Index).SpellAnim(i).Target = 0
        Player(Index).SpellAnim(i).TargetType = 0
    Next i

    Player(Index).SpellNum = 0

    For i = 1 To MAX_BLT_LINE
        BattlePMsg(i).Index = 1
        BattlePMsg(i).Time = i
        BattleMMsg(i).Index = 1
        BattleMMsg(i).Time = i
    Next i

    Inventory = 1
End Sub

Sub ClearItem(ByVal Index As Long)
    Item(Index).Name = vbNullString
    Item(Index).desc = vbNullString

    Item(Index).Type = 0
    Item(Index).Data1 = 0
    Item(Index).Data2 = 0
    Item(Index).Data3 = 0
    Item(Index).StrReq = 0
    Item(Index).DefReq = 0
    Item(Index).SpeedReq = 0
    Item(Index).MagicReq = 0
    Item(Index).ClassReq = -1
    Item(Index).AccessReq = 0
    Item(Index).LevelReq = 0

    Item(Index).AddHP = 0
    Item(Index).AddMP = 0
    Item(Index).AddSP = 0
    Item(Index).AddSTR = 0
    Item(Index).AddDef = 0
    Item(Index).AddMAGI = 0
    Item(Index).AddSpeed = 0
    Item(Index).AddEXP = 0
    Item(Index).AttackSpeed = 1000
    Item(Index).Stackable = 0
End Sub

Sub ClearItems()
    Dim i As Long

    For i = 1 To MAX_ITEMS
        Call ClearItem(i)
    Next i
End Sub

Sub ClearMapItem(ByVal Index As Long)
    MapItem(Index).num = 0
    MapItem(Index).Value = 0
    MapItem(Index).X = 0
    MapItem(Index).Y = 0
    MapItem(Index).Ammo = -1
End Sub

Sub ClearMap()
    Dim i As Long
    Dim X As Long
    Dim Y As Long

    For i = 1 To MAX_MAPS
        Map(i).Name = vbNullString
        Map(i).Revision = 0
        Map(i).Moral = 0
        Map(i).Up = 0
        Map(i).Down = 0
        Map(i).Left = 0
        Map(i).Right = 0
        Map(i).Indoors = 0

        For Y = 0 To MAX_MAPY
            For X = 0 To MAX_MAPX
                Map(i).Tile(X, Y).Ground = 0
                Map(i).Tile(X, Y).Mask = 0
                Map(i).Tile(X, Y).Anim = 0
                Map(i).Tile(X, Y).Mask2 = 0
                Map(i).Tile(X, Y).M2Anim = 0
                Map(i).Tile(X, Y).Fringe = 0
                Map(i).Tile(X, Y).FAnim = 0
                Map(i).Tile(X, Y).Fringe2 = 0
                Map(i).Tile(X, Y).F2Anim = 0
                Map(i).Tile(X, Y).Type = 0
                Map(i).Tile(X, Y).Data1 = 0
                Map(i).Tile(X, Y).Data2 = 0
                Map(i).Tile(X, Y).Data3 = 0
                Map(i).Tile(X, Y).String1 = vbNullString
                Map(i).Tile(X, Y).String2 = vbNullString
                Map(i).Tile(X, Y).String3 = vbNullString
                Map(i).Tile(X, Y).light = 0
                Map(i).Tile(X, Y).GroundSet = 0
                Map(i).Tile(X, Y).MaskSet = 0
                Map(i).Tile(X, Y).AnimSet = 0
                Map(i).Tile(X, Y).Mask2Set = 0
                Map(i).Tile(X, Y).M2AnimSet = 0
                Map(i).Tile(X, Y).FringeSet = 0
                Map(i).Tile(X, Y).FAnimSet = 0
                Map(i).Tile(X, Y).Fringe2Set = 0
                Map(i).Tile(X, Y).F2AnimSet = 0
            Next X
        Next Y
    Next i
End Sub

Sub ClearMapItems()
    Dim X As Long

    For X = 1 To MAX_MAP_ITEMS
        Call ClearMapItem(X)
    Next X
End Sub

Sub ClearMapNpc(ByVal Index As Long)
    MapNpc(Index).num = 0
    MapNpc(Index).Target = 0
    MapNpc(Index).HP = 0
    MapNpc(Index).MP = 0
    MapNpc(Index).SP = 0
    MapNpc(Index).Map = 0
    MapNpc(Index).X = 0
    MapNpc(Index).Y = 0
    MapNpc(Index).Dir = 0

    ' Client use only
    MapNpc(Index).xOffset = 0
    MapNpc(Index).yOffset = 0
    MapNpc(Index).Moving = 0
    MapNpc(Index).Attacking = 0
    MapNpc(Index).AttackTimer = 0
End Sub

Sub ClearMapNpcs()
    Dim i As Long

    For i = 1 To MAX_MAP_NPCS
        Call ClearMapNpc(i)
    Next i
End Sub

Function GetPlayerName(ByVal Index As Long) As String
    If Index < 1 Or Index > MAX_PLAYERS Then
        Exit Function
    End If
    GetPlayerName = Trim$(Player(Index).Name)
End Function

Sub SetPlayerName(ByVal Index As Long, ByVal Name As String)
    Player(Index).Name = Name
End Sub

Function GetPlayerGuild(ByVal Index As Long) As String
    GetPlayerGuild = Trim$(Player(Index).Guild)
End Function

Sub SetPlayerGuild(ByVal Index As Long, ByVal Guild As String)
    Player(Index).Guild = Guild
End Sub

Function GetPlayerGuildAccess(ByVal Index As Long) As Long
    GetPlayerGuildAccess = Player(Index).Guildaccess
End Function

Sub SetPlayerGuildAccess(ByVal Index As Long, ByVal Guildaccess As Long)
    Player(Index).Guildaccess = Guildaccess
End Sub

Function GetPlayerClass(ByVal Index As Long) As Long
    GetPlayerClass = Player(Index).Class
End Function

Sub SetPlayerClass(ByVal Index As Long, ByVal ClassNum As Long)
    Player(Index).Class = ClassNum
End Sub

Function GetPlayerSprite(ByVal Index As Long) As Long
    GetPlayerSprite = Player(Index).Sprite
End Function

Sub SetPlayerSprite(ByVal Index As Long, ByVal Sprite As Long)
    Player(Index).Sprite = Sprite
End Sub

Function GetPlayerLevel(ByVal Index As Long) As Long
    GetPlayerLevel = Player(Index).Level
End Function

Sub SetPlayerLevel(ByVal Index As Long, ByVal Level As Long)
    Player(Index).Level = Level
End Sub

Function GetPlayerExp(ByVal Index As Long) As Long
    GetPlayerExp = Player(Index).Exp
End Function

Sub SetPlayerExp(ByVal Index As Long, ByVal Exp As Long)
    Player(Index).Exp = Exp
End Sub

Function GetPlayerNextLvlExp(ByVal Index As Long) As Long
    GetPlayerNextLvlExp = Player(Index).NextLvlExp
End Function

Sub SetPlayerNextLvlExp(ByVal Index As Long, ByVal Exp As Long)
    Player(Index).NextLvlExp = Exp
End Sub

Function GetPlayerAccess(ByVal Index As Long) As Long
    GetPlayerAccess = Player(Index).Access
End Function

Sub SetPlayerAccess(ByVal Index As Long, ByVal Access As Long)
    Player(Index).Access = Access
End Sub

Function GetPlayerPK(ByVal Index As Long) As Long
    GetPlayerPK = Player(Index).PK
End Function

Sub SetPlayerPK(ByVal Index As Long, ByVal PK As Long)
    Player(Index).PK = PK
End Sub

Function GetPlayerHP(ByVal Index As Long) As Long
    GetPlayerHP = Player(Index).HP
End Function

Sub SetPlayerHP(ByVal Index As Long, ByVal HP As Long)
    Player(Index).HP = HP

    If GetPlayerHP(Index) > GetPlayerMaxHP(Index) Then
        Player(Index).HP = GetPlayerMaxHP(Index)
    End If
End Sub

Function GetPlayerMP(ByVal Index As Long) As Long
    GetPlayerMP = Player(Index).MP
End Function

Sub SetPlayerMP(ByVal Index As Long, ByVal MP As Long)
    Player(Index).MP = MP

    If GetPlayerMP(Index) > GetPlayerMaxMP(Index) Then
        Player(Index).MP = GetPlayerMaxMP(Index)
    End If
End Sub

Function GetPlayerSP(ByVal Index As Long) As Long
    GetPlayerSP = Player(Index).SP
End Function

Sub SetPlayerSP(ByVal Index As Long, ByVal SP As Long)
    Player(Index).SP = SP

    If GetPlayerSP(Index) > GetPlayerMaxSP(Index) Then
        Player(Index).SP = GetPlayerMaxSP(Index)
    End If
End Sub

Function GetPlayerMaxHP(ByVal Index As Long) As Long
    GetPlayerMaxHP = Player(Index).MaxHp
End Function

Function GetPlayerMaxMP(ByVal Index As Long) As Long
    GetPlayerMaxMP = Player(Index).MaxMP
End Function

Function GetPlayerMaxSP(ByVal Index As Long) As Long
    GetPlayerMaxSP = Player(Index).MaxSP
End Function

Function GetPlayerSTR(ByVal Index As Long) As Long
    GetPlayerSTR = Player(Index).STR
End Function

Sub SetPlayerSTR(ByVal Index As Long, ByVal STR As Long)
    Player(Index).STR = STR
End Sub

Function GetPlayerDEF(ByVal Index As Long) As Long
    GetPlayerDEF = Player(Index).DEF
End Function

Sub SetPlayerDEF(ByVal Index As Long, ByVal DEF As Long)
    Player(Index).DEF = DEF
End Sub

Function GetPlayerSPEED(ByVal Index As Long) As Long
    GetPlayerSPEED = Player(Index).speed
End Function

Sub SetPlayerSPEED(ByVal Index As Long, ByVal speed As Long)
    Player(Index).speed = speed
End Sub

Function GetPlayerStache(ByVal Index As Long) As Long
    GetPlayerStache = Player(Index).MAGI
End Function

Sub SetPlayerStache(ByVal Index As Long, ByVal Stache As Long)
    Player(Index).MAGI = Stache
End Sub

Function GetPlayerCritHitChance(ByVal Index As Long) As Long
    GetPlayerCritHitChance = Player(Index).CritHitChance
End Function

Sub SetPlayerCritHitChance(ByVal Index As Long, ByVal CritHitChance As Long)
    Player(Index).CritHitChance = CritHitChance
End Sub

Function GetPlayerBlockChance(ByVal Index As Long) As Long
    GetPlayerBlockChance = Player(Index).BlockChance
End Function

Sub SetPlayerBlockChance(ByVal Index As Long, ByVal BlockChance As Long)
    Player(Index).BlockChance = BlockChance
End Sub

Function GetPlayerPOINTS(ByVal Index As Long) As Long
    GetPlayerPOINTS = Player(Index).POINTS
End Function

Sub SetPlayerPOINTS(ByVal Index As Long, ByVal POINTS As Long)
    Player(Index).POINTS = POINTS
End Sub

Function GetPlayerMap(ByVal Index As Long) As Long
    If Index <= 0 Then
        Exit Function
    End If
    GetPlayerMap = Player(Index).Map
End Function

Sub SetPlayerMap(ByVal Index As Long, ByVal MapNum As Long)
    Player(Index).Map = MapNum
End Sub

Function GetPlayerX(ByVal Index As Long) As Long
    GetPlayerX = Player(Index).X
End Function

Sub SetPlayerX(ByVal Index As Long, ByVal X As Long)
    Player(Index).X = X
End Sub

Function GetPlayerY(ByVal Index As Long) As Long
    GetPlayerY = Player(Index).Y
End Function

Sub SetPlayerY(ByVal Index As Long, ByVal Y As Long)
    Player(Index).Y = Y
End Sub
Sub SetPlayerLoc(ByVal Index As Long, ByVal X As Long, ByVal Y As Long)
    Player(Index).X = X
    Player(Index).Y = Y
End Sub

Function GetPlayerDir(ByVal Index As Long) As Long
    GetPlayerDir = Player(Index).Dir
End Function

Sub SetPlayerDir(ByVal Index As Long, ByVal Dir As Long)
    Player(Index).Dir = Dir
End Sub

Function GetPlayerInvItemNum(ByVal Index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemNum = Player(Index).NewInv(InvSlot).num
End Function

Sub SetPlayerInvItemNum(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemNum As Long)
    Player(Index).NewInv(InvSlot).num = ItemNum
End Sub

Function GetPlayerInvItemValue(ByVal Index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemValue = Player(Index).NewInv(InvSlot).Value
End Function

Sub SetPlayerInvItemValue(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemValue As Long)
    Player(Index).NewInv(InvSlot).Value = ItemValue
End Sub

Function GetPlayerBankItemNum(ByVal Index As Long, ByVal BankSlot As Long) As Long
    If BankSlot > MAX_BANK Then
        Exit Function
    End If
    GetPlayerBankItemNum = Player(Index).Bank(BankSlot).num
End Function

Sub SetPlayerBankItemNum(ByVal Index As Long, ByVal BankSlot As Long, ByVal ItemNum As Long)
    Player(Index).Bank(BankSlot).num = ItemNum
End Sub

Function GetPlayerBankItemValue(ByVal Index As Long, ByVal BankSlot As Long) As Long
    GetPlayerBankItemValue = Player(Index).Bank(BankSlot).Value
End Function

Sub SetPlayerBankItemValue(ByVal Index As Long, ByVal BankSlot As Long, ByVal ItemValue As Long)
    Player(Index).Bank(BankSlot).Value = ItemValue
End Sub

Sub SetPlayerHead(ByVal Index As Long, ByVal head As Long)
    If Index > 0 And Index < MAX_PLAYERS Then
        Player(Index).head = head
    End If
End Sub

Sub SetPlayerBody(ByVal Index As Long, ByVal body As Long)
    If Index > 0 And Index < MAX_PLAYERS Then
        Player(Index).body = body
    End If
End Sub

Sub SetPlayerLeg(ByVal Index As Long, ByVal leg As Long)
    If Index > 0 And Index < MAX_PLAYERS Then
        Player(Index).leg = leg
    End If
End Sub

Function GetPlayerInvItemAmmo(ByVal Index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemAmmo = Player(Index).NewInv(InvSlot).Ammo
End Function

Sub SetPlayerInvItemAmmo(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemAmmo As Long)
    Player(Index).NewInv(InvSlot).Ammo = ItemAmmo
End Sub

Function GetPlayerBankItemAmmo(ByVal Index As Long, ByVal BankSlot As Long) As Long
    GetPlayerBankItemAmmo = Player(Index).Bank(BankSlot).Ammo
End Function

Sub SetPlayerBankItemAmmo(ByVal Index As Long, ByVal BankSlot As Long, ByVal ItemAmmo As Long)
    Player(Index).Bank(BankSlot).Ammo = ItemAmmo
End Sub

Function GetPlayerEquipSlotNum(ByVal Index As Long, ByVal EquipSlot As Long) As Long
    GetPlayerEquipSlotNum = Player(Index).Equipment(EquipSlot).num
End Function

Function GetPlayerEquipSlotValue(ByVal Index As Long, ByVal EquipSlot As Long) As Long
    GetPlayerEquipSlotValue = Player(Index).Equipment(EquipSlot).Value
End Function

Function GetPlayerEquipSlotAmmo(ByVal Index As Long, ByVal EquipSlot As Long) As Long
    GetPlayerEquipSlotAmmo = Player(Index).Equipment(EquipSlot).Ammo
End Function

Sub SetPlayerEquipSlotNum(ByVal Index As Long, ByVal EquipSlot As Long, ByVal ItemNum As Long)
    Player(Index).Equipment(EquipSlot).num = ItemNum
End Sub

Sub SetPlayerEquipSlotValue(ByVal Index As Long, ByVal EquipSlot As Long, ByVal Value As Long)
    Player(Index).Equipment(EquipSlot).Value = Value
End Sub

Sub SetPlayerEquipSlotAmmo(ByVal Index As Long, ByVal EquipSlot As Long, ByVal Ammo As Long)
    Player(Index).Equipment(EquipSlot).Ammo = Ammo
End Sub

Sub SetPlayerHeight(ByVal Height As Integer)
    Player(MyIndex).Height = Height
    
    If Player(MyIndex).Height < 0 Then
        Player(MyIndex).Height = 0
    End If
End Sub

Function GetPlayerAttackSpeed(ByVal Index As Long) As Long
    Dim i As Integer
    
    GetPlayerAttackSpeed = 1000
    
    For i = 1 To 7
        If GetPlayerEquipSlotNum(Index, i) > 0 Then
            GetPlayerAttackSpeed = GetPlayerAttackSpeed - (1000 - Item(GetPlayerEquipSlotNum(Index, i)).AttackSpeed)
        End If
    Next i
    
    GetPlayerAttackSpeed = GetPlayerAttackSpeed - (GetPlayerSPEED(Index) * 3)
End Function
