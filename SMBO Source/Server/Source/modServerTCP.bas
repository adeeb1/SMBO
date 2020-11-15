Attribute VB_Name = "modServerTCP"
Option Explicit

 Sub SendPlayerNewXY(ByVal Index As Long)
    Call SendDataTo(Index, SPackets.Splayernewxy & SEP_CHAR & GetPlayerX(Index) & SEP_CHAR & GetPlayerY(Index) & END_CHAR)
 End Sub

Sub UpdateTitle()
    frmServer.Caption = "Super Mario Bros. Online (" & frmServer.Socket(0).LocalIP & ":" & CStr(frmServer.Socket(0).LocalPort) & ") - Eclipse Evolution Server"
End Sub

Sub UpdateTOP()
    frmServer.TPO.Caption = "Total Players Online: " & TotalOnlinePlayers
End Sub

Sub CreateFullMapCache()
    Dim i As Long

    For i = 1 To MAX_MAPS
        Call MapCache_Create(i)
    Next i
End Sub

Function IsConnected(ByVal Index As Long) As Boolean
    On Error Resume Next
    
    If frmServer.Socket(Index).State = sckConnected Then
        IsConnected = True
    End If
End Function

Function IsPlaying(ByVal Index As Long) As Boolean
    If Index < 1 Or Index > MAX_PLAYERS Then
        Exit Function
    End If

    If IsConnected(Index) Then
        If Player(Index).InGame Then
            IsPlaying = True
        End If
    End If
End Function

Function IsLoggedIn(ByVal Index As Long) As Boolean
    If Index < 1 Or Index > MAX_PLAYERS Then
        Exit Function
    End If

    If IsConnected(Index) Then
        If Trim$(Player(Index).Login) <> vbNullString Then
            IsLoggedIn = True
        End If
    End If
End Function

Function IsMultiAccounts(ByVal Login As String) As Boolean
    Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsConnected(i) Then
            If LCase$(Trim$(Player(i).Login)) = LCase$(Trim$(Login)) Then
                IsMultiAccounts = True
                Exit Function
            End If
        End If
    Next i
End Function

Function IsMuted(ByVal Index As Long) As Boolean
    Dim FileName As String
    Dim i As Integer

    FileName = App.Path & "\SMBOMuteList.ini"
        
    For i = 1 To 100
        If GetVar(FileName, "Mute List", CStr(i)) = Trim$(Player(Index).Login) & "/" & GetPlayerName(Index) Then
            IsMuted = True
            Exit Function
        End If
    Next
    
    IsMuted = False
End Function

Function IsBanned(ByVal IPAddr As String) As Boolean
    Dim FileName As String, BanEntry As String
    Dim BanInfo() As String
    Dim i As Integer

    FileName = App.Path & "\SMBOBanList.ini"

    For i = 1 To 100
        BanEntry = GetVar(FileName, "Ban List", CStr(i))
        
        If BanEntry <> "" Then
            BanInfo = Split(BanEntry, " ")
            
            If BanInfo(0) = IPAddr Then
                IsBanned = True
                Exit Function
            End If
        End If
    Next
End Function

Sub SendDataTo(ByVal Index As Long, ByVal Data As String)
    Dim dbytes() As Byte

    dbytes = StrConv(Data, vbFromUnicode)
    If IsConnected(Index) Then
        frmServer.Socket(Index).SendData dbytes
        DoEvents
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

Sub GlobalMsg(ByVal Msg As String, ByVal Color As Byte)
    Call SendDataToAll(SPackets.Sglobalmsg & SEP_CHAR & Msg & SEP_CHAR & Color & END_CHAR)
End Sub

Sub PlayerMsg(ByVal Index As Long, ByVal Msg As String, ByVal Color As Byte)
    Call SendDataTo(Index, SPackets.Splayermsg & SEP_CHAR & Msg & SEP_CHAR & Color & END_CHAR)
End Sub

Sub OtherMsg(ByVal Index As Long, ByVal Msg As String, ByVal Color As Byte)
    Call SendDataTo(Index, SPackets.Sothermsg & SEP_CHAR & Msg & SEP_CHAR & Color & END_CHAR)
End Sub

Sub AdminMsg(ByVal Msg As String, ByVal Color As Byte)
    Dim packet As String
    Dim i As Long

    packet = SPackets.Sadminmsg & SEP_CHAR & Msg & SEP_CHAR & Color & END_CHAR

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            If GetPlayerAccess(i) > 1 Then
                Call SendDataTo(i, packet)
            End If
        End If
    Next i
End Sub

Sub MapMsg(ByVal MapNum As Long, ByVal Msg As String, ByVal Color As Byte)
    Call SendDataToMap(MapNum, SPackets.Smapmsg & SEP_CHAR & Msg & SEP_CHAR & Color & END_CHAR)
End Sub

Sub AlertMsg(ByVal Index As Long, ByVal Msg As String)
    Call SendDataTo(Index, SPackets.Salertmsg & SEP_CHAR & Msg & END_CHAR)
    Call CloseSocket(Index)
End Sub

Sub PlainMsg(ByVal Index As Long, ByVal Msg As String, ByVal num As Long)
    Call SendDataTo(Index, SPackets.Splainmsg & SEP_CHAR & Msg & SEP_CHAR & num & END_CHAR)
End Sub

Sub HackingAttempt(ByVal Index As Long, ByVal Reason As String)
    If Index < 1 Or Index > MAX_PLAYERS Then
        Exit Sub
    End If

    If IsPlaying(Index) Then
        Call AdminMsg(GetPlayerLogin(Index) & "/" & GetPlayerName(Index) & " has been booted for (" & Reason & ")", BRIGHTRED)
        Call AlertMsg(Index, "You have lost your connection with Super Mario Bros. Online.")
    End If
End Sub

Sub AcceptConnection(ByVal Index As Long, ByVal SocketId As Long)
    Dim i As Long
    
    If (Index = 0) Then
        i = FindOpenPlayerSlot

        If i <> 0 Then
            ' we can connect them
            frmServer.Socket(i).Close
            frmServer.Socket(i).Accept SocketId
            Call SocketConnected(i)
        End If
    End If
End Sub

Sub SocketConnected(ByVal Index As Long)
    If Index < 1 Or Index > MAX_PLAYERS Then
        Exit Sub
    End If
    
    If Not IsBanned(GetPlayerIP(Index)) Then
        Call TextAdd(frmServer.txtText(0), "Received connection from " & GetPlayerIP(Index) & ".", True)
    Else
        Call AlertMsg(Index, "You have been banned from Super Mario Bros. Online, and can no longer play.")
    End If
End Sub

Sub IncomingData(ByVal Index As Long, ByVal DataLength As Long)
    On Error Resume Next
    Dim Buffer As String, packet As String
    Dim Start As Long

    If Index > 0 Then
        frmServer.Socket(Index).GetData Buffer, vbString, DataLength
            
        Player(Index).Buffer = Player(Index).Buffer & Buffer
        
        Start = InStr(Player(Index).Buffer, END_CHAR)
        Do While Start > 0
            packet = Mid(Player(Index).Buffer, 1, Start - 1)
            Player(Index).Buffer = Mid(Player(Index).Buffer, Start + 1, Len(Player(Index).Buffer))
            Player(Index).DataPackets = Player(Index).DataPackets + 1
            Start = InStr(Player(Index).Buffer, END_CHAR)
            If Len(packet) > 0 Then
                Call HandleData(Index, packet)
            End If
        Loop
                
        ' Check if elapsed time has passed
        Player(Index).DataBytes = Player(Index).DataBytes + DataLength
        If GetTickCount >= Player(Index).DataTimer + 1000 Then
            Player(Index).DataTimer = GetTickCount
            Player(Index).DataBytes = 0
            Player(Index).DataPackets = 0
            Exit Sub
        End If
        
        ' Check for data flooding
        If Player(Index).DataBytes > 1500 And GetPlayerAccess(Index) <= 0 Then
            Exit Sub
        End If
        
        ' Check for packet flooding
        If Player(Index).DataPackets > 100 And GetPlayerAccess(Index) <= 0 Then
            Exit Sub
        End If
    End If
End Sub

Sub CloseSocket(ByVal Index As Long)
    If Index > 0 Then
        Call LeftGame(Index)
        Call TextAdd(frmServer.txtText(0), "Connection from " & GetPlayerIP(Index) & " has been terminated.", True)
        frmServer.Socket(Index).Close
        Call UpdateTOP
    End If
End Sub

Sub SendOnlineList()
    Dim packet As String
    Dim i As Long, PlayerCount As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            packet = packet & SEP_CHAR & GetPlayerName(i)
            PlayerCount = PlayerCount + 1
        End If
    Next i

    Call SendDataToAll(SPackets.Sonlinelist & SEP_CHAR & PlayerCount & packet & END_CHAR)
End Sub

Sub SendChars(ByVal Index As Long)
    Dim packet As String
    Dim i As Long

    packet = SPackets.Sallchars & SEP_CHAR
    
    For i = 1 To MAX_CHARS
        packet = packet & Player(Index).Char(i).Name & SEP_CHAR & ClassData(Player(Index).Char(i).Class).Name & SEP_CHAR & Player(Index).Char(i).LEVEL & SEP_CHAR
    Next i
    
    packet = packet & END_CHAR

    Call SendDataTo(Index, packet)
End Sub

Sub SendJoinMap(ByVal Index As Long)
    Dim packet As String
    Dim i As Long

    packet = vbNullString

    ' Send all players on current map to index
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            If GetPlayerMap(i) = GetPlayerMap(Index) Then
                packet = SPackets.Splayerdata & SEP_CHAR
                packet = packet & i & SEP_CHAR
                packet = packet & GetPlayerName(i) & SEP_CHAR
                packet = packet & GetPlayerSprite(i) & SEP_CHAR
                packet = packet & GetPlayerMap(i) & SEP_CHAR
                packet = packet & GetPlayerDir(i) & SEP_CHAR
                packet = packet & GetPlayerAccess(i) & SEP_CHAR
                packet = packet & GetPlayerPK(i) & SEP_CHAR
                packet = packet & GetPlayerGuild(i) & SEP_CHAR
                packet = packet & GetPlayerGuildAccess(i) & SEP_CHAR
                packet = packet & GetPlayerClass(i) & SEP_CHAR
                packet = packet & GetPlayerLevel(i) & END_CHAR
                Call SendDataTo(Index, packet)
                
                Call SendPlayerXY(i)
            End If
        End If
    Next i

    Call SendPlayerData(Index)
    Call SendPlayerXY(Index)
End Sub

Sub SendLeaveMap(ByVal Index As Long, ByVal MapNum As Long)
    Dim packet As String

    packet = SPackets.Sleave & SEP_CHAR & Index & END_CHAR
    Call SendDataToMapBut(Index, MapNum, packet)
End Sub

Sub SendPlayerData(ByVal Index As Long)
    Dim packet As String
    
    ' Send index's player data to everyone including himself on the map
    packet = SPackets.Splayerdata & SEP_CHAR
    packet = packet & Index & SEP_CHAR
    packet = packet & GetPlayerName(Index) & SEP_CHAR
    packet = packet & GetPlayerSprite(Index) & SEP_CHAR
    packet = packet & GetPlayerMap(Index) & SEP_CHAR
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

Public Sub SendMap(ByVal Index As Long, ByVal MapNum As Long)
    Call SendDataTo(Index, MapCache(MapNum))
End Sub

Public Sub MapCache_Create(ByVal MapNum As Long)
    Dim MapData As String
    Dim X As Long, Y As Long

    MapData = SPackets.Smapdata & SEP_CHAR & MapNum & SEP_CHAR & Trim$(Map(MapNum).Name) & SEP_CHAR & Map(MapNum).Revision & SEP_CHAR & Map(MapNum).Moral & SEP_CHAR & Map(MapNum).Up & SEP_CHAR & Map(MapNum).Down & SEP_CHAR & Map(MapNum).Left & SEP_CHAR & Map(MapNum).Right & SEP_CHAR & Map(MapNum).music & SEP_CHAR & Map(MapNum).BootMap & SEP_CHAR & Map(MapNum).BootX & SEP_CHAR & Map(MapNum).BootY & SEP_CHAR & Map(MapNum).Indoors & SEP_CHAR & Map(MapNum).Weather & SEP_CHAR

    For Y = 0 To MAX_MAPY
        For X = 0 To MAX_MAPX
            With Map(MapNum).Tile(X, Y)
                MapData = MapData & (.Ground & SEP_CHAR & .Mask & SEP_CHAR & .Anim & SEP_CHAR & .Mask2 & SEP_CHAR & .M2Anim & SEP_CHAR & .Fringe & SEP_CHAR & .FAnim & SEP_CHAR & .Fringe2 & SEP_CHAR & .F2Anim & SEP_CHAR & .Type & SEP_CHAR & .Data1 & SEP_CHAR & .Data2 & SEP_CHAR & .Data3 & SEP_CHAR & Trim$(.String1) & SEP_CHAR & Trim$(.String2) & SEP_CHAR & Trim$(.String3) & SEP_CHAR & .Light & SEP_CHAR)
                MapData = MapData & (.GroundSet & SEP_CHAR & .MaskSet & SEP_CHAR & .AnimSet & SEP_CHAR & .Mask2Set & SEP_CHAR & .M2AnimSet & SEP_CHAR & .FringeSet & SEP_CHAR & .FAnimSet & SEP_CHAR & .Fringe2Set & SEP_CHAR & .F2AnimSet & SEP_CHAR)
            End With
        Next X
    Next Y

    For X = 1 To MAX_MAP_NPCS
        MapData = MapData & (Map(MapNum).NPC(X) & SEP_CHAR & Map(MapNum).SpawnX(X) & SEP_CHAR & Map(MapNum).SpawnY(X) & SEP_CHAR)
    Next X
    
    For Y = 0 To MAX_MAPY
        For X = 0 To MAX_MAPX
            With QuestionBlock(MapNum, X, Y)
                MapData = MapData & (.Item1 & SEP_CHAR & .Item2 & SEP_CHAR & .Item3 & SEP_CHAR & .Item4 & SEP_CHAR & .Item5 & SEP_CHAR & .Item6 & SEP_CHAR & .Chance1 & SEP_CHAR & .Chance2 & SEP_CHAR & .Chance3 & SEP_CHAR & .Chance4 & SEP_CHAR & .Chance5 & SEP_CHAR & .Chance6 & SEP_CHAR & .Value1 & SEP_CHAR & .Value2 & SEP_CHAR & .Value3 & SEP_CHAR & .Value4 & SEP_CHAR & .Value5 & SEP_CHAR & .Value6 & SEP_CHAR)
            End With
        Next X
    Next Y
    
    MapData = MapData & END_CHAR

    MapCache(MapNum) = MapData
End Sub

Sub SendMapItemsTo(ByVal Index As Long, ByVal MapNum As Long)
    Dim packet As String
    Dim i As Long

    packet = SPackets.Smapitemdata
    
    For i = 1 To MAX_MAP_ITEMS
        If MapNum > 0 Then
            packet = packet & SEP_CHAR & MapItem(MapNum, i).num & SEP_CHAR & MapItem(MapNum, i).Value & SEP_CHAR & MapItem(MapNum, i).X & SEP_CHAR & MapItem(MapNum, i).Y & SEP_CHAR & MapItem(MapNum, i).Ammo
        End If
    Next i
    
    packet = packet & END_CHAR

    Call SendDataTo(Index, packet)
End Sub

Sub SendMapItemsToMap(ByVal MapNum As Long)
    Dim packet As String
    Dim i As Long

    packet = SPackets.Smapitemdata
    
    For i = 1 To MAX_MAP_ITEMS
        packet = packet & SEP_CHAR & MapItem(MapNum, i).num & SEP_CHAR & MapItem(MapNum, i).Value & SEP_CHAR & MapItem(MapNum, i).X & SEP_CHAR & MapItem(MapNum, i).Y & SEP_CHAR & MapItem(MapNum, i).Ammo
    Next i
    
    packet = packet & END_CHAR

    Call SendDataToMap(MapNum, packet)
End Sub

Sub SendMapNpcsTo(ByVal Index As Long, ByVal MapNum As Long)
    Dim packet As String
    Dim i As Long

    packet = SPackets.Smapnpcdata
    
    For i = 1 To MAX_MAP_NPCS
        If MapNum > 0 Then
            packet = packet & SEP_CHAR & MapNPC(MapNum, i).num & SEP_CHAR & MapNPC(MapNum, i).X & SEP_CHAR & MapNPC(MapNum, i).Y & SEP_CHAR & MapNPC(MapNum, i).Dir & SEP_CHAR & MapNPC(MapNum, i).Target & SEP_CHAR & MapNPC(MapNum, i).InBattle
        End If
    Next i
    
    packet = packet & END_CHAR

    Call SendDataTo(Index, packet)
End Sub

Sub SendMapNpcsToMap(ByVal MapNum As Long)
    Dim packet As String
    Dim i As Long

    packet = SPackets.Smapnpcdata
    
    For i = 1 To MAX_MAP_NPCS
        packet = packet & SEP_CHAR & MapNPC(MapNum, i).num & SEP_CHAR & MapNPC(MapNum, i).X & SEP_CHAR & MapNPC(MapNum, i).Y & SEP_CHAR & MapNPC(MapNum, i).Dir & SEP_CHAR & MapNPC(MapNum, i).Target & SEP_CHAR & MapNPC(MapNum, i).InBattle
    Next i
    
    packet = packet & END_CHAR

    Call SendDataToMap(MapNum, packet)
End Sub

Sub SendItems(ByVal Index As Long)
    Dim i As Long

    For i = 1 To MAX_ITEMS
        If Trim$(Item(i).Name) <> vbNullString Then
            Call SendUpdateItemTo(Index, i)
        End If
    Next i
End Sub

Sub SendElements(ByVal Index As Long)
    Dim i As Long

    For i = 0 To MAX_ELEMENTS
        If Trim$(Element(i).Name) <> vbNullString Then
            Call SendUpdateElementTo(Index, i)
        End If
    Next i
End Sub
Sub SendEmoticons(ByVal Index As Long)
    Dim i As Long

    For i = 0 To MAX_EMOTICONS
        If Trim$(Emoticons(i).Command) <> vbNullString Then
            Call SendUpdateEmoticonTo(Index, i)
        End If
    Next i
End Sub

Sub SendArrows(ByVal Index As Long)
    Dim i As Long

    For i = 1 To MAX_ARROWS
        If Trim$(Arrows(i).Name) <> vbNullString Then
            Call SendUpdateArrowTo(Index, i)
        End If
    Next i
End Sub

Sub SendNpcs(ByVal Index As Long)
    Dim i As Long

    For i = 1 To MAX_NPCS
        If Trim$(NPC(i).Name) <> vbNullString Then
            Call SendUpdateNpcTo(Index, i)
        End If
    Next i
End Sub

Sub SendBank(ByVal Index As Long)
    Dim packet As String
    Dim i As Long

    packet = SPackets.Splayerbank
    
    For i = 1 To MAX_BANK
        packet = packet & SEP_CHAR & GetPlayerBankItemNum(Index, i) & SEP_CHAR & GetPlayerBankItemValue(Index, i) & SEP_CHAR & GetPlayerBankItemAmmo(Index, i)
    Next i
    
    packet = packet & END_CHAR

    Call SendDataTo(Index, packet)
End Sub

Sub SendBankUpdate(ByVal Index As Long, ByVal BankSlot As Long)
    Call SendDataTo(Index, SPackets.Splayerbankupdate & SEP_CHAR & BankSlot & SEP_CHAR & GetPlayerBankItemNum(Index, BankSlot) & SEP_CHAR & GetPlayerBankItemValue(Index, BankSlot) & SEP_CHAR & GetPlayerBankItemAmmo(Index, BankSlot) & END_CHAR)
End Sub

Sub SendInventory(ByVal Index As Long)
    Dim packet As String
    Dim i As Long

    packet = SPackets.Splayerinv & SEP_CHAR & Index & SEP_CHAR & GetPlayerMaxInv(Index)
    
    For i = 1 To GetPlayerMaxInv(Index)
        packet = packet & SEP_CHAR & GetPlayerInvItemNum(Index, i) & SEP_CHAR & GetPlayerInvItemValue(Index, i) & SEP_CHAR & GetPlayerInvItemAmmo(Index, i)
    Next i
    
    packet = packet & END_CHAR

    Call SendDataToMap(GetPlayerMap(Index), packet)
End Sub

Sub SendInventoryUpdate(ByVal Index As Long, ByVal InvSlot As Long)
    Call SendDataToMap(GetPlayerMap(Index), SPackets.Splayerinvupdate & SEP_CHAR & InvSlot & SEP_CHAR & Index & SEP_CHAR & GetPlayerInvItemNum(Index, InvSlot) & SEP_CHAR & GetPlayerInvItemValue(Index, InvSlot) & SEP_CHAR & GetPlayerInvItemAmmo(Index, InvSlot) & END_CHAR)
End Sub

Sub SendIndexInventoryFromMap(ByVal Index As Long)
    Dim packet As String
    Dim n As Long
    Dim i As Long
    
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            If GetPlayerMap(i) = GetPlayerMap(Index) Then
                
                packet = SPackets.Splayerinv & SEP_CHAR & i & SEP_CHAR & GetPlayerMaxInv(i)
                
                For n = 1 To GetPlayerMaxInv(i)
                    packet = packet & SEP_CHAR & GetPlayerInvItemNum(i, n) & SEP_CHAR & GetPlayerInvItemValue(i, n) & SEP_CHAR & GetPlayerInvItemAmmo(i, n)
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
        packet = SPackets.Splayerworneq & SEP_CHAR & Index
        
        For i = 1 To 7
            packet = packet & SEP_CHAR & GetPlayerEquipSlotNum(Index, i)
        Next i
        
        packet = packet & END_CHAR
        
        Call SendDataToMap(GetPlayerMap(Index), packet)
    End If
End Sub

Sub SendHP(ByVal Index As Long)
    Dim Exp As Long, i As Long
  
    Call SendDataTo(Index, SPackets.Splayerhp & SEP_CHAR & Index & SEP_CHAR & GetPlayerMaxHP(Index) & SEP_CHAR & GetPlayerHP(Index) & END_CHAR)
    
    If GetPlayerGuild(Index) <> vbNullString Then
        Call SendGuildMemberHP(Index)
    End If
    
    With Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index))
        If .Type = TILE_TYPE_LOWER_STAT Or .Type = TILE_TYPE_KILL Then
            ' Kill player and restore stats
            If GetPlayerHP(Index) < 1 Then
                ' Triggers Close Call
                If CheckCloseCall(Index) = True Then
                    Exit Sub
                End If
                
                ' Triggers Life Shroom
                If HasItem(Index, 43) = 1 Then
                    Call LifeShroom(Index)
                    Exit Sub
                End If
                
                Call PlayerDeath(Index)
            End If
        End If
    End With
    
    ' Poison Cave check
    If IsInPoisonCave(Index) Then
        ' Kill player and restore stats
        If GetPlayerHP(Index) < 1 Then
            ' Triggers Close Call
            If CheckCloseCall(Index) = True Then
                Exit Sub
            End If
                
            ' Triggers Life Shroom
            If HasItem(Index, 43) = 1 Then
                Call LifeShroom(Index)
                Exit Sub
            End If
            
            If GetPlayerInBattle(Index) = True Then
                Call TurnBasedDeath(Index, GetPlayerMap(Index), Player(Index).TargetNPC, 1, True)
            Else
                Call PlayerDeath(Index)
            End If
        End If
    End If
    
    ' Handle sounds
    If GetPlayerHP(Index) > 0 And GetPlayerHP(Index) <= 5 And Val(GetVar(App.Path & "\Sounds.ini", GetPlayerName(Index), "Sound")) = 0 Then
        Call PutVar(App.Path & "\Sounds.ini", GetPlayerName(Index), "Sound", "1")
        Call SendSoundTo(Index, "sms_lowhealth.wav")
    ElseIf GetPlayerHP(Index) > 5 And Val(GetVar(App.Path & "\Sounds.ini", GetPlayerName(Index), "Sound")) = 1 Then
        Call PutVar(App.Path & "\Sounds.ini", GetPlayerName(Index), "Sound", "0")
    End If
End Sub

Sub SendMP(ByVal Index As Long)
    Call SendDataTo(Index, SPackets.Splayermp & SEP_CHAR & GetPlayerMaxMP(Index) & SEP_CHAR & GetPlayerMP(Index) & END_CHAR)
End Sub

Sub SendSP(ByVal Index As Long)
    Call SendDataTo(Index, SPackets.Splayersp & SEP_CHAR & GetPlayerMaxSP(Index) & SEP_CHAR & GetPlayerSP(Index) & END_CHAR)
End Sub

Sub SendPTS(ByVal Index As Long)
    Call SendDataTo(Index, SPackets.Splayerpoints & SEP_CHAR & GetPlayerPOINTS(Index) & END_CHAR)
End Sub

Sub SendEXP(ByVal Index As Long)
    Call SendDataTo(Index, SPackets.Splayerexp & SEP_CHAR & GetPlayerExp(Index) & SEP_CHAR & GetPlayerNextLevel(Index) & END_CHAR)
End Sub

Sub SendStats(ByVal Index As Long)
    Call SendDataTo(Index, SPackets.Splayerstatspacket & SEP_CHAR & GetPlayerSTR(Index) & SEP_CHAR & GetPlayerDEF(Index) & SEP_CHAR & GetPlayerSPEED(Index) & SEP_CHAR & GetPlayerStache(Index) & SEP_CHAR & GetPlayerNextLevel(Index) & SEP_CHAR & GetPlayerExp(Index) & SEP_CHAR & GetPlayerLevel(Index) & SEP_CHAR & GetPlayerCritHitChance(Index) & SEP_CHAR & GetPlayerBlockChance(Index) & END_CHAR)
End Sub

Sub SendPlayerLevelToAll(ByVal Index As Long)
    Call SendDataToAll(SPackets.Splayerlevel & SEP_CHAR & Index & SEP_CHAR & GetPlayerLevel(Index) & END_CHAR)
End Sub

Sub SendClasses(ByVal Index As Long)
    Dim packet As String
    Dim i As Long

    packet = SPackets.Sclassesdata
    
    For i = 0 To MAX_CLASSES
        packet = packet & SEP_CHAR & GetClassName(i) & SEP_CHAR & GetClassMaxHP(i) & SEP_CHAR & GetClassMaxMP(i) & SEP_CHAR & GetClassMaxSP(i) & SEP_CHAR & ClassData(i).STR & SEP_CHAR & ClassData(i).DEF & SEP_CHAR & ClassData(i).Speed & SEP_CHAR & ClassData(i).Magi & SEP_CHAR & ClassData(i).Locked & SEP_CHAR & ClassData(i).Desc
    Next i
    
    packet = packet & END_CHAR

    Call SendDataTo(Index, packet)
End Sub

Sub SendNewCharClasses(ByVal Index As Long)
    Dim packet As String
    Dim i As Long

    packet = SPackets.Snewcharclasses
    
    For i = 0 To MAX_CLASSES
        packet = packet & SEP_CHAR & GetClassName(i) & SEP_CHAR & GetClassMaxHP(i) & SEP_CHAR & GetClassMaxMP(i) & SEP_CHAR & GetClassMaxSP(i) & SEP_CHAR & ClassData(i).STR & SEP_CHAR & ClassData(i).DEF & SEP_CHAR & ClassData(i).Speed & SEP_CHAR & ClassData(i).Magi & SEP_CHAR & ClassData(i).MaleSprite & SEP_CHAR & ClassData(i).FemaleSprite & SEP_CHAR & ClassData(i).Locked & SEP_CHAR & ClassData(i).Desc
    Next i
    
    packet = packet & END_CHAR

    Call SendDataTo(Index, packet)
End Sub

Sub SendLeftGame(ByVal Index As Long)
    Call SendDataToAllBut(Index, SPackets.Sleft & SEP_CHAR & Index & END_CHAR)
End Sub

Sub SendPlayerXY(ByVal Index As Long)
    Call SendDataToMap(GetPlayerMap(Index), SPackets.Splayerxy & SEP_CHAR & Index & SEP_CHAR & GetPlayerX(Index) & SEP_CHAR & GetPlayerY(Index) & END_CHAR)
End Sub

Sub SendUpdateItemToAll(ByVal ItemNum As Long)
    Dim packet As String
    
    With Item(ItemNum)
        packet = SPackets.Supdateitem & SEP_CHAR & ItemNum & SEP_CHAR & .Name & SEP_CHAR & .Pic & SEP_CHAR & .Type & SEP_CHAR & .Data1 & SEP_CHAR & .Data2 & SEP_CHAR & .Data3 & SEP_CHAR & .StrReq & SEP_CHAR & .DefReq & SEP_CHAR & .SpeedReq & SEP_CHAR & .MagicReq & SEP_CHAR & .ClassReq & SEP_CHAR & .AccessReq & SEP_CHAR
        packet = packet & .addHP & SEP_CHAR & .addMP & SEP_CHAR & .addSP & SEP_CHAR & .AddStr & SEP_CHAR & .AddDef & SEP_CHAR & .AddMagi & SEP_CHAR & .AddSpeed & SEP_CHAR & .AddEXP & SEP_CHAR & .Desc & SEP_CHAR & .AttackSpeed & SEP_CHAR & .Price & SEP_CHAR & .Stackable & SEP_CHAR & .Bound & SEP_CHAR & .LevelReq & SEP_CHAR & .HPReq & SEP_CHAR & .FPReq & SEP_CHAR & .Ammo
        packet = packet & SEP_CHAR & .AddCritChance & SEP_CHAR & .AddBlockChance & SEP_CHAR & .Cookable & END_CHAR
    End With
    
    Call SendDataToAll(packet)
End Sub

Sub SendUpdateItemTo(ByVal Index As Long, ByVal ItemNum As Long)
    Dim packet As String

    With Item(ItemNum)
        packet = SPackets.Supdateitem & SEP_CHAR & ItemNum & SEP_CHAR & .Name & SEP_CHAR & .Pic & SEP_CHAR & .Type & SEP_CHAR & .Data1 & SEP_CHAR & .Data2 & SEP_CHAR & .Data3 & SEP_CHAR & .StrReq & SEP_CHAR & .DefReq & SEP_CHAR & .SpeedReq & SEP_CHAR & .MagicReq & SEP_CHAR & .ClassReq & SEP_CHAR & .AccessReq & SEP_CHAR
        packet = packet & .addHP & SEP_CHAR & .addMP & SEP_CHAR & .addSP & SEP_CHAR & .AddStr & SEP_CHAR & .AddDef & SEP_CHAR & .AddMagi & SEP_CHAR & .AddSpeed & SEP_CHAR & .AddEXP & SEP_CHAR & .Desc & SEP_CHAR & .AttackSpeed & SEP_CHAR & .Price & SEP_CHAR & .Stackable & SEP_CHAR & .Bound & SEP_CHAR & .LevelReq & SEP_CHAR & .HPReq & SEP_CHAR & .FPReq & SEP_CHAR & .Ammo
        packet = packet & SEP_CHAR & .AddCritChance & SEP_CHAR & .AddBlockChance & SEP_CHAR & .Cookable & END_CHAR
    End With
    
    Call SendDataTo(Index, packet)
End Sub

Sub SendEditItemTo(ByVal Index As Long, ByVal ItemNum As Long)
    Call SendDataTo(Index, SPackets.Sedititem & SEP_CHAR & ItemNum & SEP_CHAR & Item(ItemNum).Name & END_CHAR)
End Sub

Sub SendUpdateEmoticonToAll(ByVal ItemNum As Long)
    Call SendDataToAll(SPackets.Supdateemoticon & SEP_CHAR & ItemNum & SEP_CHAR & Emoticons(ItemNum).Command & SEP_CHAR & Emoticons(ItemNum).Pic & END_CHAR)
End Sub

Sub SendUpdateEmoticonTo(ByVal Index As Long, ByVal ItemNum As Long)
    Call SendDataTo(Index, SPackets.Supdateemoticon & SEP_CHAR & ItemNum & SEP_CHAR & Emoticons(ItemNum).Command & SEP_CHAR & Emoticons(ItemNum).Pic & END_CHAR)
End Sub

Sub SendEditEmoticonTo(ByVal Index As Long, ByVal EmoNum As Long)
    Call SendDataTo(Index, SPackets.Seditemoticon & SEP_CHAR & EmoNum & SEP_CHAR & Emoticons(EmoNum).Command & SEP_CHAR & Emoticons(EmoNum).Pic & END_CHAR)
End Sub

Sub SendUpdateElementToAll(ByVal ElementNum As Long)
    Call SendDataToAll(SPackets.Supdateelement & SEP_CHAR & ElementNum & SEP_CHAR & Element(ElementNum).Name & SEP_CHAR & Element(ElementNum).Strong & SEP_CHAR & Element(ElementNum).Weak & END_CHAR)
End Sub

Sub SendUpdateElementTo(ByVal Index As Long, ByVal ElementNum As Long)
    Call SendDataTo(Index, SPackets.Supdateelement & SEP_CHAR & ElementNum & SEP_CHAR & Element(ElementNum).Name & SEP_CHAR & Element(ElementNum).Strong & SEP_CHAR & Element(ElementNum).Weak & END_CHAR)
End Sub

Sub SendEditElementTo(ByVal Index As Long, ByVal ElementNum As Long)
    Call SendDataTo(Index, SPackets.Seditelement & SEP_CHAR & ElementNum & SEP_CHAR & Element(ElementNum).Name & SEP_CHAR & Element(ElementNum).Strong & SEP_CHAR & Element(ElementNum).Weak & END_CHAR)
End Sub

Sub SendUpdateArrowToAll(ByVal ItemNum As Long)
    Call SendDataToAll(SPackets.Supdatearrow & SEP_CHAR & ItemNum & SEP_CHAR & Arrows(ItemNum).Name & SEP_CHAR & Arrows(ItemNum).Pic & SEP_CHAR & Arrows(ItemNum).Range & SEP_CHAR & Arrows(ItemNum).Amount & END_CHAR)
End Sub

Sub SendUpdateArrowTo(ByVal Index As Long, ByVal ItemNum As Long)
    Call SendDataTo(Index, SPackets.Supdatearrow & SEP_CHAR & ItemNum & SEP_CHAR & Arrows(ItemNum).Name & SEP_CHAR & Arrows(ItemNum).Pic & SEP_CHAR & Arrows(ItemNum).Range & SEP_CHAR & Arrows(ItemNum).Amount & END_CHAR)
End Sub

Sub SendEditArrowTo(ByVal Index As Long, ByVal EmoNum As Long)
    Call SendDataTo(Index, SPackets.Seditarrow & SEP_CHAR & EmoNum & SEP_CHAR & Arrows(EmoNum).Name & END_CHAR)
End Sub

Sub SendUpdateNpcToAll(ByVal NpcNum As Long)
    Dim packet As String
    Dim i As Long
    
    With NPC(NpcNum)
        packet = SPackets.Supdatenpc & SEP_CHAR & NpcNum & SEP_CHAR & .Name & SEP_CHAR & .AttackSay & SEP_CHAR & .Sprite & SEP_CHAR & .SpawnSecs & SEP_CHAR & .Behavior & SEP_CHAR & .Range & SEP_CHAR & .STR & SEP_CHAR & .DEF & SEP_CHAR & .Speed & SEP_CHAR & .Magi & SEP_CHAR & .Big & SEP_CHAR & .MAXHP & SEP_CHAR & .Exp & SEP_CHAR & .SpawnTime & SEP_CHAR & .Element & SEP_CHAR & .SPRITESIZE & SEP_CHAR
        
        For i = 1 To MAX_NPC_DROPS
            packet = packet & (.ItemNPC(i).Chance & SEP_CHAR & .ItemNPC(i).ItemNum & SEP_CHAR & .ItemNPC(i).ItemValue & SEP_CHAR)
        Next i
        
        packet = packet & .AttackSay2 & SEP_CHAR & .LEVEL & END_CHAR
    End With
    
    Call SendDataToAll(packet)
End Sub

Sub SendUpdateNpcTo(ByVal Index As Long, ByVal NpcNum As Long)
    Dim packet As String
    Dim i As Long

    With NPC(NpcNum)
        packet = SPackets.Supdatenpc & SEP_CHAR & NpcNum & SEP_CHAR & .Name & SEP_CHAR & .AttackSay & SEP_CHAR & .Sprite & SEP_CHAR & .SpawnSecs & SEP_CHAR & .Behavior & SEP_CHAR & .Range & SEP_CHAR & .STR & SEP_CHAR & .DEF & SEP_CHAR & .Speed & SEP_CHAR & .Magi & SEP_CHAR & .Big & SEP_CHAR & .MAXHP & SEP_CHAR & .Exp & SEP_CHAR & .SpawnTime & SEP_CHAR & .Element & SEP_CHAR & .SPRITESIZE & SEP_CHAR
        
        For i = 1 To MAX_NPC_DROPS
            packet = packet & (.ItemNPC(i).Chance & SEP_CHAR & .ItemNPC(i).ItemNum & SEP_CHAR & .ItemNPC(i).ItemValue & SEP_CHAR)
        Next i
        
        packet = packet & .AttackSay2 & SEP_CHAR & .LEVEL & END_CHAR
    End With
    
    Call SendDataTo(Index, packet)
End Sub

Sub SendEditNpcTo(ByVal Index As Long, ByVal NpcNum As Long)
    Call SendDataTo(Index, SPackets.Seditnpc & SEP_CHAR & NpcNum & SEP_CHAR & NPC(NpcNum).Name & END_CHAR)
End Sub

Sub SendShops(ByVal Index As Long)
    Dim i As Long

    For i = 1 To MAX_SHOPS
        If Trim$(Shop(i).Name) <> vbNullString Then
            Call SendUpdateShopTo(Index, i)
        End If
    Next i
End Sub

Sub SendUpdateShopToAll(ByVal ShopNum As Long)
    Dim packet As String
    Dim i As Long
    
    With Shop(ShopNum)
        packet = SPackets.Supdateshop & SEP_CHAR & ShopNum & SEP_CHAR & .Name & SEP_CHAR & .BuysItems & SEP_CHAR & .ShowInfo & SEP_CHAR & .CurrencyItem & SEP_CHAR
        
        For i = 1 To MAX_SHOP_ITEMS
            packet = packet & (.ShopItem(i).ItemNum & SEP_CHAR & .ShopItem(i).Amount & SEP_CHAR & .ShopItem(i).Price & SEP_CHAR & .ShopItem(i).CurrencyItem & SEP_CHAR)
        Next i
    End With
    
    packet = packet & END_CHAR

    Call SendDataToAll(packet)
End Sub

Sub SendUpdateShopTo(ByVal Index As Long, ByVal ShopNum)
    Dim packet As String
    Dim i As Long

    With Shop(ShopNum)
        packet = SPackets.Supdateshop & SEP_CHAR & ShopNum & SEP_CHAR & .Name & SEP_CHAR & .BuysItems & SEP_CHAR & .ShowInfo & SEP_CHAR & .CurrencyItem & SEP_CHAR
        For i = 1 To MAX_SHOP_ITEMS
            packet = packet & (.ShopItem(i).ItemNum & SEP_CHAR & .ShopItem(i).Amount & SEP_CHAR & .ShopItem(i).Price & SEP_CHAR & .ShopItem(i).CurrencyItem & SEP_CHAR)
        Next i
    End With
    
    packet = packet & END_CHAR

    Call SendDataTo(Index, packet)
End Sub

Sub SendEditShopTo(ByVal Index As Long, ByVal ShopNum As Long)
    Call SendDataTo(Index, SPackets.Seditshop & SEP_CHAR & ShopNum & SEP_CHAR & Shop(ShopNum).Name & END_CHAR)
End Sub

Sub SendSpells(ByVal Index As Long)
    Dim i As Long

    For i = 1 To MAX_SPELLS
        If Trim$(Spell(i).Name) <> vbNullString Then
            Call SendUpdateSpellTo(Index, i)
        End If
    Next i
End Sub

Sub SendUpdateSpellToAll(ByVal SpellNum As Long)
    Call SendDataToAll(SPackets.Supdatespell & SEP_CHAR & SpellNum & SEP_CHAR & Spell(SpellNum).Name & SEP_CHAR & Spell(SpellNum).ClassReq & SEP_CHAR & Spell(SpellNum).LevelReq & SEP_CHAR & Spell(SpellNum).Type & SEP_CHAR & Spell(SpellNum).Data1 & SEP_CHAR & Spell(SpellNum).Data2 & SEP_CHAR & Spell(SpellNum).Data3 & SEP_CHAR & Spell(SpellNum).MPCost & SEP_CHAR & Spell(SpellNum).Sound & SEP_CHAR & Spell(SpellNum).Range & SEP_CHAR & Spell(SpellNum).SpellAnim & SEP_CHAR & Spell(SpellNum).SpellTime & SEP_CHAR & Spell(SpellNum).SpellDone & SEP_CHAR & Spell(SpellNum).AE & SEP_CHAR & Spell(SpellNum).Big & SEP_CHAR & Spell(SpellNum).Element & SEP_CHAR & Spell(SpellNum).Stat & SEP_CHAR & Spell(SpellNum).StatTime & SEP_CHAR & Spell(SpellNum).Multiplier & SEP_CHAR & Spell(SpellNum).PassiveStat & SEP_CHAR & Spell(SpellNum).PassiveStatChange & SEP_CHAR & Spell(SpellNum).UsePassiveStat & SEP_CHAR & Spell(SpellNum).SelfSpell & END_CHAR)
End Sub

Sub SendUpdateSpellTo(ByVal Index As Long, ByVal SpellNum As Long)
    Call SendDataTo(Index, SPackets.Supdatespell & SEP_CHAR & SpellNum & SEP_CHAR & Spell(SpellNum).Name & SEP_CHAR & Spell(SpellNum).ClassReq & SEP_CHAR & Spell(SpellNum).LevelReq & SEP_CHAR & Spell(SpellNum).Type & SEP_CHAR & Spell(SpellNum).Data1 & SEP_CHAR & Spell(SpellNum).Data2 & SEP_CHAR & Spell(SpellNum).Data3 & SEP_CHAR & Spell(SpellNum).MPCost & SEP_CHAR & Spell(SpellNum).Sound & SEP_CHAR & Spell(SpellNum).Range & SEP_CHAR & Spell(SpellNum).SpellAnim & SEP_CHAR & Spell(SpellNum).SpellTime & SEP_CHAR & Spell(SpellNum).SpellDone & SEP_CHAR & Spell(SpellNum).AE & SEP_CHAR & Spell(SpellNum).Big & SEP_CHAR & Spell(SpellNum).Element & SEP_CHAR & Spell(SpellNum).Stat & SEP_CHAR & Spell(SpellNum).StatTime & SEP_CHAR & Spell(SpellNum).Multiplier & SEP_CHAR & Spell(SpellNum).PassiveStat & SEP_CHAR & Spell(SpellNum).PassiveStatChange & SEP_CHAR & Spell(SpellNum).UsePassiveStat & SEP_CHAR & Spell(SpellNum).SelfSpell & END_CHAR)
End Sub

Sub SendEditSpellTo(ByVal Index As Long, ByVal SpellNum As Long)
    Call SendDataTo(Index, SPackets.Seditspell & SEP_CHAR & SpellNum & SEP_CHAR & Spell(SpellNum).Name & END_CHAR)
End Sub

Sub SendRecipes(ByVal Index As Long)
    Dim i As Long
    
    For i = 1 To MAX_RECIPES
        If Trim$(Recipe(i).Name) <> vbNullString Then
            Call SendUpdateRecipeTo(Index, i)
        End If
    Next i
End Sub

Sub SendUpdateRecipeToAll(ByVal RecipeNum As Long)
    Call SendDataToAll(SPackets.Supdaterecipe & SEP_CHAR & RecipeNum & SEP_CHAR & Recipe(RecipeNum).Ingredient1 & SEP_CHAR & Recipe(RecipeNum).Ingredient2 & SEP_CHAR & Recipe(RecipeNum).ResultItem & SEP_CHAR & Recipe(RecipeNum).Name & END_CHAR)
End Sub

Sub SendUpdateRecipeTo(ByVal Index As Long, ByVal RecipeNum As Long)
    Call SendDataTo(Index, SPackets.Supdaterecipe & SEP_CHAR & RecipeNum & SEP_CHAR & Recipe(RecipeNum).Ingredient1 & SEP_CHAR & Recipe(RecipeNum).Ingredient2 & SEP_CHAR & Recipe(RecipeNum).ResultItem & SEP_CHAR & Recipe(RecipeNum).Name & END_CHAR)
End Sub

Sub SendEditRecipeTo(ByVal Index As Long, ByVal RecipeNum As Long)
    Call SendDataTo(Index, SPackets.Seditrecipe & SEP_CHAR & RecipeNum & SEP_CHAR & Recipe(RecipeNum).Ingredient1 & SEP_CHAR & Recipe(RecipeNum).Ingredient2 & SEP_CHAR & Recipe(RecipeNum).ResultItem & SEP_CHAR & Recipe(RecipeNum).Name & END_CHAR)
End Sub

Sub SendTrade(ByVal Index As Long, ByVal ShopNum As Long)
    Call SendDataTo(Index, SPackets.Sgoshop & SEP_CHAR & ShopNum & END_CHAR)
End Sub

Sub SendPlayerSpells(ByVal Index As Long)
    Dim packet As String
    Dim i As Long

    packet = SPackets.Sspells & SEP_CHAR
    
    For i = 1 To MAX_PLAYER_SPELLS
        packet = packet & (GetPlayerSpell(Index, i) & SEP_CHAR)
    Next i
    
    packet = packet & END_CHAR

    Call SendDataTo(Index, packet)
End Sub

Sub SendWeatherTo(ByVal Index As Long)
    If WeatherLevel <= 0 Then
        WeatherLevel = 1
    End If

    Call SendDataTo(Index, SPackets.Sweather & SEP_CHAR & WeatherType & SEP_CHAR & WeatherLevel & END_CHAR)
End Sub

Sub SendWeatherToAll()
    Dim i As Long
    Dim Weather As String

    Select Case WeatherType
        Case 0
            Weather = "None"
        Case 1
            Weather = "Rain"
        Case 2
            Weather = "Snow"
        Case 3
            Weather = "Thunder"
    End Select

    frmServer.Label5.Caption = "Current Weather: " & Weather

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            Call SendWeatherTo(i)
        End If
    Next i
End Sub

Sub SendNewsTo(ByVal Index As Long)
    Dim packet As String
    Dim RED As String, GREEN As String, BLUE As String

    On Error GoTo NewsError
    
    RED = ReadINI("COLOR", "Red", App.Path & "\News.ini", "255")
    GREEN = ReadINI("COLOR", "Green", App.Path & "\News.ini", "255")
    BLUE = ReadINI("COLOR", "Blue", App.Path & "\News.ini", "255")

    packet = SPackets.Snews & SEP_CHAR & ReadINI("DATA", "NewsTitle", App.Path & "\News.ini", vbNullString) & SEP_CHAR & ReadINI("DATA", "NewsBody", App.Path & "\News.ini", vbNullString) & SEP_CHAR & RED & SEP_CHAR & BLUE & SEP_CHAR & GREEN & END_CHAR

    Call SendDataTo(Index, packet)
    Exit Sub

NewsError:
    ' Error reading the news, so just send white
    RED = "255"
    GREEN = "255"
    BLUE = "255"
    
    packet = SPackets.Snews & SEP_CHAR & ReadINI("DATA", "NewsTitle", App.Path & "\News.ini", vbNullString) & SEP_CHAR & ReadINI("DATA", "NewsBody", App.Path & "\News.ini", vbNullString) & SEP_CHAR & RED & SEP_CHAR & BLUE & SEP_CHAR & GREEN & END_CHAR

    Call SendDataTo(Index, packet)
End Sub

Sub MapMsg2(ByVal MapNum As Long, ByVal Msg As String, ByVal Index As Long)
    Call SendDataToMap(MapNum, SPackets.Smapmsg2 & SEP_CHAR & Msg & SEP_CHAR & Index & END_CHAR)
End Sub

Sub Spin(ByVal Index As Long)
    Dim X As Long, Y As Long, MapNum As Long
    
    MapNum = GetPlayerMap(Index)

    Select Case GetPlayerDir(Index)
        Case DIR_DOWN
            X = GetPlayerX(Index)
            Y = GetPlayerY(Index) + 1
            
            Do While Y <= MAX_MAPY
                If Map(MapNum).Tile(X, Y).Type = TILE_TYPE_HOOKSHOT Then
                    Player(Index).HookShotX = X
                    Player(Index).HookShotY = Y - 1
                    
                    Call SpinFinish(Index, 1)
                    Exit Sub
                Else
                    If Map(MapNum).Tile(X, Y).Type = TILE_TYPE_BLOCKED Then
                        Call SpinFinish(Index, 0)
                        Exit Sub
                    End If
                End If
                
                Y = Y + 1
            Loop
            
            Call SpinFinish(Index, 0)
        Case DIR_UP
            X = GetPlayerX(Index)
            Y = GetPlayerY(Index) - 1
            
            Do While Y >= 0
                If Map(MapNum).Tile(X, Y).Type = TILE_TYPE_HOOKSHOT Then
                    Player(Index).HookShotX = X
                    Player(Index).HookShotY = Y + 1
                    
                    Call SpinFinish(Index, 1)
                    Exit Sub
                Else
                    If Map(MapNum).Tile(X, Y).Type = TILE_TYPE_BLOCKED Then
                        Call SpinFinish(Index, 0)
                        Exit Sub
                    End If
                End If
                Y = Y - 1
            Loop
            
            Call SpinFinish(Index, 0)
        Case DIR_RIGHT
            X = GetPlayerX(Index) + 1
            Y = GetPlayerY(Index)
            
            Do While X <= MAX_MAPX
                If Map(MapNum).Tile(X, Y).Type = TILE_TYPE_HOOKSHOT Then
                    Player(Index).HookShotX = X - 1
                    Player(Index).HookShotY = Y
                    
                    Call SpinFinish(Index, 1)
                    Exit Sub
                Else
                    If Map(MapNum).Tile(X, Y).Type = TILE_TYPE_BLOCKED Then
                        Call SpinFinish(Index, 0)
                        Exit Sub
                    End If
                End If
                X = X + 1
            Loop
            
            Call SpinFinish(Index, 0)
        Case DIR_LEFT
            X = GetPlayerX(Index) - 1
            Y = GetPlayerY(Index)
            
            Do While X >= 0
                If Map(MapNum).Tile(X, Y).Type = TILE_TYPE_HOOKSHOT Then
                    Player(Index).HookShotX = X + 1
                    Player(Index).HookShotY = Y
                    
                    Call SpinFinish(Index, 1)
                    Exit Sub
                Else
                    If Map(MapNum).Tile(X, Y).Type = TILE_TYPE_BLOCKED Then
                        Call SpinFinish(Index, 0)
                        Exit Sub
                    End If
                End If
                
                X = X - 1
            Loop
            
            Call SpinFinish(Index, 0)
    End Select
End Sub

Function GetFreeSlots(ByVal Index As Long)
   Dim i As Long, Slots As Integer

   Slots = 0

   For i = 1 To GetPlayerMaxInv(Index)
      If GetPlayerInvItemNum(Index, i) = 0 Then
         Slots = Slots + 1
      End If
   Next i

   GetFreeSlots = Slots
End Function

Function CanTake(ByVal Index As Long, ByVal ItemNum As Integer, ByVal Amount As Integer) As Boolean
   Dim i As Long
   
   For i = 1 To GetPlayerMaxInv(Index)
      If GetPlayerInvItemNum(Index, i) = ItemNum Then
         If GetPlayerInvItemValue(Index, i) >= Amount Then
            CanTake = True
            Exit Function
         End If
      End If
   Next i
   
   CanTake = False
End Function

Function CheckCloseCall(ByVal Index As Long) As Boolean
    If GetPlayerEquipSlotNum(Index, 4) <> 246 Then
        Exit Function
    End If
    
    Dim RandNum As Integer
    
    RandNum = Rand(1, 5)
    
    If RandNum = 1 Then
        Call SetPlayerHP(Index, 1)
        Call SendHP(Index)
        
        Call SendSoundToMap(GetPlayerMap(Index), "smrpg_1up.wav")

        ' Take away any PK status
        If GetPlayerPK(Index) = YES Then
            Call SetPlayerPK(Index, NO)
            Call SendPlayerData(Index)
            Call SendPlayerXY(Index)
        End If
        
        Call PlayerMsg(Index, "Your Close Call badge saved you!", YELLOW)
        
        CheckCloseCall = True
    End If
End Function

Sub LifeShroom(ByVal Index As Long)
    If HasItem(Index, 43) = 1 Then
        Call SetPlayerHP(Index, Item(43).Data1)
        Call SendHP(Index)
        Call TakeItem(Index, 43, 1)
        Call SendSoundToMap(GetPlayerMap(Index), "smrpg_1up.wav")
        Call PlayerMsg(Index, "Your Life Shroom revived you!", YELLOW)

        ' Take away any PK status
        If GetPlayerPK(Index) = YES Then
            Call SetPlayerPK(Index, NO)
            Call SendPlayerData(Index)
            Call SendPlayerXY(Index)
        End If
    End If
End Sub

Sub BlockPlayer(ByVal Index As Long)
    Dim PlayerDir As Long
    
    PlayerDir = GetPlayerDir(Index)
    
    Select Case PlayerDir
        Case 0
            Call PlayerWarp(Index, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index) + 1)
        Case 1
            Call PlayerWarp(Index, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index) - 1)
        Case 2
            Call PlayerWarp(Index, GetPlayerMap(Index), GetPlayerX(Index) + 1, GetPlayerY(Index))
        Case 3
            Call PlayerWarp(Index, GetPlayerMap(Index), GetPlayerX(Index) - 1, GetPlayerY(Index))
    End Select
End Sub

Sub PlayerDeath(ByVal Index As Long)
    Dim Exp As Integer

    Call OnDeath(Index)
    Call SetPlayerHP(Index, GetPlayerMaxHP(Index))
    Call SetPlayerMP(Index, GetPlayerMaxMP(Index))
    Call SetPlayerSP(Index, GetPlayerMaxSP(Index))
    Call SendHP(Index)
    Call SendMP(Index)
    Call SendSP(Index)
    
    ' Get rid of PK
    If GetPlayerPK(Index) = YES Then
        Call SetPlayerPK(Index, NO)
        Call SendPlayerData(Index)
        Call SendPlayerXY(Index)
    End If
    
    Exp = (GetPlayerExp(Index) \ 6)
    
    ' Makes sure exp cannot be less than 0
    If Exp <= 0 Then
        Exp = 0
    End If

    Call SetPlayerExp(Index, GetPlayerExp(Index) - Exp)
    Call BattleMsg(Index, "Oh no! You've died! You lost " & Exp & " experience points.", BRIGHTRED, 0)

    Call SendEXP(Index)
End Sub

Sub CanGetItemFromQuestionBlock(ByVal Index As Long, ByVal Items As Long, ByVal Value As Long)
    Dim Msg As String
    
    If Value > 1 Then
        If ItemIsStackable(Items) = True Then
            If (CanTake(Index, Items, 1) And GetFreeSlots(Index) >= 0) Or GetFreeSlots(Index) > 0 Then
                Call GiveItem(Index, Items, Value)
                Call PlayerMsg(Index, "You got " & Value & " " & Trim$(Item(Items).Name) & "s from the ? Block!", YELLOW)
            Else
                Call PlayerMsg(Index, "Your inventory is full! Please make some room in your inventory before you hit this block!", BRIGHTRED)
                Exit Sub
            End If
        Else
            If GetFreeSlots(Index) > 0 Then
                Call GiveItem(Index, Items, Value)
                Call PlayerMsg(Index, "You got " & Value & " " & Trim$(Item(Items).Name) & "s from the ? Block!", YELLOW)
            Else
                Call PlayerMsg(Index, "Your inventory is full! Please make some room in your inventory before you hit this block!", BRIGHTRED)
                Exit Sub
            End If
        End If
    Else
        If ItemIsStackable(Items) = True Then
            If (CanTake(Index, Items, 1) And GetFreeSlots(Index) >= 0) Or GetFreeSlots(Index) > 0 Then
                Call GiveItem(Index, Items, Value)
                
                If FindItemVowels(Items) = True Then
                    Msg = "You got an " & Trim$(Item(Items).Name) & " from the ? Block!"
                Else
                    Msg = "You got a " & Trim$(Item(Items).Name) & " from the ? Block!"
                End If
                
                Call PlayerMsg(Index, Msg, YELLOW)
            Else
                Call PlayerMsg(Index, "Your inventory is full! Please make some room in your inventory before you hit this block!", BRIGHTRED)
                Exit Sub
            End If
        Else
            If GetFreeSlots(Index) > 0 Then
                Call GiveItem(Index, Items, Value)
                
                If FindItemVowels(Items) = True Then
                    Msg = "You got an " & Trim$(Item(Items).Name) & " from the ? Block!"
                Else
                    Msg = "You got a " & Trim$(Item(Items).Name) & " from the ? Block!"
                End If
                
                Call PlayerMsg(Index, Msg, YELLOW)
            Else
                Call PlayerMsg(Index, "Your inventory is full! Please make some room in your inventory before you hit this block!", BRIGHTRED)
                Exit Sub
            End If
        End If
    End If
    
    Call SendSoundToMap(GetPlayerMap(Index), "m&lss_Block Hit.wav")
    Call PutVar(App.Path & "\Question Blocks.ini", GetPlayerName(Index), "Map: " & GetPlayerMap(Index) & "/X: " & GetPlayerX(Index) & "/Y: " & GetPlayerY(Index), "Hit")
End Sub

Sub HitQuestionBlock(ByVal Index As Long)
    Dim MapNum As Long, X As Long, Y As Long
    Dim Item1 As Long, Item2 As Long, Item3 As Long, Item4 As Long, Item5 As Long, Item6 As Long
    Dim Chance1 As Long, Chance2 As Long, Chance3 As Long, Chance4 As Long, Chance5 As Long, Chance6 As Long, OverallChance As Long
    Dim Value1 As Long, Value2 As Long, Value3 As Long, Value4 As Long, Value5 As Long, Value6 As Long
    
    If IsPlaying(Index) = False Then
        Exit Sub
    End If
    
    MapNum = GetPlayerMap(Index)
    X = GetPlayerX(Index)
    Y = GetPlayerY(Index)
    
    If GetVar(App.Path & "\Question Blocks.ini", GetPlayerName(Index), "Map: " & MapNum & "/X: " & X & "/Y: " & Y) <> "Hit" And Map(MapNum).Tile(X, Y).Type = TILE_TYPE_QUESTIONBLOCK Then
        With QuestionBlock(MapNum, X, Y)
            Item1 = .Item1
            Item2 = .Item2
            Item3 = .Item3
            Item4 = .Item4
            Item5 = .Item5
            Item6 = .Item6
            Chance1 = .Chance1
            Chance2 = .Chance2
            Chance3 = .Chance3
            Chance4 = .Chance4
            Chance5 = .Chance5
            Chance6 = .Chance6
            Value1 = .Value1
            Value2 = .Value2
            Value3 = .Value3
            Value4 = .Value4
            Value5 = .Value5
            Value6 = .Value6
        End With
    
        OverallChance = (Chance1 + Chance2 + Chance3 + Chance4 + Chance5 + Chance6)
        OverallChance = Int(Rand(1, OverallChance))
        Chance2 = Chance2 + Chance1
        Chance3 = Chance3 + Chance2
        Chance4 = Chance4 + Chance3
        Chance5 = Chance5 + Chance4
        Chance6 = Chance6 + Chance5
    
        If OverallChance <= Chance1 Then
            Call CanGetItemFromQuestionBlock(Index, Item1, Value1)
        ElseIf OverallChance <= Chance2 Then
            Call CanGetItemFromQuestionBlock(Index, Item2, Value2)
        ElseIf OverallChance <= Chance3 Then
            Call CanGetItemFromQuestionBlock(Index, Item3, Value3)
        ElseIf OverallChance <= Chance4 Then
            Call CanGetItemFromQuestionBlock(Index, Item4, Value4)
        ElseIf OverallChance <= Chance5 Then
           Call CanGetItemFromQuestionBlock(Index, Item5, Value5)
        ElseIf OverallChance <= Chance6 Then
            Call CanGetItemFromQuestionBlock(Index, Item6, Value6)
        End If
    End If
End Sub

Sub PlayerQueryBox(ByVal Index As Long, ByVal Message As String, ByVal Script As Long)
    Call SendDataTo(Index, SPackets.Squerybox & SEP_CHAR & Message & SEP_CHAR & Script & END_CHAR)
End Sub

Sub SendFavorTo(ByVal Index As Long, ByVal Title As String, ByVal Progress As String, ByVal Message As String, Optional ByVal Message2 As String = vbNullString)
    Call SendDataTo(Index, SPackets.Sfavor & SEP_CHAR & Title & SEP_CHAR & Progress & SEP_CHAR & Message & SEP_CHAR & Message2 & END_CHAR)
End Sub

Sub SendNpcTalkTo(ByVal Index As Long, ByVal NpcNum As Long, ByVal NPCText1 As String, Optional ByVal NPCText2 As String = vbNullString)
    If NPCText1 <> vbNullString Then
        Call SendDataTo(Index, SPackets.Snpctalk & SEP_CHAR & NpcNum & SEP_CHAR & NPCText1 & SEP_CHAR & NPCText2 & END_CHAR)
    End If
End Sub

Sub SendNpcTalkYesNoTo(ByVal Index As Long, ByVal NpcNum As Long, ByVal NpcText As String, ByVal YesText As String, ByVal NoText As String)
    Call SendDataTo(Index, SPackets.Snpctalkyesno & SEP_CHAR & NpcNum & SEP_CHAR & NpcText & SEP_CHAR & YesText & SEP_CHAR & NoText & END_CHAR)
End Sub

Sub SendMapPlayersInBattle(ByVal Index As Long, ByVal MapNum As Long)
    Dim packet As String
    Dim PlayerCount As Integer
    Dim i As Long
    
    If GetTotalMapPlayers(MapNum) <= 1 Then
        Exit Sub
    End If
    
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            If GetPlayerMap(i) = MapNum And i <> Index Then
                packet = packet & SEP_CHAR & i & SEP_CHAR & GetPlayerInBattle(i)
                PlayerCount = PlayerCount + 1
            End If
        End If
    Next i
    
    packet = SPackets.Smapplayersinbattle & SEP_CHAR & PlayerCount & packet & END_CHAR
    
    Call SendDataTo(Index, packet)
End Sub

Sub SendPlayerHeight(ByVal Index As Long)
    Call SendDataTo(Index, SPackets.Splayerheight & SEP_CHAR & Player(Index).Char(Player(Index).CharNum).Height & END_CHAR)
End Sub

Sub SendIsHiderFrozen(ByVal Index As Long, ByVal CanMove As Boolean)
    Call SendDataTo(Index, SPackets.Shiderfreeze & SEP_CHAR & CanMove & END_CHAR)
End Sub

Sub SendPlayingHideNSneak(ByVal Index As Long, ByVal IsPlaying As Boolean)
    Call SendDataTo(Index, SPackets.Shidensneak & SEP_CHAR & IsPlaying & END_CHAR)
End Sub

Sub SendWelcomeMsg(ByVal Index As Long, ByVal TitleText As String, ByVal MessageText As String)
    Call SendDataTo(Index, SPackets.Swelcomemsg & SEP_CHAR & TitleText & SEP_CHAR & MessageText & END_CHAR)
End Sub
