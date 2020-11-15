Attribute VB_Name = "modClientTCP"
Option Explicit

Sub TcpInit()
    SEP_CHAR = Chr$(0)
    END_CHAR = Chr$(237)

    PlayerBuffer = vbNullString

    frmMirage.Socket.RemoteHost = "127.0.0.1"
    frmMirage.Socket.RemotePort = 0000
End Sub

Sub TcpDestroy()
    frmMirage.Socket.Close
End Sub

Sub IncomingData(ByVal DataLength As Long)
    Dim Buffer As String, packet As String
    Dim Start As Long

    frmMirage.Socket.GetData Buffer, vbString, DataLength

    PlayerBuffer = PlayerBuffer & Buffer

    Start = InStr(PlayerBuffer, END_CHAR)
    Do While Start > 0
        packet = Mid$(PlayerBuffer, 1, Start - 1)
        PlayerBuffer = Mid$(PlayerBuffer, Start + 1, Len(PlayerBuffer))
        Start = InStr(PlayerBuffer, END_CHAR)
        If LenB(packet) > 0 Then
            Call HandleData(packet)
        End If
    Loop
End Sub

Function ConnectToServer() As Boolean
    If IsConnected Then
        ConnectToServer = True
        Exit Function
    End If

    Call TcpDestroy
    frmMirage.Socket.Connect

    If IsConnected Then
        ConnectToServer = True
    Else
        ConnectToServer = False
    End If
End Function

Function IsConnected() As Boolean
    If frmMirage.Socket.State = sckConnected Then
        IsConnected = True
    End If
End Function

Function IsPlaying(ByVal Index As Long) As Boolean
    If GetPlayerName(Index) <> vbNullString Then
        IsPlaying = True
    End If
End Function

Function IsAlphaNumeric(TestString As String) As Boolean
    Dim LoopID As Integer
    Dim sChar As String

    If LenB(TestString) > 0 Then
        For LoopID = 1 To Len(TestString)
            sChar = Mid(TestString, LoopID, 1)
            If Not sChar Like "[0-9A-Za-z]" Then
                Exit Function
            End If
        Next

        IsAlphaNumeric = True
    End If
End Function

Sub SendData(ByVal Data As String)
    Dim DBytes() As Byte
   
    DBytes = StrConv(Data, vbFromUnicode)

    If IsConnected Then
        frmMirage.Socket.SendData DBytes
    End If

    DoEvents
End Sub

Sub SendNewAccount(ByVal Name As String, ByVal Password As String)
    Call SendData(CPackets.Cnewaccount & SEP_CHAR & Trim$(Name) & SEP_CHAR & Trim$(Password) & SEP_CHAR & True & END_CHAR)
End Sub

Sub SendDelAccount(ByVal Name As String, ByVal Password As String)
    Call SendData(CPackets.Cdelaccount & SEP_CHAR & Trim$(Name) & SEP_CHAR & Trim$(Password) & SEP_CHAR & True & END_CHAR)
End Sub

Sub SendLogin(ByVal Name As String, ByVal Password As String, ByVal OwnerStatus As Boolean)
    Call SendData(CPackets.Cacclogin & SEP_CHAR & Trim$(Name) & SEP_CHAR & Trim$(Password) & SEP_CHAR & App.Major & SEP_CHAR & App.Minor & SEP_CHAR & App.Revision & SEP_CHAR & SEC_CODE & SEP_CHAR & OwnerStatus & END_CHAR)
End Sub

Sub SendAddChar(ByVal Name As String, ByVal Sex As Long, ByVal ClassNum As Long, ByVal slot As Long)
    Call SendData(CPackets.Caddchar & SEP_CHAR & Trim$(Name) & SEP_CHAR & Sex & SEP_CHAR & ClassNum & SEP_CHAR & slot & END_CHAR)
End Sub

Sub SendDelChar(ByVal slot As Long)
    Call SendData(CPackets.Cdelchar & SEP_CHAR & slot & END_CHAR)
End Sub

Sub SendGetClasses()
    Call SendData(CPackets.Cgetclasses & END_CHAR)
End Sub

Sub SendUseChar(ByVal CharSlot As Long)
    Call SendData(CPackets.Cusechar & SEP_CHAR & CharSlot & END_CHAR)
End Sub

Sub SayMsg(ByVal Text As String)
    Call SendData(CPackets.Csaymsg & SEP_CHAR & Text & END_CHAR)
End Sub

Sub GlobalMsg(ByVal Text As String)
    Call SendData(CPackets.Cglobalmsg & SEP_CHAR & Text & END_CHAR)
End Sub

Sub BroadcastMsg(ByVal Text As String)
    Call SendData(CPackets.Cbroadcastmsg & SEP_CHAR & Text & END_CHAR)
End Sub

Sub GroupMsg(ByVal Text As String)
    Call SendData(CPackets.Cgroupmsg & SEP_CHAR & Text & END_CHAR)
End Sub

Sub MapMsg(ByVal Text As String)
    Call SendData(CPackets.Cmapmsg & SEP_CHAR & Text & END_CHAR)
End Sub

Sub PlayerMsg(ByVal Text As String, ByVal MsgTo As String)
    Call SendData(CPackets.Cplayermsg & SEP_CHAR & MsgTo & SEP_CHAR & Text & END_CHAR)
End Sub

Sub OtherMsg(ByVal Text As String, ByVal MsgTo As String)
    Call SendData(CPackets.Cothermsg & SEP_CHAR & MsgTo & SEP_CHAR & Text & END_CHAR)
End Sub

Sub AdminMsg(ByVal Text As String)
    Call SendData(CPackets.Cadminmsg & SEP_CHAR & Text & END_CHAR)
End Sub

Sub SendPlayerMove()
    Call SendData(CPackets.Cplayermove & SEP_CHAR & GetPlayerDir(MyIndex) & SEP_CHAR & Player(MyIndex).Moving & SEP_CHAR & GetPlayerX(MyIndex) & SEP_CHAR & GetPlayerY(MyIndex) & END_CHAR)
End Sub

Sub SendPlayerDir()
    Call SendData(CPackets.Cplayerdir & SEP_CHAR & GetPlayerDir(MyIndex) & END_CHAR)
End Sub

Sub SendPlayerRequestNewMap(ByVal Dir As Long)
    Call SendData(CPackets.Crequestnewmap & SEP_CHAR & Dir & END_CHAR)
End Sub

Sub SendMap()
    Dim packet As String
    Dim x As Byte, y As Byte

    packet = CPackets.Cmapdata & SEP_CHAR & GetPlayerMap(MyIndex) & SEP_CHAR & Trim$(Map(GetPlayerMap(MyIndex)).Name) & SEP_CHAR & Map(GetPlayerMap(MyIndex)).Revision & SEP_CHAR & Map(GetPlayerMap(MyIndex)).Moral & SEP_CHAR & Map(GetPlayerMap(MyIndex)).Up & SEP_CHAR & Map(GetPlayerMap(MyIndex)).Down & SEP_CHAR & Map(GetPlayerMap(MyIndex)).Left & SEP_CHAR & Map(GetPlayerMap(MyIndex)).Right & SEP_CHAR & Map(GetPlayerMap(MyIndex)).music & SEP_CHAR & Map(GetPlayerMap(MyIndex)).BootMap & SEP_CHAR & Map(GetPlayerMap(MyIndex)).BootX & SEP_CHAR & Map(GetPlayerMap(MyIndex)).BootY & SEP_CHAR & Map(GetPlayerMap(MyIndex)).Indoors & SEP_CHAR & Map(GetPlayerMap(MyIndex)).Weather & SEP_CHAR

    For y = 0 To MAX_MAPY
        For x = 0 To MAX_MAPX
            With Map(GetPlayerMap(MyIndex)).Tile(x, y)
                packet = packet & (.Ground & SEP_CHAR & .Mask & SEP_CHAR & .Anim & SEP_CHAR & .Mask2 & SEP_CHAR & .M2Anim & SEP_CHAR & .Fringe & SEP_CHAR & .FAnim & SEP_CHAR & .Fringe2 & SEP_CHAR & .F2Anim & SEP_CHAR & .Type & SEP_CHAR & .Data1 & SEP_CHAR & .Data2 & SEP_CHAR & .Data3 & SEP_CHAR & .String1 & SEP_CHAR & .String2 & SEP_CHAR & .String3 & SEP_CHAR & .light & SEP_CHAR)
                packet = packet & (.GroundSet & SEP_CHAR & .MaskSet & SEP_CHAR & .AnimSet & SEP_CHAR & .Mask2Set & SEP_CHAR & .M2AnimSet & SEP_CHAR & .FringeSet & SEP_CHAR & .FAnimSet & SEP_CHAR & .Fringe2Set & SEP_CHAR & .F2AnimSet & SEP_CHAR)
            End With
        Next x
    Next y

    With Map(GetPlayerMap(MyIndex))
        For x = 1 To MAX_MAP_NPCS
            packet = packet & (.Npc(x) & SEP_CHAR & .SpawnX(x) & SEP_CHAR & .SpawnY(x) & SEP_CHAR)
        Next x
    End With
    
    For y = 0 To MAX_MAPY
        For x = 0 To MAX_MAPX
            With QuestionBlock(GetPlayerMap(MyIndex), x, y)
                packet = packet & (.Item1 & SEP_CHAR & .Item2 & SEP_CHAR & .Item3 & SEP_CHAR & .Item4 & SEP_CHAR & .Item5 & SEP_CHAR & .Item6 & SEP_CHAR & .Chance1 & SEP_CHAR & .Chance2 & SEP_CHAR & .Chance3 & SEP_CHAR & .Chance4 & SEP_CHAR & .Chance5 & SEP_CHAR & .Chance6 & SEP_CHAR & .Value1 & SEP_CHAR & .Value2 & SEP_CHAR & .Value3 & SEP_CHAR & .Value4 & SEP_CHAR & .Value5 & SEP_CHAR & .Value6 & SEP_CHAR)
            End With
        Next x
    Next y
    
    packet = packet & Map(GetPlayerMap(MyIndex)).owner & END_CHAR

    Call SendData(packet)
End Sub

Sub WarpMeTo(ByVal Name As String)
    Call SendData(CPackets.Cwarpmeto & SEP_CHAR & Name & END_CHAR)
End Sub

Sub WarpToMe(ByVal Name As String)
    Call SendData(CPackets.Cwarptome & SEP_CHAR & Name & END_CHAR)
End Sub

Sub WarpTo(ByVal MapNum As Long, ByVal x As Long, ByVal y As Long)
    Call SendData(CPackets.Cwarpto & SEP_CHAR & MapNum & SEP_CHAR & x & SEP_CHAR & y & END_CHAR)
End Sub

Sub LocalWarp(ByVal x As Long, ByVal y As Long)
    Call SendData(CPackets.Clocalwarp & SEP_CHAR & x & SEP_CHAR & y & END_CHAR)
End Sub

Sub SendSetAccess(ByVal Name As String, ByVal Access As Byte)
    Call SendData(CPackets.Csetaccess & SEP_CHAR & Name & SEP_CHAR & Access & END_CHAR)
End Sub

Sub SendKick(ByVal Name As String)
    Call SendData(CPackets.Ckickplayer & SEP_CHAR & Name & END_CHAR)
End Sub

Sub SendMute(ByVal Name As String)
    Call SendData(CPackets.Cmuteplayer & SEP_CHAR & Trim$(Name) & END_CHAR)
End Sub

Sub SendUnmute(ByVal Name As String)
    Call SendData(CPackets.Cunmuteplayer & SEP_CHAR & Trim$(Name) & END_CHAR)
End Sub

Sub SendMuteList()
    Call SendData(CPackets.Cgetmutelist & END_CHAR)
End Sub

Sub SendBan(ByVal Name As String)
    Call SendData(CPackets.Cbanplayer & SEP_CHAR & Trim$(Name) & SEP_CHAR & True & END_CHAR)
End Sub

Sub SendUnban(ByVal BanListNum As String)
    Call SendData(CPackets.Cunbanplayer & SEP_CHAR & Trim$(BanListNum) & SEP_CHAR & True & END_CHAR)
End Sub

Sub SendBanList()
    Call SendData(CPackets.Cgetbanlist & END_CHAR)
End Sub

Sub SendRequestEditItem()
    Call SendData(CPackets.Crequestedititem & END_CHAR)
End Sub

Sub SendSaveItem(ByVal ItemNum As Long)
    Dim packet As String
    packet = CPackets.Csaveitem & SEP_CHAR & ItemNum & SEP_CHAR & Trim$(Item(ItemNum).Name) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & SEP_CHAR & Item(ItemNum).Data1 & SEP_CHAR & Item(ItemNum).Data2 & SEP_CHAR & Item(ItemNum).Data3 & SEP_CHAR & Item(ItemNum).StrReq & SEP_CHAR & Item(ItemNum).DefReq & SEP_CHAR & Item(ItemNum).SpeedReq & SEP_CHAR & Item(ItemNum).MagicReq & SEP_CHAR & Item(ItemNum).ClassReq & SEP_CHAR & Item(ItemNum).AccessReq & SEP_CHAR & Item(ItemNum).AddHP & SEP_CHAR & Item(ItemNum).AddMP & SEP_CHAR & Item(ItemNum).AddSP & SEP_CHAR & Item(ItemNum).AddSTR & SEP_CHAR & Item(ItemNum).AddDef & SEP_CHAR & Item(ItemNum).AddSpeed & SEP_CHAR & Item(ItemNum).AddMAGI & SEP_CHAR & Item(ItemNum).AddEXP & SEP_CHAR & Item(ItemNum).desc & SEP_CHAR & Item(ItemNum).AttackSpeed & SEP_CHAR & Item(ItemNum).Price & SEP_CHAR & Item(ItemNum).Stackable & SEP_CHAR & Item(ItemNum).Bound & SEP_CHAR & Item(EditorIndex).LevelReq & SEP_CHAR & Item(EditorIndex).HPReq
    packet = packet & SEP_CHAR & Item(EditorIndex).FPReq & SEP_CHAR & Item(EditorIndex).Ammo & SEP_CHAR & Item(EditorIndex).AddCritChance & SEP_CHAR & Item(EditorIndex).AddBlockChance & SEP_CHAR & Item(EditorIndex).Cookable & END_CHAR
    Call SendData(packet)
End Sub

Sub SendRequestEditEmoticon()
    Call SendData(CPackets.Crequesteditemoticon & END_CHAR)
End Sub

Sub SendRequestEditElement()
    Call SendData(CPackets.Crequesteditelement & END_CHAR)
End Sub

Sub SendSaveEmoticon(ByVal EmoNum As Long)
    Call SendData(CPackets.Csaveemoticon & SEP_CHAR & EmoNum & SEP_CHAR & Trim$(Emoticons(EmoNum).Command) & SEP_CHAR & Emoticons(EmoNum).Pic & END_CHAR)
End Sub

Sub SendSaveElement(ByVal ElementNum As Long)
    Call SendData(CPackets.Csaveelement & SEP_CHAR & ElementNum & SEP_CHAR & Trim$(Element(ElementNum).Name) & SEP_CHAR & Element(ElementNum).Strong & SEP_CHAR & Element(ElementNum).Weak & END_CHAR)
End Sub

Sub SendRequestEditArrow()
    Call SendData(CPackets.Crequesteditarrow & END_CHAR)
End Sub

Sub SendSaveArrow(ByVal ArrowNum As Long)
    Call SendData(CPackets.Csavearrow & SEP_CHAR & ArrowNum & SEP_CHAR & Trim$(Arrows(ArrowNum).Name) & SEP_CHAR & Arrows(ArrowNum).Pic & SEP_CHAR & Arrows(ArrowNum).Range & SEP_CHAR & Arrows(ArrowNum).Amount & END_CHAR)
End Sub

Sub SendRequestEditNPC()
    Call SendData(CPackets.Crequesteditnpc & END_CHAR)
End Sub

Sub SendSaveNPC(ByVal NpcNum As Long)
    Dim packet As String
    Dim i As Long

    packet = CPackets.Csavenpc & SEP_CHAR & NpcNum & SEP_CHAR & Trim$(Npc(NpcNum).Name) & SEP_CHAR & Trim$(Npc(NpcNum).AttackSay) & SEP_CHAR & Npc(NpcNum).Sprite & SEP_CHAR & Npc(NpcNum).SpawnSecs & SEP_CHAR & Npc(NpcNum).Behavior & SEP_CHAR & Npc(NpcNum).Range & SEP_CHAR & Npc(NpcNum).STR & SEP_CHAR & Npc(NpcNum).DEF & SEP_CHAR & Npc(NpcNum).speed & SEP_CHAR & Npc(NpcNum).MAGI & SEP_CHAR & Npc(NpcNum).Big & SEP_CHAR & Npc(NpcNum).MaxHp & SEP_CHAR & Npc(NpcNum).Exp & SEP_CHAR & Npc(NpcNum).SpawnTime & SEP_CHAR & Npc(NpcNum).Element & SEP_CHAR & Npc(NpcNum).SpriteSize

    For i = 1 To MAX_NPC_DROPS
        packet = packet & (SEP_CHAR & Npc(NpcNum).ItemNPC(i).chance & SEP_CHAR & Npc(NpcNum).ItemNPC(i).ItemNum & SEP_CHAR & Npc(NpcNum).ItemNPC(i).ItemValue)
    Next i

    packet = packet & SEP_CHAR & Trim$(Npc(NpcNum).AttackSay2) & SEP_CHAR & Npc(NpcNum).Level & END_CHAR

    Call SendData(packet)
End Sub

Sub SendMapRespawn()
    Call SendData(CPackets.Cmaprespawn & END_CHAR)
End Sub

Sub SendUseItem(ByVal InvNum As Long)
    Call SendData(CPackets.Cuseitem & SEP_CHAR & InvNum & END_CHAR)
End Sub

Sub SendUseTurnBasedItem(ByVal InvNum As Long, ByVal ItemNum As Long)
    ' Only allow players to use HP, FP, and SP modifying items and scripted items
    If Item(ItemNum).Type = ITEM_TYPE_CHANGEHPFPSP Or Item(ItemNum).Type = ITEM_TYPE_SCRIPTED Then
        CanUseItem = False
        Call SendData(CPackets.Cuseturnbaseditem & SEP_CHAR & InvNum & END_CHAR)
    End If
End Sub

Sub SendDropItem(ByVal InvNum As Long, ByVal Amount As Long)
    Call SendData(CPackets.Cmapdropitem & SEP_CHAR & InvNum & SEP_CHAR & Amount & END_CHAR)
End Sub

Sub SendUseTurnBasedSpecial(ByVal SpellSlot As Long)
    Dim SpellNum As Long
    
    SpellNum = Player(MyIndex).Spell(SpellSlot)
    
    ' Check if the player has enough FP
    If GetPlayerMP(MyIndex) < FlowerSaver(SpellNum) Then
        Call AddText("You don't have enough FP to use this special attack!", BRIGHTRED)
        Exit Sub
    End If

    ' Make sure the player is at a high enough level
    If Spell(SpellNum).LevelReq > GetPlayerLevel(MyIndex) Then
        Call AddText("You need to be at level " & Spell(SpellNum).LevelReq & " to perform this special attack.", WHITE)
        Exit Sub
    End If
    
    CanUseSpecial = False
    
    Call SendData(CPackets.Cuseturnbasedspecial & SEP_CHAR & SpellSlot & END_CHAR)
End Sub

Sub SendOnlineList()
    Call SendData(CPackets.Conlinelist & END_CHAR)
End Sub

Sub SendMOTDChange(ByVal MOTD As String)
    Call SendData(CPackets.Csetmotd & SEP_CHAR & MOTD & END_CHAR)
End Sub

Sub SendRequestEditShop()
    Call SendData(CPackets.Crequesteditshop & END_CHAR)
End Sub

Sub SendSaveShop(ByVal ShopNum As Long)
    Dim packet As String
    Dim i As Integer

    packet = CPackets.Csaveshop & SEP_CHAR & ShopNum & SEP_CHAR & Trim$(Shop(ShopNum).Name) & SEP_CHAR & Shop(ShopNum).BuysItems & SEP_CHAR & Shop(ShopNum).ShowInfo & SEP_CHAR & Shop(ShopNum).currencyItem & SEP_CHAR

    For i = 1 To MAX_SHOP_ITEMS
        packet = packet & (Shop(ShopNum).ShopItem(i).ItemNum & SEP_CHAR & Shop(ShopNum).ShopItem(i).Amount & SEP_CHAR & Shop(ShopNum).ShopItem(i).Price & SEP_CHAR & Shop(ShopNum).ShopItem(i).currencyItem & SEP_CHAR)
    Next i

    packet = packet & END_CHAR

    Call SendData(packet)
End Sub

Sub SendRequestEditSpell()
    Call SendData(CPackets.Crequesteditspell & END_CHAR)
End Sub

Sub SendRequestEditRecipe()
    Call SendData(CPackets.Crequesteditrecipe & END_CHAR)
End Sub

Sub SendSaveSpell(ByVal SpellNum As Long)
    Call SendData(CPackets.Csavespell & SEP_CHAR & SpellNum & SEP_CHAR & Trim$(Spell(SpellNum).Name) & SEP_CHAR & Spell(SpellNum).ClassReq & SEP_CHAR & Spell(SpellNum).LevelReq & SEP_CHAR & Spell(SpellNum).Type & SEP_CHAR & Spell(SpellNum).Data1 & SEP_CHAR & Spell(SpellNum).Data2 & SEP_CHAR & Spell(SpellNum).Data3 & SEP_CHAR & Spell(SpellNum).MPCost & SEP_CHAR & Trim$(Spell(SpellNum).Sound) & SEP_CHAR & Spell(SpellNum).Range & SEP_CHAR & Spell(SpellNum).SpellAnim & SEP_CHAR & Spell(SpellNum).SpellTime & SEP_CHAR & Spell(SpellNum).SpellDone & SEP_CHAR & Spell(SpellNum).AE & SEP_CHAR & Spell(SpellNum).Big & SEP_CHAR & Spell(SpellNum).Element & SEP_CHAR & Spell(SpellNum).Stat & SEP_CHAR & Spell(SpellNum).StatTime & SEP_CHAR & Spell(SpellNum).Multiplier & SEP_CHAR & Spell(SpellNum).PassiveStat & SEP_CHAR & Spell(SpellNum).PassiveStatChange & SEP_CHAR & Spell(SpellNum).UsePassiveStat & SEP_CHAR & Spell(SpellNum).SelfSpell & END_CHAR)
End Sub

Sub SendSaveRecipe(ByVal RecipeNum As Long)
    Call SendData(CPackets.Csaverecipe & SEP_CHAR & RecipeNum & SEP_CHAR & Recipe(RecipeNum).Ingredient1 & SEP_CHAR & Recipe(RecipeNum).Ingredient2 & SEP_CHAR & Recipe(RecipeNum).ResultItem & SEP_CHAR & Trim$(Recipe(RecipeNum).Name) & END_CHAR)
End Sub

Sub SendRequestEditMap()
    Call SendData(CPackets.Crequesteditmap & END_CHAR)
End Sub

Sub SendTradeRequest(ByVal Name As String)
    If IsBanking = True Or IsCooking = True Or IsShopping = True Then
        Call AddText("You cannot trade with another player while you are busy with another activity!", RED)
        Exit Sub
    End If
    
    ' Stop trading in STS and Dodgebill
    If GetPlayerMap(MyIndex) = 33 Or GetPlayerMap(MyIndex) = 188 Then
        Call AddText("There's no time to trade while you're in a minigame!", RED)
        Exit Sub
    End If
    
    Call SendData(CPackets.Ctraderequest & SEP_CHAR & Name & END_CHAR)
End Sub

Sub SendAcceptTrade()
    Call SendData(CPackets.Caccepttrade & END_CHAR)
    frmTradeBox.Visible = False
End Sub

Sub SendDeclineTrade()
    Call SendData(CPackets.Cdeclinetrade & END_CHAR)
    frmTradeBox.Visible = False
End Sub

Sub SendPartyRequest(ByVal Name As String)
    Call SendData(CPackets.Cparty & SEP_CHAR & Name & END_CHAR)
End Sub

Sub SendJoinParty()
    Call SendData(CPackets.Cjoinparty & END_CHAR)
End Sub

Sub SendDeclineParty()
    Call SendData(CPackets.Cpartydecline & END_CHAR)
End Sub

Sub SendLeaveParty()
    Call SendData(CPackets.Cleaveparty & END_CHAR)
End Sub

Sub SendSetPlayerSprite(ByVal Name As String, ByVal SpriteNum As Byte)
    Call SendData(CPackets.Csetplayersprite & SEP_CHAR & Name & SEP_CHAR & SpriteNum & END_CHAR)
End Sub

Sub SendHotScript(ByVal Value As Byte)
    Call SendData(CPackets.Chotscript & SEP_CHAR & Value & END_CHAR)
End Sub

Sub SendPlayerMoveMouse()
    Call SendData(CPackets.Cplayermovemouse & SEP_CHAR & GetPlayerDir(MyIndex) & END_CHAR)
End Sub

Sub SendChangeDir()
    Call SendData(CPackets.Cwarp & SEP_CHAR & GetPlayerDir(MyIndex) & END_CHAR)
End Sub

Sub SendGuildLeave()
    Call SendData(CPackets.Cguildleave & END_CHAR)
End Sub

Sub SendGuildMember(ByVal Name As String)
    Call SendData(CPackets.Cguildmember & SEP_CHAR & Name & END_CHAR)
End Sub

Sub SendGuildMemberRequest(ByVal Name As String, ByVal Trainee As Integer)
    Call SendData(CPackets.Cguildmemberrequest & SEP_CHAR & Name & SEP_CHAR & Trainee & END_CHAR)
End Sub

Sub SendGuildMemberDecline(ByVal Name As String)
    Call SendData(CPackets.Cguildmemberdecline & SEP_CHAR & Name & END_CHAR)
End Sub

Sub SendRequestSpells()
    Call SendData(CPackets.Cspells & END_CHAR)
End Sub

Sub SendForgetSpell(ByVal SpellID As Long)
    If Player(MyIndex).Spell(SpellID) > 0 Then
        If GetTickCount > Player(MyIndex).AttackTimer + 1000 Then
            If MsgBox("Are you sure you want to forget this special attack?", vbYesNo, "Forget Special Attack") = vbYes Then
                Call SendData(CPackets.Cforgetspell & SEP_CHAR & SpellID & END_CHAR)
                frmMirage.picPlayerSpells.Visible = False
            End If
        End If
    Else
        Call AddText("There's no special attack here!", WHITE)
    End If
End Sub

Sub SendRequestMyStats()
    Call SendData(CPackets.Cgetstats & SEP_CHAR & GetPlayerName(MyIndex) & END_CHAR)
End Sub

Sub SendSetTrainee(ByVal Name As String)
    Call SendData(CPackets.Cguildtrainee & SEP_CHAR & Name & END_CHAR)
End Sub

Sub SendGuildDisown(ByVal Name As String)
    Call SendData(CPackets.Cguilddisown & SEP_CHAR & Name & END_CHAR)
End Sub

Sub SendChangeGuildAccess(ByVal Name As String, ByVal AccessLvl As Long)
    Call SendData(CPackets.Cguildchangeaccess & SEP_CHAR & Name & SEP_CHAR & AccessLvl & END_CHAR)
End Sub

Sub SendPlayerChat(ByVal Name As String)
    Call SendData(CPackets.Cplayerchat & SEP_CHAR & Name & END_CHAR)
End Sub

Sub SendMakeAdmin()
    Call SendData(CPackets.Cmakeadmin & END_CHAR)
End Sub

Sub SendRunFromBattle()
    Call SendData(CPackets.Crunfrombattle & END_CHAR)
End Sub

Sub SendFinishPlayerBattle()
    Call SendData(CPackets.Cfinishplayerbattle & SEP_CHAR & BattleNPC & END_CHAR)
End Sub

Sub CookItem(ByVal FirstItemSlot As Long, Optional ByVal SecondItemSlot As Long = 0)
    Call SendData(CPackets.Ccookitem & SEP_CHAR & FirstItemSlot & SEP_CHAR & SecondItemSlot & SEP_CHAR & CookNpcNum & END_CHAR)
    IsCooking = True
End Sub

Sub FinishCooking(ByVal RecipeNum As Long)
    Call SendData(CPackets.Ccooking & SEP_CHAR & RecipeNum & SEP_CHAR & CookNpcNum & END_CHAR)
End Sub

Sub SendUseSpecialBadge(ByVal ItemNum As Long)
    Call SendData(CPackets.Cusespecialbadge & SEP_CHAR & ItemNum & END_CHAR)
End Sub

Sub NotifyOtherPlayer(ByVal Name As String)
    Call SendData(CPackets.Cnotifyotherplayer & SEP_CHAR & Name & END_CHAR)
End Sub

Sub JugemsCloudWarp()
    Call SendData(CPackets.Cjugemscloudwarp & END_CHAR)
End Sub

Sub SendGetPlayerInfo(ByVal Name As String)
    Call SendData(CPackets.Cgetplayerinfo & SEP_CHAR & Name & END_CHAR)
End Sub
