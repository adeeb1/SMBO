Attribute VB_Name = "modGameLogic"
Option Explicit

Function GetPlayerDamage(ByVal Index As Long) As Long

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or Index <= 0 Or Index > MAX_PLAYERS Then
        Exit Function
    End If

    GetPlayerDamage = GetPlayerSTR(Index)

    If GetPlayerDamage < 0 Then
        GetPlayerDamage = 0
    End If
End Function

Function GetPlayerProtection(ByVal Index As Long) As Long

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or Index <= 0 Or Index > MAX_PLAYERS Then
        Exit Function
    End If

    GetPlayerProtection = GetPlayerDEF(Index)
End Function

Function FindOpenPlayerSlot() As Long
    Dim i As Long

    For i = 1 To MAX_PLAYERS
        If Not IsConnected(i) Then
            FindOpenPlayerSlot = i
            Exit Function
        End If
    Next i
End Function

Public Function FindOpenInvSlot(ByVal Index As Long, ByVal ItemNum As Long) As Long
    Dim i As Long

    ' Check for subscript out of range.
    If ItemNum < 1 Or ItemNum > MAX_ITEMS Then
        Exit Function
    End If

    ' Check for subscript out of range.
    If Not IsPlaying(Index) Then
        Exit Function
    End If

    ' Check to see if they already have the item.
    If ItemIsStackable(ItemNum) = True Then
        For i = 1 To GetPlayerMaxInv(Index)
            If GetPlayerInvItemNum(Index, i) = ItemNum Then
                FindOpenInvSlot = i
                Exit Function
            End If
        Next i
    End If

    ' Try to find an open inventory slot.
    For i = 1 To GetPlayerMaxInv(Index)
        If GetPlayerInvItemNum(Index, i) = 0 Then
            FindOpenInvSlot = i
            Exit Function
        End If
    Next i
End Function

Function FindOpenBankSlot(ByVal Index As Long, ByVal ItemNum As Long) As Long
    Dim i As Long

    ' Check for subscript out of range.
    If ItemNum < 1 Or ItemNum > MAX_ITEMS Then
        Exit Function
    End If

    ' Check for subscript out of range.
    If Not IsPlaying(Index) Then
        Exit Function
    End If

    ' Check to see if they already have the item.
    If ItemIsStackable(ItemNum) = True Then
        For i = 1 To MAX_BANK
            If GetPlayerBankItemNum(Index, i) = ItemNum Then
                FindOpenBankSlot = i
                Exit Function
            End If
        Next i
    End If

    ' Try to find an open bank slot.
    For i = 1 To MAX_BANK
        If GetPlayerBankItemNum(Index, i) = 0 Then
            FindOpenBankSlot = i
            Exit Function
        End If
    Next i
End Function

Function FindOpenMapItemSlot(ByVal MapNum As Long) As Long
    Dim i As Long

    ' Check for subscript out of range.
    If MapNum < 1 Or MapNum > MAX_MAPS Then
        Exit Function
    End If

    For i = 1 To MAX_MAP_ITEMS
        If MapItem(MapNum, i).num = 0 Then
            FindOpenMapItemSlot = i
            Exit Function
        End If
    Next i
End Function

Function FindOpenSpellSlot(ByVal Index As Long) As Long
    Dim i As Long

    For i = 1 To MAX_PLAYER_SPELLS
        If GetPlayerSpell(Index, i) = 0 Then
            FindOpenSpellSlot = i
            Exit Function
        End If
    Next i
End Function

Function HasSpell(ByVal Index As Long, ByVal SpellNum As Long) As Boolean
    Dim i As Long

    For i = 1 To MAX_PLAYER_SPELLS
        If GetPlayerSpell(Index, i) = SpellNum Then
            HasSpell = True
            Exit Function
        End If
    Next i
End Function

Function TotalOnlinePlayers() As Long
    Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            TotalOnlinePlayers = TotalOnlinePlayers + 1
        End If
    Next i
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

Function HasItem(ByVal Index As Long, ByVal ItemNum As Long) As Long
    Dim i As Long

    ' Check for subscript out of range.
    If ItemNum < 1 Or ItemNum > MAX_ITEMS Then
        Exit Function
    End If

    ' Check for subscript out of range.
    If Not IsPlaying(Index) Then
        Exit Function
    End If

    ' Check to see if the player has the item.
    For i = 1 To GetPlayerMaxInv(Index)
        If GetPlayerInvItemNum(Index, i) = ItemNum Then
            If ItemIsStackable(ItemNum) = True Then
                HasItem = GetPlayerInvItemValue(Index, i)
            Else
                HasItem = 1
            End If

            Exit Function
        End If
    Next i
End Function

Sub TakeSpecificItem(ByVal Index As Long, ByVal ItemSlot As Long, ByVal ItemVal As Long)
    Dim ItemNum As Long
    Dim TakeSpecificItem As Boolean

    TakeSpecificItem = False
    ItemNum = GetPlayerInvItemNum(Index, ItemSlot)
    
    ' Check for subscript out of range
    If IsPlaying(Index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Sub
    End If
    
    ' Check to see if the player has the item
    If ItemIsStackable(ItemNum) = True Then
        ' Is what we are trying to take away more then what they have?  If so just set it to zero
        If ItemVal >= GetPlayerInvItemValue(Index, ItemSlot) Then
            TakeSpecificItem = True
        Else
            Call SetPlayerInvItemValue(Index, ItemSlot, GetPlayerInvItemValue(Index, ItemSlot) - ItemVal)
            Call SendInventoryUpdate(Index, ItemSlot)
        End If
    Else
        TakeSpecificItem = True
    End If
        
    If TakeSpecificItem = True Then
        Call SetPlayerInvItemNum(Index, ItemSlot, 0)
        Call SetPlayerInvItemValue(Index, ItemSlot, 0)
        Call SetPlayerInvItemAmmo(Index, ItemSlot, -1)

        ' Send the inventory update
        Call SendInventoryUpdate(Index, ItemSlot)
    End If
End Sub

Sub TakeItem(ByVal Index As Long, ByVal ItemNum As Long, ByVal ItemVal As Long)
    Dim i As Long, ItemValue As Long
    Dim TakeItem As Boolean, TakeCurrency As Boolean

    TakeItem = False
    TakeCurrency = False
    ItemValue = ItemVal
    
    ' Check for subscript out of range
    If IsPlaying(Index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Sub
    End If

    For i = 1 To GetPlayerMaxInv(Index)
        ' Check to see if the player has the item
        If GetPlayerInvItemNum(Index, i) = ItemNum Then
            If ItemIsStackable(ItemNum) = True Then
                ' Is what we are trying to take away more then what they have?  If so just set it to zero
                If ItemVal >= GetPlayerInvItemValue(Index, i) Then
                    TakeCurrency = True
                Else
                    Call SetPlayerInvItemValue(Index, i, GetPlayerInvItemValue(Index, i) - ItemVal)
                    Call SendInventoryUpdate(Index, i)
                    Exit Sub
                End If
            Else
                TakeItem = True
            End If
            
            If TakeCurrency = True Then
                Call SetPlayerInvItemNum(Index, i, 0)
                Call SetPlayerInvItemValue(Index, i, 0)
                Call SetPlayerInvItemAmmo(Index, i, -1)

                ' Send the inventory update
                Call SendInventoryUpdate(Index, i)
                Exit Sub
            End If
            
            If TakeItem = True And ItemValue > 0 Then
                Call SetPlayerInvItemNum(Index, i, 0)
                Call SetPlayerInvItemValue(Index, i, 0)
                Call SetPlayerInvItemAmmo(Index, i, -1)
                ItemValue = ItemValue - 1
            ElseIf TakeItem = True And ItemValue <= 0 Then
                Exit For
            End If
        End If
    Next i
    Call SendInventory(Index)
End Sub

Sub GiveItem(ByVal Index As Long, ByVal ItemNum As Long, ByVal ItemVal As Long)
    Dim i As Long, j As Long, ItemValue As Long
    
    ItemValue = ItemVal
    
    ' Check for subscript out of range.
    If ItemNum < 1 Or ItemNum > MAX_ITEMS Then
        Exit Sub
    End If

    ' Check for subscript out of range.
    If Not IsPlaying(Index) Then
        Exit Sub
    End If
    
    ' Check to see if inventory is full
    For i = 1 To GetPlayerMaxInv(Index)
        j = FindOpenInvSlot(Index, ItemNum)
        If j > 0 Then
            If ItemIsStackable(ItemNum) = True Then
                Call SetPlayerInvItemNum(Index, j, ItemNum)
                Call SetPlayerInvItemValue(Index, j, GetPlayerInvItemValue(Index, j) + ItemVal)
                Call SetPlayerInvItemAmmo(Index, j, Item(ItemNum).Ammo)
                Call SendInventoryUpdate(Index, j)
                Exit Sub
            Else
                If ItemValue > 0 Then
                    Call SetPlayerInvItemNum(Index, j, ItemNum)
                    Call SetPlayerInvItemValue(Index, j, 1)
                    Call SetPlayerInvItemAmmo(Index, j, Item(ItemNum).Ammo)
                    Call SendInventoryUpdate(Index, j)
                    ItemValue = ItemValue - 1
                Else
                    Exit Sub
                End If
            End If
        Else
            If GetFreeSlots(Index) = 0 And ItemValue > 0 Then
                Call PlayerMsg(Index, "Your inventory is full! Please make some room for this item.", BRIGHTRED)
                Exit Sub
            End If
        End If
    Next i
End Sub

Sub GiveItemForBank(ByVal Index As Long, ByVal ItemNum As Long, ByVal ItemVal As Long, ByVal BankSlot As Long)
    Dim i As Long

    ' Check for subscript out of range.
    If ItemNum < 1 Or ItemNum > MAX_ITEMS Then
        Exit Sub
    End If

    ' Check for subscript out of range.
    If Not IsPlaying(Index) Then
        Exit Sub
    End If

    i = FindOpenInvSlot(Index, ItemNum)

    ' Check to see if inventory is full
    If i > 0 Then
        Call SetPlayerInvItemNum(Index, i, ItemNum)
        Call SetPlayerInvItemValue(Index, i, GetPlayerInvItemValue(Index, i) + ItemVal)
        Call SetPlayerInvItemAmmo(Index, i, GetPlayerBankItemAmmo(Index, BankSlot))

        Call SendInventoryUpdate(Index, i)
    Else
        Call PlayerMsg(Index, "Your inventory is full! Please make some room for this item.", BRIGHTRED)
    End If
End Sub

Sub TakeBankItem(ByVal Index As Long, ByVal ItemNum As Long, ByVal ItemVal As Long)
    Dim i As Long, n As Long
    Dim TakeBankItem As Boolean
    Dim BankItem As Long

    TakeBankItem = False

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Sub
    End If

    For i = 1 To MAX_BANK
      BankItem = GetPlayerBankItemNum(Index, i)
        ' Check to see if the player has the item
        If BankItem = ItemNum Then
            If ItemIsStackable(ItemNum) = True Then
                ' Is what we are trying to take away more then what they have? If so just set it to zero
                If ItemVal >= GetPlayerBankItemValue(Index, i) Then
                    TakeBankItem = True
                Else
                    Call SetPlayerBankItemValue(Index, i, GetPlayerBankItemValue(Index, i) - ItemVal)
                    Call SendBankUpdate(Index, i)
                End If
            Else
                TakeBankItem = True
            End If

            If TakeBankItem = True Then
                Call SetPlayerBankItemNum(Index, i, 0)
                Call SetPlayerBankItemValue(Index, i, 0)
                Call SetPlayerBankItemAmmo(Index, i, -1)

                ' Send the Bank update
                Call SendBankUpdate(Index, i)
                Exit Sub
            End If
        End If
    Next i
End Sub

Sub GiveBankItem(ByVal Index As Long, ByVal ItemNum As Long, ByVal ItemVal As Long, ByVal BankSlot As Long, ByVal InvNum As Long)
    Dim i As Long

    ' Check for subscript out of range.
    If ItemNum < 1 Or ItemNum > MAX_ITEMS Then
        Exit Sub
    End If

    ' Check for subscript out of range.
    If Not IsPlaying(Index) Then
        Exit Sub
    End If

    i = BankSlot

    ' Check to see if Bank inventory is full
    If i > 0 Then
        Call SetPlayerBankItemNum(Index, i, ItemNum)
        Call SetPlayerBankItemValue(Index, i, GetPlayerBankItemValue(Index, i) + ItemVal)
        Call SetPlayerBankItemAmmo(Index, i, GetPlayerInvItemAmmo(Index, InvNum))
    Else
        Call SendDataTo(Index, SPackets.Sbankmsg & SEP_CHAR & "Bank full!" & END_CHAR)
    End If
End Sub

Sub SpawnItem(ByVal ItemNum As Long, ByVal ItemVal As Long, ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long)
    Dim i As Long, n As Long

    ' Check for subscript out of range.
    If ItemNum < 0 Or ItemNum > MAX_ITEMS Then
        Exit Sub
    End If

    ' Check for subscript out of range.
    If MapNum < 1 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If

    ' Find open map item slot.
    If ItemIsStackable(ItemNum) = False Then
        For n = 1 To ItemVal
            i = FindOpenMapItemSlot(MapNum)
            
            If i > 0 Then
                Call SpawnItemSlot(i, ItemNum, 1, Item(ItemNum).Ammo, MapNum, X, Y)
            End If
        Next
    Else
        i = FindOpenMapItemSlot(MapNum)
    
        If i > 0 Then
            Call SpawnItemSlot(i, ItemNum, ItemVal, Item(ItemNum).Ammo, MapNum, X, Y)
        End If
    End If
End Sub

Sub SpawnItemSlot(ByVal MapItemSlot As Long, ByVal ItemNum As Long, ByVal ItemVal As Long, ByVal ItemAmmo As Long, ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long)
    Dim i As Long

    ' Check for subscript out of range.
    If MapNum < 1 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If

    ' Check for subscript out of range.
    If ItemNum < 0 Or ItemNum > MAX_ITEMS Then
        Exit Sub
    End If

    ' Check for subscript out of range.
    If MapItemSlot < 1 Or MapItemSlot > MAX_MAP_ITEMS Then
        Exit Sub
    End If

    i = MapItemSlot

    If i > 0 Then
        MapItem(MapNum, i).num = ItemNum
        MapItem(MapNum, i).Value = ItemVal

        If Item(ItemNum).Type >= ITEM_TYPE_WEAPON And Item(ItemNum).Type <= ITEM_TYPE_MUSHROOMBADGE Then
            MapItem(MapNum, i).Ammo = ItemAmmo
        Else
            MapItem(MapNum, i).Ammo = -1
        End If
    
        MapItem(MapNum, i).X = X
        MapItem(MapNum, i).Y = Y
        
        Call SendSpawnItemSlot(MapNum, i, ItemNum, ItemVal, X, Y, MapItem(MapNum, i).Ammo)
    End If
End Sub

Sub SpawnAllMapsItems()
    Dim i As Long

    For i = 1 To MAX_MAPS
        Call SpawnMapItems(i)
    Next i
End Sub

Sub SpawnMapItems(ByVal MapNum As Long)
    Dim X As Integer
    Dim Y As Integer

    ' Check for subscript out of range.
    If MapNum < 1 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If

    ' Spawn all the mapped items on their specified tile.
    For Y = 0 To MAX_MAPY
        For X = 0 To MAX_MAPX
            With Map(MapNum).Tile(X, Y)
                If .Type = TILE_TYPE_ITEM Then
                    If ItemIsStackable(.Data1) = True And .Data2 <= 0 Then
                        Call SpawnItem(.Data1, 1, MapNum, X, Y)
                    Else
                        Call SpawnItem(.Data1, .Data2, MapNum, X, Y)
                    End If
                End If
            End With
        Next X
    Next Y
End Sub

Sub PlayerMapGetItem(ByVal Index As Long)
    Dim i As Long, n As Long, MapNum As Long, ItemNum As Long
    Dim Msg As String

    If IsPlaying(Index) = False Then
        Exit Sub
    End If

    MapNum = GetPlayerMap(Index)

    For i = 1 To MAX_MAP_ITEMS
        ' See if theres even an item here
        With MapItem(MapNum, i)
            If .num > 0 And .num <= MAX_ITEMS Then
                ' Check if item is at the same location as the player
                If .X = GetPlayerX(Index) Then
                    If .Y = GetPlayerY(Index) Then
                        ' Find open slot
                        n = FindOpenInvSlot(Index, .num)
                        ' Open slot available?
                        If n <> 0 Then
                            ' Set item in players inventory
                            Call SetPlayerInvItemNum(Index, n, .num)
                            ItemNum = GetPlayerInvItemNum(Index, n)
                            
                            If ItemIsStackable(ItemNum) = True And Item(ItemNum).Type <> ITEM_TYPE_AMMO Then
                                If MapItem(MapNum, i).Value = 1 Then
                                    Call SetPlayerInvItemValue(Index, n, GetPlayerInvItemValue(Index, n) + .Value)
                                           
                                    If FindItemVowels(ItemNum) = True Then
                                        Msg = "You pick up an " & Trim$(Item(ItemNum).Name) & "."
                                    Else
                                        Msg = "You pick up a " & Trim$(Item(ItemNum).Name) & "."
                                    End If
                                Else
                                    Call SetPlayerInvItemValue(Index, n, GetPlayerInvItemValue(Index, n) + .Value)
                                    Msg = "You pick up " & .Value & " " & Trim$(Item(ItemNum).Name) & "s."
                                End If
                            ElseIf Item(ItemNum).Type = ITEM_TYPE_AMMO Then
                                If .Value = 1 Then
                                    Call SetPlayerInvItemValue(Index, n, GetPlayerInvItemValue(Index, n) + .Value)
                                    Msg = "You pick up " & .Value & " " & Trim$(Item(ItemNum).Name) & "."
                                Else
                                    Call SetPlayerInvItemValue(Index, n, GetPlayerInvItemValue(Index, n) + .Value)
                                    Msg = "You pick up " & .Value & " " & Trim$(Item(ItemNum).Name) & "s."
                                End If
                            Else
                                Call SetPlayerInvItemValue(Index, n, 1)
                                
                                If FindItemVowels(ItemNum) = True Then
                                    Msg = "You pick up an " & Trim$(Item(ItemNum).Name) & "."
                                Else
                                    Msg = "You pick up a " & Trim$(Item(ItemNum).Name) & "."
                                End If
                            End If
                                
                            Call SetPlayerInvItemAmmo(Index, n, .Ammo)
            
                            ' Erase item from the map
                            Call ClearMapItem(i, MapNum)
            
                            Call SendInventoryUpdate(Index, n)
                            Call SpawnItemSlot(i, 0, 0, -1, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index))
                            Call PlayerMsg(Index, Msg, YELLOW)
                            Exit Sub
                        Else
                            Call PlayerMsg(Index, "Your inventory is full! Please make some room for this item.", BRIGHTRED)
                            Exit Sub
                        End If
                    End If
                End If
            End If
        End With
    Next i
End Sub

Sub PlayerMapDropItem(ByVal Index As Long, ByVal InvNum As Long, ByVal Amount As Long)
    Dim i As Long, ItemNum As Long
    
    ' Check for subscript out of range
    If IsPlaying(Index) = False Or InvNum <= 0 Or InvNum > GetPlayerMaxInv(Index) Then
        Exit Sub
    End If
    
    ItemNum = GetPlayerInvItemNum(Index, InvNum)
    
    If ItemNum > 0 And ItemNum <= MAX_ITEMS Then
        i = FindOpenMapItemSlot(GetPlayerMap(Index))
    
        If i <> 0 Then
            With MapItem(GetPlayerMap(Index), i)
                .Ammo = -1
                    
                ' Check to see if its any sort of ArmorSlot/WeaponSlot
                Select Case Item(ItemNum).Type
                    Case ITEM_TYPE_WEAPON
                        .Ammo = GetPlayerInvItemAmmo(Index, InvNum)
                    Case ITEM_TYPE_ARMOR, ITEM_TYPE_HELMET, ITEM_TYPE_SPECIALBADGE, ITEM_TYPE_LEGS, ITEM_TYPE_FLOWERBADGE, ITEM_TYPE_MUSHROOMBADGE
                        .Ammo = -1
                End Select
        
                .num = ItemNum
                .X = GetPlayerX(Index)
                .Y = GetPlayerY(Index)
        
                If ItemIsStackable(ItemNum) = True Then
                    ' Check if its more then they have and if so drop it all
                    If Amount >= GetPlayerInvItemValue(Index, InvNum) Then
                        .Value = GetPlayerInvItemValue(Index, InvNum)
                        Call SetPlayerInvItemNum(Index, InvNum, 0)
                        Call SetPlayerInvItemValue(Index, InvNum, 0)
                        Call SetPlayerInvItemAmmo(Index, InvNum, -1)
                    Else
                        .Value = Amount
                        Call SetPlayerInvItemValue(Index, InvNum, GetPlayerInvItemValue(Index, InvNum) - Amount)
                    End If
                Else
                    ' Not a stackable item
                    .Value = 0
        
                    Call SetPlayerInvItemNum(Index, InvNum, 0)
                    Call SetPlayerInvItemValue(Index, InvNum, 0)
                    Call SetPlayerInvItemAmmo(Index, InvNum, -1)
                End If
        
                ' Send inventory update
                Call SendInventoryUpdate(Index, InvNum)
                ' Spawn the item before we set the num or we'll get a different free map item slot
                Call SpawnItemSlot(i, .num, Amount, .Ammo, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index))
            End With
        Else
            Call PlayerMsg(Index, "There are already too many items on the ground!", BRIGHTRED)
        End If
    End If
End Sub

Sub SpawnNpc(ByVal MapNpcNum As Long, ByVal MapNum As Long)
    Dim packet As String
    Dim NpcNum As Long, i As Long, X As Long, Y As Long
    Dim Spawned As Boolean

    ' Check for subscript out of range
    If MapNpcNum < 1 Or MapNpcNum > MAX_MAP_NPCS Or MapNum < 1 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If
    
    NpcNum = Map(MapNum).NPC(MapNpcNum)

    If NpcNum > 0 Then
        MapNPC(MapNum, MapNpcNum).num = NpcNum
        MapNPC(MapNum, MapNpcNum).Target = 0

        MapNPC(MapNum, MapNpcNum).HP = GetNpcMaxHP(NpcNum)
        MapNPC(MapNum, MapNpcNum).MP = GetNpcMaxMP(NpcNum)
        MapNPC(MapNum, MapNpcNum).SP = GetNpcMaxSP(NpcNum)

        MapNPC(MapNum, MapNpcNum).Dir = Int(Rnd2 * 4)

        ' This means the admin wants to do a random spawn. [Mellowz]
        If Map(MapNum).SpawnX(MapNpcNum) = 0 Or Map(MapNum).SpawnY(MapNpcNum) = 0 Then
            For i = 1 To 100
                X = Int(Rnd2 * MAX_MAPX)
                Y = Int(Rnd2 * MAX_MAPY)
    
                If Map(MapNum).Tile(X, Y).Type = TILE_TYPE_WALKABLE Then
                    MapNPC(MapNum, MapNpcNum).X = X
                    MapNPC(MapNum, MapNpcNum).Y = Y
                    Spawned = True
                    Exit For
                End If
            Next i

            ' Didn't spawn, so now we'll just try to find a free tile
            If Not Spawned Then
                For Y = 0 To MAX_MAPY
                    For X = 0 To MAX_MAPX
                        If Map(MapNum).Tile(X, Y).Type = TILE_TYPE_WALKABLE Then
                            MapNPC(MapNum, MapNpcNum).X = X
                            MapNPC(MapNum, MapNpcNum).Y = Y
                            Spawned = True
                        End If
                    Next X
                Next Y
            End If
        Else
            ' We subtract one because Rand is ListIndex 0. [Mellowz]
            MapNPC(MapNum, MapNpcNum).X = Map(MapNum).SpawnX(MapNpcNum) - 1
            MapNPC(MapNum, MapNpcNum).Y = Map(MapNum).SpawnY(MapNpcNum) - 1
            Spawned = True
        End If

        ' If we suceeded in spawning then send it to everyone
        If Spawned Then
            packet = SPackets.Sspawnnpc & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNPC(MapNum, MapNpcNum).num & SEP_CHAR & MapNPC(MapNum, MapNpcNum).X & SEP_CHAR & MapNPC(MapNum, MapNpcNum).Y & SEP_CHAR & MapNPC(MapNum, MapNpcNum).Dir & SEP_CHAR & NPC(MapNPC(MapNum, MapNpcNum).num).Big & END_CHAR
            Call SendDataToMap(MapNum, packet)
        End If
    End If
End Sub

Sub SpawnMapNpcs(ByVal MapNum As Long)
    Dim i As Long

    For i = 1 To MAX_MAP_NPCS
        If Map(MapNum).NPC(i) > 0 Then
            Call SpawnNpc(i, MapNum)
        End If
    Next i
End Sub

Sub SpawnAllMapNpcs()
    Dim i As Long

    For i = 1 To MAX_MAPS
        If PlayersOnMap(i) = YES Then
            Call SpawnMapNpcs(i)
        End If
    Next i
End Sub

Function CanAttackPlayer(ByVal Attacker As Long, ByVal Victim As Long) As Boolean
    Dim i As Byte
    
    ' Check for subscript out of range
    If IsPlaying(Attacker) = False Or IsPlaying(Victim) = False Then
        Exit Function
    End If

    ' Make sure they have more than 0 hp
    If GetPlayerHP(Victim) <= 0 Then
        Exit Function
    End If

    ' Make sure we dont attack the player if they are switching maps
    If Player(Victim).GettingMap = YES Then
        Exit Function
    End If
    
    ' Make sure the player isn't in a battle
    If GetPlayerInBattle(Victim) = True Then
        Exit Function
    End If
    
    ' Make sure they're on the same map
    If GetPlayerMap(Attacker) <> GetPlayerMap(Victim) Then
        Exit Function
    End If

    ' Make sure the player hasn't just attacked
    If GetTickCount < Player(Attacker).AttackTimer + GetPlayerAttackSpeed(Attacker) Then
        Exit Function
    End If
    
    ' Check if they're at same coordinates
    Select Case GetPlayerDir(Attacker)
        Case DIR_UP
            If (GetPlayerY(Victim) + 1 = GetPlayerY(Attacker)) And (GetPlayerX(Victim) = GetPlayerX(Attacker)) Then
                GoTo DetermineAction
            End If
        Case DIR_DOWN
            If (GetPlayerY(Victim) - 1 = GetPlayerY(Attacker)) And (GetPlayerX(Victim) = GetPlayerX(Attacker)) Then
                GoTo DetermineAction
            End If
        Case DIR_LEFT
            If (GetPlayerY(Victim) = GetPlayerY(Attacker)) And (GetPlayerX(Victim) + 1 = GetPlayerX(Attacker)) Then
                GoTo DetermineAction
            End If
        Case DIR_RIGHT
            If (GetPlayerY(Victim) = GetPlayerY(Attacker)) And (GetPlayerX(Victim) - 1 = GetPlayerX(Attacker)) Then
                GoTo DetermineAction
            End If
    End Select
    
    Exit Function
DetermineAction:
    ' If they are in STS or Hide n' Sneak then they can attack each other
    If GetPlayerMap(Attacker) = 33 Or GetPlayerMap(Attacker) = 271 Or GetPlayerMap(Attacker) = 272 Or GetPlayerMap(Attacker) = 273 Then
        CanAttackPlayer = True
        Exit Function
    End If
    ' If they are on an arena tile, then they can be attacked
    If GetPlayerMap(Attacker) = 218 And Map(GetPlayerMap(Attacker)).Tile(GetPlayerX(Attacker), GetPlayerY(Attacker)).Type = TILE_TYPE_ARENA And Map(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).Type = TILE_TYPE_ARENA Then
        CanAttackPlayer = True
        Exit Function
    End If
    ' Check if map is attackable
    If (Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_SAFE Or Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_MINIGAME) And GetPlayerPK(Victim) = NO Then
        Call PlayerMsg(Attacker, "This is not a PvP area!", BRIGHTRED)
        Exit Function
    End If
    ' Check if they are in a guild and if they are a pker
    If GetPlayerGuild(Attacker) <> vbNullString And GetPlayerGuild(Victim) <> vbNullString Then
        If GetPlayerGuild(Attacker) = GetPlayerGuild(Victim) Then
            Call PlayerMsg(Attacker, GetPlayerName(Victim) & " is in the same Group as you, so you cannot attack him/her!", BRIGHTRED)
            Exit Function
        End If
    End If
    CanAttackPlayer = True
End Function

Function CanAttackNpc(ByVal Attacker As Long, ByVal MapNpcNum As Long) As Boolean
    Dim i As Byte
    Dim MapNum As Long, NpcNum As Long

    ' Check for subscript out of range
    If Not IsPlaying(Attacker) Then
        Exit Function
    End If

    ' Check for subscript out of range
    If MapNpcNum < 1 Or MapNpcNum > MAX_MAP_NPCS Then
        Exit Function
    End If
    
    MapNum = GetPlayerMap(Attacker)
    NpcNum = MapNPC(MapNum, MapNpcNum).num
    
    ' Check for subscript out of range and the Rockade NPC
    If NpcNum = 0 Or NpcNum = 195 Then
        Exit Function
    End If
    
    ' Make sure the npc isn't in battle with another player
    If MapNPC(MapNum, MapNpcNum).InBattle = True And MapNPC(MapNum, MapNpcNum).Target <> Attacker Then
        Exit Function
    End If
    
    ' Make sure the player hasn't just attacked
    If GetTickCount < Player(Attacker).AttackTimer + GetPlayerAttackSpeed(Attacker) Then
        Exit Function
    End If
    
    ' Make sure the npc isn't already dead
    If MapNPC(MapNum, MapNpcNum).HP <= 0 Then
        Exit Function
    End If

    ' Check to see if the player is within the appropriate range
    Select Case GetPlayerDir(Attacker)
        Case DIR_UP
            If (MapNPC(MapNum, MapNpcNum).Y + 1 = GetPlayerY(Attacker)) And (MapNPC(MapNum, MapNpcNum).X = GetPlayerX(Attacker)) Then
                GoTo DetermineAction
            End If
        Case DIR_DOWN
            If (MapNPC(MapNum, MapNpcNum).Y - 1 = GetPlayerY(Attacker)) And (MapNPC(MapNum, MapNpcNum).X = GetPlayerX(Attacker)) Then
                GoTo DetermineAction
            End If
        Case DIR_LEFT
            If (MapNPC(MapNum, MapNpcNum).Y = GetPlayerY(Attacker)) And (MapNPC(MapNum, MapNpcNum).X + 1 = GetPlayerX(Attacker)) Then
                GoTo DetermineAction
            End If
        Case DIR_RIGHT
            If (MapNPC(MapNum, MapNpcNum).Y = GetPlayerY(Attacker)) And (MapNPC(MapNum, MapNpcNum).X - 1 = GetPlayerX(Attacker)) Then
                GoTo DetermineAction
            End If
    End Select

    Exit Function

DetermineAction:
        If NPC(NpcNum).Behavior = NPC_BEHAVIOR_SCRIPTED Then
            Call ScriptedNPC(Attacker, NpcNum, NPC(NpcNum).SpawnSecs)
            Exit Function
        ElseIf NPC(NpcNum).Behavior = NPC_BEHAVIOR_FRIENDLY Then
            Call SendNpcTalkTo(Attacker, NpcNum, NPC(NpcNum).AttackSay, NPC(NpcNum).AttackSay2)
            Exit Function
        End If
        CanAttackNpc = True
End Function

Function CanNpcAttackPlayer(ByVal MapNpcNum As Long, ByVal Index As Long) As Boolean
    Dim MapNum As Long, NpcNum As Long
    
    ' Prevent subscript out of range
    If IsPlaying(Index) = False Then
        Exit Function
    End If

    ' Make sure the map npc number is valid
    If MapNpcNum < 1 Or MapNpcNum > MAX_MAP_NPCS Then
        Exit Function
    End If
    
    MapNum = GetPlayerMap(Index)
    NpcNum = MapNPC(MapNum, MapNpcNum).num
    
    ' Prevent subscript out of range
    If NpcNum = 0 Then
        Exit Function
    End If

    ' Make sure that the NPC isn't already dead
    If MapNPC(MapNum, MapNpcNum).HP < 1 Then
        Exit Function
    End If

    ' Make sure the NPCs don't attack more than once every second
    If GetTickCount < MapNPC(MapNum, MapNpcNum).AttackTimer + 1000 Then
        Exit Function
    End If

    ' Make sure we don't attack a player switching maps
    If Player(Index).GettingMap = YES Then
        Exit Function
    End If
    
    ' Make sure the npc cannot attack a player in battle with another npc
    If GetPlayerInBattle(Index) = True And Player(Index).TargetNPC <> MapNpcNum Then
        If MapNPC(MapNum, MapNpcNum).Target = Index Then
            MapNPC(MapNum, MapNpcNum).Target = 0
        End If
        
        Exit Function
    End If
    
    ' Prevent the npc from harming a player during his/her recovery period
    If GetTickCount < GetPlayerRecoverTime(Index) + 2000 And GetPlayerTurnBased(Index) = True Then
        Exit Function
    End If

    MapNPC(MapNum, MapNpcNum).AttackTimer = GetTickCount
    
    ' Check if the npc is at the appropriate attack range
    If (GetPlayerY(Index) + 1 = MapNPC(MapNum, MapNpcNum).Y) And (GetPlayerX(Index) = MapNPC(MapNum, MapNpcNum).X) Then
        CanNpcAttackPlayer = True
    Else
        If (GetPlayerY(Index) - 1 = MapNPC(MapNum, MapNpcNum).Y) And (GetPlayerX(Index) = MapNPC(MapNum, MapNpcNum).X) Then
            CanNpcAttackPlayer = True
        Else
            If (GetPlayerY(Index) = MapNPC(MapNum, MapNpcNum).Y) And (GetPlayerX(Index) + 1 = MapNPC(MapNum, MapNpcNum).X) Then
                CanNpcAttackPlayer = True
            Else
                If (GetPlayerY(Index) = MapNPC(MapNum, MapNpcNum).Y) And (GetPlayerX(Index) - 1 = MapNPC(MapNum, MapNpcNum).X) Then
                    CanNpcAttackPlayer = True
                End If
            End If
        End If
    End If
End Function

Sub STSDeath(ByVal Attacker As Long, ByVal Victim As Long)
    Call SetPlayerHP(Victim, 0)

    ' Set target to 0
    Player(Attacker).Target = 0
    Player(Attacker).TargetType = 0
            
    ' Restore vitals
    Call SetPlayerHP(Victim, GetPlayerMaxHP(Victim))
    Call SetPlayerMP(Victim, GetPlayerMaxMP(Victim))
    Call SetPlayerSP(Victim, GetPlayerMaxSP(Victim))
    Call SendHP(Victim)
    Call SendMP(Victim)
    Call SendSP(Victim)
    
    Call SendAttackSound(GetPlayerMap(Victim))
    Call HandleSTSDeath(Attacker, Victim)
End Sub

Sub HandleSTSDeath(ByVal Attacker As Long, ByVal Victim As Long)
    Dim FlagStatus As String, FilePath As String
    Dim i As Long, R As Long, Q As Long
    Dim Team As Byte
    
    FilePath = STSPath
    Team = GetPlayerTeam(Victim, STSPath, STSMaxPlayers)
    
    ' Determine the team the victim is on
    If Team = 1 Then
        Call SetPlayerX(Victim, 16)
        Call SetPlayerY(Victim, 26)
        Call SendPlayerXY(Victim)
        FlagStatus = GetVar(FilePath, "Flag", "Blue")
    ElseIf Team = 0 Then
        Call SetPlayerX(Victim, 15)
        Call SetPlayerY(Victim, 4)
        Call SendPlayerXY(Victim)
        FlagStatus = GetVar(FilePath, "Flag", "Red")
    End If
        
    If GetPlayerPK(Victim) = 1 And FlagStatus = "HasFlag" Then
        Call SetPlayerPK(Victim, NO)
        Call SendPlayerData(Victim)
        Call SendPlayerXY(Victim)
        Call SendSoundToMap(GetPlayerMap(Victim), "smw_reservedrop.wav")
        
        ' Returns flag to Red team when Blue player dies
        If Team = 1 Then
            Call PutVar(FilePath, "Flag", "Blue", "NoFlag")
        ' Returns flag to Blue team when Red player dies
        ElseIf Team = 0 Then
            Call PutVar(FilePath, "Flag", "Red", "NoFlag")
        End If
        
        For i = 1 To 4
            FlagStatus = CStr(i)
            R = FindPlayer(GetVar(FilePath, "Red", FlagStatus))
            Q = FindPlayer(GetVar(FilePath, "Blue", FlagStatus))
                
            If Q <> Victim Then
                Call PlayerMsg(Q, GetPlayerName(Victim) & " has lost the enemy Shroom!", WHITE)
            Else
                Call PlayerMsg(Victim, "You lost the enemy Shroom!", WHITE)
            End If
            If R <> Attacker Then
                Call PlayerMsg(R, "Your Shroom has been returned!", WHITE)
            Else
                Call PlayerMsg(Attacker, "You returned your Shroom!", WHITE)
            End If
        Next i
    End If
End Sub

Sub DodgeBillDeath(ByVal Attacker As Long, ByVal Victim As Long)
    Dim FilePath As String, TeamName As String, PlayerNum As String
    Dim TeamNumber As Byte, Outs As Byte, Team As Byte, Point As Byte, BulletBillCount As Byte
    Dim i As Long, R As Long, Q As Long, MapItemNum As Long
    
    FilePath = DodgeBillPath
    
    ' Find out the team color
    Team = GetPlayerTeam(Victim, DodgeBillPath, DodgeBillMaxPlayers)
    
    If Team = 0 Then
        TeamName = "Red"
    ElseIf Team = 1 Then
        TeamName = "Blue"
    End If
    
    ' Find out how many players are on the victim's team and how many outs the team currently has
    TeamNumber = CByte(GetVar(FilePath, "Team", TeamName))
    Outs = CByte(GetVar(FilePath, "Outs", TeamName))
    
    ' Set target to 0
    Player(Attacker).Target = 0
    Player(Attacker).TargetType = 0
    
    ' Send the sound
    Call SendAttackSound(GetPlayerMap(Victim))
    
    ' Spawn the item
    MapItemNum = FindOpenMapItemSlot(188)
    Call SpawnItemSlot(MapItemNum, 186, 1, 1, 188, GetPlayerX(Victim), GetPlayerY(Victim))
    
    ' Check for any bullet bills and spawn them when the player gets out
    BulletBillCount = FindOpenInvSlot(Victim, 186)
    
    If BulletBillCount > 0 Then
        BulletBillCount = GetPlayerInvItemValue(Victim, BulletBillCount)
        
        If BulletBillCount > 0 Then
            MapItemNum = FindOpenMapItemSlot(188)
            Call TakeItem(Victim, 186, BulletBillCount)
            Call SpawnItemSlot(MapItemNum, 186, BulletBillCount, 1, 188, GetPlayerX(Victim), GetPlayerY(Victim))
        End If
    End If
    
    ' Send message
    Call MapMsg(188, GetPlayerName(Attacker) & " has gotten " & GetPlayerName(Victim) & " out!", WHITE)
    
    ' Warp to jail
    Call SetPlayerX(Victim, 15)
    Call SetPlayerY(Victim, 5 + (Team * 21))
    Call SendPlayerXY(Victim)
    
    ' Add number of outs to total
    Call PutVar(FilePath, "Outs", TeamName, (Outs + 1))
    
    ' Check if all players on the team are out
    If (Outs + 1) >= TeamNumber Then
        ' Find out opposing team
        If TeamName = "Red" Then
            TeamName = "Blue"
        ElseIf TeamName = "Blue" Then
            TeamName = "Red"
        End If
        
        ' Get current number of points
        Point = CByte(GetVar(FilePath, "Points", TeamName))
            
        ' Add a point to the opposing team's score
        Call PutVar(FilePath, "Points", TeamName, (Point + 1))
        
        ' Put all players back onto the field and state which team won the round
        Call MapMsg(GetPlayerMap(Victim), "The " & TeamName & " team has won the round!", YELLOW)
        
        ' Warp players
        For i = 1 To 5
            PlayerNum = CStr(i)
                    
            R = FindPlayer(GetVar(FilePath, "Blue", PlayerNum))
            Q = FindPlayer(GetVar(FilePath, "Red", PlayerNum))
                            
            ' Warp players on the blue team
            If IsPlaying(R) Then
                Call SetPlayerX(R, (12 + i))
                Call SetPlayerY(R, 20)
                Call SendPlayerXY(R)
            End If
                    
            ' Warp players on the blue team
            If IsPlaying(Q) Then
                Call SetPlayerX(Q, (12 + i))
                Call SetPlayerY(Q, 10)
                Call SendPlayerXY(Q)
            End If
        Next i
                
        ' Reset outs
        Call PutVar(FilePath, "Outs", "Blue", "0")
        Call PutVar(FilePath, "Outs", "Red", "0")
    End If
End Sub

Sub HideNSneakOut(ByVal Attacker As Long, ByVal Victim As Long)
    Dim PlayerNum As String, HiderIndex As String
    Dim i As Byte, NumberHidersFound As Byte
    Dim NumHiderIndex As Long
    
    ' Loop through all seekers
    For i = 1 To MaxSeekers
        PlayerNum = CStr(i)
        
        ' Make sure the Attacker is a Seeker
        If FindPlayer(GetVar(HideNSneakPath, "Seekers", PlayerNum)) = Attacker Then
            HiderIndex = CStr(FindHiderIndex(GetPlayerName(Victim)))
            
            ' Check if the player is already out
            If GetVar(HideNSneakPath, "PlayersOut", HiderIndex) <> "Out" Then
                NumHiderIndex = CLng(HiderIndex)
            
                ' Find out how many players are currently out
                NumberHidersFound = CByte(GetVar(HideNSneakPath, "PlayersFound", "PlayersFound"))
                
                ' Record that the player is now out
                Call PutVar(HideNSneakPath, "PlayersOut", HiderIndex, "Out")
                Call PutVar(HideNSneakPath, "PlayersFound", "PlayersFound", (NumberHidersFound + 1))
                
                ' Notify all players that the Seeker got the Hider out
                Call SendHideNSneakMsg(GetPlayerName(Attacker) & " has gotten " & GetPlayerName(Victim) & " out!", YELLOW, True)
                
                ' Warp the player away to the "Out Players" area
                Call PlayerWarp(Victim, 271, (12 + NumHiderIndex), 26)
                
                ' Check if the game can be ended
                If (NumberHidersFound + 1) >= CByte(GetVar(HideNSneakPath, "Team", "Hiders")) Then
                    ' Set the TimeLeft to 0
                    Call PutVar(HideNSneakPath, "GameTime", "TimeLeft", "0")
                    
                    ' Call the timer to end the game immediately
                    Call HideNSneakPlayTime(CLng(GetVar(HideNSneakPath, "TimerIndex", "TimerIndex")))
                End If
            Else
                ' The player is already out, so notify the Attacker of this
                Call PlayerMsg(Attacker, GetPlayerName(Victim) & " is already out!", WHITE)
            End If
                
            Exit Sub
        End If
    Next
End Sub

Public Function FindHiderIndex(ByVal PlayerName As String) As Long
    Dim i As Byte
    
    For i = 1 To MaxHiders
        If GetVar(HideNSneakPath, "Hiders", CStr(i)) = PlayerName Then
            FindHiderIndex = i
            
            Exit Function
        End If
    Next
    
    FindHiderIndex = 0
End Function

Sub AttackPlayer(ByVal Attacker As Long, ByVal Victim As Long, ByVal Damage As Long)
    Dim Exp As Long
    Dim Sound As Byte
    
    ' Make sure the attacker or the victim is a valid index
    If IsPlaying(Attacker) = False Or IsPlaying(Victim) = False Then
        Exit Sub
    End If

    ' If damage is below one, exit this sub routine.
    If Damage < 1 Then
        Exit Sub
    End If
    
    Sound = 0
    
    If Damage >= GetPlayerHP(Victim) Then
        ' See if the player can be revived by the Close Call badge
        If CheckCloseCall(Victim) = True Then
            Player(Attacker).AttackTimer = GetTickCount
            
            Exit Sub
        End If
    
        ' Revives player if the player has a Life Shroom
        If HasItem(Victim, 43) = 1 Then
            Call LifeShroom(Victim)
            
            Player(Attacker).AttackTimer = GetTickCount
            
            Exit Sub
        End If
        
        Call SetPlayerHP(Victim, 0)
        Call OnPVPDeath(Attacker, Victim)
        
        If GetPlayerMap(Attacker) <> 218 Then
            If Map(GetPlayerMap(Victim)).Moral <> MAP_MORAL_NO_PENALTY Then
                ' Calculate exp to give attacker
                Exp = (GetPlayerExp(Victim) \ 20)
    
                ' Make sure we dont get fewer than 0 experience points
                If Exp < 0 Then
                    Exp = 0
                End If
    
                If Exp = 0 Then
                    Call SetPlayerExp(Victim, 0)
                    Call PlayerMsg(Victim, "Oh no! You've died!", BRIGHTRED)
                    Call PlayerMsg(Attacker, GetPlayerName(Victim) & " doesn't have any experience points for you to gain!", WHITE)
                Else
                    Call SetPlayerExp(Victim, GetPlayerExp(Victim) - Exp)
                    Call PlayerMsg(Victim, "Oh no! You've died! You lost " & Exp & " experience points.", BRIGHTRED)
                    Call SetPlayerExp(Attacker, GetPlayerExp(Attacker) + Exp)
                    Call PlayerMsg(Attacker, "You gained " & Exp & " experience points for killing " & GetPlayerName(Victim) & "!", YELLOW)
                End If
                        
                Call SendEXP(Victim)
                Call SendEXP(Attacker)
            End If
        End If

        ' Warp player away
        If GetPlayerMap(Victim) <> 218 Then
            Call OnDeath(Victim)
        Else
            Call OnArenaDeath(Attacker, Victim)
        End If

        ' Restore vitals
        Call SetPlayerHP(Victim, GetPlayerMaxHP(Victim))
        Call SetPlayerMP(Victim, GetPlayerMaxMP(Victim))
        Call SetPlayerSP(Victim, GetPlayerMaxSP(Victim))
        Call SendHP(Victim)
        Call SendMP(Victim)
        Call SendSP(Victim)

        ' Check for a level up
        Call CheckPlayerLevelUp(Attacker)

        ' Check if target is player who died and if so set target to 0
        If Player(Attacker).TargetType = TARGET_TYPE_PLAYER And Player(Attacker).Target = Victim Then
            Player(Attacker).Target = 0
            Player(Attacker).TargetType = 0
        End If

        If GetPlayerMap(Attacker) <> 218 Then
            If GetPlayerPK(Victim) = NO Then
                If GetPlayerPK(Attacker) = NO Then
                    Call SetPlayerPK(Attacker, YES)
                    Call SendPlayerData(Attacker)
                    Call SendPlayerXY(Attacker)
                    Call GlobalMsg(GetPlayerName(Attacker) & " is now a player killer! Go get " & GetPlayerName(Attacker) & "!", BRIGHTRED)
                End If
            Else
                Call SetPlayerPK(Victim, NO)
                Call SendPlayerData(Victim)
                Call SendPlayerXY(Victim)
                Call GlobalMsg(GetPlayerName(Victim) & " now understands what it's like to be killed by another player!", BRIGHTRED)
            End If
        End If
    Else
        ' Check which sound to send
        If GetPlayerHP(Victim) > 5 And (GetPlayerHP(Victim) - Damage) <= 5 Then
            Sound = 1
        End If
        Call SetPlayerHP(Victim, GetPlayerHP(Victim) - Damage)
        Call SendHP(Victim)
    End If
    
    Player(Attacker).AttackTimer = GetTickCount

    If Sound = 0 Then
        Call SendDataToMap(GetPlayerMap(Victim), SPackets.Ssound & SEP_CHAR & "pain" & END_CHAR)
    End If
End Sub

Sub NpcAttackPlayer(ByVal MapNpcNum As Long, ByVal Victim As Long, ByVal Damage As Long)
    Dim Name As String
    Dim Exp As Long, MapNum As Long
    Dim Sound As Byte
    
    If MapNpcNum < 1 Or MapNpcNum > MAX_MAP_NPCS Then
        Exit Sub
    End If
    
    If IsPlaying(Victim) = False Then
        Exit Sub
    End If
    
    Damage = DamageUpDamageDown(Victim, Damage)
    
    If Damage < 1 Then
        Damage = 1
    End If

    If MapNPC(GetPlayerMap(Victim), MapNpcNum).num <= 0 Then
        Exit Sub
    End If

    ' Send this packet so they can see the npc attacking
    Call SendDataToMap(GetPlayerMap(Victim), SPackets.Snpcattack & SEP_CHAR & MapNpcNum & END_CHAR)

    MapNum = GetPlayerMap(Victim)

    Name = Trim$(NPC(MapNPC(MapNum, MapNpcNum).num).Name)
    
    ' Starts a turn-based battle if one hasn't been started yet and continues the battle if one is in progress
    If Map(MapNum).Moral <> MAP_MORAL_MINIGAME Then
        If IsInPoisonCave(Victim) Then
            If (MapNPC(MapNum, MapNpcNum).InBattle = False And GetPlayerInBattle(Victim) = False) Then
                Call NpcFirstStrike(Victim, MapNpcNum)
                Exit Sub
            End If
        Else
            ' Check if we can start the battle, and if we can, start it
            If GetPlayerTurnBased(Victim) = True Then
                If (MapNPC(MapNum, MapNpcNum).InBattle = False And GetPlayerInBattle(Victim) = False) Then
                    Call NpcFirstStrike(Victim, MapNpcNum)
                    Exit Sub
                End If
            End If
        End If
        
        ' Continue the turn-based battle that is already in progress
        If MapNPC(MapNum, MapNpcNum).InBattle = True And GetPlayerInBattle(Victim) = True Then
            Call TurnBasedNpcAttackPlayer(Victim, MapNum, MapNpcNum, Damage)
            Exit Sub
        End If
    End If

    Sound = 0
    
    If Damage >= GetPlayerHP(Victim) Then
        ' See if the player can be revived by the Close Call badge
        If CheckCloseCall(Victim) = True Then
            MapNPC(MapNum, MapNpcNum).Target = 0
            Call SendDataTo(Victim, SPackets.Sblitnpcdmg & SEP_CHAR & Damage & END_CHAR)
            
            Exit Sub
        End If
    
        ' Allows a Life Shroom to revive a player
        If HasItem(Victim, 43) = 1 Then
            Call LifeShroom(Victim)
            
            MapNPC(MapNum, MapNpcNum).Target = 0
            Call SendDataTo(Victim, SPackets.Sblitnpcdmg & SEP_CHAR & Damage & END_CHAR)
            
            Exit Sub
        End If
       
        If FindNpcVowels(MapNPC(MapNum, MapNpcNum).num) = True Then
            Call PlayerMsg(Victim, "You have been killed by an " & Name, BRIGHTRED)
        Else
            Call PlayerMsg(Victim, "You have been killed by a " & Name, BRIGHTRED)
        End If
        
        If Map(GetPlayerMap(Victim)).Moral <> MAP_MORAL_NO_PENALTY Then
            ' Calculate exp to give attacker
            Exp = (GetPlayerExp(Victim) \ 6)

            ' Make sure we don't get fewer than 0 experience points
            If Exp < 0 Then
                Exp = 0
            End If
          ' Prevents you from losing Exp during the Daily Event
            If Map(GetPlayerMap(Victim)).Moral <> MAP_MORAL_MINIGAME Then
                Call SetPlayerExp(Victim, GetPlayerExp(Victim) - Exp)
                Call PlayerMsg(Victim, "Oh no! You've died! You lost " & Exp & " experience points.", BRIGHTRED)
            Else
                Call PlayerMsg(Victim, "Oh no! You've died!", BRIGHTRED)
                Call GiveItem(Victim, 54, 1)
            End If
            Call SendEXP(Victim)
        End If

        ' Warp player away
        Call OnDeath(Victim)
            
        ' Restore vitals
        Call SetPlayerHP(Victim, GetPlayerMaxHP(Victim))
        Call SetPlayerMP(Victim, GetPlayerMaxMP(Victim))
        Call SetPlayerSP(Victim, GetPlayerMaxSP(Victim))
        Call SendHP(Victim)
        Call SendMP(Victim)
        Call SendSP(Victim)

        ' Set NPC target to 0
        MapNPC(MapNum, MapNpcNum).Target = 0

        ' If the player the attacker killed was a pk then take it away
        If GetPlayerPK(Victim) = YES Then
            Call SetPlayerPK(Victim, NO)
            Call SendPlayerData(Victim)
            Call SendPlayerXY(Victim)
        End If
    Else
        If GetPlayerHP(Victim) > 5 And (GetPlayerHP(Victim) - Damage) <= 5 Then
            Sound = 1
        End If
        
        Call SetPlayerHP(Victim, GetPlayerHP(Victim) - Damage)
        Call SendHP(Victim)
    End If
    
    Call SendDataTo(Victim, SPackets.Sblitnpcdmg & SEP_CHAR & Damage & END_CHAR)
    
    If Sound = 0 Then
        Call SendDataToMap(GetPlayerMap(Victim), SPackets.Ssound & SEP_CHAR & "pain" & END_CHAR)
    End If
End Sub

Sub AttackNpc(ByVal Attacker As Long, ByVal MapNpcNum As Long, ByVal Damage As Long)
    Dim MapNum As Long, NpcNum As Long
    
    ' Check for subscript out of range
    If IsPlaying(Attacker) = False Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Then
        Exit Sub
    End If
    
    MapNum = GetPlayerMap(Attacker)
    NpcNum = MapNPC(MapNum, MapNpcNum).num

    ' Different provisions for turn-based battles
    If Map(MapNum).Moral <> MAP_MORAL_MINIGAME Then
        If GetPlayerInBattle(Attacker) = True And MapNPC(MapNum, MapNpcNum).InBattle = True Then
            Call TurnBasedAttackNpc(Attacker, MapNpcNum, Damage)
            Exit Sub
        End If
        
        ' Starts turn based battle if it hasn't been started yet
        If IsInPoisonCave(Attacker) Then
            If NpcNum <> 195 Then
                If MapNPC(MapNum, MapNpcNum).InBattle = False And GetPlayerInBattle(Attacker) = False Then
                    Call PlayerFirstStrike(Attacker, MapNpcNum)
                    Exit Sub
                End If
            End If
        Else
            If GetPlayerTurnBased(Attacker) = True Then
                If NpcNum <> 195 Then
                    If MapNPC(MapNum, MapNpcNum).InBattle = False And GetPlayerInBattle(Attacker) = False Then
                        Call PlayerFirstStrike(Attacker, MapNpcNum)
                        Exit Sub
                    End If
                End If
            End If
        End If
    End If

    If Damage >= MapNPC(MapNum, MapNpcNum).HP Then
        Dim Add As String, Msg As String
        Dim i As Long, n As Long, Exp As Long, Index As Long, ItemNum As Long, ItemValue As Long, PartyNum As Long
        Dim CoinBadge As Boolean
    
        ' Since the player defeated the Npc, set his/her target to 0
        Player(Attacker).TargetNPC = 0

        Call OnNPCDeath(Attacker, MapNum, NpcNum, MapNpcNum)

        Add = 0
        
        For i = 1 To 7
            If GetPlayerEquipSlotNum(Attacker, i) > 0 Then
                Add = Add + Item(GetPlayerEquipSlotNum(Attacker, i)).AddEXP
            End If
        Next i

        If Add > 0 Then
            If Add < 100 Then
                If Add < 10 Then
                    Add = 0 & ".0" & Right$(Add, 2)
                Else
                    Add = 0 & "." & Right$(Add, 2)
                End If
            Else
                Add = Mid$(Add, 1, 1) & "." & Right$(Add, 2)
            End If
        ElseIf Add < 0 Then
            If Val(Add) > -100 Then
                If Val(Add) > -10 Then
                    Add = "-" & "." & 0 & Left$(-1 * Val(Add), 2)
                Else
                    Add = "-" & "." & Right$(-1 * Val(Add), 2)
                End If
            Else
                Add = "-" & Mid$(-1 * Val(Add), 1, 1) & "." & Right$(-1 * Val(Add), 2)
            End If
        End If
        
        ' Calculate exp to give attacker
        If Add = 0 Then
            Exp = NPC(NpcNum).Exp
        Else
            Exp = NPC(NpcNum).Exp + (NPC(NpcNum).Exp * Val(Add))
        End If

        ' Make sure we dont get fewer than 0 experience points
        If Exp < 0 Then
            Exp = 0
        End If
        
        PartyNum = GetPlayerPartyNum(Attacker)
       
        ' Doesn't display or give exp if player is in the Whack-A-Monty minigame
        If Map(MapNum).Moral <> MAP_MORAL_MINIGAME Then
            ' Check if the player is in a party and distribute party experience among members if applicable
            If PartyNum <= 0 Or PartyNum > MAX_PLAYERS Or GetPartyShareCount(Attacker) <= 1 Then
                Call SetPlayerExp(Attacker, GetPlayerExp(Attacker) + Exp)
                
                ' Change message depending on how many experience points the player earned
                Msg = "You gained " & Exp & " experience "
                If Exp = 1 Then
                    Msg = Msg & "point."
                Else
                    Msg = Msg & "points."
                End If
                
                Call BattleMsg(Attacker, Msg, BRIGHTBLUE, 0)
            Else
                Exp = Int((Exp * (1 / (MAX_PARTY_MEMBERS - (MAX_PARTY_MEMBERS - GetPartyMembers(PartyNum))))))
                
                ' Make sure we dont get fewer than 0 experience points
                If Exp < 0 Then
                    Exp = 0
                End If
                
                ' Change message depending on how many experience points the player earned
                Msg = "You gained " & Exp & " party experience "
                If Exp = 1 Then
                    Msg = Msg & "point."
                Else
                    Msg = Msg & "points."
                End If
                
                ' Distribute experience points to all of the party members
                For n = 1 To MAX_PARTY_MEMBERS
                    Index = GetPartyMember(PartyNum, n)
                    If Index > 0 Then
                        If GetPlayerPartyShare(Index) = True Then
                            Call SetPlayerExp(Index, GetPlayerExp(Index) + Exp)
                            Call BattleMsg(Index, Msg, BRIGHTBLUE, 0)
                        End If
                    End If
                Next n
            End If
        End If
       
        CoinBadge = False
        
        If GetPlayerEquipSlotNum(Attacker, 4) = 142 Then
            CoinBadge = True
        End If
        
        ' Handle any item drops
        For i = 1 To MAX_NPC_DROPS
            ItemNum = NPC(NpcNum).ItemNPC(i).ItemNum
            
            If ItemNum > 0 Then
                n = Int(Rnd2 * NPC(NpcNum).ItemNPC(i).Chance) + 1
                
                If n = 1 Then
                    ItemValue = NPC(NpcNum).ItemNPC(i).ItemValue
                    
                    If ItemNum = 1 Or ItemNum = 271 Then
                        ' Gives 2x more coins with the Payoff Badge
                        If CoinBadge = True Then
                            ItemValue = (ItemValue * 2)
                        End If
                    End If
                    
                    Call SpawnItem(ItemNum, ItemValue, MapNum, MapNPC(MapNum, MapNpcNum).X, MapNPC(MapNum, MapNpcNum).Y)
                End If
            End If
        Next i

        ' Now set HP to 0 so we know to actually kill them in the server loop (this prevents subscript out of range)
        MapNPC(MapNum, MapNpcNum).num = 0
        MapNPC(MapNum, MapNpcNum).SpawnWait = GetTickCount
        MapNPC(MapNum, MapNpcNum).HP = 0
        Call SendNpcDead(MapNum, MapNpcNum)

        ' Check for level up
        Call CheckPlayerLevelUp(Attacker)

        ' Check for level up party member
        If PartyNum > 0 Then
            For n = 1 To MAX_PARTY_MEMBERS
                Index = GetPartyMember(PartyNum, n)
                If Index > 0 Then
                    Call CheckPlayerLevelUp(Index)
                End If
            Next n
        End If

        ' Check if target is npc that died and if so set target to 0
        If Player(Attacker).TargetType = TARGET_TYPE_NPC And Player(Attacker).Target = MapNpcNum Then
            Player(Attacker).Target = 0
            Player(Attacker).TargetType = 0
        End If
        
        ' Check for the Heart Power badge
        If GetPlayerEquipSlotNum(Attacker, 7) = 317 Then
            Dim RandNum As Long
            
            RandNum = Rand(1, 20)
            
            ' Give it a 15% chance of activating
            If RandNum <= 3 Then
                Dim HPHeal As Long
                
                ' Heal the player for half the player's Stache
                HPHeal = (GetPlayerStache(Attacker) \ 2)
                
                Call SetPlayerHP(Attacker, GetPlayerHP(Attacker) + HPHeal)
                Call SendHP(Attacker)
                
                Call PlayerMsg(Attacker, "You healed " & HPHeal & " HP!", YELLOW)
            End If
        End If
    Else
        ' NPC not dead, just do the damage
        MapNPC(MapNum, MapNpcNum).HP = MapNPC(MapNum, MapNpcNum).HP - Damage
        Player(Attacker).TargetNPC = MapNpcNum

        ' Set the NPC target to the player
        MapNPC(MapNum, MapNpcNum).Target = Attacker
    End If
    
    Call SendDataToMap(MapNum, SPackets.Snpchp & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNPC(MapNum, MapNpcNum).HP & SEP_CHAR & GetNpcMaxHP(NpcNum) & END_CHAR)

    ' Reset attack timer
    Player(Attacker).AttackTimer = GetTickCount
End Sub

Sub TurnBasedAttackNpc(ByVal Attacker As Long, ByVal MapNpcNum As Long, ByVal Damage As Long)
    Dim MapNum As Long, NpcNum As Long
    
    MapNum = GetPlayerMap(Attacker)
    NpcNum = MapNPC(MapNum, MapNpcNum).num
    
    If Damage >= MapNPC(MapNum, MapNpcNum).HP Then
        Dim Add As String
        Dim i As Long, n As Long, Exp As Long, Index As Long, ItemNum As Long, ItemValue As Long, PartyNum As Long
        Dim CoinBadge As Boolean
        
        ' Since the player defeated the Npc, set his/her target to 0
        Player(Attacker).TargetNPC = 0

        Call OnNPCDeath(Attacker, MapNum, NpcNum, MapNpcNum)

        Add = 0
        
        For i = 1 To 7
            If GetPlayerEquipSlotNum(Attacker, i) > 0 Then
                Add = Add + Item(GetPlayerEquipSlotNum(Attacker, i)).AddEXP
            End If
        Next i

        If Add > 0 Then
            If Add < 100 Then
                If Add < 10 Then
                    Add = 0 & ".0" & Right$(Add, 2)
                Else
                    Add = 0 & "." & Right$(Add, 2)
                End If
            Else
                Add = Mid$(Add, 1, 1) & "." & Right$(Add, 2)
            End If
        ElseIf Add < 0 Then
            If Val(Add) > -100 Then
                If Val(Add) > -10 Then
                    Add = "-" & "." & 0 & Left$(-1 * Val(Add), 2)
                Else
                    Add = "-" & "." & Right$(-1 * Val(Add), 2)
                End If
            Else
                Add = "-" & Mid$(-1 * Val(Add), 1, 1) & "." & Right$(-1 * Val(Add), 2)
            End If
        End If

        ' Calculate exp to give attacker
        If Add = 0 Then
            Exp = (NPC(NpcNum).Exp * 1.2)
        Else
            Exp = (NPC(NpcNum).Exp * 1.2) + (NPC(NpcNum).Exp * Val(Add))
        End If

        ' Make sure we dont get fewer than 0 experience points
        If Exp < 0 Then
            Exp = 0
        End If
        
        PartyNum = GetPlayerPartyNum(Attacker)
        
        ' Check if the player is in a party and distribute party experience among members if applicable
        If PartyNum <= 0 Or PartyNum > MAX_PLAYERS Or GetPartyShareCount(Attacker) <= 1 Then
            Call SetPlayerExp(Attacker, GetPlayerExp(Attacker) + Exp)
        Else
            Exp = Int((Exp * (1 / (MAX_PARTY_MEMBERS - (MAX_PARTY_MEMBERS - GetPartyMembers(PartyNum))))))
                
            ' Make sure we dont get fewer than 0 experience points
            If Exp < 0 Then
                Exp = 0
            End If
                
            ' Distribute experience points to all of the party members
            For n = 1 To MAX_PARTY_MEMBERS
                Index = GetPartyMember(PartyNum, n)
                If Index > 0 Then
                    If GetPlayerPartyShare(Index) = True Then
                        Call SetPlayerExp(Index, GetPlayerExp(Index) + Exp)
                    End If
                End If
            Next n
        End If
        
        For i = 1 To 7
            VictoryInfo(Attacker, i) = "Placeholder"
        Next
        
        VictoryInfo(Attacker, 1) = CStr(Exp)
        
        CoinBadge = False
        
        If GetPlayerEquipSlotNum(Attacker, 4) = 142 Then
            CoinBadge = True
        End If
        
        Dim CoinValue As Long
        Dim ItemCount As Byte
        
        ItemCount = 2
        
        ' Handle any item drops
        For i = 1 To MAX_NPC_DROPS
            ItemNum = NPC(NpcNum).ItemNPC(i).ItemNum
            
            If ItemNum > 0 Then
                n = Int(Rnd2 * NPC(NpcNum).ItemNPC(i).Chance) + 1
                
                If n = 1 Then
                    ItemValue = NPC(NpcNum).ItemNPC(i).ItemValue
                    
                    If ItemIsStackable(ItemNum) = False Then
                        If GetFreeSlots(Attacker) > 0 Then
                            Call GiveItem(Attacker, ItemNum, ItemValue)
                        Else
                            Call SpawnItem(ItemNum, ItemValue, MapNum, GetPlayerOldX(Attacker), GetPlayerOldY(Attacker))
                        End If
                    Else
                        If GetFreeSlots(Attacker) = 0 Then
                            If CanTake(Attacker, ItemNum, 1) Then
                                Call GiveItem(Attacker, ItemNum, ItemValue)
                            Else
                                Call SpawnItem(ItemNum, ItemValue, MapNum, GetPlayerOldX(Attacker), GetPlayerOldY(Attacker))
                            End If
                        Else
                            Call GiveItem(Attacker, ItemNum, ItemValue)
                        End If
                    End If
                    
                    If ItemNum = 1 Or ItemNum = 271 Then
                        ' Gives 2x more coins with the Payoff Badge
                        If CoinBadge = True Then
                            ItemValue = (ItemValue * 2)
                        End If
                            
                        CoinValue = CoinValue + ItemValue
                    Else
                        If ItemCount < 7 Then
                            ItemCount = ItemCount + 1
                        
                            If ItemValue > 1 Then
                                VictoryInfo(Attacker, ItemCount) = Trim$(Item(ItemNum).Name) & " (" & ItemValue & ")"
                            Else
                                VictoryInfo(Attacker, ItemCount) = Trim$(Item(ItemNum).Name)
                            End If
                        End If
                    End If
                End If
            End If
        Next i
        
        VictoryInfo(Attacker, 2) = CStr(CoinValue)

        ' Now set HP to 0 so we know to actually kill them in the server loop (this prevents subscript out of range)
        Call EndBattleVictory(Attacker, MapNpcNum)
        
        MapNPC(MapNum, MapNpcNum).num = 0
        MapNPC(MapNum, MapNpcNum).SpawnWait = GetTickCount
        MapNPC(MapNum, MapNpcNum).HP = 0
        Call SendNpcDead(MapNum, MapNpcNum)

        ' Check if target is npc that died and if so set target to 0
        If Player(Attacker).TargetType = TARGET_TYPE_NPC And Player(Attacker).Target = MapNpcNum Then
            Player(Attacker).Target = 0
            Player(Attacker).TargetType = 0
        End If
        
        ' Check for the Heart Power badge
        If GetPlayerEquipSlotNum(Attacker, 7) = 317 Then
            Dim RandNum As Long
            
            RandNum = Rand(1, 20)
            
            ' Give it a 15% chance of activating
            If RandNum <= 3 Then
                Dim HPHeal As Long
                
                ' Heal the player for half the player's Stache
                HPHeal = (GetPlayerStache(Attacker) \ 2)
                
                Call SetPlayerHP(Attacker, GetPlayerHP(Attacker) + HPHeal)
                Call SendHP(Attacker)
                
                Call PlayerMsg(Attacker, "You healed " & HPHeal & " HP!", YELLOW)
            End If
        End If
    Else
        ' NPC not dead, just do the damage
        MapNPC(MapNum, MapNpcNum).HP = MapNPC(MapNum, MapNpcNum).HP - Damage
        Player(Attacker).TargetNPC = MapNpcNum

        ' Set the NPC target to the player
        MapNPC(MapNum, MapNpcNum).Target = Attacker
    End If
    
    Call SendDataTo(Attacker, SPackets.Sblitplayerdmg & SEP_CHAR & Damage & SEP_CHAR & MapNpcNum & END_CHAR)
    Call SendDataToMap(MapNum, SPackets.Snpchp & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNPC(MapNum, MapNpcNum).HP & SEP_CHAR & GetNpcMaxHP(NpcNum) & END_CHAR)

    ' Reset attack timer
    Player(Attacker).AttackTimer = GetTickCount
End Sub

Sub PlayerWarp(ByVal Index As Long, ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long)
    Dim OldMap As Long

    On Error GoTo WarpErr

    ' Check for subscript out of range.
    If Not IsPlaying(Index) Then
        Exit Sub
    End If

    ' Check for subscript out of range.
    If MapNum < 1 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If

    ' Save current number map the player is on.
    OldMap = GetPlayerMap(Index)
    
    If Not OldMap = MapNum Then
        Call SendLeaveMap(Index, OldMap)
    End If

    Call SetPlayerMap(Index, MapNum)
    Call SetPlayerX(Index, X)
    Call SetPlayerY(Index, Y)

    ' Check to see if anyone is on the map.
    If GetTotalMapPlayers(OldMap) = 0 Then
        PlayersOnMap(OldMap) = NO
    End If

    ' Sets it so we know to process npcs on the map
    PlayersOnMap(MapNum) = YES
    
    Call SendDataToMap(GetPlayerMap(Index), SPackets.Ssound & SEP_CHAR & "warp" & END_CHAR)

    Player(Index).GettingMap = YES
    
    Call SendCheckForMap(Index)

    Call SendInventory(Index)
    Call SendIndexInventoryFromMap(Index)
    Call SendIndexWornEquipmentFromMap(Index)
    Call SendMapPlayersInBattle(Index, MapNum)

    If GetPlayerMap(Index) = 243 Then
        If GetVar(App.Path & "\Scripts\" & "WelcomeMsg.ini", GetPlayerName(Index), "BeanbeanMsg") = vbNullString Then
            Call SendWelcomeMsg(Index, "Welcome to the Beanbean Kingdom!", "You've reached a new kingdom, the Beanbean Kingdom! Here, you'll find many interesting creatures, including the Beanish race. This is a kingdom unlike any other you have seen.")
            Call PutVar(App.Path & "\Scripts\" & "WelcomeMsg.ini", GetPlayerName(Index), "BeanbeanMsg", "Saw Box")
        End If
    End If

    Exit Sub

WarpErr:
    Call AddLog("PlayerWarp error for player index " & Index & " on map " & GetPlayerMap(Index) & ".", "logs\ErrorLog.txt")
End Sub

Sub PlayerMove(ByVal Index As Long, ByVal Dir As Long, ByVal Movement As Long, Xpos As Long, Ypos As Long)
    Dim i As Long, MapNum As Long, X As Long, Y As Long, pmap As Long, Xold As Long, Yold As Long
    Dim Moved As Byte

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Sub
    End If
    
    If Movement < 1 Or Movement > 2 Then
        Call HackingAttempt(Index, "Trying to move at a different speed other than walking or running.")
        Exit Sub
    End If
        
    If Player(Index).GettingMap = True Then
        Exit Sub
    End If

    Call SetPlayerDir(Index, Dir)

    ' Remove SP if the player is running
    If Movement = MOVING_RUNNING Then
        If GetPlayerAccess(Index) < 2 Then
            If GetPlayerSP(Index) > 0 Then
                Call SetPlayerSP(Index, GetPlayerSP(Index) - 1)
                Call SendSP(Index)
            End If
        End If
    End If

    Moved = NO

     ' save the current location
    Xold = GetPlayerX(Index)
    Yold = GetPlayerY(Index)
    pmap = GetPlayerMap(Index)

    ' validate map number
    If pmap <= 0 Or pmap > MAX_MAPS Then
        Call HackingAttempt(Index, vbNullString)
        Exit Sub
    End If

    ' Next check to see if we have gone outside of map boundaries
    ' If we have, then we need to try to warp to the next map if there is one
    If Xpos < 0 Or Xpos > MAX_MAPX Or Ypos < 0 Or Ypos > MAX_MAPY Then
        Select Case Dir
            Case DIR_UP
                If Ypos < 0 And Map(pmap).Up > 0 Then
                    Call PlayerWarp(Index, Map(pmap).Up, Xpos, MAX_MAPY)
                    Exit Sub
                End If
            Case DIR_DOWN
                If Ypos > MAX_MAPY And Map(pmap).Down > 0 Then
                    Call PlayerWarp(Index, Map(pmap).Down, Xpos, 0)
                    Exit Sub
                End If
            Case DIR_LEFT
                If Xpos < 0 And Map(pmap).Left > 0 Then
                    Call PlayerWarp(Index, Map(pmap).Left, MAX_MAPX, Ypos)
                    Exit Sub
                End If
            Case DIR_RIGHT
                If Xpos > MAX_MAPX And Map(pmap).Right > 0 Then
                    Call PlayerWarp(Index, Map(pmap).Right, 0, Ypos)
                    Exit Sub
                End If
        End Select
    End If
    
    ' Check to make sure our position is on the map if we haven't been warped
    If Xpos < 0 Or Xpos > MAX_MAPX Or Ypos < 0 Or Ypos > MAX_MAPY Then
        Call HackingAttempt(Index, vbNullString)
        Exit Sub
    End If
    
    ' Update coordinates to match the client - this will be correct 99% of the time
    Call SetPlayerX(Index, Xpos)
    Call SetPlayerY(Index, Ypos)
    
    ' Check to make sure that the tile is walkable
    If Map(pmap).Tile(Xpos, Ypos).Type <> TILE_TYPE_BLOCKED Then
      ' Check to see if the tile is a key and if it is check if its opened
        If (Map(pmap).Tile(Xpos, Ypos).Type <> TILE_TYPE_KEY And Map(pmap).Tile(Xpos, Ypos).Type <> TILE_TYPE_DOOR) Or ((Map(pmap).Tile(Xpos, Ypos).Type = TILE_TYPE_DOOR Or Map(pmap).Tile(Xpos, Ypos).Type = TILE_TYPE_KEY) And TempTile(pmap).DoorOpen(Xpos, Ypos) = YES) Then
            Call SendDataToMapBut(Index, pmap, SPackets.Splayermove & SEP_CHAR & Index & SEP_CHAR & Xpos & SEP_CHAR & Ypos & SEP_CHAR & Dir & SEP_CHAR & Movement & END_CHAR)
            Moved = YES
        End If
    End If

    ' At this point we have either moved, or there is a problem with the new location
    ' If we didn't move, we need to go back to the previous location and exit the sub
    If Moved <> YES Then
        Call SetPlayerX(Index, Xold)
        Call SetPlayerY(Index, Yold)
        Call SendPlayerNewXY(Index)
        Exit Sub
    End If

    ' Check if the player can start a battle by moving onto an Npc
    If (GetPlayerTurnBased(Index) = True Or IsInPoisonCave(Index) = True) And GetPlayerInBattle(Index) = False Then
        For i = 1 To MAX_MAP_NPCS
            With MapNPC(GetPlayerMap(Index), i)
                If .num > 0 Then
                    ' Prevent the player from going into battle with Rockades
                    If .num <> 195 Then
                        If NPC(.num).Behavior <> NPC_BEHAVIOR_FRIENDLY And NPC(.num).Behavior <> NPC_BEHAVIOR_SCRIPTED Then
                            If .InBattle = False Then
                                If GetPlayerX(Index) = .X And GetPlayerY(Index) = .Y Then
                                    Call OnTurnBasedBattle(Index, i, False, Xold, Yold)
                                    Exit Sub
                                End If
                            End If
                        End If
                    End If
                End If
            End With
        Next i
    End If

    If GetPlayerX(Index) + 1 <= MAX_MAPX Then
        If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index)).Type = TILE_TYPE_DOOR Then
            X = GetPlayerX(Index) + 1
            Y = GetPlayerY(Index)

            If TempTile(GetPlayerMap(Index)).DoorOpen(X, Y) = NO Then
                TempTile(GetPlayerMap(Index)).DoorOpen(X, Y) = YES
                TempTile(GetPlayerMap(Index)).DoorTimer(X, Y) = GetTickCount
                
                Call SendMapKey(GetPlayerMap(Index), X, Y, 1)
                Call SendDataToMap(GetPlayerMap(Index), SPackets.Ssound & SEP_CHAR & "door" & END_CHAR)
            End If
        End If
    End If
    If GetPlayerX(Index) - 1 >= 0 Then
        If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index)).Type = TILE_TYPE_DOOR Then
            X = GetPlayerX(Index) - 1
            Y = GetPlayerY(Index)

            If TempTile(GetPlayerMap(Index)).DoorOpen(X, Y) = NO Then
                TempTile(GetPlayerMap(Index)).DoorOpen(X, Y) = YES
                TempTile(GetPlayerMap(Index)).DoorTimer(X, Y) = GetTickCount

                Call SendMapKey(GetPlayerMap(Index), X, Y, 1)
                Call SendDataToMap(GetPlayerMap(Index), SPackets.Ssound & SEP_CHAR & "door" & END_CHAR)
            End If
        End If
    End If
    If GetPlayerY(Index) - 1 >= 0 Then
        If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) - 1).Type = TILE_TYPE_DOOR Then
            X = GetPlayerX(Index)
            Y = GetPlayerY(Index) - 1

            If TempTile(GetPlayerMap(Index)).DoorOpen(X, Y) = NO Then
                TempTile(GetPlayerMap(Index)).DoorOpen(X, Y) = YES
                TempTile(GetPlayerMap(Index)).DoorTimer(X, Y) = GetTickCount

                Call SendMapKey(GetPlayerMap(Index), X, Y, 1)
                Call SendDataToMap(GetPlayerMap(Index), SPackets.Ssound & SEP_CHAR & "door" & END_CHAR)
            End If
        End If
    End If
    If GetPlayerY(Index) + 1 <= MAX_MAPY Then
        If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) + 1).Type = TILE_TYPE_DOOR Then
            X = GetPlayerX(Index)
            Y = GetPlayerY(Index) + 1

            If TempTile(GetPlayerMap(Index)).DoorOpen(X, Y) = NO Then
                TempTile(GetPlayerMap(Index)).DoorOpen(X, Y) = YES
                TempTile(GetPlayerMap(Index)).DoorTimer(X, Y) = GetTickCount

                Call SendMapKey(GetPlayerMap(Index), X, Y, 1)
                Call SendDataToMap(GetPlayerMap(Index), SPackets.Ssound & SEP_CHAR & "door" & END_CHAR)
            End If
        End If
    End If
    
    ' Resets value for Respawn Points and Heart Blocks
    If GetVar(App.Path & "\Respawn Points.ini", GetPlayerName(Index), "CTRL") = "1" Then
        Call PutVar(App.Path & "\Respawn Points.ini", GetPlayerName(Index), "CTRL", "0")
    End If
    
    If GetVar(App.Path & "\Heart Blocks.ini", GetPlayerName(Index), "CTRL") = "1" Then
        Call PutVar(App.Path & "\Heart Blocks.ini", GetPlayerName(Index), "CTRL", "0")
    End If
    
    ' Reset value for Simultaneous Blocks
    If Map(GetPlayerMap(Index)).Tile(Xold, Yold).Type = TILE_TYPE_SIMULBLOCK Then
        Call PutVar(App.Path & "\SimulBlocks.ini", GetPlayerMap(Index), Xold & "/" & Yold, vbNullString)
    End If
    
    ' Execute tile
    Select Case Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type
      ' Heal tile
        Case TILE_TYPE_HEAL
            Call SetPlayerHP(Index, GetPlayerMaxHP(Index))
            Call SetPlayerMP(Index, GetPlayerMaxMP(Index))
            Call SetPlayerSP(Index, GetPlayerMaxSP(Index))
            Call SendHP(Index)
            Call SendMP(Index)
            Call SendSP(Index)
            Call PlayerMsg(Index, "You just got fully healed! You feel amazing!", BRIGHTGREEN)
            Call SendSoundTo(Index, "spm_get_health.wav")
            Call SpellAnim(4, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index))
      ' Kill tile
        Case TILE_TYPE_KILL
            Call SetPlayerHP(Index, 0)
            Call SendHP(Index)
      ' Warp tile
        Case TILE_TYPE_WARP
            MapNum = Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1
            X = Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data2
            Y = Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data3

            Call PlayerWarp(Index, MapNum, X, Y)
      ' Key Open tile
        Case TILE_TYPE_KEYOPEN
            X = Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1
            Y = Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data2

            If Map(GetPlayerMap(Index)).Tile(X, Y).Type = TILE_TYPE_KEY And TempTile(GetPlayerMap(Index)).DoorOpen(X, Y) = NO Then
                TempTile(GetPlayerMap(Index)).DoorOpen(X, Y) = YES
                TempTile(GetPlayerMap(Index)).DoorTimer(X, Y) = GetTickCount
                
                Call SendMapKey(GetPlayerMap(Index), X, Y, 1)
                
                If Trim$(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String1) = vbNullString Then
                    Call MapMsg(GetPlayerMap(Index), "A door has been unlocked!", WHITE)
                Else
                    Call MapMsg(GetPlayerMap(Index), Trim$(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String1), WHITE)
                End If
                Call SendDataToMap(GetPlayerMap(Index), SPackets.Ssound & SEP_CHAR & "key" & END_CHAR)
            End If
      ' Shop tile
        Case TILE_TYPE_SHOP
            If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1 > 0 Then
                Call SendTrade(Index, Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1)
            End If
      ' Sprite Change tile
        Case TILE_TYPE_SPRITE_CHANGE
            If GetPlayerSprite(Index) <> Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1 Then
                Call SetPlayerSprite(Index, Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1)
                Call SendDataToMap(GetPlayerMap(Index), SPackets.Schecksprite & SEP_CHAR & Index & SEP_CHAR & GetPlayerSprite(Index) & END_CHAR)
                Exit Sub
            End If
      ' Class Change tile
        Case TILE_TYPE_CLASS_CHANGE
            If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data2 > -1 Then
                If GetPlayerClass(Index) <> Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data2 Then
                    Call PlayerMsg(Index, "You arent the required class!", BRIGHTRED)
                    Exit Sub
                End If
            End If

            If GetPlayerClass(Index) = Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1 Then
                Call PlayerMsg(Index, "You are already this class!", BRIGHTRED)
            Else
                If Player(Index).Char(Player(Index).CharNum).Sex = 0 Then
                    If GetPlayerSprite(Index) = ClassData(GetPlayerClass(Index)).MaleSprite Then
                        Call SetPlayerSprite(Index, ClassData(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1).MaleSprite)
                    End If
                Else
                    If GetPlayerSprite(Index) = ClassData(GetPlayerClass(Index)).FemaleSprite Then
                        Call SetPlayerSprite(Index, ClassData(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1).FemaleSprite)
                    End If
                End If

                Call SetPlayerSTR(Index, (Player(Index).Char(Player(Index).CharNum).STR - ClassData(GetPlayerClass(Index)).STR))
                Call SetPlayerDEF(Index, (Player(Index).Char(Player(Index).CharNum).DEF - ClassData(GetPlayerClass(Index)).DEF))
                Call SetPlayerStache(Index, (Player(Index).Char(Player(Index).CharNum).Magi - ClassData(GetPlayerClass(Index)).Magi))
                Call SetPlayerSPEED(Index, (Player(Index).Char(Player(Index).CharNum).Speed - ClassData(GetPlayerClass(Index)).Speed))

                Call SetPlayerClassData(Index, Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1)

                Call SetPlayerSTR(Index, (Player(Index).Char(Player(Index).CharNum).STR + ClassData(GetPlayerClass(Index)).STR))
                Call SetPlayerDEF(Index, (Player(Index).Char(Player(Index).CharNum).DEF + ClassData(GetPlayerClass(Index)).DEF))
                Call SetPlayerStache(Index, (Player(Index).Char(Player(Index).CharNum).Magi + ClassData(GetPlayerClass(Index)).Magi))
                Call SetPlayerSPEED(Index, (Player(Index).Char(Player(Index).CharNum).Speed + ClassData(GetPlayerClass(Index)).Speed))

                Call PlayerMsg(Index, "Your new class is a " & Trim$(ClassData(GetPlayerClass(Index)).Name) & "!", BRIGHTGREEN)

                Call SendStats(Index)
                Call SendHP(Index)
                Call SendMP(Index)
                Call SendSP(Index)
                Player(Index).Char(Player(Index).CharNum).MAXHP = GetPlayerMaxHP(Index)
                Player(Index).Char(Player(Index).CharNum).MAXMP = GetPlayerMaxMP(Index)
                Player(Index).Char(Player(Index).CharNum).MAXSP = GetPlayerMaxSP(Index)
                Call SendDataToMap(GetPlayerMap(Index), SPackets.Schecksprite & SEP_CHAR & Index & SEP_CHAR & GetPlayerSprite(Index) & END_CHAR)
            End If
      ' Check if player stepped on notice tile
        Case TILE_TYPE_NOTICE
            If Trim$(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String1) <> vbNullString Then
                Call PlayerMsg(Index, Trim$(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String1), BLACK)
            End If
            If Trim$(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String2) <> vbNullString Then
                Call PlayerMsg(Index, Trim$(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String2), DARKGREY)
            End If
            If Not Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String3 = vbNullString Or Not Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String3 = vbNullString Then
                Call SendSoundToMap(GetPlayerMap(Index), Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String3)
            End If
      ' Check if player steppted on minus stat tile
        Case TILE_TYPE_LOWER_STAT
            If Trim$(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String1) <> vbNullString Then
                Call PlayerMsg(Index, Trim$(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String1), BRIGHTRED)
            End If
            If Trim$(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1) <> 0 Then
                Call SetPlayerHP(Index, GetPlayerHP(Index) - Trim$(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1))
                Call SendHP(Index)
            End If
            If Trim$(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data2) <> 0 Then
                Call SetPlayerMP(Index, GetPlayerMP(Index) - Trim$(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data2))
                Call SendMP(Index)
            End If
            If Trim$(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data3) <> 0 Then
                Call SetPlayerSP(Index, GetPlayerSP(Index) - Trim$(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data3))
                Call SendSP(Index)
            End If
      ' Check if player stepped on sound tile
        Case TILE_TYPE_SOUND
            Call SendSoundToMap(GetPlayerMap(Index), Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String1)
      ' Check if player stepped on scripted tile
        Case TILE_TYPE_SCRIPTED
            Call ScriptedTile(Index, Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1)
      ' Check if player stepped on Bank tile
        Case TILE_TYPE_BANK
            Call SendDataTo(Index, SPackets.Sopenbank & END_CHAR)
        Case Else
            Exit Sub
    End Select
End Sub

Function CanNpcMove(ByVal MapNum As Long, ByVal MapNpcNum As Long, ByVal Dir As Long) As Boolean
    Dim i As Long
    Dim TileType As Long
    Dim X As Long
    Dim Y As Long

    ' Check for sub-script out of range.
    If MapNum < 1 Or MapNum > MAX_MAPS Then
        Exit Function
    End If

    ' Check for sub-script out of range.
    If MapNpcNum < 1 Or MapNpcNum > MAX_MAP_NPCS Then
        Exit Function
    End If

    ' Check for sub-script out of range.
    If Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Function
    End If
    
    If MapNPC(MapNum, MapNpcNum).InBattle = True Then
        Exit Function
    End If
    
    X = MapNPC(MapNum, MapNpcNum).X
    Y = MapNPC(MapNum, MapNpcNum).Y

    Select Case Dir
        Case DIR_UP
            If Y <= 0 Then
                Exit Function
            End If
            
            Y = Y - 1
        Case DIR_DOWN
            If Y >= MAX_MAPY Then
                Exit Function
            End If
            
            Y = Y + 1
        Case DIR_LEFT
            If X <= 0 Then
                Exit Function
            End If
        
            X = X - 1
        Case DIR_RIGHT
            If X >= MAX_MAPX Then
                Exit Function
            End If
            
            X = X + 1
    End Select
    
    ' Get the attribute on the tile
    TileType = Map(MapNum).Tile(X, Y).Type
    
    ' Check to make sure that the tile is walkable
    If TileType = TILE_TYPE_BLOCKED Or TileType = TILE_TYPE_NPCAVOID Or TileType = TILE_TYPE_SIGN Or TileType = TILE_TYPE_HOOKSHOT Or TileType = TILE_TYPE_ROOFBLOCK Or TileType = TILE_TYPE_SWITCH Or TileType = TILE_TYPE_KEY Or TileType = TILE_TYPE_JUMPBLOCK Or TileType = TILE_TYPE_WARP Then
        Exit Function
    End If

    ' Check to make sure that there isn't a player in the way
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            If GetPlayerMap(i) = MapNum Then
                If GetPlayerX(i) = X Then
                    If GetPlayerY(i) = Y Then
                        If Map(MapNum).Moral <> MAP_MORAL_MINIGAME Then
                            If NPC(MapNPC(MapNum, MapNpcNum).num).Behavior <> NPC_BEHAVIOR_FRIENDLY And NPC(MapNPC(MapNum, MapNpcNum).num).Behavior <> NPC_BEHAVIOR_SCRIPTED Then
                                If GetPlayerTurnBased(i) = False And IsInPoisonCave(i) = False Then
                                    Exit Function
                                Else
                                    If GetPlayerInBattle(i) = False Then
                                        If GetTickCount <= (GetPlayerRecoverTime(i) + 2000) Then
                                            Exit Function
                                        End If
                                    Else
                                        Exit Function
                                    End If
                                End If
                            Else
                                Exit Function
                            End If
                        Else
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
    Next i

    ' Check to make sure that there is not another npc in the way
    For i = 1 To MAX_MAP_NPCS
        If i <> MapNpcNum Then
            If MapNPC(MapNum, i).num > 0 Then
                If MapNPC(MapNum, i).X = X Then
                    If MapNPC(MapNum, i).Y = Y Then
                        Exit Function
                    End If
                End If
            End If
        End If
    Next i
    
    CanNpcMove = True
End Function

Public Sub NpcMove(ByVal MapNum As Long, ByVal MapNpcNum As Long, ByVal Dir As Long, ByVal Movement As Long)
    ' Check to make sure it's a valid map.
    If MapNum < 1 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If

    ' Check to make sure it's a valid NPC.
    If MapNpcNum < 1 Or MapNpcNum > MAX_MAP_NPCS Then
        Exit Sub
    End If

    ' Check to make sure it's a valid direction.
    If Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Sub
    End If

    ' Check to make sure it's a valid movement speed.
    If Movement < 1 Or Movement > 2 Then
        Exit Sub
    End If

    MapNPC(MapNum, MapNpcNum).Dir = Dir

    Select Case Dir
        Case DIR_UP
            MapNPC(MapNum, MapNpcNum).Y = MapNPC(MapNum, MapNpcNum).Y - 1
        Case DIR_DOWN
            MapNPC(MapNum, MapNpcNum).Y = MapNPC(MapNum, MapNpcNum).Y + 1
        Case DIR_LEFT
            MapNPC(MapNum, MapNpcNum).X = MapNPC(MapNum, MapNpcNum).X - 1
        Case DIR_RIGHT
            MapNPC(MapNum, MapNpcNum).X = MapNPC(MapNum, MapNpcNum).X + 1
    End Select

    Call SendDataToMap(MapNum, SPackets.Snpcmove & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNPC(MapNum, MapNpcNum).X & SEP_CHAR & MapNPC(MapNum, MapNpcNum).Y & SEP_CHAR & MapNPC(MapNum, MapNpcNum).Dir & SEP_CHAR & Movement & END_CHAR)

    Dim i As Long
    
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            If GetPlayerMap(i) = MapNum Then
                If GetPlayerX(i) = MapNPC(MapNum, MapNpcNum).X And GetPlayerY(i) = MapNPC(MapNum, MapNpcNum).Y Then
                    If NPC(MapNPC(MapNum, MapNpcNum).num).Behavior <> NPC_BEHAVIOR_FRIENDLY And NPC(MapNPC(MapNum, MapNpcNum).num).Behavior <> NPC_BEHAVIOR_SCRIPTED Then
                        Call OnTurnBasedBattle(i, MapNpcNum)
                        Exit Sub
                    End If
                End If
            End If
        End If
    Next
End Sub

Public Sub NpcDir(ByVal MapNum As Long, ByVal MapNpcNum As Long, ByVal Dir As Long)
    ' Check to make sure it's a valid map.
    If MapNum < 1 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If

    ' Check to make sure it's a valid NPC.
    If MapNpcNum < 1 Or MapNpcNum > MAX_MAP_NPCS Then
        Exit Sub
    End If

    ' Check to make sure it's a valid direction.
    If Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Sub
    End If

    MapNPC(MapNum, MapNpcNum).Dir = Dir

    Call SendDataToMap(MapNum, SPackets.Snpcdir & SEP_CHAR & MapNpcNum & SEP_CHAR & Dir & END_CHAR)
End Sub

Public Sub JoinGame(ByVal Index As Long)
    Dim MOTD As String
    Dim i As Integer
    
    MOTD = GetVar(App.Path & "\MOTD.ini", "MOTD", "Msg")
    
    ' Set the flag so we know the person is in the game
    Player(Index).InGame = True

    ' Send an ok to client to start receiving in game data
    Call SendDataTo(Index, SPackets.Sloginok & SEP_CHAR & Index & END_CHAR)

    Call CheckEquippedItems(Index)
    Call SendClasses(Index)
    Call SendItems(Index)
    Call SendEmoticons(Index)
    Call SendElements(Index)
    Call SendArrows(Index)
    Call SendNpcs(Index)
    Call SendShops(Index)
    Call SendSpells(Index)
    Call SendRecipes(Index)
    Call SendInventory(Index)
    Call SendBank(Index)
    Call SendHP(Index)
    Call SendMP(Index)
    Call SendSP(Index)
    Call SendPTS(Index)
    Call SendStats(Index)
    Call SendWeatherTo(Index)
    Call SendPlayerSpells(Index)
    Call SendOnlineList
    Call SendPlayerHeight(Index)
    
    ' Sets max Attack, Defense, Speed, and Stache
    Call SetPlayerMaxSTR(Index, Player(Index).Char(Player(Index).CharNum).STR)
    Call SetPlayerMaxDEF(Index, Player(Index).Char(Player(Index).CharNum).DEF)
    Call SetPlayerMaxSpeed(Index, Player(Index).Char(Player(Index).CharNum).Speed)
    Call SetPlayerMaxStache(Index, Player(Index).Char(Player(Index).CharNum).Magi)
    
    ' Give the player passive bonuses from special attacks if the player didn't receive them yet
    If GetVar(App.Path & "\Passive Bonuses.ini", GetPlayerName(Index), "GotBonuses") <> "Yes" Then
        Dim SpellNum As Long
        
        For i = 1 To MAX_PLAYER_SPELLS
            SpellNum = GetPlayerSpell(Index, i)
        
            If SpellNum > 0 Then
                If Spell(SpellNum).UsePassiveStat = True Then
                    Call AddPassiveStatBonus(Index, SpellNum)
                End If
            End If
        Next i
        
        Call PutVar(App.Path & "\Passive Bonuses.ini", GetPlayerName(Index), "GotBonuses", "Yes")
    End If
    
    ' Set the player's sprite
    If GetPlayerTempSprite(Index) > 0 Then
        If GetPlayerSprite(Index) <> GetPlayerTempSprite(Index) Then
            Call SetPlayerSprite(Index, GetPlayerTempSprite(Index))
        End If
    End If
    
    ' Warp the player to his saved location
    Call PlayerWarp(Index, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index))
    
    ' Refresh recovery time from turn-based battles
    Call SetPlayerRecoverTime(Index, 0)
    
    ' Set player mute status
    If IsMuted(Index) Then
        Player(Index).Mute = True
    Else
        Player(Index).Mute = False
    End If
    
    ' Sends different messages on login depending on access
    Select Case GetPlayerAccess(Index)
        Case 0
            Call GlobalMsg(GetPlayerName(Index) & " has joined Super Mario Bros. Online!", GREEN)
        Case 1
            Call GlobalMsg("Moderator " & GetPlayerName(Index) & " has joined Super Mario Bros. Online!", BLACK)
        Case 2
            Call GlobalMsg("Designer " & GetPlayerName(Index) & " has joined Super Mario Bros. Online!", BRIGHTBLUE)
        Case 3
            Call GlobalMsg("Developer " & GetPlayerName(Index) & " has joined Super Mario Bros. Online!", BROWN)
        Case 4
            Call GlobalMsg("Administrator " & GetPlayerName(Index) & " has joined Super Mario Bros. Online!", WHITE)
        Case 5
            Call GlobalMsg("Creator " & GetPlayerName(Index) & " has joined Super Mario Bros. Online!", YELLOW)
    End Select

    Call PlayerMsg(Index, "Welcome to Super Mario Bros. Online!", GREEN)

    If LenB(MOTD) <> 0 Then
        Call PlayerMsg(Index, "MOTD: " & MOTD, BRIGHTCYAN)
    End If
    
    ' Tell the client the player is in-game.
    Call SendDataTo(Index, SPackets.Singame & SEP_CHAR & GetPlayerName(Index) & END_CHAR)

    ' Give the player the number of inventory slots he/she deserves
    Call ObtainInventorySlots(Index)

    ' Give the player any special attacks he/she missed
    Call LearnSpecialAttacks(Index)

    ' Update the server console.
    Call ShowPLR(Index)
    
    ' Welcome new players
    If GetVar(App.Path & "\Scripts\" & "WelcomeMsg.ini", GetPlayerName(Index), "WelcomeMsg") = vbNullString Then
        Call SendWelcomeMsg(Index, "Welcome to Super Mario Bros. Online!", "Welcome to Super Mario Bros. Online! We hope you enjoy your stay!" & vbNewLine & vbNewLine & "Use the arrow keys to move, and press Enter to read signs once you're in front of them. Click Ok to close this menu.")
        Call PutVar(App.Path & "\Scripts\" & "WelcomeMsg.ini", GetPlayerName(Index), "WelcomeMsg", "Saw Box")
    End If
    
    ' Loads Friends Lists
    For i = 1 To 10
        Call SendDataTo(Index, SPackets.Scaption & SEP_CHAR & GetVar(App.Path & "\SMBOAccounts\" & "Friend Lists.ini", GetPlayerName(Index), CStr(i)) & SEP_CHAR & i & END_CHAR)
    Next i
End Sub

Public Sub LeftGame(ByVal Index As Long)
    Dim MapNum As Long, MapNpcNum As Long
    Dim FilePath As String
    Dim AttackChange As String, DefenseChange As String, SpeedChange As String, StacheChange As String
    Dim AttackItemChange As String, DefenseItemChange As String, SpeedItemChange As String, StacheItemChange As String
    
    MapNum = GetPlayerMap(Index)
    
    FilePath = App.Path & "\Scripts\" & "StatIncreases.ini"
    
    AttackItemChange = Trim$(GetVar(FilePath, GetPlayerName(Index), "ItemAttack"))
    DefenseItemChange = Trim$(GetVar(FilePath, GetPlayerName(Index), "ItemDefense"))
    SpeedItemChange = Trim$(GetVar(FilePath, GetPlayerName(Index), "ItemSpeed"))
    StacheItemChange = Trim$(GetVar(FilePath, GetPlayerName(Index), "ItemStache"))
    AttackChange = Trim$(GetVar(FilePath, GetPlayerName(Index), "Attack"))
    DefenseChange = Trim$(GetVar(FilePath, GetPlayerName(Index), "Defense"))
    SpeedChange = Trim$(GetVar(FilePath, GetPlayerName(Index), "Speed"))
    StacheChange = Trim$(GetVar(FilePath, GetPlayerName(Index), "Stache"))
    
' Removes current stat modifications to players
  ' Attack
    If AttackChange = "Yes" Or AttackItemChange = "Yes" Then
        Call SetPlayerSTR(Index, Player(Index).Char(Player(Index).CharNum).MAXSTR)
    End If
    If AttackChange = "Yes" Then
        Call PutVar(FilePath, GetPlayerName(Index), "Attack", "No")
    End If
    If AttackItemChange = "Yes" Then
        Call PutVar(FilePath, GetPlayerName(Index), "ItemAttack", "No")
        Call PutVar(App.Path & "\Scripts\" & "StatIncreases.ini", GetPlayerName(Index), "RedPeppers", "No")
    End If
  ' Defense
    If DefenseChange = "Yes" Or DefenseItemChange = "Yes" Then
        Call SetPlayerDEF(Index, Player(Index).Char(Player(Index).CharNum).MAXDEF)
    End If
    If DefenseChange = "Yes" Then
        Call PutVar(FilePath, GetPlayerName(Index), "Defense", "No")
    End If
    If DefenseItemChange = "Yes" Then
        Call PutVar(FilePath, GetPlayerName(Index), "ItemDefense", "No")
        Call PutVar(App.Path & "\Scripts\" & "StatIncreases.ini", GetPlayerName(Index), "GreenPeppers", "No")
    End If
  ' Speed
    If SpeedChange = "Yes" Or SpeedItemChange = "Yes" Then
        Call SetPlayerSPEED(Index, Player(Index).Char(Player(Index).CharNum).MAXSpeed)
    End If
    If SpeedChange = "Yes" Then
        Call PutVar(FilePath, GetPlayerName(Index), "Speed", "No")
    End If
    If SpeedItemChange = "Yes" Then
        Call PutVar(FilePath, GetPlayerName(Index), "ItemSpeed", "No")
    End If
  ' Stache
    If StacheChange = "Yes" Or StacheItemChange = "Yes" Then
        Call SetPlayerStache(Index, Player(Index).Char(Player(Index).CharNum).MAXStache)
    End If
    If StacheChange = "Yes" Then
        Call PutVar(FilePath, GetPlayerName(Index), "Stache", "No")
    End If
    If StacheItemChange = "Yes" Then
        Call PutVar(FilePath, GetPlayerName(Index), "ItemStache", "No")
    End If
    
    ' Removes players from a battle if they're in one
    If GetPlayerInBattle(Index) = True Then
        MapNpcNum = GetPlayerTargetNpc(Index)
        
        ' Get the player and NPC out of the battle
        If MapNpcNum > 0 Then
            MapNPC(MapNum, MapNpcNum).InBattle = False
            MapNPC(MapNum, MapNpcNum).Turn = False
            MapNPC(MapNum, MapNpcNum).X = MapNPC(MapNum, MapNpcNum).OldX
            MapNPC(MapNum, MapNpcNum).Y = MapNPC(MapNum, MapNpcNum).OldY
            MapNPC(MapNum, MapNpcNum).HP = GetNpcMaxHP(MapNPC(MapNum, MapNpcNum).num)
            Call SendTurnBasedBattle(Index, 0, MapNpcNum)
            ' Set target to 0
            MapNPC(MapNum, MapNpcNum).Target = 0
        End If
        
        Call SetPlayerInBattle(Index, False)
        Call SetPlayerTurn(Index, False)
        Call SetPlayerX(Index, GetPlayerOldX(Index))
        Call SetPlayerY(Index, GetPlayerOldY(Index))
        
        ' Set that the player is not in the victory animation
        IsInVictoryAnim(Index) = False
        
      ' Set target to 0
        Player(Index).TargetNPC = 0
        Call SendMapNpcsToMap(MapNum)
    End If
    
    ' Takes players out of Steal the Shroom minigame
    Call LeftSTS(Index)
    
    ' Takes players out of Whack-A-Monty minigame
    Call LeftWhackAMonty(Index)
    
    ' Takes players out of Dodgebill minigame
    Call LeftDodgeBill(Index)
    
    ' Takes players out of Hide n' Sneak minigame
    Call LeftHideNSneak(Index)
    
    If Player(Index).InGame Then
        Player(Index).InGame = False

        ' Stop processing NPCs if no one is on it.
        If MapNum > 0 Then
            If GetTotalMapPlayers(MapNum) = 0 Then
                PlayersOnMap(MapNum) = NO
            End If
        End If

        ' If player is in party, remove experience decrease.
        If GetPlayerPartyNum(Index) > 0 Then
            Call LeaveParty(Index)
        End If

        ' Sends different messages on logout depending on access
        Select Case GetPlayerAccess(Index)
            Case 0
                Call GlobalMsg(GetPlayerName(Index) & " has logged off Super Mario Bros. Online!", GREEN)
            Case 1
                Call GlobalMsg("Moderator " & GetPlayerName(Index) & " has logged off Super Mario Bros. Online!", BLACK)
            Case 2
                Call GlobalMsg("Designer " & GetPlayerName(Index) & " has logged off Super Mario Bros. Online!", BRIGHTBLUE)
            Case 3
                Call GlobalMsg("Developer " & GetPlayerName(Index) & " has logged off Super Mario Bros. Online!", BROWN)
            Case 4
                Call GlobalMsg("Administrator " & GetPlayerName(Index) & " has logged off Super Mario Bros. Online!", WHITE)
            Case 5
                Call GlobalMsg("Creator " & GetPlayerName(Index) & " has logged off Super Mario Bros. Online!", YELLOW)
        End Select
        
        Call SavePlayer(Index)
        Call SendLeftGame(Index)

        Call TextAdd(frmServer.txtText(0), GetPlayerName(Index) & " has disconnected from Super Mario Bros. Online.", True)

        Call RemovePLR(Index)
    End If

    Call ClearPlayer(Index)
    Call SendOnlineList
End Sub

Private Sub LeftSTS(ByVal Index As Long)
    If GetPlayerMap(Index) = 32 Or GetPlayerMap(Index) = 33 Then
        Dim i As Integer
        Dim PlayerNum As String
    
        For i = 1 To 4
            PlayerNum = CStr(i)
        
            If GetVar(STSPath, "Red", PlayerNum) = GetPlayerName(Index) Then
                Call PutVar(STSPath, "Team", "Red", CInt(GetVar(STSPath, "Team", "Red")) - 1)
                Call PutVar(STSPath, "Red", PlayerNum, "")
            
                Exit For
            ElseIf GetVar(STSPath, "Blue", PlayerNum) = GetPlayerName(Index) Then
                Call PutVar(STSPath, "Team", "Blue", CInt(GetVar(STSPath, "Team", "Blue")) - 1)
                Call PutVar(STSPath, "Blue", PlayerNum, "")
            
                Exit For
            End If
        Next i
      
        Call SetPlayerMap(Index, 31)
        Call SetPlayerX(Index, 20)
        Call SetPlayerY(Index, 13)
    End If
End Sub

Private Sub LeftWhackAMonty(ByVal Index As Long)
    If GetPlayerMap(Index) >= 72 And GetPlayerMap(Index) <= 74 Then
        Dim n As Integer, i As Integer
        
        ' Remove Montys from the game
        For n = 1 To MAX_MAP_NPCS
            Call ClearMapNpc(n, GetPlayerMap(Index))
        Next n
        
        Call SetPlayerMap(Index, 71)
        Call SetPlayerX(Index, 25)
        Call SetPlayerY(Index, 11)
        
        Dim FilePath As String, GameNum As String
        
        ' Resets the game
        FilePath = App.Path & "\Scripts\" & "Whack.ini"
        
        For i = 1 To 3
            GameNum = "Game" & i
        
            If GetVar(FilePath, GameNum, "Player") = GetPlayerName(Index) Then
                Call PutVar(FilePath, GameNum, "TimeLeft", "0")
                Call PutVar(FilePath, GameNum, "InGame", "No")
                Call PutVar(FilePath, GameNum, "Points", "0")
                Call PutVar(FilePath, GameNum, "Player", "")
                
                Exit For
            End If
        Next i
    End If
End Sub

Private Sub LeftDodgeBill(ByVal Index As Long)
    If GetPlayerMap(Index) = 188 Or GetPlayerMap(Index) = 191 Then
        Dim i As Integer
        Dim PlayerNum As String
        
        For i = 1 To 5
            PlayerNum = CStr(i)
            
            If GetVar(DodgeBillPath, "Red", PlayerNum) = GetPlayerName(Index) Then
                Call PutVar(DodgeBillPath, "Team", "Red", CInt(GetVar(DodgeBillPath, "Team", "Red")) - 1)
                Call PutVar(DodgeBillPath, "Red", PlayerNum, "")
            
                Exit For
            ElseIf GetVar(DodgeBillPath, "Blue", PlayerNum) = GetPlayerName(Index) Then
                Call PutVar(DodgeBillPath, "Team", "Blue", CInt(GetVar(DodgeBillPath, "Team", "Blue")) - 1)
                Call PutVar(DodgeBillPath, "Blue", PlayerNum, "")
                
                Exit For
            End If
        Next i
        
        ' Remove bullet bills from the player's inventory if he/she has any
        If GetPlayerMap(Index) = 188 Then
            Dim n As Integer
            
            For i = 1 To GetPlayerMaxInv(Index)
                If GetPlayerInvItemNum(Index, i) = 186 Then
                    n = i
                    Exit For
                End If
            Next i
            
            If n > 0 Then
                Dim BulletBillCount As Byte
                Dim MapItemNum As Integer
                
                BulletBillCount = GetPlayerInvItemValue(Index, n)
                
                ' Spawn the item
                If BulletBillCount > 0 Then
                    MapItemNum = FindOpenMapItemSlot(188)
                    
                    Call SetPlayerInvItemNum(Index, n, 0)
                    Call SetPlayerInvItemValue(Index, n, 0)
                    Call SetPlayerInvItemAmmo(Index, n, -1)
                    Call SpawnItemSlot(MapItemNum, 186, BulletBillCount, 1, 188, GetPlayerX(Index), GetPlayerY(Index))
                End If
            End If
        End If
        
        Call SetPlayerMap(Index, 190)
        Call SetPlayerX(Index, 21)
        Call SetPlayerY(Index, 10)
    End If
End Sub

Private Sub LeftHideNSneak(ByVal Index As Long)
    If GetPlayerMap(Index) >= 270 And GetPlayerMap(Index) <= 273 Then
        Dim i As Integer, NumHiders As Integer, NumSeekers As Integer
        Dim PlayerNum As String
        
        NumHiders = 1
        NumSeekers = 1
        
        For i = 1 To MaxHiders
            PlayerNum = CStr(i)
            
            ' Remove the player from the Hiders list
            If GetVar(HideNSneakPath, "Hiders", PlayerNum) = GetPlayerName(Index) Then
                NumHiders = CInt(GetVar(HideNSneakPath, "Team", "Hiders")) - 1
                
                Call PutVar(HideNSneakPath, "Team", "Hiders", CStr(NumHiders))
                Call PutVar(HideNSneakPath, "Hiders", PlayerNum, "")
                
                Exit For
            End If
            
            If i <= MaxSeekers Then
                ' Remove the player from the Seekers list
                If GetVar(HideNSneakPath, "Seekers", PlayerNum) = GetPlayerName(Index) Then
                    NumSeekers = CInt(GetVar(HideNSneakPath, "Team", "Seekers")) - 1
                    
                    Call PutVar(HideNSneakPath, "Team", "Seekers", CStr(NumSeekers))
                    Call PutVar(HideNSneakPath, "Seekers", PlayerNum, "")
                    
                    Exit For
                End If
            End If
        Next i
        
        ' Check if there are any hiders or seekers remaining in the game
        If NumHiders <= 0 Or NumSeekers <= 0 Then
            ' State that the game ended because a player left
            HasLeftHideNSneak = True
            
            ' State that the Seekers have already seeked
            Call PutVar(HideNSneakPath, "GameTime", "IsSeeking", "Yes")
            
            ' Set the TimeLeft to 0
            Call PutVar(HideNSneakPath, "GameTime", "TimeLeft", "0")
            
            On Error Resume Next
            
            ' End the game immediately if there are either no hiders or seekers left in the game
            Call HideNSneakPlayTime(CLng(GetVar(HideNSneakPath, "TimerIndex", "TimerIndex")))
        End If
        
        ' Change the player's sprite back to the old sprite
        Call SetPlayerSprite(Index, GetPlayerTempSprite(Index))
        
        ' Warp the player to the Beanbean Arcade
        Call SetPlayerMap(Index, 269)
        Call SetPlayerX(Index, 15)
        Call SetPlayerY(Index, 7)
    End If
End Sub

Function GetTotalMapPlayers(ByVal MapNum As Long) As Long
    Dim i As Long

    If MapNum < 1 Or MapNum > MAX_MAPS Then
        Exit Function
    End If

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            If GetPlayerMap(i) = MapNum Then
                GetTotalMapPlayers = GetTotalMapPlayers + 1
            End If
        End If
    Next i
End Function

Function GetNpcMaxHP(ByVal NpcNum As Long) As Long
    If NpcNum < 1 Or NpcNum > MAX_NPCS Then
        Exit Function
    End If
    
    GetNpcMaxHP = NPC(NpcNum).MAXHP
End Function

Function GetNpcMaxMP(ByVal NpcNum As Long) As Long
    If NpcNum < 1 Or NpcNum > MAX_NPCS Then
        Exit Function
    End If

    GetNpcMaxMP = NPC(NpcNum).Magi * 2
End Function

Function GetNpcMaxSP(ByVal NpcNum As Long) As Long
    If NpcNum < 1 Or NpcNum > MAX_NPCS Then
        Exit Function
    End If

    GetNpcMaxSP = NPC(NpcNum).Speed * 2
End Function

Function GetPlayerSPRegen(ByVal Index As Long) As Integer
    Dim Total As Integer

    If Index < 1 Or Index > MAX_PLAYERS Then
        Exit Function
    End If

    If Not IsPlaying(Index) Then
        Exit Function
    End If

    Total = (GetPlayerSPEED(Index) \ 2)
    If Total < 2 Then
        Total = 2
    End If

    GetPlayerSPRegen = Total
End Function

Sub AddPassiveStatBonus(ByVal Index As Long, ByVal SpellNum As Long)
    If Spell(SpellNum).UsePassiveStat = True Then
        Select Case Spell(SpellNum).PassiveStat
            Case 0 ' HP
                Dim HPLvl As Integer
                
                HPLvl = Val(GetVar(App.Path & "\Level Up.ini", GetPlayerName(Index), "HP"))
                HPLvl = HPLvl + Spell(SpellNum).PassiveStatChange
                
                Call PutVar(App.Path & "\Level Up.ini", GetPlayerName(Index), "HP", CStr(HPLvl))
            Case 1 ' FP
                Dim FPLvl As Integer
                
                FPLvl = Val(GetVar(App.Path & "\Level Up.ini", GetPlayerName(Index), "FP"))
                FPLvl = FPLvl + Spell(SpellNum).PassiveStatChange
                
                Call PutVar(App.Path & "\Level Up.ini", GetPlayerName(Index), "FP", CStr(FPLvl))
            Case 2 ' Attack
                Call SetPlayerSTR(Index, Player(Index).Char(Player(Index).CharNum).STR + Spell(SpellNum).PassiveStatChange)
                Call SetPlayerMaxSTR(Index, Player(Index).Char(Player(Index).CharNum).MAXSTR + Spell(SpellNum).PassiveStatChange)
            Case 3 ' Defense
                Call SetPlayerDEF(Index, Player(Index).Char(Player(Index).CharNum).DEF + Spell(SpellNum).PassiveStatChange)
                Call SetPlayerMaxDEF(Index, Player(Index).Char(Player(Index).CharNum).MAXDEF + Spell(SpellNum).PassiveStatChange)
            Case 4 ' Speed
                Call SetPlayerSPEED(Index, Player(Index).Char(Player(Index).CharNum).Speed + Spell(SpellNum).PassiveStatChange)
                Call SetPlayerMaxSpeed(Index, Player(Index).Char(Player(Index).CharNum).MAXSpeed + Spell(SpellNum).PassiveStatChange)
            Case 5 ' Stache
                Call SetPlayerStache(Index, Player(Index).Char(Player(Index).CharNum).Magi + Spell(SpellNum).PassiveStatChange)
                Call SetPlayerMaxStache(Index, Player(Index).Char(Player(Index).CharNum).MAXStache + Spell(SpellNum).PassiveStatChange)
        End Select
    End If
End Sub

Sub CheckPlayerLevelUp(ByVal Index As Long)
    If GetPlayerLevel(Index) <> MAX_LEVEL Then
        If GetPlayerExp(Index) >= GetPlayerNextLevel(Index) Then
            Call PlayerLevelUp(Index)
                
            Call SendDataToMap(GetPlayerMap(Index), SPackets.Slevelup & SEP_CHAR & Index & END_CHAR)
            Call SendPlayerLevelToAll(Index)
        
            ' Learn special attacks on level up
            Call LearnSpecialAttacks(Index)
                
            ' Gain inventory slots on level up
            Call ObtainInventorySlots(Index)
        End If
    Else
        If GetPlayerExp(Index) >= Experience(MAX_LEVEL) Then
            Call SetPlayerExp(Index, Experience(MAX_LEVEL))
        End If
    End If
    
    Call SendHP(Index)
    Call SendMP(Index)
    Call SendSP(Index)
    Call SendPTS(Index)

    Player(Index).Char(Player(Index).CharNum).MAXHP = GetPlayerMaxHP(Index)
    Player(Index).Char(Player(Index).CharNum).MAXMP = GetPlayerMaxMP(Index)
    Player(Index).Char(Player(Index).CharNum).MAXSP = GetPlayerMaxSP(Index)

    Call SendStats(Index)
End Sub

Sub LearnSpecialAttacks(ByVal Index As Long)
    Dim i As Long, SpellSlot As Long, ClassReq As Long
    
    ' Learning special attacks on level up
    For i = 1 To MAX_SPELLS
        SpellSlot = GetSpellReqLevel(i)
        ClassReq = Spell(i).ClassReq
    
        If GetPlayerLevel(Index) >= SpellSlot And SpellSlot > 0 Then
            If GetPlayerClass(Index) = (ClassReq - 1) Or ClassReq = 0 Then
                SpellSlot = FindOpenSpellSlot(Index)

                If SpellSlot > 0 Then
                    If Not HasSpell(Index, i) Then
                        Call SetPlayerSpell(Index, SpellSlot, i)
                        
                        ' Add any passive bonuses to the player's stats
                        Call AddPassiveStatBonus(Index, i)
                        
                        Call PlayerMsg(Index, "You've learned " & Trim$(Spell(i).Name) & "!", WHITE)
                    End If
                End If
            End If
        End If
    Next i

    Call SendPlayerSpells(Index)
End Sub

Sub ObtainInventorySlots(ByVal Index As Long)
    ' Stop players from getting more than 28 inventory slots
    If GetPlayerMaxInv(Index) >= 28 Then
        Exit Sub
    End If
    
    Dim PlayerMaxInv As Integer, TempPlayerMaxInv As Integer
    
    TempPlayerMaxInv = GetPlayerMaxInv(Index)
    PlayerMaxInv = MAX_INV
    
    Dim i As Integer
    
    ' Loop through the values: 10, 20, 30, 40
    For i = 10 To 40 Step 10
        If GetPlayerLevel(Index) >= i Then
            PlayerMaxInv = PlayerMaxInv + 1
        Else
            Exit For
        End If
    Next i
        
    If GetPlayerMaxInv(Index) <> PlayerMaxInv Then
        Player(Index).Char(Player(Index).CharNum).MaxInv = PlayerMaxInv
        
        ReDim Preserve Player(Index).Char(Player(Index).CharNum).NewInv(1 To PlayerMaxInv) As NewPlayerInvRec
        
        ' Clear out the new slots' information
        Dim j As Integer
        
        For j = (TempPlayerMaxInv + 1) To PlayerMaxInv
            Call SetPlayerInvItemNum(Index, j, 0)
            Call SetPlayerInvItemValue(Index, j, 0)
            Call SetPlayerInvItemAmmo(Index, j, -1)
        Next
        
        Call SavePlayer(Index)
        
        Dim InvDifference As Integer
        
        InvDifference = (PlayerMaxInv - TempPlayerMaxInv)
        
        ' Send a different message for one new inventory slot and more than one new inventory slot
        If InvDifference = 1 Then
            Call PlayerMsg(Index, "You've gained another inventory slot! You can now hold " & GetPlayerMaxInv(Index) & " items in your inventory!", WHITE)
        ElseIf InvDifference > 1 Then
            Call PlayerMsg(Index, "You've gained " & InvDifference & " inventory slots! You can now hold " & GetPlayerMaxInv(Index) & " items in your inventory!", WHITE)
        End If
        
        Call SendInventory(Index)
        Exit Sub
    End If
End Sub

Sub UseSpecialAttack(ByVal Index As Long, ByVal SpellSlot As Long)
    Dim SpellNum As Long
    
    SpellNum = GetPlayerSpell(Index, SpellSlot)
    
    ' Spell doesn't exist
    If SpellNum <= 0 Then
        Exit Sub
    End If
    
    ' Players can't cast the Flower Saver special attack
    If SpellNum = 43 Then
        Exit Sub
    End If
    
    ' Determine which Sub to use
    
    ' Area effect
    If Spell(SpellNum).AE = 1 Then
        If GetPlayerTurnBased(Index) = True Or IsInPoisonCave(Index) = True Then
            If Player(Index).TargetNPC > 0 Then
                Call UseSpecialOnNpc(Index, SpellSlot)
                Exit Sub
            Else
                Call UseSpecialArea(Index, SpellSlot)
                Exit Sub
            End If
        Else
            Call UseSpecialArea(Index, SpellSlot)
            Exit Sub
        End If
    End If
    
    ' Self-spell
    If Spell(SpellNum).SelfSpell = True Then
        Call UseSpecialOnSelf(Index, SpellSlot)
        Exit Sub
    End If
    
    ' Non-area effect
    If Player(Index).TargetType = TARGET_TYPE_PLAYER Then
        ' Cast the special attack on ourselves if we're the target
        If Player(Index).Target = Index Then
            Call UseSpecialOnSelf(Index, SpellSlot)
        Else
            Call UseSpecialOnPlayer(Index, SpellSlot)
        End If
    ' Cast the special attack
    Else
        Call UseSpecialOnNpc(Index, SpellSlot)
    End If
End Sub

Sub UseSpecialOnSelf(ByVal Index As Long, ByVal SpellSlot As Long)
    Dim SpellWorking As String
    Dim SpellNum As Long, Damage As Long, MapNum As Long
    
    ' Make sure all the requirements are met
    If CanUseSpecial(Index, SpellSlot) = False Then
        Exit Sub
    End If
    
    MapNum = GetPlayerMap(Index)
    
    ' Stops players from using special attacks in minigames
    If Map(MapNum).Moral = MAP_MORAL_MINIGAME Then
        Call PlayerMsg(Index, "You cannot use special attacks in a minigame!", BRIGHTRED)
        Exit Sub
    End If
    
    SpellNum = GetPlayerSpell(Index, SpellSlot)
    
    If Spell(SpellNum).Type = SPELL_TYPE_SUBHP Or Spell(SpellNum).Type = SPELL_TYPE_SUBMP Or Spell(SpellNum).Type = SPELL_TYPE_SUBSP Then
        Call PlayerMsg(Index, "You cannot use that type of special attack on yourself!", BRIGHTRED)
        Exit Sub
    End If
    
    SpellWorking = Trim$(GetVar(App.Path & "\Scripts\" & "StatIncreases.ini", GetPlayerName(Index), GetSpellStatName(Spell(SpellNum).Stat)))
    
    Select Case Spell(SpellNum).Type
        Case SPELL_TYPE_ADDHP
            Call SetPlayerHP(Index, GetPlayerHP(Index) + Spell(SpellNum).Data1)
            Call SendHP(Index)
        Case SPELL_TYPE_ADDMP
            Call SetPlayerMP(Index, GetPlayerMP(Index) + Spell(SpellNum).Data1)
            Call SendMP(Index)
        Case SPELL_TYPE_ADDSP
            Call SetPlayerSP(Index, GetPlayerSP(Index) + Spell(SpellNum).Data1)
            Call SendSP(Index)
        Case SPELL_TYPE_STATCHANGE
            If SpellWorking <> "Yes" Then
                Call StartStatMod(Index, Spell(SpellNum).Multiplier, Spell(SpellNum).StatTime, Spell(SpellNum).Stat)
            Else
                Call PlayerMsg(Index, "You must wait for the effects of your special attack to end!", WHITE)
                Exit Sub
            End If
        Case SPELL_TYPE_SCRIPTED
            If GetVar(App.Path & "\Scripts\" & "StatIncreases.ini", GetPlayerName(Index), "Attack") = "Yes" Or GetVar(App.Path & "\Scripts\" & "StatIncreases.ini", GetPlayerName(Index), "Defense") = "Yes" Then
                Call PlayerMsg(Index, "You cannot use this special attack when there is another special attack modifying your Attack or Defense!", BRIGHTRED)
                Exit Sub
            Else
                Dim ItemAttack As String, ItemDefense As String
                
                ' Check to see if any Attack or Defense modifying items are currently being used
                ItemAttack = GetVar(App.Path & "\Scripts\" & "StatIncreases.ini", GetPlayerName(Index), "ItemAttack")
                ItemDefense = GetVar(App.Path & "\Scripts\" & "StatIncreases.ini", GetPlayerName(Index), "ItemDefense")
            
                If ItemAttack = "Yes" Or ItemDefense = "Yes" Then
                    Call PlayerMsg(Index, "You would be too powerful if you used this Special Attack right now.", WHITE)
                    Exit Sub
                End If
            
                Call ScriptedSpell(Index, Spell(SpellNum).Data1)
            End If
    End Select
                    
    ' Take away the FP
    Call SetPlayerMP(Index, GetPlayerMP(Index) - FlowerSaver(Index, SpellNum))
    Call SendMP(Index)
    
    ' State that the player used the special attack
    Player(Index).CastedSpell = YES
    
    ' Send the animation
    Call SendDataToMap(MapNum, SPackets.Sspellanim & SEP_CHAR & SpellNum & SEP_CHAR & Spell(SpellNum).SpellAnim & SEP_CHAR & Spell(SpellNum).SpellTime & SEP_CHAR & Spell(SpellNum).SpellDone & SEP_CHAR & Index & SEP_CHAR & TARGET_TYPE_PLAYER & SEP_CHAR & Index & SEP_CHAR & YES & SEP_CHAR & Spell(SpellNum).Big & END_CHAR)
    
    ' Send the sound
    Call SendSoundToMap(MapNum, Spell(SpellNum).Sound)
    
    ' Reset attack timer
    Player(Index).AttackTimer = GetTickCount
End Sub

Sub UseSpecialOnPlayer(ByVal Index As Long, ByVal SpellSlot As Long)
    Dim SpellWorking As String
    Dim SpellNum As Long, Damage As Long, MapNum As Long
    Dim Target As Byte
    
    ' Make sure all the requirements are met
    If CanUseSpecial(Index, SpellSlot) = False Then
        Exit Sub
    End If
    
    Target = Player(Index).Target
    
    ' Make sure the target is playing
    If IsPlaying(Target) = False Then
        Exit Sub
    End If
    
    MapNum = GetPlayerMap(Index)
    
    ' Make sure both players are on the same map
    If GetPlayerMap(Target) <> MapNum Then
        Exit Sub
    End If
    
    ' Make sure the target is alive
    If GetPlayerHP(Target) <= 0 Then
        Exit Sub
    End If
    
    ' Stops players from using special attacks in minigames
    If Map(MapNum).Moral = MAP_MORAL_MINIGAME Then
        Call PlayerMsg(Index, "You cannot use special attacks in a minigame!", BRIGHTRED)
        Exit Sub
    End If
    
    SpellNum = GetPlayerSpell(Index, SpellSlot)
    
    If GetPlayerInBattle(Target) = True Then
        Call PlayerMsg(Index, "You cannot use special attacks on a player that's in battle!", BRIGHTRED)
        Exit Sub
    End If
    
    If MapNum <> 218 Then
        If Map(MapNum).Moral <> MAP_MORAL_NONE And Spell(SpellNum).Type = SPELL_TYPE_SUBHP Then
            If GetPlayerPK(Target) = NO Then
                Call PlayerMsg(Index, "This is not a PvP area!", BRIGHTRED)
                Exit Sub
            End If
        End If
    Else
        If Map(MapNum).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type <> TILE_TYPE_ARENA Or Map(MapNum).Tile(GetPlayerX(Target), GetPlayerY(Target)).Type <> TILE_TYPE_ARENA Then
            Call PlayerMsg(Index, "This is not a PvP area!", BRIGHTRED)
            Exit Sub
        End If
    End If
    
    SpellWorking = Trim$(GetVar(App.Path & "\Scripts\" & "StatIncreases.ini", GetPlayerName(Target), GetSpellStatName(Spell(SpellNum).Stat)))
    
    ' Check if we can reach the target
    If CInt(Sqr((GetPlayerX(Index) - GetPlayerX(Target)) ^ 2 + ((GetPlayerY(Index) - GetPlayerY(Target)) ^ 2))) > Spell(SpellNum).Range Then
        Call PlayerMsg(Index, "Your special attack's range isn't far enough to hit your target.", BRIGHTRED)
        Exit Sub
    End If
                
    Select Case Spell(SpellNum).Type
        Case SPELL_TYPE_ADDHP
            Call SetPlayerHP(Target, GetPlayerHP(Target) + Spell(SpellNum).Data1)
            Call SendHP(Target)
        Case SPELL_TYPE_ADDMP
            Call SetPlayerMP(Target, GetPlayerMP(Target) + Spell(SpellNum).Data1)
            Call SendMP(Target)
        Case SPELL_TYPE_ADDSP
            Call SetPlayerSP(Target, GetPlayerSP(Target) + Spell(SpellNum).Data1)
            Call SendSP(Target)
        Case SPELL_TYPE_SUBHP
            If GetPlayerGuild(Index) = GetPlayerGuild(Target) And GetPlayerGuild(Index) <> vbNullString Then
                Call PlayerMsg(Index, GetPlayerName(Target) & " is in the same Group as you, so you cannot attack him/her!", BRIGHTRED)
                Exit Sub
            End If
            
            Damage = ((GetPlayerSTR(Index) * 0.6) + PityFlower(Index, Spell(SpellNum).Data1)) - GetPlayerProtection(Target)
            Damage = DamageUpDamageDown(Index, Damage, Target)
            Damage = Int(Rand(Damage - 2, Damage + 2))
            
            If Damage > 0 Then
                Call AttackPlayer(Index, Target, Damage)
            Else
                Call BattleMsg(Index, "The special attack was too weak to harm " & GetPlayerName(Target) & "!", BRIGHTRED, 0)
            End If
        Case SPELL_TYPE_SUBMP
            Call SetPlayerMP(Target, GetPlayerMP(Target) - Spell(SpellNum).Data1)
            Call SendMP(Target)
        Case SPELL_TYPE_SUBSP
            Call SetPlayerSP(Target, GetPlayerSP(Target) - Spell(SpellNum).Data1)
            Call SendSP(Target)
        Case SPELL_TYPE_STATCHANGE
            If SpellWorking <> "Yes" Then
                Call StartStatMod(Target, Spell(SpellNum).Multiplier, Spell(SpellNum).StatTime, Spell(SpellNum).Stat)
            Else
                Call PlayerMsg(Index, "You must wait for the effects of " & GetPlayerName(Target) & "'s special attack to end!", WHITE)
                Exit Sub
            End If
        Case SPELL_TYPE_SCRIPTED
            If GetVar(App.Path & "\Scripts\" & "StatIncreases.ini", GetPlayerName(Target), "Attack") = "Yes" Or GetVar(App.Path & "\Scripts\" & "StatIncreases.ini", GetPlayerName(Target), "Defense") = "Yes" Then
                Call PlayerMsg(Index, "You cannot use this special attack when there is another special attack modifying " & GetPlayerName(Target) & "'s Attack or Defense!", BRIGHTRED)
                Exit Sub
            Else
                Dim ItemAttack As String, ItemDefense As String
                
                ' Check to see if any Attack or Defense modifying items are currently being used
                ItemAttack = GetVar(App.Path & "\Scripts\" & "StatIncreases.ini", GetPlayerName(Target), "ItemAttack")
                ItemDefense = GetVar(App.Path & "\Scripts\" & "StatIncreases.ini", GetPlayerName(Target), "ItemDefense")
            
                If ItemAttack = "Yes" Or ItemDefense = "Yes" Then
                    Call PlayerMsg(Index, GetPlayerName(Target) & " would be too powerful if you used this Special Attack right now.", WHITE)
                    Exit Sub
                End If
                
                Call ScriptedSpell(Target, Spell(SpellNum).Data1)
            End If
    End Select
    
    ' Take away the FP
    Call SetPlayerMP(Index, GetPlayerMP(Index) - FlowerSaver(Index, SpellNum))
    Call SendMP(Index)
    
    ' State that the player used the special attack
    Player(Index).CastedSpell = YES
    
    ' Send the animation
    Call SendDataToMap(MapNum, SPackets.Sspellanim & SEP_CHAR & SpellNum & SEP_CHAR & Spell(SpellNum).SpellAnim & SEP_CHAR & Spell(SpellNum).SpellTime & SEP_CHAR & Spell(SpellNum).SpellDone & SEP_CHAR & Index & SEP_CHAR & Player(Index).TargetType & SEP_CHAR & Target & SEP_CHAR & YES & SEP_CHAR & Spell(SpellNum).Big & END_CHAR)
    
    ' Send the sound
    Call SendSoundToMap(MapNum, Spell(SpellNum).Sound)
    
    ' Reset attack timer
    Player(Index).AttackTimer = GetTickCount
End Sub

Sub UseSpecialOnNpc(ByVal Index As Long, ByVal SpellSlot As Long)
    ' Make sure all the requirements are met
    If CanUseSpecial(Index, SpellSlot) = False Then
        Exit Sub
    End If
    
    Dim Target As Byte
    
    Target = Player(Index).TargetNPC
    
    ' Make sure the npc exists
    If Target <= 0 Then
        Exit Sub
    End If
    
    Dim MapNum As Long
    
    MapNum = GetPlayerMap(Index)
    
    ' Stops players from using special attacks in minigames
    If Map(MapNum).Moral = MAP_MORAL_MINIGAME Then
        Call PlayerMsg(Index, "You cannot use special attacks in a minigame!", BRIGHTRED)
        Exit Sub
    End If
    
    Dim NpcNum As Long
    
    NpcNum = MapNPC(MapNum, Target).num
    
    ' Make sure the npc is attackable
    If NPC(NpcNum).Behavior = NPC_BEHAVIOR_FRIENDLY Or NPC(NpcNum).Behavior = NPC_BEHAVIOR_SHOPKEEPER Or NPC(NpcNum).Behavior = NPC_BEHAVIOR_SCRIPTED Then
        Call PlayerMsg(Index, "You cannot use a special attack on that npc!", BRIGHTRED)
        Exit Sub
    End If
    
    Dim SpellNum As Long
    
    SpellNum = GetPlayerSpell(Index, SpellSlot)
     
    If CInt(Sqr((GetPlayerX(Index) - MapNPC(MapNum, Target).X) ^ 2 + ((GetPlayerY(Index) - MapNPC(MapNum, Target).Y) ^ 2))) > Spell(SpellNum).Range Then
        Call PlayerMsg(Index, "Your special attack's range isn't far enough to hit your target.", BRIGHTRED)
        Exit Sub
    End If
    
    Select Case Spell(SpellNum).Type
        Case SPELL_TYPE_ADDHP
            MapNPC(MapNum, Target).HP = MapNPC(MapNum, Target).HP + Spell(SpellNum).Data1
            ' resets HP to max if it goes above
            If MapNPC(MapNum, Target).HP >= GetNpcMaxHP(NpcNum) Then
                MapNPC(MapNum, Target).HP = GetNpcMaxHP(NpcNum)
            End If
                
            ' Update npc hp
            Call SendDataToMap(MapNum, SPackets.Snpchp & SEP_CHAR & Target & SEP_CHAR & MapNPC(MapNum, Target).HP & SEP_CHAR & GetNpcMaxHP(NpcNum) & END_CHAR)
        Case SPELL_TYPE_ADDMP
                MapNPC(MapNum, Target).MP = MapNPC(MapNum, Target).MP + Spell(SpellNum).Data1
        Case SPELL_TYPE_ADDSP
                MapNPC(MapNum, Target).SP = MapNPC(MapNum, Target).SP + Spell(SpellNum).Data1
        Case SPELL_TYPE_SUBHP
            If NpcNum <> 195 Then
                If SpellNum = 45 Then
                    Call PlayerMsg(Index, "You can't use this special attack on the " & Trim$(NPC(NpcNum).Name) & "!", BRIGHTRED)
                    Exit Sub
                End If
                
                Dim Damage As Long
            
                Damage = ((GetPlayerSTR(Index) * 0.6) + PityFlower(Index, Spell(SpellNum).Data1)) - Int(NPC(NpcNum).DEF)
                Damage = DamageUpDamageDown(Index, Damage)
                Damage = Int(Rand(Damage - 2, Damage + 2))
                
                If Damage > 0 Then
                    Call AttackNpc(Index, Target, Damage)
                Else
                    Call BattleMsg(Index, "The special attack was too weak to harm " & Trim$(NPC(NpcNum).Name) & "!", BRIGHTRED, 0)
                End If
            Else
                If SpellNum <> 45 Then
                    Exit Sub
                Else
                    Call AttackNpc(Index, Target, 1)
                End If
            End If
        Case SPELL_TYPE_SUBMP
            MapNPC(MapNum, Target).MP = MapNPC(MapNum, Target).MP - Spell(SpellNum).Data1
        Case SPELL_TYPE_SUBSP
            MapNPC(MapNum, Target).SP = MapNPC(MapNum, Target).SP - Spell(SpellNum).Data1
        Case Else
            Exit Sub
    End Select
    
    ' Take away the FP
    Call SetPlayerMP(Index, GetPlayerMP(Index) - FlowerSaver(Index, SpellNum))
    Call SendMP(Index)
    
    ' State that the player used the special attack
    Player(Index).CastedSpell = YES
    
    ' Send the animation
    Call SendDataToMap(MapNum, SPackets.Sspellanim & SEP_CHAR & SpellNum & SEP_CHAR & Spell(SpellNum).SpellAnim & SEP_CHAR & Spell(SpellNum).SpellTime & SEP_CHAR & Spell(SpellNum).SpellDone & SEP_CHAR & Index & SEP_CHAR & Player(Index).TargetType & SEP_CHAR & Target & SEP_CHAR & YES & SEP_CHAR & Spell(SpellNum).Big & END_CHAR)
    
    ' Send the sound
    Call SendSoundToMap(MapNum, Spell(SpellNum).Sound)
    
    ' Reset attack timer
    Player(Index).AttackTimer = GetTickCount
End Sub

Sub UseSpecialArea(ByVal Index As Long, ByVal SpellSlot As Long)
    Dim SpellWorking As String
    Dim i As Long, StartX As Long, EndX As Long, StartY As Long, EndY As Long, SpellNum As Long, Damage As Long, MapNum As Long
    Dim Target As Byte, HitCount As Byte
    Dim Casted As Boolean
    
    ' Make sure all the requirements are met
    If CanUseSpecial(Index, SpellSlot) = False Then
        Exit Sub
    End If
    
    SpellNum = GetPlayerSpell(Index, SpellSlot)
    
    MapNum = GetPlayerMap(Index)

    ' Stops players from using special attacks in minigames
    If Map(MapNum).Moral = MAP_MORAL_MINIGAME Then
        Call PlayerMsg(Index, "You cannot use special attacks in a minigame!", BRIGHTRED)
        Exit Sub
    End If
    
    ' Find targets within range
    StartX = GetPlayerX(Index) - Spell(SpellNum).Range
    EndX = GetPlayerX(Index) + Spell(SpellNum).Range
    
    StartY = GetPlayerY(Index) - Spell(SpellNum).Range
    EndY = GetPlayerY(Index) + Spell(SpellNum).Range
    
    If StartX < 0 Then StartX = 0
    If EndX > 30 Then EndX = 30
    If StartY < 0 Then StartY = 0
    If EndY > 30 Then EndY = 30
    
    Player(Index).TargetType = TARGET_TYPE_PLAYER
    
    ' Loop through the players first
    For i = 1 To MAX_PLAYERS
        Casted = False
        If IsPlaying(i) Then
            If i <> Index Then
                If GetPlayerMap(i) = MapNum Then
                    ' Make sure the player isn't in battle
                    If GetPlayerInBattle(i) = False Then
                        ' Make sure the player is in range
                        If GetPlayerY(i) >= StartY And GetPlayerY(i) <= EndY And GetPlayerX(i) >= StartX And GetPlayerX(i) <= EndX Then
                            Target = CByte(i)
                            ' Start the effects
                            SpellWorking = Trim$(GetVar(App.Path & "\Scripts\" & "StatIncreases.ini", GetPlayerName(Target), GetSpellStatName(Spell(SpellNum).Stat)))
                                
                            Select Case Spell(SpellNum).Type
                                Case SPELL_TYPE_ADDHP
                                    Call SetPlayerHP(Target, GetPlayerHP(Target) + Spell(SpellNum).Data1)
                                    Call SendHP(Target)
                                Case SPELL_TYPE_ADDMP
                                    Call SetPlayerMP(Target, GetPlayerMP(Target) + Spell(SpellNum).Data1)
                                    Call SendMP(Target)
                                Case SPELL_TYPE_ADDSP
                                    Call SetPlayerSP(Target, GetPlayerSP(Target) + Spell(SpellNum).Data1)
                                    Call SendSP(Target)
                                Case SPELL_TYPE_SUBHP
                                    If Map(MapNum).Moral = MAP_MORAL_NONE Or GetPlayerPK(Target) = YES Then
                                        If Not (GetPlayerGuild(Index) = GetPlayerGuild(Target) And GetPlayerGuild(Index) <> vbNullString) Then
                                            Damage = ((GetPlayerSTR(Index) * 0.6) + PityFlower(Index, Spell(SpellNum).Data1)) - GetPlayerProtection(Target)
                                            Damage = DamageUpDamageDown(Index, Damage, Target)
                                            Damage = Int(Rand(Damage - 2, Damage + 2))

                                            If Damage > 0 Then
                                                Call AttackPlayer(Index, Target, Damage)
                                            Else
                                                Call BattleMsg(Index, "The special attack was too weak to harm " & GetPlayerName(Target) & "!", BRIGHTRED, 0)
                                            End If
                                        End If
                                    End If
                                Case SPELL_TYPE_SUBMP
                                    Call SetPlayerMP(Target, GetPlayerMP(Target) - Spell(SpellNum).Data1)
                                    Call SendMP(Target)
                                Case SPELL_TYPE_SUBSP
                                    Call SetPlayerSP(Target, GetPlayerSP(Target) - Spell(SpellNum).Data1)
                                    Call SendSP(Target)
                                Case SPELL_TYPE_STATCHANGE
                                    If SpellWorking <> "Yes" Then
                                        Call StartStatMod(Target, Spell(SpellNum).Multiplier, Spell(SpellNum).StatTime, Spell(SpellNum).Stat)
                                    Else
                                        Call PlayerMsg(Index, "You must wait for the effects of " & GetPlayerName(Target) & "'s special attack to end!", WHITE)
                                    End If
                                Case SPELL_TYPE_SCRIPTED
                                    If GetVar(App.Path & "\Scripts\" & "StatIncreases.ini", GetPlayerName(Target), "Attack") = "Yes" Or GetVar(App.Path & "\Scripts\" & "StatIncreases.ini", GetPlayerName(Target), "Defense") = "Yes" Then
                                        Call PlayerMsg(Index, "You cannot use this special attack when there is another special attack modifying " & GetPlayerName(Target) & "'s Attack or Defense!", BRIGHTRED)
                                    Else
                                        Dim ItemAttack As String, ItemDefense As String
                
                                        ' Check to see if any Attack or Defense modifying items are currently being used
                                        ItemAttack = GetVar(App.Path & "\Scripts\" & "StatIncreases.ini", GetPlayerName(Target), "ItemAttack")
                                        ItemDefense = GetVar(App.Path & "\Scripts\" & "StatIncreases.ini", GetPlayerName(Target), "ItemDefense")
                                    
                                        If ItemAttack = "Yes" Or ItemDefense = "Yes" Then
                                            Call PlayerMsg(Index, GetPlayerName(Target) & " would be too powerful if you used this Special Attack right now.", WHITE)
                                            Exit Sub
                                        End If
                                        
                                        Call ScriptedSpell(Target, Spell(SpellNum).Data1)
                                    End If
                            End Select
                        
                            Casted = True
                            HitCount = HitCount + 1
                        End If
                    End If
                End If
            Else
                Target = Index
            
                Select Case Spell(SpellNum).Type
                    Case SPELL_TYPE_ADDHP
                        Call SetPlayerHP(Index, GetPlayerHP(Index) + Spell(SpellNum).Data1)
                        Call SendHP(Index)
                        
                        Casted = True
                        HitCount = HitCount + 1
                    Case SPELL_TYPE_ADDMP
                        Call SetPlayerMP(Index, GetPlayerMP(Index) + Spell(SpellNum).Data1)
                        Call SendMP(Index)
                        
                        Casted = True
                        HitCount = HitCount + 1
                    Case SPELL_TYPE_ADDSP
                        Call SetPlayerSP(Index, GetPlayerSP(Index) + Spell(SpellNum).Data1)
                        Call SendSP(Index)
                        
                        Casted = True
                        HitCount = HitCount + 1
                End Select
            End If
        End If
        
        If Casted = True Then
            ' State that the player used the special attack
            Player(Index).CastedSpell = YES
            
            ' Send the animation
            Call SendDataToMap(MapNum, SPackets.Sspellanim & SEP_CHAR & SpellNum & SEP_CHAR & Spell(SpellNum).SpellAnim & SEP_CHAR & Spell(SpellNum).SpellTime & SEP_CHAR & Spell(SpellNum).SpellDone & SEP_CHAR & Index & SEP_CHAR & Player(Index).TargetType & SEP_CHAR & Target & SEP_CHAR & YES & SEP_CHAR & Spell(SpellNum).Big & END_CHAR)
        End If
    Next i
                    
    Player(Index).TargetType = TARGET_TYPE_NPC
    
    ' Loop through npcs
    For i = 1 To MAX_MAP_NPCS
        Casted = False
        If MapNPC(MapNum, i).num > 0 Then
            If NPC(MapNPC(MapNum, i).num).Behavior <> NPC_BEHAVIOR_FRIENDLY And NPC(MapNPC(MapNum, i).num).Behavior <> NPC_BEHAVIOR_SHOPKEEPER And NPC(MapNPC(MapNum, i).num).Behavior <> NPC_BEHAVIOR_SCRIPTED Then
                ' Make sure the npc isn't in battle
                If MapNPC(MapNum, i).InBattle = False Then
                    ' Make sure the npc is in range
                    If MapNPC(MapNum, i).Y >= StartY And MapNPC(MapNum, i).Y <= EndY And MapNPC(MapNum, i).X >= StartX And MapNPC(MapNum, i).X <= EndX Then
                        Target = CByte(i)
                        
                        Casted = True
                        HitCount = HitCount + 1
                        
                        ' Start the effects
                        Select Case Spell(SpellNum).Type
                            Case SPELL_TYPE_ADDHP
                                MapNPC(MapNum, Target).HP = MapNPC(MapNum, Target).HP + Spell(SpellNum).Data1
                                ' resets HP to max if it goes above
                                If MapNPC(MapNum, Target).HP >= GetNpcMaxHP(MapNPC(MapNum, Target).num) Then
                                    MapNPC(MapNum, Target).HP = GetNpcMaxHP(MapNPC(MapNum, Target).num)
                                End If
                                            
                                ' Update npc hp
                                Call SendDataToMap(MapNum, SPackets.Snpchp & SEP_CHAR & CLng(Target) & SEP_CHAR & MapNPC(MapNum, Target).HP & SEP_CHAR & GetNpcMaxHP(MapNPC(MapNum, Target).num) & END_CHAR)
                            Case SPELL_TYPE_ADDMP
                                MapNPC(MapNum, Target).MP = MapNPC(MapNum, Target).MP + Spell(SpellNum).Data1
                            Case SPELL_TYPE_ADDSP
                                MapNPC(MapNum, Target).SP = MapNPC(MapNum, Target).SP + Spell(SpellNum).Data1
                            Case SPELL_TYPE_SUBHP
                                If MapNPC(MapNum, Target).num <> 195 Then
                                    Damage = ((GetPlayerSTR(Index) * 0.6) + PityFlower(Index, Spell(SpellNum).Data1)) - Int(NPC(MapNPC(MapNum, Target).num).DEF)
                                    Damage = DamageUpDamageDown(Index, Damage)
                                    Damage = Int(Rand(Damage - 2, Damage + 2))
                                            
                                    If Damage > 0 Then
                                        Call AttackNpc(Index, Target, Damage)
                                    Else
                                        Call BattleMsg(Index, "The special attack was too weak to harm " & Trim$(NPC(MapNPC(MapNum, Target).num).Name) & "!", BRIGHTRED, 0)
                                    End If
                                    
                                    ' Stop the AoE special attack from harming other Npcs if the user got into a battle
                                    If GetPlayerTurnBased(Index) = True Or IsInPoisonCave(Index) = True Then
                                        ' State that the player used the special attack
                                        Player(Index).CastedSpell = YES
                                            
                                        ' Send the animation
                                        Call SendDataToMap(MapNum, SPackets.Sspellanim & SEP_CHAR & SpellNum & SEP_CHAR & Spell(SpellNum).SpellAnim & SEP_CHAR & Spell(SpellNum).SpellTime & SEP_CHAR & Spell(SpellNum).SpellDone & SEP_CHAR & Index & SEP_CHAR & Player(Index).TargetType & SEP_CHAR & Target & SEP_CHAR & YES & SEP_CHAR & Spell(SpellNum).Big & END_CHAR)
                                            
                                        Exit For
                                    End If
                                End If
                            Case SPELL_TYPE_SUBMP
                                MapNPC(MapNum, Target).MP = MapNPC(MapNum, Target).MP - Spell(SpellNum).Data1
                            Case SPELL_TYPE_SUBSP
                                MapNPC(MapNum, Target).SP = MapNPC(MapNum, Target).SP - Spell(SpellNum).Data1
                            Case Else
                                Casted = False
                                HitCount = HitCount - 1
                        End Select
                    End If
                End If
            End If
        End If
        
        If Casted = True Then
            ' State that the player used the special attack
            Player(Index).CastedSpell = YES
            
            ' Send the animation
            Call SendDataToMap(MapNum, SPackets.Sspellanim & SEP_CHAR & SpellNum & SEP_CHAR & Spell(SpellNum).SpellAnim & SEP_CHAR & Spell(SpellNum).SpellTime & SEP_CHAR & Spell(SpellNum).SpellDone & SEP_CHAR & Index & SEP_CHAR & Player(Index).TargetType & SEP_CHAR & Target & SEP_CHAR & YES & SEP_CHAR & Spell(SpellNum).Big & END_CHAR)
        End If
    Next i
        
    ' Only do all of this if we actually hit any players or npcs
    If HitCount > 0 Then
        ' Take away the FP
        Call SetPlayerMP(Index, GetPlayerMP(Index) - FlowerSaver(Index, SpellNum))
        Call SendMP(Index)
        
        ' Send the sound
        Call SendSoundToMap(MapNum, Spell(SpellNum).Sound)
        
        ' Reset attack timer
        Player(Index).AttackTimer = GetTickCount
    Else
        Call PlayerMsg(Index, "Your special attack's range isn't far enough to affect any targets.", BRIGHTRED)
        Exit Sub
    End If
End Sub

Sub UseSpecialTurnBased(ByVal Index As Long, ByVal SpellSlot As Long)
    Dim SpellWorking As String
    Dim SpellNum As Long, Damage As Long, MapNum As Long, Target As Long
    Dim TargetType As Byte
    
    MapNum = GetPlayerMap(Index)
    SpellNum = GetPlayerSpell(Index, SpellSlot)
    
    SpellWorking = Trim$(GetVar(App.Path & "\Scripts\" & "StatIncreases.ini", GetPlayerName(Index), GetSpellStatName(Spell(SpellNum).Stat)))
    
    Select Case Spell(SpellNum).Type
        Case SPELL_TYPE_ADDHP
            Call SetPlayerHP(Index, GetPlayerHP(Index) + Spell(SpellNum).Data1)
            Call SendHP(Index)
            
            Target = Index
            TargetType = TARGET_TYPE_PLAYER
        Case SPELL_TYPE_ADDMP
            Call SetPlayerMP(Index, GetPlayerMP(Index) + Spell(SpellNum).Data1)
            Call SendMP(Index)
            
            Target = Index
            TargetType = TARGET_TYPE_PLAYER
        Case SPELL_TYPE_ADDSP
            Call SetPlayerSP(Index, GetPlayerSP(Index) + Spell(SpellNum).Data1)
            Call SendSP(Index)
            
            Target = Index
            TargetType = TARGET_TYPE_PLAYER
        Case SPELL_TYPE_SUBHP
            Target = GetPlayerTargetNpc(Index)

            ' Invalid target
            If Target <= 0 Then Exit Sub
            
            TargetType = TARGET_TYPE_NPC
            
            ' Armored Koopa can only be hurt if the player has the Turtle Badge equipped
            If MapNPC(MapNum, Target).num <> 194 Then
                Damage = ((GetPlayerSTR(Index) * 0.6) + PityFlower(Index, Spell(SpellNum).Data1)) - Int(NPC(MapNPC(MapNum, Target).num).DEF)
                Damage = Int(Rand(Damage - 2, Damage + 2))
            Else
                ' If the Turtle Badge isn't equipped, then deal 0 damage
                If GetPlayerEquipSlotNum(Index, 4) = 276 Then
                    Damage = ((GetPlayerSTR(Index) * 0.6) + PityFlower(Index, Spell(SpellNum).Data1)) - Int(NPC(MapNPC(MapNum, Target).num).DEF)
                    Damage = Int(Rand(Damage - 2, Damage + 2))
                Else
                    Damage = 0
                End If
            End If
            
            If Damage > 0 Then
                Call AttackNpc(Index, Target, Damage)
            Else
                Call BattleMsg(Index, "The special attack was too weak to harm " & Trim$(NPC(MapNPC(MapNum, Target).num).Name) & "!", BRIGHTRED, 0)
            End If
        Case SPELL_TYPE_SUBMP
            Target = GetPlayerTargetNpc(Index)
            TargetType = TARGET_TYPE_NPC
            
            ' Invalid target
            If Target <= 0 Then Exit Sub
            
            MapNPC(MapNum, Target).MP = MapNPC(MapNum, Target).MP - Spell(SpellNum).Data1
        Case SPELL_TYPE_SUBSP
            Target = GetPlayerTargetNpc(Index)
            TargetType = TARGET_TYPE_NPC
            
            ' Invalid target
            If Target <= 0 Then Exit Sub
            
            MapNPC(MapNum, Target).SP = MapNPC(MapNum, Target).SP - Spell(SpellNum).Data1
        Case SPELL_TYPE_STATCHANGE
            If SpellWorking <> "Yes" Then
                Call StartStatMod(Index, Spell(SpellNum).Multiplier, Spell(SpellNum).StatTime, Spell(SpellNum).Stat)
            Else
                Call PlayerMsg(Index, "You must wait for the effects of your special attack to end!", WHITE)
                
                Call SendCanUseSpecial(Index)
                Exit Sub
            End If
            
            Target = Index
            TargetType = TARGET_TYPE_PLAYER
        Case SPELL_TYPE_SCRIPTED
            If GetVar(App.Path & "\Scripts\" & "StatIncreases.ini", GetPlayerName(Index), "Attack") = "Yes" Or GetVar(App.Path & "\Scripts\" & "StatIncreases.ini", GetPlayerName(Index), "Defense") = "Yes" Then
                Call PlayerMsg(Index, "You cannot use this special attack when there is another special attack modifying your Attack or Defense!", BRIGHTRED)
                
                Call SendCanUseSpecial(Index)
                Exit Sub
            Else
                Dim ItemAttack As String, ItemDefense As String
                
                ' Check to see if any Attack or Defense modifying items are currently being used
                ItemAttack = GetVar(App.Path & "\Scripts\" & "StatIncreases.ini", GetPlayerName(Index), "ItemAttack")
                ItemDefense = GetVar(App.Path & "\Scripts\" & "StatIncreases.ini", GetPlayerName(Index), "ItemDefense")
            
                If ItemAttack = "Yes" Or ItemDefense = "Yes" Then
                    Call PlayerMsg(Index, "You would be too powerful if you used this Special Attack right now.", WHITE)
                    Exit Sub
                End If
            
                Call ScriptedSpell(Index, Spell(SpellNum).Data1)
            End If
            
            Target = Index
            TargetType = TARGET_TYPE_PLAYER
    End Select
    
    ' Take away the FP
    Call SetPlayerMP(Index, GetPlayerMP(Index) - FlowerSaver(Index, SpellNum))
    Call SendMP(Index)
    
    ' State that the player used the special attack
    Player(Index).CastedSpell = YES
    
    ' Send the animation
    Call SendDataTo(Index, SPackets.Sspellanim & SEP_CHAR & SpellNum & SEP_CHAR & Spell(SpellNum).SpellAnim & SEP_CHAR & Spell(SpellNum).SpellTime & SEP_CHAR & Spell(SpellNum).SpellDone & SEP_CHAR & Index & SEP_CHAR & TargetType & SEP_CHAR & Target & SEP_CHAR & YES & SEP_CHAR & Spell(SpellNum).Big & END_CHAR)
    
    ' Send the sound
    Call SendSoundTo(Index, Spell(SpellNum).Sound)
    
    ' Reset attack timer
    Player(Index).AttackTimer = GetTickCount
    
    ' Show the player attacking
    Call SendDataTo(Index, SPackets.Sattack & SEP_CHAR & Index & END_CHAR)
    
    MapNPC(MapNum, Target).Turn = True
    Call StartNpcTurn(Index)
End Sub

Function CanPlayerCriticalHit(ByVal Index As Long) As Boolean
    ' If Red Peppers are active, stop the player from getting a critical hit
    If GetVar(App.Path & "\Scripts\" & "StatIncreases.ini", GetPlayerName(Index), "RedPeppers") = "Yes" Then
        Exit Function
    End If

    If (Int(Rnd2 * 100) + 1) <= GetPlayerCritHitChance(Index) Then
        CanPlayerCriticalHit = True
    End If
End Function

Function CanPlayerBlockHit(ByVal Index As Long) As Boolean
    ' If Green Peppers are active, stop the player from getting a critical hit
    If GetVar(App.Path & "\Scripts\" & "StatIncreases.ini", GetPlayerName(Index), "GreenPeppers") = "Yes" Then
        Exit Function
    End If
    
    If (Int(Rnd2 * 100) + 1) <= GetPlayerBlockChance(Index) Then
        CanPlayerBlockHit = True
    End If
End Function

Public Sub CheckEquippedItems(ByVal Index As Long)
    Dim i As Long, ItemNum As Long
    Dim Change As Boolean
    
    Change = False
    
    ' Check to make sure the item exists
    For i = 1 To 7
        ItemNum = GetPlayerEquipSlotNum(Index, i)
    
        If ItemNum > 0 Then
            Select Case i
                ' Weapon
                Case 1
                    If Item(ItemNum).Type <> ITEM_TYPE_WEAPON Then
                        Call SetPlayerEquipSlotNum(Index, i, 0)
                        Call SetPlayerEquipSlotValue(Index, i, 0)
                        Call SetPlayerEquipSlotAmmo(Index, i, -1)
                        
                        Change = True
                    End If
                ' Special Badge
                Case 4
                    If Item(ItemNum).Type <> ITEM_TYPE_WEAPON And Item(ItemNum).Type <> i + 1 Then
                        Call SetPlayerEquipSlotNum(Index, i, 0)
                        Call SetPlayerEquipSlotValue(Index, i, 0)
                        Call SetPlayerEquipSlotAmmo(Index, i, -1)
                
                        Change = True
                    End If
                ' Everything else
                Case Else
                    If Item(ItemNum).Type <> i + 1 Then
                        Call SetPlayerEquipSlotNum(Index, i, 0)
                        Call SetPlayerEquipSlotValue(Index, i, 0)
                        Call SetPlayerEquipSlotAmmo(Index, i, -1)
                
                        Change = True
                    End If
            End Select
        End If
    Next i
    
    If Change = True Then
        Call SendWornEquipment(Index)
    End If
End Sub

Public Sub ShowPLR(ByVal Index As Long)
    Dim LS As ListItem

    On Error Resume Next

    If frmServer.lvUsers.ListItems.Count > 0 And IsPlaying(Index) Then
        frmServer.lvUsers.ListItems.Remove Index
    End If

    Set LS = frmServer.lvUsers.ListItems.Add(Index, , Index)

    If IsPlaying(Index) Then
        LS.SubItems(1) = GetPlayerLogin(Index)
        LS.SubItems(2) = GetPlayerName(Index)
        LS.SubItems(3) = GetPlayerLevel(Index)
        LS.SubItems(4) = GetPlayerSprite(Index)
        LS.SubItems(5) = GetPlayerAccess(Index)
    End If
End Sub

Public Sub RemovePLR(ByVal Index As Long)
    Dim LS As ListItem
    
    On Error Resume Next

    If Not IsPlaying(Index) Then
        frmServer.lvUsers.ListItems.Remove Index
    
        Set LS = frmServer.lvUsers.ListItems.Add(Index, , Index)
        
        LS.SubItems(1) = vbNullString
        LS.SubItems(2) = vbNullString
        LS.SubItems(3) = vbNullString
        LS.SubItems(4) = vbNullString
        LS.SubItems(5) = vbNullString
    End If
End Sub

Function CanAttackPlayerWithArrow(ByVal Attacker As Long, ByVal Victim As Long) As Boolean
    ' Make sure we dont attack the player if they are switching maps
    If Player(Victim).GettingMap = YES Then
        Exit Function
    End If
    
    ' Make sure the player isn't in a battle
    If GetPlayerInBattle(Victim) = True Then
        Exit Function
    End If
    
    ' If they are in STS, Dodgebill, or Hide n' Sneak then they can attack each other
    If GetPlayerMap(Attacker) = 33 Or GetPlayerMap(Attacker) = 188 Or GetPlayerMap(Attacker) = 271 Or GetPlayerMap(Attacker) = 272 Or GetPlayerMap(Attacker) = 273 Then
        CanAttackPlayerWithArrow = True
        Exit Function
    End If
    ' If they are on an arena tile, then they can be attacked
    If GetPlayerMap(Attacker) = 218 And Map(GetPlayerMap(Attacker)).Tile(GetPlayerX(Attacker), GetPlayerY(Attacker)).Type = TILE_TYPE_ARENA And Map(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).Type = TILE_TYPE_ARENA Then
        CanAttackPlayerWithArrow = True
        Exit Function
    End If
    ' Check if map is attackable
    If (Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_SAFE Or Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_MINIGAME) And GetPlayerPK(Victim) = NO Then
        Call PlayerMsg(Attacker, "This is not a PvP area!", BRIGHTRED)
        Exit Function
    End If
    ' Check if they are in a guild and if they are a pker
    If GetPlayerGuild(Attacker) <> vbNullString And GetPlayerGuild(Victim) <> vbNullString Then
        If GetPlayerGuild(Attacker) = GetPlayerGuild(Victim) Then
            Call PlayerMsg(Attacker, GetPlayerName(Victim) & " is in the same Group as you, so you cannot attack him/her!", BRIGHTRED)
            Exit Function
        End If
    End If
    CanAttackPlayerWithArrow = True
End Function

Function CanAttackNpcWithArrow(ByVal Attacker As Long, ByVal MapNpcNum As Long) As Boolean
    Dim i As Byte
    Dim MapNum As Long, NpcNum As Long
    
    CanAttackNpcWithArrow = False

    ' Check for subscript out of range
    If IsPlaying(Attacker) = False Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Then
        Exit Function
    End If
    
    MapNum = GetPlayerMap(Attacker)
    NpcNum = MapNPC(MapNum, MapNpcNum).num
    
    ' Check for subscript out of range
    If NpcNum = 0 Or NpcNum = 195 Then
        Exit Function
    End If
    
    ' Make sure the npc isn't in battle with another player
    If MapNPC(MapNum, MapNpcNum).InBattle = True And MapNPC(MapNum, MapNpcNum).Target <> Attacker Then
        Exit Function
    End If

    ' Make sure the npc isn't already dead
    If MapNPC(MapNum, MapNpcNum).HP <= 0 Then
        Exit Function
    End If
    
    ' Make sure the player hasn't just attacked
    If GetTickCount < Player(Attacker).AttackTimer + GetPlayerAttackSpeed(Attacker) Then
        Exit Function
    End If
    
    If NPC(NpcNum).Behavior = NPC_BEHAVIOR_SCRIPTED Then
        Call ScriptedNPC(Attacker, NpcNum, NPC(NpcNum).SpawnSecs)
        Exit Function
    ElseIf NPC(NpcNum).Behavior = NPC_BEHAVIOR_FRIENDLY Then
        Call SendNpcTalkTo(Attacker, NpcNum, NPC(NpcNum).AttackSay, NPC(NpcNum).AttackSay2)
        Exit Function
    End If
    
    CanAttackNpcWithArrow = True
End Function

Sub SendIndexWornEquipment(ByVal Index As Long)
    Dim i As Long
    Dim packet As String
    
    packet = SPackets.Sitemworn & SEP_CHAR & Index
    
    For i = 1 To 7
        packet = packet & SEP_CHAR & GetPlayerEquipSlotNum(Index, i)
    Next i
    
    packet = packet & END_CHAR

    Call SendDataToMap(GetPlayerMap(Index), packet)
End Sub

Sub SendIndexWornEquipmentTo(ByVal Index As Long, ByVal From As Long)
    Dim i As Long
    Dim packet As String
    
    packet = SPackets.Sitemworn & SEP_CHAR & From
    
    For i = 1 To 7
        packet = packet & SEP_CHAR & GetPlayerEquipSlotNum(From, i)
    Next i
    
    packet = packet & END_CHAR

    Call SendDataTo(Index, packet)
End Sub

Sub SendIndexWornEquipmentFromMap(ByVal Index As Long)
    Dim packet As String
    Dim i As Long, X As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            If GetPlayerMap(i) = GetPlayerMap(Index) Then
                packet = SPackets.Sitemworn & SEP_CHAR & i
                
                For X = 1 To 7
                    packet = packet & SEP_CHAR & GetPlayerEquipSlotNum(i, X)
                Next X
                
                packet = packet & END_CHAR

                Call SendDataTo(Index, packet)
            End If
        End If
    Next i
End Sub

Sub AddNewTimer(ByVal Index As Long, ByVal Number As Long, ByVal Interval As Long, Optional ByVal Parameter1 As Long = 0, Optional ByVal Parameter2 As Long = 0, Optional ByVal Parameter3 As Long = 0)
    Dim i As Long
    
    For i = 1 To MAX_PLAYERS
        If LenB(Timer(i).Player) = 0 Then
            Timer(i).Index = Index
            Timer(i).Player = GetPlayerName(Index)
            Timer(i).num = Number
            Timer(i).Interval = Interval
            Timer(i).WaitTime = GetTickCount + Interval
            Timer(i).Parameter1 = Parameter1
            Timer(i).Parameter2 = Parameter2
            Timer(i).Parameter3 = Parameter3
            Exit Sub
        End If
    Next i
End Sub

Sub GetRidOfTimer(ByVal Index As Long, ByVal TimerNum As Long, Optional ByVal Parameter As Long = -1)
    Dim i As Long
    
    If Index < 1 Then
        Exit Sub
    End If
    
    If Parameter <> -1 Then
        For i = 1 To MAX_PLAYERS
            If Timer(i).Index = Index And Timer(i).num = TimerNum And Timer(i).Parameter1 = Parameter Then
                Timer(i).Index = 0
                Timer(i).Player = vbNullString
                Timer(i).num = 0
                Timer(i).Interval = 0
                Timer(i).WaitTime = 0
                Timer(i).Parameter1 = 0
                Timer(i).Parameter2 = 0
                Timer(i).Parameter3 = 0
                Exit Sub
            End If
        Next i
    Else
        For i = 1 To MAX_PLAYERS
            If Timer(i).Index = Index And Timer(i).num = TimerNum Then
                Timer(i).Index = 0
                Timer(i).Player = vbNullString
                Timer(i).num = 0
                Timer(i).Interval = 0
                Timer(i).WaitTime = 0
                Timer(i).Parameter1 = 0
                Timer(i).Parameter2 = 0
                Timer(i).Parameter3 = 0
                Exit Sub
            End If
        Next i
    End If
End Sub

Function ItemIsUsable(ByVal Index As Long, ByVal InvNum As Long) As Boolean
    Dim PlayerBaseHP As Long, PlayerBaseFP As Long, LvlUpHP As Long, LvlUpFP As Long, ItemNum As Long
    
    LvlUpHP = Val(GetVar(App.Path & "\Level Up.ini", GetPlayerName(Index), "HP")) * 5
    LvlUpFP = Val(GetVar(App.Path & "\Level Up.ini", GetPlayerName(Index), "FP")) * 5
    
 ' Sets player base HP
    If GetPlayerClass(Index) = 1 Or GetPlayerClass(Index) = 3 Or GetPlayerClass(Index) = 5 Then
        PlayerBaseHP = 15 + LvlUpHP
    Else
        PlayerBaseHP = 10 + LvlUpHP
    End If
    
 ' Sets player base FP
    If GetPlayerClass(Index) = 4 Then
        PlayerBaseFP = 10 + LvlUpFP
    Else
        PlayerBaseFP = 5 + LvlUpFP
    End If
    
    ItemNum = GetPlayerInvItemNum(Index, InvNum)
    
    ' Check if the player meets the class requirement.
    If Item(ItemNum).ClassReq > -1 Then
        If GetPlayerClass(Index) <> Item(ItemNum).ClassReq Then
            Call PlayerMsg(Index, "You must be " & GetClassName(Item(ItemNum).ClassReq) & " to use this item!", BRIGHTRED)
            Exit Function
        End If
    End If

    ' Check if the player meets the access requirement.
    If GetPlayerAccess(Index) < Item(ItemNum).AccessReq Then
        Call PlayerMsg(Index, "Your access must be higher than " & Item(ItemNum).AccessReq & "!", BRIGHTRED)
        Exit Function
    End If

    ' Check if the player meets the strength requirement.
    If Player(Index).Char(Player(Index).CharNum).STR < Item(ItemNum).StrReq Then
        Call PlayerMsg(Index, "Your base Attack is too low to equip this item!", BRIGHTRED)
        Exit Function
    End If

    ' Check if the player meets the defense requirement.
    If Player(Index).Char(Player(Index).CharNum).DEF < Item(ItemNum).DefReq Then
        Call PlayerMsg(Index, "Your base Defense is too low to equip this item!", BRIGHTRED)
        Exit Function
    End If

    ' Check if the player meets the stache requirement.
    If Player(Index).Char(Player(Index).CharNum).Magi < Item(ItemNum).MagicReq Then
        Call PlayerMsg(Index, "Your base Stache is too low to equip this item!", BRIGHTRED)
        Exit Function
    End If

    ' Check if the player meets the speed requirement.
    If Player(Index).Char(Player(Index).CharNum).Speed < Item(ItemNum).SpeedReq Then
        Call PlayerMsg(Index, "Your base Speed is too low to equip this item!", BRIGHTRED)
        Exit Function
    End If
    
    ' Check if the player meets the level requirement.
    If GetPlayerLevel(Index) < Item(ItemNum).LevelReq Then
        Call PlayerMsg(Index, "You need to be at level " & Item(ItemNum).LevelReq & " to equip this item!", BRIGHTRED)
        Exit Function
    End If
    
    ' Check if the player meets the HP requirement.
    If PlayerBaseHP < Item(ItemNum).HPReq Then
        Call PlayerMsg(Index, "You need to have at least " & Item(ItemNum).HPReq & " base HP to equip this item!", BRIGHTRED)
        Exit Function
    End If
    
    ' Check if the player meets the FP requirement.
    If PlayerBaseFP < Item(ItemNum).FPReq Then
        Call PlayerMsg(Index, "You need to have at least " & Item(ItemNum).FPReq & " base FP to equip this item!", BRIGHTRED)
        Exit Function
    End If
    
    ItemIsUsable = True
End Function

Sub SpellAnim(ByVal SpellNum As Long, ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long, Optional ByVal Index As Long = 0)
    Call SendDataToMap(MapNum, SPackets.Sscriptspellanim & SEP_CHAR & SpellNum & SEP_CHAR & Spell(SpellNum).SpellAnim & SEP_CHAR & Spell(SpellNum).SpellTime & SEP_CHAR & Spell(SpellNum).SpellDone & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & Spell(SpellNum).Big & SEP_CHAR & Index & END_CHAR)
End Sub

Sub OnAttack(ByVal Index As Long, ByVal Damage As Long)
    Dim Target As Long
    Dim IndexTeam As Byte, TargetTeam As Byte
 
    If Player(Index).TargetType = TARGET_TYPE_PLAYER Then
        Target = Player(Index).Target
        
        ' Ignore attack if we don't have a target
        If Target <= 0 Then
            Exit Sub
        End If
        
        ' Handle minigame-specific events
        If Map(GetPlayerMap(Index)).Moral = MAP_MORAL_MINIGAME Then
            Dim PlayerMinigamePath As String
            
            PlayerMinigamePath = FindPlayerMinigame(Index)
            
            IndexTeam = GetPlayerTeam(Index, PlayerMinigamePath, DodgeBillMaxPlayers)
            TargetTeam = GetPlayerTeam(Target, PlayerMinigamePath, DodgeBillMaxPlayers)
            
            ' Stop players from harming their teammate in a minigame
            If IndexTeam = TargetTeam And IndexTeam <> 2 Then
                Call PlayerMsg(Index, "You cannot harm your teammate!", BRIGHTRED)
                Exit Sub
            End If
            
            ' Kills players in Steal the Shroom in 1 hit
            If GetPlayerMap(Index) = 33 And IndexTeam <> TargetTeam Then
                Call STSDeath(Index, Target)
                Exit Sub
            End If
            
            ' Gets players out in Hide n' Sneak
            If GetPlayerMap(Index) >= 271 And GetPlayerMap(Index) <= 273 Then
                Call HideNSneakOut(Index, Target)
                Exit Sub
            End If
        End If
        
        ' Stops players from harming another player in a waiting room
        If GetPlayerMap(Index) = 32 Or GetPlayerMap(Index) = 191 Or GetPlayerMap(Index) = 270 Then
            Call PlayerMsg(Index, "You cannot fight in a waiting room!", BRIGHTRED)
            Exit Sub
        End If
        
        Call AttackPlayer(Index, Target, Damage)
    ElseIf Player(Index).TargetType = TARGET_TYPE_NPC Then
        Target = Player(Index).TargetNPC
        
        ' Ignore attack if we don't have a target
        If Target <= 0 Then
            Exit Sub
        End If
        
        ' Kills Monty Moles in 1 hit in Whack-A-Monty
        If GetPlayerMap(Index) = 72 Or GetPlayerMap(Index) = 73 Or GetPlayerMap(Index) = 74 Then
            Call AttackNpc(Index, Target, 999)
            Exit Sub
        End If
        
        Call AttackNpc(Index, Target, Damage)
    End If
End Sub

Sub OnArrowHit(ByVal Index As Long, ByVal Damage As Long)
    Dim Target As Long, MapItemSlot As Long, MapItemNum As Long
    Dim IndexTeam As Byte, TargetTeam As Byte
    
    If Player(Index).TargetType = TARGET_TYPE_PLAYER Then
        Target = Player(Index).Target
        
        ' Ignore attack if we don't have a target
        If Target <= 0 Then
            Exit Sub
        End If
               
        ' Handle minigame-specific events
        If Map(GetPlayerMap(Index)).Moral = MAP_MORAL_MINIGAME Then
            Dim PlayerMinigamePath As String
            
            PlayerMinigamePath = FindPlayerMinigame(Index)
            
            IndexTeam = GetPlayerTeam(Index, PlayerMinigamePath, DodgeBillMaxPlayers)
            TargetTeam = GetPlayerTeam(Target, PlayerMinigamePath, DodgeBillMaxPlayers)
            
            ' Stop players from harming their teammate in a minigame
            If IndexTeam = TargetTeam And IndexTeam <> 2 Then
                ' Only spawn the item in Dodgebill
                If GetPlayerMap(Index) = 188 Then
                    ' Spawn the item
                    MapItemNum = FindOpenMapItemSlot(188)
                    Call SpawnItemSlot(MapItemNum, 186, 1, 1, 188, GetPlayerX(Target), GetPlayerY(Target))
                    
                    Exit Sub
                End If
                
                ' Notify the user that he/she cannot harm a teammate in STS
                If GetPlayerMap(Index) = 33 Then
                    Call PlayerMsg(Index, "You cannot harm your teammate!", BRIGHTRED)
                    Exit Sub
                End If
            End If

            If IndexTeam <> TargetTeam Then
                ' Kills players in Steal the Shroom in 1 hit
                If GetPlayerMap(Index) = 33 Then
                    Call STSDeath(Index, Target)
                    Exit Sub
                End If
        
                ' Kills players in Dodgebill in 1 hit
                If GetPlayerMap(Index) = 188 Then
                    Call DodgeBillDeath(Index, Target)
                    Exit Sub
                End If
                
                ' Gets players out in Hide n' Sneak
                If GetPlayerMap(Index) >= 271 And GetPlayerMap(Index) <= 273 Then
                    Call HideNSneakOut(Index, Target)
                    Exit Sub
                End If
            End If
        End If
        
        ' Stops players from harming another player in a waiting room
        If GetPlayerMap(Index) = 32 Or GetPlayerMap(Index) = 191 Or GetPlayerMap(Index) = 270 Then
            Call PlayerMsg(Index, "You cannot fight in a waiting room!", BRIGHTRED)
            Exit Sub
        End If
        
        Call AttackPlayer(Index, Target, Damage)
    ElseIf Player(Index).TargetType = TARGET_TYPE_NPC Then
        Target = Player(Index).TargetNPC
        
        ' Ignore attack if we don't have a target
        If Target <= 0 Then
            Exit Sub
        End If
        
        ' Kills Monty Moles in 1 hit in Whack-A-Monty
        If GetPlayerMap(Index) = 72 Or GetPlayerMap(Index) = 73 Or GetPlayerMap(Index) = 74 Then
            Call AttackNpc(Index, Target, 999)
            Exit Sub
        End If
        
        Call AttackNpc(Index, Target, Damage)
    End If
End Sub

Sub StartStatMod(ByVal Index As Long, ByVal Multiplier As Double, ByVal Time As Long, ByVal Stat As Integer)
    Call PutVar(App.Path & "\Scripts\" & "StatIncreases.ini", GetPlayerName(Index), Trim$(GetSpellStatName(Stat)), "Yes")
    
    Select Case Stat
    ' Attack
       Case 0
           Call SetPlayerSTR(Index, (Player(Index).Char(Player(Index).CharNum).STR * Multiplier))
    ' Defense
       Case 1
           Call SetPlayerDEF(Index, (Player(Index).Char(Player(Index).CharNum).DEF * Multiplier))
    ' Speed
       Case 2
           Call SetPlayerSPEED(Index, (Player(Index).Char(Player(Index).CharNum).Speed * Multiplier))
    ' Stache
       Case 3
           Call SetPlayerStache(Index, (Player(Index).Char(Player(Index).CharNum).Magi * Multiplier))
     End Select
     
    Call AddNewTimer(Index, 1, (Time * 1000), Stat)
    Call SendStats(Index)
End Sub

Function GetSpellStatName(ByVal Stat As Integer) As String
    Select Case Stat
        Case 0
            GetSpellStatName = "Attack"
        Case 1
            GetSpellStatName = "Defense"
        Case 2
            GetSpellStatName = "Speed"
        Case 3
            GetSpellStatName = "Stache"
        End Select
End Function

' Executes whenever somebody dies outside of an arena
Sub OnDeath(ByVal Index As Long)
    Dim Map As Long
    Dim X As Long
    Dim Y As Long
    
    If GetVar(App.Path & "\Respawn Points.ini", GetPlayerName(Index), "Map #") <> vbNullString Then
        Map = GetVar(App.Path & "\Respawn Points.ini", GetPlayerName(Index), "Map #")
        X = GetVar(App.Path & "\Respawn Points.ini", GetPlayerName(Index), "X")
        Y = GetVar(App.Path & "\Respawn Points.ini", GetPlayerName(Index), "Y")

        Call PlayerWarp(Index, Map, X, Y)
        Call SendPlayerData(Index)
        Call SendPlayerXY(Index)
    Else
        Call PlayerWarp(Index, 1, 7, 6)
        Call SendPlayerData(Index)
        Call SendPlayerXY(Index)
    End If
End Sub

Sub OnPVPDeath(ByVal Attacker As Long, ByVal Victim As Long)
    Call GlobalMsg(GetPlayerName(Victim) & " has been defeated by " & GetPlayerName(Attacker) & "!", BRIGHTRED)
End Sub

Sub OnNPCDeath(ByVal Index As Long, ByVal Map As Long, ByVal NpcNum As Long, ByVal NPCIndex As Long)
    Dim a As Integer
    
    ' Monty Mole kills for Whack-A-Monty minigame
    If NpcNum = 47 Or NpcNum = 48 Then
        For a = 1 To 3
            If Trim$(GetVar(App.Path & "\Scripts\" & "Whack.ini", "Game" & Int(a), "Player")) = GetPlayerName(Index) Then
                Call KillMonty(Index, a, NpcNum)
                Exit Sub
            End If
        Next a
    End If
    
    ' Lakitu kills for First Kill Quest
    If GetVar(App.Path & "\Scripts\Quests.ini", GetPlayerName(Index), "KillQuest1") = "InProgress" Then
        a = CInt(GetVar(App.Path & "\Scripts\Quests.ini", GetPlayerName(Index), "LakitusKilled"))
        If NpcNum = 22 And a < 10 Then
            Call PutVar(App.Path & "\Scripts\Quests.ini", GetPlayerName(Index), "LakitusKilled", Int(a) + 1)
        End If
    End If
End Sub

Sub SendGuildMemberHP(ByVal Index As Long)
    If IsPlaying(Index) = False Then
        Exit Sub
    End If
        
    If LenB(GetPlayerGuild(Index)) < 1 Then
        Exit Sub
    End If
    
    Dim i As Long
    
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            If GetPlayerGuild(i) = GetPlayerGuild(Index) And GetPlayerMap(i) = GetPlayerMap(Index) Then
                Call SendDataTo(i, SPackets.Splayerhp & SEP_CHAR & Index & SEP_CHAR & GetPlayerMaxHP(Index) & SEP_CHAR & GetPlayerHP(Index) & END_CHAR)
            End If
        End If
    Next i
End Sub

Function FindItemVowels(ByVal ItemNum As Long) As Boolean
    Dim Beginning As String * 1
    
    Beginning = LCase$(Left$(Trim$(Item(ItemNum).Name), 1))
    
    If Beginning Like "[aeiou]" Then
        FindItemVowels = True
    Else
        FindItemVowels = False
    End If
End Function

Function FindNpcVowels(ByVal NpcNum As Long) As Boolean
    Dim Beginning As String * 1
    
    Beginning = LCase$(Left$(Trim$(NPC(NpcNum).Name), 1))
    
    If Beginning Like "[aeiou]" Then
        FindNpcVowels = True
    Else
        FindNpcVowels = False
    End If
End Function

Sub SendCardShop(ByVal Index As Long)
    Dim i As Long
    Dim packet As String
    
    packet = SPackets.Scardshop
    
    For i = 94 To MAX_ITEMS
        If Trim$(GetVar(App.Path & "\Scripts\" & "Cards.ini", CStr(i), GetPlayerName(Index))) = "Has" Then
            packet = packet & SEP_CHAR & Item(i).Name
        Else
            packet = packet & SEP_CHAR & vbNullString
        End If
    Next i
    
    packet = packet & END_CHAR
    
    Call SendDataTo(Index, packet)
End Sub

Sub SendRecipeLog(ByVal Index As Long)
    Dim i As Long
    Dim packet As String
    
    packet = SPackets.Srecipelog
    
    For i = 1 To MAX_RECIPES
        If Recipe(i).ResultItem > 0 Then
            If Trim$(GetVar(App.Path & "\Scripts\" & "Recipes.ini", CInt(i), GetPlayerName(Index))) = "Has" Then
                packet = packet & SEP_CHAR & Recipe(i).Name
            Else
                packet = packet & SEP_CHAR & vbNullString
            End If
        End If
    Next i
    
    packet = packet & END_CHAR
    
    Call SendDataTo(Index, packet)
End Sub

Sub WhackAMontyTime(ByVal Index As Long, ByVal GameNum As Integer)
    Dim TimeLeft As Integer, WAMPoints As Integer, Tokens As Integer
    Dim MapNum As Long, MapNpcIndex As Long
    
    Call GetRidOfTimer(Index, 7, GameNum)
    
    If IsConnected(Index) = False Or IsPlaying(Index) = False Then
        Exit Sub
    End If
    
    MapNum = GetPlayerMap(Index)
    TimeLeft = CInt(GetVar(App.Path & "\Scripts\" & "Whack.ini", "Game" & GameNum, "TimeLeft")) - 10
    
    ' Continues game
    If TimeLeft > 0 Then
        Call PutVar(App.Path & "\Scripts\" & "Whack.ini", "Game" & GameNum, "TimeLeft", CStr(TimeLeft))
        Call BattleMsg(Index, "Time remaining: " & TimeLeft & " seconds", WHITE, 1)
        Call AddNewTimer(Index, 7, 15800, GameNum)
    ' Stops game when time is up
    ElseIf TimeLeft <= 0 Then
        WAMPoints = CInt(GetVar(App.Path & "\Scripts\" & "Whack.ini", "Game" & GameNum, "Points"))
     
        ' Clears Map
        For MapNpcIndex = 1 To 15
            Call ClearMapNpc(MapNpcIndex, MapNum)
        Next MapNpcIndex
     
        If WAMPoints < 0 Then
            Tokens = 0
        Else
            Tokens = WAMPoints
        End If
     
        If Tokens > 0 Then
            Call GiveItem(Index, 87, Tokens)
        End If
     
        Call SetPlayerSP(Index, GetPlayerMaxSP(Index))
        Call SendSP(Index)
        Call PlayerWarp(Index, 71, 25, 11)
        
        Call PlayerMsg(Index, "Thanks for playing Whack-A-Monty! Your score for this round was: " & WAMPoints & ". You've earned " & Tokens & " Whack-A-Monty Reward Coins!", YELLOW)
    
        ' Checks to see if the player got enough points to equip Power Rush and Damage Dodge
        If WAMPoints >= 50 Then
            If Val(GetVar(App.Path & "\Scripts\" & "WhackFame.ini", GetPlayerName(Index), "Points")) < 50 Then
                Call PutVar(App.Path & "\Scripts\" & "WhackFame.ini", GetPlayerName(Index), "Points", CStr(WAMPoints))
                Call PlayerMsg(Index, "Congratulations! You've earned at least 50 points in Whack-A-Monty!", YELLOW)
            End If
        End If
    
        ' Checks to see if the player got a high enough score to get a special item
        If WAMPoints >= 80 Then
            If Trim$(GetVar(App.Path & "\Scripts\" & "WhackFame.ini", GetPlayerName(Index), "In")) <> "Yes" Then
                Call PutVar(App.Path & "\Scripts\" & "WhackFame.ini", GetPlayerName(Index), "In", "Yes")
                Call PlayerMsg(Index, "Congratulations! You've earned at least 85 points in Whack-A-Monty!", YELLOW)
            End If
        End If

        ' Resets game
        Call PutVar(App.Path & "\Scripts\" & "Whack.ini", "Game" & GameNum, "TimeLeft", "0")
        Call PutVar(App.Path & "\Scripts\" & "Whack.ini", "Game" & GameNum, "InGame", "No")
        Call PutVar(App.Path & "\Scripts\" & "Whack.ini", "Game" & GameNum, "Points", "0")
        Call PutVar(App.Path & "\Scripts\" & "Whack.ini", "Game" & GameNum, "Player", vbNullString)
    End If
End Sub

Sub MontyRespawn(ByVal Index As Long, ByVal GameNum As Integer, MapNum As Long)
    Dim CanSpawnNpc As Boolean
    Dim MoleNumber As Integer, Randomize As Integer
    Dim MapNpcIndex As Long, NpcNum As Long
    Dim X As Byte, Y As Byte
    
    Call GetRidOfTimer(Index, 8, GameNum)
    
    If IsConnected(Index) = False Or IsPlaying(Index) = False Then
        Exit Sub
    End If
  
    If CInt(GetVar(App.Path & "\Scripts\" & "Whack.ini", "Game" & GameNum, "TimeLeft")) <= 0 Then
        Exit Sub
    End If
 
    For MapNpcIndex = 1 To MAX_MAP_NPCS
        Call ClearMapNpc(MapNpcIndex, MapNum)
    Next MapNpcIndex
    
    MoleNumber = Int(Rand(6, 11))
    
    For MapNpcIndex = 1 To MoleNumber
        CanSpawnNpc = False
        Do
            Randomize = Int(Rand(1, 9))
    
            Select Case Randomize
                Case 1
                    X = 8
                Case 2
                    X = 11
                Case 3
                    X = 14
                Case 4
                    X = 15
                Case 5
                    X = 17
                Case 6
                    X = 19
                Case 7
                    X = 20
                Case 8
                    X = 23
                Case 9
                    X = 26
            End Select
    
            Y = (Int(Rand(9, 12)) * 2)
      
            If X = 15 Or X = 19 Then
                Y = 24
            End If
      
            If Y = 24 Then
                Randomize = Int(Rand(1, 2))
         
                If Randomize = 1 Then
                    X = 15
                Else
                    X = 19
                End If
            End If
    
            Randomize = Int(Rand(1, 4))
                Select Case Randomize
                    Case 1
                        NpcNum = 47
                    Case 2
                        NpcNum = 47
                    Case 3
                        NpcNum = 47
                    Case 4
                        NpcNum = 48
                End Select
            
            If MapNpcSpotOpen(MapNum, X, Y) = True Then
                Call ScriptSpawnNpc(MapNpcIndex, MapNum, X, Y, NpcNum)
                CanSpawnNpc = True
            End If
        Loop While CanSpawnNpc = False
    Next MapNpcIndex
  
    Call SendMapNpcsToMap(MapNum)
    Call AddNewTimer(Index, 8, (Int(Rand(7, 11)) * 1000), GameNum, MapNum)
End Sub

Function MapNpcSpotOpen(ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long) As Boolean
    Dim MapNpcIndex As Long
    
    MapNpcSpotOpen = True
    
    For MapNpcIndex = 1 To MAX_MAP_NPCS
        If X = MapNPC(MapNum, MapNpcIndex).X And Y = MapNPC(MapNum, MapNpcIndex).Y Then
            MapNpcSpotOpen = False
            Exit Function
        End If
    Next MapNpcIndex
End Function

' Executes every second, based on the server time.
Sub TimedEvent(Hours, Minutes, Seconds)
    Dim i As Long
    
    ' Daily Event
    If Hours = 12 And Minutes = 0 And Seconds = 0 Then
        For i = 1 To MAX_PLAYERS
            If IsPlaying(i) Then
                If GetPlayerLevel(i) > 3 Then
                    If Map(GetPlayerMap(i)).Moral = MAP_MORAL_MINIGAME Or GetPlayerInBattle(i) = False Then
                        Player(i).GetsDE = True
                    ElseIf Map(GetPlayerMap(i)).Moral <> MAP_MORAL_MINIGAME And GetPlayerInBattle(i) = False Then
                        Call Prompt(i, "It's time for the daily event. Would you like to start it? Note: You will be teleported to another area.", 0)
                    End If
                End If
            End If
        Next i
    End If
End Sub

Sub CheckForDE(ByVal Index As Long)
    If IsPlaying(Index) Then
        If Player(Index).GetsDE = True Then
            If Map(GetPlayerMap(Index)).Moral <> MAP_MORAL_MINIGAME And GetPlayerInBattle(Index) = False Then
                Call Prompt(Index, "It's time for the daily event. Would you like to start it? Note: You will be teleported to another area.", 0)
                
                Player(Index).GetsDE = False
            End If
        End If
    End If
End Sub

' Executes when a player steps onto a scripted tile.
Sub ScriptedTile(ByVal Index As Long, ByVal ScriptNum As Long)
    Dim PlayerNum As String, FilePath As String
    Dim Team As Byte, TeamNumberRed As Byte, TeamNumberBlue As Byte, Randomize As Byte, POINTS As Byte, BluePlayerCount As Byte, RedPlayerCount As Byte
    Dim i As Long, Q As Long, R As Long
    
    Select Case ScriptNum
   ' Reduce HP to 1 (Daily Event)
        Case 0
            If GetPlayerHP(Index) > 1 Then
                Call SetPlayerHP(Index, 1)
                Call SendHP(Index)
            End If
        Exit Sub
   ' Daily Event Rewards
        Case 1
            If GetPlayerMap(Index) = 35 Then
                Call GiveItem(Index, 54, 5)
            Else
                Call GiveItem(Index, 54, 3)
            End If
                
            Call PlayerWarp(Index, 11, 2, 8)
            Call PlayerMsg(Index, "Congratulations! You've completed the Daily Event for today! Enjoy your reward!", YELLOW)
            Call SetPlayerHP(Index, GetPlayerMaxHP(Index))
            Call SendHP(Index)
        Exit Sub
   ' Save Blocks
        Case 2
            Call PutVar(App.Path & "\Respawn Points.ini", GetPlayerName(Index), "CTRL", "1")
        Exit Sub
   ' Warp and play sound (in tutorial)
        Case 3
            Call SendSoundTo(Index, "pm_Long Fall.wav")
            Call PlayerWarp(Index, 6, 9, 10)
                
            Call PutVar(App.Path & "\Respawn Points.ini", GetPlayerName(Index), "Map #", "7")
            Call PutVar(App.Path & "\Respawn Points.ini", GetPlayerName(Index), "X", "10")
            Call PutVar(App.Path & "\Respawn Points.ini", GetPlayerName(Index), "Y", "5")
        Exit Sub
    ' Blocking player for Mushroom Toy Quest
        Case 4
            If GetVar(App.Path & "\Scripts\Quests.ini", GetPlayerName(Index), "ItemQuest2") <> "Done" Then
                Call PlayerMsg(Index, "You feel that you should help out the Toad before moving on...", WHITE)
                Call BlockPlayer(Index)
            End If
        Exit Sub
    ' Entering Steal the Shroom rooms
        Case 5
            FilePath = STSPath
            ' Red Team door
            If GetPlayerMap(Index) = 31 And GetPlayerX(Index) = 22 And GetPlayerY(Index) = 12 And HasItem(Index, 49) >= 1 Then
                If GetVar(FilePath, "Team", "Red") = "4" Then
                    Call PlayerMsg(Index, "The red team is currently full!", WHITE)
                    Exit Sub
                ElseIf GetVar(FilePath, "Ingame", "Ingame") = "Yes" Then
                    Call PlayerMsg(Index, "There's already a game in progress!", WHITE)
                    Exit Sub
                Else
                    Call SetPlayerTeam(Index, 0, STSPath, STSMaxPlayers)
                    Call TakeItem(Index, 49, 1)
                    Call PlayerWarp(Index, 32, 23, 17)
                End If
            ' Random door
            ElseIf GetPlayerMap(Index) = 31 And GetPlayerX(Index) = 20 And GetPlayerY(Index) = 12 And HasItem(Index, 50) >= 1 Then
                If GetVar(FilePath, "Team", "Red") = "4" And GetVar(FilePath, "Team", "Blue") = "4" Then
                    Call PlayerMsg(Index, "All teams are currently full!", WHITE)
                    Exit Sub
                ElseIf GetVar(FilePath, "Ingame", "Ingame") = "Yes" Then
                    Call PlayerMsg(Index, "There's already a game in progress!", WHITE)
                    Exit Sub
                Else
                    ' Checks number of team members for the Blue Team
                    If GetVar(FilePath, "Team", "Blue") = vbNullString Then
                        TeamNumberBlue = 0
                    Else
                        TeamNumberBlue = CByte(GetVar(FilePath, "Team", "Blue"))
                    End If
                    ' Checks number of team members for the Red Team
                    If GetVar(FilePath, "Team", "Red") = vbNullString Then
                        TeamNumberRed = 0
                    Else
                        TeamNumberRed = CByte(GetVar(FilePath, "Team", "Red"))
                    End If
                    
                    If HasItem(Index, 50) >= 1 Then
                        Team = Int(Rand(1, 2))
                        ' Sets Teams
                        If Team = 1 Then
                            Call SetPlayerTeam(Index, 0, STSPath, STSMaxPlayers)
                                
                            If TeamNumberRed >= 4 And TeamNumberBlue < 4 Then
                                Call PlayerWarp(Index, 32, 7, 17)
                            Else
                                Call PlayerWarp(Index, 32, 23, 17)
                            End If
                        ElseIf Team = 2 Then
                            Call SetPlayerTeam(Index, 1, STSPath, STSMaxPlayers)
                            
                            If TeamNumberBlue >= 4 And TeamNumberRed < 4 Then
                                Call PlayerWarp(Index, 32, 23, 17)
                            Else
                                Call PlayerWarp(Index, 32, 7, 17)
                            End If
                        End If
                        
                        Call TakeItem(Index, 50, 1)
                    End If
                End If
            ' Blue Team door
            ElseIf GetPlayerMap(Index) = 31 And GetPlayerX(Index) = 18 And GetPlayerY(Index) = 12 And HasItem(Index, 48) >= 1 Then
                If GetVar(FilePath, "Team", "Blue") = "4" Then
                    Call PlayerMsg(Index, "The blue team is currently full!", WHITE)
                    Exit Sub
                ElseIf GetVar(FilePath, "Ingame", "Ingame") = "Yes" Then
                    Call PlayerMsg(Index, "There's already a game in progress!", WHITE)
                    Exit Sub
                Else
                    Call SetPlayerTeam(Index, 1, STSPath, STSMaxPlayers)
                    Call TakeItem(Index, 48, 1)
                    Call PlayerWarp(Index, 32, 7, 17)
                End If
            End If
        Exit Sub
    ' Random Warping (For Mushroom PvP Area)
        Case 6
            Randomize = Int(Rand(1, 5))
            Select Case Randomize
                Case 1
                    Call PlayerWarp(Index, 21, 2, 8)
                Case 2
                    Call PlayerWarp(Index, 21, 16, 26)
                Case 3
                    Call PlayerWarp(Index, 21, 23, 12)
                Case 4
                    Call PlayerWarp(Index, 21, 15, 12)
                Case 5
                    Call PlayerWarp(Index, 21, 5, 28)
            End Select
        Exit Sub
    ' Heart Blocks
        Case 7
            If GetVar(App.Path & "\Heart Blocks.ini", GetPlayerName(Index), "CTRL") <> "1" Then
                Call PutVar(App.Path & "\Heart Blocks.ini", GetPlayerName(Index), "CTRL", "1")
            End If
        Exit Sub
    ' Exiting STS Waiting Rooms
        Case 8
            FilePath = STSPath
            If GetPlayerMap(Index) = 32 And GetPlayerX(Index) = 23 And GetPlayerY(Index) = 22 Then
                TeamNumberRed = CByte(GetVar(FilePath, "Team", "Red"))
                For i = 1 To 4
                    PlayerNum = CStr(i)
                    If GetVar(FilePath, "Red", PlayerNum) = GetPlayerName(Index) Then
                        Call PutVar(FilePath, "Team", "Red", Int(TeamNumberRed) - 1)
                        Call PutVar(FilePath, "Red", PlayerNum, "")
                        Call PlayerWarp(Index, 31, 22, 13)
                    End If
                Next i
            ElseIf GetPlayerMap(Index) = 32 And GetPlayerX(Index) = 7 And GetPlayerY(Index) = 22 Then
                TeamNumberBlue = CByte(GetVar(FilePath, "Team", "Blue"))
                For i = 1 To 4
                    PlayerNum = CStr(i)
                    If GetVar(FilePath, "Blue", PlayerNum) = GetPlayerName(Index) Then
                        Call PutVar(FilePath, "Team", "Blue", Int(TeamNumberBlue) - 1)
                        Call PutVar(FilePath, "Blue", PlayerNum, "")
                        Call PlayerWarp(Index, 31, 18, 13)
                    End If
                Next i
            End If
        Exit Sub
    ' Stealing Shroom & Scoring Points
        Case 9
            FilePath = STSPath
            ' Stealing for Blue
            If GetPlayerTeam(Index, FilePath, STSMaxPlayers) = 1 And GetVar(FilePath, "Flag", "Blue") <> "HasFlag" And GetPlayerMap(Index) = 33 And GetPlayerX(Index) = 15 And GetPlayerY(Index) = 1 Then
                Call PutVar(FilePath, "Flag", "Blue", "HasFlag")
                    
                Call PlayerMsg(Index, "You stole the enemy Shroom!", MAGENTA)
                    
                For i = 1 To 4
                    PlayerNum = CStr(i)
                    R = FindPlayer(GetVar(FilePath, "Blue", PlayerNum))
                    Q = FindPlayer(GetVar(FilePath, "Red", PlayerNum))
                    
                    If IsPlaying(Q) Then
                        Call PlayerMsg(Q, GetPlayerName(Index) & " has stolen your Shroom!", MAGENTA)
                    End If
                    
                    If IsPlaying(R) Then
                        If R <> Index Then
                            Call PlayerMsg(R, GetPlayerName(Index) & " has stolen the enemy Shroom!", MAGENTA)
                        End If
                    End If
                Next i
                
                Call SetPlayerPK(Index, YES)
                Call SendPlayerData(Index)
                Call SendSoundToMap(GetPlayerMap(Index), "mpds_pickupitem.wav")
            ' Scoring for Blue
            ElseIf GetPlayerTeam(Index, STSPath, STSMaxPlayers) = 1 And GetVar(FilePath, "Flag", "Blue") = "HasFlag" And GetPlayerMap(Index) = 33 And GetPlayerX(Index) = 16 And GetPlayerY(Index) = 29 And GetPlayerPK(Index) = 1 Then
                If GetVar(FilePath, "Flag", "Red") = "HasFlag" Then
                    Call PlayerMsg(Index, "You must retrieve your own Shroom before you can score!", BRIGHTRED)
                    Exit Sub
                End If
                    
                POINTS = CByte(GetVar(FilePath, "Points", "Blue"))
                    
                Call PlayerMsg(Index, "You scored for your team!", YELLOW)
                Call PutVar(FilePath, "Flag", "Blue", "NoFlag")
                        
                For i = 1 To 4
                    PlayerNum = CStr(i)
                    R = FindPlayer(GetVar(FilePath, "Blue", PlayerNum))
                    Q = FindPlayer(GetVar(FilePath, "Red", PlayerNum))
                        
                    If IsPlaying(Q) Then
                        Call PlayerMsg(Q, GetPlayerName(Index) & " has scored for the enemy team!", YELLOW)
                    End If
                    
                    If IsPlaying(R) Then
                        If R <> Index Then
                            Call PlayerMsg(R, GetPlayerName(Index) & " has scored for your team!", YELLOW)
                        End If
                    End If
                Next i
                    
                Call SetPlayerPK(Index, NO)
                Call SendPlayerData(Index)
                Call PutVar(FilePath, "Points", "Blue", POINTS + 1)
                Call SendSoundToMap(GetPlayerMap(Index), "smrpg_coinfrog.wav")
            ' Stealing for Red
            ElseIf GetPlayerTeam(Index, STSPath, STSMaxPlayers) = 0 And GetVar(FilePath, "Flag", "Red") <> "HasFlag" And GetPlayerMap(Index) = 33 And GetPlayerX(Index) = 16 And GetPlayerY(Index) = 29 Then
                Call PutVar(FilePath, "Flag", "Red", "HasFlag")
                    
                Call PlayerMsg(Index, "You stole the enemy Shroom!", MAGENTA)
                    
                For i = 1 To 4
                    PlayerNum = CStr(i)
                    R = FindPlayer(GetVar(FilePath, "Red", PlayerNum))
                    Q = FindPlayer(GetVar(FilePath, "Blue", PlayerNum))
                            
                    If IsPlaying(Q) Then
                        Call PlayerMsg(Q, GetPlayerName(Index) & " has stolen your Shroom!", MAGENTA)
                    End If
                    
                    If IsPlaying(R) Then
                        If R <> Index Then
                            Call PlayerMsg(R, GetPlayerName(Index) & " has stolen the enemy Shroom!", MAGENTA)
                        End If
                    End If
                Next i

                Call SetPlayerPK(Index, YES)
                Call SendPlayerData(Index)
                Call SendSoundToMap(GetPlayerMap(Index), "mpds_pickupitem.wav")
            ' Scoring for Red
            ElseIf GetPlayerTeam(Index, STSPath, STSMaxPlayers) = 0 And GetVar(FilePath, "Flag", "Red") = "HasFlag" And GetPlayerMap(Index) = 33 And GetPlayerX(Index) = 15 And GetPlayerY(Index) = 1 And GetPlayerPK(Index) = 1 Then
                If GetVar(FilePath, "Flag", "Blue") = "HasFlag" Then
                    Call PlayerMsg(Index, "You must retrieve your own Shroom before you can score!", BRIGHTRED)
                    Exit Sub
                End If
                    
                POINTS = CByte(GetVar(FilePath, "Points", "Red"))
                
                Call PlayerMsg(Index, "You scored for your team!", YELLOW)
                Call PutVar(FilePath, "Flag", "Red", "NoFlag")
                
                For i = 1 To 4
                    PlayerNum = CStr(i)
                    R = FindPlayer(GetVar(FilePath, "Red", PlayerNum))
                    Q = FindPlayer(GetVar(FilePath, "Blue", PlayerNum))
                        
                    If IsPlaying(Q) Then
                        Call PlayerMsg(Q, GetPlayerName(Index) & " has scored for the enemy team!", YELLOW)
                    End If
                    
                    If IsPlaying(R) Then
                        If R <> Index Then
                            Call PlayerMsg(R, GetPlayerName(Index) & " has scored for your team!", YELLOW)
                        End If
                    End If
                Next i
                    
                Call SetPlayerPK(Index, NO)
                Call SendPlayerData(Index)
                Call PutVar(FilePath, "Points", "Red", POINTS + 1)
                Call SendSoundToMap(GetPlayerMap(Index), "smrpg_coinfrog.wav")
            End If
        Exit Sub
    ' Starting STS game
        Case 10
            FilePath = STSPath
            TeamNumberBlue = CByte(GetVar(FilePath, "Team", "Blue"))
            TeamNumberRed = CByte(GetVar(FilePath, "Team", "Red"))
                
            ' Exit if there aren't enough players
            If TeamNumberBlue < 3 Or TeamNumberRed < 3 Then
                Exit Sub
            End If
                
            ' Exit if the teams aren't even
            If TeamNumberBlue <> TeamNumberRed Then
                Exit Sub
            End If
                
            BluePlayerCount = 0
            RedPlayerCount = 0
                
            ' Loop through all slots to determine whether all players from both teams are on the lighted tiles
            For i = 1 To 4
                PlayerNum = CStr(i)
                    
                R = FindPlayer(GetVar(FilePath, "Blue", PlayerNum))
                Q = FindPlayer(GetVar(FilePath, "Red", PlayerNum))
                            
                ' Check if the player on the Blue team is on a lighted tile
                If IsPlaying(R) Then
                    If GetPlayerMap(R) = 32 And GetPlayerX(R) = 4 And GetPlayerY(R) >= 16 And GetPlayerY(R) <= 20 Then
                        BluePlayerCount = BluePlayerCount + 1
                    End If
                End If
                    
                ' Check if the player on the Red team is on a lighted tile
                If IsPlaying(Q) Then
                    If GetPlayerMap(Q) = 32 And GetPlayerX(Q) = 26 And GetPlayerY(Q) >= 16 And GetPlayerY(Q) <= 20 Then
                        RedPlayerCount = RedPlayerCount + 1
                    End If
                End If
            Next i
                            
            ' Make sure all the players are standing on the lighted tiles
            If BluePlayerCount = TeamNumberBlue And RedPlayerCount = TeamNumberRed Then
                For i = 1 To 4
                    PlayerNum = CStr(i)
                    R = FindPlayer(GetVar(FilePath, "Blue", PlayerNum))
                    Q = FindPlayer(GetVar(FilePath, "Red", PlayerNum))
               
                    If IsPlaying(R) Then
                        Call PlayerWarp(R, 33, 13 + i, 26)
                        Call SetPlayerPK(R, NO)
                    End If
                        
                    If IsPlaying(Q) Then
                        Call PlayerWarp(Q, 33, 13 + i, 4)
                        Call SetPlayerPK(Q, NO)
                    End If
                Next i
                        
                Call PutVar(FilePath, "GameTime", "TimeLeft", "240")
                Call PutVar(FilePath, "Ingame", "Ingame", "Yes")
                Call PutVar(FilePath, "Points", "Red", "0")
                Call PutVar(FilePath, "Points", "Blue", "0")
                Call AddNewTimer(Index, 0, 20000)
            End If
        Exit Sub
    ' Start Whack-A-Monty game
        Case 11
            ' First Room
            If GetPlayerMap(Index) = 71 And GetPlayerX(Index) = 22 And GetPlayerY(Index) = 16 And HasItem(Index, 84) >= 1 Then
                ' Get Game Number
                TeamNumberBlue = 1
            ' Second Room
            ElseIf GetPlayerMap(Index) = 71 And GetPlayerX(Index) = 25 And GetPlayerY(Index) = 16 And HasItem(Index, 84) >= 1 Then
                TeamNumberBlue = 2
            ' Third Room
            ElseIf GetPlayerMap(Index) = 71 And GetPlayerX(Index) = 28 And GetPlayerY(Index) = 16 And HasItem(Index, 84) >= 1 Then
                TeamNumberBlue = 3
            End If
            
            If GetVar(App.Path & "\Scripts\" & "Whack.ini", "Game" & TeamNumberBlue, "InGame") = "Yes" Then
                Call PlayerMsg(Index, "There is already a game in progress here! Try another room or wait until a room is available.", WHITE)
                Exit Sub
            End If
            
            If HasItem(Index, 84) < 1 Then
                Exit Sub
            End If
                    
            Call TakeItem(Index, 84, 1)
            Call PutVar(App.Path & "\Scripts\" & "Whack.ini", "Game" & TeamNumberBlue, "TimeLeft", "90")
            Call PutVar(App.Path & "\Scripts\" & "Whack.ini", "Game" & TeamNumberBlue, "InGame", "Yes")
            Call PutVar(App.Path & "\Scripts\" & "Whack.ini", "Game" & TeamNumberBlue, "Points", "0")
            Call PutVar(App.Path & "\Scripts\" & "Whack.ini", "Game" & TeamNumberBlue, "Player", GetPlayerName(Index))
            Call PlayerWarp(Index, (71 + TeamNumberBlue), 17, 15)
            Call AddNewTimer(Index, 7, 15800, TeamNumberBlue)
            Call BattleMsg(Index, "Time remaining: 90 seconds", WHITE, 1)
            Call AddNewTimer(Index, 8, 1000, TeamNumberBlue, GetPlayerMap(Index))
        Exit Sub
    ' Kuribo's Shoe Special Badge
        Case 12
            If GetPlayerEquipSlotNum(Index, 4) <> 93 Then
                Call BlockPlayer(Index)
                Call PlayerMsg(Index, "This looks too dangerous to walk on. It seems like you'll need an item to cross it.", WHITE)
            End If
        Exit Sub
    ' Viewing Card Collection
        Case 13
            Call SendCardShop(Index)
        Exit Sub
    ' Entering Card Shop
        Case 14
            If GetVar(App.Path & "\Scripts\Quests.ini", GetPlayerName(Index), "ItemQuest5") <> "Done" Then
                Call BlockPlayer(Index)
                Call PlayerMsg(Index, "Card Shop Owner: Get out! We're closed!", WHITE)
                Exit Sub
            End If
        Exit Sub
    ' Little Maze for Card Shop Favor
        Case 15
            If (GetPlayerMap(Index) = 95 And GetPlayerX(Index) = 27 And GetPlayerY(Index) = 15) Or (GetPlayerMap(Index) = 95 And GetPlayerX(Index) = 27 And GetPlayerY(Index) = 16) Then
                If GetVar(App.Path & "\Scripts\Quests.ini", GetPlayerName(Index), "SwitchesHit") <> "4" Then
                    Call PlayerMsg(Index, "There are still switches blocking your way! You need to hit them all before you can pass.", WHITE)
                    Call BlockPlayer(Index)
                    Exit Sub
                Else
                    Exit Sub
                End If
            End If
                
            If GetVar(App.Path & "\Scripts\Quests.ini", GetPlayerName(Index), "Map: " & GetPlayerMap(Index) & "/X: " & GetPlayerX(Index) & "/Y: " & GetPlayerY(Index)) <> "Got" Then
                If GetVar(App.Path & "\Scripts\Quests.ini", GetPlayerName(Index), "SwitchesHit") = vbNullString Then
                    Call PutVar(App.Path & "\Scripts\Quests.ini", GetPlayerName(Index), "SwitchesHit", "0")
                End If
                    
                Call PutVar(App.Path & "\Scripts\Quests.ini", GetPlayerName(Index), "Map: " & GetPlayerMap(Index) & "/X: " & GetPlayerX(Index) & "/Y: " & GetPlayerY(Index), "Got")
                Call PlayerMsg(Index, "You step on the switch.", WHITE)
                Call PutVar(App.Path & "\Scripts\Quests.ini", GetPlayerName(Index), "SwitchesHit", CInt(GetVar(App.Path & "\Scripts\Quests.ini", GetPlayerName(Index), "SwitchesHit")) + 1)
                Call SendSoundTo(Index, "smrpg_!switch.wav")
            End If
        Exit Sub
    ' Random Warping (For Dry Bones Desert PvP Area)
        Case 16
            Randomize = Int(Rand(1, 4))
            Select Case Randomize
                Case 1
                    Call PlayerWarp(Index, 102, 15, 23)
                Case 2
                    Call PlayerWarp(Index, 102, 8, 17)
                Case 3
                    Call PlayerWarp(Index, 102, 22, 10)
                Case 4
                    Call PlayerWarp(Index, 102, 14, 12)
            End Select
        Exit Sub
    ' Golden Bullet Bill for Dry Bones Desert
        Case 17
            If GetVar(App.Path & "\Scripts\Quests.ini", GetPlayerName(Index), "Bullet Bill") <> "Has" Then
                If GetPlayerMap(Index) = 105 And GetPlayerX(Index) = 22 And GetPlayerY(Index) = 26 Then
                    If GetVar(App.Path & "\Scripts\Quests.ini", GetPlayerName(Index), "Bullet Bill") <> "Spot2" Then
                        Call PlayerMsg(Index, "You see a note on the floor. It reads: " & vbNewLine & "Now that I think about it, the Bullet Bill was in an area that had a particular plant in the corners.", WHITE)
                        Call PutVar(App.Path & "\Scripts\Quests.ini", GetPlayerName(Index), "Bullet Bill", "Spot1")
                    End If
                ElseIf GetPlayerMap(Index) = 104 And GetPlayerX(Index) = 13 And GetPlayerY(Index) = 13 Then
                    If GetVar(App.Path & "\Scripts\Quests.ini", GetPlayerName(Index), "Bullet Bill") = "Spot1" Or GetVar(App.Path & "\Scripts\Quests.ini", GetPlayerName(Index), "Bullet Bill") = "Spot2" Then
                        Call PlayerMsg(Index, "You see a note on the floor. It reads: " & vbNewLine & "Ah! I remember now! The Bullet Bill is buried in the sand near two cliffs. If I wasn't so old, I'd get it myself...", WHITE)
                        Call PutVar(App.Path & "\Scripts\Quests.ini", GetPlayerName(Index), "Bullet Bill", "Spot2")
                    End If
                ElseIf GetPlayerMap(Index) = 106 And GetPlayerX(Index) = 11 And GetPlayerY(Index) = 8 And GetVar(App.Path & "\Scripts\Quests.ini", GetPlayerName(Index), "Bullet Bill") = "Spot2" Then
                    If GetFreeSlots(Index) > 0 Then
                        Call GiveItem(Index, 112, 1)
                        Call PlayerMsg(Index, "You dug up a Golden Bullet Bill!", YELLOW)
                        Call PutVar(App.Path & "\Scripts\Quests.ini", GetPlayerName(Index), "Bullet Bill", "Has")
                    Else
                        Call PlayerMsg(Index, "Your inventory is full! Please make some room for this item.", BRIGHTRED)
                    End If
                End If
            Else
                If GetPlayerMap(Index) = 106 And GetPlayerX(Index) = 11 And GetPlayerY(Index) = 8 And HasItem(Index, 112) <> 1 Then
                    If GetFreeSlots(Index) > 0 Then
                        Call GiveItem(Index, 112, 1)
                        Call PlayerMsg(Index, "You dug up a Golden Bullet Bill!", YELLOW)
                    Else
                        Call PlayerMsg(Index, "Your inventory is full! Please make some room for this item.", BRIGHTRED)
                    End If
                End If
            End If
        Exit Sub
    ' Getting Super Shroom Supply for Super Shroom Favor
        Case 18
            If GetVar(App.Path & "\Scripts\Quests.ini", GetPlayerName(Index), "ItemQuest6") = "InProgress" And CanTake(Index, 122, 1) = False Then
                If GetFreeSlots(Index) > 0 Then
                    Call GiveItem(Index, 122, 1)
                    Call PlayerMsg(Index, "You've found the Super Shroom supply!", YELLOW)
                Else
                    Call PlayerMsg(Index, "You found the Super Shroom supply, but you need to make some room in your inventory to carry it!", BRIGHTRED)
                End If
            End If
        Exit Sub
    ' Recipe Log
        Case 19
            Call SendRecipeLog(Index)
        Exit Sub
    ' Random Warping (For Mushroom Kingdom PvP Area)
        Case 20
            Randomize = Int(Rand(1, 4))
            Select Case Randomize
                Case 1
                    Call PlayerWarp(Index, 182, 7, 12)
                Case 2
                    Call PlayerWarp(Index, 182, 19, 23)
                Case 3
                    Call PlayerWarp(Index, 182, 25, 4)
                Case 4
                    Call PlayerWarp(Index, 182, 17, 28)
            End Select
        Exit Sub
    ' Dodgeball Minigame rooms
        Case 21
            FilePath = DodgeBillPath
            ' Red Team door
            If ((GetPlayerX(Index) >= 17 And GetPlayerX(Index) <= 20 And GetPlayerY(Index) = 6) Or (GetPlayerX(Index) = 12 And GetPlayerY(Index) = 8)) And HasItem(Index, 187) >= 1 Then
                If GetVar(FilePath, "Team", "Red") = "5" Then
                    Call PlayerMsg(Index, "The red team is currently full!", WHITE)
                    Exit Sub
                ElseIf GetVar(FilePath, "Ingame", "Ingame") = "Yes" Then
                    Call PlayerMsg(Index, "There's already a game in progress!", WHITE)
                    Exit Sub
                Else
                    Call SetPlayerTeam(Index, 0, DodgeBillPath, DodgeBillMaxPlayers)
                    Call TakeItem(Index, 187, 1)
                    Call PlayerWarp(Index, 191, 16, 9)
                End If
            ' Blue Team door
            ElseIf ((GetPlayerX(Index) >= 22 And GetPlayerX(Index) <= 25 And GetPlayerY(Index) = 6) Or (GetPlayerX(Index) = 14 And GetPlayerY(Index) = 8)) And HasItem(Index, 188) >= 1 Then
                If GetVar(FilePath, "Team", "Blue") = "5" Then
                    Call PlayerMsg(Index, "The blue team is currently full!", WHITE)
                    Exit Sub
                ElseIf GetVar(FilePath, "Ingame", "Ingame") = "Yes" Then
                    Call PlayerMsg(Index, "There's already a game in progress!", WHITE)
                    Exit Sub
                Else
                    Call SetPlayerTeam(Index, 1, DodgeBillPath, DodgeBillMaxPlayers)
                    Call TakeItem(Index, 188, 1)
                    Call PlayerWarp(Index, 191, 16, 24)
                End If
            End If
        Exit Sub
    ' Exiting Dodgeball Waiting Rooms
        Case 22
            FilePath = DodgeBillPath
            
            If GetPlayerX(Index) = 16 And GetPlayerY(Index) = 12 Then
                TeamNumberRed = CByte(GetVar(FilePath, "Team", "Red"))
                
                For i = 1 To 5
                    PlayerNum = CStr(i)
                    If GetVar(FilePath, "Red", PlayerNum) = GetPlayerName(Index) Then
                        Call PutVar(FilePath, "Team", "Red", (TeamNumberRed - 1))
                        Call PutVar(FilePath, "Red", PlayerNum, "")
                        Call PlayerWarp(Index, 190, 20, 9)
                    End If
                Next i
            ElseIf GetPlayerX(Index) = 16 And GetPlayerY(Index) = 27 Then
                TeamNumberBlue = CByte(GetVar(FilePath, "Team", "Blue"))
                
                For i = 1 To 5
                    PlayerNum = CStr(i)
                    If GetVar(FilePath, "Blue", PlayerNum) = GetPlayerName(Index) Then
                        Call PutVar(FilePath, "Team", "Blue", (TeamNumberBlue - 1))
                        Call PutVar(FilePath, "Blue", PlayerNum, "")
                        Call PlayerWarp(Index, 190, 20, 9)
                    End If
                Next i
            End If
        Exit Sub
    ' Starting Dodgeball game
        Case 23
            FilePath = DodgeBillPath
            TeamNumberBlue = CByte(GetVar(FilePath, "Team", "Blue"))
            TeamNumberRed = CByte(GetVar(FilePath, "Team", "Red"))
                
            ' Exit if there aren't enough players
            If TeamNumberBlue < 2 Or TeamNumberRed < 2 Then
                Exit Sub
            End If
                
            ' Exit if the teams aren't even
            If TeamNumberBlue <> TeamNumberRed Then
                Exit Sub
            End If
                
            BluePlayerCount = 0
            RedPlayerCount = 0
                
            ' Loop through all slots to determine whether all players from both teams are on the lighted tiles
            For i = 1 To 5
                PlayerNum = CStr(i)
                    
                R = FindPlayer(GetVar(FilePath, "Blue", PlayerNum))
                Q = FindPlayer(GetVar(FilePath, "Red", PlayerNum))
                            
                ' Check if the player on the Red team is on a lighted tile
                If IsPlaying(R) Then
                    If (GetPlayerMap(R) = 191 And GetPlayerX(R) >= 14 And GetPlayerX(R) <= 18 And GetPlayerY(R) = 21) Then
                        BluePlayerCount = BluePlayerCount + 1
                    End If
                End If
                    
                ' Check if the player on the Blue team is on a lighted tile
                If IsPlaying(Q) Then
                    If (GetPlayerMap(Q) = 191 And GetPlayerX(Q) >= 14 And GetPlayerX(Q) <= 18 And GetPlayerY(Q) = 6) Then
                        RedPlayerCount = RedPlayerCount + 1
                    End If
                End If
            Next i
            
            ' Make sure all the players are standing on the lighted tiles
            If BluePlayerCount = TeamNumberBlue And RedPlayerCount = TeamNumberRed Then
                ' Respawn the map just incase
                Call RespawnMap(188)
                
                ' Warp players
                For i = 1 To 5
                    PlayerNum = CStr(i)
                    
                    R = FindPlayer(GetVar(FilePath, "Blue", PlayerNum))
                    Q = FindPlayer(GetVar(FilePath, "Red", PlayerNum))
                            
                    ' Warp players on the blue team
                    If IsPlaying(R) Then
                        Call PlayerWarp(R, 188, (12 + i), 20)
                    End If
                    
                    ' Warp players on the blue team
                    If IsPlaying(Q) Then
                        Call PlayerWarp(Q, 188, (12 + i), 10)
                    End If
                Next i
                
                Call PutVar(FilePath, "Outs", "Blue", "0")
                Call PutVar(FilePath, "Outs", "Red", "0")
                Call PutVar(FilePath, "GameTime", "TimeLeft", "300")
                Call PutVar(FilePath, "Ingame", "Ingame", "Yes")
                Call PutVar(FilePath, "Points", "Red", "0")
                Call PutVar(FilePath, "Points", "Blue", "0")
                Call AddNewTimer(Index, 9, 20000)
            End If
        Exit Sub
    ' Mushroom Ball Favor tile
        Case 24
            If GetVar(App.Path & "\Scripts\Quests.ini", GetPlayerName(Index), "ItemQuest13") = "InProgress" Then
                If GetVar(App.Path & "\Scripts\" & "MBFavor.ini", GetPlayerName(Index), "mbquest /" & GetPlayerMap(Index) & "/" & GetPlayerX(Index) & "/" & GetPlayerY(Index)) <> "1" Then
                    Call SendDataTo(Index, SPackets.Scheckemoticons & SEP_CHAR & Index & SEP_CHAR & Emoticons(0).Pic & END_CHAR)
                End If
            End If
        Exit Sub
    ' Key Falling out of tree tile (Peach Favor)
        Case 25
            
        Exit Sub
    ' Going into a house with the key (Peach Favor)
        Case 26
            If HasItem(Index, 227) >= 1 Then
                Call PlayerWarp(Index, 208, 15, 29)
            End If
        Exit Sub
    ' Plant for Seventeenth Quest Script (Getting Salestoad's Stuff again)
        Case 27
             
        Exit Sub
    ' Hide n' Sneak waiting rooms
        Case 28
            FilePath = HideNSneakPath
            
            ' Hider door
            If (GetPlayerX(Index) >= 13 And GetPlayerX(Index) <= 18) Then
                If HasItem(Index, 279) >= 1 Then
                    If GetVar(FilePath, "Team", "Hiders") = "5" Then
                        Call PlayerMsg(Index, "The Hider team is currently full!", WHITE)
                        Exit Sub
                    ElseIf GetVar(FilePath, "Ingame", "Ingame") = "Yes" Then
                        Call PlayerMsg(Index, "There's already a game in progress!", WHITE)
                        Exit Sub
                    Else
                        Call SetPlayerTeam(Index, 0, HideNSneakPath, MaxHiders)
                        Call TakeItem(Index, 279, 1)
                        Call PlayerWarp(Index, 270, 13, 25)
                    End If
                End If
            ' Seeker door
            ElseIf (GetPlayerX(Index) >= 21 And GetPlayerX(Index) <= 23) Then
                If HasItem(Index, 278) >= 1 Then
                    If GetVar(FilePath, "Team", "Seekers") = "3" Then
                        Call PlayerMsg(Index, "The Seeker team is currently full!", WHITE)
                        Exit Sub
                    ElseIf GetVar(FilePath, "Ingame", "Ingame") = "Yes" Then
                        Call PlayerMsg(Index, "There's already a game in progress!", WHITE)
                        Exit Sub
                    Else
                        Call SetPlayerTeam(Index, 1, HideNSneakPath, MaxSeekers)
                        Call TakeItem(Index, 278, 1)
                        Call PlayerWarp(Index, 270, 12, 17)
                    End If
                End If
            End If
        Exit Sub
    ' Exiting Hide n' Sneak waiting rooms
        Case 29
            FilePath = HideNSneakPath
            
            ' Stop players from exiting if there is a game in progress
            If GetVar(FilePath, "Ingame", "Ingame") = "Yes" Then
                Call PlayerMsg(Index, "You cannot leave the waiting room while the game is in progress!", BRIGHTRED)
                Exit Sub
            End If
            
            ' Hider = Red; Seeker = Blue
            
            If GetPlayerX(Index) = 19 And GetPlayerY(Index) = 25 Then
                TeamNumberRed = CByte(GetVar(FilePath, "Team", "Hiders"))
                
                For i = 1 To MaxHiders
                    PlayerNum = CStr(i)
                    
                    If GetVar(FilePath, "Hiders", PlayerNum) = GetPlayerName(Index) Then
                        Call PutVar(FilePath, "Team", "Hiders", (TeamNumberRed - 1))
                        Call PutVar(FilePath, "Hiders", PlayerNum, "")
                        Call PlayerWarp(Index, 269, 15, 7)
                    End If
                Next i
            ElseIf GetPlayerX(Index) = 9 And GetPlayerY(Index) = 19 Then
                TeamNumberBlue = CByte(GetVar(FilePath, "Team", "Seekers"))
                
                For i = 1 To MaxSeekers
                    PlayerNum = CStr(i)
                    
                    If GetVar(FilePath, "Seekers", PlayerNum) = GetPlayerName(Index) Then
                        Call PutVar(FilePath, "Team", "Seekers", (TeamNumberBlue - 1))
                        Call PutVar(FilePath, "Seekers", PlayerNum, "")
                        Call PlayerWarp(Index, 269, 15, 7)
                    End If
                Next i
            End If
        Exit Sub
    ' Starting Hide n' Sneak game
        Case 30
            FilePath = HideNSneakPath
        
            TeamNumberBlue = CByte(GetVar(FilePath, "Team", "Seekers"))
            TeamNumberRed = CByte(GetVar(FilePath, "Team", "Hiders"))
            
            ' Exit if there aren't enough players
            If TeamNumberBlue < 1 Or TeamNumberRed < 1 Then
                Exit Sub
            End If
                
            BluePlayerCount = 0 ' Seekers
            RedPlayerCount = 0 ' Hiders
            
            ' Loop through all slots to determine whether all players from both teams are on the lighted tiles
            For i = 1 To MaxHiders
                PlayerNum = CStr(i)
                
                ' Only check seekers from values 1 - 3 since there are only 3 max seekers
                If i < 4 Then
                    R = FindPlayer(GetVar(FilePath, "Seekers", PlayerNum))
                
                    ' Check if the Seeker is on a lighted tile
                    If IsPlaying(R) Then
                        If (GetPlayerX(R) = 8 And GetPlayerY(R) >= 15 And GetPlayerY(R) <= 17) Then
                            BluePlayerCount = BluePlayerCount + 1
                        End If
                    End If
                End If
                
                Q = FindPlayer(GetVar(FilePath, "Hiders", PlayerNum))
                
                ' Check if the player on the Blue team is on a lighted tile
                If IsPlaying(Q) Then
                    If (GetPlayerX(Q) = 21 And GetPlayerY(Q) >= 19 And GetPlayerY(Q) <= 23) Then
                        RedPlayerCount = RedPlayerCount + 1
                    End If
                End If
            Next i
            
            ' Make sure there is at least one player standing on the lighted tiles
            If BluePlayerCount = TeamNumberBlue And RedPlayerCount = TeamNumberRed Then
                ' State that no one has left Hide n' Sneak
                HasLeftHideNSneak = False
            
                Dim RandomSprite As Byte
                
                ' Warp the Hiders to the game arena and start the game
                ' Seekers need to wait 2 minutes for the Hiders to hide
                For i = 1 To MaxHiders
                    PlayerNum = CStr(i)
                    
                    Q = FindPlayer(GetVar(FilePath, "Hiders", PlayerNum))
                    
                    ' Warp the Hiders
                    If IsPlaying(Q) Then
                        ' Store the hiders' old sprites
                        Call SetPlayerTempSprite(Q, GetPlayerSprite(Q))
                    
                        ' Set the sprite of the hider to one of the babies
                        RandomSprite = Rand(72, 73)
                        Call SetPlayerSprite(Q, RandomSprite)
                        
                        ' Heal the player fully
                        Call SetPlayerHP(Q, GetPlayerMaxHP(Q))
                        Call SetPlayerMP(Q, GetPlayerMaxMP(Q))
                        Call SetPlayerSP(Q, GetPlayerMaxSP(Q))
                        
                        Call SendHP(Q)
                        Call SendMP(Q)
                        Call SendSP(Q)
                        
                        ' Send that the hider is playing
                        Call SendPlayingHideNSneak(Q, True)
                        
                        ' Warp the hider to the minigame
                        Call PlayerWarp(Q, 271, 15, 23)
                    End If
                Next i
                
                Call PutVar(FilePath, "PlayersFound", "PlayersFound", "0")
                Call PutVar(FilePath, "GameTime", "TimeLeft", "120")
                Call PutVar(FilePath, "GameTime", "IsSeeking", "No")
                Call PutVar(FilePath, "Ingame", "Ingame", "Yes")
                
                ' Store the Timer's Index (Player's Index)
                Call PutVar(FilePath, "TimerIndex", "TimerIndex", CStr(Index))
                
                Call AddNewTimer(Index, 10, 20000)
            End If
        Exit Sub
    ' Bean Fruit for Nineteenth Quest Script (Bean Fruit Finder)
        Case 31
        
        Exit Sub
    ' Rockade tile
        Case 32
        
        Exit Sub
    ' Currency Exchanger
        Case 33
            Call SendNpcTalkYesNoTo(Index, 221, "Would you like to exchange currency?", vbNullString, vbNullString)
        Exit Sub
    ' Anything else
        Case Else
            Call PlayerMsg(Index, "No tile script found. Please contact an admin to solve this problem.", WHITE)
        Exit Sub
    End Select
End Sub

' Executes whenever a scripted NPC does an action.
Sub ScriptedNPC(ByVal Index As Long, ByVal NpcNum As Long, ByVal Script As Long)
    Dim FilePath As String, FavorStatus As String
    Dim b As Integer, InvSlot As Integer
    Dim CardCount As Byte
    
    Select Case Script
    ' First Quest Script (Getting Koopa Shells)
        Case 0
        'Starts quest
            FilePath = App.Path & "\Scripts\Quests.ini"
            FavorStatus = GetVar(FilePath, GetPlayerName(Index), "ItemQuest1")
            
            If FavorStatus <> "Done" Then
                If FavorStatus <> "InProgress" Then
                    Call SendFavorTo(Index, "The Shell Collection", "Favor", "Hey! Would you mind getting me 3 Koopa Shells? I need them to complete my shell collection.")
                    Call PutVar(FilePath, GetPlayerName(Index), "ItemQuest1", "InProgress")
                Else
                    If CountInvItemNum(Index, 42) < 3 Then
                        Call SendFavorTo(Index, "The Shell Collection", "Need Shells!", "Did you get those shells yet?")
                    Else
                        Call TakeItem(Index, 42, 3)
                        Call GiveItem(Index, 45, 1)
                        Call SendFavorTo(Index, "The Shell Collection", "Favor Complete!", "Thanks a lot for the help! Here, take this!")
                        Call PlayerMsg(Index, "You received a Small Key!", YELLOW)
                        Call PutVar(FilePath, GetPlayerName(Index), "ItemQuest1", "Done")
                    End If
                End If
            Else
               Call SendFavorTo(Index, "The Shell Collection", "Favor Already Complete!", "Thanks for the help before; I really appreciate it!")
            End If
        Exit Sub
    ' Allows players to change their password by talking to an NPC
        Case 1
            Call PlayerQueryBox(Index, "Hey, I'm the password changer. I can change your password for you. Please enter your password.", 0)
        Exit Sub
    ' Second Quest Script (Killing Lakitus)
        Case 2
            FilePath = App.Path & "\Scripts\Quests.ini"
            FavorStatus = GetVar(FilePath, GetPlayerName(Index), "KillQuest1")
            
            If FavorStatus <> "Done" Then
                If FavorStatus <> "InProgress" Then
                    Call SendFavorTo(Index, "The Picnic Crisis", "Favor", "These Lakitus always stop me from having a nice picnic. Do you think you can get rid of them all for me? I need you to get rid of all 10 of them.")
                    Call PutVar(FilePath, GetPlayerName(Index), "KillQuest1", "InProgress")
                    Call PutVar(FilePath, GetPlayerName(Index), "LakitusKilled", "0")
                Else
                    b = 10 - CInt(GetVar(FilePath, GetPlayerName(Index), "LakitusKilled"))
                    
                    If b <= 10 And b > 1 Then
                        Call SendFavorTo(Index, "The Picnic Crisis", "Need More Kills!", "How are you doing? It seems to me like there are " & b & " Lakitus left!")
                    ElseIf b = 1 Then
                        Call SendFavorTo(Index, "The Picnic Crisis", "Need More Kills!", "How are you doing? It seems to me like there is " & b & " Lakitu left!")
                    ElseIf b = 0 Then
                        Call SendFavorTo(Index, "The Picnic Crisis", "Favor Complete!", "Thanks a lot for getting rid of them! Now they can't bother me anymore. Take this for your help.")
                        
                        If GetFreeSlots(Index) <= 0 And Not CanTake(Index, 1, 1) Then
                            Call PlayerMsg(Index, "You don't have enough inventory space to take the reward! Come back to him when you have room.", WHITE)
                            Exit Sub
                        Else
                            Call GiveItem(Index, 1, 100)
                            Call PlayerMsg(Index, "You got 100 Coins!", YELLOW)
                            Call PutVar(FilePath, GetPlayerName(Index), "KillQuest1", "Done")
                        End If
                    End If
                End If
            Else
                Call SendFavorTo(Index, "The Picnic Crisis", "Favor Already Complete!", "Thanks for helping me out before! Now I can have a picnic in peace!")
            End If
        Exit Sub
    ' Third Quest Script (Getting Mushroom Toy)
        Case 3
            FilePath = App.Path & "\Scripts\Quests.ini"
            FavorStatus = GetVar(FilePath, GetPlayerName(Index), "ItemQuest2")
            
            If FavorStatus <> "Done" Then
                If FavorStatus <> "InProgress" Then
                    Call SendFavorTo(Index, "The Lost Toy", "Favor", "I lost a Mushroom Toy that I had since I was a small Toad on this island. That thing was so important to me...")
                    Call PutVar(FilePath, GetPlayerName(Index), "ItemQuest2", "InProgress")
                Else
                    If CountInvItemNum(Index, 65) < 1 Then
                        Call SendFavorTo(Index, "The Lost Toy", "Need Toy!", "I really wonder where my toy is...")
                    Else
                        Call TakeItem(Index, 65, 1)
                        Call PutVar(FilePath, GetPlayerName(Index), "ItemQuest2", "Done")
                        Call SendFavorTo(Index, "The Lost Toy", "Favor Complete!", "Wow! You found my toy! Thanks so much!")
                    End If
                End If
            Else
                Call SendFavorTo(Index, "The Lost Toy", "Favor Already Complete!", "Hey! Thanks again for finding my toy!")
            End If
        Exit Sub
    ' Fourth Quest Script (Getting Salestoad's Stuff)
        Case 4
            FilePath = App.Path & "\Scripts\Quests.ini"
            FavorStatus = GetVar(FilePath, GetPlayerName(Index), "ItemQuest3")
            
            If FavorStatus <> "Done" Then
                If FavorStatus <> "InProgress" And FavorStatus <> "InProgress2" Then
                    Call SendFavorTo(Index, "The Package", "Favor", "Hey there! I would normally show you what I have for sale, but I lent my sick friend in Mushroom Town all my stuff! Can you please get it for me?")
                    Call PutVar(FilePath, GetPlayerName(Index), "ItemQuest3", "InProgress")
                    Exit Sub
                Else
                    If CountInvItemNum(Index, 68) < 1 Then
                        Call SendFavorTo(Index, "The Package", "Retrieve Salestoad's Stuff!", "Hey again! Did you get my stuff back from my friend yet? He lives in Mushroom Town and wasn't feeling well the last time I visited him.")
                    Else
                        Call TakeItem(Index, 68, 1)
                        Call PutVar(FilePath, GetPlayerName(Index), "ItemQuest3", "Done")
                        Call SendFavorTo(Index, "The Package", "Favor Complete!", "Wow! You got my stuff back! Now I can sell all the great items I have! Thanks a lot!")
                        Call PlayerMsg(Index, "You can now buy items from the Salestoad!", WHITE)
                    End If
                End If
            Else
                Call SendTrade(Index, 5)
            End If
        Exit Sub
    ' Patient for Fourth Quest Script
        Case 5
            FilePath = App.Path & "\Scripts\Quests.ini"
            FavorStatus = GetVar(FilePath, GetPlayerName(Index), "ItemQuest3")
            
            If FavorStatus = "InProgress" Then
                Call SendNpcTalkTo(Index, NpcNum, "Oh, the Salestoad's stuff? Here you go; tell him I'm very sorry.")
                If GetFreeSlots(Index) > 0 Then
                    Call GiveItem(Index, 68, 1)
                    Call PlayerMsg(Index, "You got the Salestoad's stuff!", WHITE)
                    Call PutVar(FilePath, GetPlayerName(Index), "ItemQuest3", "InProgress2")
                Else
                    Call SendNpcTalkTo(Index, NpcNum, "Sorry, you need to make room in your inventory to hold this.")
                    Exit Sub
                End If
            Else
                Call SendNpcTalkTo(Index, NpcNum, "I'm not feeling too well... the doctor told me to stay here until he can get me a Doctor's Check.")
            End If
        Exit Sub
    ' Guild Creating NPC
        Case 6
            Call PlayerQueryBox(Index, "Hey, you look like you'd like to form a Group. Just tell me what you'd like to make your Group name. It's only 5,000 Coins!", 2)
        Exit Sub
    ' Fifth Quest Script (Getting 3 Drillbit Crab Shells)
        Case 7
            FilePath = App.Path & "\Scripts\Quests.ini"
            FavorStatus = GetVar(FilePath, GetPlayerName(Index), "ItemQuest4")
            
            If FavorStatus <> "Done" Then
                If FavorStatus <> "InProgress" Then
                    Call SendFavorTo(Index, "The Rare Shell Collector", "Favor", "Hello! I've heard that the Drillbit Crabs on this island have very rare shells. Do you think you can get me 3 of them for my collection?")
                    Call PutVar(FilePath, GetPlayerName(Index), "ItemQuest4", "InProgress")
                    Exit Sub
                Else
                    If CountInvItemNum(Index, 69) < 3 Then
                        Call SendFavorTo(Index, "The Rare Shell Collector", "Need Crab Shells!", "Hey! I see that you don't have all 3 of the shells I'm looking for. Would you please try getting them all?")
                    Else
                        Call TakeItem(Index, 69, 3)
                        Call PutVar(FilePath, GetPlayerName(Index), "ItemQuest4", "Done")
                        Call SendFavorTo(Index, "The Rare Shell Collector", "Favor Complete!", "Wow! Thank you so much! Now I can continue to be one of the world's greatest shell collectors!")
                        Call GiveItem(Index, 98, 5)
                        Call PlayerMsg(Index, "You got 5 Melons!", WHITE)
                    End If
                End If
            Else
                Call SendFavorTo(Index, "The Rare Shell Collector", "Favor Already Complete!", "Hey, buddy! Thanks again for your help before; I don't know how else I can repay you.")
            End If
        Exit Sub
    ' Sixth Quest Script (Delivering Mysterious Envelope)
        Case 8
            FilePath = App.Path & "\Scripts\Quests.ini"
            FavorStatus = GetVar(FilePath, GetPlayerName(Index), "ItemQuest5")
            
            If FavorStatus <> "Done" Then
                If FavorStatus <> "InProgress" Then
                    Call SendFavorTo(Index, "Toadsdale's Request", "Favor", "Hello. Having trouble getting in the Card Shop? I'll help you out if you get me the envelope I dropped. It's very important...")
                    Call PutVar(FilePath, GetPlayerName(Index), "ItemQuest5", "InProgress")
                    Exit Sub
                Else
                    If CountInvItemNum(Index, 99) < 1 Then
                        Call SendFavorTo(Index, "Toadsdale's Request", "Missing Item!", "You still didn't get it yet? I'll just have to wait longer then...")
                    Else
                        Call TakeItem(Index, 99, 1)
                        Call PutVar(FilePath, GetPlayerName(Index), "ItemQuest5", "Done")
                        Call SendFavorTo(Index, "Toadsdale's Request", "Favor Complete!", "Ahh, thank you very much! I will convince the Card Shop Owner to let you in.")
                    End If
                End If
            Else
                Call SendFavorTo(Index, "Toadsdale's Request", "Favor Already Complete!", "Thanks again. I will go about my business now.")
            End If
        Exit Sub
    ' Seventh Quest Script (Super Shroom Quest)
        Case 9
            FilePath = App.Path & "\Scripts\Quests.ini"
            FavorStatus = GetVar(FilePath, GetPlayerName(Index), "ItemQuest6")
            
            If FavorStatus <> "Done" Then
                If FavorStatus <> "InProgress" Then
                    Call SendFavorTo(Index, "Restock The Shop", "Favor", "Oh, no! I can't find out where all of the Super Shrooms went! They must've all gotten lost in the Desert somewhere. Do you think you can help?")
                    Call PlayerMsg(Index, "It doesn't seem like the supplies could be very far away, so you probably won't need to cross the vines to get to them.", WHITE)
                    Call PutVar(FilePath, GetPlayerName(Index), "ItemQuest6", "InProgress")
                    Exit Sub
                Else
                    If CountInvItemNum(Index, 122) < 1 Then
                        Call SendFavorTo(Index, "Restock The Shop", "Missing Item!", "Did you get that supply back yet? Please get it back soon because I really need it.")
                    Else
                        Call TakeItem(Index, 122, 1)
                        Call PutVar(FilePath, GetPlayerName(Index), "ItemQuest6", "Done")
                        Call SendFavorTo(Index, "Restock The Shop", "Favor Complete!", "Thanks for getting back that Super Shroom Supply! Now you can purchase those delicious Shrooms from the shop!")
                    End If
                End If
            Else
                Call SendFavorTo(Index, "Restock The Shop", "Favor Already Complete!", "Thank you very much for your help; since we got our Super Shrooms back, sales have gone up 50%!")
            End If
        Exit Sub
    ' Eight Quest Script (Toy in Kinopio Village)
        Case 10
            FilePath = App.Path & "\Scripts\Quests.ini"
            FavorStatus = GetVar(FilePath, GetPlayerName(Index), "ItemQuest7")
            
            If FavorStatus <> "Done" Then
                If FavorStatus <> "InProgress" And FavorStatus <> "InProgress2" And FavorStatus <> "InProgress3" Then
                    Call SendFavorTo(Index, "The Missing Toy", "Favor", "Help! I don't know where my toy went. Can you please help me find it?")
                    Call PutVar(FilePath, GetPlayerName(Index), "ItemQuest7", "InProgress")
                    Exit Sub
                Else
                    If CountInvItemNum(Index, 130) < 1 Then
                        Call SendFavorTo(Index, "The Missing Toy", "Need Toy!", "Were you able to get my toy back?")
                    Else
                        Call TakeItem(Index, 130, 1)
                        Call PutVar(FilePath, GetPlayerName(Index), "ItemQuest7", "Done")
                        Call SendFavorTo(Index, "The Missing Toy", "Favor Complete!", "Thank you so much! There must be some way I can repay you...I think I know!")
                        Call GiveItem(Index, 116, 1)
                        Call PlayerMsg(Index, "You got a Pokey Card!", YELLOW)
                    End If
                End If
            Else
                Call SendFavorTo(Index, "The Missing Toy", "Favor Already Complete!", "I'm very grateful for your help. Now, I can relax since my toy is safe.")
            End If
        Exit Sub
    ' Toad #1 (Eighth Quest Script)
        Case 11
            FilePath = App.Path & "\Scripts\Quests.ini"
            FavorStatus = GetVar(FilePath, GetPlayerName(Index), "ItemQuest7")
            
            If FavorStatus = "InProgress" Then
                Call SendNpcTalkTo(Index, NpcNum, "You're looking for the Goomba Toy? I remember I lent it out to someone else in the village.")
                Call PutVar(FilePath, GetPlayerName(Index), "ItemQuest7", "InProgress2")
            Else
                Call SendNpcTalkTo(Index, NpcNum, "How's everything? Are you enjoying yourself here in Kinopio Village?", "Anyway, make yourself at home!")
            End If
        Exit Sub
    ' Toad #2 (Eighth Quest Script)
        Case 12
            FilePath = App.Path & "\Scripts\Quests.ini"
            FavorStatus = GetVar(FilePath, GetPlayerName(Index), "ItemQuest7")
            
            If FavorStatus = "InProgress2" Then
                Call SendNpcTalkTo(Index, NpcNum, "The Goomba Toy? Oh, I think I have it here somewhere...Here it is!")
                
                If GetFreeSlots(Index) > 0 Then
                    Call GiveItem(Index, 130, 1)
                    Call PlayerMsg(Index, "You got the Goomba Toy!", WHITE)
                    Call PutVar(FilePath, GetPlayerName(Index), "ItemQuest7", "InProgress3")
                Else
                    Call SendNpcTalkTo(Index, NpcNum, "Sorry, you'll need to make some room in your inventory for the Goomba Toy.")
                    Exit Sub
                End If
            Else
                Call SendNpcTalkTo(Index, NpcNum, "Hey! Welcome to Kinopio Village!", "Everyone is welcome at my house; enjoy your stay!")
            End If
        Exit Sub
    ' Ninth Quest Script (Getting 10 Pokey Flowers)
        Case 13
            FilePath = App.Path & "\Scripts\Quests.ini"
            FavorStatus = GetVar(FilePath, GetPlayerName(Index), "ItemQuest8")
            
            If FavorStatus <> "Done" Then
                If FavorStatus <> "InProgress" Then
                    Call SendFavorTo(Index, "The Pokey Hunter", "Favor", "I really love Pokey Flowers! They're so soft and fluffy! Can you get me 10?")
                    Call PutVar(FilePath, GetPlayerName(Index), "ItemQuest8", "InProgress")
                    Exit Sub
                Else
                    If CountInvItemNum(Index, 91) < 10 Then
                        Call SendFavorTo(Index, "The Pokey Hunter", "Need Pokey Flowers!", "I really want those 10 Pokey Flowers! Please hurry up!")
                    Else
                        Call TakeItem(Index, 91, 10)
                        Call PutVar(FilePath, GetPlayerName(Index), "ItemQuest8", "Done")
                        Call SendFavorTo(Index, "The Pokey Hunter", "Favor Complete!", "Wow, thanks a lot! They're so nice and fluffy! Here's a reward for helping me.")
                        Call SetPlayerExp(Index, GetPlayerExp(Index) + 400)
                        Call GiveItem(Index, 47, 10)
                        Call SendPlayerData(Index)
                        Call CheckPlayerLevelUp(Index)
                        Call PlayerMsg(Index, "You got 400 experience points and 10 Dim Star Pieces!", YELLOW)
                    End If
                End If
            Else
                Call SendFavorTo(Index, "The Pokey Hunter", "Favor Already Complete!", "Thanks for those flowers before; I'm really enjoying them!")
            End If
        Exit Sub
    ' Tenth Quest Script (Counting Dry Bones Statues)
        Case 14
            FilePath = App.Path & "\Scripts\Quests.ini"
            FavorStatus = GetVar(FilePath, GetPlayerName(Index), "OtherQuest1")
            
            If FavorStatus <> "Done" Then
                If FavorStatus <> "InProgress" And FavorStatus <> "InProgress2" Then
                    For b = 1 To 7
                        If GetPlayerEquipSlotNum(Index, b) > 0 Then
                            Call SendNpcTalkTo(Index, NpcNum, "Another enemy? Get away from me!")
                            Call PlayerMsg(Index, "He seems to be scared of you with your equipment on. Unequip everything and then try talking to him again.", WHITE)
                            Exit Sub
                        End If
                    Next b
                    
                    Call SendFavorTo(Index, "The Explorer", "Favor", "Oh, you're friendly! Mind doing me a favor? I'm an explorer, and I'm collecting data on different areas.", "Can you find out how many Dry Bones statues there are in this desert for me?")
                    Call PutVar(FilePath, GetPlayerName(Index), "OtherQuest1", "InProgress")
                    Exit Sub
                Else
                    Call PlayerQueryBox(Index, "So, how many Dry Bones statues did you find?", 3)
                End If
            Else
                Call SendFavorTo(Index, "The Explorer", "Favor Already Complete!", "Thanks for helping me collect some data on this place. Maybe I'll see you another time!")
            End If
        Exit Sub
    ' Cooking
        Case 15
            Call SendDataTo(Index, SPackets.Scookitem & SEP_CHAR & NpcNum & END_CHAR)
        Exit Sub
    ' Yoshi Color Changing NPC
        Case 16
            If GetPlayerClass(Index) = 4 Then
                Call PlayerQueryBox(Index, "Hello! Which color would you like me to spray paint you?", 4)
            Else
                Call SendNpcTalkTo(Index, NpcNum, "Sorry, but I can only serve Yoshis.")
            End If
        Exit Sub
    ' Toad Color Changing NPC
        Case 17
            If GetPlayerClass(Index) = 5 Then
                Call PlayerQueryBox(Index, "Hello! Which color would you like me to spray paint you?", 5)
            Else
                Call SendNpcTalkTo(Index, NpcNum, "Sorry, but I can only serve Toads.")
            End If
        Exit Sub
    ' Eleventh Quest Script (Getting a Shroom Cake)
        Case 18
            FilePath = App.Path & "\Scripts\Quests.ini"
            FavorStatus = GetVar(FilePath, GetPlayerName(Index), "ItemQuest9")
            
            If FavorStatus <> "Done" Then
                If FavorStatus <> "InProgress" Then
                    Call SendFavorTo(Index, "The Toad's Craving", "Favor", "Oh man...I'm really in the mood for a Shroom Cake right now. Those Paratroopas stole all the Cake Mix, though.")
                    Call PutVar(FilePath, GetPlayerName(Index), "ItemQuest9", "InProgress")
                    Exit Sub
                Else
                    If CountInvItemNum(Index, 146) < 1 Then
                        Call SendFavorTo(Index, "The Toad's Craving", "Need Shroom Cake!", "Oh man...I'm really in the mood for a Shroom Cake right now. Those Paratroopas stole all the Cake Mix, though.")
                    Else
                        Call TakeItem(Index, 146, 1)
                        Call PutVar(FilePath, GetPlayerName(Index), "ItemQuest9", "Done")
                        Call SendFavorTo(Index, "The Toad's Craving", "Favor Complete!", "Oh, thanks a lot! *Eats* Wow, that was amazing! Here, take a look at what I can offer you.")
                        Call PlayerMsg(Index, "You can now buy items from the Toad!", WHITE)
                    End If
                End If
            Else
                Call SendTrade(Index, 15)
            End If
        Exit Sub
    ' Toad that sells Cake Mix
        Case 19
            Call PlayerQueryBox(Index, "Hello! Would you like to buy a Cake Mix for 1 Dim Star Piece?", 6)
        Exit Sub
    ' Toad that sells orange Bob-Omb
        Case 20
            Call PlayerQueryBox(Index, "Hey! Having trouble getting farther? I have the Bob-Omb that you'll need! It'll cost you 50 Dim Star Pieces. Would you like to buy it?", 7)
        Exit Sub
    ' Card Shop Owner
        Case 21
            FilePath = App.Path & "\Scripts\" & "CardRewards.ini"
            For b = 94 To MAX_ITEMS
                If GetVar(App.Path & "\Scripts\" & "Cards.ini", CStr(b), GetPlayerName(Index)) = "Has" Then
                    CardCount = CardCount + 1
                    
                    Select Case CardCount
                        ' 5 Cards reward
                        Case 5, 6, 7, 8, 9
                            If GetVar(FilePath, GetPlayerName(Index), "5") <> "Got" Then
                                Call SendNpcTalkTo(Index, NpcNum, "Oh, you collected 5 Cards! Excellent job!", "Here, take this!")
                                
                                InvSlot = FindOpenInvSlot(Index, 1)
                                
                                If InvSlot > 0 Then
                                    Call GiveItem(Index, 1, 100)
                                    Call PlayerMsg(Index, "You got 100 Coins!", YELLOW)
                                    Call PutVar(FilePath, GetPlayerName(Index), "5", "Got")
                                    Exit Sub
                                Else
                                    Call SendNpcTalkTo(Index, NpcNum, "You've earned a reward, but you can't carry this now! Make room for it and come back later!")
                                    Exit Sub
                                End If
                            End If
                        ' 10 Cards reward
                        Case 10, 11, 12, 13, 14
                            If GetVar(FilePath, GetPlayerName(Index), "10") <> "Got" Then
                                Call SendNpcTalkTo(Index, NpcNum, "Oh, you collected 10 Cards! Excellent job!", "Here, take this!")
                                
                                InvSlot = FindOpenInvSlot(Index, 1)
                                
                                If InvSlot > 0 Then
                                    Call GiveItem(Index, 1, 200)
                                    Call PlayerMsg(Index, "You got 200 Coins!", YELLOW)
                                    Call PutVar(FilePath, GetPlayerName(Index), "10", "Got")
                                    Exit Sub
                                Else
                                    Call SendNpcTalkTo(Index, NpcNum, "You've earned a reward, but you can't carry this now! Make room for it and come back later!")
                                    Exit Sub
                                End If
                            End If
                        ' 15 Cards reward
                        Case 15, 16, 17, 18, 19
                            If GetVar(FilePath, GetPlayerName(Index), "15") <> "Got" Then
                                Call SendNpcTalkTo(Index, NpcNum, "Oh, you collected 15 Cards! Excellent job!", "Here, take this!")
                                
                                InvSlot = FindOpenInvSlot(Index, 47)
                                
                                If InvSlot > 0 Then
                                    Call GiveItem(Index, 47, 20)
                                    Call PlayerMsg(Index, "You got 20 Dim Star Pieces!", YELLOW)
                                    Call PutVar(FilePath, GetPlayerName(Index), "15", "Got")
                                    Exit Sub
                                Else
                                    Call SendNpcTalkTo(Index, NpcNum, "You've earned a reward, but you can't carry this now! Make room for it and come back later!")
                                    Exit Sub
                                End If
                            End If
                        ' 20 Cards reward
                        Case 20, 21, 22, 23, 24
                            If GetVar(FilePath, GetPlayerName(Index), "20") <> "Got" Then
                                Call SendNpcTalkTo(Index, NpcNum, "Oh, you collected 20 Cards! Excellent job!", "Here, take this!")
                                
                                InvSlot = FindOpenInvSlot(Index, 179)
                                
                                If InvSlot > 0 Then
                                    Call GiveItem(Index, 179, 5)
                                    Call PlayerMsg(Index, "You got 5 Star Pieces!", YELLOW)
                                    Call PutVar(FilePath, GetPlayerName(Index), "20", "Got")
                                    Exit Sub
                                Else
                                    Call SendNpcTalkTo(Index, NpcNum, "You've earned a reward, but you can't carry this now! Make room for it and come back later!")
                                    Exit Sub
                                End If
                            End If
                        ' 25 Cards reward
                        Case 25, 26, 27, 29, 29
                            If GetVar(FilePath, GetPlayerName(Index), "25") <> "Got" Then
                                Call SendNpcTalkTo(Index, NpcNum, "Oh, you collected 25 Cards! Excellent job!", "Here, take this!")
                                
                                InvSlot = FindOpenInvSlot(Index, 1)
                                
                                If InvSlot > 0 Then
                                    Call GiveItem(Index, 1, 500)
                                    Call PlayerMsg(Index, "You got 500 Coins!", YELLOW)
                                    Call PutVar(FilePath, GetPlayerName(Index), "25", "Got")
                                    Exit Sub
                                Else
                                    Call SendNpcTalkTo(Index, NpcNum, "You've earned a reward, but you can't carry this now! Make room for it and come back later!")
                                    Exit Sub
                                End If
                            End If
                    End Select
                End If
            Next b
            
            ' Boss card rewards
            ' Tutankoopa
            If GetVar(FilePath, GetPlayerName(Index), "Tutankoopa") <> "Got" Then
                If GetVar(App.Path & "\Scripts\" & "Cards.ini", CStr(119), GetPlayerName(Index)) = "Has" Then
                    Call SendNpcTalkTo(Index, NpcNum, "Wow, you've managed to collect the Tutankoopa Card! Amazing job!", "Here, take this!")
                    
                    InvSlot = FindOpenInvSlot(Index, 193)
                    
                    If InvSlot > 0 Then
                        Call GiveItem(Index, 193, 1)
                        Call PlayerMsg(Index, "You got an Ultra Shroom!", YELLOW)
                        Call PutVar(FilePath, GetPlayerName(Index), "Tutankoopa", "Got")
                        Exit Sub
                    Else
                        Call SendNpcTalkTo(Index, NpcNum, "You've earned a reward, but you can't carry this now! Make room for it and come back later!")
                        Exit Sub
                    End If
                End If
            End If
            ' Runaway Chain Chomp
            If GetVar(FilePath, GetPlayerName(Index), "RCC") <> "Got" Then
                If GetVar(App.Path & "\Scripts\" & "Cards.ini", CStr(166), GetPlayerName(Index)) = "Has" Then
                    Call SendNpcTalkTo(Index, NpcNum, "Wow, you've managed to collect the Runaway Chain Chomp Card! Amazing job!", "Here, take this!")
                    
                    InvSlot = FindOpenInvSlot(Index, 47)
                    
                    If InvSlot > 0 Then
                        Call GiveItem(Index, 47, 5)
                        Call GiveItem(Index, 1, 1000)
                        Call PlayerMsg(Index, "You got 1,000 Coins and 5 Dim Star Pieces!", YELLOW)
                        Call PutVar(FilePath, GetPlayerName(Index), "RCC", "Got")
                        Exit Sub
                    Else
                        Call SendNpcTalkTo(Index, NpcNum, "You've earned a reward, but you can't carry this now! Make room for it and come back later!")
                        Exit Sub
                    End If
                End If
            End If
            ' Kamek
            If GetVar(FilePath, GetPlayerName(Index), "Kamek") <> "Got" Then
                If GetVar(App.Path & "\Scripts\" & "Cards.ini", CStr(220), GetPlayerName(Index)) = "Has" Then
                    Call SendNpcTalkTo(Index, NpcNum, "Wow, you've managed to collect the Kamek Card! Amazing job!", "Here, take this!")
                    
                    InvSlot = FindOpenInvSlot(Index, 194)
                    
                    If InvSlot > 0 Then
                        Call GiveItem(Index, 194, 1)
                        Call GiveItem(Index, 1, 500)
                        Call PlayerMsg(Index, "You got 500 Coins and a Jammin' Jelly!", YELLOW)
                        Call PutVar(FilePath, GetPlayerName(Index), "Kamek", "Got")
                        Exit Sub
                    Else
                        Call SendNpcTalkTo(Index, NpcNum, "You've earned a reward, but you can't carry this now! Make room for it and come back later!")
                        Exit Sub
                    End If
                End If
            End If
            ' Shy Guy General
            If GetVar(FilePath, GetPlayerName(Index), "Shy Guy General") <> "Got" Then
                If GetVar(App.Path & "\Scripts\" & "Cards.ini", CStr(255), GetPlayerName(Index)) = "Has" Then
                    Call SendNpcTalkTo(Index, NpcNum, "Wow, you've managed to collect the Shy Guy General Card! Amazing job!", "Here, take this!")
                    
                    InvSlot = FindOpenInvSlot(Index, 179)
                    
                    If InvSlot > 0 Then
                        Call GiveItem(Index, 179, 5)
                        Call GiveItem(Index, 1, 750)
                        Call PlayerMsg(Index, "You got 750 Coins and 5 Star Pieces!", YELLOW)
                        Call PutVar(FilePath, GetPlayerName(Index), "Shy Guy General", "Got")
                        Exit Sub
                    Else
                        Call SendNpcTalkTo(Index, NpcNum, "You've earned a reward, but you can't carry this now! Make room for it and come back later!")
                        Exit Sub
                    End If
                End If
            End If
             
            Call SendNpcTalkTo(Index, NpcNum, NPC(NpcNum).AttackSay, NPC(NpcNum).AttackSay2)
        Exit Sub
    ' Twelfth Quest Script (Getting the Gold Card)
        Case 22
            FilePath = App.Path & "\Scripts\Quests.ini"
            FavorStatus = GetVar(FilePath, GetPlayerName(Index), "ItemQuest10")
            
            If FavorStatus <> "Done" Then
                ' Check if the player has the Gold Card
                If HasItem(Index, 178) >= 1 Then
                    If GetFreeSlots(Index) > 0 Then
                        Call SendNpcTalkTo(Index, NpcNum, "Oh, what's that you have there? A Gold Card? Can I have it, please?", "Thanks a lot! I've always wanted one! Here, take these!")
                        Call TakeItem(Index, 178, 1)
                        Call GiveItem(Index, 43, 2)
                        
                        Call PlayerMsg(Index, "You got 2 Life Shrooms! They'll revive you when you get defeated in combat!", YELLOW)
                        Call PutVar(FilePath, GetPlayerName(Index), "ItemQuest10", "Done")
                    Else
                        Call SendNpcTalkTo(Index, NpcNum, "Oh, what's that you have there? A Gold Card? Can I have it, please?", "I'd love to take it, but I can't until you have enough inventory space for my reward.")
                    End If
                Else
                    Call SendNpcTalkTo(Index, NpcNum, "Did you know that this was the first house built in the Mushroom Kingdom?", "It's really cool to be living in it!")
                End If
            Else
                Call SendNpcTalkTo(Index, NpcNum, "Did you know that this was the first house built in the Mushroom Kingdom?", "It's really cool to be living in it!")
            End If
        Exit Sub
    ' Thirteenth Quest Script (Getting the Jelly Ultra)
        Case 23
            FilePath = App.Path & "\Scripts\Quests.ini"
            FavorStatus = GetVar(FilePath, GetPlayerName(Index), "ItemQuest11")
            
            If FavorStatus <> "Done" Then
                ' Check if the player has the Jelly Ultra
                If HasItem(Index, 212) >= 1 Then
                    If GetFreeSlots(Index) > 0 Then
                        Call SendNpcTalkTo(Index, NpcNum, "Ah, I beg your pardon. Would you mind getting me a Jelly Ultra?", "I really need it for my son; he's studying abroad in Goomba Village, and it would cheer him up.")
                        Call TakeItem(Index, 212, 1)
                        Call GiveItem(Index, 1, 200)
                        Call GiveItem(Index, 179, 1)
                        
                        Call SendNpcTalkTo(Index, NpcNum, "Oh, thank you so much! My son will love this! Here's a gift for your generosity!")
                        Call PlayerMsg(Index, "You got 200 Coins and a Star Piece!", YELLOW)
                        Call PutVar(FilePath, GetPlayerName(Index), "ItemQuest11", "Done")
                    Else
                        Call SendNpcTalkTo(Index, NpcNum, "Oh, thank you so much! Unfortunately, you don't have enough inventory space to accept my reward. Please make some room for it.")
                    End If
                Else
                    Call SendNpcTalkTo(Index, NpcNum, "Ah, I beg your pardon. Would you mind getting me a Jelly Ultra?", "I really need it for my son; he's studying abroad in Goomba Village, and it would cheer him up.")
                End If
            Else
                Call SendNpcTalkTo(Index, NpcNum, "Thanks again for that Jelly Ultra! I'll send it to my son as soon as I get a chance!")
            End If
        Exit Sub
    ' Fourteenth Quest Script (Counting buildings in the Mushroom Kingdom)
        Case 24
            FilePath = App.Path & "\Scripts\Quests.ini"
            FavorStatus = GetVar(FilePath, GetPlayerName(Index), "OtherQuest2")
            
            If FavorStatus <> "Done" Then
                If FavorStatus <> "InProgress" And FavorStatus <> "InProgress2" Then
                    For b = 1 To 7
                        If GetPlayerEquipSlotNum(Index, b) > 0 Then
                            Call SendNpcTalkTo(Index, NpcNum, "Another enemy? Get away from me!")
                            Call PlayerMsg(Index, "He seems to be scared of you with your equipment on. Unequip everything and then try talking to him again.", WHITE)
                            Exit Sub
                        End If
                    Next b
                    
                    Call SendFavorTo(Index, "The Explorer 2", "Favor", "Oh, it's you again. I'm sorry.", "Can you help me out again by finding out how many buildings are in the Mushroom Kingdom, including the castle?")
                    Call PutVar(FilePath, GetPlayerName(Index), "OtherQuest2", "InProgress")
                    Exit Sub
                Else
                    Call PlayerQueryBox(Index, "So, how many buildings did you find?", 8)
                End If
            Else
                Call SendFavorTo(Index, "The Explorer 2", "Favor Already Complete!", "Thanks for helping me collect some data on this place. Maybe I'll see you another time!")
            End If
        Exit Sub
    ' Fifteenth Quest Script (Retrieving a letter for Peach)
        Case 25
            FilePath = App.Path & "\Scripts\Quests.ini"
            FavorStatus = GetVar(FilePath, GetPlayerName(Index), "OtherQuest3")
            
            Select Case FavorStatus
                Case "Done"
                    Call SendFavorTo(Index, "Princess Peach's Lost Possession", "Favor Already Complete!", "Thank you very much for all that you did for me! Your name will be honored in this kingdom forever!")
                Case "InProgress"
                    ' Check if the player has the letter
                    If HasItem(Index, 239) >= 1 Then
                        Call TakeItem(Index, 239, 1)
                        Call SendFavorTo(Index, "Princess Peach's Lost Possession", "Delivered Letter!", "Oh, that's my letter! Thank you very much! *Reads letter* Oh, I remember now! I entrusted a Goomba in Mushroom Kingdom Path with my precious item.", "If it wouldn't trouble you, please talk to him and get me my item back.")
                    
                        Call PutVar(FilePath, GetPlayerName(Index), "OtherQuest3", "InProgress2")
                    Else
                        Call SendFavorTo(Index, "Princess Peach's Lost Possession", "Favor", "Hello there! I really need my precious item, but I forgot who I entrusted it to.", "Would you mind helping me retrieve it? I will reward you handsomely.")
                    End If
                Case "InProgress2"
                    ' Check if the player has peach's parcel
                    If HasItem(Index, 240) >= 1 Then
                        If GetFreeSlots(Index) >= 2 Then
                            Call TakeItem(Index, 240, 1)
                            Call SendFavorTo(Index, "Princess Peach's Lost Possession", "Favor Complete!", "Oh, this is it! Thank you so much! For your efforts, I will bestow you with wealth!")
                            
                            Call GiveItem(Index, 1, 2000)
                            Call GiveItem(Index, 223, 1)
                            Call GiveItem(Index, 241, 1)
                            
                            Call PlayerMsg(Index, "You got 2,000 Coins, a Deluxe Meal, and an HP FP Down, Attack Defense Up badge!", YELLOW)
                            Call PutVar(FilePath, GetPlayerName(Index), "OtherQuest3", "Done")
                        Else
                            Call SendFavorTo(Index, "Princess Peach's Lost Possession", "Favor Complete!", "Oh, this is it! Thank you so much! For your efforts, I will bestow you with wealth!", "It seems like you cannot hold this wealth in your inventory. Come back once you have made more room in your inventory.")
                        End If
                    Else
                        Call SendFavorTo(Index, "Princess Peach's Lost Possession", "Need Parcel!", "I entrusted a Goomba in Mushroom Kingdom Path with my precious item. If it wouldn't trouble you, please talk to him and get me my item back.")
                    End If
                Case Else
                    Call SendFavorTo(Index, "Princess Peach's Lost Possession", "Favor", "Hello there! I really need my precious item, but I forgot who I entrusted it to.", "Would you mind helping me retrieve it? I will reward you handsomely.")
                    Call PutVar(FilePath, GetPlayerName(Index), "OtherQuest3", "InProgress")
            End Select
        Exit Sub
    ' Goomba that gives you the parcel
        Case 26
            FilePath = App.Path & "\Scripts\Quests.ini"
            FavorStatus = GetVar(FilePath, GetPlayerName(Index), "ItemQuest12")
            
            Select Case FavorStatus
                Case "Done"
                    Call SendNpcTalkTo(Index, NpcNum, "Well, this place seems okay to stay in for a bit.", "Maybe I should stick around here more often.")
                Case "InProgress"
                    ' Check if the player has the Honey Super
                    If HasItem(Index, 138) >= 1 Then
                        Call SendNpcTalkTo(Index, NpcNum, "Oh, thanks a lot; I was starving! Here's the parcel Peach gave me!")
                        
                        Call TakeItem(Index, 138, 1)
                        Call GiveItem(Index, 240, 1)
                        
                        Call PlayerMsg(Index, "You got Peach's Parcel!", YELLOW)
                        Call PutVar(FilePath, GetPlayerName(Index), "ItemQuest12", "Done")
                    Else
                        Call SendNpcTalkTo(Index, NpcNum, "What? Peach's item? Oh yeah! I have that!", "But I am kind of hungry...mind getting me a Honey Super?")
                    End If
                Case Else
                    If GetVar(FilePath, GetPlayerName(Index), "OtherQuest3") <> "InProgress2" Then
                        Call SendNpcTalkTo(Index, NpcNum, "Well, this place seems okay to stay in for a bit.", "Maybe I should stick around here more often.")
                    Else
                        Call SendNpcTalkTo(Index, NpcNum, "What? Peach's item? Oh yeah! I have that!", "But I am kind of hungry...mind getting me a Honey Super?")
                        Call PutVar(FilePath, GetPlayerName(Index), "ItemQuest12", "InProgress")
                    End If
            End Select
        Exit Sub
    ' Sixteenth Quest Script (Mushroom Balls)
        Case 27
            FilePath = App.Path & "\Scripts\Quests.ini"
            FavorStatus = GetVar(FilePath, GetPlayerName(Index), "ItemQuest13")
            
            Select Case FavorStatus
                Case "Done"
                    Call SendFavorTo(Index, "Find the Mushroom Balls!", "Favor Already Complete!", "Thanks for getting my Mushroom Balls before; you're the best!")
                Case "InProgress"
                    ' Check if the player has the Honey Super
                    If HasItem(Index, 190) >= 10 Then
                        Call SendFavorTo(Index, "Find the Mushroom Balls!", "Favor Complete!", "Oh, thank you soo much! Here, take this!")
                        
                        Call TakeItem(Index, 190, 10)
                        Call GiveItem(Index, 1, 300)
                        Call GiveItem(Index, 179, 3)
                        Call GiveItem(Index, 225, 1)
                        
                        Call PlayerMsg(Index, "You received 300 Coins!", YELLOW)
                        Call PlayerMsg(Index, "You received 3 Star Pieces!", YELLOW)
                        Call PlayerMsg(Index, "You received a Blue Candy!", YELLOW)
                        Call PutVar(FilePath, GetPlayerName(Index), "ItemQuest13", "Done")
                    Else
                        Call SendFavorTo(Index, "Find the Mushroom Balls!", "Need Mushroom Balls!", "Did you get my Mushroom Balls yet? Drill into the ground when you see a '!' above your head to get them.")
                    End If
                Case Else
                    Call SendFavorTo(Index, "Find the Mushroom Balls!", "Favor", "Oh no! I lost all my Mushroom Balls! Can you help me?", "They are hidden in the ground around the city. Drill into the ground when you see a '!' above your head.")
                    Call PutVar(FilePath, GetPlayerName(Index), "ItemQuest13", "InProgress")
            End Select
        Exit Sub
    ' Seventeenth Quest Script (Getting Salestoad's Stuff again)
        Case 28
            FilePath = App.Path & "\Scripts\Quests.ini"
            FavorStatus = GetVar(FilePath, GetPlayerName(Index), "ItemQuest14")
            
            If FavorStatus <> "Done" Then
                If FavorStatus <> "InProgress" And FavorStatus <> "InProgress2" Then
                    Call SendFavorTo(Index, "The Salestoad's Lost Package", "Favor", "Oh, it's you again! One of these fiends took my stuff and hid it in some weird plant! Can you please get it back for me?")
                    Call PutVar(FilePath, GetPlayerName(Index), "ItemQuest14", "InProgress")
                    Exit Sub
                Else
                    If CountInvItemNum(Index, 266) < 1 Then
                        Call SendFavorTo(Index, "The Salestoad's Lost Package", "Retrieve Salestoad's Stuff!", "Did you get my stuff back yet? I'll be here waiting.")
                    Else
                        Call TakeItem(Index, 266, 1)
                        Call PutVar(FilePath, GetPlayerName(Index), "ItemQuest14", "Done")
                        Call SendFavorTo(Index, "The Salestoad's Lost Package", "Favor Complete!", "Oh, thank you so much! Stop by anytime to check out the stuff I have!")
                        Call PlayerMsg(Index, "You can now buy items from the Salestoad!", WHITE)
                    End If
                End If
            Else
                Call SendTrade(Index, 22)
            End If
        Exit Sub
    ' Eighteenth Quest Script (Chef Bean B.'s Dilemma)
        Case 29
            FilePath = App.Path & "\Scripts\Quests.ini"
            FavorStatus = GetVar(FilePath, GetPlayerName(Index), "ItemQuest15")
            
            Select Case FavorStatus
                Case "Done"
                    ' Refreshshroom Favor
                    If HasItem(Index, 281) = 1 Then
                        If GetVar(FilePath, GetPlayerName(Index), "ItemQuest19-Chef") <> "Done" Then
                            Call TakeItem(Index, 281, 1)
                            Call SendNpcTalkTo(Index, NpcNum, "Oh, this ingredient looks particularly interesting...", "Hmm, I'll make something fantastic with it! Let me just add a secret ingredient...there!")
                            
                            Call GiveItem(Index, 282, 1)
                            Call PlayerMsg(Index, "You got a Memory Cake!", YELLOW)
                            
                            Call PutVar(FilePath, GetPlayerName(Index), "ItemQuest19-Chef", "Done")
                            
                            ' Add the Memory Cake to the player's recipe log
                            Call PutVar(App.Path & "\Scripts\" & "Recipes.ini", "47", GetPlayerName(Index), "Has")
                        Else
                            Call SendDataTo(Index, SPackets.Scookitem & SEP_CHAR & NpcNum & END_CHAR)
                        End If
                    Else
                        Call SendDataTo(Index, SPackets.Scookitem & SEP_CHAR & NpcNum & END_CHAR)
                    End If
                Case "InProgress"
                    If CanTake(Index, 277, 1) Then
                        Call TakeItem(Index, 277, 1)
                        Call SendNpcTalkTo(Index, NpcNum, "No way; it's my Cooking Pan! Oh, how I missed it so much!", "And for you, I'll be able to make you anything now that I have my Cooking Pan again! Thanks a lot!")
                        
                        Call PutVar(FilePath, GetPlayerName(Index), "ItemQuest15", "Done")
                        Call AddCastleTownFavorComplete(Index)
                    Else
                        Call SendNpcTalkTo(Index, NpcNum, "No, no, no, no!!! Where'd my precious Cooking Pan go now?!", "I cleared out my entire house and STILL can't find it! Someone must have stolen it!")
                    End If
                Case Else
                    Call SendNpcTalkTo(Index, NpcNum, "No, no, no, no!!! Where'd my precious Cooking Pan go now?!", "I cleared out my entire house and STILL can't find it! Someone must have stolen it!")
                    Call PutVar(FilePath, GetPlayerName(Index), "ItemQuest15", "InProgress")
            End Select
        Exit Sub
    ' Nineteenth Quest Script (Bean Fruit Finder)
        Case 30
            FilePath = App.Path & "\Scripts\Quests.ini"
            FavorStatus = GetVar(FilePath, GetPlayerName(Index), "ItemQuest16")
            
            Select Case FavorStatus
                Case "Done"
                    Call SendFavorTo(Index, "Bean Fruit Finder", "Favor Already Complete!", "Thanks again for all your help! I've been eating the Bean Fruit, and they're absolutely delicious!")
                Case "InProgress"
                    Dim i As Byte
                    
                    For i = 0 To 9
                        If CanTake(Index, 285 + i, 1) = False Then
                            Call SendFavorTo(Index, "Bean Fruit Finder", "Need Bean Fruits!", "Oh, it doesn't seem like you have the 10 kinds of Bean Fruit yet. Look for them in certain flower and plant patterns; they're hard to miss.")
                            
                            Exit Sub
                        End If
                    Next i
                    
                    For i = 0 To 9
                        Call TakeItem(Index, 285 + i, 1)
                    Next i
                    
                    If GetFreeSlots(Index) >= 3 Then
                        Call SendFavorTo(Index, "Bean Fruit Finder", "Favor Complete", "Thank you very much for getting me all the fruit I've been wanting to eat for a long time! I know a good reward for an excellent person such as yourself!")
                        
                        Call GiveItem(Index, 271, 150)
                        Call GiveItem(Index, 284, 1)
                        Call GiveItem(Index, 323, 1)
                        
                        Call PlayerMsg(Index, "You got Ultra Nuts, a Wrench, 150 Beanbean Coins, and 300 experience points!", YELLOW)
                        
                        Call SetPlayerExp(Index, GetPlayerExp(Index) + 300)
                        Call CheckPlayerLevelUp(Index)
                        
                        Call PutVar(FilePath, GetPlayerName(Index), "ItemQuest16", "Done")
                        Call AddCastleTownFavorComplete(Index)
                    Else
                        Call SendFavorTo(Index, "Bean Fruit Finder", "Favor Complete", "Thank you very much for getting me all the fruit I've been wanting to eat for a long time! I know a good reward for an excellent person such as yourself!", "It seems like your inventory is full. Please come back to get your reward when you can hold it.")
                    End If
                Case Else
                    Call SendFavorTo(Index, "Bean Fruit Finder", "Favor", "Oh, you look like a nice person. Can you help an old man out? I want you to get me all 10 Bean Fruit in Beanbean Outskirts. You can find them in flower and plant patterns.")
                    Call PutVar(FilePath, GetPlayerName(Index), "ItemQuest16", "InProgress")
            End Select
        Exit Sub
    ' 20th Quest Script (Lonely Bean)
        Case 31
            FilePath = App.Path & "\Scripts\Quests.ini"
            FavorStatus = GetVar(FilePath, GetPlayerName(Index), "ItemQuest17")
            
            Select Case FavorStatus
                Case "Done"
                    Call SendNpcTalkTo(Index, NpcNum, "Hey, how are you doing? Peasley and I are getting along just fine!")
                Case "InProgress"
                    If CanTake(Index, 318, 1) Then
                        If GetFreeSlots(Index) >= 2 Then
                            Call TakeItem(Index, 318, 1)
                            Call SendNpcTalkTo(Index, NpcNum, "Oh, you found one! Thank you so much! I think I'll name it...Peasley! Here, take this!")

                            Call GiveItem(Index, 179, 5)
                            Call GiveItem(Index, 321, 1)
                            
                            Call PlayerMsg(Index, "You got 5 Star Pieces and a Rock Candy!", YELLOW)
                            
                            Call PutVar(FilePath, GetPlayerName(Index), "ItemQuest17", "Done")
                            Call AddCastleTownFavorComplete(Index)
                        Else
                            Call SendNpcTalkTo(Index, NpcNum, "Oh, you found one! Thank you so much! I think I'll name it...Peasley! Here, take this!", "It seems like your inventory is full. Please come back to get your reward when you can hold it.")
                        End If
                    Else
                        Call SendNpcTalkTo(Index, NpcNum, "Did you get me a pet rock yet? I'll be waiting!")
                    End If
                Case Else
                    Call SendNpcTalkTo(Index, NpcNum, "Hey, you there! I've been feeling kinda lonely, and well...I've always wanted a pet rock!", "Can you get me one? Pretty please?!")
                    Call PutVar(FilePath, GetPlayerName(Index), "ItemQuest17", "InProgress")
            End Select
        Exit Sub
    ' Armored Koopa
        Case 32
            Call SendNpcTalkYesNoTo(Index, NpcNum, "Nyeck Nyeck! They say I'm the toughest fighter around. You look pretty strong; will you challenge me?", "Nyeck Nyeck! Show me your best!", "Nyeck Nyeck! Come back when you toughen up!")
        Exit Sub
    ' 21st Quest Script (Sewer Blockade!)
        Case 33
            FilePath = App.Path & "\Scripts\Quests.ini"
            FavorStatus = GetVar(FilePath, GetPlayerName(Index), "ItemQuest18")
            
            If FavorStatus <> "Done" Then
                If CanTake(Index, 323, 1) = True Then
                    Call TakeItem(Index, 323, 1)
                    Call SendNpcTalkTo(Index, NpcNum, "Oh, there it is! My wrench! What? You're saying it's not a wrench!?", "Anyway, thanks a lot! Here, take this!")
                        
                    Call GiveItem(Index, 271, 100)
                        
                    Call PlayerMsg(Index, "You got 100 Beanbean Coins and 500 experience points!", YELLOW)
                        
                    Call SetPlayerExp(Index, GetPlayerExp(Index) + 500)
                    Call CheckPlayerLevelUp(Index)
                        
                    Call PutVar(FilePath, GetPlayerName(Index), "ItemQuest18", "Done")
                    Call AddCastleTownFavorComplete(Index)
                Else
                    Call SendNpcTalkTo(Index, NpcNum, "The sewers are currently unavailable thanks to this rock. How did it even get here, anyway?", "I'll need my wrench to work this out; do you think you can find it for me?")
                End If
            Else
                Call SendNpcTalkTo(Index, NpcNum, "I'll get working on removing this rock right away. I'll finish eventually...")
            End If
        Exit Sub
    ' Sewer Blockade's Friend
        Case 34
            FilePath = App.Path & "\Scripts\Quests.ini"
            
            If GetVar(FilePath, GetPlayerName(Index), "ItemQuest18") <> "Done" Then
                Call SendNpcTalkTo(Index, NpcNum, "We tried everything so far and couldn't get this rock out of the way.", "If we can find his wrench then maybe he can clear the pipe.")
            Else
                Call SendNpcTalkTo(Index, NpcNum, "Great, now we can finally remove this annoying rock! Thanks for your help!")
            End If
        Exit Sub
    ' Random Guy in West Beanbean Outskirts
        Case 35
            FilePath = App.Path & "\Scripts\Quests.ini"
            FavorStatus = GetVar(FilePath, GetPlayerName(Index), "ItemQuest19")
            
            If FavorStatus <> "Done" Then
                If CanTake(Index, 282, 1) = True Then
                    Call TakeItem(Index, 282, 1)
                    Call SendNpcTalkTo(Index, NpcNum, "Hmm? You want me to eat this? *Eats cake*", "Wow, I remember everything now! I must've hit my head really hard on a rock or something...Here, look at what I can offer you!")
                    
                    Call PutVar(FilePath, GetPlayerName(Index), "ItemQuest19", "Done")
                    
                    Call SendTrade(Index, 28)
                Else
                    Call SendNpcTalkTo(Index, NpcNum, "...What...Where am I?")
                End If
            Else
                Call SendTrade(Index, 28)
            End If
        Exit Sub
    ' Doctor in Castle Town
        Case 36
            Call SendNpcTalkYesNoTo(Index, NpcNum, "Hello, you look hurt. Would you like me to nurse you back to full health for 10 coins?", "Okay, I'd like you to rest for a minute.", "Oh, so you're feeling well? Come back if you need anything.")
        Exit Sub
    ' Guard in East Beanbean Outskirts
        Case 37
            FavorStatus = GetVar(App.Path & "\Scripts\CastleTownFavors.ini", GetPlayerName(Index), "FavorsCompleted")
            
            If FavorStatus <> vbNullString Then
                Dim FavorStatusNum As Byte
                
                FavorStatusNum = CByte(FavorStatus)
                
                If FavorStatusNum >= 4 Then
                    Call SendNpcTalkTo(Index, NpcNum, "Hey, you did a lot for the people in Castle Town! Okay, I'll let you through.")
                    
                    If GetPlayerY(Index) = 28 Then ' Talking to the Guard facing down
                        Call SetPlayerY(Index, 30)
                    ElseIf GetPlayerY(Index) = 30 Then ' Talking to the Guard facing up
                        Call SetPlayerY(Index, 28)
                    End If
                    
                    Call SendPlayerXY(Index)
                Else
                    Call SendNpcTalkTo(Index, NpcNum, "Sorry, but no one is allowed beyond this point. It's Castle Town regulations.")
                End If
            Else
                Call SendNpcTalkTo(Index, NpcNum, "Sorry, but no one is allowed beyond this point. It's Castle Town regulations.")
            End If
        Exit Sub
    ' Anything else
        Case Else
            Call PlayerMsg(Index, "No NPC script found. Please contact an admin to solve this problem.", WHITE)
        Exit Sub
    End Select
End Sub

' Sub for adding points after killing the Monty Moles in Whack-A-Monty
Sub KillMonty(ByVal Index As Long, ByVal GameNum As Integer, ByVal NpcNum As Long)
    Dim FilePath As String
    Dim POINTS As Integer

    FilePath = App.Path & "\Scripts\" & "Whack.ini"
    POINTS = CInt(GetVar(FilePath, "Game" & GameNum, "Points"))
    
    If NpcNum = 47 Then
        Call PutVar(FilePath, "Game" & GameNum, "Points", (POINTS + 1))
        Call BattleMsg(Index, "Points: " & (POINTS + 1), YELLOW, 0)
    ElseIf NpcNum = 48 Then
        Call PutVar(FilePath, "Game" & GameNum, "Points", (POINTS - 1))
        Call BattleMsg(Index, "Points: " & (POINTS - 1), YELLOW, 0)
    End If
End Sub

Function CountInvItemNum(ByVal Index As Long, ByVal ItemNum As Long) As Long
    Dim i As Long
    
    If ItemIsStackable(ItemNum) = True Then
        Exit Function
    End If
    
    CountInvItemNum = 0
    
    For i = 1 To GetPlayerMaxInv(Index)
        If GetPlayerInvItemNum(Index, i) = ItemNum Then
            CountInvItemNum = CountInvItemNum + 1
        End If
    Next i
End Function

Sub OnArenaDeath(ByVal Attacker As Long, ByVal Victim As Long)
    Call PlayerWarp(Attacker, 218, 8, 25)
    Call PlayerWarp(Victim, 218, 21, 25)
End Sub

Sub QueryBox(ByVal Index As Long, ByVal Response As String, ByVal PromptNum As Long)
    Dim RealPassword As String
    
    Select Case PromptNum
     ' Password Changing NPC
        Case 0
            RealPassword = Trim$(GetVar(App.Path & "\SMBOAccounts\" & GetPlayerLogin(Index) & "_Info.ini", "ACCESS", "Password"))
            Select Case Response
              ' Asks player to input new password
                Case RealPassword
                    Call PlayerQueryBox(Index, "Please enter your new password.", 1)
              ' Rejects password if it's incorrect
                Case Else
                    Call SendMsgBoxTo(Index, "Incorrect Password!", "I'm sorry, but that password is incorrect. Please visit me when you know your password and want to change it.")
            End Select
        Exit Sub
    ' Password Changing NPC (continued)
        Case 1
            RealPassword = Trim$(GetVar(App.Path & "\SMBOAccounts\" & GetPlayerLogin(Index) & "_Info.ini", "ACCESS", "Password"))
            Select Case Response
              ' Asks you to input a new password if the password you inputted is the same as the current one
                Case RealPassword
                    Call PlayerQueryBox(Index, "I'm sorry, but that's your current password. Please choose a different password.", 1)
                Case Else
              ' Sets password length requirement; must be at least 3 characters long
                    If Len(Response) < 3 Then
                        Call PlayerQueryBox(Index, "Your new password must be at least 3 characters in length! Please choose another password.", 1)
                    Else
            ' Changes password
                        Player(Index).Password = Response
                        Call SendMsgBoxTo(Index, "Password Changed Successfully!", "Your password has been changed to: " & Response & ". Please come back to me whenever you want to change your password again.")
                    End If
            End Select
        Exit Sub
    ' Group Creating NPC
        Case 2
            If Len(Response) > 0 Then
                If Len(GetPlayerGuild(Index)) = 0 Then
                    If CanTake(Index, 1, 5000) Then
                        Call SetPlayerGuild(Index, Response)
                        Call SetPlayerGuildAccess(Index, 4)
                        Call TakeItem(Index, 1, 5000)
                        Call SendPlayerData(Index)
                        Call SendMsgBoxTo(Index, "Group Created Successfully!", "Thank you very much! Enjoy your Group!")
                    Else
                        Call SendMsgBoxTo(Index, "Not Enough Coins!", "I'm sorry, but you don't have enough Coins to form your own Group.")
                    End If
                Else
                    Call SendMsgBoxTo(Index, "Already In Group!", "In order to create your own Group, you must not already be in one!")
                End If
            Else
                Call SendMsgBoxTo(Index, "Invalid Group Name!", "Please select a Group name that is at least one character long.")
            End If
        Exit Sub
    ' Koopa Troopa Explorer for Tenth Favor
        Case 3
            If Val(Response) <> 7 And LCase$(Response) <> "seven" Then
                Call SendFavorTo(Index, "The Explorer", "Wrong Number!", "Really? It didn't seem like there were that many statues when I was on my way here. Try again.")
            Else
                Call SendFavorTo(Index, "The Explorer", "Favor Complete!", "That seems about right. Thanks a lot! Here, take this. I found it along my travels.")
                Call PutVar(App.Path & "\Scripts\Quests.ini", GetPlayerName(Index), "OtherQuest1", "InProgress2")
                If GetFreeSlots(Index) > 0 Then
                    Call GiveItem(Index, 131, 1)
                    Call PlayerMsg(Index, "You got a F-Defense!", YELLOW)
                    Call PutVar(App.Path & "\Scripts\Quests.ini", GetPlayerName(Index), "OtherQuest1", "Done")
                Else
                    Call PlayerMsg(Index, "You don't have enough inventory space to take the reward! Come back to him when you have room.", BRIGHTRED)
                End If
            End If
        Exit Sub
    ' Yoshi color changing NPC
        Case 4
            Call ChangeYoshiColor(Index, LCase$(Response))
        Exit Sub
    ' Toad color changing NPC
        Case 5
            Call ChangeToadColor(Index, LCase$(Response))
        Exit Sub
    ' Cake Mix selling Toad
        Case 6
            If LCase$(Response) = "yes" Then
                If GetFreeSlots(Index) > 0 Then
                    If CanTake(Index, 47, 1) Then
                        Call TakeItem(Index, 47, 1)
                        Call GiveItem(Index, 145, 1)
                        Call SendNpcTalkTo(Index, 83, "Thanks! Come back any time you want some more!")
                    Else
                        Call SendNpcTalkTo(Index, 83, "Sorry, you don't have enough Dim Star Pieces to buy this.", "Please come back another time.")
                    End If
                Else
                    Call SendNpcTalkTo(Index, 83, "Oh, your inventory is full. Please come back once you have room for some Cake Mix.")
                End If
            Else
                Call SendNpcTalkTo(Index, 83, "Oh, it doesn't seem like you want to buy some Cake Mix. Feel free to come back whenever you want some!")
            End If
        Exit Sub
    ' Orange Bob-Omb selling Toad
        Case 7
            If LCase$(Response) = "yes" Then
                If GetFreeSlots(Index) > 0 Then
                    If CanTake(Index, 47, 50) Then
                        Call TakeItem(Index, 47, 50)
                        Call GiveItem(Index, 162, 1)
                        Call SendNpcTalkTo(Index, 84, "Thanks! Come back any time you need another!")
                    Else
                        Call SendNpcTalkTo(Index, 84, "Sorry, you don't have enough Dim Star Pieces to buy this.", "Please come back another time.")
                    End If
                Else
                    Call SendNpcTalkTo(Index, 84, "Oh, your inventory is full. Please come back once you have room for the Bob-Omb.")
                End If
            Else
                Call SendNpcTalkTo(Index, 84, "Oh, it doesn't seem like you want to buy the Bob-Omb. Please come back whenever you need it!")
            End If
        Exit Sub
    ' Koopa Troopa Explorer for Fourteenth Favor
        Case 8
            If Val(Response) <> 23 Then
                Call SendFavorTo(Index, "The Explorer 2", "Wrong Number!", "Really? It didn't seem like there were that many buildings here. Try again.")
            Else
                Call SendFavorTo(Index, "The Explorer 2", "Favor Complete!", "That seems about right once again! Thanks a lot! Here, take this. I found it along my travels.")
                Call PutVar(App.Path & "\Scripts\Quests.ini", GetPlayerName(Index), "OtherQuest2", "InProgress2")
                
                If GetFreeSlots(Index) > 0 Then
                    Call GiveItem(Index, 216, 1)
                    Call PlayerMsg(Index, "You got a F-Attack!", YELLOW)
                    Call PutVar(App.Path & "\Scripts\Quests.ini", GetPlayerName(Index), "OtherQuest2", "Done")
                Else
                    Call PlayerMsg(Index, "You don't have enough inventory space to take the reward! Come back to him when you have room.", BRIGHTRED)
                End If
            End If
        Exit Sub
    End Select
End Sub

Sub ChangeYoshiColor(ByVal Index As Long, ByVal Response As String)
    Select Case Response
        Case "green"
            Call SetPlayerSprite(Index, 4)
        Case "red"
            Call SetPlayerSprite(Index, 42)
        Case "blue"
            Call SetPlayerSprite(Index, 43)
        Case "light blue"
            Call SetPlayerSprite(Index, 44)
        Case "yellow"
            Call SetPlayerSprite(Index, 45)
        Case "pink"
            Call SetPlayerSprite(Index, 46)
        Case "black"
            Call SetPlayerSprite(Index, 47)
        Case "white"
            Call SetPlayerSprite(Index, 48)
        Case Else
            Call SendNpcTalkTo(Index, 77, "I'm sorry, but we don't have that color. Please read the sign to find out which colors we have.")
            Exit Sub
    End Select
    
    ' Set the player's temp sprite to the new sprite
    Call SetPlayerTempSprite(Index, GetPlayerSprite(Index))
    
    Call SendDataToMap(GetPlayerMap(Index), SPackets.Schecksprite & SEP_CHAR & Index & SEP_CHAR & GetPlayerSprite(Index) & END_CHAR)
    Call SendNpcTalkTo(Index, 77, "There you go! Wow, you look great! Come back whenever you want to spray paint yourself a different color.")
End Sub

Sub ChangeToadColor(ByVal Index As Long, ByVal Response As String)
    Select Case Response
        Case "standard"
            Call SetPlayerSprite(Index, 5)
        Case "red"
            Call SetPlayerSprite(Index, 49)
        Case "light green"
            Call SetPlayerSprite(Index, 50)
        Case "blue"
            Call SetPlayerSprite(Index, 51)
        Case "yellow"
            Call SetPlayerSprite(Index, 93)
        Case Else
            Call SendNpcTalkTo(Index, 76, "I'm sorry, but we don't have that color. Please read the sign to find out which colors we have.")
            Exit Sub
    End Select
    
    ' Set the player's temp sprite to the new sprite
    Call SetPlayerTempSprite(Index, GetPlayerSprite(Index))
    
    Call SendDataToMap(GetPlayerMap(Index), SPackets.Schecksprite & SEP_CHAR & Index & SEP_CHAR & GetPlayerSprite(Index) & END_CHAR)
    Call SendNpcTalkTo(Index, 76, "There you go! Wow, you look great! Come back whenever you want to spray paint yourself a different color.")
End Sub

Sub PlayerLevelUp(ByVal Index As Long)
    Dim TotalExp As Long

    Do While GetPlayerExp(Index) >= GetPlayerNextLevel(Index)
        TotalExp = GetPlayerExp(Index) - GetPlayerNextLevel(Index)
        Call SetPlayerLevel(Index, GetPlayerLevel(Index) + 1)

        Call SetPlayerPOINTS(Index, GetPlayerPOINTS(Index) + 1)
        Call SetPlayerExp(Index, TotalExp)
    Loop

    Call PlayerMsg(Index, "Congratulations, you've leveled up!", BRIGHTBLUE)
End Sub

Sub UsingStatPoints(ByVal Index As Long, ByVal PointType As Integer)
    Dim X As Integer
    
    Select Case PointType
        Case 0
            Call SetPlayerSTR(Index, Player(Index).Char(Player(Index).CharNum).STR + 1)
            Call SetPlayerMaxSTR(Index, Player(Index).Char(Player(Index).CharNum).MAXSTR + 1)
            Call PlayerMsg(Index, "You've increased your Attack!", WHITE)
        Case 1
            Call SetPlayerDEF(Index, Player(Index).Char(Player(Index).CharNum).DEF + 1)
            Call SetPlayerMaxDEF(Index, Player(Index).Char(Player(Index).CharNum).MAXDEF + 1)
            Call PlayerMsg(Index, "You've increased your Defense!", WHITE)
        Case 2
            Call SetPlayerStache(Index, Player(Index).Char(Player(Index).CharNum).Magi + 1)
            Call SetPlayerMaxStache(Index, Player(Index).Char(Player(Index).CharNum).MAXStache + 1)
            Call PlayerMsg(Index, "You've increased your Stache!", WHITE)
        Case 3
            Call SetPlayerSPEED(Index, Player(Index).Char(Player(Index).CharNum).Speed + 1)
            Call SetPlayerMaxSpeed(Index, Player(Index).Char(Player(Index).CharNum).MAXSpeed + 1)
            Call PlayerMsg(Index, "You've increased your Speed!", WHITE)
        Case 4
            X = Val(GetVar(App.Path & "\Level Up.ini", GetPlayerName(Index), "HP"))
            Call PutVar(App.Path & "\Level Up.ini", GetPlayerName(Index), "HP", Int(X + 1))

            Call PlayerMsg(Index, "You've increased your HP!", WHITE)
        Case 5
            X = Val(GetVar(App.Path & "\Level Up.ini", GetPlayerName(Index), "FP"))
            Call PutVar(App.Path & "\Level Up.ini", GetPlayerName(Index), "FP", Int(X + 1))

            Call PlayerMsg(Index, "You've increased your FP!", WHITE)
    End Select

    ' Remove one point
    Call SetPlayerPOINTS(Index, GetPlayerPOINTS(Index) - 1)
End Sub

Sub PlayerPrompt(ByVal Index As Long, ByVal Prompt As Long, ByVal Value As Integer)
    ' If players select "Yes"
    If Prompt = 6 Then
        Select Case Value
            Case 0
                Call PlayerWarp(Index, 34, 15, 25)
                Call PlayerMsg(Index, "You chose to begin the daily event for today. Good luck!", YELLOW)
            Case 1
                Call GlobalMsg("This is case 1?", CYAN)
        End Select
    ' If players select "No"
    Else
        Select Case Value
            Case 0
                Call PlayerMsg(Index, "You will have another opportunity to try the daily event tomorrow.", GREEN)
            Case 1
                Call GlobalMsg("This is no case 1?", RED)
        End Select
    End If
End Sub

Sub Prompt(ByVal Index As Long, ByVal Question As String, ByVal Value As Long)
    Call SendDataTo(Index, SPackets.Sprompt & SEP_CHAR & Question & SEP_CHAR & Value & END_CHAR)
End Sub

Sub PlayerFirstStrike(ByVal Index As Long, ByVal MapNpcNum As Long)
    Call BattleMsg(Index, "You made the First Strike!!", YELLOW, 0)

    Call OnTurnBasedBattle(Index, MapNpcNum, True)
    Call PlayerTurn(Index, MapNpcNum)
    
    If MapNPC(GetPlayerMap(Index), MapNpcNum).HP > 0 Then
        Call DetermineBattleTurn(Index, GetPlayerMap(Index), MapNpcNum)
    End If
End Sub

Sub NpcFirstStrike(ByVal Index As Long, ByVal MapNpcNum As Long)
    Call BattleMsg(Index, "The enemy made the First Strike!!", RED, 0)
    
    Call OnTurnBasedBattle(Index, MapNpcNum, True)
    Call NpcTurn(Index, MapNpcNum)
    
    If GetPlayerHP(Index) > 0 Then
        Call DetermineBattleTurn(Index, GetPlayerMap(Index), MapNpcNum)
    End If
End Sub

Sub OnTurnBasedBattle(ByVal Index As Long, ByVal MapNpcNum As Long, Optional ByVal IsFirstStrike As Boolean = False, Optional ByVal PlayerOldX As Long = -5, Optional ByVal PlayerOldY As Long = -5)
    Dim MapNum As Long
    Dim Randomize As Integer
    
    MapNum = GetPlayerMap(Index)
    
    Call SetPlayerInBattle(Index, True)
    MapNPC(MapNum, MapNpcNum).InBattle = True
    Player(Index).TargetNPC = MapNpcNum
    MapNPC(MapNum, MapNpcNum).Target = Index
    
    If PlayerOldX >= 0 And PlayerOldY >= 0 Then
        Call SetPlayerOldX(Index, PlayerOldX)
        Call SetPlayerOldY(Index, PlayerOldY)
    Else
        Call SetPlayerOldX(Index, GetPlayerX(Index))
        Call SetPlayerOldY(Index, GetPlayerY(Index))
    End If
    
    MapNPC(MapNum, MapNpcNum).OldX = MapNPC(MapNum, MapNpcNum).X
    MapNPC(MapNum, MapNpcNum).OldY = MapNPC(MapNum, MapNpcNum).Y
    
    Call PlayerWarp(Index, MapNum, Map(MapNum).BootX, Map(MapNum).BootY)
    MapNPC(MapNum, MapNpcNum).X = GetPlayerX(Index) + 6
    MapNPC(MapNum, MapNpcNum).Y = GetPlayerY(Index)
    MapNPC(MapNum, MapNpcNum).Dir = DIR_LEFT
    Call SetPlayerDir(Index, DIR_RIGHT)
    Call SendMapNpcsToMap(MapNum)
    
    If IsFirstStrike = False Then
        Call SendTurnBasedBattle(Index, 1, MapNpcNum)
        
        Call DetermineBattleTurn(Index, MapNum, MapNpcNum)
    Else
        Call SendTurnBasedBattle(Index, 2, MapNpcNum)
    End If
End Sub

Sub TurnBasedBattle(ByVal Index As Long, ByVal Target As Long)
    If GetPlayerTurn(Index) = True And MapNPC(GetPlayerMap(Index), Target).Turn = False Then
        Call PlayerTurn(Index, Target)
    End If
    If MapNPC(GetPlayerMap(Index), Target).Turn = True And GetPlayerTurn(Index) = False Then
        Call NpcTurn(Index, Target)
    End If
End Sub

Sub DetermineBattleTurn(ByVal Index As Long, ByVal MapNum As Long, ByVal MapNpcNum As Long)
    Dim Randomize As Integer
    
    If GetPlayerSPEED(Index) > NPC(MapNPC(MapNum, MapNpcNum).num).Speed Then
        Call SetPlayerTurn(Index, True)
        Call StartPlayerTurn(Index, MapNpcNum)
        Exit Sub
    ElseIf GetPlayerSPEED(Index) < NPC(MapNPC(MapNum, MapNpcNum).num).Speed Then
        MapNPC(MapNum, MapNpcNum).Turn = True
        Call StartNpcTurn(Index)
        Exit Sub
    ElseIf GetPlayerSPEED(Index) = NPC(MapNPC(MapNum, MapNpcNum).num).Speed Then
        Randomize = Int(Rand(1, 2))
            
        Select Case Randomize
            Case 1
                Call SetPlayerTurn(Index, True)
                Call StartPlayerTurn(Index, MapNpcNum)
                Exit Sub
            Case 2
                MapNPC(MapNum, MapNpcNum).Turn = True
                Call StartNpcTurn(Index)
                Exit Sub
        End Select
    End If
End Sub

Sub PlayerTurn(ByVal Index As Long, ByVal Target As Long)
    Dim Damage As Long
    Dim CritHit As Boolean
    Dim packet As String
    
    ' Adds critical hits to turn-based battles
    If Not CanPlayerCriticalHit(Index) Then
        Damage = GetPlayerDamage(Index) - Int(NPC(MapNPC(GetPlayerMap(Index), Target).num).DEF)
        packet = SPackets.Ssound & SEP_CHAR & "attack" & END_CHAR
    Else
        Damage = Int((GetPlayerDamage(Index) - Int(NPC(MapNPC(GetPlayerMap(Index), Target).num).DEF)) * 1.5)
        packet = SPackets.Ssound & SEP_CHAR & "critical" & END_CHAR
        
        Call BattleMsg(Index, "Critical hit!", BRIGHTGREEN, 0)
        
        CritHit = True
    End If
    
    Damage = DamageUpDamageDown(Index, Damage)
    
    ' Randomizes damage
    Damage = Int(Rand(Damage - 2, Damage + 2))
    
    ' Make it so you cannot hit less than 1 if your attack is greater than the target's defense
    If Damage <= 0 And GetPlayerSTR(Index) > Int(NPC(MapNPC(GetPlayerMap(Index), Target).num).DEF) Then
        Damage = 1
    End If
    
    ' The Armored Koopa cannot be hurt with normal attacks
    If MapNPC(GetPlayerMap(Index), Target).num = 194 Then
        Damage = 0
    End If
    
    If Damage <= 0 Then
        Damage = 0
        
        If CritHit = False Then
            packet = SPackets.Ssound & SEP_CHAR & "miss" & END_CHAR
            
            Call PlayerMsg(Index, "Your attack was too weak to harm the enemy!", WHITE)
        End If
    End If
    
    Call AttackNpc(Index, Target, Damage)
    Call SendDataTo(Index, packet)
    Call StartNpcTurn(Index)
End Sub

Sub NpcTurn(ByVal Index As Long, ByVal Target As Long)
    Dim Damage As Long
    
    ' Check for block chance
    If Not CanPlayerBlockHit(Index) Then
        Damage = Int(NPC(MapNPC(GetPlayerMap(Index), Target).num).STR - GetPlayerProtection(Index))
        
        Call NpcAttackPlayer(Target, Index, Damage)
    Else
        Call BattleMsg(Index, "You blocked the " & Trim$(NPC(MapNPC(GetPlayerMap(Index), Target).num).Name) & "'s hit!", BRIGHTCYAN, 0)
        Call SendDataTo(Index, SPackets.Ssound & SEP_CHAR & "miss" & END_CHAR)
    End If
    
    Call StartPlayerTurn(Index, Target)
End Sub

Sub EndBattle(ByVal Index As Long, ByVal MapNum As Long, ByVal MapNpcNum As Long, ByVal OldX As Integer, ByVal OldY As Integer)
    Call SetPlayerInBattle(Index, False)
    Call SetPlayerTurn(Index, False)
    
    Call SetPlayerX(Index, OldX)
    Call SetPlayerY(Index, OldY)
    Call SendPlayerXY(Index)
    
    Call SendTurnBasedBattle(Index, 0, MapNpcNum)
    Call SetPlayerRecoverTime(Index, GetTickCount)
    
    ' Set that the player is not in the victory animation
    IsInVictoryAnim(Index) = False
    
    ' Check for level up
    Call CheckPlayerLevelUp(Index)
        
    ' Check for level up party member
    If GetPlayerPartyNum(Index) > 0 Then
        Dim n As Long, PartyMember As Long
        
        For n = 1 To MAX_PARTY_MEMBERS
            PartyMember = GetPartyMember(GetPlayerPartyNum(Index), n)
            
            If PartyMember > 0 Then
                Call CheckPlayerLevelUp(PartyMember)
            End If
        Next n
    End If
End Sub

Sub EndBattleVictory(ByVal Index As Long, ByVal MapNpcNum As Long)
    Dim packet As String
    Dim i As Integer
    
    ' Set that the player is in the victory animation
    IsInVictoryAnim(Index) = True
    
    packet = SPackets.Sturnbasedvictory & SEP_CHAR & MapNpcNum
    
    For i = 1 To 7
        packet = packet & SEP_CHAR & VictoryInfo(Index, i)
    Next
    
    packet = packet & END_CHAR

    Call SendDataTo(Index, packet)
    
    ' Get the NPC out of battle
    MapNPC(GetPlayerMap(Index), MapNpcNum).InBattle = False
    MapNPC(GetPlayerMap(Index), MapNpcNum).Turn = False
End Sub

Sub TurnBasedNpcAttackPlayer(ByVal Index As Long, ByVal MapNum As Long, ByVal MapNpcNum As Long, ByVal Damage As Long)
    Dim Sound As Byte
    
    Sound = 0
    
    ' Make a 1/5 chance of NPCs hitting a 0 if the player's Defense is higher than their Attack (in a turn-based battle)
    If Damage = 1 Then
        If GetPlayerDEF(Index) >= NPC(MapNPC(MapNum, MapNpcNum).num).STR Then
            Dim Random As Integer
            
            Random = Rand(1, 4)
            
            If Random = 1 Then
                Damage = 0
            End If
        End If
    End If
    
    If Damage >= GetPlayerHP(Index) Then
        Call TurnBasedDeath(Index, MapNum, MapNpcNum, Damage)
    Else
        If GetPlayerHP(Index) > 5 And (GetPlayerHP(Index) - Damage) <= 5 Then
            Sound = 1
        End If
        
        Call SetPlayerHP(Index, GetPlayerHP(Index) - Damage)
        Call SendHP(Index)
    End If
    
    Call SendDataTo(Index, SPackets.Sblitnpcdmg & SEP_CHAR & Damage & END_CHAR)
    
    If Damage <> 0 Then
        If Sound = 0 Then
            Call SendDataTo(Index, SPackets.Ssound & SEP_CHAR & "pain" & END_CHAR)
        End If
    Else
        Call SendDataTo(Index, SPackets.Ssound & SEP_CHAR & "miss" & END_CHAR)
    End If
End Sub

Sub StartNpcTurn(ByVal Index As Long)
    Call SetPlayerTurn(Index, False)
    Call SendDataTo(Index, SPackets.Sturnbasedtime & SEP_CHAR & 1 & SEP_CHAR & 1 & END_CHAR)
End Sub

Sub StartPlayerTurn(ByVal Index As Long, ByVal Target As Long)
    MapNPC(GetPlayerMap(Index), Target).Turn = False
    Call SendDataTo(Index, SPackets.Sturnbasedtime & SEP_CHAR & 1 & SEP_CHAR & 0 & END_CHAR)
End Sub

Sub TurnBasedDeath(ByVal Index As Long, ByVal MapNum As Long, ByVal MapNpcNum As Long, ByVal Damage As Long, Optional ByVal PoisonDeath As Boolean = False)
    ' See if the player can be revived by the Close Call badge
    If CheckCloseCall(Index) = True Then
        Call SendDataTo(Index, SPackets.Sblitnpcdmg & SEP_CHAR & Damage & END_CHAR)
            
        Exit Sub
    End If
        
    ' Allows a Life Shroom to revive a player
    If HasItem(Index, 43) = 1 Then
        Call LifeShroom(Index)
        Call SendDataTo(Index, SPackets.Sblitnpcdmg & SEP_CHAR & Damage & END_CHAR)
        
        Exit Sub
    End If
    
    If PoisonDeath = False Then
        Dim Name As String
        
        Name = Trim$(NPC(MapNPC(MapNum, MapNpcNum).num).Name)
        
        If FindNpcVowels(MapNPC(MapNum, MapNpcNum).num) = True Then
            Call PlayerMsg(Index, "You have been killed by an " & Name & ".", BRIGHTRED)
        Else
            Call PlayerMsg(Index, "You have been killed by a " & Name & ".", BRIGHTRED)
        End If
    Else
        Call PlayerMsg(Index, "You have been killed by the poison.", BRIGHTRED)
    End If
        
    ' Checks if the player should lose exp
    If Map(MapNum).Moral <> MAP_MORAL_NO_PENALTY Then
        Dim Exp As Long
    
        ' Calculate exp to give attacker
        Exp = (GetPlayerExp(Index) \ 6)
        ' Make sure we dont get fewer than 0 experience points
        If Exp < 0 Then
            Exp = 0
        End If
            
        ' Subtracts Exp
        Call SetPlayerExp(Index, GetPlayerExp(Index) - Exp)
        Call BattleMsg(Index, "Oh no! You've died! You lost " & Exp & " experience points.", BRIGHTRED, 0)
            
        Call SendEXP(Index)
    End If
        
    ' Set targets to 0
    MapNPC(MapNum, MapNpcNum).Target = 0
    Player(Index).TargetNPC = 0
        
    ' Get the player and NPC out of the battle
    MapNPC(MapNum, MapNpcNum).InBattle = False
    MapNPC(MapNum, MapNpcNum).Turn = False
    MapNPC(MapNum, MapNpcNum).X = MapNPC(MapNum, MapNpcNum).OldX
    MapNPC(MapNum, MapNpcNum).Y = MapNPC(MapNum, MapNpcNum).OldY
    Call SetPlayerInBattle(Index, False)
    Call SetPlayerTurn(Index, False)
    Call SendTurnBasedBattle(Index, 0, MapNpcNum)
    Call SendMapNpcsToMap(MapNum)
        
    ' Warp player away
    Call OnDeath(Index)
        
    ' Restore vitals
    Call SetPlayerHP(Index, GetPlayerMaxHP(Index))
    Call SetPlayerMP(Index, GetPlayerMaxMP(Index))
    Call SetPlayerSP(Index, GetPlayerMaxSP(Index))
    Call SendHP(Index)
    Call SendMP(Index)
    Call SendSP(Index)
        
    ' If the player the attacker killed was a pk then take it away
    If GetPlayerPK(Index) = YES Then
        Call SetPlayerPK(Index, NO)
        Call SendPlayerData(Index)
        Call SendPlayerXY(Index)
    End If
End Sub

Sub StatIncrease(ByVal Index As Long, ByVal Stat As Integer)
    If IsConnected(Index) = False Or IsPlaying(Index) = False Then
        Exit Sub
    End If
    
    Call GetRidOfTimer(Index, 1, Stat)

    Call PlayerMsg(Index, "The effects of your special attack have worn off.", WHITE)
    Call PutVar(App.Path & "\Scripts\" & "StatIncreases.ini", GetPlayerName(Index), Trim$(GetSpellStatName(Stat)), "No")

     Select Case Stat
    ' Attack
       Case 0
           Call SetPlayerSTR(Index, Player(Index).Char(Player(Index).CharNum).MAXSTR)
    ' Defense
       Case 1
           Call SetPlayerDEF(Index, Player(Index).Char(Player(Index).CharNum).MAXDEF)
    ' Speed
       Case 2
           Call SetPlayerSPEED(Index, Player(Index).Char(Player(Index).CharNum).MAXSpeed)
    ' Stache
       Case 3
           Call SetPlayerStache(Index, Player(Index).Char(Player(Index).CharNum).MAXStache)
     End Select
     
     Call SendStats(Index)
End Sub

' Timer for Steal the Shroom
Sub STSPlayTime(ByVal Index As Long)
    Dim FilePath As String, PlayerNum As String
    Dim i As Long, R As Long, Q As Long
    Dim Point As Integer, POINTS As Integer, TimeLeft As Integer

    Call GetRidOfTimer(Index, 0)
    
    FilePath = STSPath
    TimeLeft = CInt(GetVar(FilePath, "GameTime", "TimeLeft")) - 20

    ' Continues game
    If TimeLeft > 0 Then
        Call PutVar(FilePath, "GameTime", "TimeLeft", CStr(TimeLeft))
        ' Sends time left to players every 20 seconds
        Call MapBattleMsg(33, "Time remaining: " & TimeLeft & " seconds", WHITE, 1)
                
        Point = CInt(GetVar(FilePath, "Points", "Red"))
        POINTS = CInt(GetVar(FilePath, "Points", "Blue"))
        ' Sends points to players every 20 seconds
        Call MapBattleMsg(33, "Red Team: " & Point, YELLOW, 0)
        Call MapBattleMsg(33, "Blue Team: " & POINTS, YELLOW, 0)
        ' Start timer again
        Call AddNewTimer(Index, 0, 20000)
    ' Exits game when time is up, and gives out rewards
    ElseIf TimeLeft <= 0 Then
        Point = CInt(GetVar(FilePath, "Points", "Red"))
        POINTS = CInt(GetVar(FilePath, "Points", "Blue"))
    
        For i = 1 To 4
            PlayerNum = CStr(i)
            
            Q = FindPlayer(GetVar(FilePath, "Red", PlayerNum))
            R = FindPlayer(GetVar(FilePath, "Blue", PlayerNum))
            
            ' Rewards and resetting game for Red
            If IsPlaying(Q) Then
                Call PlayerWarp(Q, 31, 12 + i, 18)
                Call SetPlayerPK(Q, NO)
                Call SendPlayerData(Q)
                Call SendPlayerXY(Q)
                
                ' Gives out rewards
                If Point > POINTS Then
                    Call PlayerMsg(Q, "Congratulations! Your team won! You receive 2 Reward Coins!", BLUE)
                    Call GiveItem(Q, 51, 2)
                ElseIf POINTS > Point Then
                    Call PlayerMsg(Q, "Sorry, your team lost. Try again next time! You receive 1 Reward Coin!", BLUE)
                    Call GiveItem(Q, 51, 1)
                ElseIf Point = POINTS Then
                    Call PlayerMsg(Q, "Wow, it was a tie! You receive 1 Reward Coin!", BLUE)
                    Call GiveItem(Q, 51, 1)
                End If
                Call PlayerMsg(Q, "The final score was: Red: " & Point & "; Blue: " & POINTS, YELLOW)
            End If
            
            ' Rewards and resetting game for Blue
            If IsPlaying(R) Then
                Call PlayerWarp(R, 31, 12 + i, 17)
                Call SetPlayerPK(R, NO)
                Call SendPlayerData(R)
                Call SendPlayerXY(R)
                
                ' Gives out rewards
                If Point > POINTS Then
                    Call PlayerMsg(R, "Sorry, your team lost. Try again next time! You receive 1 Reward Coin!", BLUE)
                    Call GiveItem(R, 51, 1)
                ElseIf POINTS > Point Then
                    Call PlayerMsg(R, "Congratulations! Your team won! You receive 2 Reward Coins!", BLUE)
                    Call GiveItem(R, 51, 2)
                ElseIf Point = POINTS Then
                    Call PlayerMsg(R, "Wow, it was a tie! You receive 1 Reward Coin!", BLUE)
                    Call GiveItem(R, 51, 1)
                End If
                Call PlayerMsg(R, "The final score was: Red: " & Point & "; Blue: " & POINTS, YELLOW)
            End If
    
            If GetVar(FilePath, "Red", PlayerNum) <> vbNullString Then
                Call PutVar(FilePath, "Red", PlayerNum, "")
            End If
            
            If GetVar(FilePath, "Blue", PlayerNum) <> vbNullString Then
                Call PutVar(FilePath, "Blue", PlayerNum, "")
            End If
        Next i

        ' Resets game
        Call PutVar(FilePath, "Ingame", "Ingame", "No")
        Call PutVar(FilePath, "Flag", "Red", "NoFlag")
        Call PutVar(FilePath, "Flag", "Blue", "NoFlag")
        Call PutVar(FilePath, "Points", "Red", "0")
        Call PutVar(FilePath, "Points", "Blue", "0")
        Call PutVar(FilePath, "Team", "Red", "0")
        Call PutVar(FilePath, "Team", "Blue", "0")
        Call PutVar(FilePath, "GameTime", "TimeLeft", "0")
    End If
End Sub

Sub AttackDouble(ByVal Index As Long)
    Call GetRidOfTimer(Index, 5)

    If IsConnected(Index) = False Or IsPlaying(Index) = False Then
        Exit Sub
    End If

    Call PlayerMsg(Index, "The effects of your item have worn off.", WHITE)
    Call PutVar(App.Path & "\Scripts\" & "StatIncreases.ini", GetPlayerName(Index), "ItemAttack", "No")

    Call SetPlayerSTR(Index, Player(Index).Char(Player(Index).CharNum).MAXSTR)
    Call SendStats(Index)
End Sub

Sub DefenseDouble(ByVal Index As Long)
    Call GetRidOfTimer(Index, 6)

    If IsConnected(Index) = False Or IsPlaying(Index) = False Then
        Exit Sub
    End If

    Call PlayerMsg(Index, "The effects of your item have worn off.", WHITE)
    Call PutVar(App.Path & "\Scripts\" & "StatIncreases.ini", GetPlayerName(Index), "ItemDefense", "No")

    Call SetPlayerDEF(Index, Player(Index).Char(Player(Index).CharNum).MAXDEF)
    Call SendStats(Index)
End Sub

Sub DryTreatHeal(ByVal Index As Long, ByVal Time As Long)
    Call GetRidOfTimer(Index, 2, Time)

    If IsConnected(Index) = False Or IsPlaying(Index) = False Then
        Exit Sub
    End If

    If Time <= 6 And GetPlayerSP(Index) < GetPlayerMaxSP(Index) Then
        Call SetPlayerSP(Index, GetPlayerSP(Index) + (GetPlayerMaxSP(Index) / 10))
        Call SendSP(Index)
        Call AddNewTimer(Index, 2, 5000, (Time + 1))
    End If
End Sub

Sub CookieHeal(ByVal Index As Long, ByVal Time As Integer)
    Call GetRidOfTimer(Index, 3, Time)

    If IsConnected(Index) = False Or IsPlaying(Index) = False Then
        Exit Sub
    End If

    If Time <= 5 And GetPlayerHP(Index) < GetPlayerMaxHP(Index) Then
        Call SendSoundTo(Index, "spm_get_health.wav")
        Call SetPlayerHP(Index, GetPlayerHP(Index) + 2)
        Call SendHP(Index)
        Call AddNewTimer(Index, 3, 5000, (Time + 1))
    End If
End Sub

Sub StatSwap(ByVal Index As Long)
    Call GetRidOfTimer(Index, 4)
    
    If IsConnected(Index) = False Or IsPlaying(Index) = False Then
        Exit Sub
    End If
        
    Call PlayerMsg(Index, "The effects of your special attack have worn off.", WHITE)

    Call SetPlayerSTR(Index, Player(Index).Char(Player(Index).CharNum).MAXSTR)
    Call SetPlayerDEF(Index, Player(Index).Char(Player(Index).CharNum).MAXDEF)
    
    Call PutVar(App.Path & "\Scripts\" & "StatIncreases.ini", GetPlayerName(Index), "Attack", "No")
    Call PutVar(App.Path & "\Scripts\" & "StatIncreases.ini", GetPlayerName(Index), "Defense", "No")
    
    Call SendStats(Index)
End Sub

Sub RedGreenPeppers(ByVal Index As Long, ByVal PepperType As Long)
    ' Red Peppers = 1; Green Peppers = 2
    
    Call GetRidOfTimer(Index, 11)
    
    If IsConnected(Index) = False Or IsPlaying(Index) = False Then
        Exit Sub
    End If

    Call PlayerMsg(Index, "The effects of your item have worn off.", WHITE)
    
    Select Case PepperType
        Case 1 ' Red Peppers
            Call PutVar(App.Path & "\Scripts\" & "StatIncreases.ini", GetPlayerName(Index), "ItemAttack", "No")
            Call PutVar(App.Path & "\Scripts\" & "StatIncreases.ini", GetPlayerName(Index), "RedPeppers", "No")
    
            Call SetPlayerSTR(Index, Player(Index).Char(Player(Index).CharNum).MAXSTR)
        Case 2 ' Green Peppers
            Call PutVar(App.Path & "\Scripts\" & "StatIncreases.ini", GetPlayerName(Index), "ItemDefense", "No")
            Call PutVar(App.Path & "\Scripts\" & "StatIncreases.ini", GetPlayerName(Index), "GreenPeppers", "No")
    
            Call SetPlayerDEF(Index, Player(Index).Char(Player(Index).CharNum).MAXDEF)
    End Select
            
    Call SendStats(Index)
End Sub

Sub Nuts(ByVal Index As Long, ByVal Time As Integer)
    Call GetRidOfTimer(Index, 12, Time)

    If IsConnected(Index) = False Or IsPlaying(Index) = False Then
        Exit Sub
    End If

    If Time <= 3 And GetPlayerHP(Index) < GetPlayerMaxHP(Index) Then
        Call SendSoundTo(Index, "spm_get_health.wav")
        
        Call SetPlayerHP(Index, GetPlayerHP(Index) + 4)
        Call SendHP(Index)
        
        Call AddNewTimer(Index, 12, 3000, (Time + 1))
    End If
End Sub

Sub UltraNuts(ByVal Index As Long, ByVal Time As Integer)
    Call GetRidOfTimer(Index, 13, Time)

    If IsConnected(Index) = False Or IsPlaying(Index) = False Then
        Exit Sub
    End If

    If Time <= 7 And GetPlayerHP(Index) < GetPlayerMaxHP(Index) Then
        Call SendSoundTo(Index, "spm_get_health.wav")
        
        Call SetPlayerHP(Index, GetPlayerHP(Index) + 2)
        Call SetPlayerMP(Index, GetPlayerMP(Index) + 1)
        
        Call SendHP(Index)
        Call SendMP(Index)
        
        Call AddNewTimer(Index, 13, 1000, (Time + 1))
    End If
End Sub

Sub SendHideNSneakMsg(ByVal Msg As String, ByVal MsgColor As Byte, Optional ByVal IsPlayerMsg As Boolean = False, Optional ByVal Msg2 As String = vbNullString)
    Dim PlayerNum As String
    Dim R As Long, Q As Long
    Dim i As Byte
    
    For i = 1 To MaxHiders
        PlayerNum = CStr(i)
        
        ' Only check seekers from values 1 - 3 since there are only 3 max seekers
        If i < 4 Then
            R = FindPlayer(GetVar(HideNSneakPath, "Seekers", PlayerNum))
                
            ' Send the time remaining to the Seekers
            If IsPlaying(R) Then
                If IsPlayerMsg = False Then
                    Call BattleMsg(R, Msg, MsgColor, 1)
                Else
                    Call PlayerMsg(R, Msg, MsgColor)
                End If
                
                If Msg2 <> vbNullString Then
                    If IsPlayerMsg = False Then
                        Call BattleMsg(R, Msg2, YELLOW, 0)
                    Else
                        Call PlayerMsg(R, Msg2, YELLOW)
                    End If
                End If
            End If
        End If
        
        Q = FindPlayer(GetVar(HideNSneakPath, "Hiders", PlayerNum))
        
        ' Send the time remaining to the Hiders
        If IsPlaying(Q) Then
            If IsPlayerMsg = False Then
                Call BattleMsg(Q, Msg, MsgColor, 1)
            Else
                Call PlayerMsg(Q, Msg, MsgColor)
            End If
            
            If Msg2 <> vbNullString Then
                If IsPlayerMsg = False Then
                    Call BattleMsg(Q, Msg2, YELLOW, 0)
                Else
                    Call PlayerMsg(Q, Msg2, YELLOW)
                End If
            End If
        End If
    Next i
End Sub

Sub HideNSneakPlayTime(ByVal Index As Long)
    Dim PlayerNum As String, PlayerGetVar As String
    Dim R As Long, Q As Long
    Dim TimeLeft As Integer
    Dim i As Byte, PlayersFound As Byte
    
    Call GetRidOfTimer(Index, 10)
    TimeLeft = CInt(GetVar(HideNSneakPath, "GameTime", "TimeLeft")) - 20

    ' Continues game
    If TimeLeft > 0 Then
        Call PutVar(HideNSneakPath, "GameTime", "TimeLeft", CStr(TimeLeft))
        PlayersFound = CByte(GetVar(HideNSneakPath, "PlayersFound", "PlayersFound"))
        
        ' Notify all players of the time remaining and current score
        Call SendHideNSneakMsg("Time remaining: " & TimeLeft & " seconds", WHITE, False, "Hiders found: " & PlayersFound & " / " & CByte(GetVar(HideNSneakPath, "Team", "Hiders")))
        
        ' Start timer again
        Call AddNewTimer(Index, 10, 20000)
    ' Exits game when time is up, and gives out rewards
    ElseIf TimeLeft <= 0 Then
        ' It's now time for the seekers to start looking for players
        If GetVar(HideNSneakPath, "GameTime", "IsSeeking") = "No" Then
            Call PutVar(HideNSneakPath, "GameTime", "TimeLeft", "120")
            Call PutVar(HideNSneakPath, "GameTime", "IsSeeking", "Yes")
            
            Call SendHideNSneakMsg("The hiders' time is up, and the seekers have began searching! Come out, come out, wherever you are!", YELLOW, True)
        
            Dim RandomSprite As Byte
        
            ' Warp the seekers to the map
            For i = 1 To MaxHiders
                PlayerNum = CStr(i)
                
                If i <= MaxSeekers Then
                    R = FindPlayer(GetVar(HideNSneakPath, "Seekers", PlayerNum))
                    
                    If IsPlaying(R) Then
                        ' Store the seekers' old sprites
                        Call SetPlayerTempSprite(R, GetPlayerSprite(R))
                    
                        ' Set the sprite of the seeker to one of the boos
                        RandomSprite = Rand(74, 75)
                        Call SetPlayerSprite(R, RandomSprite)
                        
                        ' Heal the player fully
                        Call SetPlayerHP(R, GetPlayerMaxHP(R))
                        Call SetPlayerMP(R, GetPlayerMaxMP(R))
                        Call SetPlayerSP(R, GetPlayerMaxSP(R))
                        
                        Call SendHP(R)
                        Call SendMP(R)
                        Call SendSP(R)
                        
                        ' State that the seeker is playing
                        Call SendPlayingHideNSneak(R, True)
                        
                        ' Warp the seeker to the minigame
                        Call PlayerWarp(R, 271, 15, 23)
                    End If
                End If
                
                ' Stop the hiders from moving
                Q = FindPlayer(GetVar(HideNSneakPath, "Hiders", PlayerNum))
                
                If IsPlaying(Q) Then
                    Call SendIsHiderFrozen(Q, True)
                End If
            Next
                
            ' Start the timer again
            Call AddNewTimer(Index, 10, 20000)
        Else
            Dim TotalPlayers As Integer, NumPlayersFound As Integer, NumHiders As Integer
            
            ' Find out the number of hiders in the game
            NumHiders = CInt(GetVar(HideNSneakPath, "Team", "Hiders"))
            
            ' Find out the number of total players in the game by adding up the number of hiders and seekers
            TotalPlayers = NumHiders + CInt(GetVar(HideNSneakPath, "Team", "Seekers"))
            
            ' Find out how many players were found
            NumPlayersFound = CInt(GetVar(HideNSneakPath, "PlayersFound", "PlayersFound"))
        
            Dim NumHiderCoins As Integer, NumSeekerCoins As Integer
            
            NumSeekerCoins = 2 * TotalPlayers
            NumHiderCoins = NumSeekerCoins
            
            Dim HiderMsg As String, SeekerMsg As String
            
            ' Check if the Seekers won and multiply the number of reward coins they receive by 2 if they won
            If NumPlayersFound = NumHiders Then
                NumSeekerCoins = NumSeekerCoins * 2
                
                ' Set the Hider and Seeker game finish messages - Seekers won, Hiders lost
                SeekerMsg = "Congratulations! Your team won! You receive "
                HiderMsg = "Sorry, your team lost. Try again next time! You receive "
            Else ' Check if the Hiders won and multiply the number of reward coins they receive by 2 if they won
                NumHiderCoins = NumHiderCoins * 2
                
                ' Set the Hider and Seeker game finish messages - Hiders won, Seekers lost
                SeekerMsg = "Sorry, your team lost. Try again next time! You receive "
                HiderMsg = "Congratulations! Your team won! You receive "
            End If
            
            ' If the game ended because of a disconnect, then divide the number of coins received by 2
            If HasLeftHideNSneak = True Then
                NumSeekerCoins = NumSeekerCoins / 2
                NumHiderCoins = NumHiderCoins / 2
            End If
            
            ' Finish up the Seeker message
            If NumSeekerCoins <> 1 Then
                SeekerMsg = SeekerMsg & NumSeekerCoins & " Reward Coins!"
            Else
                SeekerMsg = SeekerMsg & NumSeekerCoins & " Reward Coin!"
            End If

            ' Finish up the Hider message
            If NumHiderCoins <> 1 Then
                HiderMsg = HiderMsg & NumHiderCoins & " Reward Coins!"
            Else
                HiderMsg = HiderMsg & NumHiderCoins & " Reward Coin!"
            End If
            
            ' Reset the game and give out rewards
            For i = 1 To MaxHiders
                PlayerNum = CStr(i)
                
                PlayerGetVar = GetVar(HideNSneakPath, "Hiders", PlayerNum)
                Q = FindPlayer(PlayerGetVar)
                
                If IsPlaying(Q) Then
                    Call SetPlayerSprite(Q, GetPlayerTempSprite(Q))
                    Call PlayerWarp(Q, 269, (12 + i), 8)
                    Call SetPlayerPK(Q, NO)
                    Call SendPlayerData(Q)
                    Call SendPlayerXY(Q)
                    
                    ' State that the hider is no longer playing
                    Call SendPlayingHideNSneak(Q, False)
                    
                    ' Allow the hider to move
                    Call SendIsHiderFrozen(Q, False)
                    
                    ' Give out rewards
                    Call GiveItem(Q, 280, NumHiderCoins)
                    
                    ' Send the player the game finish message
                    Call PlayerMsg(Q, HiderMsg, BLUE)
                End If
                
                Call PutVar(HideNSneakPath, "PlayersOut", PlayerNum, "")
                
                If PlayerGetVar <> vbNullString Then
                    Call PutVar(HideNSneakPath, "Hiders", PlayerNum, "")
                End If
                
                If i <= MaxSeekers Then
                    PlayerGetVar = GetVar(HideNSneakPath, "Seekers", PlayerNum)
                    
                    R = FindPlayer(PlayerGetVar)
                    
                    If IsPlaying(R) Then
                        Call SetPlayerSprite(R, GetPlayerTempSprite(R))
                        Call PlayerWarp(R, 269, (13 + i), 7)
                        Call SetPlayerPK(R, NO)
                        Call SendPlayerData(R)
                        Call SendPlayerXY(R)
                        
                        ' State that the seeker is no longer playing
                        Call SendPlayingHideNSneak(R, False)
                        
                        ' Gives out rewards
                        Call GiveItem(R, 280, NumSeekerCoins)
                        
                        ' Send the player the game finish message
                        Call PlayerMsg(R, SeekerMsg, BLUE)
                    End If
                   
                    If PlayerGetVar <> vbNullString Then
                        Call PutVar(HideNSneakPath, "Seekers", PlayerNum, "")
                    End If
                End If
            Next i
            
            ' State that no one has left Hide n' Sneak since this game is over
            HasLeftHideNSneak = False
            
            ' Resets game
            Call PutVar(HideNSneakPath, "PlayersFound", "PlayersFound", "0")
            Call PutVar(HideNSneakPath, "GameTime", "TimeLeft", "0")
            Call PutVar(HideNSneakPath, "GameTime", "IsSeeking", "No")
            Call PutVar(HideNSneakPath, "Ingame", "Ingame", "No")
            Call PutVar(HideNSneakPath, "Team", "Hiders", "0")
            Call PutVar(HideNSneakPath, "Team", "Seekers", "0")
            Call PutVar(HideNSneakPath, "TimerIndex", "TimerIndex", "")
        End If
    End If
End Sub

Sub DodgeBallPlayTime(ByVal Index As Long)
    Dim FilePath As String, PlayerNum As String
    Dim i As Long, R As Long, Q As Long
    Dim Point As Integer, POINTS As Integer, TimeLeft As Integer
    
    FilePath = DodgeBillPath
    
    Call GetRidOfTimer(Index, 9)
    TimeLeft = CInt(GetVar(FilePath, "GameTime", "TimeLeft")) - 20

    ' Continues game
    If TimeLeft > 0 Then
        Call PutVar(FilePath, "GameTime", "TimeLeft", CStr(TimeLeft))
        ' Sends time left to players every 20 seconds
        Call MapBattleMsg(188, "Time remaining: " & TimeLeft & " seconds", WHITE, 1)
                
        Point = CInt(GetVar(FilePath, "Points", "Red"))
        POINTS = CInt(GetVar(FilePath, "Points", "Blue"))
        ' Sends points to players every 20 seconds
        Call MapBattleMsg(188, "Red Team: " & Point, YELLOW, 0)
        Call MapBattleMsg(188, "Blue Team: " & POINTS, YELLOW, 0)
        ' Start timer again
        Call AddNewTimer(Index, 9, 20000)
    ' Exits game when time is up, and gives out rewards
    ElseIf TimeLeft <= 0 Then
        Point = CInt(GetVar(FilePath, "Points", "Red"))
        POINTS = CInt(GetVar(FilePath, "Points", "Blue"))
    
        ' Reset the game and give out rewards
        For i = 1 To 5
            PlayerNum = CStr(i)
            R = FindPlayer(GetVar(FilePath, "Blue", PlayerNum))
            Q = FindPlayer(GetVar(FilePath, "Red", PlayerNum))
            
            If IsPlaying(Q) Then
                Call PlayerWarp(Q, 190, 19 + i, 10)
                Call SetPlayerPK(Q, NO)
                Call SendPlayerData(Q)
                Call SendPlayerXY(Q)
                
                ' Remove bullet bills from inventory
                Call TakeItem(Q, 186, 4)
                
                ' Gives out rewards
                If Point > POINTS Then
                    Call PlayerMsg(Q, "Congratulations! Your team won! You receive 3 Reward Coins!", BLUE)
                    Call GiveItem(Q, 189, 3)
                ElseIf POINTS > Point Then
                    Call PlayerMsg(Q, "Sorry, your team lost. Try again next time! You receive 1 Reward Coin!", BLUE)
                    Call GiveItem(Q, 189, 1)
                ElseIf Point = POINTS Then
                    Call PlayerMsg(Q, "Wow, it was a tie! You receive 2 Reward Coins!", BLUE)
                    Call GiveItem(Q, 189, 2)
                End If
                
                Call PlayerMsg(Q, "The final score was: Red: " & Point & "; Blue: " & POINTS, YELLOW)
            End If
            
            If IsPlaying(R) Then
                Call PlayerWarp(R, 190, 19 + i, 9)
                Call SetPlayerPK(R, NO)
                Call SendPlayerData(R)
                Call SendPlayerXY(R)
                
                ' Remove bullet bills from inventory
                Call TakeItem(R, 186, 4)
                
                ' Gives out rewards
                If Point > POINTS Then
                    Call PlayerMsg(R, "Sorry, your team lost. Try again next time! You receive 1 Reward Coin!", BLUE)
                    Call GiveItem(R, 189, 1)
                ElseIf POINTS > Point Then
                    Call PlayerMsg(R, "Congratulations! Your team won! You receive 3 Reward Coins!", BLUE)
                    Call GiveItem(R, 189, 3)
                ElseIf Point = POINTS Then
                    Call PlayerMsg(R, "Wow, it was a tie! You receive 2 Reward Coin!", BLUE)
                    Call GiveItem(R, 189, 2)
                End If
                
                Call PlayerMsg(R, "The final score was: Red: " & Point & "; Blue: " & POINTS, YELLOW)
            End If
            
            If GetVar(FilePath, "Red", PlayerNum) <> vbNullString Then
                Call PutVar(FilePath, "Red", PlayerNum, "")
            End If
            
            If GetVar(FilePath, "Blue", PlayerNum) <> vbNullString Then
                Call PutVar(FilePath, "Blue", PlayerNum, "")
            End If
        Next i
        
        ' Respawn the map just incase
        Call RespawnMap(188)
        
        ' Resets game
        Call PutVar(FilePath, "Ingame", "Ingame", "No")
        Call PutVar(FilePath, "Points", "Red", "0")
        Call PutVar(FilePath, "Points", "Blue", "0")
        Call PutVar(FilePath, "Team", "Red", "0")
        Call PutVar(FilePath, "Team", "Blue", "0")
        Call PutVar(FilePath, "GameTime", "TimeLeft", "0")
        Call PutVar(FilePath, "Outs", "Blue", "0")
        Call PutVar(FilePath, "Outs", "Red", "0")
    End If
End Sub

Sub EquipItem(ByVal Index As Long, ByVal InvNum As Long, EquipSlot As Long)
    Dim i As Long, ItemNum As Long, EquippedItemNum As Long, EquippedItemValue As Long, EquippedItemAmmo As Long
    
    ' Weapon = 1; Shirt = 2; Cap = 3; Special Badge = 4; Pants = 5; Flower Badge = 6; Mushroom Badge = 7
    
    ItemNum = GetPlayerInvItemNum(Index, InvNum)
    
    ' Prevents players from equipping items that would make their HP/FP less than 5
    If GetPlayerMaxHP(Index) + Item(ItemNum).addHP < 5 Then
        Call PlayerMsg(Index, "You cannot equip this item because you will have less than 5 HP!", WHITE)
        Exit Sub
    End If
    If GetPlayerMaxMP(Index) + Item(ItemNum).addMP < 5 Then
        Call PlayerMsg(Index, "You cannot equip this item because you will have less than 5 FP!", WHITE)
        Exit Sub
    End If
    
    If ItemIsUsable(Index, InvNum) = False Then
        Exit Sub
    End If
    
    EquippedItemNum = GetPlayerEquipSlotNum(Index, EquipSlot)
    
    If EquippedItemNum > 0 Then
        EquippedItemValue = GetPlayerEquipSlotValue(Index, EquipSlot)
        EquippedItemAmmo = GetPlayerEquipSlotAmmo(Index, EquipSlot)
    End If
    
    Call SetPlayerEquipSlotNum(Index, EquipSlot, GetPlayerInvItemNum(Index, InvNum))
    Call SetPlayerEquipSlotValue(Index, EquipSlot, GetPlayerInvItemValue(Index, InvNum))
    Call SetPlayerEquipSlotAmmo(Index, EquipSlot, GetPlayerInvItemAmmo(Index, InvNum))
    Call TakeSpecificItem(Index, InvNum, 1)
            
    If EquippedItemNum > 0 Then
        i = FindOpenInvSlot(Index, EquippedItemNum)
        Call GiveItem(Index, EquippedItemNum, 1)
        Call SetPlayerInvItemValue(Index, i, EquippedItemValue)
        Call SetPlayerInvItemAmmo(Index, i, EquippedItemAmmo)
        Call SendInventoryUpdate(Index, i)
    End If
            
    Call SendWornEquipment(Index)
    Call SendEquipmentUpdate(Index, EquipSlot)
    ' Update information
    Call SendStats(Index)
    Call SendHP(Index)
    Call SendMP(Index)
    Call SendSP(Index)
End Sub

Sub UnequipItem(ByVal Index As Long, ByVal EquipSlot As Long)
    Dim InvSlot As Long
    
    InvSlot = FindOpenInvSlot(Index, GetPlayerEquipSlotNum(Index, EquipSlot))
    
    If InvSlot > 0 Then
        Call GiveItem(Index, GetPlayerEquipSlotNum(Index, EquipSlot), 1)
        
        ' Retain item value and ammo
        Call SetPlayerInvItemValue(Index, InvSlot, GetPlayerEquipSlotValue(Index, EquipSlot))
        Call SetPlayerInvItemAmmo(Index, InvSlot, GetPlayerEquipSlotAmmo(Index, EquipSlot))
        
        ' Clear out equipment
        Call SetPlayerEquipSlotNum(Index, EquipSlot, 0)
        Call SetPlayerEquipSlotValue(Index, EquipSlot, 0)
        Call SetPlayerEquipSlotAmmo(Index, EquipSlot, -1)
        
        Call SendWornEquipment(Index)
        Call SendInventoryUpdate(Index, InvSlot)
        ' Update information
        Call SendStats(Index)
        Call SendHP(Index)
        Call SendMP(Index)
        Call SendSP(Index)
    Else
        Call PlayerMsg(Index, "Your inventory is full! Please make some room for this item.", BRIGHTRED)
    End If
End Sub

Sub SendEquipmentUpdate(ByVal Index As Long, ByVal EquipSlot As Long)
    Call SendDataToMap(GetPlayerMap(Index), SPackets.Splayerequipupdate & SEP_CHAR & EquipSlot & SEP_CHAR & Index & SEP_CHAR & GetPlayerEquipSlotNum(Index, EquipSlot) & SEP_CHAR & GetPlayerEquipSlotValue(Index, EquipSlot) & SEP_CHAR & GetPlayerEquipSlotAmmo(Index, EquipSlot) & END_CHAR)
End Sub

Sub Drill(ByVal Index As Long)
    Dim MapNum As Long, X As Long, Y As Long
    
    MapNum = GetPlayerMap(Index)
    X = GetPlayerX(Index)
    Y = GetPlayerY(Index)
        
    With Map(MapNum).Tile(X, Y)
        
        If .Type <> TILE_TYPE_DRILL Then
            Call SpellAnim(28, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index))
            Call SendSoundToMap(GetPlayerMap(Index), "m&lpit_babydrill.wav")

            ' Allows players to get Mushroom Ball items for the Favor
            If .Type = TILE_TYPE_SCRIPTED Then
                If .Data1 = 24 Then
                    If GetVar(App.Path & "\Scripts\Quests.ini", GetPlayerName(Index), "ItemQuest13") = "InProgress" Then
                        Dim VarName As String
                        
                        VarName = "mbquest /" & GetPlayerMap(Index) & "/" & GetPlayerX(Index) & "/" & GetPlayerY(Index)
                        
                        If GetVar(App.Path & "\Scripts\" & "MBFavor.ini", GetPlayerName(Index), VarName) <> "1" Then
                            Call PutVar(App.Path & "\Scripts\" & "MBFavor.ini", GetPlayerName(Index), VarName, "1")
                            Call GiveItem(Index, 190, 1)
                            
                            Call PlayerMsg(Index, "You found a Mushroom Ball!", WHITE)
                        End If
                    End If
                End If
                
                Call CheckForBeanFruit(Index)
            End If
            
            ' Allows players to get beans
            If .Type = TILE_TYPE_BEAN Then
                If .Data3 = 0 Then
                    If ItemIsStackable(.Data1) = True Then
                        If FindOpenInvSlot(Index, .Data1) > 0 Then
                            GoTo DigUpBean
                        Else
                            Call PlayerMsg(Index, "Your inventory is full! Please make some room for this item.", BRIGHTRED)
                        End If
                    Else
                        If GetFreeSlots(Index) >= .Data2 Then
                            GoTo DigUpBean
                        Else
                            Call PlayerMsg(Index, "Your inventory is full! Please make some room for this item.", BRIGHTRED)
                        End If
                    End If
                End If
            End If
            
            Exit Sub
        End If
    End With
    
    MapNum = Map(MapNum).Tile(X, Y).Data1
    X = Map(GetPlayerMap(Index)).Tile(X, Y).Data2
    Y = Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), Y).Data3
    
    Call SpellAnim(28, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index))
    Call SendSoundToMap(GetPlayerMap(Index), "m&lpit_babydrill.wav")
    Call PlayerMsg(Index, "You use the Drill ability to dig beneath the ground and reach your destination.", WHITE)
    Call PlayerWarp(Index, MapNum, X, Y)
    
    Exit Sub
    
DigUpBean:
    With Map(MapNum).Tile(X, Y)
        Call GiveItem(Index, .Data1, .Data2)
                                
        If .Data2 > 1 Then
            Call PlayerMsg(Index, "You dig up " & .Data2 & " " & Trim$(Item(.Data1).Name) & "s.", WHITE)
        Else
            If FindItemVowels(.Data1) = True Then
                Call PlayerMsg(Index, "You dig up an " & Trim$(Item(.Data1).Name) & ".", WHITE)
            Else
                Call PlayerMsg(Index, "You dig up a " & Trim$(Item(.Data1).Name) & ".", WHITE)
            End If
        End If
        
        .Data3 = 1
    End With
End Sub

Sub CheckForBeanFruit(ByVal Index As Long)
    ' Allows players to get Bean Fruits for the Favor
    If GetVar(App.Path & "\Scripts\Quests.ini", GetPlayerName(Index), "ItemQuest16") = "InProgress" Then
        If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1 = 31 Then
            Dim PlayerCoords As String
            
            PlayerCoords = GetPlayerMap(Index) & "/" & GetPlayerX(Index) & "/" & GetPlayerY(Index)
            
            If GetVar(App.Path & "\Scripts\" & "BFFavor.ini", GetPlayerName(Index), PlayerCoords) <> "Got" Then
                Dim ItemNum As Long
            
                Select Case PlayerCoords
                    Case "246/11/18"
                        ItemNum = 285
                    Case "249/10/24"
                        ItemNum = 288
                    Case "254/27/23"
                        ItemNum = 291
                    Case "274/5/15"
                        ItemNum = 286
                    Case "281/4/12"
                        ItemNum = 289
                    Case "284/27/26"
                        ItemNum = 294
                    Case "290/25/2"
                        ItemNum = 287
                    Case "303/4/6"
                        ItemNum = 290
                    Case "307/27/8"
                        ItemNum = 292
                    Case "314/3/9"
                        ItemNum = 293
                End Select
                
                If GetFreeSlots(Index) > 0 Then
                    Call PutVar(App.Path & "\Scripts\" & "BFFavor.ini", GetPlayerName(Index), PlayerCoords, "Got")
                    Call GiveItem(Index, ItemNum, 1)
                            
                    If FindItemVowels(ItemNum) = True Then
                        Call PlayerMsg(Index, "You found an " & Trim$(Item(ItemNum).Name) & "!", WHITE)
                    Else
                        Call PlayerMsg(Index, "You found a " & Trim$(Item(ItemNum).Name) & "!", WHITE)
                    End If
                Else
                    Call PlayerMsg(Index, "Your inventory is full! Please make some room for this item.", BRIGHTRED)
                End If
            End If
        End If
    End If
End Sub

Sub SendSpawnItemSlot(ByVal MapNum As Long, MapItemIndex As Long, ByVal ItemNum As Long, ByVal Value As Long, ByVal X As Byte, ByVal Y As Byte, ByVal Ammo As Long)
    Call SendDataToMap(MapNum, SPackets.Sspawnitem & SEP_CHAR & MapItemIndex & SEP_CHAR & ItemNum & SEP_CHAR & Value & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & Ammo & END_CHAR)
End Sub

Sub SendNpcDead(ByVal MapNum As Long, ByVal MapNpcNum As Long)
    Call SendDataToMap(MapNum, SPackets.Snpcdead & SEP_CHAR & MapNpcNum & END_CHAR)
End Sub

Sub SendAttackSound(ByVal MapNum As Long)
    Call SendDataToMap(MapNum, SPackets.Ssound & SEP_CHAR & "attack" & END_CHAR)
End Sub

Sub SendMapKey(ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long, ByVal Status As Long)
    Call SendDataToMap(MapNum, SPackets.Smapkey & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & Status & END_CHAR)
End Sub

Sub SendTurnBasedBattle(ByVal Index As Long, ByVal Status As Byte, ByVal MapNpcNum As Long)
    Call SendDataToMap(GetPlayerMap(Index), SPackets.Sturnbasedbattle & SEP_CHAR & Index & SEP_CHAR & Status & SEP_CHAR & MapNpcNum & END_CHAR)
End Sub

Sub SendMsgBoxTo(ByVal Index As Long, ByVal Title As String, ByVal Message As String)
    Call SendDataTo(Index, SPackets.Smsgbox & SEP_CHAR & Title & SEP_CHAR & Message & END_CHAR)
End Sub

Sub SendCheckForMap(ByVal Index As Long)
    Call SendDataTo(Index, SPackets.Scheckformap & SEP_CHAR & GetPlayerMap(Index) & SEP_CHAR & Map(GetPlayerMap(Index)).Revision & END_CHAR)
End Sub

Sub SendSoundTo(ByVal Index As Long, ByVal Sound As String)
    Call SendDataTo(Index, SPackets.Ssound & SEP_CHAR & "soundattribute" & SEP_CHAR & Sound & END_CHAR)
End Sub

Sub SendSoundToMap(ByVal MapNum As Long, ByVal Sound As String)
    Call SendDataToMap(MapNum, SPackets.Ssound & SEP_CHAR & "soundattribute" & SEP_CHAR & Sound & END_CHAR)
End Sub

Sub SendCanUseSpecial(ByVal Index As Long)
    Call SendDataTo(Index, SPackets.Scanusespecial & END_CHAR)
End Sub

Sub ScriptSpawnNpc(ByVal MapNpcNum As Long, ByVal MapNum As Long, ByVal SpawnX As Byte, ByVal SpawnY As Byte, ByVal NpcNum As Long)
    ' Check for subscript out of range
    If MapNpcNum < 0 Or MapNpcNum > MAX_MAP_NPCS Or MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If

    If NpcNum = 0 Then
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

        Call SaveMap(MapNum)
    End If

    Map(MapNum).Revision = Map(MapNum).Revision + 1

    MapNPC(MapNum, MapNpcNum).num = NpcNum
    Map(MapNum).NPC(MapNpcNum) = NpcNum

    MapNPC(MapNum, MapNpcNum).Target = 0

    MapNPC(MapNum, MapNpcNum).HP = GetNpcMaxHP(NpcNum)
    MapNPC(MapNum, MapNpcNum).MP = GetNpcMaxMP(NpcNum)
    MapNPC(MapNum, MapNpcNum).SP = GetNpcMaxSP(NpcNum)
   
    MapNPC(MapNum, MapNpcNum).Dir = Int(Rnd2 * 4)

    MapNPC(MapNum, MapNpcNum).X = SpawnX
    MapNPC(MapNum, MapNpcNum).Y = SpawnY
    
    Call SendDataToMap(MapNum, SPackets.Sspawnnpc & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNPC(MapNum, MapNpcNum).num & SEP_CHAR & MapNPC(MapNum, MapNpcNum).X & SEP_CHAR & MapNPC(MapNum, MapNpcNum).Y & SEP_CHAR & MapNPC(MapNum, MapNpcNum).Dir & SEP_CHAR & NPC(MapNPC(MapNum, MapNpcNum).num).Big & END_CHAR)
End Sub

Sub LeaveParty(ByVal Index As Long)
    Dim i As Long, PlayerIndex As Long, PartyNum As Long, PartyMember As Long
    
    PartyNum = GetPlayerPartyNum(Index)
    
    If PartyNum = 0 Then
        Call PlayerMsg(Index, "You're not in a party!", BRIGHTRED)
        Exit Sub
    End If
    
    If PartyNum <= 0 Or PartyNum > MAX_PLAYERS Then Exit Sub
    
    ' If the player is the leader, then replace him/her as leader and notify everyone in the party
    If Index = Party(PartyNum).Leader Then
        Call SetPlayerPartyShare(Index, False)
        Call RemovePartyMember(PartyNum, Index)
        Call SetPartyLeader(PartyNum, 0)
        Call PlayerMsg(Index, "You have left the party.", WHITE)
        
        ' Set new leader if the player was a leader
        For i = 1 To MAX_PARTY_MEMBERS
            PartyMember = Party(PartyNum).Member(i)
            If PartyMember > 0 Then
                Call SetPartyLeader(PartyNum, PartyMember)
                Call SetPlayerPartyShare(PartyMember, True)
                Exit For
            End If
        Next i
        
        ' If there still is no leader, then that means there are no other party members
        If GetPartyLeader(PartyNum) = 0 Then
            Call PlayerMsg(Index, "The party has been disbanded.", WHITE)
            Exit Sub
        End If
        
        ' Notify everyone in the party that the leader left and was replaced
        For i = 1 To MAX_PARTY_MEMBERS
            PartyMember = Party(PartyNum).Member(i)
            If PartyMember > 0 Then
                Call PlayerMsg(PartyMember, GetPlayerName(Index) & " has left the party. The new leader is: " & GetPlayerName(Party(PartyNum).Leader), WHITE)
            End If
        Next i
    ' Simply remove the player from the party if he/she is not the leader
    Else
        Call SetPlayerPartyShare(Index, False)
        Call RemovePartyMember(PartyNum, Index)
        Call PlayerMsg(Index, "You have left the party.", WHITE)
        
        ' Notify everyone in the party that the player left
        For i = 1 To MAX_PARTY_MEMBERS
            PartyMember = Party(PartyNum).Member(i)
            If PartyMember > 0 Then
                Call PlayerMsg(PartyMember, GetPlayerName(Index) & " has left the party.", WHITE)
            End If
        Next i
    End If
End Sub

Function CanUseSpecial(ByVal Index As Long, ByVal SpellSlot As Long) As Boolean
    If IsInVictoryAnim(Index) = True Then
        Exit Function
    End If

    ' Prevent subscript out of range
    If SpellSlot <= 0 Or SpellSlot > MAX_PLAYER_SPELLS Then
        Exit Function
    End If

    Dim SpellNum As Long

    SpellNum = GetPlayerSpell(Index, SpellSlot)

    If SpellNum <= 0 Then
        Exit Function
    End If

    ' Make sure player has the spell
    If Not HasSpell(Index, SpellNum) Then
        Call BattleMsg(Index, "You don't have this special attack!", BRIGHTRED, 0)
        Exit Function
    End If
    
    ' Check if the player has enough FP
    If SpellNum <> 45 Then
        If GetPlayerMP(Index) < FlowerSaver(Index, SpellNum) Then
            Call BattleMsg(Index, "You don't have enough FP to use this special attack!", BRIGHTRED, 0)
            Exit Function
        End If
    End If

    Dim RequiredLevel As Long

    RequiredLevel = GetSpellReqLevel(SpellNum)

    ' Make sure the player is at a high enough level
    If RequiredLevel > GetPlayerLevel(Index) Then
        Call PlayerMsg(Index, "You need to be at level " & RequiredLevel & " to perform this special attack.", WHITE)
        Exit Function
    End If

    ' Check if timer is ok
    If GetTickCount < Player(Index).AttackTimer + 1000 Then
        Exit Function
    End If
    
    CanUseSpecial = True
End Function

Sub HotScript(ByVal Index As Long, ByVal KeyID As Byte)
    Select Case KeyID
        ' Executes when players press the CTRL key
        Case 1
            ' Golden Cannon Launch
            If GetPlayerMap(Index) = 103 And GetPlayerX(Index) = 10 And GetPlayerY(Index) = 8 And GetPlayerDir(Index) = 2 And CanTake(Index, 112, 1) Then
                Call TakeItem(Index, 112, 1)
                Call PlayerWarp(Index, 107, 22, 20)
                Call SendSoundTo(Index, "sm64_cannon_fire.wav")
                Call PlayerMsg(Index, "You put the Golden Bullet Bill in the cannon and fire it at yourself!", WHITE)
            
                Exit Sub
            End If
            
            ' Salestoad Plant
            If GetPlayerMap(Index) = 229 Then
                Dim X As Long, Y As Long
                
                X = GetPlayerX(Index)
                Y = GetPlayerY(Index)
                
                Select Case GetPlayerDir(Index)
                    Case DIR_LEFT
                        X = X - 1
                    Case DIR_RIGHT
                        X = X + 1
                    Case DIR_UP
                        Y = Y - 1
                    Case DIR_DOWN
                        Y = Y + 1
                End Select
                
                ' Make sure the tile is scripted and is script #27
                If Map(GetPlayerMap(Index)).Tile(X, Y).Type = TILE_TYPE_SCRIPTED And Map(GetPlayerMap(Index)).Tile(X, Y).Data1 = 27 Then
                    If GetVar(App.Path & "\Scripts\Quests.ini", GetPlayerName(Index), "ItemQuest14-Got") <> "Got" Then
                        If GetFreeSlots(Index) < 1 Then
                            Call PlayerMsg(Index, "You've found something, but you don't have any room to hold it!", BRIGHTRED)
                            Exit Sub
                        Else
                            Call GiveItem(Index, 266, 1)
                            Call PutVar(App.Path & "\Scripts\Quests.ini", GetPlayerName(Index), "ItemQuest14-Got", "Got")
                            Call PlayerMsg(Index, "You got the " & Trim$(Item(266).Name) & "!", YELLOW)
                            Exit Sub
                        End If
                    End If
                End If
            End If
        ' Executes when players hit the Delete key
        Case 2
            ' Check if the player hit a Question Block
            If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type = TILE_TYPE_QUESTIONBLOCK Then
                Call HitQuestionBlock(Index)
                
                Exit Sub
            End If
        
            ' Respawn Point
            If GetVar(App.Path & "\Respawn Points.ini", GetPlayerName(Index), "CTRL") = "1" Then
                Call PutVar(App.Path & "\Respawn Points.ini", GetPlayerName(Index), "Map #", GetPlayerMap(Index))
                Call PutVar(App.Path & "\Respawn Points.ini", GetPlayerName(Index), "X", GetPlayerX(Index))
                Call PutVar(App.Path & "\Respawn Points.ini", GetPlayerName(Index), "Y", GetPlayerY(Index))
                Call SendSoundToMap(GetPlayerMap(Index), "m&lpit_savealbum.wav")
                Call SpellAnim(1, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index))
                Call PlayerMsg(Index, "Your respawn point has been saved!", YELLOW)

                Call PutVar(App.Path & "\Respawn Points.ini", GetPlayerName(Index), "CTRL", "0")
                
                Exit Sub
            End If
    
            ' Heart Blocks
            If GetVar(App.Path & "\Heart Blocks.ini", GetPlayerName(Index), "CTRL") = "1" And CanTake(Index, 1, GetPlayerLevel(Index) * 3) Then
                Call TakeItem(Index, 1, GetPlayerLevel(Index) * 3)
                Call SetPlayerHP(Index, GetPlayerMaxHP(Index))
                Call SetPlayerMP(Index, GetPlayerMaxMP(Index))
                Call SetPlayerSP(Index, GetPlayerMaxSP(Index))
                
                Call SendHP(Index)
                Call SendMP(Index)
                Call SendSP(Index)
          
                Call PlayerMsg(Index, "You just got fully healed! You feel amazing!", BRIGHTGREEN)
                Call SendSoundTo(Index, "spm_get_health.wav")
                Call SpellAnim(4, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index))
                Call PutVar(App.Path & "\Heart Blocks.ini", GetPlayerName(Index), "CTRL", "0")
                
                Exit Sub
            ElseIf GetVar(App.Path & "\Heart Blocks.ini", GetPlayerName(Index), "CTRL") = "1" And Not CanTake(Index, 1, GetPlayerLevel(Index) * 3) Then
                Call PlayerMsg(Index, "You don't have enough coins to use this block!", WHITE)
                Call PutVar(App.Path & "\Heart Blocks.ini", GetPlayerName(Index), "CTRL", "0")
                
                Exit Sub
            End If
            
            ' Check if the player should learn the Rock Smasher special attack
            If GetPlayerMap(Index) = 309 Then
                If GetPlayerX(Index) = 12 Then
                    If GetPlayerY(Index) = 25 Then
                       ' Teach the player the Rock Smasher special attack if he/she didn't already learn it
                       If Not HasSpell(Index, 45) Then
                           Call SetPlayerSpell(Index, FindOpenSpellSlot(Index), 45)
                           Call PlayerMsg(Index, "You've learned the Rock Smasher special attack!", WHITE)
                       End If
                    End If
                End If
            End If
            
            ' Check if the player hit a Simultaneous Block
            With Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index))
                If .Type = TILE_TYPE_SIMULBLOCK Then
                    Dim SimulBlockCoords() As String, SimulBlockXYCoords() As String
                    Dim SimulBlockXCoords(0 To 3) As Long, SimulBlockYCoords(0 To 3) As Long
                    Dim i As Byte
                    
                    SimulBlockCoords = Split(.String1, "/")
                    
                    For i = 0 To 3
                        SimulBlockXYCoords = Split(SimulBlockCoords(i), ",")
                        
                        SimulBlockXCoords(i) = CLng(SimulBlockXYCoords(0))
                        SimulBlockYCoords(i) = CLng(SimulBlockXYCoords(1))
                    Next i
                    
                    Call PutVar(App.Path & "\SimulBlocks.ini", GetPlayerMap(Index), GetPlayerX(Index) & "/" & GetPlayerY(Index), GetPlayerName(Index))
                    
                    Dim SimulBlockVars(0 To 3) As String
                    
                    For i = 0 To 3
                        If SimulBlockXCoords(i) > 0 And SimulBlockYCoords(i) > 0 Then
                            SimulBlockVars(i) = GetVar(App.Path & "\SimulBlocks.ini", GetPlayerMap(Index), SimulBlockXCoords(i) & "/" & SimulBlockYCoords(i))
                            
                            If SimulBlockVars(i) = vbNullString Then
                                Exit Sub
                            End If
                        End If
                    Next i
                    
                    For i = 0 To 3
                        Call PlayerWarp(FindPlayer(SimulBlockVars(i)), GetPlayerMap(Index), .Data1, .Data2)
                        
                        If SimulBlockXCoords(i) > 0 And SimulBlockYCoords(i) > 0 Then
                            Call PutVar(App.Path & "\SimulBlocks.ini", GetPlayerMap(Index), SimulBlockXCoords(i) & "/" & SimulBlockYCoords(i), vbNullString)
                        End If
                    Next i
                    
                    Exit Sub
                End If
            End With
        ' Other
        Case Else
            Exit Sub
    End Select
End Sub

Sub ScriptedSpell(ByVal Index As Long, ByVal Script As Long)
    Dim Attack As Long, Defense As Long
    
    Select Case Script
        ' Attack-Defense Swap
        Case 0
            Attack = Player(Index).Char(Player(Index).CharNum).STR
            Defense = Player(Index).Char(Player(Index).CharNum).DEF
            
            Call SetPlayerSTR(Index, Defense)
            Call SetPlayerDEF(Index, Attack)
            
            Call SendStats(Index)
            
            Call AddNewTimer(Index, 4, 20000)
            Call PlayerMsg(Index, "Your Attack and Defense have been swapped!", WHITE)
            
            Call PutVar(App.Path & "\Scripts\" & "StatIncreases.ini", GetPlayerName(Index), "Attack", "Yes")
            Call PutVar(App.Path & "\Scripts\" & "StatIncreases.ini", GetPlayerName(Index), "Defense", "Yes")
        Exit Sub
        ' Other
        Case Else
            Exit Sub
    End Select
End Sub

Sub ScriptedItem(ByVal Index As Long, ByVal Script As Long)
    Select Case Script
        ' Yoshi Cookie
        Case 0
            Call AddNewTimer(Index, 3, 5000, 1)
            Call TakeItem(Index, 66, 1)
        Exit Sub
        ' Red Gnarantula Cola
        Case 1
            If GetVar(App.Path & "\Scripts\" & "StatIncreases.ini", GetPlayerName(Index), "ItemAttack") <> "Yes" Then
                Call AddNewTimer(Index, 5, 10000)
                Call TakeItem(Index, 70, 1)
                Call PutVar(App.Path & "\Scripts\" & "StatIncreases.ini", GetPlayerName(Index), "ItemAttack", "Yes")
                Call SetPlayerSTR(Index, Player(Index).Char(Player(Index).CharNum).STR * 2)
            Else
                Call PlayerMsg(Index, "You must wait for the effects of your item to end!", WHITE)
            End If
        Exit Sub
        ' Dry Bones Treat
        Case 2
            Call AddNewTimer(Index, 2, 5000, 1)
            Call TakeItem(Index, 92, 1)
        Exit Sub
        ' Blue Gnarantula Cola
        Case 3
            If GetVar(App.Path & "\Scripts\" & "StatIncreases.ini", GetPlayerName(Index), "ItemDefense") <> "Yes" Then
                Call AddNewTimer(Index, 6, 10000)
                Call TakeItem(Index, 135, 1)
                Call PutVar(App.Path & "\Scripts\" & "StatIncreases.ini", GetPlayerName(Index), "ItemDefense", "Yes")
                Call SetPlayerDEF(Index, Player(Index).Char(Player(Index).CharNum).DEF * 2)
            Else
                Call PlayerMsg(Index, "You must wait for the effects of your item to end!", WHITE)
            End If
        Exit Sub
        ' Point Swap - Swaps current HP and FP
        Case 4
            Dim NewHP As Integer, NewFP As Integer
            Dim MaxStat As Long
            
            ' Set NewHP to the player's current FP
            NewHP = GetPlayerMP(Index)
            MaxStat = GetPlayerMaxHP(Index)
            
            ' Don't allow the player to get to 0 HP left
            If NewHP = 0 Then
                NewHP = 1
            ' If the player's new HP value is greater than his/her max HP, set NewHP to the player's max HP
            ElseIf NewHP > MaxStat Then
                NewHP = MaxStat
            End If
            
            ' Set NewFP to the player's current HP
            NewFP = GetPlayerHP(Index)
            MaxStat = GetPlayerMaxMP(Index)
            
            ' If the player's new FP value is greater than his/her max FP, set NewFP to the player's max FP
            If NewFP > MaxStat Then
                NewFP = MaxStat
            End If
            
            Call SetPlayerHP(Index, NewHP)
            Call SetPlayerMP(Index, NewFP)
            
            Call SendHP(Index)
            Call SendMP(Index)
            
            Call TakeItem(Index, 334, 1)
            
            ' Send the player a message
            Call PlayerMsg(Index, "You used the Point Swap to switch your HP and FP!", WHITE)
        Exit Sub
        ' Red Peppers
        Case 5
            If GetVar(App.Path & "\Scripts\" & "StatIncreases.ini", GetPlayerName(Index), "ItemAttack") <> "Yes" Then
                Call AddNewTimer(Index, 11, 20000, 1)
                Call TakeItem(Index, 314, 1)
                Call PutVar(App.Path & "\Scripts\" & "StatIncreases.ini", GetPlayerName(Index), "ItemAttack", "Yes")
                Call PutVar(App.Path & "\Scripts\" & "StatIncreases.ini", GetPlayerName(Index), "RedPeppers", "Yes")
                Call SetPlayerSTR(Index, Player(Index).Char(Player(Index).CharNum).STR * 1.5)
            Else
                Call PlayerMsg(Index, "You must wait for the effects of your item to end!", WHITE)
            End If
        Exit Sub
        ' Green Peppers
        Case 6
            If GetVar(App.Path & "\Scripts\" & "StatIncreases.ini", GetPlayerName(Index), "ItemDefense") <> "Yes" Then
                Call AddNewTimer(Index, 11, 20000, 2)
                Call TakeItem(Index, 315, 1)
                Call PutVar(App.Path & "\Scripts\" & "StatIncreases.ini", GetPlayerName(Index), "ItemDefense", "Yes")
                Call PutVar(App.Path & "\Scripts\" & "StatIncreases.ini", GetPlayerName(Index), "GreenPeppers", "Yes")
                Call SetPlayerDEF(Index, Player(Index).Char(Player(Index).CharNum).DEF * 1.5)
            Else
                Call PlayerMsg(Index, "You must wait for the effects of your item to end!", WHITE)
            End If
        Exit Sub
        ' Nuts
        Case 7
            Call AddNewTimer(Index, 12, 3000, 1)
            Call TakeItem(Index, 283, 1)
        Exit Sub
        ' Ultra Nuts
        Case 8
            Call AddNewTimer(Index, 13, 1000, 1)
            Call TakeItem(Index, 284, 1)
        Exit Sub
        ' Other
        Case Else
            Exit Sub
    End Select
End Sub

Sub RespawnMap(ByVal MapNum As Long)
    Dim i As Long, Y As Long
    
    ' Respawn the beans on the map
    For Y = 1 To MAX_MAPY
        For i = 1 To MAX_MAPX
            With Map(MapNum).Tile(i, Y)
                If .Type = TILE_TYPE_BEAN Then
                    .Data3 = 0
                End If
            End With
        Next i
    Next Y
    
    ' Clear out all of the floor items
    For i = 1 To MAX_MAP_ITEMS
        Call SpawnItemSlot(i, 0, 0, -1, MapNum, MapItem(MapNum, i).X, MapItem(MapNum, i).Y)
        Call ClearMapItem(i, MapNum)
    Next i

    ' Respawn all of the floor items
    Call SpawnMapItems(MapNum)

    ' Respawn NPCS
    For i = 1 To MAX_MAP_NPCS
        Call SpawnNpc(i, MapNum)
    Next i
End Sub

Sub HammerBarrage(ByVal Index As Long)
    Dim MapNum As Long, X As Long, Y As Long
    
    MapNum = GetPlayerMap(Index)
    X = GetPlayerX(Index)
    Y = GetPlayerY(Index)
    
    ' Send the animation and sound regardless
    Call SendSoundToMap(MapNum, "hammer_barrage_attack.wav")
    Call SpellAnim(29, MapNum, X, Y)
    
    ' Warp the player if possible
    If Map(MapNum).Tile(X, Y).Type = TILE_TYPE_HAMMERBARRAGE Then
        Call PlayerWarp(Index, Map(MapNum).Tile(X, Y).Data1, Map(MapNum).Tile(X, Y).Data2, Map(MapNum).Tile(X, Y).Data3)
    
        Call PlayerMsg(Index, "You use the Hammer Barrage ability to break through the blocks and reach your destination.", WHITE)
    End If
End Sub

Sub JugemsCloud(ByVal Index As Long)
    Dim MapNum As Long, OldX As Long, OldY As Long, X As Long, Y As Long
    
    MapNum = GetPlayerMap(Index)
    OldX = GetPlayerX(Index)
    OldY = GetPlayerY(Index)
    
    X = OldX
    Y = OldY
    
    ' Stop the player from using the Jugem's Cloud when facing a direction that would render the Jugem's Cloud past the map borders
    ' Also, find out the X or Y to render the Jugem's Cloud
    Select Case GetPlayerDir(Index)
        Case DIR_LEFT
            X = OldX - 1
            
            If X < 0 Then
                Exit Sub
            End If
        Case DIR_RIGHT
            X = OldX + 1
            
            If X > MAX_MAPX Then
                Exit Sub
            End If
        Case DIR_UP
            Y = OldY - 1
            
            If Y < 0 Then
                Exit Sub
            End If
        Case DIR_DOWN
            Y = OldY + 1
            
            If Y > MAX_MAPY Then
                Exit Sub
            End If
    End Select
    
    ' Stop the Jugem's Cloud from being used on a sign
    If Map(MapNum).Tile(X, Y).Type = TILE_TYPE_SIGN Then
        Exit Sub
    End If
    
    ' Send the animation and sound regardless
    Call SendSoundToMap(MapNum, "smas-smb3_tankooki_leaf.wav")
    
    ' Play the poison Jugem's Cloud animation in the Poison Cave
    If MapNum >= 320 And MapNum <= 328 Then
        Call SpellAnim(44, MapNum, X, Y, Index)
        
        ' Send the player a message
        Call PlayerMsg(Index, "Your Jugem's Cloud is poisoned in here! It's too dangerous to ride!", WHITE)
    Else
        ' Play the normal animation otherwise
        Call SpellAnim(42, MapNum, X, Y, Index)
    End If
End Sub

Sub MutePlayer(ByVal PlayerBeingMuted As Long, Index As Long)
    Dim FileName As String, MuteList As String, NumString As String
    Dim i As Integer
    Dim FileID As Long

    If Player(PlayerBeingMuted).Mute = True Then
        If Index > 0 Then
            Call PlayerMsg(Index, GetPlayerName(PlayerBeingMuted) & " is already muted!", WHITE)
        End If
        
        Exit Sub
    End If

    FileName = App.Path & "\SMBOMuteList.ini"
    
    For i = 1 To 100
        NumString = CStr(i)
        
        If GetVar(FileName, "Mute List", NumString) = "" Then
            ' Add the player to the mute list INI file
            Call PutVar(FileName, "Mute List", NumString, Trim$(Player(PlayerBeingMuted).Login) & "/" & GetPlayerName(PlayerBeingMuted))
            Exit For
        End If
    Next
    
    Player(PlayerBeingMuted).Mute = True
    Call PlayerMsg(PlayerBeingMuted, "You have been muted!", WHITE)
    
    If Index > 0 Then
        Call AddLog(GetPlayerName(Index) & " has muted " & GetPlayerName(PlayerBeingMuted) & ".", ADMIN_LOG)
        Call PlayerMsg(Index, "You have muted " & GetPlayerName(PlayerBeingMuted) & "!", WHITE)
    Else
        Call AddLog(GetPlayerName(PlayerBeingMuted) & " has been muted by the server.", ADMIN_LOG)
    End If
End Sub

Sub UnmutePlayer(Optional ByVal MuteListNum As Integer = 0, Optional ByVal PlayerBeingUnmuted As Long = 0, Optional ByVal Index As Long = 0)
    Dim FileName As String
    Dim i As Integer
    
    FileName = App.Path & "\SMBOMuteList.INI"
    
    If MuteListNum > 0 Then
        Call PutVar(FileName, "Mute List", CStr(MuteListNum), "")
        
        If Index > 0 Then
            Call PlayerMsg(Index, "You have successfully removed entry #" & MuteListNum & " from the mute list!", YELLOW)
        End If
    Else
        Player(PlayerBeingUnmuted).Mute = False
        Call PlayerMsg(PlayerBeingUnmuted, "You have been unmuted!", YELLOW)
        
        If Index > 0 Then
            Call PlayerMsg(Index, "You have unmuted " & GetPlayerName(PlayerBeingUnmuted) & "!", YELLOW)
        End If
        
        Dim NumString As String
        
        For i = 1 To 100
            NumString = CStr(i)
            
            If GetVar(FileName, "Mute List", NumString) = Trim$(Player(PlayerBeingUnmuted).Login) & "/" & GetPlayerName(PlayerBeingUnmuted) Then
                Call PutVar(FileName, "Mute List", NumString, "")
                Exit Sub
            End If
        Next
    End If
End Sub

Function IsPlayerMuted(ByVal Index As Long) As Boolean
    If Player(Index).Mute = True And GetPlayerAccess(Index) < 1 Then
        IsPlayerMuted = True
        Exit Function
    End If
    
    IsPlayerMuted = False
End Function

Function DamageUpDamageDown(ByVal Index As Long, ByVal Damage As Long, Optional ByVal Target As Long = 0) As Long
    ' Target MUST be a player

    DamageUpDamageDown = Damage
    
    ' Damage Up badge - increase damage dealt and taken by 30%
    If GetPlayerEquipSlotNum(Index, 4) = 268 Then
        DamageUpDamageDown = Damage * 1.3
        ' Damage Down badge - decrease damage dealt and taken by 30%
    ElseIf GetPlayerEquipSlotNum(Index, 4) = 269 Then
        DamageUpDamageDown = Damage * 0.7
    End If
    
    If Target > 0 Then
        ' Damage Up badge - increase damage dealt and taken by 30%
        If GetPlayerEquipSlotNum(Target, 4) = 268 Then
            DamageUpDamageDown = Damage * 1.3
            ' Damage Down badge - decrease damage dealt and taken by 30%
        ElseIf GetPlayerEquipSlotNum(Target, 4) = 269 Then
            DamageUpDamageDown = Damage * 0.7
        End If
    End If
End Function

Function PityFlower(ByVal Index As Long, ByVal BaseDamage As Long) As Long
    ' Pity Flower badge - increases base special attack damage by 30%
    If GetPlayerEquipSlotNum(Index, 6) = 316 Then
        PityFlower = BaseDamage * 1.3
    Else
        PityFlower = BaseDamage
    End If
End Function

Function FlowerSaver(ByVal Index As Long, ByVal SpellNum As Long) As Long
    ' Account for the Flower Saver special attack
    FlowerSaver = Spell(SpellNum).MPCost
    
    ' Check if the player has the Flower Saver special attack
    If HasSpell(Index, 43) = True Then
        ' Reduce the cost of the special attack by 30% and always round up
        FlowerSaver = Int((-FlowerSaver * 0.7)) * -1
    End If
End Function

Sub AddCastleTownFavorComplete(ByVal Index As Long)
    Dim FavorsCompleted As String
    Dim FavorNum As Byte
                        
    FavorsCompleted = GetVar(App.Path & "\Scripts\CastleTownFavors.ini", GetPlayerName(Index), "FavorsCompleted")
                        
    If FavorsCompleted <> vbNullString Then
        FavorNum = CByte(FavorsCompleted) + 1
    Else
        FavorNum = 1
    End If
    
    Call PutVar(App.Path & "\Scripts\CastleTownFavors.ini", GetPlayerName(Index), "FavorsCompleted", CStr(FavorNum))
End Sub

Function IsInPoisonCave(ByVal Index As Long) As Boolean
    If GetPlayerMap(Index) >= 320 And GetPlayerMap(Index) <= 328 Then
        IsInPoisonCave = True
    End If
End Function
