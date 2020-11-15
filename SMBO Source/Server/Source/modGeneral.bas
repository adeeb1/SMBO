Attribute VB_Name = "modGeneral"
Option Explicit

Sub InitServer()
    Dim Index As Integer

    Call SetStatus("Checking Folders...")

    ' Check folders
    If Not FolderExists(App.Path & "\Maps") Then
        Call MkDir(App.Path & "\Maps")
    End If

    If Not FolderExists(App.Path & "\Logs") Then
        Call MkDir(App.Path & "\Logs")
    End If

    If Not FolderExists(App.Path & "\SMBOAccounts") Then
        Call MkDir(App.Path & "\SMBOAccounts")
    End If

    If Not FolderExists(App.Path & "\NPCs") Then
        Call MkDir(App.Path & "\NPCs")
    End If

    If Not FolderExists(App.Path & "\Items") Then
        Call MkDir(App.Path & "\Items")
    End If

    If Not FolderExists(App.Path & "\Spells") Then
        Call MkDir(App.Path & "\Spells")
    End If

    If Not FolderExists(App.Path & "\Shops") Then
        Call MkDir(App.Path & "\Shops")
    End If
    
    If Not FolderExists(App.Path & "\SMBOClasses") Then
        Call MkDir(App.Path & "\SMBOClasses")
    End If
    
    If Not FolderExists(App.Path & "\Recipes") Then
        Call MkDir(App.Path & "\Recipes")
    End If
    
    Call SetStatus("Checking Files...")

    If Not FileExists("News.ini") Then
        PutVar App.Path & "\News.ini", "DATA", "NewsTitle", "Change this message in News.ini."
        PutVar App.Path & "\News.ini", "DATA", "NewsBody", "Change this message in News.ini."
        PutVar App.Path & "\News.ini", "COLOR", "Red", 255
        PutVar App.Path & "\News.ini", "COLOR", "Green", 255
        PutVar App.Path & "\News.ini", "COLOR", "Blue", 255
    End If
    
    If Not FileExists("MOTD.ini") Then
        PutVar App.Path & "\MOTD.ini", "MOTD", "Msg", "Change this message in MOTD.ini."
    End If

    If Not FileExists("Tiles.ini") Then
        For Index = 0 To 100
            PutVar App.Path & "\Tiles.ini", "Names", "Tile" & Index, CStr(Index)
        Next Index
    End If

    ' Set the minigame file paths
    STSPath = App.Path & "\Scripts\" & "CTF.ini"
    DodgeBillPath = App.Path & "\Scripts\" & "Dodgeball.ini"
    HideNSneakPath = App.Path & "\Scripts\" & "HideNSneak.ini"

    ' Check if the master charlist file exists for checking duplicate names, and if it doesnt make it
    If Not FileExists("SMBOAccounts\Charlist.txt") Then
        Index = FreeFile
        Open App.Path & "\SMBOAccounts\CharList.txt" For Output As #Index
        Close #Index
    End If

    Call SetStatus("Loading Settings...")

    On Error GoTo LoadErr
    
    addSP.LEVEL = 2
    addSP.STR = 0
    addSP.DEF = 0
    addSP.Magi = 0
    addSP.Speed = 3

    SAVETIME = 15
    LEVEL = 1

    ' Weather variables
    WeatherType = WEATHER_NONE
    WeatherLevel = 25
    
    ' Get the listening socket ready to go
    frmServer.Socket(0).RemoteHost = frmServer.Socket(0).LocalIP
    frmServer.Socket(0).LocalPort = GAME_PORT
    
    SEP_CHAR = Chr$(0)
    END_CHAR = Chr$(237)

    ServerLog = True
    
    GoTo LoadSuccess

LoadErr:
    Call MsgBox(Err.Description, vbOKOnly)
    End

LoadSuccess:
    ' Restore error handling
    On Error GoTo 0

    For Index = 1 To MAX_MAPS
        ReDim Map(Index).Tile(0 To MAX_MAPX, 0 To MAX_MAPY) As TileRec
        ReDim TempTile(Index).DoorOpen(0 To MAX_MAPX, 0 To MAX_MAPY) As Byte
        ReDim TempTile(Index).DoorTimer(0 To MAX_MAPX, 0 To MAX_MAPY) As Long
    Next Index
    
    ReDim QuestionBlock(1 To MAX_MAPS, 0 To MAX_MAPX, 0 To MAX_MAPY) As QuestionBlockRec
    
    START_MAP = 1
    START_X = MAX_MAPX / 2
    START_Y = MAX_MAPY / 2

    Call IncrementBar
    
    On Error GoTo 0

    ' Init all the player sockets
    Call SetStatus("Initializing player array...")
    
    For Index = 1 To MAX_PLAYERS
        Call ClearPlayer(Index)
        Load frmServer.Socket(Index)
    Next Index

    For Index = 1 To MAX_PLAYERS
        Call ShowPLR(Index)
    Next Index
    
    Call IncrementBar

    Call SetStatus("Clearing temp tile fields...")
    Call ClearTempTile
    Call SetStatus("Clearing maps...")
    Call ClearMaps
    Call SetStatus("Clearing map items...")
    Call ClearMapItems
    Call SetStatus("Clearing map npcs...")
    Call ClearMapNpcs
    Call SetStatus("Clearing npcs...")
    Call ClearNpcs
    Call SetStatus("Clearing items...")
    Call ClearItems
    Call SetStatus("Clearing shops...")
    Call ClearShops
    Call SetStatus("Clearing spells...")
    Call ClearSpells
    Call SetStatus("Clearing recipes...")
    Call ClearRecipes
    Call SetStatus("Clearing exp...")
    Call ClearExperience
    Call SetStatus("Clearing emoticons...")
    Call ClearEmoticon
    Call IncrementBar
    Call SetStatus("Loading emoticons...")
    Call IncrementBar
    Call LoadEmoticon
    Call SetStatus("Loading elements...")
    Call IncrementBar
    Call LoadElements
    Call SetStatus("Clearing arrows...")
    Call ClearArrows
    Call SetStatus("Loading arrows...")
    Call IncrementBar
    Call LoadArrows
    Call SetStatus("Loading exp...")
    Call IncrementBar
    Call LoadExperience
    Call SetStatus("Loading classes...")
    Call IncrementBar
    Call LoadClasses
    Call SetStatus("Loading maps...")
    Call IncrementBar
    Call LoadMaps
    Call SetStatus("Loading items...")
    Call IncrementBar
    Call LoadItems
    Call SetStatus("Loading npcs...")
    Call IncrementBar
    Call LoadNpcs
    Call SetStatus("Loading shops...")
    Call IncrementBar
    Call LoadShops
    Call SetStatus("Loading spells...")
    Call IncrementBar
    Call LoadSpells
    Call SetStatus("Loading recipes...")
    Call IncrementBar
    Call LoadRecipes
    Call SetStatus("Spawning map items...")
    Call SpawnAllMapsItems
    Call SetStatus("Spawning map npcs...")
    Call SpawnAllMapNpcs
    Call IncrementBar
    Call SetStatus("Creating map cache...")
    Call IncrementBar
    Call CreateFullMapCache
    Call SetStatus("Loading question blocks...")
    Call IncrementBar
    Call LoadQuestionBlocks

    frmServer.MapList.Clear

    For Index = 1 To MAX_MAPS
        frmServer.MapList.AddItem Index & ": " & Map(Index).Name
    Next Index

    frmServer.MapList.Selected(0) = True
    frmServer.tmrPlayerSave.Enabled = True
    frmServer.tmrSpawnMapItems.Enabled = True
    frmServer.Timer1.Enabled = True

    ' Error handling for 'Address in use' error
    Err.Clear
    On Error Resume Next

    ' Start listening
    frmServer.Socket(0).Listen

    ' RTE 10048 occured
    If Err.Number = 10048 Then
        Call MsgBox("The port on this address is already busy.", vbOKOnly)
        End
    End If

    ' Restore error handling
    On Error GoTo 0

    Call UpdateTitle
    Call UpdateTOP

    frmLoad.Visible = False
    frmServer.Show

    SpawnSeconds = 0
    frmServer.tmrGameAI.Enabled = True
End Sub

Sub DestroyServer()
    Dim i As Long
    
    Call SaveAllPlayersOnline

    frmServer.Visible = False
    frmLoad.Visible = True

    For i = 1 To MAX_PLAYERS
        temp = i / MAX_PLAYERS * 100
        Call SetStatus("Unloading Sockets... " & temp & "%")
        Unload frmServer.Socket(CInt(i))
    Next i

    End
End Sub

Sub SetStatus(ByVal Status As String)
    frmLoad.lblStatus.Caption = Status
    DoEvents
End Sub

Sub IncrementBar()
    On Error Resume Next
    ' Increment prog bar
    frmLoad.loadProgressBar.Value = frmLoad.loadProgressBar.Value + 1
End Sub

Sub ServerLogic()
    Call CheckGiveVitals
    Call GameAI
    Call StartTimers
    Call CheckForDisconnectionsAndDE
End Sub

Sub CheckSpawnMapItems()
    Dim X As Long, Y As Long, i As Long

    ' Used for map item respawning
    SpawnSeconds = SpawnSeconds + 1

    ' Respawns the map items.
    If SpawnSeconds >= 120 Then
        ' 2 minutes have passed
        For Y = 1 To MAX_MAPS
            ' Make sure no one is on the map when it respawns
            If Not PlayersOnMap(Y) And Y <> 188 Then
                ' Respawn the beans on the map
                For i = 1 To MAX_MAPY
                    For X = 1 To MAX_MAPX
                        With Map(Y).Tile(X, i)
                            If .Type = TILE_TYPE_BEAN Then
                                .Data3 = 0
                            End If
                        End With
                    Next X
                Next i
                
                ' Clear out unnecessary junk
                For X = 1 To MAX_MAP_ITEMS
                    Call ClearMapItem(X, Y)
                Next X

                ' Spawn the items
                Call SpawnMapItems(Y)
                Call SendMapItemsToMap(Y)
            End If
        Next Y

        SpawnSeconds = 0
    End If
End Sub

Sub CheckForDisconnectionsAndDE()
    Dim i As Long
    
    For i = 1 To MAX_PLAYERS
        ' Check if the player is still connected
        If frmServer.Socket(i).State > sckConnected Then
            Call CloseSocket(i)
        End If
        
        ' Check if the player should get the Daily Event
        Call CheckForDE(i)
    Next i
End Sub

Public Sub GameAI()
    Dim i As Long, X As Long, Y As Long, n As Long, x1 As Long, y1 As Long, TickCount As Long
    Dim Damage As Long, DistanceX As Long, DistanceY As Long, NpcNum As Long, Target As Long
    Dim DidWalk As Boolean

    On Error Resume Next

    For Y = 1 To MAX_MAPS
        If PlayersOnMap(Y) = YES Then
            TickCount = GetTickCount
            
            ' ////////////////////////////////////
            ' // This is used for closing doors //
            ' ////////////////////////////////////
            For y1 = 0 To MAX_MAPY
                For x1 = 0 To MAX_MAPX
                    If TickCount > TempTile(Y).DoorTimer(x1, y1) + 5000 Then
                        If Map(Y).Tile(x1, y1).Type = TILE_TYPE_KEY Then
                            If TempTile(Y).DoorOpen(x1, y1) = YES Then
                                TempTile(Y).DoorOpen(x1, y1) = NO
                                Call SendMapKey(Y, x1, y1, 0)
                            End If
                        ElseIf Map(Y).Tile(x1, y1).Type = TILE_TYPE_DOOR Then
                            If TempTile(Y).DoorOpen(x1, y1) = YES Then
                                TempTile(Y).DoorOpen(x1, y1) = NO
                                Call SendMapKey(Y, x1, y1, 0)
                            End If
                        End If
                    End If
                Next x1
            Next y1
            
            For X = 1 To MAX_MAP_NPCS
                NpcNum = MapNPC(Y, X).num
                
                ' /////////////////////////////////////////
                ' // This is used for ATTACKING ON SIGHT //
                ' /////////////////////////////////////////
                ' Make sure theres a npc with the map
                If Map(Y).NPC(X) > 0 Then
                    If MapNPC(Y, X).num > 0 Then
                        ' If the npc is a attack on sight, search for a player on the map
                        If NPC(NpcNum).Behavior = NPC_BEHAVIOR_ATTACKONSIGHT Or NPC(NpcNum).Behavior = NPC_BEHAVIOR_GUARD Then
                            For i = 1 To MAX_PLAYERS
                                If IsPlaying(i) Then
                                    If GetPlayerMap(i) = Y Then
                                        If MapNPC(Y, X).Target = 0 Then
                                            If GetPlayerAccess(i) <= ADMIN_MONITER Then
                                              If GetNpcLevel(MapNPC(Y, X).num) = 0 Or GetNpcLevel(MapNPC(Y, X).num) + 4 > GetPlayerLevel(i) Then
                                                n = NPC(NpcNum).Range
                                                
                                                DistanceX = MapNPC(Y, X).X - GetPlayerX(i)
                                                DistanceY = MapNPC(Y, X).Y - GetPlayerY(i)
                                                
                                                ' Make sure we get a positive value
                                                If DistanceX < 0 Then DistanceX = DistanceX * -1
                                                If DistanceY < 0 Then DistanceY = DistanceY * -1
                                                
                                                ' Are they in range?  if so GET'M!
                                                If DistanceX <= n And DistanceY <= n Then
                                                    If NPC(NpcNum).Behavior = NPC_BEHAVIOR_ATTACKONSIGHT Then
                                                        If GetPlayerInBattle(i) = False Then
                                                            MapNPC(Y, X).Target = i
                                                        End If
                                                    End If
                                                End If
                                              End If
                                            End If
                                        End If
                                    End If
                                End If
                            Next i
                        End If
                    End If
                End If
                                                                        
                ' /////////////////////////////////////////////
                ' // This is used for NPC walking/targetting //
                ' /////////////////////////////////////////////
                ' Make sure theres a npc with the map
                If Map(Y).NPC(X) > 0 Then
                    If MapNPC(Y, X).num > 0 Then
                        Target = MapNPC(Y, X).Target
                        
                        ' Check to see if its time for the npc to walk
                        If NPC(NpcNum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then
                            ' Check to see if we are following a player or not
                            If Target > 0 Then
                                ' Check if the player is even playing, if so follow'm
                                If IsPlaying(Target) Then
                                    If GetPlayerMap(Target) = Y Then
                                        DidWalk = False
                                        
                                        i = Int(Rnd2 * 4)
                                        
                                        ' Lets move the npc
                                        Select Case i
                                            Case 0
                                                ' Up
                                                If MapNPC(Y, X).Y > GetPlayerY(Target) Then
                                                    If DidWalk = False Then
                                                        If CanNpcMove(Y, X, DIR_UP) Then
                                                            Call NpcMove(Y, X, DIR_UP, MOVING_WALKING)
                                                            DidWalk = True
                                                        End If
                                                    End If
                                                End If
                                                ' Down
                                                If MapNPC(Y, X).Y < GetPlayerY(Target) Then
                                                    If DidWalk = False Then
                                                        If CanNpcMove(Y, X, DIR_DOWN) Then
                                                            Call NpcMove(Y, X, DIR_DOWN, MOVING_WALKING)
                                                            DidWalk = True
                                                        End If
                                                    End If
                                                End If
                                                ' Left
                                                If MapNPC(Y, X).X > GetPlayerX(Target) Then
                                                    If DidWalk = False Then
                                                        If CanNpcMove(Y, X, DIR_LEFT) Then
                                                            Call NpcMove(Y, X, DIR_LEFT, MOVING_WALKING)
                                                            DidWalk = True
                                                        End If
                                                    End If
                                                End If
                                                ' Right
                                                If MapNPC(Y, X).X < GetPlayerX(Target) Then
                                                    If DidWalk = False Then
                                                        If CanNpcMove(Y, X, DIR_RIGHT) Then
                                                            Call NpcMove(Y, X, DIR_RIGHT, MOVING_WALKING)
                                                            DidWalk = True
                                                        End If
                                                    End If
                                                End If
                                            
                                            Case 1
                                                ' Right
                                                If MapNPC(Y, X).X < GetPlayerX(Target) Then
                                                    If DidWalk = False Then
                                                        If CanNpcMove(Y, X, DIR_RIGHT) Then
                                                            Call NpcMove(Y, X, DIR_RIGHT, MOVING_WALKING)
                                                            DidWalk = True
                                                        End If
                                                    End If
                                                End If
                                                ' Left
                                                If MapNPC(Y, X).X > GetPlayerX(Target) Then
                                                    If DidWalk = False Then
                                                        If CanNpcMove(Y, X, DIR_LEFT) Then
                                                            Call NpcMove(Y, X, DIR_LEFT, MOVING_WALKING)
                                                            DidWalk = True
                                                        End If
                                                    End If
                                                End If
                                                ' Down
                                                If MapNPC(Y, X).Y < GetPlayerY(Target) Then
                                                    If DidWalk = False Then
                                                        If CanNpcMove(Y, X, DIR_DOWN) Then
                                                            Call NpcMove(Y, X, DIR_DOWN, MOVING_WALKING)
                                                            DidWalk = True
                                                        End If
                                                    End If
                                                End If
                                                ' Up
                                                If MapNPC(Y, X).Y > GetPlayerY(Target) Then
                                                    If DidWalk = False Then
                                                        If CanNpcMove(Y, X, DIR_UP) Then
                                                            Call NpcMove(Y, X, DIR_UP, MOVING_WALKING)
                                                            DidWalk = True
                                                        End If
                                                    End If
                                                End If
                                                
                                            Case 2
                                                ' Down
                                                If MapNPC(Y, X).Y < GetPlayerY(Target) Then
                                                    If DidWalk = False Then
                                                        If CanNpcMove(Y, X, DIR_DOWN) Then
                                                            Call NpcMove(Y, X, DIR_DOWN, MOVING_WALKING)
                                                            DidWalk = True
                                                        End If
                                                    End If
                                                End If
                                                ' Up
                                                If MapNPC(Y, X).Y > GetPlayerY(Target) Then
                                                    If DidWalk = False Then
                                                        If CanNpcMove(Y, X, DIR_UP) Then
                                                            Call NpcMove(Y, X, DIR_UP, MOVING_WALKING)
                                                            DidWalk = True
                                                        End If
                                                    End If
                                                End If
                                                ' Right
                                                If MapNPC(Y, X).X < GetPlayerX(Target) Then
                                                    If DidWalk = False Then
                                                        If CanNpcMove(Y, X, DIR_RIGHT) Then
                                                            Call NpcMove(Y, X, DIR_RIGHT, MOVING_WALKING)
                                                            DidWalk = True
                                                        End If
                                                    End If
                                                End If
                                                ' Left
                                                If MapNPC(Y, X).X > GetPlayerX(Target) Then
                                                    If DidWalk = False Then
                                                        If CanNpcMove(Y, X, DIR_LEFT) Then
                                                            Call NpcMove(Y, X, DIR_LEFT, MOVING_WALKING)
                                                            DidWalk = True
                                                        End If
                                                    End If
                                                End If
                                            
                                            Case 3
                                                ' Left
                                                If MapNPC(Y, X).X > GetPlayerX(Target) Then
                                                    If DidWalk = False Then
                                                        If CanNpcMove(Y, X, DIR_LEFT) Then
                                                            Call NpcMove(Y, X, DIR_LEFT, MOVING_WALKING)
                                                            DidWalk = True
                                                        End If
                                                    End If
                                                End If
                                                ' Right
                                                If MapNPC(Y, X).X < GetPlayerX(Target) Then
                                                    If DidWalk = False Then
                                                        If CanNpcMove(Y, X, DIR_RIGHT) Then
                                                            Call NpcMove(Y, X, DIR_RIGHT, MOVING_WALKING)
                                                            DidWalk = True
                                                        End If
                                                    End If
                                                End If
                                                ' Up
                                                If MapNPC(Y, X).Y > GetPlayerY(Target) Then
                                                    If DidWalk = False Then
                                                        If CanNpcMove(Y, X, DIR_UP) Then
                                                            Call NpcMove(Y, X, DIR_UP, MOVING_WALKING)
                                                            DidWalk = True
                                                        End If
                                                    End If
                                                End If
                                                ' Down
                                                If MapNPC(Y, X).Y < GetPlayerY(Target) Then
                                                    If DidWalk = False Then
                                                        If CanNpcMove(Y, X, DIR_DOWN) Then
                                                            Call NpcMove(Y, X, DIR_DOWN, MOVING_WALKING)
                                                            DidWalk = True
                                                        End If
                                                    End If
                                                End If
                                        End Select
                                    
                                        ' Check if we can't move and if player is behind something and if we can just switch dirs
                                        If Not DidWalk Then
                                            If MapNPC(Y, X).X - 1 = GetPlayerX(Target) Then
                                                If MapNPC(Y, X).Y = GetPlayerY(Target) Then
                                                    If MapNPC(Y, X).Dir <> DIR_LEFT Then
                                                        Call NpcDir(Y, X, DIR_LEFT)
                                                    End If
                                                End If
                                                DidWalk = True
                                            End If
                                            If MapNPC(Y, X).X + 1 = GetPlayerX(Target) Then
                                                If MapNPC(Y, X).Y = GetPlayerY(Target) Then
                                                    If MapNPC(Y, X).Dir <> DIR_RIGHT Then
                                                        Call NpcDir(Y, X, DIR_RIGHT)
                                                    End If
                                                End If
                                                DidWalk = True
                                            End If
                                            If MapNPC(Y, X).X = GetPlayerX(Target) Then
                                                If MapNPC(Y, X).Y - 1 = GetPlayerY(Target) Then
                                                    If MapNPC(Y, X).Dir <> DIR_UP Then
                                                        Call NpcDir(Y, X, DIR_UP)
                                                    End If
                                                End If
                                                DidWalk = True
                                            End If
                                            If MapNPC(Y, X).X = GetPlayerX(Target) Then
                                                If MapNPC(Y, X).Y + 1 = GetPlayerY(Target) Then
                                                    If MapNPC(Y, X).Dir <> DIR_DOWN Then
                                                        Call NpcDir(Y, X, DIR_DOWN)
                                                    End If
                                                End If
                                                DidWalk = True
                                            End If
                                            
                                            ' We could not move so player must be behind something, walk randomly.
                                            If Not DidWalk Then
                                                i = Int(Rnd2 * 2)
                                                If i = 1 Then
                                                    i = Int(Rnd2 * 4)
                                                    If CanNpcMove(Y, X, i) Then
                                                        Call NpcMove(Y, X, i, MOVING_WALKING)
                                                    End If
                                                End If
                                            End If
                                        End If
                                    Else
                                        MapNPC(Y, X).Target = 0
                                    End If
                                End If
                                
                            Else
                                i = Int(Rnd2 * 4)
                                If i = 1 Then
                                    i = Int(Rnd2 * 4)
                                    If CanNpcMove(Y, X, i) Then
                                        Call NpcMove(Y, X, i, MOVING_WALKING)
                                    End If
                                End If
                            End If
                            
                        End If
                    End If
                    
                End If
                
                ' /////////////////////////////////////////////
                ' // This is used for npcs to attack players //
                ' /////////////////////////////////////////////
                ' Make sure theres a npc with the map
                If Map(Y).NPC(X) > 0 Then
                    If MapNPC(Y, X).num > 0 Then
                        Target = MapNPC(Y, X).Target
                        
                        ' Check if the npc can attack the targeted player player
                        If Target > 0 Then
                            ' Is the target playing and on the same map?
                            If IsPlaying(Target) Then
                                If GetPlayerMap(Target) = Y Then
                                    ' Can the npc attack the player?
                                    If CanNpcAttackPlayer(X, Target) Then
                                        If Not CanPlayerBlockHit(Target) Then
                                        
                                            Damage = NPC(NpcNum).STR - GetPlayerProtection(Target)
                                            
                                            Call NpcAttackPlayer(X, Target, Damage)
                                        Else
                                            Call BattleMsg(Target, "You blocked the " & Trim(NPC(NpcNum).Name) & "'s hit!", BRIGHTCYAN, 1)
                                            Call SendDataToMap(GetPlayerMap(Target), SPackets.Ssound & SEP_CHAR & "miss" & END_CHAR)
                                        End If
                                    End If
                                End If
                            Else
                                ' Player left map or game, set target to 0
                                MapNPC(Y, X).Target = 0
                            End If
                        End If

                    End If
                End If
                
                ' /////////////////////////////////////////
                ' // This is used for Turn-based battles //
                ' /////////////////////////////////////////
                ' Make sure theres a npc with the map
                If Map(Y).NPC(X) > 0 Then
                    If MapNPC(Y, X).num > 0 Then
                        Target = MapNPC(Y, X).Target
                        If CanNpcAttackPlayer(X, Target) Then
                            If IsPlaying(Target) Then
                                If GetPlayerMap(Target) = Y Then
                                    If MapNPC(Y, X).InBattle = True Then
                                        Call TurnBasedBattle(Target, X)
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
                
                ' //////////////////////////////////////
                ' // This is used for spawning an NPC //
                ' //////////////////////////////////////
                ' Check if we are supposed to spawn an npc or not
                If MapNPC(Y, X).num = 0 Then
                    If Map(Y).NPC(X) > 0 Then
                        If TickCount > MapNPC(Y, X).SpawnWait + (NPC(Map(Y).NPC(X)).SpawnSecs * 1000) Then
                            Call SpawnNpc(X, Y)
                        End If
                    End If
                End If
            Next X
            
            ' Poison Cave - Subtract 1 HP per second from players inside it
            If Y >= 320 And Y <= 328 Then
                If TickCount > LoseHPTimer + 1000 Then
                    For i = 1 To MAX_PLAYERS
                        If IsPlaying(i) Then
                            If IsInPoisonCave(i) Then
                                If IsInVictoryAnim(i) = False Then
                                    Call SetPlayerHP(i, GetPlayerHP(i) - 1)
                                    Call SendHP(i)
                                        
                                    ' Show the damage being done to the player
                                    Call SendDataTo(i, SPackets.Sblitnpcdmg & SEP_CHAR & 1 & END_CHAR)
                                End If
                            End If
                        End If
                    Next i
                    
                    LoseHPTimer = TickCount
                End If
            End If
        End If
    Next Y

    ' Resets the timer for door closing
    If TickCount > KeyTimer + 15000 Then
        KeyTimer = TickCount
    End If
End Sub

Sub Timers(ByVal Index As Long, ByVal TimeNum As Long)
    Dim TimerNum As Long, Parameter1 As Long, Parameter2 As Long, Parameter3 As Long
    
    TimerNum = Timer(TimeNum).num
    Parameter1 = Timer(TimeNum).Parameter1
    Parameter2 = Timer(TimeNum).Parameter2
    Parameter3 = Timer(TimeNum).Parameter3
    
    Select Case TimerNum
        Case 0
            Call STSPlayTime(Index)
        Case 1
            Call StatIncrease(Index, Parameter1)
        Case 2
            Call DryTreatHeal(Index, Parameter1)
        Case 3
            Call CookieHeal(Index, Parameter1)
        Case 4
            Call StatSwap(Index)
        Case 5
            Call AttackDouble(Index)
        Case 6
            Call DefenseDouble(Index)
        Case 7
            Call WhackAMontyTime(Index, Parameter1)
        Case 8
            Call MontyRespawn(Index, Parameter1, Parameter2)
        Case 9
            Call DodgeBallPlayTime(Index)
        Case 10
            Call HideNSneakPlayTime(Index)
        Case 11
            Call RedGreenPeppers(Index, Parameter1)
        Case 12
            Call Nuts(Index, Parameter1)
        Case 13
            Call UltraNuts(Index, Parameter1)
        Case Else
            Exit Sub
    End Select
End Sub

Sub CheckGiveVitals()
    Dim i As Long

    ' SP Regeneration
    If GetTickCount >= GiveSPTimer + 3000 Then
        For i = 1 To MAX_PLAYERS
            If IsPlaying(i) Then
                If GetPlayerSP(i) < GetPlayerMaxSP(i) Then
                    Call SetPlayerSP(i, GetPlayerSP(i) + GetPlayerSPRegen(i))
                    Call SendSP(i)
                End If
            End If
        Next i

        GiveSPTimer = GetTickCount
    End If
End Sub

Sub StartTimers()
    Dim i As Long, Index As Long, TimerNum As Long, MapNum As Long
    
    For i = 1 To MAX_PLAYERS
        If Timer(i).WaitTime > 0 Then
            TimerNum = Timer(i).num
            
            If GetTickCount > Timer(i).WaitTime Then
                Index = FindPlayer(Timer(i).Player)
                
                If Index > 0 Then
                    Call Timers(Index, i)
                Else
                    ' Get rid of the multiplayer-based minigame timers if there are no players on the map
                    Select Case TimerNum
                        ' STS
                        Case 0
                            If PlayersOnMap(33) = NO Then
                                Call GetRidOfTimer(Timer(i).Index, TimerNum)
                            Else
                                Call Timers(Timer(i).Index, i)
                            End If
                        ' Dodgebill
                        Case 9
                            If PlayersOnMap(188) = NO Then
                                Call GetRidOfTimer(Timer(i).Index, TimerNum)
                            Else
                                Call Timers(Timer(i).Index, i)
                            End If
                        ' Hide n' Sneak
                        Case 10
                            If PlayersOnMap(271) = NO And PlayersOnMap(272) = NO And PlayersOnMap(273) = NO Then
                                Call GetRidOfTimer(Timer(i).Index, TimerNum)
                            Else
                                Call Timers(Timer(i).Index, i)
                            End If
                        ' Non-multiplayer-based minigame timers
                        Case Else
                            ' Get rid of all timers except multiplayer-based minigames
                            Call GetRidOfTimer(Timer(i).Index, TimerNum)
                    End Select
                End If
            End If
        End If
    Next i
End Sub

Sub PlayerSaveTimer()
    Dim i As Long

    PLYRSAVE_TIMER = PLYRSAVE_TIMER + 1

    If SAVETIME <> 0 Then
        If PLYRSAVE_TIMER >= SAVETIME Then
            For i = 1 To MAX_PLAYERS
                If IsPlaying(i) Then
                    Call SavePlayer(i)
                End If
            Next i
    
            PLYRSAVE_TIMER = 0
        End If
    Else
        PLYRSAVE_TIMER = 0
    End If
End Sub

Function IsAlphaNumeric(TestString As String) As Boolean
    Dim LoopID As Integer
    Dim sChar As String

    IsAlphaNumeric = False

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

Public Function Rand(ByVal Low As Long, ByVal High As Long) As Long
    Randomize
    Rand = Int((High - Low + 1) * Rnd) + Low
End Function

Public Function Rnd2()
    Randomize
    Rnd2 = Rnd
End Function
