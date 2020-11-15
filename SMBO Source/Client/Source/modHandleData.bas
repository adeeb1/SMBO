Attribute VB_Name = "modHandleData"
Option Explicit

Sub HandleData(ByVal Data As String)
    Dim parse() As String, Name As String, Msg As String, packet As String, s As String, MapNeed As String
    Dim Dir As Long, Level As Long, i As Long, n As Long, x As Long, y As Long, p As Long, a As Long, ShopNum As Long, z As Long
    Dim Q As Integer, F As Integer

    ' Handle Data
    parse = Split(Data, SEP_CHAR)
    
    ' Start handling packets
    Select Case CInt(parse(0))
        Case SPackets.Smaxinfo
            lvl = CInt(parse(1))
    
            For i = 1 To MAX_MAPS
                ReDim Map(i).Tile(0 To MAX_MAPX, 0 To MAX_MAPY) As TileRec
            Next
            
            ReDim QuestionBlock(0 To MAX_MAPS, 0 To MAX_MAPX, 0 To MAX_MAPY) As QuestionBlockRec
            ReDim TempTile(0 To MAX_MAPX, 0 To MAX_MAPY) As TempTileRec
            
            For i = 0 To MAX_EMOTICONS
                Emoticons(i).Pic = 0
                Emoticons(i).Command = vbNullString
            Next i
        
            Call ClearTempTile
                
            MAX_BLT_LINE = 6

            ReDim BattlePMsg(1 To MAX_BLT_LINE) As BattleMsgRec
            ReDim BattleMMsg(1 To MAX_BLT_LINE) As BattleMsgRec
                
            ' Clear out players
            For i = 1 To MAX_PLAYERS
                Call ClearPlayer(i)
            Next i

            For i = 1 To MAX_MAPS
                Call LoadMap(i)
            Next i
            
            frmMirage.Caption = "Super Mario Bros. Online"
            App.Title = "Super Mario Bros. Online"
            
            AllDataReceived = True
        Exit Sub

    ' :::::::::::::::::::
    ' :: Npc hp packet ::
    ' :::::::::::::::::::
        Case SPackets.Snpchp
            n = CLng(parse(1))

            MapNpc(n).HP = CLng(parse(2))
            MapNpc(n).MaxHp = CLng(parse(3))
        Exit Sub

    ' ::::::::::::::::::::::::::
    ' :: Alert message packet ::
    ' ::::::::::::::::::::::::::
        Case SPackets.Salertmsg
            Msg = parse(1)
            Call MsgBox(Msg, vbOKOnly, "Super Mario Bros. Online")
            
            Call GameDestroy
        Exit Sub

    ' ::::::::::::::::::::::::::
    ' :: Plain message packet ::
    ' ::::::::::::::::::::::::::
        Case SPackets.Splainmsg
            frmSendGetData.Visible = False
            n = CLng(parse(2))
            
            Select Case n
                Case 0
                    frmMainMenu.Show
                Case 1
                    frmNewAccount.Show
                Case 2
                    frmDeleteAccount.Show
                Case 3
                    frmLogin.Show
                Case 4
                    frmNewChar.Show
                Case 5
                    frmChars.Show
            End Select

            Msg = parse(1)
            Call MsgBox(Msg, vbOKOnly, "Super Mario Bros. Online")
        Exit Sub

    ' :::::::::::::::::::::::::::
    ' :: All characters packet ::
    ' :::::::::::::::::::::::::::
        Case SPackets.Sallchars
            n = 1

            frmChars.Visible = True
            frmSendGetData.Visible = False

            frmChars.lstChars.Clear

            For i = 1 To MAX_CHARS
                Name = Trim$(parse(n))
                Msg = Trim$(parse(n + 1))
                Level = CInt(parse(n + 2))

                If Name = vbNullString Then
                    frmChars.lstChars.addItem "Free Character Slot"
                Else
                    frmChars.lstChars.addItem Name & " (" & Msg & ": Level " & Level & ")"
                End If

                n = n + 3
            Next i

            frmChars.lstChars.ListIndex = 0
        Exit Sub

    ' :::::::::::::::::::::::::::::::::
    ' :: Login was successful packet ::
    ' :::::::::::::::::::::::::::::::::
        Case SPackets.Sloginok
            ' Now we can receive game data
            MyIndex = CLng(parse(1))
            
            InventorySlotsIndex = 1
            
            frmSendGetData.Visible = True
            frmChars.Visible = False
            
            JugemsCloudHolder = "SMBO"

            Call SetStatus("Receiving game data...")
        Exit Sub

    ' :::::::::::::::::::::::::::::::::
    ' ::     News Recieved packet    ::
    ' :::::::::::::::::::::::::::::::::
        Case SPackets.Snews
            Call ParseNews(Trim$(parse(1)), Trim$(parse(2)), CInt(parse(3)), CInt(parse(4)), CInt(parse(5)))
        Exit Sub

    ' :::::::::::::::::::::::::::::::::::::::
    ' :: New character classes data packet ::
    ' :::::::::::::::::::::::::::::::::::::::
        Case SPackets.Snewcharclasses
            n = 1

            For i = 0 To MAX_CLASSES
                Class(i).Name = parse(n)
                Class(i).HP = CLng(parse(n + 1))
                Class(i).MP = CLng(parse(n + 2))
                Class(i).SP = CLng(parse(n + 3))
                Class(i).STR = CLng(parse(n + 4))
                Class(i).DEF = CLng(parse(n + 5))
                Class(i).speed = CLng(parse(n + 6))
                Class(i).MAGI = CLng(parse(n + 7))
                Class(i).MaleSprite = CLng(parse(n + 8))
                Class(i).FemaleSprite = CLng(parse(n + 9))
                Class(i).Locked = CLng(parse(n + 10))
                Class(i).desc = Trim$(parse(n + 11))

                n = n + 12
            Next i

            ' Used for if the player is creating a new character
            frmNewChar.Visible = True
            frmSendGetData.Visible = False

            frmNewChar.cmbClass.Clear
            
            For i = 0 To MAX_CLASSES
                If Class(i).Locked = 0 Then
                    frmNewChar.cmbClass.addItem Class(i).Name
                End If
            Next i
            
            frmNewChar.cmbClass.ListIndex = 0
            frmNewChar.lblClassDesc = Class(0).desc
            
            frmNewChar.cmbClass.Visible = True
            frmNewChar.lblClassDesc.Visible = True
            
            ' Changes stat display values depending on class
            frmNewChar.lblHP.Caption = CStr(Class(frmNewChar.cmbClass.ListIndex).HP)
            frmNewChar.lblMP.Caption = CStr(Class(frmNewChar.cmbClass.ListIndex).MP)
            frmNewChar.lblSP.Caption = CStr(Class(frmNewChar.cmbClass.ListIndex).SP)
            frmNewChar.lblSTR.Caption = CStr(Class(frmNewChar.cmbClass.ListIndex).STR)
            frmNewChar.lblDEF.Caption = CStr(Class(frmNewChar.cmbClass.ListIndex).DEF)
            frmNewChar.lblSpeed.Caption = CStr(Class(frmNewChar.cmbClass.ListIndex).speed)
            frmNewChar.lblMAGI.Caption = CStr(Class(frmNewChar.cmbClass.ListIndex).MAGI)

            frmNewChar.lblClassDesc.Caption = Class(0).desc
        Exit Sub

    ' :::::::::::::::::::::::::
    ' :: Classes data packet ::
    ' :::::::::::::::::::::::::
        Case SPackets.Sclassesdata
            n = 1

            For i = 0 To MAX_CLASSES
                Class(i).Name = parse(n)
                Class(i).HP = CLng(parse(n + 1))
                Class(i).MP = CLng(parse(n + 2))
                Class(i).SP = CLng(parse(n + 3))
                Class(i).STR = CLng(parse(n + 4))
                Class(i).DEF = CLng(parse(n + 5))
                Class(i).speed = CLng(parse(n + 6))
                Class(i).MAGI = CLng(parse(n + 7))
                Class(i).Locked = CLng(parse(n + 8))
                Class(i).desc = parse(n + 9)

                n = n + 10
            Next i
        Exit Sub

    ' ::::::::::::::::::::::::
    ' :: Desynch Fix packet ::
    ' ::::::::::::::::::::::::
        Case SPackets.Splayernewxy
            x = CLng(parse(1))
            y = CLng(parse(2))

            If Not GetPlayerX(MyIndex) = x Then Call SetPlayerX(MyIndex, x)
            If Not GetPlayerY(MyIndex) = y Then Call SetPlayerY(MyIndex, y)
        Exit Sub

    ' ::::::::::::::::::::::::::
    ' :: Players Online packet::
    ' ::::::::::::::::::::::::::
        Case SPackets.Stotalonline
            z = CLng(parse(1))
            
            If z <> 1 Then
                frmMainMenu.lblPlayers.Caption = "There are " & z & " players online."
            Else
                frmMainMenu.lblPlayers.Caption = "There is " & z & " player online."
            End If
        Exit Sub
    
    ' ::::::::::::::::::::
    ' :: In game packet ::
    ' ::::::::::::::::::::
        Case SPackets.Singame
            Player(MyIndex).Name = Trim$(parse(1))
        
            ' Check for TurnBasedConfig
            Dim TurnBasedConfig As String
            
            TurnBasedConfig = ReadINI("CONFIG", GetPlayerName(MyIndex) & " - TurnBased", App.Path & "\config.ini")
            
            If TurnBasedConfig = vbNullString Then
                Call WriteINI("CONFIG", GetPlayerName(MyIndex) & " - TurnBased", "0", App.Path & "\config.ini")
                TurnBasedConfig = "0"
            End If
        
            frmMirage.chkTurnBased.Value = Unchecked
        
            Call GameInit
            Call GameLoop
        Exit Sub

    ' :::::::::::::::::::::::::::::
    ' :: Player inventory packet ::
    ' :::::::::::::::::::::::::::::
        Case SPackets.Splayerinv
            n = 3
            z = CLng(parse(1))
            Player(z).MaxInv = CInt(parse(2))

            ReDim Player(z).NewInv(1 To Player(z).MaxInv) As PlayerInvRec

            For i = 1 To Player(z).MaxInv
                Call SetPlayerInvItemNum(z, i, CLng(parse(n)))
                Call SetPlayerInvItemValue(z, i, CLng(parse(n + 1)))
                Call SetPlayerInvItemAmmo(z, i, CLng(parse(n + 2)))
            
                n = n + 3
            Next i

            If z = MyIndex Then
                Call UpdateVisInv
            End If
        Exit Sub

    ' ::::::::::::::::::::::::::::::::::::
    ' :: Player inventory update packet ::
    ' ::::::::::::::::::::::::::::::::::::
        Case SPackets.Splayerinvupdate
            n = CLng(parse(1))
            z = CLng(parse(2))

            Call SetPlayerInvItemNum(z, n, CLng(parse(3)))
            Call SetPlayerInvItemValue(z, n, CLng(parse(4)))
            Call SetPlayerInvItemAmmo(z, n, CLng(parse(5)))
        
            If z = MyIndex Then
                ' Shows ammo of weapons and updates it
                If GetPlayerInvItemAmmo(z, n) > -1 Then
                    frmMirage.descName.Caption = Trim$(Item(GetPlayerInvItemNum(z, n)).Name) & " (Ammo: " & GetPlayerInvItemAmmo(z, n) & ")"
                End If
                
                Call UpdateVisInv
            End If
        Exit Sub
    
    ' ::::::::::::::::::::::::
    ' :: Player bank packet ::
    ' ::::::::::::::::::::::::
        Case SPackets.Splayerbank
            n = 1
            
            For i = 1 To MAX_BANK
                Call SetPlayerBankItemNum(MyIndex, i, CLng(parse(n)))
                Call SetPlayerBankItemValue(MyIndex, i, CLng(parse(n + 1)))
                Call SetPlayerBankItemAmmo(MyIndex, i, CLng(parse(n + 2)))

                n = n + 3
            Next i

            If frmBank.Visible = True Then
                Call UpdateBank
            End If
        Exit Sub

    ' :::::::::::::::::::::::::::::::
    ' :: Player bank update packet ::
    ' :::::::::::::::::::::::::::::::
        Case SPackets.Splayerbankupdate
            n = CLng(parse(1))

            Call SetPlayerBankItemNum(MyIndex, n, CLng(parse(2)))
            Call SetPlayerBankItemValue(MyIndex, n, CLng(parse(3)))
            Call SetPlayerBankItemAmmo(MyIndex, n, CLng(parse(4)))
        
            If frmBank.Visible = True Then
                Call UpdateBank
            End If
        Exit Sub

    ' :::::::::::::::::::::::::::::
    ' :: Player bank open packet ::
    ' :::::::::::::::::::::::::::::
        Case SPackets.Sopenbank
            Call frmBank.OpenBank
        Exit Sub

        Case SPackets.Sbankmsg
            frmBank.lblMsg.Caption = parse(1)
        Exit Sub

    ' ::::::::::::::::::::::::::::::::::
    ' :: Player worn equipment packet ::
    ' ::::::::::::::::::::::::::::::::::
        Case SPackets.Splayerworneq
            z = CLng(parse(1))
            
            If z <= 0 Then
                Exit Sub
            End If
            
            For i = 2 To 8
                Call SetPlayerEquipSlotNum(z, (i - 1), CLng(parse(i)))
            Next i

            If z = MyIndex Then
                Call UpdateVisInv
            End If
        Exit Sub
        
    ' ::::::::::::::::::::::::::::::::::::
    ' :: Player equipment update packet ::
    ' ::::::::::::::::::::::::::::::::::::
        Case SPackets.Splayerequipupdate
            n = CLng(parse(1))
            z = CLng(parse(2))

            Call SetPlayerEquipSlotNum(z, n, CLng(parse(3)))
            Call SetPlayerEquipSlotValue(z, n, CLng(parse(4)))
            Call SetPlayerEquipSlotAmmo(z, n, CLng(parse(5)))
        
            If z = MyIndex Then
                ' Shows ammo of weapons and updates it
                If GetPlayerEquipSlotAmmo(z, n) > -1 Then
                    frmMirage.descName.Caption = Trim$(Item(GetPlayerEquipSlotNum(z, n)).Name) & " (Ammo: " & GetPlayerEquipSlotAmmo(z, n) & ")"
                End If
            
                Call UpdateVisInv
            End If
        Exit Sub

    ' ::::::::::::::::::::::::::
    ' :: Player points packet ::
    ' ::::::::::::::::::::::::::
        Case SPackets.Splayerpoints
            Player(MyIndex).POINTS = CLng(parse(1))
            
            If Player(MyIndex).POINTS > 0 Then
                frmMirage.lblLevelUp.Visible = True
            Else
                frmMirage.lblLevelUp.Visible = False
            End If
        Exit Sub

    ' ::::::::::::::::::::::::::::::::::::
    ' :: Player stat points used packet ::
    ' ::::::::::::::::::::::::::::::::::::
        Case SPackets.Sstatpointused
            frmLevelUp.Visible = False
        
            Call MapMusic(Map(GetPlayerMap(MyIndex)).music)
        Exit Sub

    ' ::::::::::::::::::::::
    ' :: Player hp packet ::
    ' ::::::::::::::::::::::
        Case SPackets.Splayerhp
            n = CLng(parse(1))
        
            Player(n).MaxHp = CLng(parse(2))
            Call SetPlayerHP(n, CLng(parse(3)))
        Exit Sub

    ' ::::::::::::::::::::::
    ' :: Player mp packet ::
    ' ::::::::::::::::::::::
        Case SPackets.Splayermp
            Player(MyIndex).MaxMP = CLng(parse(1))
            Call SetPlayerMP(MyIndex, CLng(parse(2)))
        Exit Sub
    
    ' ::::::::::::::::::::::
    ' :: Player sp packet ::
    ' ::::::::::::::::::::::
        Case SPackets.Splayersp
            Player(MyIndex).MaxSP = CLng(parse(1))
            Call SetPlayerSP(MyIndex, CLng(parse(2)))
            
            frmMirage.lblSP.Caption = Int((GetPlayerSP(MyIndex) / GetPlayerMaxSP(MyIndex)) * 100) & "%"
        Exit Sub
    
        Case SPackets.Splayerexp
            Player(MyIndex).Exp = CLng(parse(1))
            Player(MyIndex).NextLvlExp = CLng(parse(2))
        Exit Sub

    ' speech bubble packet
        Case SPackets.Smapmsg2
            n = CLng(parse(2))
        
            Bubble(n).Text = parse(1)
            Bubble(n).Created = GetTickCount()
        Exit Sub

    ' :::::::::::::::::::::::::
    ' :: Player Stats Packet ::
    ' :::::::::::::::::::::::::
        Case SPackets.Splayerstatspacket
            Dim SubDef As Long, SubMagi As Long, SubSpeed As Long, SubStr As Long
        
            Player(MyIndex).STR = CLng(parse(1))
            Player(MyIndex).DEF = CLng(parse(2))
            Player(MyIndex).speed = CLng(parse(3))
            Player(MyIndex).MAGI = CLng(parse(4))
            
            For i = 1 To 7
                p = GetPlayerEquipSlotNum(MyIndex, i)
                
                If p > 0 Then
                    SubStr = SubStr + Item(p).AddSTR
                    SubDef = SubDef + Item(p).AddDef
                    SubMagi = SubMagi + Item(p).AddMAGI
                    SubSpeed = SubSpeed + Item(p).AddSpeed
                End If
            Next i

            If SubStr > 0 Then
                frmMirage.lblSTR.Caption = parse(1) - SubStr & " (+" & SubStr & ")"
            ElseIf SubStr < 0 Then
                frmMirage.lblSTR.Caption = parse(1) - SubStr & " (" & SubStr & ")"
            Else
                frmMirage.lblSTR.Caption = parse(1)
            End If
            
            If SubDef > 0 Then
                frmMirage.lblDEF.Caption = parse(2) - SubDef & " (+" & SubDef & ")"
            ElseIf SubDef < 0 Then
                frmMirage.lblDEF.Caption = parse(2) - SubDef & " (" & SubDef & ")"
            Else
                frmMirage.lblDEF.Caption = parse(2)
            End If
            
            If SubSpeed > 0 Then
                frmMirage.lblSpeed.Caption = parse(3) - SubSpeed & " (+" & SubSpeed & ")"
            ElseIf SubSpeed < 0 Then
                frmMirage.lblSpeed.Caption = parse(3) - SubSpeed & " (" & SubSpeed & ")"
            Else
                frmMirage.lblSpeed.Caption = parse(3)
            End If
            
            If SubMagi > 0 Then
                frmMirage.lblMAGI.Caption = parse(4) - SubMagi & " (+" & SubMagi & ")"
            ElseIf SubMagi < 0 Then
                frmMirage.lblMAGI.Caption = parse(4) - SubMagi & " (" & SubMagi & ")"
            Else
                frmMirage.lblMAGI.Caption = parse(4)
            End If
        
            Call SetPlayerNextLvlExp(MyIndex, CLng(parse(5)))
            Call SetPlayerExp(MyIndex, CLng(parse(6)))
        
            frmMirage.lblLevel.Caption = parse(7)
            Player(MyIndex).Level = CLng(parse(7))
            Player(MyIndex).CritHitChance = CDbl(parse(8))
            Player(MyIndex).BlockChance = CDbl(parse(9))
        Exit Sub

    ' ::::::::::::::::::::::::::
    ' :: General MsgBox Packet::
    ' ::::::::::::::::::::::::::
        Case SPackets.Smsgbox
            Call MsgBox(parse(2), 0, parse(1))
        Exit Sub
    
    ' :::::::::::::::::
    ' :: Quest Packet::
    ' :::::::::::::::::
        Case SPackets.Sfavor
            Call frmQuest.FavorStart(parse(1), parse(2), parse(3), parse(4))
        Exit Sub
    
    ' :::::::::::::::::::::
    ' :: Card Shop Packet::
    ' :::::::::::::::::::::
        Case SPackets.Scardshop
            frmCard.lstCards.Clear
        
            For i = 94 To MAX_ITEMS
                Msg = Trim$(parse(i - 93))
            
                If Msg <> vbNullString Then
                    frmCard.lstCards.addItem Msg
                Else
                    If Item(i).Type = ITEM_TYPE_CARD Then
                        frmCard.lstCards.addItem "<Empty Card Slot>"
                    End If
                End If
            Next i
        
            Call frmCard.Show(vbModal, frmMirage)
        Exit Sub
    
    ' ::::::::::::::::::::::::::::
    ' :: Update Card Shop Packet::
    ' ::::::::::::::::::::::::::::
        Case SPackets.Supdatecardshop
            Call frmCard.UpdateInfo(CInt(parse(1)), CInt(parse(2)), CInt(parse(3)), CInt(parse(4)), CInt(parse(5)), Trim$(parse(6)), CLng(parse(7)))
        Exit Sub
    
    ' ::::::::::::::::::::::::
    ' :: Player data packet ::
    ' ::::::::::::::::::::::::
        Case SPackets.Splayerdata
            i = CLng(parse(1))
            
            Call SetPlayerName(i, parse(2))
            Call SetPlayerSprite(i, CLng(parse(3)))
            Call SetPlayerMap(i, CLng(parse(4)))
            Call SetPlayerDir(i, CLng(parse(5)))
            Call SetPlayerAccess(i, CLng(parse(6)))
            Call SetPlayerPK(i, CLng(parse(7)))
            Call SetPlayerGuild(i, parse(8))
            Call SetPlayerGuildAccess(i, CLng(parse(9)))
            Call SetPlayerClass(i, CLng(parse(10)))
            Call SetPlayerHead(i, 0)
            Call SetPlayerBody(i, 0)
            Call SetPlayerLeg(i, 0)
            Call SetPlayerLevel(i, CLng(parse(11)))

            ' Check if the player is the client player, and if so reset directions
            If i = MyIndex Then
                DirUp = False
                DirDown = False
                DirLeft = False
                DirRight = False
            End If
        Exit Sub
    
    ' ::::::::::::::::::::::::::::::
    ' :: Players In Battle Packet ::
    ' ::::::::::::::::::::::::::::::
        Case SPackets.Smapplayersinbattle
            z = CInt(parse(1))
            n = 2
            a = n + 1
        
            Do While n <= (z * 2)
                Player(CLng(parse(n))).InBattle = CBool(parse(a))
                
                n = a + 1
                a = n + 1
            Loop
        Exit Sub
    
    ' Leaving map packet
        Case SPackets.Sleave
            Call SetPlayerMap(CLng(parse(1)), 0)
        Exit Sub
        
    ' Exiting game packet
        Case SPackets.Sleft
            Call ClearPlayer(CLng(parse(1)))
        Exit Sub

    ' ::::::::::::::::::::::::::::
    ' :: Welcome Message Packet ::
    ' ::::::::::::::::::::::::::::
        Case SPackets.Swelcomemsg
            Call PlaySound("yi_messageBoxAppear.wav")
            
            Call frmWelcome.ShowWelcomeMsg(parse(1), parse(2))
        Exit Sub
    
    ' ::::::::::::::::::::::::::
    ' :: Friend Status Packet ::
    ' ::::::::::::::::::::::::::
        Case SPackets.Sfriendstatus
            i = CLng(parse(1))
            Q = parse(2)
            
            If i = 0 Then
                frmMirage.lblFriend(Q).ForeColor = &HFF&
            Else
                frmMirage.lblFriend(Q).ForeColor = &H8080&
            End If
       Exit Sub
     
    ' :::::::::::::::::::::::::
    ' :: Friends List Packet ::
    ' :::::::::::::::::::::::::
        Case SPackets.Scaption
            F = CInt(parse(2))
            
            frmMirage.lblFriend(F - 1).Caption = Trim$(parse(1))
        Exit Sub
    
    ' ::::::::::::::::::::::::::
    ' :: Player Level Packet  ::
    ' ::::::::::::::::::::::::::
        Case SPackets.Splayerlevel
            n = CLng(parse(1))
            Player(n).Level = CLng(parse(2))
        Exit Sub

    ' ::::::::::::::::::::::::::::
    ' :: Player movement packet ::
    ' ::::::::::::::::::::::::::::
        Case SPackets.Splayermove
            i = CLng(parse(1))
            x = CLng(parse(2))
            y = CLng(parse(3))
            Dir = CLng(parse(4))
            n = CLng(parse(5))

            If Dir < DIR_UP Or Dir > DIR_RIGHT Then
                Exit Sub
            End If
            
            Call SetPlayerX(i, x)
            Call SetPlayerY(i, y)
            Call SetPlayerDir(i, Dir)

            Player(i).xOffset = 0
            Player(i).yOffset = 0
            Player(i).Moving = n
        
            Select Case GetPlayerDir(i)
                Case DIR_UP
                    Player(i).yOffset = PIC_Y
                Case DIR_DOWN
                    Player(i).yOffset = PIC_Y * -1
                Case DIR_LEFT
                    Player(i).xOffset = PIC_X
                Case DIR_RIGHT
                    Player(i).xOffset = PIC_X * -1
            End Select
        Exit Sub

    ' :::::::::::::::::::::::::
    ' :: Npc movement packet ::
    ' :::::::::::::::::::::::::
        Case SPackets.Snpcmove
            i = CLng(parse(1))
            x = CByte(parse(2))
            y = CByte(parse(3))
            Dir = CByte(parse(4))
            n = CLng(parse(5))

            MapNpc(i).x = x
            MapNpc(i).y = y
            MapNpc(i).Dir = Dir
            MapNpc(i).xOffset = 0
            MapNpc(i).yOffset = 0
            MapNpc(i).Moving = 1

            If n <> 1 Then
                Select Case MapNpc(i).Dir
                    Case DIR_UP
                        MapNpc(i).yOffset = PIC_Y * (n - 1)
                    Case DIR_DOWN
                        MapNpc(i).yOffset = PIC_Y * -n
                    Case DIR_LEFT
                        MapNpc(i).xOffset = PIC_X * Val(n - 1)
                    Case DIR_RIGHT
                        MapNpc(i).xOffset = PIC_X * -n
                End Select
            Else
                Select Case MapNpc(i).Dir
                    Case DIR_UP
                        MapNpc(i).yOffset = PIC_Y
                    Case DIR_DOWN
                        MapNpc(i).yOffset = PIC_Y * -1
                    Case DIR_LEFT
                        MapNpc(i).xOffset = PIC_X
                    Case DIR_RIGHT
                        MapNpc(i).xOffset = PIC_X * -1
                End Select
            End If
        Exit Sub

    ' :::::::::::::::::::::::::::::
    ' :: Player direction packet ::
    ' :::::::::::::::::::::::::::::
        Case SPackets.Splayerdir
            i = CLng(parse(1))
            Dir = CLng(parse(2))

            If Dir < DIR_UP Or Dir > DIR_RIGHT Then
                Exit Sub
            End If

            Call SetPlayerDir(i, Dir)

            Player(i).xOffset = 0
            Player(i).yOffset = 0
            Player(i).MovingH = 0
            Player(i).MovingV = 0
            Player(i).Moving = 0
        Exit Sub

    ' ::::::::::::::::::::::::::
    ' :: NPC direction packet ::
    ' ::::::::::::::::::::::::::
        Case SPackets.Snpcdir
            i = CLng(parse(1))
            Dir = CLng(parse(2))
            
            MapNpc(i).Dir = Dir
            MapNpc(i).xOffset = 0
            MapNpc(i).yOffset = 0
            MapNpc(i).Moving = 0
        Exit Sub

    ' :::::::::::::::::::::::::::::::
    ' :: Player XY location packet ::
    ' :::::::::::::::::::::::::::::::
        Case SPackets.Splayerxy
            i = CLng(parse(1))
            x = CLng(parse(2))
            y = CLng(parse(3))

            Call SetPlayerX(i, x)
            Call SetPlayerY(i, y)

            ' Make sure they aren't walking
            Player(i).Moving = 0
            Player(i).xOffset = 0
            Player(i).yOffset = 0
        Exit Sub

    ' ::::::::::::::::::::::::::
    ' :: Player attack packet ::
    ' ::::::::::::::::::::::::::
        Case SPackets.Sattack
            i = CLng(parse(1))

            ' Set player to attacking
            Player(i).Attacking = 1
            Player(i).AttackTimer = GetTickCount
        Exit Sub

    ' :::::::::::::::::::::::
    ' :: NPC attack packet ::
    ' :::::::::::::::::::::::
        Case SPackets.Snpcattack
            i = CLng(parse(1))

            ' Set player to attacking
            MapNpc(i).Attacking = 1
            MapNpc(i).AttackTimer = GetTickCount
        Exit Sub

    ' ::::::::::::::::::::::::::
    ' :: Check for map packet ::
    ' ::::::::::::::::::::::::::
        Case SPackets.Scheckformap
            GettingMap = True
    
            ' Erase all players except self
            For i = 1 To MAX_PLAYERS
                If i <> MyIndex Then
                    Call SetPlayerMap(i, 0)
                End If
            Next i

            ' Erase all temporary tile values
            Call ClearTempTile

            ' Get map num
            x = CLng(parse(1))
            
            ' Get revision
            y = CInt(parse(2))
        
            ' Close map editor if player leaves current map
            If InEditor Then
                ScreenMode = 0
                GridMode = 0
                InEditor = False
                Unload frmMapEditor
                frmMapEditor.MousePointer = 1
                frmMirage.MousePointer = 1
            End If
        
            If FileExists("maps\map" & x & ".dat") Then
                ' Check to see if the revisions match
                If GetMapRevision(x) = y Then
                    ' Load the map
                    Call LoadMap(x)
                    
                    MapNeed = "no"
                Else
                    MapNeed = "yes"
                End If
            Else
                MapNeed = "yes"
            End If

            ' Either the revisions didn't match or we dont have the map, so we need it
            Call SendData(CPackets.Cneedmap & SEP_CHAR & MapNeed & END_CHAR)
            
            ' When switching maps, clear all scripted spell animations
            For i = 1 To MAX_SCRIPTSPELLS
                ' Don't clear out the Drill or Hammer Barrage animations
                If ScriptSpell(i).SpellNum <> 28 And ScriptSpell(i).SpellNum <> 29 Then
                    If ScriptSpell(i).CastedSpell = Yes Then
                        ScriptSpell(i).CastedSpell = No
                    End If
                End If
            Next i
        Exit Sub

    ' :::::::::::::::::::::
    ' :: Map data packet ::
    ' :::::::::::::::::::::
        Case SPackets.Smapdata
            n = 1
            a = CLng(parse(1))
            
            With Map(a)
                .Name = Trim$(parse(n + 1))
                .Revision = CInt(parse(n + 2))
                .Moral = CByte(parse(n + 3))
                .Up = CInt(parse(n + 4))
                .Down = CInt(parse(n + 5))
                .Left = CInt(parse(n + 6))
                .Right = CInt(parse(n + 7))
                .music = parse(n + 8)
                .BootMap = CInt(parse(n + 9))
                .BootX = CByte(parse(n + 10))
                .BootY = CByte(parse(n + 11))
                .Indoors = CByte(parse(n + 12))
                .Weather = CInt(parse(n + 13))
                
                n = n + 14
            End With
            
            For y = 0 To MAX_MAPY
                For x = 0 To MAX_MAPX
                    With Map(a).Tile(x, y)
                        .Ground = CLng(parse(n))
                        .Mask = CLng(parse(n + 1))
                        .Anim = CLng(parse(n + 2))
                        .Mask2 = CLng(parse(n + 3))
                        .M2Anim = CLng(parse(n + 4))
                        .Fringe = CLng(parse(n + 5))
                        .FAnim = CLng(parse(n + 6))
                        .Fringe2 = CLng(parse(n + 7))
                        .F2Anim = CLng(parse(n + 8))
                        .Type = CByte(parse(n + 9))
                        .Data1 = CLng(parse(n + 10))
                        .Data2 = CLng(parse(n + 11))
                        .Data3 = CLng(parse(n + 12))
                        .String1 = Trim$(parse(n + 13))
                        .String2 = Trim$(parse(n + 14))
                        .String3 = Trim$(parse(n + 15))
                        .light = CLng(parse(n + 16))
                        .GroundSet = CByte(parse(n + 17))
                        .MaskSet = CByte(parse(n + 18))
                        .AnimSet = CByte(parse(n + 19))
                        .Mask2Set = CByte(parse(n + 20))
                        .M2AnimSet = CByte(parse(n + 21))
                        .FringeSet = CByte(parse(n + 22))
                        .FAnimSet = CByte(parse(n + 23))
                        .Fringe2Set = CByte(parse(n + 24))
                        .F2AnimSet = CByte(parse(n + 25))
                
                        n = n + 26
                    End With
                Next x
            Next y

            For x = 1 To 15
                Map(a).Npc(x) = CInt(parse(n))
                Map(a).SpawnX(x) = CByte(parse(n + 1))
                Map(a).SpawnY(x) = CByte(parse(n + 2))
                
                n = n + 3
            Next x
        
            For y = 0 To MAX_MAPY
                For x = 0 To MAX_MAPX
                    With QuestionBlock(a, x, y)
                        .Item1 = CLng(parse(n))
                        .Item2 = CLng(parse(n + 1))
                        .Item3 = CLng(parse(n + 2))
                        .Item4 = CLng(parse(n + 3))
                        .Item5 = CLng(parse(n + 4))
                        .Item6 = CLng(parse(n + 5))
                        .Chance1 = CLng(parse(n + 6))
                        .Chance2 = CLng(parse(n + 7))
                        .Chance3 = CLng(parse(n + 8))
                        .Chance4 = CLng(parse(n + 9))
                        .Chance5 = CLng(parse(n + 10))
                        .Chance6 = CLng(parse(n + 11))
                        .Value1 = CLng(parse(n + 12))
                        .Value2 = CLng(parse(n + 13))
                        .Value3 = CLng(parse(n + 14))
                        .Value4 = CLng(parse(n + 15))
                        .Value5 = CLng(parse(n + 16))
                        .Value6 = CLng(parse(n + 17))
                
                        n = n + 18
                    End With
                Next x
            Next y
        
            ' Save the map
            Call SaveLocalMap(a)

            ' Check if we get a map from someone else and if we were editing a map cancel it out
            If InEditor Then
                InEditor = False
                frmMapEditor.Visible = False
                frmMirage.Show

                If frmMapWarp.Visible Then
                    Unload frmMapWarp
                End If

                If frmMapProperties.Visible Then
                    Unload frmMapProperties
                End If
            End If
        Exit Sub
    ' :::::::::::::::::::::::::::
    ' :: Map items data packet ::
    ' :::::::::::::::::::::::::::
        Case SPackets.Smapitemdata
            n = 1

            For i = 1 To MAX_MAP_ITEMS
                SaveMapItem(i).num = CLng(parse(n))
                SaveMapItem(i).Value = CLng(parse(n + 1))
                SaveMapItem(i).x = CByte(parse(n + 2))
                SaveMapItem(i).y = CByte(parse(n + 3))
                SaveMapItem(i).Ammo = CLng(parse(n + 4))
                MapItem(i) = SaveMapItem(i)

                n = n + 5
            Next i
        Exit Sub

    ' :::::::::::::::::::::::::
    ' :: Map npc data packet ::
    ' :::::::::::::::::::::::::
        Case SPackets.Smapnpcdata
            n = 1

            For i = 1 To MAX_MAP_NPCS
                SaveMapNpc(i).num = CLng(parse(n))
                SaveMapNpc(i).x = CByte(parse(n + 1))
                SaveMapNpc(i).y = CByte(parse(n + 2))
                SaveMapNpc(i).Dir = CByte(parse(n + 3))
                SaveMapNpc(i).Target = CLng(parse(n + 4))
                SaveMapNpc(i).InBattle = CBool(parse(n + 5))
                MapNpc(i) = SaveMapNpc(i)

                n = n + 6
            Next i
        Exit Sub
    
    ' :::::::::::::::::::::::::::::::
    ' :: Map send completed packet ::
    ' :::::::::::::::::::::::::::::::
        Case SPackets.Smapdone
            GettingMap = False

            ' Play music
            If Trim$(Map(GetPlayerMap(MyIndex)).music) <> "None" Then
                Call MapMusic(Map(GetPlayerMap(MyIndex)).music)
            End If

            If GameWeather = WEATHER_RAINING Then
                Call PlayBGS("rain.wav")
            ElseIf GameWeather = WEATHER_THUNDER Then
                Call PlayBGS("thunder.wav")
            End If
        
            frmMirage.Caption = "Super Mario Bros. Online - " & Trim$(Map(GetPlayerMap(MyIndex)).Name)
        Exit Sub
    
    ' :::::::::::::::::::::::::::::
    ' :: Turn-Based Timer packet ::
    ' :::::::::::::::::::::::::::::
        Case SPackets.Sturnbasedtime
            TurnBasedTimeToWait = CLng(parse(1)) * 1000
            PlayerTurn = CInt(parse(2))
            TurnBasedTime = GetTickCount()
            TurnBasedTimer = True
        
            If PlayerTurn = 0 Then
                IsPlayerTurn = True
            Else
                IsPlayerTurn = False
            End If
        Exit Sub
    
    ' ::::::::::::::::::::
    ' :: Social packets ::
    ' ::::::::::::::::::::
        Case SPackets.Sglobalmsg, SPackets.Splayermsg, SPackets.Sothermsg, SPackets.Smapmsg, SPackets.Sadminmsg
            Call AddText(parse(1), CInt(parse(2)))
        Exit Sub

    ' :::::::::::::::::::::::
    ' :: Item spawn packet ::
    ' :::::::::::::::::::::::
        Case SPackets.Sspawnitem
            n = CLng(parse(1))

            MapItem(n).num = CLng(parse(2))
            MapItem(n).Value = CLng(parse(3))
            MapItem(n).x = CByte(parse(4))
            MapItem(n).y = CByte(parse(5))
            MapItem(n).Ammo = CLng(parse(6))
        Exit Sub

    ' ::::::::::::::::::::::::
    ' :: Item editor packet ::
    ' ::::::::::::::::::::::::
        Case SPackets.Sitemeditor
            InItemsEditor = True

            Call frmIndex.Show(vbModeless, frmMirage)
            frmIndex.lstIndex.Clear

            ' Add the names
            For i = 1 To MAX_ITEMS
                frmIndex.lstIndex.addItem i & ": " & Trim$(Item(i).Name)
            Next i

            frmIndex.lstIndex.ListIndex = 0
        Exit Sub

    ' ::::::::::::::::::::::::
    ' :: Update item packet ::
    ' ::::::::::::::::::::::::
        Case SPackets.Supdateitem
            n = CLng(parse(1))

            ' Update the item
            Item(n).Name = Trim$(parse(2))
            Item(n).Pic = CLng(parse(3))
            Item(n).Type = CByte(parse(4))
            Item(n).Data1 = CLng(parse(5))
            Item(n).Data2 = CLng(parse(6))
            Item(n).Data3 = CLng(parse(7))
            Item(n).StrReq = CLng(parse(8))
            Item(n).DefReq = CLng(parse(9))
            Item(n).SpeedReq = CLng(parse(10))
            Item(n).MagicReq = CLng(parse(11))
            Item(n).ClassReq = CLng(parse(12))
            Item(n).AccessReq = CByte(parse(13))

            Item(n).AddHP = CLng(parse(14))
            Item(n).AddMP = CLng(parse(15))
            Item(n).AddSP = CLng(parse(16))
            Item(n).AddSTR = CLng(parse(17))
            Item(n).AddDef = CLng(parse(18))
            Item(n).AddMAGI = CLng(parse(19))
            Item(n).AddSpeed = CLng(parse(20))
            Item(n).AddEXP = CLng(parse(21))
            Item(n).desc = Trim$(parse(22))
            Item(n).AttackSpeed = CLng(parse(23))
            Item(n).Price = CLng(parse(24))
            Item(n).Stackable = CLng(parse(25))
            Item(n).Bound = CLng(parse(26))
            Item(n).LevelReq = CLng(parse(27))
            Item(n).HPReq = CLng(parse(28))
            Item(n).FPReq = CLng(parse(29))
            Item(n).Ammo = CLng(parse(30))
            Item(n).AddCritChance = CDbl(parse(31))
            Item(n).AddBlockChance = CDbl(parse(32))
            Item(n).Cookable = CBool(parse(33))
        Exit Sub

    ' ::::::::::::::::::::::
    ' :: Edit item packet :: <- Used for item editor admins only
    ' ::::::::::::::::::::::
        Case SPackets.Sedititem
            ' Update the item
            Item(CLng(parse(1))).Name = Trim$(parse(2))
            
            ' Initialize the item editor
            Call ItemEditorInit
        Exit Sub

    ' ::::::::::::::::::::::
    ' :: Npc spawn packet ::
    ' ::::::::::::::::::::::
        Case SPackets.Sspawnnpc
            n = CLng(parse(1))

            MapNpc(n).num = CLng(parse(2))
            MapNpc(n).x = CByte(parse(3))
            MapNpc(n).y = CByte(parse(4))
            MapNpc(n).Dir = CByte(parse(5))
            MapNpc(n).Big = CByte(parse(6))

            ' Client use only
            MapNpc(n).xOffset = 0
            MapNpc(n).yOffset = 0
            MapNpc(n).Moving = 0
        Exit Sub

    ' :::::::::::::::::::::
    ' :: Npc dead packet ::
    ' :::::::::::::::::::::
        Case SPackets.Snpcdead
            n = CLng(parse(1))

            MapNpc(n).num = 0
            MapNpc(n).x = 0
            MapNpc(n).y = 0
            MapNpc(n).Dir = 0

            ' Client use only
            MapNpc(n).xOffset = 0
            MapNpc(n).yOffset = 0
            MapNpc(n).Moving = 0
        Exit Sub

    ' :::::::::::::::::::::::
    ' :: Npc editor packet ::
    ' :::::::::::::::::::::::
        Case SPackets.Snpceditor
            InNpcEditor = True
        
            Call frmIndex.Show(vbModeless, frmMirage)
            frmIndex.lstIndex.Clear

            ' Add the names
            For i = 1 To MAX_NPCS
                frmIndex.lstIndex.addItem i & ": " & Trim$(Npc(i).Name)
            Next i

            frmIndex.lstIndex.ListIndex = 0
        Exit Sub
    
    ' :::::::::::::::::::::
    ' :: npc talk packet ::
    ' :::::::::::::::::::::
        Case SPackets.Snpctalk
            z = CLng(parse(1))
            Msg = Trim$(parse(2))
            
            If z = 224 Then ' Chef Bean B.
                If Mid$(Msg, 1, 6) = "No, no" Then
                    IsChefBeanB = True
                End If
            End If
            
            Call frmNpcTalk.NpcTalk(z, Msg, Trim$(parse(3)))
        Exit Sub
    
    ' :::::::::::::::::::::::::::::::
    ' :: npc talk yes or no packet ::
    ' :::::::::::::::::::::::::::::::
        Case SPackets.Snpctalkyesno
            Call frmNpcTalkYesNo.NpcTalk(CLng(parse(1)), parse(2), parse(3), parse(4))
        Exit Sub
    
    ' :::::::::::::::::::::::
    ' :: Update npc packet ::
    ' :::::::::::::::::::::::
        Case SPackets.Supdatenpc
            n = CLng(parse(1))

            ' Update the npc
            Npc(n).Name = Trim$(parse(2))
            Npc(n).AttackSay = Trim$(parse(3))
            Npc(n).Sprite = CLng(parse(4))
            Npc(n).SpawnSecs = CLng(parse(5))
            Npc(n).Behavior = CByte(parse(6))
            Npc(n).Range = CByte(parse(7))
            Npc(n).STR = CLng(parse(8))
            Npc(n).DEF = CLng(parse(9))
            Npc(n).speed = CLng(parse(10))
            Npc(n).MAGI = CLng(parse(11))
            Npc(n).Big = CLng(parse(12))
            Npc(n).MaxHp = CLng(parse(13))
            Npc(n).Exp = CLng(parse(14))
            Npc(n).SpawnTime = CLng(parse(15))
            Npc(n).Element = CLng(parse(16))
            Npc(n).SpriteSize = CLng(parse(17))
            z = 18
        
            For i = 1 To MAX_NPC_DROPS
                Npc(n).ItemNPC(i).chance = CLng(parse(z))
                Npc(n).ItemNPC(i).ItemNum = CLng(parse(z + 1))
                Npc(n).ItemNPC(i).ItemValue = CLng(parse(z + 2))
                z = z + 3
            Next i

            Npc(n).AttackSay2 = Trim$(parse(48))
            Npc(n).Level = CLng(parse(49))
        Exit Sub

    ' :::::::::::::::::::::
    ' :: Edit npc packet ::
    ' :::::::::::::::::::::
        Case SPackets.Seditnpc
            Npc(CLng(parse(1))).Name = Trim$(parse(2))
            
            ' Initialize the npc editor
            Call NpcEditorInit
        Exit Sub
    
    ' ::::::::::::::::::::::::::
    ' :: Recipe editor packet ::
    ' ::::::::::::::::::::::::::
        Case SPackets.Srecipeeditor
            InRecipeEditor = True

            Call frmIndex.Show(vbModeless, frmMirage)
            frmIndex.lstIndex.Clear
        
            ' Add the names
            For i = 1 To MAX_RECIPES
                frmIndex.lstIndex.addItem i & ": " & Trim$(Recipe(i).Name)
            Next i

            frmIndex.lstIndex.ListIndex = 0
        Exit Sub
    
    ' ::::::::::::::::::::::::::
    ' :: Update recipe packet ::
    ' ::::::::::::::::::::::::::
        Case SPackets.Supdaterecipe
            n = CLng(parse(1))

            ' Update the recipe
            Recipe(n).Ingredient1 = CLng(parse(2))
            Recipe(n).Ingredient2 = CLng(parse(3))
            Recipe(n).ResultItem = CLng(parse(4))
            Recipe(n).Name = Trim$(parse(5))
        Exit Sub
    
    ' :::::::::::::::::::::::::
    ' :: Edit recipe packet ::: <- Used for recipe editor admins only
    ' :::::::::::::::::::::::::
        Case SPackets.Seditrecipe
            n = CLng(parse(1))

            ' Update the recipe
            Recipe(n).Ingredient1 = CLng(parse(2))
            Recipe(n).Ingredient2 = CLng(parse(3))
            Recipe(n).ResultItem = CLng(parse(4))
            Recipe(n).Name = Trim$(parse(5))
        
            ' Initialize the recipe editor
            Call RecipeEditorInit
        Exit Sub
    
    ' :::::::::::::::::::::::::::
    ' :: Begin Cooking packet :::
    ' :::::::::::::::::::::::::::
        Case SPackets.Scookitem
            Call ShowCookForm(CLng(parse(1)))
        Exit Sub
    
    ' :::::::::::::::::::::
    ' :: Cooking packet :::
    ' :::::::::::::::::::::
        Case SPackets.Scooking
            RecipeNumber = CLng(parse(1))
            CookingTimer = True
            CookingTime = GetTickCount()
        Exit Sub
        
    ' :::::::::::::::::::::::
    ' :: Recipe Log Packet ::
    ' :::::::::::::::::::::::
        Case SPackets.Srecipelog
            frmRecipe.lstRecipes.Clear
        
            For i = 1 To MAX_RECIPES
                If Recipe(i).ResultItem > 0 Then
                    Msg = Trim$(parse(i))
                
                    If Msg <> vbNullString Then
                        frmRecipe.lstRecipes.addItem Msg
                    Else
                        frmRecipe.lstRecipes.addItem "<Empty Slot>"
                    End If
                End If
            Next i
        
            Call frmRecipe.Show(vbModal, frmMirage)
        Exit Sub
    
    ' ::::::::::::::::::::
    ' :: Map key packet ::
    ' ::::::::::::::::::::
        Case SPackets.Smapkey
            x = CLng(parse(1))
            y = CLng(parse(2))
            n = CLng(parse(3))

            TempTile(x, y).DoorOpen = n
        Exit Sub

    ' :::::::::::::::::::::
    ' :: Edit map packet ::
    ' :::::::::::::::::::::
        Case SPackets.Seditmap
            Call EditorInit
        Exit Sub

    ' ::::::::::::::::::::::::
    ' :: Shop editor packet ::
    ' ::::::::::::::::::::::::
        Case SPackets.Sshopeditor
            InShopEditor = True

            Call frmIndex.Show(vbModeless, frmMirage)
            frmIndex.lstIndex.Clear

            ' Add the names
            For i = 1 To MAX_SHOPS
                frmIndex.lstIndex.addItem i & ": " & Trim$(Shop(i).Name)
            Next i

            frmIndex.lstIndex.ListIndex = 0
        Exit Sub

    ' ::::::::::::::::::::::::
    ' :: Update shop packet ::
    ' ::::::::::::::::::::::::
        Case SPackets.Supdateshop
            n = CLng(parse(1))

            ' Update the shop name
            Shop(n).Name = Trim$(parse(2))
            Shop(n).FixesItems = 0
            Shop(n).BuysItems = CByte(parse(3))
            Shop(n).ShowInfo = CByte(parse(4))
            Shop(n).currencyItem = CInt(parse(5))

            a = 6
            
            ' Get shop items
            For i = 1 To MAX_SHOP_ITEMS
                Shop(n).ShopItem(i).ItemNum = CLng(parse(a))
                Shop(n).ShopItem(i).Amount = CDbl(parse(a + 1))
                Shop(n).ShopItem(i).Price = CDbl(parse(a + 2))
                Shop(n).ShopItem(i).currencyItem = CInt(parse(a + 3))
                
                a = a + 4
            Next i
        Exit Sub

    ' ::::::::::::::::::::::
    ' :: Edit shop packet :: <- Used for shop editor admins only
    ' ::::::::::::::::::::::
        Case SPackets.Seditshop
            ShopNum = CLng(parse(1))

            ' Update the shop
            Shop(ShopNum).Name = Trim$(parse(2))

            ' Initialize the shop editor
            Call ShopEditorInit
        Exit Sub

    ' :::::::::::::::::::::::::
    ' :: Spell editor packet ::
    ' :::::::::::::::::::::::::
        Case SPackets.Sspelleditor
            InSpellEditor = True

            Call frmIndex.Show(vbModeless, frmMirage)
            frmIndex.lstIndex.Clear

            ' Add the names
            For i = 1 To MAX_SPELLS
                frmIndex.lstIndex.addItem i & ": " & Trim$(Spell(i).Name)
            Next i

            frmIndex.lstIndex.ListIndex = 0
        Exit Sub

    ' ::::::::::::::::::::::::
    ' :: Update spell packet ::
    ' ::::::::::::::::::::::::
        Case SPackets.Supdatespell
            n = CLng(parse(1))
        
            ' Update the spell
            Spell(n).Name = Trim$(parse(2))
            Spell(n).ClassReq = CLng(parse(3))
            Spell(n).LevelReq = CLng(parse(4))
            Spell(n).Type = CLng(parse(5))
            Spell(n).Data1 = CLng(parse(6))
            Spell(n).Data2 = CLng(parse(7))
            Spell(n).Data3 = CLng(parse(8))
            Spell(n).MPCost = CLng(parse(9))
            Spell(n).Sound = Trim$(parse(10))
            Spell(n).Range = CByte(parse(11))
            Spell(n).SpellAnim = CLng(parse(12))
            Spell(n).SpellTime = CLng(parse(13))
            Spell(n).SpellDone = CLng(parse(14))
            Spell(n).AE = CLng(parse(15))
            Spell(n).Big = CLng(parse(16))
            Spell(n).Element = CLng(parse(17))
            Spell(n).Stat = CInt(parse(18))
            Spell(n).StatTime = CLng(parse(19))
            Spell(n).Multiplier = CDbl(parse(20))
            Spell(n).PassiveStat = CInt(parse(21))
            Spell(n).PassiveStatChange = CInt(parse(22))
            Spell(n).UsePassiveStat = CBool(parse(23))
            Spell(n).SelfSpell = CBool(parse(24))
        Exit Sub

    ' :::::::::::::::::::::::
    ' :: Edit spell packet :: <- Used for spell editor admins only
    ' :::::::::::::::::::::::
        Case SPackets.Seditspell
            n = CLng(parse(1))
            
            ' Update the spell name
            Spell(n).Name = Trim$(parse(2))

            ' Initialize the spell editor
            Call SpellEditorInit
        Exit Sub

    ' ::::::::::::::::::
    ' :: Trade packet ::
    ' ::::::::::::::::::
        Case SPackets.Sgoshop
            IsShopping = True
        
            ShopNum = CLng(parse(1))
            
            ' Show the shop
            Call GoShop(ShopNum)
        Exit Sub

    ' :::::::::::::::::::
    ' :: Spells packet ::
    ' :::::::::::::::::::
        Case SPackets.Sspells
            a = frmMirage.lstSpells.ListIndex
            
            If a < 0 Then
                a = 0
            End If
            
            frmMirage.picPlayerSpells.Visible = True
            frmMirage.lstSpells.Clear

            ' Put spells known in player record
            For i = 1 To MAX_PLAYER_SPELLS
                Player(MyIndex).Spell(i) = CLng(parse(i))
                
                If Player(MyIndex).Spell(i) > 0 Then
                    frmMirage.lstSpells.addItem i & ": " & Trim$(Spell(Player(MyIndex).Spell(i)).Name)
                Else
                    frmMirage.lstSpells.addItem "--- Slot Free ---"
                End If
            Next i

            frmMirage.lstSpells.ListIndex = a
        Exit Sub
    
    ' ::::::::::::::::::::::::
    ' :: Spells info packet ::
    ' ::::::::::::::::::::::::
        Case SPackets.Sspecialattackinfo
            frmMirage.lblSpecialAtkName.Caption = Trim$(parse(1))
            frmMirage.lblSpecialAtkDmg = Trim$(parse(2))
            frmMirage.lblSpecialAtkFPCost.Caption = "FP Cost: " & parse(3)
            frmMirage.lblSpecialAtkRange.Caption = "Range: " & parse(4)
            frmMirage.lblSpecialAtkDesc.Caption = Trim$(parse(5))
        
            frmMirage.picSpecialAtkDetails.Visible = True
        Exit Sub

    ' ::::::::::::::::::::
    ' :: Weather packet ::
    ' ::::::::::::::::::::
        Case SPackets.Sweather
            z = CByte(parse(1))
            
            If z = WEATHER_RAINING And GameWeather <> WEATHER_RAINING Then
                Call AddText("You see drops of rain falling from the sky above!", BRIGHTGREEN)
                Call PlayBGS("rain.mp3")
            End If
            If z = WEATHER_THUNDER And GameWeather <> WEATHER_THUNDER Then
                Call AddText("You see thunder in the sky above!", BRIGHTGREEN)
                Call PlayBGS("thunder.mp3")
            End If
            If z = WEATHER_SNOWING And GameWeather <> WEATHER_SNOWING Then
                Call AddText("You see snow falling from the sky above!", BRIGHTGREEN)
            End If

            If z = WEATHER_NONE Then
                If GameWeather = WEATHER_RAINING Then
                    Call AddText("The rain beings to calm.", BRIGHTGREEN)
                    Call frmMirage.BGSPlayer.StopMedia
                ElseIf GameWeather = WEATHER_SNOWING Then
                    Call AddText("The snow is melting away.", BRIGHTGREEN)
                ElseIf GameWeather = WEATHER_THUNDER Then
                    Call AddText("The thunder begins to disapear.", BRIGHTGREEN)
                    Call frmMirage.BGSPlayer.StopMedia
                End If
            End If
        
            GameWeather = z
            RainIntensity = CLng(parse(2))
            
            If MAX_RAINDROPS <> RainIntensity Then
                MAX_RAINDROPS = RainIntensity
                ReDim DropRain(1 To MAX_RAINDROPS) As DropRainRec
                ReDim DropSnow(1 To MAX_RAINDROPS) As DropRainRec
            End If
        Exit Sub

    ' :::::::::::::::::::::
    ' :: Get Online List ::
    ' :::::::::::::::::::::
        Case SPackets.Sonlinelist
            frmMirage.lstOnline.Clear

            n = 2
            z = CLng(parse(1))
            
            Do While n <= (z + 1)
                frmMirage.lstOnline.addItem Trim$(parse(n))
                n = n + 1
            Loop
        Exit Sub
    
    ' ::::::::::::::::::::::::::
    ' :: Group Member Request ::
    ' ::::::::::::::::::::::::::
        Case SPackets.Sguildmemberrequest
            frmGroupMember.lblMessage.Caption = parse(1)
            Call frmGroupMember.SetGuildRequesterName(parse(2), CInt(parse(3)))
            
            Call frmGroupMember.Show(vbModeless, frmMirage)
        Exit Sub
        
    ' :::::::::::::::::::::::
    ' :: Group Member List ::
    ' :::::::::::::::::::::::
        Case SPackets.Sgroupmemberlist
            frmMirage.lstGroupMembers.Clear
        
            n = 2
            z = CInt(parse(1))
        
            Do While n <= (z + 1)
                frmMirage.lstGroupMembers.addItem Trim$(parse(n))
                n = n + 1
            Loop
        
            frmMirage.picGroupMembers.Visible = True
        Exit Sub
            
    ' ::::::::::::::::::::::::
    ' :: Blit Player Damage ::
    ' ::::::::::::::::::::::::
        Case SPackets.Sblitplayerdmg
            DmgDamage = CLng(parse(1))
            NPCWho = CLng(parse(2))
            DmgTime = GetTickCount
            iii = 0
        Exit Sub

    ' :::::::::::::::::::::
    ' :: Blit NPC Damage ::
    ' :::::::::::::::::::::
        Case SPackets.Sblitnpcdmg
            NPCDmgDamage = CLng(parse(1))
            NPCDmgTime = GetTickCount
            ii = 0
        Exit Sub

    ' :::::::::::::::::::
    ' :: Trade Request ::
    ' :::::::::::::::::::
        Case SPackets.Straderequest
            If IsBanking = True Or IsCooking = True Or IsShopping = True Then
                Call NotifyOtherPlayer(parse(1))
                Exit Sub
            End If
        
            frmTradeBox.Label1.Caption = parse(1) & " wants to trade with you."
            Call frmTradeBox.Show(vbModal, frmMirage)
        Exit Sub
            
    ' :::::::::::::::::
    ' :: Start Trade ::
    ' :::::::::::::::::
        Case SPackets.Sstarttrade
            ' Clear out the trade Listboxes
            frmPlayerTrade.Items1.Clear
            frmPlayerTrade.Items2.Clear
        
            ' Clear out any previous settings
            For i = 1 To MAX_PLAYER_TRADES
                PlayerTrading(i).InvName = vbNullString
                PlayerTrading(i).InvNum = 0
                PlayerTrading(i).InvVal = 0
                
                OtherPlayerTrading(i).InvName = vbNullString
                OtherPlayerTrading(i).InvNum = 0
                OtherPlayerTrading(i).InvVal = 0
                
                frmPlayerTrade.Items1.addItem i & ": <Nothing>"
                frmPlayerTrade.Items2.addItem i & ": <Nothing>"
            Next i
        
            frmPlayerTrade.TradingWith.Caption = "Trading With: " & parse(1)
            frmPlayerTrade.Items1.ListIndex = 0
            
            Call DisplayInventoryInTrade
            Call frmPlayerTrade.Show(vbModeless, frmMirage)
            
            IsTrading = True
        Exit Sub
        
    ' ::::::::::::::::::
    ' :: Stop Trading ::
    ' ::::::::::::::::::
        Case SPackets.Sstoptrading
            ' Reset any trade settings
            For i = 1 To MAX_PLAYER_TRADES
                PlayerTrading(i).InvName = vbNullString
                PlayerTrading(i).InvNum = 0
                PlayerTrading(i).InvVal = 0
                
                OtherPlayerTrading(i).InvName = vbNullString
                OtherPlayerTrading(i).InvNum = 0
                OtherPlayerTrading(i).InvVal = 0
            Next i
        
            frmPlayerTrade.Accepted.Caption = vbNullString
            IsTrading = False
            Unload frmPlayerTrade
        Exit Sub
        
    ' :::::::::::::::::::::::::
    ' :: Update Trade Offers ::
    ' :::::::::::::::::::::::::
        Case SPackets.Supdatetradeoffers
            n = CInt(parse(1))

            OtherPlayerTrading(n).InvNum = CLng(parse(2))
            OtherPlayerTrading(n).InvName = Trim$(parse(3))
            OtherPlayerTrading(n).InvVal = CLng(parse(4))
            
            a = CLng(parse(5))

            If OtherPlayerTrading(n).InvNum <= 0 Then
                frmPlayerTrade.Items2.List(n - 1) = n & ": <Nothing>"
                Exit Sub
            End If

            If Item(a).Type = ITEM_TYPE_CURRENCY Or Item(a).Stackable = 1 Then
                frmPlayerTrade.Items2.List(n - 1) = n & ": " & OtherPlayerTrading(n).InvName & " (" & OtherPlayerTrading(n).InvVal & ")"
            Else
                frmPlayerTrade.Items2.List(n - 1) = n & ": " & OtherPlayerTrading(n).InvName
            End If
        Exit Sub
    
    ' ::::::::::::::::::::
    ' :: Trade Messages ::
    ' ::::::::::::::::::::
        Case SPackets.Strademessage
            i = CLng(parse(1))
            a = CLng(parse(2))
        
            If CByte(parse(3)) = 0 Then
                frmPlayerTrade.Accepted.Caption = vbNullString
            Else
                If i = MyIndex Then
                    frmPlayerTrade.Accepted.Caption = GetPlayerName(a) & " has accepted the trade."
                Else
                    frmPlayerTrade.Accepted.Caption = "You have accepted the trade. Waiting for " & GetPlayerName(i) & "..."
                End If
            End If
        Exit Sub
    
    ' :::::::::::::::::::::::::
    ' :: Chat System Packets ::
    ' :::::::::::::::::::::::::
        Case SPackets.Sppchatting
            frmPlayerChat.txtChat.Text = vbNullString
            frmPlayerChat.txtSay.Text = vbNullString
            frmPlayerChat.Label1.Caption = "Chatting With: " & Trim$(parse(1))
            frmPlayerChat.Show vbModeless, frmMirage
        Exit Sub

        Case SPackets.Sqchat
            frmPlayerChat.txtChat.Text = vbNullString
            frmPlayerChat.txtSay.Text = vbNullString
            frmPlayerChat.Visible = False
        Exit Sub

        Case SPackets.Ssendchat
            frmPlayerChat.txtChat.SelStart = Len(frmPlayerChat.txtChat.Text)
            frmPlayerChat.txtChat.SelColor = QBColor(GREEN)
            frmPlayerChat.txtChat.SelText = vbNewLine & Trim$(parse(2)) & ": " & Trim$(parse(1))
            frmPlayerChat.txtChat.SelStart = Len(frmPlayerChat.txtChat.Text) - 1
        Exit Sub
    ' :::::::::::::::::::::::::::::
    ' :: END Chat System Packets ::
    ' :::::::::::::::::::::::::::::

    ' :::::::::::::::::::::::
    ' :: Play Sound Packet ::
    ' :::::::::::::::::::::::
        Case SPackets.Ssound
            s = parse(1)
            
            Select Case s
                Case "attack"
                    Call PlaySound("smw_stomp.wav")
                Case "critical"
                    Call PlaySound("mpds_critical.wav")
                Case "miss"
                    Call PlaySound("yi_notTryLevelAgain.wav")
                Case "key"
                    Call PlaySound("yi_doorUnlock.wav")
                Case "warp"
                    If FileExists("SFX\warp.wav") Then
                        Call PlaySound("warp.wav")
                    End If
                Case "pain"
                    Call PlaySound("smas-smb3_hit_with_shell.wav")
                Case "soundattribute"
                    If parse(2) <> "No Sound" Then
                        Call PlaySound(parse(2))
                    End If
            End Select
        Exit Sub

    ' :::::::::::::::::::
    ' :: Prompt Packet ::
    ' :::::::::::::::::::
        Case SPackets.Sprompt
            i = MsgBox(Trim$(parse(1)), vbYesNo)
            Call SendData(CPackets.Cprompt & SEP_CHAR & i & SEP_CHAR & parse(2) & END_CHAR)
        Exit Sub

    ' :::::::::::::::::::
    ' :: Prompt Packet ::
    ' :::::::::::::::::::
        Case SPackets.Squerybox
            frmQuery.Label1.Caption = Trim$(parse(1))
            frmQuery.Label2.Caption = parse(2)
            frmQuery.Show vbModal
        Exit Sub
    
    ' ::::::::::::::::::::::::::::::
    ' :: Turn-based Battle Packet ::
    ' ::::::::::::::::::::::::::::::
        Case SPackets.Sturnbasedbattle
            z = CLng(parse(1))
            n = CByte(parse(2))
            
            Select Case n
                Case 0 ' End battle
                    Player(z).InBattle = False
                    MapNpc(CLng(parse(3))).InBattle = False
                    
                    If z = MyIndex Then
                        BattleVictoryTimer = GetTickCount
                        IsPlayerTurn = False
                        Call PlayBGM(Map(GetPlayerMap(MyIndex)).music)
                    End If
                Case 1 ' Start Battle
                    Player(z).InBattle = True
                    MapNpc(CLng(parse(3))).InBattle = True
                    
                    If z = MyIndex Then
                        ButtonHighlighted = 2
                        IsPlayerTurn = True
                        Call PlayBattleSong
                    End If
                Case 2 ' First Strike
                    Player(z).InBattle = True
                    MapNpc(CLng(parse(3))).InBattle = True
                    
                    If z = MyIndex Then
                        ButtonHighlighted = 2
                        Call PlayBattleSong
                    End If
            End Select
        Exit Sub
        
    ' ::::::::::::::::::::::::::::::::::::::
    ' :: Turn-based Battle Victory Packet ::
    ' ::::::::::::::::::::::::::::::::::::::
        Case SPackets.Sturnbasedvictory
            For i = 2 To 8
                VictoryInfo(i - 1) = parse(i)
            Next
            
            BattleFrameCount = 0
            BattleVictoryTimer = GetTickCount
            StartedVictoryAnim = True
            BattleNPC = CLng(parse(1))
            
            Player(MyIndex).BattleVictory = True
            CanFinishBattle = False
            
            Call PlayBGM("Mario & Luigi Partners In Time - Victory.mp3")
        Exit Sub
    
    ' ::::::::::::::::::::::::::::
    ' :: Emoticon editor packet ::
    ' ::::::::::::::::::::::::::::
        Case SPackets.Semoticoneditor
            InEmoticonEditor = True

            Call frmIndex.Show(vbModeless, frmMirage)
            frmIndex.lstIndex.Clear

            For i = 0 To MAX_EMOTICONS
                frmIndex.lstIndex.addItem i & ": " & Trim$(Emoticons(i).Command)
            Next i

            frmIndex.lstIndex.ListIndex = 0
        Exit Sub
    
        Case SPackets.Seditemoticon
            n = CLng(parse(1))

            Emoticons(n).Command = Trim$(parse(2))
            Emoticons(n).Pic = CLng(parse(3))

            Call EmoticonEditorInit
        Exit Sub

        Case SPackets.Supdateemoticon
            n = CLng(parse(1))

            Emoticons(n).Command = Trim$(parse(2))
            Emoticons(n).Pic = CLng(parse(3))
        Exit Sub
    
    ' :::::::::::::::::::::::::::
    ' :: Element editor packet ::
    ' :::::::::::::::::::::::::::
        Case SPackets.Selementeditor
            InElementEditor = True

            Call frmIndex.Show(vbModeless, frmMirage)
            frmIndex.lstIndex.Clear

            For i = 0 To MAX_ELEMENTS
                frmIndex.lstIndex.addItem i & ": " & Trim$(Element(i).Name)
            Next i

            frmIndex.lstIndex.ListIndex = 0
        Exit Sub

        Case SPackets.Seditelement
            n = CLng(parse(1))

            Element(n).Name = Trim$(parse(2))
            Element(n).Strong = CLng(parse(3))
            Element(n).Weak = CLng(parse(4))

            Call ElementEditorInit
        Exit Sub

        Case SPackets.Supdateelement
            n = CLng(parse(1))

            Element(n).Name = Trim$(parse(2))
            Element(n).Strong = CLng(parse(3))
            Element(n).Weak = CLng(parse(4))
        Exit Sub

    ' :::::::::::::::::::::::::
    ' :: Arrow editor packet ::
    ' :::::::::::::::::::::::::
        Case SPackets.Sarroweditor
            InArrowEditor = True

            Call frmIndex.Show(vbModeless, frmMirage)
            frmIndex.lstIndex.Clear

            For i = 1 To MAX_ARROWS
                frmIndex.lstIndex.addItem i & ": " & Trim$(Arrows(i).Name)
            Next i

            frmIndex.lstIndex.ListIndex = 0
        Exit Sub

        Case SPackets.Seditarrow
            Arrows(CLng(parse(1))).Name = Trim$(parse(2))

            Call ArrowEditorInit
        Exit Sub

        Case SPackets.Supdatearrow
            n = CLng(parse(1))

            Arrows(n).Name = Trim$(parse(2))
            Arrows(n).Pic = CLng(parse(3))
            Arrows(n).Range = CByte(parse(4))
            Arrows(n).Amount = CLng(parse(5))
        Exit Sub

        Case SPackets.Scheckarrows
            n = CLng(parse(1))
            z = CLng(parse(2))
            i = CByte(parse(3))

            For x = 1 To MAX_PLAYER_ARROWS
                If Player(n).Arrow(x).Arrow = 0 Then
                    Player(n).Arrow(x).Arrow = 1
                    Player(n).Arrow(x).ArrowNum = z
                    Player(n).Arrow(x).ArrowAnim = Arrows(z).Pic
                    Player(n).Arrow(x).ArrowTime = GetTickCount
                    Player(n).Arrow(x).ArrowVarX = 0
                    Player(n).Arrow(x).ArrowVarY = 0
                    Player(n).Arrow(x).ArrowY = GetPlayerY(n)
                    Player(n).Arrow(x).ArrowX = GetPlayerX(n)
                    Player(n).Arrow(x).ArrowAmount = p
                    
                    Select Case i
                        Case DIR_DOWN
                            Player(n).Arrow(x).ArrowY = GetPlayerY(n) + 1
                            Player(n).Arrow(x).ArrowPosition = 0
                        
                            If Player(n).Arrow(x).ArrowY - 1 > MAX_MAPY Then
                                Player(n).Arrow(x).Arrow = 0
                                Exit Sub
                            End If
                        Case DIR_UP
                            Player(n).Arrow(x).ArrowY = GetPlayerY(n) - 1
                            Player(n).Arrow(x).ArrowPosition = 1
                        
                            If Player(n).Arrow(x).ArrowY + 1 < 0 Then
                                Player(n).Arrow(x).Arrow = 0
                                Exit Sub
                            End If
                        Case DIR_RIGHT
                            Player(n).Arrow(x).ArrowX = GetPlayerX(n) + 1
                            Player(n).Arrow(x).ArrowPosition = 2
                        
                            If Player(n).Arrow(x).ArrowX - 1 > MAX_MAPX Then
                                Player(n).Arrow(x).Arrow = 0
                                Exit Sub
                            End If
                        Case DIR_LEFT
                            Player(n).Arrow(x).ArrowX = GetPlayerX(n) - 1
                            Player(n).Arrow(x).ArrowPosition = 3
                        
                            If Player(n).Arrow(x).ArrowX + 1 < 0 Then
                                Player(n).Arrow(x).Arrow = 0
                                Exit Sub
                            End If
                    End Select
                    Exit Sub
                End If
            Next x
        Exit Sub

        Case SPackets.Schecksprite
            Player(CLng(parse(1))).Sprite = CLng(parse(2))
        Exit Sub

        Case SPackets.Smapreport
            n = 1

            frmMapReport.lstIndex.Clear
            
            For i = 1 To MAX_MAPS
                frmMapReport.lstIndex.addItem i & ": " & Trim$(parse(n))
                n = n + 1
            Next i

            frmMapReport.Show vbModeless, frmMirage
        Exit Sub

    ' :::::::::::::::::::::::
    ' :: Spell anim packet ::
    ' :::::::::::::::::::::::
        Case SPackets.Sspellanim
            a = CLng(parse(1))
            z = CLng(parse(5))
            
            Spell(a).SpellAnim = CLng(parse(2))
            Spell(a).SpellTime = CLng(parse(3))
            Spell(a).SpellDone = CLng(parse(4))
            Spell(a).Big = CLng(parse(9))

            Player(z).SpellNum = a

            For i = 1 To MAX_SPELL_ANIM
                If Player(z).SpellAnim(i).CastedSpell = No Then
                    Player(z).SpellAnim(i).SpellDone = 0
                    Player(z).SpellAnim(i).SpellVar = 0
                    Player(z).SpellAnim(i).SpellTime = GetTickCount
                    Player(z).SpellAnim(i).TargetType = CLng(parse(6))
                    Player(z).SpellAnim(i).Target = CLng(parse(7))
                    Player(z).SpellAnim(i).CastedSpell = Yes
                    Exit For
                End If
            Next i
        Exit Sub

    ' :::::::::::::::::::::::
    ' :: Script Spell anim packet ::
    ' :::::::::::::::::::::::
        Case SPackets.Sscriptspellanim
            a = CLng(parse(1))
            
            Spell(a).SpellAnim = CLng(parse(2))
            Spell(a).SpellTime = CLng(parse(3))
            Spell(a).SpellDone = CLng(parse(4))
            Spell(a).Big = CLng(parse(7))

            For i = 1 To MAX_SCRIPTSPELLS
                If ScriptSpell(i).CastedSpell = No Then
                    ScriptSpell(i).SpellNum = a
                    ScriptSpell(i).SpellDone = 0
                    ScriptSpell(i).SpellVar = 0
                    ScriptSpell(i).SpellTime = GetTickCount
                    ScriptSpell(i).x = CLng(parse(5))
                    ScriptSpell(i).y = CLng(parse(6))
                    ScriptSpell(i).Index = CLng(parse(8))
                    ScriptSpell(i).CastedSpell = Yes
                    Exit For
                End If
            Next i
        Exit Sub

        Case SPackets.Scheckemoticons
            n = CLng(parse(1))

            Player(n).EmoticonNum = CLng(parse(2))
            Player(n).EmoticonTime = GetTickCount
            Player(n).EmoticonVar = 0
        Exit Sub

        Case SPackets.Slevelup
            i = CLng(parse(1))
        
            Player(i).LevelUpT = GetTickCount
            Player(i).LevelUp = 1
            
            If i = MyIndex Then
                Call PlayBGM("Paper Mario The Thousand Year Door - Level Up.mp3")
                Call frmLevelUp.Show(vbModeless, frmMirage)
            End If
        Exit Sub

        Case SPackets.Sdamagedisplay
            p = CByte(parse(1))
            s = parse(2)
        
            For i = 1 To MAX_BLT_LINE
                If p = 0 Then
                    If BattlePMsg(i).Index <= 0 Then
                        BattlePMsg(i).Index = 1
                        BattlePMsg(i).Msg = s
                        BattlePMsg(i).Color = CByte(parse(3))
                        BattlePMsg(i).Time = GetTickCount
                        BattlePMsg(i).Done = 1
                        BattlePMsg(i).y = 0
                        Exit Sub
                    Else
                        BattlePMsg(i).y = BattlePMsg(i).y - 15
                    End If
                Else
                    If BattleMMsg(i).Index <= 0 Then
                        BattleMMsg(i).Index = 1
                        BattleMMsg(i).Msg = s
                        BattleMMsg(i).Color = CByte(parse(3))
                        BattleMMsg(i).Time = GetTickCount
                        BattleMMsg(i).Done = 1
                        BattleMMsg(i).y = 0
                        Exit Sub
                    Else
                        BattleMMsg(i).y = BattleMMsg(i).y - 15
                    End If
                End If
            Next i
            
            z = 1
            
            If p = 0 Then
                For i = 1 To MAX_BLT_LINE
                    If i < MAX_BLT_LINE Then
                        If BattlePMsg(i).y < BattlePMsg(i + 1).y Then
                            z = i
                        End If
                    Else
                        If BattlePMsg(i).y < BattlePMsg(1).y Then
                            z = i
                        End If
                    End If
                Next i

                BattlePMsg(z).Index = 1
                BattlePMsg(z).Msg = s
                BattlePMsg(z).Color = CByte(parse(3))
                BattlePMsg(z).Time = GetTickCount
                BattlePMsg(z).Done = 1
                BattlePMsg(z).y = 0
            Else
                For i = 1 To MAX_BLT_LINE
                    If i < MAX_BLT_LINE Then
                        If BattleMMsg(i).y < BattleMMsg(i + 1).y Then
                            z = i
                        End If
                    Else
                        If BattleMMsg(i).y < BattleMMsg(1).y Then
                            z = i
                        End If
                    End If
                Next i

                BattleMMsg(z).Index = 1
                BattleMMsg(z).Msg = s
                BattleMMsg(z).Color = CByte(parse(3))
                BattleMMsg(z).Time = GetTickCount
                BattleMMsg(z).Done = 1
                BattleMMsg(z).y = 0
            End If
        Exit Sub

    ' ::::::::::::::::::::::::::::::::::::::::
    ' :: Index player worn equipment packet ::
    ' ::::::::::::::::::::::::::::::::::::::::
        Case SPackets.Sitemworn
            i = CLng(parse(1))
            
            For z = 2 To 8
                Call SetPlayerEquipSlotNum(i, z - 1, CLng(parse(z)))
            Next z
        Exit Sub
        
        Case SPackets.Scanusespecial
            CanUseSpecial = True
        Exit Sub
        
        Case SPackets.Splayerheight
            Player(MyIndex).Height = CInt(parse(1))
        Exit Sub
        
        Case SPackets.Sjumping
            i = CLng(parse(1))
            
            Player(i).JumpDir = CByte(parse(2))
            Player(i).TempJumpAnim = CByte(parse(3))
            Player(i).JumpAnim = CByte(parse(4))
            Player(i).Jumping = True
            
            Player(i).JumpTime = GetTickCount
            
            ' Play sound based on player's character
            Select Case GetPlayerClass(i)
                ' Mario and Luigi
                Case 0, 1
                    Call PlaySound("m&lss_Jump Normal.wav")
                ' Wario
                Case 2
                    Call PlaySound("warioland4_jump.wav")
                ' Waluigi and Toad
                Case 3, 5
                    Call PlaySound("smas-smb2_jump.wav")
                ' Yoshi
                Case 4
                    Call PlaySound("yi_jump.wav")
            End Select
        Exit Sub
        
        Case SPackets.Sendjump
            i = CLng(parse(1))
            
            Player(i).Jumping = False
        Exit Sub
        
        Case SPackets.Shiderfreeze
            IsHiderFrozen = CBool(parse(1))
        Exit Sub
        
        Case SPackets.Shidensneak
            IsPlayingHideNSneak = CBool(parse(1))
        Exit Sub
    End Select
End Sub

Sub ShowCookForm(ByVal NpcNumCook As Long)
    Dim i As Integer
    
    ' Determine the Npc Number of the cook
    CookNpcNum = NpcNumCook
            
    frmCooking.lstInventory.Clear
        
    For i = 1 To Player(MyIndex).MaxInv
        If GetPlayerInvItemNum(MyIndex, i) > 0 Then
            frmCooking.lstInventory.addItem i & ": " & Trim$(Item(GetPlayerInvItemNum(MyIndex, i)).Name)
        Else
            frmCooking.lstInventory.addItem i & ": "
        End If
    Next i
        
    Call frmCooking.Show(vbModal, frmMirage)
End Sub
