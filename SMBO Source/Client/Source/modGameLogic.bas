Attribute VB_Name = "modGameLogic"
Option Explicit

Public Function TwipsToPixels(lngTwips As Long, lngDirection As Long) As Long
    ' Handle to device
    Dim lngDC As Long
    Dim lngPixelsPerInch As Long
    Const nTwipsPerInch = 1440
    lngDC = GetDC(0)

    If (lngDirection = 0) Then       'Horizontal
        lngPixelsPerInch = GetDeviceCaps(lngDC, 88)
    Else                            'Vertical
        lngPixelsPerInch = GetDeviceCaps(lngDC, 90)
    End If
    
    lngDC = ReleaseDC(0, lngDC)
    TwipsToPixels = (lngTwips / nTwipsPerInch) * lngPixelsPerInch
End Function

Public Function PixelsToTwips(lngTwips As Long, lngDirection As Long) As Long
    ' Handle to device
    Dim lngDC As Long
    Dim lngPixelsPerInch As Long
    Const nTwipsPerInch = 1440
    lngDC = GetDC(0)

    If (lngDirection = 0) Then       'Horizontal
        lngPixelsPerInch = GetDeviceCaps(lngDC, 88)
    Else                            'Vertical
        lngPixelsPerInch = GetDeviceCaps(lngDC, 90)
    End If
    
    lngDC = ReleaseDC(0, lngDC)
    PixelsToTwips = (lngTwips / lngPixelsPerInch) * nTwipsPerInch
End Function

Sub SetStatus(ByVal Caption As String)
    frmSendGetData.lblStatus.Caption = Caption
    DoEvents
End Sub

Sub MenuState(ByVal State As Long)
    Dim OwnerStatus As Boolean

    Connected = True

    frmSendGetData.Visible = True

    Call SetStatus("Connecting to Server...")

    Select Case State
        Case MENU_STATE_NEWACCOUNT
            frmNewAccount.Visible = False

            If ConnectToServer Then
                Call SetStatus("Connected! Creating Account...")

                Call SendNewAccount(frmNewAccount.txtName.Text, frmNewAccount.txtPassword.Text)
            End If

        Case MENU_STATE_DELACCOUNT
            frmDeleteAccount.Visible = False

            If ConnectToServer Then
                Call SetStatus("Connected. Deleting Account...")

                Call SendDelAccount(frmDeleteAccount.txtName.Text, frmDeleteAccount.txtPassword.Text)
            End If

        Case MENU_STATE_LOGIN
            frmLogin.Visible = False

            If ConnectToServer Then
                Call SetStatus("Connected. Logging In...")
                
                ' Check if the user is authorized to use the "Kimimaru" and "hydrakiller4000" accounts
                If FileExists("RealGFX\Items.bmp") And FileExists("RealGUI\pvpsign.bmp") Then
                    OwnerStatus = True
                End If
                
                Call SendLogin(Trim$(frmLogin.txtName.Text), Trim$(frmLogin.txtPassword.Text), OwnerStatus)
            End If

        Case MENU_STATE_AUTO_LOGIN
            frmMainMenu.Visible = False

            If ConnectToServer Then
                Call SetStatus("Connected. Logging In...")
                
                ' Check if the user is authorized to use the "Kimimaru" and "hydrakiller4000" accounts
                If FileExists("RealGFX\Items.bmp") And FileExists("RealGUI\pvpsign.bmp") Then
                    OwnerStatus = True
                End If
                
                Call SendLogin(Trim$(frmLogin.txtName.Text), Trim$(frmLogin.txtPassword.Text), OwnerStatus)
            End If

        Case MENU_STATE_NEWCHAR
            frmChars.Visible = False

            If ConnectToServer Then
                Call SetStatus("Connected. Receiving Classes...")

                frmNewChar.Picture4.Top = frmNewChar.Picture4.Top - 32
                frmNewChar.Picture4.Height = 69
                frmNewChar.picPic.Height = 65

                Call SendGetClasses
            End If

        Case MENU_STATE_ADDCHAR
            frmNewChar.Visible = False

            If ConnectToServer Then
                Call SetStatus("Connected. Creating Character...")

                If frmNewChar.optMale.Value Then
                    Call SendAddChar(frmNewChar.txtName, 0, frmNewChar.cmbClass.ListIndex, frmChars.lstChars.ListIndex + 1)
                Else
                    Call SendAddChar(frmNewChar.txtName, 1, frmNewChar.cmbClass.ListIndex, frmChars.lstChars.ListIndex + 1)
                End If
            End If

        Case MENU_STATE_DELCHAR
            frmChars.Visible = False

            If ConnectToServer Then
                Call SetStatus("Connected. Deleting Character...")

                Call SendDelChar(frmChars.lstChars.ListIndex + 1)
            End If

        Case MENU_STATE_USECHAR
            frmChars.Visible = False

            If ConnectToServer Then
                Call SetStatus("Connected. Entering Super Mario Bros. Online...")

                Call SendUseChar(frmChars.lstChars.ListIndex + 1)
            End If
    End Select

    If Not IsConnected And Connected = True Then
        frmSendGetData.Visible = False
        frmMainMenu.Visible = True

        Call MsgBox("The server is currently offline. Please try to reconnect in a few minutes or visit: http://www.supermariobrosonline.co.cc", vbOKOnly, "Super Mario Bros. Online")
    End If
End Sub

Sub GameInit()
    Call InitDirectX
    Call StopBGM

    InGame = True
    
    ' Unload main menu forms after character logs in.
    Unload frmSendGetData
    Unload frmMainMenu
    Unload frmChars
    Unload frmNewChar
    Unload frmSendGetData
    
    ' Load pictures
    Call InitLoadPicture(App.Path & "\GUI\minimenus.smbo", frmMirage.picCharStatus)
    Call InitLoadPicture(App.Path & "\GUI\minimenusequipment.smbo", frmMirage.picEquipment)
    Call InitLoadPicture(App.Path & "\GUI\minimenusspecialattacks.smbo", frmMirage.picPlayerSpells)
    Call InitLoadPicture(App.Path & "\GUI\minimenusinventory.smbo", frmMirage.picInventory)
    Call InitLoadPicture(App.Path & "\GUI\minimenusgroup.smbo", frmMirage.picGuildAdmin)
    Call InitLoadPicture(App.Path & "\GUI\minimenus.smbo", frmMirage.picWhosOnline)
    Call InitLoadPicture(App.Path & "\GUI\minimenusgroup.smbo", frmMirage.picGuildMember)
    'Call InitLoadPicture(App.Path & "\GUI\minimenusinventory.smbo", frmMirage.picInventory3)
    
    Call StopBGM
    frmMirage.Visible = True

    On Error Resume Next
    
    ' Update the user's inventory in case it wasn't updated before
    UpdateVisInv
    
    ' Set the focus to the main form since only focused objects may set the focus
    frmMirage.SetFocus

    frmMirage.picScreen.SetFocus
End Sub

Sub GameLoop()
    Dim Tick As Long, TickFPS As Long, FPS As Long, x As Long, y As Long, i As Long, z As Long

    On Error Resume Next

    ' Used for calculating fps
    TickFPS = GetTickCount
    FPS = 0

    ' *************************************************
    ' * SUPER MARIO BROS. ONLINE MAIN GAME LOOP BEGIN *
    ' *************************************************
    Do While InGame
        Tick = GetTickCount
        
        If frmMirage.WindowState = 0 Then
        
            ' Check if we need to restore surfaces
            If NeedToRestoreSurfaces Then
                DD.RestoreAllSurfaces
                Call InitSurfaces
            End If

            If Not GettingMap Then

                ' Check to make sure they aren't trying to auto do anything
                If GetAsyncKeyState(VK_UP) >= 0 Then
                    If DirUp Then
                        DirUp = False
                    End If
                End If
                If GetAsyncKeyState(VK_DOWN) >= 0 Then
                    If DirDown Then
                        DirDown = False
                    End If
                End If
                If GetAsyncKeyState(VK_LEFT) >= 0 Then
                    If DirLeft Then
                        DirLeft = False
                    End If
                End If
                If GetAsyncKeyState(VK_RIGHT) >= 0 Then
                    If DirRight Then
                        DirRight = False
                    End If
                End If
                If GetAsyncKeyState(VK_CONTROL) >= 0 Then
                    If ControlDown Then
                        ControlDown = False
                    End If
                End If
                If GetAsyncKeyState(VK_SHIFT) >= 0 Then
                    If ShiftDown Then
                        ShiftDown = False
                    End If
                End If

                ' Check to make sure we are still connected
                If Not IsConnected Then
                    InGame = False
                    Exit Do
                End If

                ' Update the user's inventory
                Call UpdateInventory
                
                NewX = 10
                NewY = 7

                NewPlayerY = Player(MyIndex).y - NewY
                NewPlayerX = Player(MyIndex).x - NewX

                NewX = NewX * PIC_X
                NewY = NewY * PIC_Y

                NewXOffset = Player(MyIndex).xOffset
                NewYOffset = Player(MyIndex).yOffset

                If Player(MyIndex).y - 7 < 1 Then
                    NewY = Player(MyIndex).y * PIC_Y + Player(MyIndex).yOffset
                    NewYOffset = 0
                    NewPlayerY = 0
                    
                    If Player(MyIndex).y = 7 Then
                        If Player(MyIndex).Dir = DIR_UP Then
                            NewPlayerY = Player(MyIndex).y - 7
                            NewY = 7 * PIC_Y
                            NewYOffset = Player(MyIndex).yOffset
                        End If
                    End If
                ElseIf Player(MyIndex).y + 9 > MAX_MAPY + 1 Then
                    NewY = (Player(MyIndex).y - (MAX_MAPY - 14)) * PIC_Y + Player(MyIndex).yOffset
                    NewYOffset = 0
                    NewPlayerY = MAX_MAPY - 14
                    
                    If Player(MyIndex).y = MAX_MAPY - 7 Then
                        If Player(MyIndex).Dir = DIR_DOWN Then
                            NewPlayerY = Player(MyIndex).y - 7
                            NewY = 7 * PIC_Y
                            NewYOffset = Player(MyIndex).yOffset
                        End If
                    End If
                End If

                If Player(MyIndex).x - 10 < 1 Then
                    NewX = Player(MyIndex).x * PIC_X + Player(MyIndex).xOffset
                    NewXOffset = 0
                    NewPlayerX = 0
                    
                    If Player(MyIndex).x = 10 Then
                        If Player(MyIndex).Dir = DIR_LEFT Then
                            NewPlayerX = Player(MyIndex).x - 10
                            NewX = 10 * PIC_X
                            NewXOffset = Player(MyIndex).xOffset
                        End If
                    End If
                ElseIf Player(MyIndex).x + 11 > MAX_MAPX + 1 Then
                    NewX = (Player(MyIndex).x - (MAX_MAPX - 19)) * PIC_X + Player(MyIndex).xOffset
                    NewXOffset = 0
                    NewPlayerX = MAX_MAPX - 19
                    
                    If Player(MyIndex).x = MAX_MAPX - 9 Then
                        If Player(MyIndex).Dir = DIR_RIGHT Then
                            NewPlayerX = Player(MyIndex).x - 10
                            NewX = 10 * PIC_X
                            NewXOffset = Player(MyIndex).xOffset
                        End If
                    End If
                End If

                ScreenX = GetScreenLeft(MyIndex)
                ScreenY = GetScreenTop(MyIndex)
                ScreenX2 = GetScreenRight(MyIndex)
                ScreenY2 = GetScreenBottom(MyIndex)

                If ScreenX < 0 Then
                    ScreenX = 0
                    ScreenX2 = 20
                ElseIf ScreenX2 > MAX_MAPX Then
                    ScreenX2 = MAX_MAPX
                    ScreenX = MAX_MAPX - 20
                End If
            
                If ScreenY < 0 Then
                    ScreenY = 0
                    ScreenY2 = 15
                ElseIf ScreenY2 > MAX_MAPY Then
                    ScreenY2 = MAX_MAPY
                    ScreenY = MAX_MAPY - 15
                End If

                sx = 32
                If MAX_MAPX = 19 Then
                    NewX = Player(MyIndex).x * PIC_X + Player(MyIndex).xOffset
                    NewXOffset = 0
                    NewPlayerX = 0
                    NewY = Player(MyIndex).y * PIC_Y + Player(MyIndex).yOffset
                    NewYOffset = 0
                    NewPlayerY = 0
                    ScreenX = 0
                    ScreenY = 0
                    ScreenX2 = MAX_MAPX
                    ScreenY2 = MAX_MAPY
                    sx = 0
                End If

                ' Blit out tiles layers ground/anim1/anim2
                For y = ScreenY To ScreenY2
                    For x = ScreenX To ScreenX2
                        Call BltTile(x, y)
                    Next x
                Next y

                If ScreenMode = 0 Then
                
                    ' Blit out the items
                    For i = 1 To MAX_MAP_ITEMS
                        If MapItem(i).num > 0 Then
                            Call BltItem(i)
                        End If
                    Next i
                    
                    ' Blit out NPC hp bars
                    If frmMirage.chkNpcBar.Value = Checked Then
                        For i = 1 To MAX_MAP_NPCS
                            Call BltNpcBars(i)
                        Next i
                    End If
                    
                     ' Blit players bar
                    If frmMirage.chkPlayerBar.Value = Checked Then
                        For i = 1 To MAX_PLAYERS
                            If IsPlaying(i) Then
                                If GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                                  If GetPlayerGuild(i) = GetPlayerGuild(MyIndex) Then
                                    Call BltPlayerBars(i)
                                    Call BltPlayerSPBars(i)
                                  End If
                                End If
                            End If
                        Next i
                    End If
                
                    ' Rendering based on the Y-axis
                    For y = 0 To MAX_MAPY
                        ' Blit out players and arrows
                        For i = 1 To MAX_PLAYERS
                            If IsPlaying(i) Then
                                If GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                                    If Player(i).y = y Then
                                        Call BltPlayer(i)
                                        Call BltJump(i)
                                        Call BltArrow(i)
                                    End If
                                End If
                            End If
                        Next i
                    
                        ' Blit out the npc
                        For i = 1 To MAX_MAP_NPCS
                            If MapNpc(i).num > 0 Then
                                If MapNpc(i).y = y Then
                                    Call BltNpc(i)
                                End If
                            End If
                        Next i
                    Next y

                    ' Blt out the spells
                    For i = 1 To MAX_PLAYERS
                        If IsPlaying(i) Then
                            If GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                                Call BltSpell(i)
                            End If
                        End If
                    Next i

                    ' Blt out the scripted spells
                    For i = 1 To MAX_SCRIPTSPELLS
                        If ScriptSpell(i).SpellNum > 0 Then
                            If ScriptSpell(i).SpellNum <= MAX_SPELLS Then
                                If ScriptSpell(i).CastedSpell = Yes Then
                                    Call BltScriptSpell(i)
                                End If
                            End If
                        End If
                    Next i
                    
                    ' Draw 'level up!' text
                    For i = 1 To MAX_PLAYERS
                        If IsPlaying(i) Then
                            If GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                                Call BltLevelUp(i)
                            End If
                        End If
                    Next i

                End If

                ' Blit out tile layer fringe
                For y = ScreenY To ScreenY2
                    For x = ScreenX To ScreenX2
                        Call BltFringeTile(x, y)
                    Next x
                Next y

                ' Check for roof tiles
                For y = ScreenY To ScreenY2
                    For x = ScreenX To ScreenX2
                        If Not IsTileRoof(x, y) Then
                            Call BltFringe2Tile(x, y)
                        End If
                    Next x
                Next y
                
                ' Blit out emoticons
                For i = 1 To MAX_PLAYERS
                    If IsPlaying(i) Then
                        If GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                            Call BltEmoticons(i)
                        End If
                    End If
                Next i
            
                ' Draw weather (for all players)
                If Map(GetPlayerMap(MyIndex)).Indoors = 0 Then
                    If Map(GetPlayerMap(MyIndex)).Weather <> 0 Then
                        Call BltMapWeather
                    End If
            
                    Call BltWeather
                End If

                If InEditor Then
                    If GridMode = 1 Then
                        For y = ScreenY To ScreenY2
                            For x = ScreenX To ScreenX2
                                Call BltTile2(x * PIC_X, y * PIC_Y, 0)
                            Next x
                        Next y
                    End If
                End If
                
                ' Timer for turn-based battles
                If TurnBasedTimer = True Then
                    If GetTickCount > (TurnBasedTime + TurnBasedTimeToWait) Then
                        Call SendData(CPackets.Csetbattleturn & SEP_CHAR & PlayerTurn & END_CHAR)
                        TurnBasedTimer = False
                    End If
                End If
                
                ' Draw elements onto picScreen
                If Map(GetPlayerMap(MyIndex)).Moral = MAP_MORAL_NONE Then
                    Call BltPvPSign(605, 470)
                End If
                
                Call BltHPIcon(53, 38)
                Call BltFPIcon(53, 78)
                Call BltExpIcon(53, 118)
                
                If Player(MyIndex).InBattle = True Then
                    ' Draw icons while in a turn-based battle
                    If IsPlayerTurn = True Then
                        Call BltBattleIcons
                    End If
                    
                    ' Draw battle victory animation and screen
                    If Player(MyIndex).BattleVictory = True Then
                        Call BltVictory
                    End If
                End If
                
                ' Timer for Cooking
                If CookingTimer = True Then
                    If GetTickCount > CookingTime + 4000 Then
                        Call FinishCooking(RecipeNumber)
                        CookingTimer = False
                        IsCooking = False
                    End If
                End If
                
                ' Lock the backbuffer so we can draw text and names
                TexthDC = DD_BackBuffer.GetDC
                    
                    ' Draws text for HP
                    Call DrawText(TexthDC, 120, 55, GetPlayerHP(MyIndex) & " / " & GetPlayerMaxHP(MyIndex), QBColor(BLACK), GameFont2)
                    ' Draws text for FP
                    Call DrawText(TexthDC, 120, 95, GetPlayerMP(MyIndex) & " / " & GetPlayerMaxMP(MyIndex), QBColor(BLACK), GameFont2)
                    ' Draws text for Exp
                    Call DrawText(TexthDC, 99, 135, GetPlayerExp(MyIndex) & " / " & GetPlayerNextLvlExp(MyIndex), QBColor(BLACK), GameFont2)
                    
                    Dim TempX As Long, TempY As Long
                    
                    TempX = NewX + sx
                    TempY = NewY + sx
                    
                    If Player(MyIndex).InBattle = True And IsPlayerTurn = True Then
                        ' Draws text for Attack button in turn-based battles
                        Call DrawText(TexthDC, TempX + 81, TempY - 79, "Attack", QBColor(BLACK), GameFont2)
                        ' Draws text for Items button in turn-based battles
                        Call DrawText(TexthDC, TempX + 3, TempY - 99, "Items", QBColor(BLACK), GameFont2)
                        ' Draws text for Run button in turn-based battles
                        Call DrawText(TexthDC, TempX - 70, TempY - 79, "Run", QBColor(BLACK), GameFont2)
                        ' Draws text for Special button in turn-based battles
                        Call DrawText(TexthDC, TempX + 129, TempY - 14, "Special", QBColor(BLACK), GameFont2)
                    End If
                    
                    ' Draws text for the victory screen
                    If Player(MyIndex).BattleVictory = True And DisplayInfo = True Then
                        ' Draw Exp
                        Call DrawText(TexthDC, TempX - 85, TempY + 52, VictoryInfo(1), QBColor(WHITE), GameFont3)
                        ' Draw Coins
                        Call DrawText(TexthDC, TempX - 85, TempY + 88, VictoryInfo(2), QBColor(WHITE), GameFont3)
                        
                        Dim VictoryCount As Byte
                        
                        VictoryCount = 0
                        
                        For x = 3 To 7
                            If VictoryInfo(x) <> "Placeholder" Then
                                Call DrawText(TexthDC, TempX - 20, TempY + (62 + (VictoryCount * 10)), VictoryInfo(x), QBColor(WHITE), GameFont4)
                                VictoryCount = VictoryCount + 1
                            End If
                        Next
                    End If
                    
                If ScreenMode = 0 Then
                
                    ' Draw NPC's damage on player
                    If frmMirage.chkNpcDamage.Value = 1 Then
                        If frmMirage.chkPlayerName.Value = 0 Then
                            If GetTickCount < NPCDmgTime + 2000 Then
                                Call DrawText(TexthDC, (Int(Len(NPCDmgDamage)) / 2) * 3 + TempX, TempY - 22 - ii, NPCDmgDamage, QBColor(BRIGHTRED), GameFont)
                            End If
                        Else
                            If GetPlayerGuild(MyIndex) <> vbNullString Then
                                If GetTickCount < NPCDmgTime + 2000 Then
                                    Call DrawText(TexthDC, (Int(Len(NPCDmgDamage)) / 2) * 3 + TempX, TempY - 42 - ii, NPCDmgDamage, QBColor(BRIGHTRED), GameFont)
                                End If
                            Else
                                If GetTickCount < NPCDmgTime + 2000 Then
                                    Call DrawText(TexthDC, (Int(Len(NPCDmgDamage)) / 2) * 3 + TempX, TempY - 22 - ii, NPCDmgDamage, QBColor(BRIGHTRED), GameFont)
                                End If
                            End If
                        End If
                        ii = ii + 1
                    End If

                    ' Draw player's damage on NPC
                    If frmMirage.chkPlayerDamage.Value = 1 Then
                        If NPCWho > 0 Then
                            If MapNpc(NPCWho).num > 0 Then
                                TempX = (MapNpc(NPCWho).x - NewPlayerX) * PIC_X + sx + (Int(Len(DmgDamage)) / 2) * 3 + MapNpc(NPCWho).xOffset - NewXOffset
                                TempY = (MapNpc(NPCWho).y - NewPlayerY) * PIC_Y + sx + MapNpc(NPCWho).yOffset - NewYOffset - iii
                                
                                If frmMirage.chkNpcName.Value = 0 Then
                                    If Npc(MapNpc(NPCWho).num).Big = 0 Then
                                        TempY = TempY - 20
                                    Else
                                        TempY = TempY - 47
                                    End If
                                    
                                    If GetTickCount < DmgTime + 2000 Then
                                        Call DrawText(TexthDC, TempX, TempY, DmgDamage, QBColor(WHITE), GameFont)
                                    End If
                                Else
                                    If Npc(MapNpc(NPCWho).num).Big = 0 Then
                                        TempY = TempY - 30
                                    Else
                                        TempY = TempY - 57
                                    End If
                                    
                                    If GetTickCount < DmgTime + 2000 Then
                                        Call DrawText(TexthDC, TempX, TempY, DmgDamage, QBColor(WHITE), GameFont)
                                    End If
                                End If
                                
                                iii = iii + 1
                            End If
                        End If
                    End If
                    
                    ' Draw player name and guild name
                    If frmMirage.chkPlayerName.Value = 1 Then
                        For i = 1 To MAX_PLAYERS
                            If IsPlaying(i) Then
                                If GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                                  If GetPlayerMap(MyIndex) <> 21 Then
                                    Call BltPlayerGuildName(i)
                                    Call BltPlayerName(i)
                                  End If
                                End If
                            End If
                        Next i
                    End If

                    ' speech bubble stuffs
                    If ReadINI("CONFIG", "SpeechBubbles", App.Path & "\config.ini") = 1 Then
                        For i = 1 To MAX_PLAYERS
                            If IsPlaying(i) Then
                                If GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                                    If Bubble(i).Text <> vbNullString Then
                                        Call BltPlayerText(i)
                                    End If
    
                                    If GetTickCount() > Bubble(i).Created + DISPLAY_BUBBLE_TIME Then
                                        Bubble(i).Text = vbNullString
                                    End If
                                End If
                            End If
                        Next i
                    End If

                    ' Draw NPC Names
                    If ReadINI("CONFIG", "NPCName", App.Path & "\config.ini") = 1 Then
                        For i = LBound(MapNpc) To UBound(MapNpc)
                            If MapNpc(i).num > 0 Then
                                Call BltMapNPCName(i)
                            End If
                        Next i
                    End If

                    ' Blit out attribs if in editor
                    If InEditor Then
                        TempX = sx + 8 - (NewPlayerX * PIC_X) - NewXOffset
                        TempY = sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset
                        
                        For y = 0 To MAX_MAPY
                            For x = 0 To MAX_MAPX
                                With Map(GetPlayerMap(MyIndex)).Tile(x, y)
                                    Select Case .Type
                                        Case TILE_TYPE_BLOCKED
                                            Call DrawText(TexthDC, (x * PIC_X) + TempX, (y * PIC_Y) + TempY, "B", QBColor(BRIGHTRED), GameFont)
                                        Case TILE_TYPE_WARP
                                            Call DrawText(TexthDC, (x * PIC_X) + TempX, (y * PIC_Y) + TempY, "W", QBColor(BRIGHTBLUE), GameFont)
                                        Case TILE_TYPE_ITEM
                                            Call DrawText(TexthDC, (x * PIC_X) + TempX, (y * PIC_Y) + TempY, "I", QBColor(WHITE), GameFont)
                                        Case TILE_TYPE_NPCAVOID
                                            Call DrawText(TexthDC, (x * PIC_X) + TempX, (y * PIC_Y) + TempY, "N", QBColor(WHITE), GameFont)
                                        Case TILE_TYPE_KEY
                                            Call DrawText(TexthDC, (x * PIC_X) + TempX, (y * PIC_Y) + TempY, "K", QBColor(WHITE), GameFont)
                                        Case TILE_TYPE_KEYOPEN
                                            Call DrawText(TexthDC, (x * PIC_X) + TempX, (y * PIC_Y) + TempY, "O", QBColor(WHITE), GameFont)
                                        Case TILE_TYPE_HEAL
                                            Call DrawText(TexthDC, (x * PIC_X) + TempX, (y * PIC_Y) + TempY, "H", QBColor(BRIGHTGREEN), GameFont)
                                        Case TILE_TYPE_KILL
                                            Call DrawText(TexthDC, (x * PIC_X) + TempX, (y * PIC_Y) + TempY, "K", QBColor(BRIGHTRED), GameFont)
                                        Case TILE_TYPE_SHOP
                                            Call DrawText(TexthDC, (x * PIC_X) + TempX, (y * PIC_Y) + TempY, "S", QBColor(YELLOW), GameFont)
                                        Case TILE_TYPE_CBLOCK
                                            Call DrawText(TexthDC, (x * PIC_X) + TempX, (y * PIC_Y) + TempY, "CB", QBColor(BLACK), GameFont)
                                        Case TILE_TYPE_ARENA
                                            Call DrawText(TexthDC, (x * PIC_X) + TempX, (y * PIC_Y) + TempY, "A", QBColor(BRIGHTGREEN), GameFont)
                                        Case TILE_TYPE_SOUND
                                            Call DrawText(TexthDC, (x * PIC_X) + TempX, (y * PIC_Y) + TempY, "PS", QBColor(YELLOW), GameFont)
                                        Case TILE_TYPE_SPRITE_CHANGE
                                            Call DrawText(TexthDC, (x * PIC_X) + TempX, (y * PIC_Y) + TempY, "SC", QBColor(GREY), GameFont)
                                        Case TILE_TYPE_SIGN
                                            Call DrawText(TexthDC, (x * PIC_X) + TempX, (y * PIC_Y) + TempY, "SI", QBColor(YELLOW), GameFont)
                                        Case TILE_TYPE_DOOR
                                            Call DrawText(TexthDC, (x * PIC_X) + TempX, (y * PIC_Y) + TempY, "D", QBColor(BLACK), GameFont)
                                        Case TILE_TYPE_NOTICE
                                            Call DrawText(TexthDC, (x * PIC_X) + TempX, (y * PIC_Y) + TempY, "N", QBColor(BRIGHTGREEN), GameFont)
                                        Case TILE_TYPE_CHEST
                                            Call DrawText(TexthDC, (x * PIC_X) + TempX, (y * PIC_Y) + TempY, "C", QBColor(BROWN), GameFont)
                                        Case TILE_TYPE_CLASS_CHANGE
                                            Call DrawText(TexthDC, (x * PIC_X) + TempX, (y * PIC_Y) + TempY, "CG", QBColor(WHITE), GameFont)
                                        Case TILE_TYPE_SCRIPTED
                                            Call DrawText(TexthDC, (x * PIC_X) + TempX, (y * PIC_Y) + TempY, "SC", QBColor(YELLOW), GameFont)
                                        Case TILE_TYPE_BANK
                                            Call DrawText(TexthDC, (x * PIC_X) + TempX, (y * PIC_Y) + TempY, "BANK", QBColor(BRIGHTRED), GameFont)
                                        Case TILE_TYPE_GUILDBLOCK
                                            Call DrawText(TexthDC, (x * PIC_X) + TempX, (y * PIC_Y) + TempY, "GB", QBColor(MAGENTA), GameFont)
                                        Case TILE_TYPE_HOOKSHOT
                                            Call DrawText(TexthDC, (x * PIC_X) + TempX, (y * PIC_Y) + TempY, "GS", QBColor(WHITE), GameFont)
                                        Case TILE_TYPE_WALKTHRU
                                            Call DrawText(TexthDC, (x * PIC_X) + TempX, (y * PIC_Y) + TempY, "WT", QBColor(RED), GameFont)
                                        Case TILE_TYPE_ROOF
                                            Call DrawText(TexthDC, (x * PIC_X) + TempX, (y * PIC_Y) + TempY, "RF", QBColor(RED), GameFont)
                                        Case TILE_TYPE_ROOFBLOCK
                                            Call DrawText(TexthDC, (x * PIC_X) + TempX, (y * PIC_Y) + TempY, "RFB", QBColor(BRIGHTRED), GameFont)
                                        Case TILE_TYPE_ONCLICK
                                            Call DrawText(TexthDC, (x * PIC_X) + TempX, (y * PIC_Y) + TempY, "OC", QBColor(WHITE), GameFont)
                                        Case TILE_TYPE_LOWER_STAT
                                            Call DrawText(TexthDC, (x * PIC_X) + TempX, (y * PIC_Y) + TempY, "-S", QBColor(BRIGHTRED), GameFont)
                                        Case TILE_TYPE_SWITCH
                                            Call DrawText(TexthDC, (x * PIC_X) + TempX, (y * PIC_Y) + TempY, "Sw", QBColor(GREY), GameFont)
                                        Case TILE_TYPE_LVLBLOCK
                                            Call DrawText(TexthDC, (x * PIC_X) + TempX, (y * PIC_Y) + TempY, "LB", QBColor(GREEN), GameFont)
                                        Case TILE_TYPE_QUESTIONBLOCK
                                            Call DrawText(TexthDC, (x * PIC_X) + TempX, (y * PIC_Y) + TempY, "?", QBColor(YELLOW), GameFont)
                                        Case TILE_TYPE_DRILL
                                            Call DrawText(TexthDC, (x * PIC_X) + TempX, (y * PIC_Y) + TempY, "DR", QBColor(WHITE), GameFont)
                                        Case TILE_TYPE_JUMPBLOCK
                                            Call DrawText(TexthDC, (x * PIC_X) + TempX, (y * PIC_Y) + TempY, "JB", QBColor(BRIGHTBLUE), GameFont)
                                        Case TILE_TYPE_DODGEBILL
                                            Call DrawText(TexthDC, (x * PIC_X) + TempX, (y * PIC_Y) + TempY, "D", QBColor(BLUE), GameFont)
                                        Case TILE_TYPE_HAMMERBARRAGE
                                            Call DrawText(TexthDC, (x * PIC_X) + TempX, (y * PIC_Y) + TempY, "HB", QBColor(GREEN), GameFont)
                                        Case TILE_TYPE_JUGEMSCLOUD
                                            Call DrawText(TexthDC, (x * PIC_X) + TempX, (y * PIC_Y) + TempY, "JC", QBColor(YELLOW), GameFont)
                                        Case TILE_TYPE_SIMULBLOCK
                                            Call DrawText(TexthDC, (x * PIC_X) + TempX, (y * PIC_Y) + TempY, "SB", QBColor(BRIGHTBLUE), GameFont)
                                        Case TILE_TYPE_BEAN
                                            Call DrawText(TexthDC, (x * PIC_X) + TempX, (y * PIC_Y) + TempY, "BE", QBColor(YELLOW), GameFont)
                                    End Select
                                
                                    If .light > 0 Then
                                        Call DrawText(TexthDC, (x * PIC_X) + TempX, (y * PIC_Y) + TempY, "L", QBColor(YELLOW), GameFont)
                                    End If
                                End With
                            Next x
                        Next y
                    End If

                    ' draw FPS
                    If BFPS Then
                        Call DrawText(TexthDC, sx + 2, sx, "FPS: " & GameFPS, QBColor(YELLOW), GameFont)
                    End If

                    ' draw cursor and player X and Y locations
                    If BLoc Then
                        Call DrawText(TexthDC, 450 + sx, 0 + sx, "Cursor (X: " & CurX & "; Y: " & CurY & ")", QBColor(YELLOW), GameFont)
                        Call DrawText(TexthDC, 450 + sx, 15 + sx, "Location (X: " & GetPlayerX(MyIndex) & "; Y: " & GetPlayerY(MyIndex) & ")", QBColor(YELLOW), GameFont)
                        Call DrawText(TexthDC, 450 + sx, 30 + sx, "Map #" & GetPlayerMap(MyIndex), QBColor(YELLOW), GameFont)
                    End If

                    For i = 1 To MAX_BLT_LINE
                        If BattlePMsg(i).Index > 0 Then
                            If BattlePMsg(i).Time + 7000 > GetTickCount Then
                                Call DrawText(TexthDC, 1 + sx, BattlePMsg(i).y + frmMirage.picScreen.Height - 15 + sx, Trim$(BattlePMsg(i).Msg), QBColor(BattlePMsg(i).Color), GameFont)
                            Else
                                BattlePMsg(i).Done = 0
                            End If
                        End If

                        If BattleMMsg(i).Index > 0 Then
                            If BattleMMsg(i).Time + 7000 > GetTickCount Then
                                Call DrawText(TexthDC, (frmMirage.picScreen.Width - (Len(BattleMMsg(i).Msg) * 8)) + sx, BattleMMsg(i).y + frmMirage.picScreen.Height - 15 + sx, Trim$(BattleMMsg(i).Msg), QBColor(BattleMMsg(i).Color), GameFont)
                            Else
                                BattleMMsg(i).Done = 0
                            End If
                        End If
                    Next i
                        
                End If
                
            Else
                ' Lock the backbuffer so we can draw text
                TexthDC = DD_BackBuffer.GetDC
                
                ' Show player that a new map is loading
                Call DrawText(TexthDC, 36, 36, "Loading map...", QBColor(BRIGHTCYAN), GameFont)
            End If

            ' Release DC
            Call DD_BackBuffer.ReleaseDC(TexthDC)
            
            TempX = (MAX_MAPX + 1) * PIC_X
            TempY = (MAX_MAPY + 1) * PIC_Y
            
            ' Get the rect for the back buffer to blit from
            rec.Top = 0
            rec.Bottom = TempY
            rec.Left = 0
            rec.Right = TempX

            ' Get the rect to blit to
            Call DX.GetWindowRect(frmMirage.picScreen.hWnd, rec_pos)
            rec_pos.Bottom = rec_pos.Top - sx + TempY
            rec_pos.Right = rec_pos.Left - sx + TempX
            rec_pos.Top = rec_pos.Bottom - TempY
            rec_pos.Left = rec_pos.Right - TempX

            ' Blit the backbuffer
            Call DD_PrimarySurf.Blt(rec_pos, DD_BackBuffer, rec, DDBLT_WAIT)

            ' Check if player is trying to move
            Call CheckMovement

            ' Check to see if player is trying to attack
            Call CheckAttack

            ' Process player movements (actually move them)
            For i = 1 To MAX_PLAYERS
                If IsPlaying(i) Then
                    Call ProcessMovement(i)
                End If
            Next i

            ' Process npc movements (actually move them)
            For i = 1 To MAX_MAP_NPCS
                If Map(GetPlayerMap(MyIndex)).Npc(i) > 0 Then
                    Call ProcessNpcMovement(i)
                End If
            Next i

        End If

        ' Change map animation every 250 milliseconds
        If GetTickCount > MapAnimTimer + 250 Then
            If MapAnim = 0 Then
                MapAnim = 1
            Else
                MapAnim = 0
            End If
            MapAnimTimer = GetTickCount
        End If
        
        ' Lock fps
        Do While GetTickCount < Tick + 31
            DoEvents
            Sleep 1
        Loop

        ' Calculate fps
        If GetTickCount > TickFPS + 1000 Then
            GameFPS = FPS
            TickFPS = GetTickCount
            FPS = 0
        Else
            FPS = FPS + 1
        End If

        DoEvents
    Loop

    frmSendGetData.Visible = True

    Call SetStatus("Exiting game...")

    ' MsgBox "Connection lost!"

    ' Shutdown the game
    Call GameDestroy

    Exit Sub
End Sub

' Closes the game client.
Sub GameDestroy()
    ' Unloads all TCP-related things.
    Call TcpDestroy

    ' Unloads all DirectX objects.
    Call DestroyDirectX

    ' Unloads the BGM in memory (soon-to-be obsolete).
    Call StopBGM

    ' Closes the VB6 application.
    End
End Sub

Sub BltTile(ByVal x As Long, ByVal y As Long)
    Dim Ground As Long
    Dim Mask1 As Long
    Dim Anim1 As Long
    Dim Mask2 As Long
    Dim Anim2 As Long
    Dim GroundTileSet As Byte
    Dim Mask1TileSet As Byte
    Dim Anim1TileSet As Byte
    Dim Mask2TileSet As Byte
    Dim Anim2TileSet As Byte

    Ground = Map(GetPlayerMap(MyIndex)).Tile(x, y).Ground
    Mask1 = Map(GetPlayerMap(MyIndex)).Tile(x, y).Mask
    Anim1 = Map(GetPlayerMap(MyIndex)).Tile(x, y).Anim
    Mask2 = Map(GetPlayerMap(MyIndex)).Tile(x, y).Mask2
    Anim2 = Map(GetPlayerMap(MyIndex)).Tile(x, y).M2Anim

    GroundTileSet = Map(GetPlayerMap(MyIndex)).Tile(x, y).GroundSet
    Mask1TileSet = Map(GetPlayerMap(MyIndex)).Tile(x, y).MaskSet
    Anim1TileSet = Map(GetPlayerMap(MyIndex)).Tile(x, y).AnimSet
    Mask2TileSet = Map(GetPlayerMap(MyIndex)).Tile(x, y).Mask2Set
    Anim2TileSet = Map(GetPlayerMap(MyIndex)).Tile(x, y).M2AnimSet

    If TileFile(GroundTileSet) = 0 Then
        Exit Sub
    End If

    rec.Top = (Ground \ TilesInSheets) * PIC_Y
    rec.Bottom = rec.Top + PIC_Y
    rec.Left = (Ground - (Ground \ TilesInSheets) * TilesInSheets) * PIC_X
    rec.Right = rec.Left + PIC_X

    Call DD_BackBuffer.BltFast((x - NewPlayerX) * PIC_X - NewXOffset + sx, (y - NewPlayerY) * PIC_Y - NewYOffset + sx, DD_TileSurf(GroundTileSet), rec, DDBLTFAST_WAIT)

    If MapAnim = 0 Or Anim1 = 0 Then
        If Mask1 > 0 Then
            If TileFile(Mask1TileSet) = 0 Then
                Exit Sub
            End If

            If TempTile(x, y).DoorOpen = No Then
                rec.Top = (Mask1 \ TilesInSheets) * PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                rec.Left = (Mask1 - (Mask1 \ TilesInSheets) * TilesInSheets) * PIC_X
                rec.Right = rec.Left + PIC_X
                
                Call DD_BackBuffer.BltFast((x - NewPlayerX) * PIC_X - NewXOffset + sx, (y - NewPlayerY) * PIC_Y - NewYOffset + sx, DD_TileSurf(Mask1TileSet), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            End If
        End If
    Else
        If Anim1 > 0 Then
            If TileFile(Anim1TileSet) = 0 Then
                Exit Sub
            End If

            rec.Top = (Anim1 \ TilesInSheets) * PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            rec.Left = (Anim1 - (Anim1 \ TilesInSheets) * TilesInSheets) * PIC_X
            rec.Right = rec.Left + PIC_X

            Call DD_BackBuffer.BltFast((x - NewPlayerX) * PIC_X - NewXOffset + sx, (y - NewPlayerY) * PIC_Y - NewYOffset + sx, DD_TileSurf(Anim1TileSet), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    End If

    If MapAnim = 0 Or Anim2 = 0 Then
        If Mask2 > 0 Then
            If TileFile(Mask2TileSet) = 0 Then
                Exit Sub
            End If

            rec.Top = (Mask2 \ TilesInSheets) * PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            rec.Left = (Mask2 - (Mask2 \ TilesInSheets) * TilesInSheets) * PIC_X
            rec.Right = rec.Left + PIC_X

            Call DD_BackBuffer.BltFast((x - NewPlayerX) * PIC_X - NewXOffset + sx, (y - NewPlayerY) * PIC_Y - NewYOffset + sx, DD_TileSurf(Mask2TileSet), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    Else
        If Anim2 > 0 Then
            If TileFile(Anim2TileSet) = 0 Then
                Exit Sub
            End If

            rec.Top = (Anim2 \ TilesInSheets) * PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            rec.Left = (Anim2 - (Anim2 \ TilesInSheets) * TilesInSheets) * PIC_X
            rec.Right = rec.Left + PIC_X

            Call DD_BackBuffer.BltFast((x - NewPlayerX) * PIC_X - NewXOffset + sx, (y - NewPlayerY) * PIC_Y - NewYOffset + sx, DD_TileSurf(Anim2TileSet), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    End If
End Sub

Sub BltBattleIcons()
    Dim X1 As Long, x2 As Long, X3 As Long, X4 As Long, Y1 As Long, y2 As Long, Y3 As Long, Y4 As Long
    
    X1 = NewX + sx + 60 + (PIC_X \ 2)
    Y1 = NewY + sx + (PIC_Y \ 2) - 80
    x2 = X1 - 80
    y2 = Y1 - 20
    X3 = x2 - 80
    Y3 = Y1
    X4 = X1 + 50
    Y4 = Y1 + 65
    
    Select Case ButtonHighlighted
        Case 0
            Call BltFadedAttackImage(X1, Y1)
            Call BltFadedItemImage(x2, y2)
            Call BltRunImage(X3, Y3)
            Call BltFadedSpecialImage(X4, Y4)
        Case 1
            Call BltFadedAttackImage(X1, Y1)
            Call BltItemImage(x2, y2)
            Call BltFadedRunImage(X3, Y3)
            Call BltFadedSpecialImage(X4, Y4)
            Exit Sub
        Case 2
            Call BltAttackImage(X1, Y1)
            Call BltFadedItemImage(x2, y2)
            Call BltFadedRunImage(X3, Y3)
            Call BltFadedSpecialImage(X4, Y4)
        Case 3
            Call BltFadedAttackImage(X1, Y1)
            Call BltFadedItemImage(x2, y2)
            Call BltFadedRunImage(X3, Y3)
            Call BltSpecialImage(X4, Y4)
    End Select
End Sub

Sub BltItem(ByVal ItemNum As Long)
    rec.Top = (Item(MapItem(ItemNum).num).Pic \ 6) * PIC_Y
    rec.Bottom = rec.Top + PIC_Y
    rec.Left = (Item(MapItem(ItemNum).num).Pic - (Item(MapItem(ItemNum).num).Pic \ 6) * 6) * PIC_X
    rec.Right = rec.Left + PIC_X

    Call DD_BackBuffer.BltFast((MapItem(ItemNum).x - NewPlayerX) * PIC_X + sx - NewXOffset, (MapItem(ItemNum).y - NewPlayerY) * PIC_Y + sx - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub

Sub BltFringeTile(ByVal x As Long, ByVal y As Long)
    Dim Fringe As Long
    Dim FAnim As Long
    Dim FringeTileSet As Byte
    Dim FAnimTileSet As Byte

    Fringe = Map(GetPlayerMap(MyIndex)).Tile(x, y).Fringe
    FAnim = Map(GetPlayerMap(MyIndex)).Tile(x, y).FAnim

    FringeTileSet = Map(GetPlayerMap(MyIndex)).Tile(x, y).FringeSet
    FAnimTileSet = Map(GetPlayerMap(MyIndex)).Tile(x, y).FAnimSet

    If MapAnim = 0 Or FAnim = 0 Then
        If Fringe > 0 Then
            If TileFile(FringeTileSet) = 0 Then
                Exit Sub
            End If

            rec.Top = (Fringe \ TilesInSheets) * PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            rec.Left = (Fringe - (Fringe \ TilesInSheets) * TilesInSheets) * PIC_X
            rec.Right = rec.Left + PIC_X

            Call DD_BackBuffer.BltFast((x - NewPlayerX) * PIC_X + sx - NewXOffset, (y - NewPlayerY) * PIC_Y + sx - NewYOffset, DD_TileSurf(FringeTileSet), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    Else
        If FAnim > 0 Then
            If TileFile(FAnimTileSet) = 0 Then
                Exit Sub
            End If

            rec.Top = (FAnim \ TilesInSheets) * PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            rec.Left = (FAnim - (FAnim \ TilesInSheets) * TilesInSheets) * PIC_X
            rec.Right = rec.Left + PIC_X

            Call DD_BackBuffer.BltFast((x - NewPlayerX) * PIC_X + sx - NewXOffset, (y - NewPlayerY) * PIC_Y + sx - NewYOffset, DD_TileSurf(FAnimTileSet), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    End If
End Sub

Sub BltFringe2Tile(ByVal x As Integer, ByVal y As Integer)
    Dim Fringe2 As Long
    Dim F2Anim As Long
    Dim Fringe2TileSet As Byte
    Dim F2AnimTileSet As Byte

    Fringe2 = Map(GetPlayerMap(MyIndex)).Tile(x, y).Fringe2
    F2Anim = Map(GetPlayerMap(MyIndex)).Tile(x, y).F2Anim

    Fringe2TileSet = Map(GetPlayerMap(MyIndex)).Tile(x, y).Fringe2Set
    F2AnimTileSet = Map(GetPlayerMap(MyIndex)).Tile(x, y).F2AnimSet

    If MapAnim = 0 Or F2Anim = 0 Then
        If Fringe2 > 0 Then
            If TileFile(Fringe2TileSet) = 0 Then
                Exit Sub
            End If

            rec.Top = (Fringe2 \ TilesInSheets) * PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            rec.Left = (Fringe2 - (Fringe2 \ TilesInSheets) * TilesInSheets) * PIC_X
            rec.Right = rec.Left + PIC_X

            Call DD_BackBuffer.BltFast((x - NewPlayerX) * PIC_X + sx - NewXOffset, (y - NewPlayerY) * PIC_Y + sx - NewYOffset, DD_TileSurf(Fringe2TileSet), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    Else
        If F2Anim > 0 Then
            If TileFile(F2AnimTileSet) = 0 Then
                Exit Sub
            End If

            rec.Top = (F2Anim \ TilesInSheets) * PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            rec.Left = (F2Anim - (F2Anim \ TilesInSheets) * TilesInSheets) * PIC_X
            rec.Right = rec.Left + PIC_X

            Call DD_BackBuffer.BltFast((x - NewPlayerX) * PIC_X + sx - NewXOffset, (y - NewPlayerY) * PIC_Y + sx - NewYOffset, DD_TileSurf(F2AnimTileSet), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    End If
End Sub

Sub BltPlayer(ByVal Index As Long)
    Dim Anim As Byte
    Dim x As Long, y As Long, AttackSpeed As Long
    
    If CanBltPlayerGfx(Index) = False Then
        Exit Sub
    End If
    
    If Player(Index).Jumping = True Then
        Exit Sub
    End If
    
    ' Check attack speed
    AttackSpeed = GetPlayerAttackSpeed(Index)
   
    ' Check for animation
    Anim = 1
    
    If Player(Index).Attacking = 0 Then
        Select Case GetPlayerDir(Index)
            Case DIR_UP
                If (Player(Index).yOffset > 8) Then Anim = Player(Index).Step
            Case DIR_DOWN
                If (Player(Index).yOffset < -8) Then Anim = Player(Index).Step
            Case DIR_LEFT
                If (Player(Index).xOffset > 8) Then Anim = Player(Index).Step
            Case DIR_RIGHT
                If (Player(Index).xOffset < -8) Then Anim = Player(Index).Step
        End Select
    Else
        If Player(Index).AttackTimer + AttackSpeed > GetTickCount Then
            Anim = 2
        End If
    End If

    ' Check to see if we want to stop making him attack
    If Player(Index).AttackTimer + AttackSpeed < GetTickCount Then
        Player(Index).Attacking = 0
    End If

    For x = 1 To 7
        Player(Index).Equipment(x).num = GetPlayerEquipSlotNum(Index, x)
    Next x
   
    ' Start blitting out player
    rec.Left = (GetPlayerDir(Index) * 3 + Anim) * 32
    rec.Right = rec.Left + 32

    x = GetPlayerX(Index) * PIC_X + sx + Player(Index).xOffset
    y = GetPlayerY(Index) * PIC_Y + sx + Player(Index).yOffset - 32
    
    x = x - (NewPlayerX * PIC_X) - NewXOffset
    y = y - (NewPlayerY * PIC_Y) - NewYOffset
    
    ' Don't blit out any equipment if the player is in Hide n' Sneak
    If IsPlayingHideNSneak = True Then
        ' BLIT SPRITE
        rec.Top = GetPlayerSprite(Index) * 64
        rec.Bottom = rec.Top + 64
        Call DD_BackBuffer.BltFast(x, y, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            
        Exit Sub
    End If
    
    ' IF DIR = UP
    If GetPlayerDir(Index) = DIR_UP Then
        ' BLIT SHIELD IF DIR = UP
        If Player(Index).Equipment(4).num > 0 Then
            rec.Top = Item(Player(Index).Equipment(4).num).Pic * 64
            rec.Bottom = rec.Top + 64
            Call DD_BackBuffer.BltFast(x, y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
        ' BLIT WEAPON IF DIR = UP
        If Player(Index).Equipment(1).num > 0 And Player(Index).Equipment(1).num <> 124 Then
            rec.Top = Item(Player(Index).Equipment(1).num).Pic * 64
            rec.Bottom = rec.Top + 64
            Call DD_BackBuffer.BltFast(x, y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
        ' BLIT NECKLACE IF DIR = UP
        If Player(Index).Equipment(7).num > 0 Then
            rec.Top = Item(Player(Index).Equipment(7).num).Pic * 64
            rec.Bottom = rec.Top + 64
            Call DD_BackBuffer.BltFast(x, y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    End If

    ' BLIT SPRITE
    rec.Top = GetPlayerSprite(Index) * 64
    rec.Bottom = rec.Top + 64
    Call DD_BackBuffer.BltFast(x, y, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                
    If GetPlayerDir(Index) = DIR_UP Then
        If Player(Index).Equipment(1).num = 124 Then
            rec.Top = Item(Player(Index).Equipment(1).num).Pic * 64
            rec.Bottom = rec.Top + 64
            Call DD_BackBuffer.BltFast(x, y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    End If
                
    ' BLIT LEGS
    If Player(Index).Equipment(5).num > 0 Then
        rec.Top = Item(Player(Index).Equipment(5).num).Pic * 64
        rec.Bottom = rec.Top + 64
        Call DD_BackBuffer.BltFast(x, y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    End If
    ' BLIT ARMOR
    If Player(Index).Equipment(2).num > 0 Then
        rec.Top = Item(Player(Index).Equipment(2).num).Pic * 64
        rec.Bottom = rec.Top + 64
        Call DD_BackBuffer.BltFast(x, y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    End If
    ' BLIT HELMET
    If Player(Index).Equipment(3).num > 0 Then
        rec.Top = Item(Player(Index).Equipment(3).num).Pic * 64
        rec.Bottom = rec.Top + 64
        Call DD_BackBuffer.BltFast(x, y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    End If

    ' IF DIR <> UP
    If GetPlayerDir(Index) <> DIR_UP Then
        ' BLIT SHIELD IF DIR <> UP
        If Player(Index).Equipment(4).num > 0 Then
            rec.Top = Item(Player(Index).Equipment(4).num).Pic * 64
            rec.Bottom = rec.Top + 64
            Call DD_BackBuffer.BltFast(x, y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
        ' BLIT NECKLACE IF DIR <> UP
        If Player(Index).Equipment(7).num > 0 Then
            rec.Top = Item(Player(Index).Equipment(7).num).Pic * 64
            rec.Bottom = rec.Top + 64
            Call DD_BackBuffer.BltFast(x, y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
        ' BLIT WEAPON IF DIR <> UP
        If Player(Index).Equipment(1).num > 0 Then
            rec.Top = Item(Player(Index).Equipment(1).num).Pic * 64
            rec.Bottom = rec.Top + 64
            Call DD_BackBuffer.BltFast(x, y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    End If
End Sub

Sub BltJump(ByVal Index As Long)
    Dim Anim As Byte, AnimNum As Byte, JumpAnim As Byte
    Dim x As Long, y As Long
    
    If CanBltPlayerGfx(Index) = False Then
        Exit Sub
    End If
    
    ' Make sure the player is jumping
    If Player(Index).Jumping = False Then
        Exit Sub
    End If
   
    ' Set the animation to attacking
    Anim = 2

    For x = 1 To 7
        Player(Index).Equipment(x).num = GetPlayerEquipSlotNum(Index, x)
    Next x
   
    ' Determine which animation to use
    If GetTickCount > (Player(Index).JumpTime + 20) Then
        ' Set Jump Time
        Player(Index).JumpTime = GetTickCount
    
        Player(Index).TempJumpAnim = Player(Index).TempJumpAnim + 1
        
        ' Start jumping up
        If Player(Index).TempJumpAnim <= 15 Then
            Player(Index).JumpAnim = Player(Index).JumpAnim + 1
        ' Start falling down after 15 animations
        Else
            Player(Index).JumpAnim = Player(Index).JumpAnim - 1
            Player(Index).JumpDir = 1
        End If
    End If
    
    rec.Left = (GetPlayerDir(Index) * 3 + Anim) * 32
    rec.Right = rec.Left + 32
    
    x = GetPlayerX(Index) * PIC_X + sx + Player(Index).xOffset
    y = GetPlayerY(Index) * PIC_Y + sx + Player(Index).yOffset - 32 - (Player(Index).JumpAnim * 3)
                    
    x = x - (NewPlayerX * PIC_X) - NewXOffset
    y = y - (NewPlayerY * PIC_Y) - NewYOffset
    
    ' Don't blit out any equipment if the player is in Hide n' Sneak
    If IsPlayingHideNSneak = True Then
        ' BLIT SPRITE
        rec.Top = GetPlayerSprite(Index) * 64
        rec.Bottom = rec.Top + 64
        Call DD_BackBuffer.BltFast(x, y, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        
        ' End the jump
        If Player(Index).JumpDir = 1 And Player(Index).TempJumpAnim >= 30 Then
            Player(Index).Jumping = False
                
        ' Send that the jump ended
            If Index = MyIndex Then
                Call SendData(CPackets.Cendjump & SEP_CHAR & Index & END_CHAR)
            End If
        End If
        
        Exit Sub
    End If
    
    ' IF DIR = UP
    If GetPlayerDir(Index) = DIR_UP Then
        ' BLIT SHIELD IF DIR = UP
        If Player(Index).Equipment(4).num > 0 Then
            rec.Top = Item(Player(Index).Equipment(4).num).Pic * 64
            rec.Bottom = rec.Top + 64
            Call DD_BackBuffer.BltFast(x, y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
        ' BLIT WEAPON IF DIR = UP
        If Player(Index).Equipment(1).num > 0 And Player(Index).Equipment(1).num <> 124 Then
            rec.Top = Item(Player(Index).Equipment(1).num).Pic * 64
            rec.Bottom = rec.Top + 64
            Call DD_BackBuffer.BltFast(x, y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
        ' BLIT NECKLACE IF DIR = UP
        If Player(Index).Equipment(7).num > 0 Then
            rec.Top = Item(Player(Index).Equipment(7).num).Pic * 64
            rec.Bottom = rec.Top + 64
            Call DD_BackBuffer.BltFast(x, y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    End If
    
    ' BLIT SPRITE
    rec.Top = GetPlayerSprite(Index) * 64
    rec.Bottom = rec.Top + 64
    Call DD_BackBuffer.BltFast(x, y, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    
    If GetPlayerDir(Index) = DIR_UP Then
        If Player(Index).Equipment(1).num = 124 Then
            rec.Top = Item(Player(Index).Equipment(1).num).Pic * 64
            rec.Bottom = rec.Top + 64
            Call DD_BackBuffer.BltFast(x, y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    End If
                    
    ' BLIT LEGS
    If Player(Index).Equipment(5).num > 0 Then
        rec.Top = Item(Player(Index).Equipment(5).num).Pic * 64
        rec.Bottom = rec.Top + 64
        Call DD_BackBuffer.BltFast(x, y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    End If
    ' BLIT ARMOR
    If Player(Index).Equipment(2).num > 0 Then
        rec.Top = Item(Player(Index).Equipment(2).num).Pic * 64
        rec.Bottom = rec.Top + 64
        Call DD_BackBuffer.BltFast(x, y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    End If
    ' BLIT HELMET
    If Player(Index).Equipment(3).num > 0 Then
        rec.Top = Item(Player(Index).Equipment(3).num).Pic * 64
        rec.Bottom = rec.Top + 64
        Call DD_BackBuffer.BltFast(x, y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    End If
    
    ' IF DIR <> UP
    If GetPlayerDir(Index) <> DIR_UP Then
        ' BLIT SHIELD IF DIR <> UP
        If Player(Index).Equipment(4).num > 0 Then
            rec.Top = Item(Player(Index).Equipment(4).num).Pic * 64
            rec.Bottom = rec.Top + 64
            Call DD_BackBuffer.BltFast(x, y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
        ' BLIT NECKLACE IF DIR <> UP
        If Player(Index).Equipment(7).num > 0 Then
            rec.Top = Item(Player(Index).Equipment(7).num).Pic * 64
            rec.Bottom = rec.Top + 64
            Call DD_BackBuffer.BltFast(x, y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
        ' BLIT WEAPON IF DIR <> UP
        If Player(Index).Equipment(1).num > 0 Then
            rec.Top = Item(Player(Index).Equipment(1).num).Pic * 64
            rec.Bottom = rec.Top + 64
            Call DD_BackBuffer.BltFast(x, y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    End If
    
    ' End the jump
    If Player(Index).JumpDir = 1 And Player(Index).TempJumpAnim >= 30 Then
        Player(Index).Jumping = False
            
    ' Send that the jump ended
        If Index = MyIndex Then
            Call SendData(CPackets.Cendjump & SEP_CHAR & Index & END_CHAR)
        End If
    End If
End Sub

Sub BltVictory()
    If Player(MyIndex).BattleVictory = False Then
        Exit Sub
    End If
    
    If StartedVictoryAnim = True Then
        If GetTickCount <= (BattleVictoryTimer + 600) Then
            Exit Sub
        End If
    End If
    
    Dim srcRECT As RECT
    Dim x As Long, y As Long
        
    With srcRECT
        .Top = 64 * GetPlayerClass(MyIndex)
        .Left = PIC_X * BattleFrameCount
        .Right = .Left + PIC_X
        .Bottom = .Top + 64
    End With
        
    x = GetPlayerX(MyIndex) * PIC_X + sx + Player(MyIndex).xOffset
    y = GetPlayerY(MyIndex) * PIC_Y + sx + Player(MyIndex).yOffset - 32
    
    x = x - (NewPlayerX * PIC_X) - NewXOffset
    y = y - (NewPlayerY * PIC_Y) - NewYOffset
    
    If BattleFrameCount = 9 Then
        CanFinishBattle = True
        DisplayInfo = True
        Call BltVictoryImage((x - 161), (y - 36))
    End If
    
    Call DD_BackBuffer.BltFast(x, y, DD_VictoryAnimSurf, srcRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)

    If GetTickCount > (BattleVictoryTimer + 75) And BattleFrameCount < 9 Then
        StartedVictoryAnim = False
        BattleVictoryTimer = GetTickCount
        BattleFrameCount = BattleFrameCount + 1
    End If
End Sub

Sub BltMapNPCName(ByVal Index As Long)
    Dim TextX As Long, TextY As Long
    Dim NpcName As String
    
    If CanBltNpcGfx(Index) = False Then
        Exit Sub
    End If
    
    If Player(MyIndex).BattleVictory = True Then
        Exit Sub
    End If
    
    If Npc(MapNpc(Index).num).Level > 0 Then
        NpcName = Trim$(Npc(MapNpc(Index).num).Name) & " (Level " & Npc(MapNpc(Index).num).Level & ")"
    Else
        NpcName = Trim$(Npc(MapNpc(Index).num).Name)
    End If
    
    TextX = MapNpc(Index).x * PIC_X + sx + MapNpc(Index).xOffset + (PIC_X / 2) - ((Len(NpcName) / 2) * 8) - (NewPlayerX * PIC_X) - NewXOffset
    TextY = MapNpc(Index).y * PIC_Y + sx + MapNpc(Index).yOffset - (PIC_Y / 2) - (NewPlayerY * PIC_Y) - NewYOffset - 32
    
    Call DrawText(TexthDC, TextX, TextY, NpcName, QBColor(DARKGREY), GameFont)
End Sub

Sub BltNpc(ByVal MapNpcNum As Long)
    Dim Anim As Byte
    Dim x As Long, y As Long, modify As Long
    
    If CanBltNpcGfx(MapNpcNum) = False Then
        Exit Sub
    End If

    ' Check for animation
    Anim = 0
    If MapNpc(MapNpcNum).Attacking = 0 Then
        Select Case MapNpc(MapNpcNum).Dir
            Case DIR_UP
                If (MapNpc(MapNpcNum).yOffset < PIC_Y / 2) Then
                    Anim = 1
                End If
            Case DIR_DOWN
                If (MapNpc(MapNpcNum).yOffset < PIC_Y / 2 * -1) Then
                    Anim = 1
                End If
            Case DIR_LEFT
                If (MapNpc(MapNpcNum).xOffset < PIC_Y / 2) Then
                    Anim = 1
                End If
            Case DIR_RIGHT
                If (MapNpc(MapNpcNum).xOffset < PIC_Y / 2 * -1) Then
                    Anim = 1
                End If
        End Select
    Else
        If MapNpc(MapNpcNum).AttackTimer + 500 > GetTickCount Then
            Anim = 2
        End If
    End If

    ' Check to see if we want to stop making him attack
    If MapNpc(MapNpcNum).AttackTimer + 1000 < GetTickCount Then
        MapNpc(MapNpcNum).Attacking = 0
        MapNpc(MapNpcNum).AttackTimer = 0
    End If

    If Npc(MapNpc(MapNpcNum).num).Big = 1 Then
        rec.Top = Npc(MapNpc(MapNpcNum).num).Sprite * 64
        rec.Bottom = rec.Top + 64
        rec.Left = (MapNpc(MapNpcNum).Dir * 3 + Anim) * 64
        rec.Right = rec.Left + 64

        x = MapNpc(MapNpcNum).x * 32 + sx - 16 + MapNpc(MapNpcNum).xOffset
        y = MapNpc(MapNpcNum).y * 32 + sx + MapNpc(MapNpcNum).yOffset - 32

        If y < 0 Then
            modify = -y
            rec.Top = rec.Top + modify
            rec.Bottom = rec.Top + 32
            y = 0
        End If

        If x < 0 Then
            modify = -x
            rec.Left = rec.Left + modify
            rec.Right = rec.Left + 48
            x = 0
        End If

        If 32 + x >= (MAX_MAPX * 32) Then
            modify = x - (MAX_MAPX * 32)
            rec.Left = rec.Left + modify + 16
            rec.Right = rec.Left + 32 - modify
        End If

        Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_BigSpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    Else
        rec.Top = Npc(MapNpc(MapNpcNum).num).Sprite * 64
        rec.Bottom = rec.Top + 64
        rec.Left = (MapNpc(MapNpcNum).Dir * 3 + Anim) * PIC_X
        rec.Right = rec.Left + PIC_X

        x = MapNpc(MapNpcNum).x * PIC_X + sx + MapNpc(MapNpcNum).xOffset
        y = MapNpc(MapNpcNum).y * PIC_Y + sx + MapNpc(MapNpcNum).yOffset - 32

        Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    End If
End Sub

Sub BltPlayerName(ByVal Index As Long)
    Dim TextX As Long, TextY As Long, Color As Long
    
    If IsPlayingHideNSneak = True Then
        Exit Sub
    End If
    
    If CanBltPlayerGfx(Index) = False Then
        Exit Sub
    End If
    
    If Player(MyIndex).BattleVictory = True Then
        Exit Sub
    End If
    
    ' Check access level
    If GetPlayerPK(Index) = No Then
        Select Case GetPlayerAccess(Index)
            Case 0
                Color = QBColor(GREEN)
            Case 1
                Color = QBColor(BLACK)
            Case 2
                Color = QBColor(BRIGHTBLUE)
            Case 3
                Color = QBColor(BROWN)
            Case 4
                Color = QBColor(WHITE)
            Case 5
                Color = QBColor(YELLOW)
        End Select
    Else
        Color = QBColor(BRIGHTRED)
    End If
    
    Dim PlayerName As String
    
    If lvl >= 1 Then
        PlayerName = GetPlayerName(Index) & " (Level " & GetPlayerLevel(Index) & ")"
    Else
        PlayerName = GetPlayerName(Index)
    End If
    
    TextX = sx - ((Len(PlayerName) / 2) * 8) + 16
    
    ' Draw the player's name
    If Index = MyIndex Then
        TextX = TextX + NewX
        TextY = NewY + sx + 37
        
        Call DrawText(TexthDC, TextX, TextY, PlayerName, Color, GameFont)
    Else
        TextX = TextX + GetPlayerX(Index) * PIC_X + Player(Index).xOffset - (NewPlayerX * PIC_X) - NewXOffset
        TextY = GetPlayerY(Index) * PIC_Y + sx + Player(Index).yOffset - (NewPlayerY * PIC_Y) - NewYOffset + 37
        
        Call DrawText(TexthDC, TextX, TextY, PlayerName, Color, GameFont)
    End If
End Sub

Sub BltPlayerGuildName(ByVal Index As Long)
    Dim TextX As Long, TextY As Long, Color As Long
    
    If IsPlayingHideNSneak = True Then
        Exit Sub
    End If
    
    If CanBltPlayerGfx(Index) = False Then
        Exit Sub
    End If
    
    If Player(MyIndex).BattleVictory = True Then
        Exit Sub
    End If
    
    ' Check access level
    If GetPlayerPK(Index) = No Then
        Select Case GetPlayerGuildAccess(Index)
            Case 0
                Color = QBColor(RED)
            Case 1
                Color = QBColor(BRIGHTCYAN)
            Case 2
                Color = QBColor(PINK)
            Case 3
                Color = QBColor(BRIGHTGREEN)
            Case 4
                Color = QBColor(YELLOW)
        End Select
    Else
        Color = QBColor(BRIGHTRED)
    End If

    TextX = sx - ((Len(GetPlayerGuild(Index)) / 2) * 8) + 16

    ' Draw the players guild.
    If Index = MyIndex Then
        TextX = TextX + NewX
        TextY = NewY + sx + 52

        Call DrawText(TexthDC, TextX, TextY, GetPlayerGuild(MyIndex), Color, GameFont)
    Else
        TextX = TextX + GetPlayerX(Index) * PIC_X + Player(Index).xOffset - (NewPlayerX * PIC_X) - NewXOffset
        TextY = GetPlayerY(Index) * PIC_Y + sx + Player(Index).yOffset - (NewPlayerY * PIC_Y) - NewYOffset + 52

        Call DrawText(TexthDC, TextX, TextY, GetPlayerGuild(Index), Color, GameFont)
    End If
End Sub

Sub ProcessMovement(ByVal Index As Long)
    Dim MovementSpeed As Byte
    
    Select Case Player(Index).Moving
        Case MOVING_WALKING
            MovementSpeed = WALK_SPEED
        Case MOVING_RUNNING
            If GetPlayerSP(Index) = 0 Then
                Player(Index).Moving = MOVING_WALKING
                MovementSpeed = WALK_SPEED
            Else
                MovementSpeed = RUN_SPEED
            End If
    End Select
    
    Select Case GetPlayerDir(Index)
        Case DIR_UP
            Player(Index).yOffset = Player(Index).yOffset - MovementSpeed
            If Player(Index).yOffset <= 0 Then
                Player(Index).yOffset = 0
            End If
        Case DIR_DOWN
            Player(Index).yOffset = Player(Index).yOffset + MovementSpeed
            If Player(Index).yOffset >= 0 Then
                Player(Index).yOffset = 0
            End If
        Case DIR_LEFT
            Player(Index).xOffset = Player(Index).xOffset - MovementSpeed
            If Player(Index).xOffset <= 0 Then
                Player(Index).xOffset = 0
            End If
        Case DIR_RIGHT
            Player(Index).xOffset = Player(Index).xOffset + MovementSpeed
            If Player(Index).xOffset >= 0 Then
                Player(Index).xOffset = 0
            End If
    End Select
    
    ' Check if completed moving over to the next tile
    If Player(Index).Moving > 0 Then
        If Player(Index).Dir = DIR_RIGHT Or Player(Index).Dir = DIR_DOWN Then
            If (Player(Index).xOffset >= 0) And (Player(Index).yOffset >= 0) Then
                Player(Index).Moving = 0
                            
                If Player(Index).Step = 0 Then
                    Player(Index).Step = 2
                Else
                    Player(Index).Step = 0
                End If
            End If
        Else
            If (Player(Index).xOffset <= 0) And (Player(Index).yOffset <= 0) Then
                Player(Index).Moving = 0
                            
                If Player(Index).Step = 0 Then
                    Player(Index).Step = 2
                Else
                    Player(Index).Step = 0
                End If
            End If
        End If
    End If
End Sub

Sub ProcessNpcMovement(ByVal MapNpcNum As Long)
    ' Check if npc is walking, and if so process moving them over
    If MapNpc(MapNpcNum).Moving = MOVING_WALKING Then
        Select Case MapNpc(MapNpcNum).Dir
            Case DIR_UP
                MapNpc(MapNpcNum).yOffset = MapNpc(MapNpcNum).yOffset - WALK_SPEED
            Case DIR_DOWN
                MapNpc(MapNpcNum).yOffset = MapNpc(MapNpcNum).yOffset + WALK_SPEED
            Case DIR_LEFT
                MapNpc(MapNpcNum).xOffset = MapNpc(MapNpcNum).xOffset - WALK_SPEED
            Case DIR_RIGHT
                MapNpc(MapNpcNum).xOffset = MapNpc(MapNpcNum).xOffset + WALK_SPEED
        End Select

        ' Check if completed walking over to the next tile
        If (MapNpc(MapNpcNum).xOffset = 0) And (MapNpc(MapNpcNum).yOffset = 0) Then
            MapNpc(MapNpcNum).Moving = 0
        End If
    End If
End Sub

Sub HandleKeypresses(ByVal KeyAscii As Integer)
    Dim ChatText As String
    Dim Name As String
    Dim i As Long

    MyText = frmMirage.txtMyTextBox.Text
    
    If Player(MyIndex).BattleVictory = True And CanFinishBattle = True Then
        If KeyAscii = vbKeyReturn Then
            Player(MyIndex).BattleVictory = False
            BattleFrameCount = 0
            CanFinishBattle = False
            DisplayInfo = False
            
            Call SendFinishPlayerBattle
        End If
    End If
    
    ' Handle when the player presses the return key
    If (KeyAscii = vbKeyReturn) Then
        frmMirage.txtMyTextBox.Text = vbNullString
        If Player(MyIndex).y - 1 > -1 Then
            If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).Type = TILE_TYPE_SIGN And Player(MyIndex).Dir = DIR_UP And Len(MyText) < 1 Then
                If Trim$(Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).String1) <> vbNullString Then
                    Call AddText(Trim$(Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).String1), BLUE)
                End If
                If Trim$(Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).String2) <> vbNullString Then
                    Call AddText(Trim$(Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).String2), BLUE)
                End If
                If Trim$(Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).String3) <> vbNullString Then
                    Call AddText(Trim$(Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).String3), BLUE)
                End If
                Exit Sub
            End If
        End If
        ' Broadcast message
        If Mid$(MyText, 1, 1) = "'" Then
            ChatText = Mid$(MyText, 2, Len(MyText) - 1)
            If Len(Trim$(ChatText)) > 0 Then
                Call BroadcastMsg(ChatText)
            End If
            MyText = vbNullString
            Exit Sub
        End If

        ' Emote message
        If Mid$(MyText, 1, 1) = "-" Then
            ChatText = Mid$(MyText, 2, Len(MyText) - 1)
            If Len(Trim$(ChatText)) > 0 Then
                Call GroupMsg(ChatText)
            End If
            MyText = vbNullString
            Exit Sub
        End If

        ' Player message
        If Mid$(MyText, 1, 1) = "!" Then
            ChatText = Mid$(MyText, 2, Len(MyText) - 1)
            Name = vbNullString

            ' Get the desired player from the user text
            For i = 1 To Len(ChatText)
                If Mid$(ChatText, i, 1) <> " " Then
                    Name = Name & Mid$(ChatText, i, 1)
                Else
                    Exit For
                End If
            Next i

            ' Make sure they are actually sending something
            If Len(ChatText) - i > 0 Then
                ChatText = Mid$(ChatText, i + 1, Len(ChatText) - i)

                ' Send the message to the player
                Call PlayerMsg(ChatText, Name)
            Else
                Call AddText("Usage: !playername msghere", AlertColor)
            End If
            MyText = vbNullString
            Exit Sub
        End If

        ' // Commands //
        ' Verification User
        If GetPlayerAccess(MyIndex) >= 1 Then
            If LCase$(Mid$(MyText, 1, 5)) = "/info" Then
                ChatText = Mid$(MyText, 7, Len(MyText) - 5)
    
                If LenB(ChatText) <> 0 Then
                    Call SendData(CPackets.Cgetstats & SEP_CHAR & ChatText & END_CHAR)
                Else
                    Call AddText("Please enter a player's username.", BLACK)
                End If
    
                MyText = vbNullString
                Exit Sub
            End If
        End If
        
        ' Makes the Creators a normal player or an admin
        If GetPlayerName(MyIndex) = "Kimimaru" Or GetPlayerName(MyIndex) = "hydrakiller4000" Then
            If LCase$(Mid$(MyText, 1, 10)) = "/makeadmin" Then
                Call SendMakeAdmin
                MyText = vbNullString
                Exit Sub
            End If
        End If
    
        ' Checking fps
        If Mid$(MyText, 1, 4) = "/fps" Then
            If BFPS = False Then
                BFPS = True
            Else
                BFPS = False
            End If
            MyText = vbNullString
            Exit Sub
        End If

        ' Show inventory
        If LCase$(Mid$(MyText, 1, 4)) = "/inv" Then
            frmMirage.picInventory.Visible = True
            MyText = vbNullString
            Exit Sub
        End If

        ' Request stats
        If LCase$(Mid$(MyText, 1, 6)) = "/stats" Then
            Call SendData(CPackets.Cgetstats & SEP_CHAR & GetPlayerName(MyIndex) & END_CHAR)
            MyText = vbNullString
            Exit Sub
        End If

        ' Command help
        If LCase$(Mid$(MyText, 1, 5)) = "/help" Then
            Call SendData(CPackets.Chelp & END_CHAR)
            
            MyText = vbNullString
            Exit Sub
        End If
        
        ' Decline Chat
        If LCase$(Mid$(MyText, 1, 12)) = "/chatdecline" Then
            Call SendData(CPackets.Cdchat & END_CHAR)
            MyText = vbNullString
            Exit Sub
        End If

        ' Accept Chat
        If LCase$(Mid$(MyText, 1, 5)) = "/chat" Then
            Call SendData(CPackets.Cachat & END_CHAR)
            MyText = vbNullString
            Exit Sub
        End If

        If LCase$(Mid$(MyText, 1, 6)) = "/trade" Then
            ' Make sure they are actually sending something
            If Len(MyText) > 7 Then
                ChatText = Mid$(MyText, 8, Len(MyText) - 7)
                Call SendTradeRequest(ChatText)
            Else
                Call AddText("Usage: /trade playernamehere", AlertColor)
            End If
            MyText = vbNullString
            Exit Sub
        End If

        ' Accept Trade
        If LCase$(Mid$(MyText, 1, 7)) = "/accept" Then
            Call SendAcceptTrade
            MyText = vbNullString
            Exit Sub
        End If

        ' Decline Trade
        If LCase$(Mid$(MyText, 1, 8)) = "/decline" Then
            Call SendDeclineTrade
            MyText = vbNullString
            Exit Sub
        End If

        ' Party request
        If LCase$(Mid$(MyText, 1, 6)) = "/party" And LCase$(Mid$(MyText, 1, 13)) <> "/partydecline" Then
            ' Make sure the player is actually sending something
            If Len(MyText) > 9 Then
                ChatText = Mid$(MyText, 8, Len(MyText) - 7)
                Call SendPartyRequest(ChatText)
            Else
                Call AddText("Usage: /party (username)", AlertColor)
            End If
            
            MyText = vbNullString
            Exit Sub
        End If

        ' Join party
        If LCase$(Mid$(MyText, 1, 5)) = "/join" Then
            Call SendJoinParty
            
            MyText = vbNullString
            Exit Sub
        End If
        
        ' Decline party request
        If LCase$(Mid$(MyText, 1, 13)) = "/partydecline" Then
            Call SendDeclineParty
            
            MyText = vbNullString
            Exit Sub
        End If
            
        ' Leave party
        If LCase$(Mid$(MyText, 1, 6)) = "/leave" Then
            Call SendLeaveParty
            
            MyText = vbNullString
            Exit Sub
        End If

        ' // Moniter Admin Commands //
        If GetPlayerAccess(MyIndex) > 1 Then
            ' weather command
            If LCase$(Mid$(MyText, 1, 8)) = "/weather" Then
                If Len(MyText) > 8 Then
                    MyText = Mid$(MyText, 9, Len(MyText) - 8)
                    If IsNumeric(MyText) = True Then
                        Call SendData(CPackets.Cweather & SEP_CHAR & Val(MyText) & END_CHAR)
                    Else
                        If Trim$(LCase$(MyText)) = "none" Then
                            i = 0
                        End If
                        If Trim$(LCase$(MyText)) = "rain" Then
                            i = 1
                        End If
                        If Trim$(LCase$(MyText)) = "snow" Then
                            i = 2
                        End If
                        If Trim$(LCase$(MyText)) = "thunder" Then
                            i = 3
                        End If
                        Call SendData(CPackets.Cweather & SEP_CHAR & i & END_CHAR)
                    End If
                End If
                MyText = vbNullString
                Exit Sub
            End If
            
            ' Warping players to a map of their choice
            If LCase$(Mid$(MyText, 1, 3)) = "/go" And LCase$(Mid$(MyText, 1, 5)) <> "/goto" Then
               If Len(MyText) > 4 Then
                 MyText = Val(Mid$(MyText, 5, Len(MyText) - 4))
                  If MyText > 0 And MyText <= MAX_MAPS Then
                    Call WarpTo(MyText, GetPlayerX(MyIndex), GetPlayerY(MyIndex))
                  Else
                    Call AddText("Invalid map number.", BRIGHTRED)
                  End If
                End If
                MyText = vbNullString
                Exit Sub
            End If
            
            ' Warping players to other players
            If LCase$(Mid$(MyText, 1, 5)) = "/goto" Then
               If Len(MyText) > 6 Then
                 MyText = Mid$(MyText, 7, Len(MyText) - 6)
                  If Len(MyText) > 0 Then
                    Call WarpMeTo(MyText)
                  Else
                    Call AddText("You didn't enter a player's username!", BRIGHTRED)
                  End If
                End If
                MyText = vbNullString
                Exit Sub
            End If
            
            ' Warping players to developers and admins
            If LCase$(Mid$(MyText, 1, 5)) = "/call" Then
                If Len(MyText) > 6 Then
                    MyText = Mid$(MyText, 7, Len(MyText) - 6)
                    
                    If Len(MyText) > 0 Then
                        Call WarpToMe(MyText)
                    Else
                        Call AddText("You didn't enter a player's username!", BRIGHTRED)
                    End If
                End If
                
                MyText = vbNullString
                Exit Sub
            End If
        End If

        If GetPlayerAccess(MyIndex) >= 1 Then
            ' Kicking a player
            If LCase$(Mid$(MyText, 1, 5)) = "/kick" Then
                If Len(MyText) > 6 Then
                    MyText = Mid$(MyText, 7, Len(MyText) - 6)
                    Call SendKick(MyText)
                End If
                MyText = vbNullString
                Exit Sub
            End If
        
            ' Global Message
            If Mid$(MyText, 1, 1) = "'" Then
                ChatText = Mid$(MyText, 2, Len(MyText) - 1)
                If Len(Trim$(ChatText)) > 0 Then
                    Call GlobalMsg(ChatText)
                End If
                MyText = vbNullString
                Exit Sub
            End If

            ' Admin Message
            If Mid$(MyText, 1, 1) = "=" Then
                ChatText = Mid$(MyText, 2, Len(MyText) - 1)
                If Len(Trim$(ChatText)) > 0 Then
                    Call AdminMsg(ChatText)
                End If
                MyText = vbNullString
                Exit Sub
            End If
        End If

        ' // Mapper Admin Commands //
        If GetPlayerAccess(MyIndex) >= 2 Then
            ' Location
            If Mid$(MyText, 1, 4) = "/loc" Then
                If BLoc = False Then
                    BLoc = True
                Else
                    BLoc = False
                End If
                MyText = vbNullString
                Exit Sub
            End If

            ' Map Editor
            If LCase$(Mid$(MyText, 1, 10)) = "/mapeditor" Then
                Call SendRequestEditMap
                MyText = vbNullString
                Exit Sub
            End If

            ' Map report
            If LCase$(Mid$(MyText, 1, 10)) = "/mapreport" Then
                Call SendData(CPackets.Cmapreport & END_CHAR)
                MyText = vbNullString
                Exit Sub
            End If
        End If
        
            ' Setting sprite
        If GetPlayerAccess(MyIndex) = 5 Then
            If LCase$(Mid$(MyText, 1, 10)) = "/setsprite" Then
                If Len(MyText) > 11 Then
                    ' Get sprite #
                    MyText = Mid$(MyText, 12, Len(MyText) - 11)

                    Call SendSetPlayerSprite(GetPlayerName(MyIndex), Val(MyText))
                End If
                MyText = vbNullString
                Exit Sub
            End If

            ' Setting player sprite
            If LCase$(Mid$(MyText, 1, 16)) = "/setplayersprite" Then
                If Len(MyText) > 19 Then
                    i = Val(Mid$(MyText, 17, 1))
                    MyText = Mid$(MyText, 18, Len(MyText) - 17)
                    Call SendSetPlayerSprite(i, Val(MyText))
                End If
                MyText = vbNullString
                Exit Sub
            End If
        End If
            ' Respawn request
        If GetPlayerAccess(MyIndex) >= 2 Then
            If Mid$(MyText, 1, 8) = "/respawn" Then
                Call SendMapRespawn
                MyText = vbNullString
                Exit Sub
            End If

            ' MOTD change
            If Mid$(MyText, 1, 5) = "/motd" Then
                If Len(MyText) > 6 Then
                    MyText = Mid$(MyText, 7, Len(MyText) - 6)
                    If Trim$(MyText) <> vbNullString Then
                        Call SendMOTDChange(MyText)
                    End If
                End If
                MyText = vbNullString
                Exit Sub
            End If
        End If
        
        ' Checking the mute list
        If GetPlayerAccess(MyIndex) >= 1 Then
            If Mid$(MyText, 1, 12) = "/getmutelist" Then
                Call SendMuteList
                MyText = vbNullString
                Exit Sub
            End If
        End If
        
        ' Muting a player
        If GetPlayerAccess(MyIndex) >= 1 Then
            If LCase$(Mid$(MyText, 1, 5)) = "/mute" Then
                If Len(MyText) > 6 Then
                    MyText = Mid$(MyText, 7)
                    
                    If Len(MyText) > 0 Then
                        Call SendMute(MyText)
                    Else
                        Call AddText("Usage: /mute (username)", WHITE)
                    End If
                    
                    MyText = vbNullString
                End If
                Exit Sub
            End If
        End If
        
        ' Unmuting a player
        If GetPlayerAccess(MyIndex) >= 1 Then
            If LCase$(Mid$(MyText, 1, 7)) = "/unmute" Then
                If Len(MyText) > 8 Then
                    MyText = Mid$(MyText, 9)
                    
                    If Len(MyText) > 0 Then
                        Call SendUnmute(MyText)
                    Else
                        Call AddText("Usage: /unmute (username OR mute entry #)", WHITE)
                    End If
                    
                    MyText = vbNullString
                End If
                Exit Sub
            End If
        End If
        
        ' Checking the ban list
        If GetPlayerAccess(MyIndex) >= 1 Then
            If Mid$(MyText, 1, 11) = "/getbanlist" Then
                Call SendBanList
                MyText = vbNullString
                Exit Sub
            End If
        End If
        
        ' Banning a player
        If GetPlayerAccess(MyIndex) >= 1 Then
            If LCase$(Mid$(MyText, 1, 4)) = "/ban" Then
                If Len(MyText) > 5 Then
                    MyText = Mid$(MyText, 6, Len(MyText) - 5)
                    
                    If Len(MyText) > 0 Then
                        Call SendBan(MyText)
                    Else
                        Call AddText("Usage: /ban (username)", WHITE)
                    End If
                    
                    MyText = vbNullString
                End If
                Exit Sub
            End If
        End If
        
        ' Unbanning a player
        If GetPlayerAccess(MyIndex) >= 1 Then
            If LCase$(Mid$(MyText, 1, 6)) = "/unban" Then
                If Len(MyText) > 7 Then
                    MyText = Mid$(MyText, 8)
                    
                    If IsNumeric(MyText) = True Then
                        Call SendUnban(MyText)
                    Else
                        Call AddText("Usage: /unban (ban entry #)", WHITE)
                    End If
                    
                    MyText = vbNullString
                End If
                Exit Sub
            End If
        End If

        ' // Developer Admin Commands //
        If GetPlayerAccess(MyIndex) >= ADMIN_DEVELOPER Then
            ' Editing item request
            If Mid$(MyText, 1, 9) = "/edititem" Then
                Call SendRequestEditItem
                MyText = vbNullString
                Exit Sub
            End If

            ' Editing emoticon request
            If Mid$(MyText, 1, 13) = "/editemoticon" Then
                Call SendRequestEditEmoticon
                MyText = vbNullString
                Exit Sub
            End If

            ' Editing emoticon request
            If Mid$(MyText, 1, 12) = "/editelement" Then
                Call SendRequestEditElement
                MyText = vbNullString
                Exit Sub
            End If

            ' Editing arrow request
            If Mid$(MyText, 1, 13) = "/editarrow" Then
                Call SendRequestEditArrow
                MyText = vbNullString
                Exit Sub
            End If

            ' Editing npc request
            If Mid$(MyText, 1, 8) = "/editnpc" Then
                Call SendRequestEditNPC
                MyText = vbNullString
                Exit Sub
            End If

            ' Editing shop request
            If Mid$(MyText, 1, 9) = "/editshop" Then
                Call SendRequestEditShop
                MyText = vbNullString
                Exit Sub
            End If

            ' Editing spell request
            If LCase$(Trim$(MyText)) = "/editspell" Then
                Call SendRequestEditSpell
                MyText = vbNullString
                Exit Sub
            End If
            
            ' Editing recipe request
            If LCase$(Trim$(MyText)) = "/editrecipe" Then
                Call SendRequestEditRecipe
                MyText = vbNullString
                Exit Sub
            End If
        End If

        ' // Creator Admin Commands //
        If GetPlayerAccess(MyIndex) >= ADMIN_CREATOR Then
            ' Giving another player access
            If LCase$(Mid$(MyText, 1, 10)) = "/setaccess" Then
                ' Get access #
                i = Val(Mid$(MyText, 12, 1))
                
                If i > 0 Then
                    MyText = Mid$(MyText, 14, Len(MyText) - 13)
                    Call SendSetAccess(MyText, i)
                End If
                
                MyText = vbNullString
                Exit Sub
            End If
        End If

        ' Tell them its not a valid command
        If Left$(Trim$(MyText), 1) = "/" Then
            For i = 0 To MAX_EMOTICONS
                If Trim$(Emoticons(i).Command) = Trim$(MyText) And Trim$(Emoticons(i).Command) <> "/" Then
                    Call SendData(CPackets.Ccheckemoticons & SEP_CHAR & i & END_CHAR)
                    MyText = vbNullString
                    Exit Sub
                End If
            Next i
        End If
        
        ' Notify the player that no command was valid
        If Left$(MyText, 1) = "/" Then
            Call AddText("That is not a valid command!", BRIGHTRED)
            
            MyText = vbNullString
            Exit Sub
        End If
        
        ' Say message
        If Len(Trim$(MyText)) > 0 Then
            Call SayMsg(MyText)
        End If
        MyText = vbNullString
        Exit Sub
    End If
End Sub

Sub CheckMapGetItem()
    If GetTickCount > Player(MyIndex).MapGetTimer + 250 And Trim$(MyText) = vbNullString Then
        Player(MyIndex).MapGetTimer = GetTickCount
        Call SendData(CPackets.Cmapgetitem & END_CHAR)
    End If
End Sub

Sub CheckAttack()
    Dim i As Byte
    Dim AttackSpeed As Integer
    
    If IsCooking = True Or IsTrading = True Or IsShopping = True Or IsBanking = True Or IsHiderFrozen = True Then
        Exit Sub
    End If

    If ControlDown Then
        ' Check attack speed
        AttackSpeed = GetPlayerAttackSpeed(MyIndex)
        
        If Player(MyIndex).InBattle = False Then
            If Player(MyIndex).AttackTimer + AttackSpeed < GetTickCount Then
                If Player(MyIndex).Attacking = 0 Then
                    Player(MyIndex).Attacking = 1
                    Player(MyIndex).AttackTimer = GetTickCount
                    
                    Call SendData(CPackets.Cattack & END_CHAR)
                End If
            End If
        ElseIf (Player(MyIndex).InBattle = True And IsPlayerTurn = True) Then
            If Player(MyIndex).AttackTimer + AttackSpeed < GetTickCount Then
                If Player(MyIndex).Attacking = 0 Then
                    Select Case ButtonHighlighted
                        Case 0
                            Call SendRunFromBattle
                        Case 1
                            Call PlaySound("smas-smb3_itemmenu.wav")
                            frmMirage.picGuildMember.Visible = False
                            frmMirage.picGuildAdmin.Visible = False
                            frmMirage.picEquipment.Visible = False
                            frmMirage.picPlayerSpells.Visible = False
                            frmMirage.picWhosOnline.Visible = False
                            frmMirage.picCharStatus.Visible = False
                            frmMirage.picInventory.Visible = True
                            CanUseItem = True
                        Case 2
                            Call SendData(CPackets.Cattack & END_CHAR)
                        Case 3
                            Call PlaySound("smas-smb3_itemmenu.wav")
                            frmMirage.picGuildMember.Visible = False
                            frmMirage.picGuildAdmin.Visible = False
                            frmMirage.picEquipment.Visible = False
                            frmMirage.picPlayerSpells.Visible = True
                            frmMirage.picWhosOnline.Visible = False
                            frmMirage.picCharStatus.Visible = False
                            frmMirage.picInventory.Visible = False
                            CanUseSpecial = True
                    End Select
                End If
            End If
            
            Player(MyIndex).Attacking = 1
            Player(MyIndex).AttackTimer = GetTickCount
        End If
    End If
End Sub

Sub CheckInput(ByVal KeyState As Byte, ByVal KeyCode As Integer, ByVal Shift As Integer)
    If Not GettingMap Then
        If KeyState = 1 Then
            If KeyCode = vbKeyReturn Then
                Call CheckMapGetItem
            End If

            If KeyCode = vbKeyControl Then
                ControlDown = True
                Call SendHotScript(1)
            End If

            If KeyCode = vbKeyUp Then
                DirUp = True
                DirDown = False
                DirLeft = False
                DirRight = False
            End If

            If KeyCode = vbKeyDown Then
                DirUp = False
                DirDown = True
                DirLeft = False
                DirRight = False
            End If

            If KeyCode = vbKeyLeft Then
                DirUp = False
                DirDown = False
                DirLeft = True
                DirRight = False
                
                If ButtonHighlighted > 0 And IsPlayerTurn = True And Player(MyIndex).InBattle = True Then
                    ButtonHighlighted = ButtonHighlighted - 1
                    
                    Call PlaySound("smrpg_mario_kick.wav")
                    
                    CanUseItem = False
                    CanUseSpecial = False
                End If
            End If

            If KeyCode = vbKeyRight Then
                DirUp = False
                DirDown = False
                DirLeft = False
                DirRight = True
                
                If ButtonHighlighted < 3 And IsPlayerTurn = True And Player(MyIndex).InBattle = True Then
                    ButtonHighlighted = ButtonHighlighted + 1
                    
                    Call PlaySound("smrpg_mario_kick.wav")
                    
                    CanUseItem = False
                    CanUseSpecial = False
                End If
            End If

            If KeyCode = vbKeyShift Then
                ShiftDown = True
            End If
        End If
    End If
End Sub

Function IsTryingToMove() As Boolean
    If DirUp Or DirDown Or DirLeft Or DirRight Then
        IsTryingToMove = True
    End If
End Function

Function CanMove() As Boolean
    Dim JumpDirection() As String, JumpAddHeight() As String
    Dim i As Long, x As Long, y As Long
    
    CanMove = False

    If Player(MyIndex).Moving <> 0 Then
        Exit Function
    End If
    
    If Player(MyIndex).InBattle = True Then
        Exit Function
    End If
    
    If IsCooking = True Or IsTrading = True Or IsShopping = True Or IsBanking = True Or IsHiderFrozen = True Then
        Exit Function
    End If

    x = GetPlayerX(MyIndex)
    y = GetPlayerY(MyIndex)
    
    If DirUp Then
        Call SetPlayerDir(MyIndex, DIR_UP)
        y = y - 1
    ElseIf DirDown Then
        Call SetPlayerDir(MyIndex, DIR_DOWN)
        y = y + 1
    ElseIf DirLeft Then
        Call SetPlayerDir(MyIndex, DIR_LEFT)
        x = x - 1
    Else
        Call SetPlayerDir(MyIndex, DIR_RIGHT)
        x = x + 1
    End If

    If y < 0 Then
        If Map(GetPlayerMap(MyIndex)).Up > 0 Then
            Call SendPlayerRequestNewMap(DIR_UP)
            GettingMap = True
        End If
        Exit Function
    ElseIf y > MAX_MAPY Then
        If Map(GetPlayerMap(MyIndex)).Down > 0 Then
            Call SendPlayerRequestNewMap(DIR_DOWN)
            GettingMap = True
        End If
        Exit Function
    ElseIf x < 0 Then
        If Map(GetPlayerMap(MyIndex)).Left > 0 Then
            Call SendPlayerRequestNewMap(DIR_LEFT)
            GettingMap = True
        End If
        Exit Function
    ElseIf x > MAX_MAPX Then
        If Map(GetPlayerMap(MyIndex)).Right > 0 Then
            Call SendPlayerRequestNewMap(DIR_RIGHT)
            GettingMap = True
        End If
        Exit Function
    End If
    
    If Not GetPlayerDir(MyIndex) = LAST_DIR Then
        LAST_DIR = GetPlayerDir(MyIndex)
        Call SendPlayerDir
    End If
    
    ' Stop you from moving if you're stepping on a Jump Block tile
    If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex)).Type = TILE_TYPE_JUMPBLOCK Then
        If Player(MyIndex).Jumping = False Then
            Exit Function
        End If
    End If
    
    With Map(GetPlayerMap(MyIndex)).Tile(x, y)
        ' Check for Jugem's Cloud
        If .Type = TILE_TYPE_ROOFBLOCK Then
            If .String1 <> "Jugem's Cloud" Then
                Exit Function
            Else
                Call JugemsCloudWarp
                Exit Function
            End If
        End If
    
        ' Check if players can move onto certain scripted tiles
        If .Type = TILE_TYPE_SCRIPTED Then
            Select Case .Data1
                ' Seventeenth Quest Script (Getting Salestoad's Stuff again)
                Case 27
                    Exit Function
                ' Rockade Tile
                Case 32
                    For i = 1 To MAX_PLAYER_SPELLS
                        If Player(MyIndex).Spell(i) = 45 Then
                            CanMove = True
                            Exit For
                        End If
                    Next i
                    
                    If CanMove = False Then
                        Exit Function
                    End If
            End Select
        End If
    
        If .Type = TILE_TYPE_BLOCKED Or .Type = TILE_TYPE_SIGN Or .Type = TILE_TYPE_HOOKSHOT Or .Type = TILE_TYPE_SWITCH Or .Type = TILE_TYPE_DODGEBILL Then
            Exit Function
        End If
    
        If .Type = TILE_TYPE_JUMPBLOCK Then
            If Player(MyIndex).Jumping = False Then
                Exit Function
            Else
                ' Prevent us from moving if the jump block doesn't allow our direction
                JumpDirection = Split(Trim$(.String1), ",")
                
                If JumpDirection(GetPlayerDir(MyIndex)) = 0 Then
                    Exit Function
                End If
                
                ' Check if the direction we're moving in adds height or not
                JumpAddHeight = Split(Trim$(.String2), ",")
                
                ' The direction lowers height
                If JumpAddHeight(GetPlayerDir(MyIndex)) = 0 Then
                    ' Allow us to move to a lower height while jumping
                    Call SetPlayerHeight(Player(MyIndex).Height - .Data2)
                        
                    ' Update height on the server
                    Call SendData(CPackets.Cplayerheight & SEP_CHAR & Player(MyIndex).Height & END_CHAR)
                    CanMove = True
                    Exit Function
                End If
                
                ' Don't allow us to move unless we're very near/at the peak of the jump
                If Player(MyIndex).TempJumpAnim < 11 Or Player(MyIndex).TempJumpAnim > 19 Then
                    Exit Function
                Else
                    ' Allow us to move to a higher height while we're at the peak of our jump
                    If JumpAddHeight(GetPlayerDir(MyIndex)) <> 0 Then
                        If Player(MyIndex).Height >= .Data1 Then
                            Call SetPlayerHeight(Player(MyIndex).Height + 1)
                            
                            ' Update height on the server
                            Call SendData(CPackets.Cplayerheight & SEP_CHAR & Player(MyIndex).Height & END_CHAR)
                            CanMove = True
                            Exit Function
                        Else
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
        
        If .Type = TILE_TYPE_CBLOCK Then
            If .Data1 = Player(MyIndex).Class Or .Data2 = Player(MyIndex).Class Or .Data3 = Player(MyIndex).Class Then
                CanMove = True
                Exit Function
            End If
            
            Exit Function
        End If

        If .Type = TILE_TYPE_GUILDBLOCK Then
            If .String1 <> GetPlayerGuild(MyIndex) Then
                Exit Function
            End If
        End If

        If .Type = TILE_TYPE_KEY Or .Type = TILE_TYPE_DOOR Then
            If TempTile(x, y).DoorOpen = No Then
                Exit Function
            End If
        End If

        If .Type = TILE_TYPE_WALKTHRU Then
            CanMove = True
            Exit Function
        End If
    
        If .Type = TILE_TYPE_LVLBLOCK Then
            If GetPlayerLevel(MyIndex) < .Data1 Then
                Call AddText("You must be at least level " & .Data1 & " to pass!", GREEN)
                Exit Function
            End If
        End If
    End With
    
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            If GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                If GetPlayerX(i) = x Then
                    If GetPlayerY(i) = y Then
                        If Player(i).InBattle = False Then
                            If ((Player(MyIndex).Jumping = True) And (Player(MyIndex).TempJumpAnim <= 10) Or (Player(MyIndex).TempJumpAnim >= 20)) Or Player(MyIndex).Jumping = False Then
                                CanMove = False
                                Exit Function
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next i
    
    For i = 1 To MAX_MAP_NPCS
        With MapNpc(i)
            If .num > 0 Then
                If .x = x Then
                    If .y = y Then
                        ' Prevent players from moving onto a Rockade
                        If .num = 195 Then
                            CanMove = False
                            Exit Function
                        End If
                        
                        If .InBattle = False Then
                            If Map(GetPlayerMap(MyIndex)).Moral <> MAP_MORAL_MINIGAME Then
                                If Npc(.num).Behavior <> NPC_BEHAVIOR_FRIENDLY And Npc(.num).Behavior <> NPC_BEHAVIOR_SCRIPTED Then
                                    If frmMirage.chkTurnBased.Value = Unchecked And IsInPoisonCave(MyIndex) = False Then
                                        CanMove = False
                                        Exit Function
                                    Else
                                        If GetTickCount <= (BattleVictoryTimer + 2000) Then
                                            CanMove = False
                                            Exit Function
                                        End If
                                    End If
                                Else
                                    CanMove = False
                                    Exit Function
                                End If
                            Else
                                CanMove = False
                                Exit Function
                            End If
                        End If
                    End If
                End If
            End If
        End With
    Next i
    
    CanMove = True
End Function

Sub CheckMovement()
 Dim s2kX As Integer, s2kY As Integer   ' used below for temp store of X/Y
    If Not GettingMap Then
        If IsTryingToMove Then
            If CanMove Then
                ' Check if the player is holding down the Shift key or has the Auto-Run feature enabled
                If (ShiftDown Or frmMirage.chkAutoRun.Value = Checked) Or (ShiftDown And frmMirage.chkAutoRun.Value = Checked) Then
                    Player(MyIndex).Moving = MOVING_RUNNING
                Else
                    Player(MyIndex).Moving = MOVING_WALKING
                End If
                
                 Select Case GetPlayerDir(MyIndex)
                    Case DIR_UP
                       Player(MyIndex).yOffset = PIC_Y
                       Call SetPlayerY(MyIndex, GetPlayerY(MyIndex) - 1)

                    Case DIR_DOWN
                       Player(MyIndex).yOffset = PIC_Y * -1
                       Call SetPlayerY(MyIndex, GetPlayerY(MyIndex) + 1)

                    Case DIR_LEFT
                       Player(MyIndex).xOffset = PIC_X
                       Call SetPlayerX(MyIndex, GetPlayerX(MyIndex) - 1)

                    Case DIR_RIGHT
                       Player(MyIndex).xOffset = PIC_X * -1
                       Call SetPlayerX(MyIndex, GetPlayerX(MyIndex) + 1)
                  End Select
            
                    Call SendPlayerMove     '090829 moved here
                    s2kX = GetPlayerX(MyIndex)  '090829
                    s2kY = GetPlayerY(MyIndex)  '090829
            
                ' Gotta check :)
                If Map(GetPlayerMap(MyIndex)).Tile(s2kX, s2kY).Type = TILE_TYPE_WARP Or s2kX < 0 Or s2kX > MAX_MAPX Or s2kY < 0 Or s2kY > MAX_MAPY Then
                    GettingMap = True
                End If
            End If
        End If
    End If
End Sub

Function FindPlayer(ByVal Name As String) As Long
    Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            ' Make sure we dont try to check a name thats to small
            If Len(GetPlayerName(i)) >= Len(Trim$(Name)) Then
                If UCase$(Trim$(GetPlayerName(i))) = UCase$(Trim$(Name)) Then
                    FindPlayer = i
                    Exit Function
                End If
            End If
        End If
    Next i

    FindPlayer = 0
End Function

Public Sub DisplayInventoryInTrade()
    Dim i As Long
    Dim ItemNum As Long
    
    frmPlayerTrade.PlayerInv1.Clear

    For i = 1 To Player(MyIndex).MaxInv
      ItemNum = GetPlayerInvItemNum(MyIndex, i)
        If ItemNum > 0 And ItemNum <= MAX_ITEMS Then
            If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Or Item(ItemNum).Stackable = 1 Then
                frmPlayerTrade.PlayerInv1.addItem i & ": " & Trim$(Item(ItemNum).Name) & " (" & GetPlayerInvItemValue(MyIndex, i) & ")"
            Else
                frmPlayerTrade.PlayerInv1.addItem i & ": " & Trim$(Item(ItemNum).Name)
            End If
        Else
            frmPlayerTrade.PlayerInv1.addItem i & ": <Nothing>"
        End If
    Next i

    frmPlayerTrade.PlayerInv1.ListIndex = 0
End Sub

Sub PlayerSearch(Button As Integer, Shift As Integer, x As Long, y As Long)
    If CurX >= 0 And CurX <= MAX_MAPX Then
        If CurY >= 0 And CurY <= MAX_MAPY Then
            Call SendData(CPackets.Csearch & SEP_CHAR & CurX & SEP_CHAR & CurY & END_CHAR)
        End If
    End If
End Sub

Public Sub UpdateVisInv()
    Dim i As Long, ItemNum As Long
    Dim srec As RECT, drec As RECT
    
    If DD_ItemSurf Is Nothing Then
        Exit Sub
    End If
    
    For i = 1 To 7
        frmMirage.EquipImage(i - 1).Picture = LoadPicture()
        ItemNum = GetPlayerEquipSlotNum(MyIndex, i)
        
        If ItemNum > 0 Then
            drec.Top = 0
            drec.Bottom = PIC_X
            drec.Left = 0
            drec.Right = PIC_Y
            srec.Top = (Item(ItemNum).Pic \ 6) * PIC_Y
            srec.Bottom = srec.Top + PIC_X
            srec.Left = (Item(ItemNum).Pic - (Item(ItemNum).Pic \ 6) * 6) * PIC_X
            srec.Right = srec.Left + PIC_Y
            
            Call DD_ItemSurf.BltToDC(frmMirage.EquipImage(i - 1).hDC, srec, drec)
        End If
    Next i
End Sub

Public Sub UpdateInventory()
    If frmMirage.picInventory.Visible = True Then
        Dim i As Long, ItemNum As Long
        Dim srec As RECT, drec As RECT
        Dim LeftInvSlot As Integer, TopInvSlot As Integer
                    
        TopInvSlot = -1
                    
        For i = InventorySlotsIndex To Player(MyIndex).MaxInv
            ItemNum = Player(MyIndex).NewInv(i).num
                    
            If LeftInvSlot > 3 Then
                LeftInvSlot = 0
            End If
                        
            If LeftInvSlot = 0 Then
                TopInvSlot = TopInvSlot + 1
            End If
                        
            If ItemNum > 0 Then
                With srec
                    .Left = (Item(ItemNum).Pic - (Item(ItemNum).Pic \ 6) * 6) * PIC_X
                    .Right = srec.Left + PIC_X
                    .Top = (Item(ItemNum).Pic \ 6) * PIC_Y
                    .Bottom = srec.Top + PIC_Y
                End With
            Else
                With srec
                    .Left = 0
                    .Right = 1
                    .Top = 0
                    .Bottom = 1
                End With
            End If
                        
            With drec
                .Left = 6 + (LeftInvSlot * 38)
                .Right = .Left + PIC_X
                .Top = 1 + (TopInvSlot * 35)
                .Bottom = .Top + PIC_Y
            End With
                
            Call DD_ItemSurf.BltToDC(frmMirage.picInventory3.hDC, srec, drec)
            
            LeftInvSlot = LeftInvSlot + 1
        Next i
        
        Dim TempInventorySlotsIndex As Integer, LeftIndex As Integer

        TempInventorySlotsIndex = (InventorySlotsIndex + 24)
        
        If TempInventorySlotsIndex > Player(MyIndex).MaxInv Then
            TempInventorySlotsIndex = TempInventorySlotsIndex - 5
            LeftIndex = TempInventorySlotsIndex - 3
            
            For i = (Player(MyIndex).MaxInv + 1) To TempInventorySlotsIndex
                With srec
                    .Left = 0
                    .Right = 1
                    .Top = 0
                    .Bottom = 1
                End With

                With drec
                    .Left = 6 + ((i - LeftIndex) * 38)
                    .Right = .Left + PIC_X
                    .Top = 1 + (TopInvSlot * 35)
                    .Bottom = .Top + PIC_Y
                End With

                Call DD_ItemSurf.BltToDC(frmMirage.picInventory3.hDC, srec, drec)
            Next
        End If
    End If
End Sub

Sub UpdateBank()
    Dim InvListIndex As Integer, BankListIndex As Integer
    Dim i As Long
    Dim ItemNum As Long
    
    InvListIndex = frmBank.lstInventory.ListIndex
    BankListIndex = frmBank.lstBank.ListIndex
    
    frmBank.lstInventory.Clear
    frmBank.lstBank.Clear

    For i = 1 To Player(MyIndex).MaxInv
      ItemNum = GetPlayerInvItemNum(MyIndex, i)
        If ItemNum > 0 Then
            If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Or Item(ItemNum).Stackable = 1 Then
                frmBank.lstInventory.addItem i & "> " & Trim$(Item(ItemNum).Name) & " (" & GetPlayerInvItemValue(MyIndex, i) & ")"
            Else
                frmBank.lstInventory.addItem i & "> " & Trim$(Item(ItemNum).Name)
            End If
        Else
            frmBank.lstInventory.addItem i & "> Empty"
        End If
    Next i

    For i = 1 To MAX_BANK
      ItemNum = GetPlayerBankItemNum(MyIndex, i)
        If ItemNum > 0 Then
            If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Or Item(ItemNum).Stackable = 1 Then
                frmBank.lstBank.addItem i & "> " & Trim$(Item(ItemNum).Name) & " (" & GetPlayerBankItemValue(MyIndex, i) & ")"
            Else
                frmBank.lstBank.addItem i & "> " & Trim$(Item(ItemNum).Name)
            End If
        Else
            frmBank.lstBank.addItem i & "> Empty"
        End If
    Next i
    
    frmBank.lstInventory.ListIndex = InvListIndex
    frmBank.lstBank.ListIndex = BankListIndex
End Sub

Sub UseItem()
    Call SendUseItem(Inventory)
End Sub

Sub DropItem()
    Dim ItemNum As Long

    On Error GoTo DropItem_Error

    ItemNum = GetPlayerInvItemNum(MyIndex, Inventory)
    
    If ItemNum > 0 And ItemNum <= MAX_ITEMS Then
        If Item(ItemNum).Bound = 0 Then
            If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Or Item(ItemNum).Stackable = 1 Then
                Dim GoldAmount As String, itemName As String
                
                itemName = Trim$(Item(ItemNum).Name)
                GoldAmount = InputBox("How much " & itemName & "(" & GetPlayerInvItemValue(MyIndex, Inventory) & ") would you like to drop?", "Drop " & itemName, 0, frmMirage.Left, frmMirage.Top)

                If IsNumeric(GoldAmount) Then
                    Call SendDropItem(Inventory, GoldAmount)
                End If
            Else
                Call SendDropItem(Inventory, 0)
            End If
        End If
    End If

    Call UpdateVisInv

    Exit Sub

DropItem_Error:
    Call AddText("Please enter a valid amount for that item!", BRIGHTRED)
End Sub

' Sets the speed of a character based on speed
Sub SetSpeed(ByVal Run As String, ByVal speed As Long)
    If Run = "walk" Then
        SS_WALK_SPEED = speed
    ElseIf Run = "run" Then
        SS_RUN_SPEED = speed
    End If
End Sub

Public Sub AlwaysOnTop(FormName As Form, bOnTop As Boolean)
    If Not bOnTop Then
        Call SetWindowPos(FormName.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
    Else
        Call SetWindowPos(FormName.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
    End If
End Sub

Sub GoShop(ByVal Shop As Integer)
    ' Close any other shop windows
    frmNewShop.Hide

    ' Initialize the shop
    Call frmNewShop.loadShop(Shop)
    snumber = Shop

    ' Hide panel
    frmNewShop.picItemInfo.Visible = False

    On Error Resume Next
    
    ' Set focus
    frmNewShop.SetFocus

    ' Show page 1 (it starts from 0)
    frmNewShop.showPage (0)
    
    ' Show shop
    frmNewShop.Show vbModeless, frmMirage
End Sub

' Returns true if the tile is a roof tile and the player is under that section of roof
Function IsTileRoof(ByVal x As Integer, ByVal y As Integer) As Boolean
    Dim IsRoof As Boolean
    
    If Map(GetPlayerMap(MyIndex)).Tile(x, y).Type = TILE_TYPE_ROOF Or Map(GetPlayerMap(MyIndex)).Tile(x, y).Type = TILE_TYPE_ROOFBLOCK Then 'If the tile is a roof or a roofblock
        If Map(GetPlayerMap(MyIndex)).Tile(x, y).String1 = Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex)).String1 Then 'If the roof ID is the same
            IsTileRoof = True
            Exit Function
        End If
    End If

    IsTileRoof = False
End Function

Function CanBltPlayerGfx(ByVal Index As Long) As Boolean
    If Index <> MyIndex Then
        If Player(Index).InBattle = False Then
            CanBltPlayerGfx = True
            Exit Function
        End If
    Else
        If Player(Index).BattleVictory = False Then
            CanBltPlayerGfx = True
            Exit Function
        End If
    End If
    
    CanBltPlayerGfx = False
End Function

Function CanBltNpcGfx(ByVal MapNpcNum As Long) As Boolean
    If MapNpc(MapNpcNum).InBattle = True Then
        If MapNpc(MapNpcNum).Target = MyIndex Then
            CanBltNpcGfx = True
            Exit Function
        Else
            CanBltNpcGfx = False
            Exit Function
        End If
    Else
        CanBltNpcGfx = True
        Exit Function
    End If
    
    CanBltNpcGfx = False
End Function

Public Function Rand(ByVal Low As Long, ByVal High As Long) As Long
    Randomize
    Rand = Int((High - Low + 1) * Rnd) + Low
End Function

Public Function Rnd2()
    Randomize
    Rnd2 = Rnd
End Function

Sub PlayBattleSong()
    Dim Song As Integer
    
    Song = Int(Rand(1, 6))
    
    Select Case Song
        Case 1
            Call PlayBGM("Mario & Luigi Bowser's Inside Story - M&L Battle Theme.mp3")
        Case 2
            Call PlayBGM("Mario & Luigi Partners In Time - Battle Theme.mp3")
        Case 3
            Call PlayBGM("Mario & Luigi Superstar Saga - Battle Theme.mp3")
        Case 4
            Call PlayBGM("Paper Mario - Battle Theme.mp3")
        Case 5
            Call PlayBGM("Super Mario RPG - Battle Theme.mp3")
        Case 6
            Call PlayBGM("Paper Mario The Thousand Year Door - Battle Theme.mid")
    End Select
End Sub

Sub EndNpcTalkConditions(ByVal NpcNum As Long)
    Select Case NpcNum
        Case 194 ' Armored Koopa
            ' Start a turn-based battle
            Call SendData(CPackets.Cstartbattle & SEP_CHAR & NpcNum & END_CHAR)
        Case 208 ' Castle Town Doctor
            Call SendData(CPackets.Cdoctorheal & END_CHAR)
    End Select
End Sub

Sub EndNpcTalkToConditions(ByVal NpcNum As Long, ByVal IsYes As Boolean)
    Select Case NpcNum
        Case 221 ' Currency Exchanger
            If IsYes = True Then
                IsShopping = True
        
                ' Show the shop
                Call GoShop(23)
            Else
                ' Open the bank
                Call frmBank.OpenBank
            End If
    End Select
End Sub

Function FlowerSaver(ByVal SpellNum As Long) As Long
    Dim i As Integer
    
    ' Account for the Flower Saver special attack
    FlowerSaver = Spell(SpellNum).MPCost
    
    ' Check if the player has the Flower Saver special attack
    For i = 1 To MAX_PLAYER_SPELLS
        If Player(MyIndex).Spell(i) = 43 Then
            ' Reduce the cost of the special attack by 30% and always round up
            FlowerSaver = Int((-FlowerSaver * 0.7)) * -1
            
            Exit Function
        End If
    Next i
End Function

Function IsInPoisonCave(ByVal Index As Long) As Boolean
    If GetPlayerMap(Index) >= 320 And GetPlayerMap(Index) <= 328 Then
        IsInPoisonCave = True
    End If
End Function
