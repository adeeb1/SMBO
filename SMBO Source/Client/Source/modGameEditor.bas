Attribute VB_Name = "modGameEditor"
Option Explicit

Public Sub EditorInit()
    Dim i As Long
    
    InEditor = True

    Call frmMapEditor.Show(vbModeless, frmMirage)
    
    EditorSet = 0
    
    MapEditorSelectedType = 1

    For i = 0 To 10
        If frmMapEditor.Option1(i).Value = True Then
            Call InitLoadPicture(App.Path & "\GFX\Tiles" & i & ".smbo", frmMapEditor.picBackSelect)

            EditorSet = i
        End If
    Next i

    frmMapEditor.scrlPicture.Max = ((frmMapEditor.picBackSelect.Height - frmMapEditor.picBack.Height) \ PIC_Y)
    frmMapEditor.picBack.Width = 448
End Sub

Public Sub MainMenuInit()
    frmLogin.txtName.Text = Trim$(ReadINI("CONFIG", "Account", App.Path & "\config.ini"))
    frmLogin.txtPassword.Text = Trim$(ReadINI("CONFIG", "Password", App.Path & "\config.ini"))

    If frmLogin.Check1.Value = 0 Then
        frmLogin.Check2.Value = 0
    End If

    If ConnectToServer = True And AutoLogin = 1 Then
        frmMainMenu.picAutoLogin.Visible = True
        frmChars.Label1.Visible = False
    Else
        frmMainMenu.picAutoLogin.Visible = False
        frmChars.Label1.Visible = True
    End If
End Sub

Public Sub ParseNews(ByVal FileTitle As String, ByVal FileBody As String, ByVal RED As Integer, ByVal BLUE As Integer, ByVal GRN As Integer)

    frmMainMenu.picNews.Caption = FileTitle & vbNewLine & vbNewLine & FileBody

    If RED < 0 Or RED > 255 Or GRN < 0 Or GRN > 255 Or BLUE < 0 Or BLUE > 255 Then
        frmMainMenu.picNews.ForeColor = RGB(255, 255, 255)
    Else
        frmMainMenu.picNews.ForeColor = RGB(RED, GRN, BLUE)
    End If
End Sub

Public Sub EditorMouseDown(Button As Integer, Shift As Integer, x As Long, y As Long)
    Dim x2 As Long, y2 As Long, PicX As Long

    If InEditor Then

        If frmMapEditor.MousePointer = 2 Then
            If MapEditorSelectedType = 1 Then
                With Map(GetPlayerMap(MyIndex)).Tile(x, y)
                    If frmMapEditor.optGround.Value = True Then
                        PicX = .Ground
                        EditorSet = .GroundSet
                    End If
                    If frmMapEditor.optMask.Value = True Then
                        PicX = .Mask
                        EditorSet = .MaskSet
                    End If
                    If frmMapEditor.optAnim.Value = True Then
                        PicX = .Anim
                        EditorSet = .AnimSet
                    End If
                    If frmMapEditor.optMask2.Value = True Then
                        PicX = .Mask2
                        EditorSet = .Mask2Set
                    End If
                    If frmMapEditor.optM2Anim.Value = True Then
                        PicX = .M2Anim
                        EditorSet = .M2AnimSet
                    End If
                    If frmMapEditor.optFringe.Value = True Then
                        PicX = .Fringe
                        EditorSet = .FringeSet
                    End If
                    If frmMapEditor.optFAnim.Value = True Then
                        PicX = .FAnim
                        EditorSet = .FAnimSet
                    End If
                    If frmMapEditor.optFringe2.Value = True Then
                        PicX = .Fringe2
                        EditorSet = .Fringe2Set
                    End If
                    If frmMapEditor.optF2Anim.Value = True Then
                        PicX = .F2Anim
                        EditorSet = .F2AnimSet
                    End If

                    EditorTileY = (PicX \ TilesInSheets)
                    EditorTileX = (PicX - (PicX \ TilesInSheets) * TilesInSheets)
                    frmMapEditor.shpSelected.Top = Int(EditorTileY * PIC_Y)
                    frmMapEditor.shpSelected.Left = Int(EditorTileX * PIC_Y)
                    frmMapEditor.shpSelected.Height = PIC_Y
                    frmMapEditor.shpSelected.Width = PIC_X
                End With
                
            ElseIf MapEditorSelectedType = 3 Then
                EditorTileY = (Map(GetPlayerMap(MyIndex)).Tile(x, y).light \ TilesInSheets)
                EditorTileX = (Map(GetPlayerMap(MyIndex)).Tile(x, y).light - (Map(GetPlayerMap(MyIndex)).Tile(x, y).light \ TilesInSheets) * TilesInSheets)
                frmMapEditor.shpSelected.Top = Int(EditorTileY * PIC_Y)
                frmMapEditor.shpSelected.Left = Int(EditorTileX * PIC_Y)
                frmMapEditor.shpSelected.Height = PIC_Y
                frmMapEditor.shpSelected.Width = PIC_X
                
            ElseIf MapEditorSelectedType = 2 Then
                With Map(GetPlayerMap(MyIndex)).Tile(x, y)
                    Select Case .Type
                        Case TILE_TYPE_BLOCKED
                            frmMapEditor.optBlocked.Value = True
                        Case TILE_TYPE_WALKTHRU
                            frmMapEditor.optWalkThru.Value = True
                        Case TILE_TYPE_WARP
                            EditorWarpMap = .Data1
                            EditorWarpX = .Data2
                            EditorWarpY = .Data3
                            frmMapEditor.optWarp.Value = True
                        Case TILE_TYPE_HEAL
                            frmMapEditor.optHeal.Value = True
                        Case TILE_TYPE_ROOFBLOCK
                            frmMapEditor.optRoofBlock.Value = True
                            RoofId = .String1
                        Case TILE_TYPE_ROOF
                            frmMapEditor.optRoof.Value = True
                            RoofId = .String1
                        Case TILE_TYPE_KILL
                            frmMapEditor.optKill.Value = True
                        Case TILE_TYPE_ITEM
                            ItemEditorNum = .Data1
                            ItemEditorValue = .Data2
                            frmMapEditor.optItem.Value = True
                        Case TILE_TYPE_NPCAVOID
                            frmMapEditor.optNpcAvoid.Value = True
                        Case TILE_TYPE_KEY
                            KeyEditorNum = .Data1
                            KeyEditorTake = .Data2
                            KeyText = .String1
                            frmMapEditor.optKey.Value = True
                        Case TILE_TYPE_KEYOPEN
                            KeyOpenEditorX = .Data1
                            KeyOpenEditorY = .Data2
                            KeyOpenEditorMsg = .String1
                            frmMapEditor.optKeyOpen.Value = True
                        Case TILE_TYPE_SHOP
                            EditorShopNum = .Data1
                            frmMapEditor.optShop.Value = True
                        Case TILE_TYPE_CBLOCK
                            EditorItemNum1 = .Data1
                            EditorItemNum2 = .Data2
                            EditorItemNum3 = .Data3
                            frmMapEditor.optCBlock.Value = True
                        Case TILE_TYPE_ARENA
                            Arena1 = .Data1
                            Arena2 = .Data2
                            Arena3 = .Data3
                            frmMapEditor.optArena.Value = True
                        Case TILE_TYPE_SOUND
                            SoundFileName = .String1
                            frmMapEditor.optSound.Value = True
                        Case TILE_TYPE_SPRITE_CHANGE
                            SpritePic = .Data1
                            frmMapEditor.optSprite.Value = True
                        Case TILE_TYPE_SIGN
                            SignLine1 = .String1
                            SignLine2 = .String2
                            SignLine3 = .String3
                            frmMapEditor.optSign.Value = True
                        Case TILE_TYPE_DOOR
                            frmMapEditor.optDoor.Value = True
                        Case TILE_TYPE_NOTICE
                            NoticeTitle = .String1
                            NoticeText = .String2
                            NoticeSound = .String3
                            frmMapEditor.optNotice.Value = True
                        Case TILE_TYPE_CHEST
                            frmMapEditor.optChest.Value = True
                        Case TILE_TYPE_CLASS_CHANGE
                            ClassChange = .Data1
                            ClassChangeReq = .Data2
                            frmMapEditor.optClassChange.Value = True
                        Case TILE_TYPE_SCRIPTED
                            ScriptNum = .Data1
                            frmMapEditor.optScripted.Value = True
                        Case TILE_TYPE_GUILDBLOCK
                            GuildBlock = .Data1
                            frmMapEditor.optGuildBlock.Value = True
                        Case TILE_TYPE_BANK
                            frmMapEditor.optBank.Value = True
                        Case TILE_TYPE_HOOKSHOT
                            frmMapEditor.OptGHook.Value = True
                        Case TILE_TYPE_ONCLICK
                            ClickScript = .Data1
                            frmMapEditor.optClick.Value = True
                        Case TILE_TYPE_LOWER_STAT
                            MinusHp = frmMinusStat.scrlNum1.Value
                            MinusMp = frmMinusStat.scrlNum1.Value
                            MinusSp = frmMinusStat.scrlNum1.Value
                            MessageMinus = frmMinusStat.Text1.Text
                            frmMapEditor.optMinusStat.Value = True
                        Case TILE_TYPE_SWITCH
                            SwitchWarpMap = .Data1
                            SwitchWarpPos = .Data2
                            SwitchWarpFlags = .Data3
                            frmMapEditor.optSwitch.Value = True
                        Case TILE_TYPE_LVLBLOCK
                            LevelToBlock = .Data1
                            frmMapEditor.optLevelBlock.Value = True
                        Case TILE_TYPE_DRILL
                            EditorWarpMap = .Data1
                            EditorWarpX = .Data2
                            EditorWarpY = .Data3
                            frmMapEditor.optDrill.Value = True
                        Case TILE_TYPE_JUMPBLOCK
                            frmMapEditor.optJumpBlock.Value = True
                            JumpHeight = .Data1
                            JumpDecrease = .Data2
                        Case TILE_TYPE_DODGEBILL
                            frmMapEditor.optDodgeBill.Value = True
                        Case TILE_TYPE_HAMMERBARRAGE
                            EditorWarpMap = .Data1
                            EditorWarpX = .Data2
                            EditorWarpY = .Data3
                            frmMapEditor.optHammerBarrage.Value = True
                        Case TILE_TYPE_JUGEMSCLOUD
                            CloudDir = .Data1
                            frmMapEditor.optJugemsCloud.Value = True
                        Case TILE_TYPE_SIMULBLOCK
                            frmMapEditor.optSimulBlock.Value = True
                        Case TILE_TYPE_BEAN
                            BeanItemNum = .Data1
                            BeanItemQuantity = .Data2
                            frmMapEditor.optBean.Value = True
                    End Select
                End With
            
                If Map(GetPlayerMap(MyIndex)).Tile(x, y).Type = TILE_TYPE_QUESTIONBLOCK Then
                    With QuestionBlock(GetPlayerMap(MyIndex), x, y)
                        ItemThing1 = .Item1
                        ItemThing2 = .Item2
                        ItemThing3 = .Item3
                        ItemThing4 = .Item4
                        ItemThing5 = .Item5
                        ItemThing6 = .Item6
                        ChanceThing1 = .Chance1
                        ChanceThing2 = .Chance2
                        ChanceThing3 = .Chance3
                        ChanceThing4 = .Chance4
                        ChanceThing5 = .Chance5
                        ChanceThing6 = .Chance6
                        ValueThing1 = .Value1
                        ValueThing2 = .Value2
                        ValueThing3 = .Value3
                        ValueThing4 = .Value4
                        ValueThing5 = .Value5
                        ValueThing6 = .Value6
                   End With
                   
                   frmMapEditor.optQuestionBlock.Value = True
                End If
            End If
            
            frmMapEditor.MousePointer = 1
            frmMirage.MousePointer = 1
        Else
            If (Button = 1) And (x >= 0) And (x <= MAX_MAPX) And (y >= 0) And (y <= MAX_MAPY) Then
                If frmMapEditor.shpSelected.Height <= PIC_Y And frmMapEditor.shpSelected.Width <= PIC_X Then
                    If MapEditorSelectedType = 1 Then
                        With Map(GetPlayerMap(MyIndex)).Tile(x, y)
                            If frmMapEditor.optGround.Value = True Then
                                .Ground = EditorTileY * TilesInSheets + EditorTileX
                                .GroundSet = EditorSet
                            End If
                            If frmMapEditor.optMask.Value = True Then
                                .Mask = EditorTileY * TilesInSheets + EditorTileX
                                .MaskSet = EditorSet
                            End If
                            If frmMapEditor.optAnim.Value = True Then
                                .Anim = EditorTileY * TilesInSheets + EditorTileX
                                .AnimSet = EditorSet
                            End If
                            If frmMapEditor.optMask2.Value = True Then
                                .Mask2 = EditorTileY * TilesInSheets + EditorTileX
                                .Mask2Set = EditorSet
                            End If
                            If frmMapEditor.optM2Anim.Value = True Then
                                .M2Anim = EditorTileY * TilesInSheets + EditorTileX
                                .M2AnimSet = EditorSet
                            End If
                            If frmMapEditor.optFringe.Value = True Then
                                .Fringe = EditorTileY * TilesInSheets + EditorTileX
                                .FringeSet = EditorSet
                            End If
                            If frmMapEditor.optFAnim.Value = True Then
                                .FAnim = EditorTileY * TilesInSheets + EditorTileX
                                .FAnimSet = EditorSet
                            End If
                            If frmMapEditor.optFringe2.Value = True Then
                                .Fringe2 = EditorTileY * TilesInSheets + EditorTileX
                                .Fringe2Set = EditorSet
                            End If
                            If frmMapEditor.optF2Anim.Value = True Then
                                .F2Anim = EditorTileY * TilesInSheets + EditorTileX
                                .F2AnimSet = EditorSet
                            End If
                        End With
                    ElseIf MapEditorSelectedType = 3 Then
                        Map(GetPlayerMap(MyIndex)).Tile(x, y).light = EditorTileY * TilesInSheets + EditorTileX
                    ElseIf MapEditorSelectedType = 2 Then
                        With Map(GetPlayerMap(MyIndex)).Tile(x, y)
                            If frmMapEditor.optBlocked.Value = True Then
                                .Type = TILE_TYPE_BLOCKED
                            End If
                            If frmMapEditor.optRoofBlock.Value = True Then
                                .Type = TILE_TYPE_ROOFBLOCK
                                .Data1 = 0
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = RoofId
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.optRoof.Value = True Then
                                .Type = TILE_TYPE_ROOF
                                .Data1 = 0
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = RoofId
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.optWarp.Value = True Then
                                .Type = TILE_TYPE_WARP
                                .Data1 = EditorWarpMap
                                .Data2 = EditorWarpX
                                .Data3 = EditorWarpY
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.optHeal.Value = True Then
                                .Type = TILE_TYPE_HEAL
                                .Data1 = 0
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.optKill.Value = True Then
                                .Type = TILE_TYPE_KILL
                                .Data1 = 0
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If GetPlayerAccess(MyIndex) >= 3 Then
                                If frmMapEditor.optItem.Value = True Then
                                    .Type = TILE_TYPE_ITEM
                                    .Data1 = ItemEditorNum
                                    .Data2 = ItemEditorValue
                                    .Data3 = 0
                                    .String1 = vbNullString
                                    .String2 = vbNullString
                                    .String3 = vbNullString
                                End If
                            End If
                            If frmMapEditor.optNpcAvoid.Value = True Then
                                .Type = TILE_TYPE_NPCAVOID
                                .Data1 = 0
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.optKey.Value = True Then
                                .Type = TILE_TYPE_KEY
                                .Data1 = KeyEditorNum
                                .Data2 = KeyEditorTake
                                .Data3 = 0
                                .String1 = KeyText
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.optKeyOpen.Value = True Then
                                .Type = TILE_TYPE_KEYOPEN
                                .Data1 = KeyOpenEditorX
                                .Data2 = KeyOpenEditorY
                                .Data3 = 0
                                .String1 = KeyOpenEditorMsg
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.optShop.Value = True Then
                                .Type = TILE_TYPE_SHOP
                                .Data1 = EditorShopNum
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.optCBlock.Value = True Then
                                .Type = TILE_TYPE_CBLOCK
                                .Data1 = EditorItemNum1
                                .Data2 = EditorItemNum2
                                .Data3 = EditorItemNum3
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.optArena.Value = True Then
                                .Type = TILE_TYPE_ARENA
                                .Data1 = Arena1
                                .Data2 = Arena2
                                .Data3 = Arena3
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.optSound.Value = True Then
                                .Type = TILE_TYPE_SOUND
                                .Data1 = 0
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = SoundFileName
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.optSprite.Value = True Then
                                .Type = TILE_TYPE_SPRITE_CHANGE
                                .Data1 = SpritePic
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.optSign.Value = True Then
                                .Type = TILE_TYPE_SIGN
                                .Data1 = 0
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = SignLine1
                                .String2 = SignLine2
                                .String3 = SignLine3
                            End If
                            If frmMapEditor.optDoor.Value = True Then
                                .Type = TILE_TYPE_DOOR
                                .Data1 = 0
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.optNotice.Value = True Then
                                .Type = TILE_TYPE_NOTICE
                                .Data1 = 0
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = NoticeTitle
                                .String2 = NoticeText
                                .String3 = NoticeSound
                            End If
                            If frmMapEditor.optChest.Value = True Then
                                .Type = TILE_TYPE_CHEST
                                .Data1 = 0
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If GetPlayerAccess(MyIndex) >= 3 Then
                                If frmMapEditor.optClassChange.Value = True Then
                                    .Type = TILE_TYPE_CLASS_CHANGE
                                    .Data1 = ClassChange
                                    .Data2 = ClassChangeReq
                                    .Data3 = 0
                                    .String1 = vbNullString
                                    .String2 = vbNullString
                                    .String3 = vbNullString
                                End If
                            End If
                            If frmMapEditor.optScripted.Value = True Then
                                .Type = TILE_TYPE_SCRIPTED
                                .Data1 = ScriptNum
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.optGuildBlock.Value = True Then
                                .Type = TILE_TYPE_GUILDBLOCK
                                .Data1 = 0
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = GuildBlock
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.optBank.Value = True Then
                                .Type = TILE_TYPE_BANK
                                .Data1 = 0
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.OptGHook.Value = True Then
                                .Type = TILE_TYPE_HOOKSHOT
                                .Data1 = 0
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.optWalkThru.Value = True Then
                                .Type = TILE_TYPE_WALKTHRU
                                .Data1 = 0
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.optClick.Value = True Then
                                .Type = TILE_TYPE_ONCLICK
                                .Data1 = ClickScript
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.optMinusStat.Value = True Then
                                .Type = TILE_TYPE_LOWER_STAT
                                .Data1 = MinusHp
                                .Data2 = MinusMp
                                .Data3 = MinusSp
                                .String1 = MessageMinus
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.optSwitch.Value = True Then
                                .Type = TILE_TYPE_SWITCH
                                .Data1 = SwitchWarpMap
                                .Data2 = SwitchWarpPos
                                .Data3 = SwitchWarpFlags
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.optLevelBlock.Value = True Then
                                .Type = TILE_TYPE_LVLBLOCK
                                .Data1 = LevelToBlock
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.optDrill.Value = True Then
                                .Type = TILE_TYPE_DRILL
                                .Data1 = EditorWarpMap
                                .Data2 = EditorWarpX
                                .Data3 = EditorWarpY
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.optJumpBlock.Value = True Then
                                .Type = TILE_TYPE_JUMPBLOCK
                                .Data1 = JumpHeight
                                .Data2 = JumpDecrease
                                .Data3 = 0
                                .String1 = JumpDir(1) & "," & JumpDir(2) & "," & JumpDir(3) & "," & JumpDir(4)
                                .String2 = JumpDirAddHeight(1) & "," & JumpDirAddHeight(2) & "," & JumpDirAddHeight(3) & "," & JumpDirAddHeight(4)
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.optDodgeBill.Value = True Then
                                .Type = TILE_TYPE_DODGEBILL
                                .Data1 = 0
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.optHammerBarrage.Value = True Then
                                .Type = TILE_TYPE_HAMMERBARRAGE
                                .Data1 = EditorWarpMap
                                .Data2 = EditorWarpX
                                .Data3 = EditorWarpY
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.optJugemsCloud.Value = True Then
                                .Type = TILE_TYPE_JUGEMSCLOUD
                                .Data1 = CloudDir
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.optSimulBlock.Value = True Then
                                .Type = TILE_TYPE_SIMULBLOCK
                                .Data1 = SimulBlockWarpCoords(1)
                                .Data2 = SimulBlockWarpCoords(2)
                                .Data3 = 0
                                .String1 = SimulBlockCoords(1) & "/" & SimulBlockCoords(2) & "/" & SimulBlockCoords(3) & "/" & SimulBlockCoords(4) & "/"
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.optBean.Value = True Then
                                .Type = TILE_TYPE_BEAN
                                .Data1 = BeanItemNum
                                .Data2 = BeanItemQuantity
                                .Data3 = 0
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                        End With
                        If frmMapEditor.optQuestionBlock.Value = True Then
                            Map(GetPlayerMap(MyIndex)).Tile(x, y).Type = TILE_TYPE_QUESTIONBLOCK
                            
                            With QuestionBlock(GetPlayerMap(MyIndex), x, y)
                                .Item1 = ItemThing1
                                .Item2 = ItemThing2
                                .Item3 = ItemThing3
                                .Item4 = ItemThing4
                                .Item5 = ItemThing5
                                .Item6 = ItemThing6
                                .Chance1 = ChanceThing1
                                .Chance2 = ChanceThing2
                                .Chance3 = ChanceThing3
                                .Chance4 = ChanceThing4
                                .Chance5 = ChanceThing5
                                .Chance6 = ChanceThing6
                                .Value1 = ValueThing1
                                .Value2 = ValueThing2
                                .Value3 = ValueThing3
                                .Value4 = ValueThing4
                                .Value5 = ValueThing5
                                .Value6 = ValueThing6
                            End With
                        End If
                    End If
                Else
                    For y2 = 0 To (frmMapEditor.shpSelected.Height \ PIC_Y) - 1
                        For x2 = 0 To (frmMapEditor.shpSelected.Width \ PIC_X) - 1
                            If x + x2 <= MAX_MAPX Then
                                If y + y2 <= MAX_MAPY Then
                                    If MapEditorSelectedType = 1 Then
                                        With Map(GetPlayerMap(MyIndex)).Tile(x + x2, y + y2)
                                            If frmMapEditor.optGround.Value = True Then
                                                .Ground = (EditorTileY + y2) * TilesInSheets + (EditorTileX + x2)
                                                .GroundSet = EditorSet
                                            End If
                                            If frmMapEditor.optMask.Value = True Then
                                                .Mask = (EditorTileY + y2) * TilesInSheets + (EditorTileX + x2)
                                                .MaskSet = EditorSet
                                            End If
                                            If frmMapEditor.optAnim.Value = True Then
                                                .Anim = (EditorTileY + y2) * TilesInSheets + (EditorTileX + x2)
                                                .AnimSet = EditorSet
                                            End If
                                            If frmMapEditor.optMask2.Value = True Then
                                                .Mask2 = (EditorTileY + y2) * TilesInSheets + (EditorTileX + x2)
                                                .Mask2Set = EditorSet
                                            End If
                                            If frmMapEditor.optM2Anim.Value = True Then
                                                .M2Anim = (EditorTileY + y2) * TilesInSheets + (EditorTileX + x2)
                                                .M2AnimSet = EditorSet
                                            End If
                                            If frmMapEditor.optFringe.Value = True Then
                                                .Fringe = (EditorTileY + y2) * TilesInSheets + (EditorTileX + x2)
                                                .FringeSet = EditorSet
                                            End If
                                            If frmMapEditor.optFAnim.Value = True Then
                                                .FAnim = (EditorTileY + y2) * TilesInSheets + (EditorTileX + x2)
                                                .FAnimSet = EditorSet
                                            End If
                                            If frmMapEditor.optFringe2.Value = True Then
                                                .Fringe2 = (EditorTileY + y2) * TilesInSheets + (EditorTileX + x2)
                                                .Fringe2Set = EditorSet
                                            End If
                                            If frmMapEditor.optF2Anim.Value = True Then
                                                .F2Anim = (EditorTileY + y2) * TilesInSheets + (EditorTileX + x2)
                                                .F2AnimSet = EditorSet
                                            End If
                                        End With
                                    ElseIf MapEditorSelectedType = 3 Then
                                        Map(GetPlayerMap(MyIndex)).Tile(x + x2, y + y2).light = (EditorTileY + y2) * TilesInSheets + (EditorTileX + x2)
                                    End If
                                End If
                            End If
                        Next x2
                    Next y2
                End If
            End If

            If (Button = 2) And (x >= 0) And (x <= MAX_MAPX) And (y >= 0) And (y <= MAX_MAPY) Then
                If MapEditorSelectedType = 1 Then
                    With Map(GetPlayerMap(MyIndex)).Tile(x, y)
                        If frmMapEditor.optGround.Value = True Then
                            .Ground = 0
                        End If
                        If frmMapEditor.optMask.Value = True Then
                            .Mask = 0
                        End If
                        If frmMapEditor.optAnim.Value = True Then
                            .Anim = 0
                        End If
                        If frmMapEditor.optMask2.Value = True Then
                            .Mask2 = 0
                        End If
                        If frmMapEditor.optM2Anim.Value = True Then
                            .M2Anim = 0
                        End If
                        If frmMapEditor.optFringe.Value = True Then
                            .Fringe = 0
                        End If
                        If frmMapEditor.optFAnim.Value = True Then
                            .FAnim = 0
                        End If
                        If frmMapEditor.optFringe2.Value = True Then
                            .Fringe2 = 0
                        End If
                        If frmMapEditor.optF2Anim.Value = True Then
                            .F2Anim = 0
                        End If
                    End With
                ElseIf MapEditorSelectedType = 3 Then
                    Map(GetPlayerMap(MyIndex)).Tile(x, y).light = 0
                ElseIf MapEditorSelectedType = 2 Then
                    With Map(GetPlayerMap(MyIndex)).Tile(x, y)
                        .Type = 0
                        .Data1 = 0
                        .Data2 = 0
                        .Data3 = 0
                        .String1 = vbNullString
                        .String2 = vbNullString
                        .String3 = vbNullString
                    End With
                End If
            End If
        End If
    End If
End Sub

Public Sub EditorChooseTile(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        EditorTileX = (x \ PIC_X)
        EditorTileY = (y \ PIC_Y)
    End If
    frmMapEditor.shpSelected.Top = Int(EditorTileY * PIC_Y)
    frmMapEditor.shpSelected.Left = Int(EditorTileX * PIC_Y)
End Sub

Public Sub EditorTileScroll()
    frmMapEditor.picBackSelect.Top = (frmMapEditor.scrlPicture.Value * PIC_Y) * -1
End Sub

Public Sub EditorSend()
    Call SendMap
    Call EditorCancel
End Sub

Public Sub EditorCancel()
    ScreenMode = 0
    GridMode = 0

    ' Set the type back to default.
    MapEditorSelectedType = 1

    ' Set the map controls to default.
    frmMapEditor.fraAttribs.Visible = False
    frmMapEditor.fraLayers.Visible = True
    frmMapEditor.frmtile.Visible = True

    InEditor = False
    frmMapEditor.Visible = False

    frmMirage.Show
    frmMapEditor.MousePointer = 1
    frmMirage.MousePointer = 1

    Call LoadMap(GetPlayerMap(MyIndex))
End Sub

Public Sub EditorClearLayer()
    Dim Choice As Integer
    Dim x As Byte
    Dim y As Byte

    ' Ground Layer
    If frmMapEditor.optGround.Value Then
        Choice = MsgBox("Are you sure you wish to clear the ground layer?", vbYesNo, "Super Mario Bros. Online")

        If Choice = vbYes Then
            For y = 0 To MAX_MAPY
                For x = 0 To MAX_MAPX
                    Map(GetPlayerMap(MyIndex)).Tile(x, y).Ground = 0
                    Map(GetPlayerMap(MyIndex)).Tile(x, y).GroundSet = 0
                Next x
            Next y
        End If
    End If

    ' Mask Layer
    If frmMapEditor.optMask.Value Then
        Choice = MsgBox("Are you sure you wish to clear the mask layer?", vbYesNo, "Super Mario Bros. Online")

        If Choice = vbYes Then
            For y = 0 To MAX_MAPY
                For x = 0 To MAX_MAPX
                    Map(GetPlayerMap(MyIndex)).Tile(x, y).Mask = 0
                    Map(GetPlayerMap(MyIndex)).Tile(x, y).MaskSet = 0
                Next x
            Next y
        End If
    End If

    ' Mask Animation Layer
    If frmMapEditor.optAnim.Value Then
        Choice = MsgBox("Are you sure you wish to clear the animation layer?", vbYesNo, "Super Mario Bros. Online")

        If Choice = vbYes Then
            For y = 0 To MAX_MAPY
                For x = 0 To MAX_MAPX
                    Map(GetPlayerMap(MyIndex)).Tile(x, y).Anim = 0
                    Map(GetPlayerMap(MyIndex)).Tile(x, y).AnimSet = 0
                Next x
            Next y
        End If
    End If

    ' Mask 2 Layer
    If frmMapEditor.optMask2.Value Then
        Choice = MsgBox("Are you sure you wish to clear the mask 2 layer?", vbYesNo, "Super Mario Bros. Online")

        If Choice = vbYes Then
            For y = 0 To MAX_MAPY
                For x = 0 To MAX_MAPX
                    Map(GetPlayerMap(MyIndex)).Tile(x, y).Mask2 = 0
                    Map(GetPlayerMap(MyIndex)).Tile(x, y).Mask2Set = 0
                Next x
            Next y
        End If
    End If

    ' Mask 2 Animation layer
    If frmMapEditor.optM2Anim.Value Then
        Choice = MsgBox("Are you sure you wish to clear the mask 2 animation layer?", vbYesNo, "Super Mario Bros. Online")

        If Choice = vbYes Then
            For y = 0 To MAX_MAPY
                For x = 0 To MAX_MAPX
                    Map(GetPlayerMap(MyIndex)).Tile(x, y).M2Anim = 0
                    Map(GetPlayerMap(MyIndex)).Tile(x, y).M2AnimSet = 0
                Next x
            Next y
        End If
    End If

    ' Fringe Layer
    If frmMapEditor.optFringe.Value Then
        Choice = MsgBox("Are you sure you wish to clear the fringe layer?", vbYesNo, "Super Mario Bros. Online")

        If Choice = vbYes Then
            For y = 0 To MAX_MAPY
                For x = 0 To MAX_MAPX
                    Map(GetPlayerMap(MyIndex)).Tile(x, y).Fringe = 0
                    Map(GetPlayerMap(MyIndex)).Tile(x, y).FringeSet = 0
                Next x
            Next y
        End If
    End If

    ' Fringe Animation Layer
    If frmMapEditor.optFAnim.Value Then
        Choice = MsgBox("Are you sure you wish to clear the fringe animation layer?", vbYesNo, "Super Mario Bros. Online")

        If Choice = vbYes Then
            For y = 0 To MAX_MAPY
                For x = 0 To MAX_MAPX
                    Map(GetPlayerMap(MyIndex)).Tile(x, y).FAnim = 0
                    Map(GetPlayerMap(MyIndex)).Tile(x, y).FAnimSet = 0
                Next x
            Next y
        End If
    End If

    ' Fringe 2 Layer
    If frmMapEditor.optFringe2.Value Then
        Choice = MsgBox("Are you sure you wish to clear the fringe 2 layer?", vbYesNo, "Super Mario Bros. Online")

        If Choice = vbYes Then
            For y = 0 To MAX_MAPY
                For x = 0 To MAX_MAPX
                    Map(GetPlayerMap(MyIndex)).Tile(x, y).Fringe2 = 0
                    Map(GetPlayerMap(MyIndex)).Tile(x, y).Fringe2Set = 0
                Next x
            Next y
        End If
    End If

    ' Fringe 2 Animation Layer
    If frmMapEditor.optF2Anim.Value Then
        Choice = MsgBox("Are you sure you wish to clear the fringe 2 animation layer?", vbYesNo, "Super Mario Bros. Online")

        If Choice = vbYes Then
            For y = 0 To MAX_MAPY
                For x = 0 To MAX_MAPX
                    Map(GetPlayerMap(MyIndex)).Tile(x, y).F2Anim = 0
                    Map(GetPlayerMap(MyIndex)).Tile(x, y).F2AnimSet = 0
                Next x
            Next y
        End If
    End If
End Sub

Public Sub EditorClearAttribs()
    Dim Choice As Integer
    Dim x As Byte
    Dim y As Byte

    Choice = MsgBox("Are you sure you wish to clear the attributes on this map?", vbYesNo, "Super Mario Bros. Online")

    If Choice = vbYes Then
        For y = 0 To MAX_MAPY
            For x = 0 To MAX_MAPX
                Map(GetPlayerMap(MyIndex)).Tile(x, y).Type = 0
            Next x
        Next y
    End If
End Sub

Public Sub EmoticonEditorInit()
    frmEmoticonEditor.scrlEmoticon.Max = MAX_EMOTICONS
    frmEmoticonEditor.scrlEmoticon.Value = Emoticons(EditorIndex - 1).Pic
    frmEmoticonEditor.txtCommand.Text = Trim$(Emoticons(EditorIndex - 1).Command)
    
    Call InitLoadPicture(App.Path & "\GFX\Emoticons.smbo", frmEmoticonEditor.picEmoticons)
    
    frmEmoticonEditor.Show vbModal
End Sub

Public Sub ElementEditorInit()
    frmElementEditor.txtName.Text = Trim$(Element(EditorIndex - 1).Name)
    frmElementEditor.scrlStrong.Value = Element(EditorIndex - 1).Strong
    frmElementEditor.scrlWeak.Value = Element(EditorIndex - 1).Weak
    frmElementEditor.Show vbModal
End Sub

Public Sub EmoticonEditorOk()
    Emoticons(EditorIndex - 1).Pic = frmEmoticonEditor.scrlEmoticon.Value
    If frmEmoticonEditor.txtCommand.Text <> "/" Then
        Emoticons(EditorIndex - 1).Command = frmEmoticonEditor.txtCommand.Text
    Else
        Emoticons(EditorIndex - 1).Command = vbNullString
    End If

    Call SendSaveEmoticon(EditorIndex - 1)
    Call EmoticonEditorCancel
End Sub

Public Sub ElementEditorOk()
    Element(EditorIndex - 1).Name = frmElementEditor.txtName.Text
    Element(EditorIndex - 1).Strong = frmElementEditor.scrlStrong.Value
    Element(EditorIndex - 1).Weak = frmElementEditor.scrlWeak.Value
    Call SendSaveElement(EditorIndex - 1)
    Call ElementEditorCancel
End Sub

Public Sub EmoticonEditorCancel()
    InEmoticonEditor = False
    Unload frmEmoticonEditor
End Sub

Public Sub ElementEditorCancel()
    InElementEditor = False
    Unload frmElementEditor
End Sub

Public Sub ArrowEditorInit()
    frmEditArrows.scrlArrow.Max = MAX_ARROWS
    If Arrows(EditorIndex).Pic = 0 Then
        Arrows(EditorIndex).Pic = 1
    End If
    frmEditArrows.scrlArrow.Value = Arrows(EditorIndex).Pic
    frmEditArrows.txtName.Text = Arrows(EditorIndex).Name
    If Arrows(EditorIndex).Range = 0 Then
        Arrows(EditorIndex).Range = 1
    End If
    frmEditArrows.scrlRange.Value = Arrows(EditorIndex).Range
    If Arrows(EditorIndex).Amount = 0 Then
        Arrows(EditorIndex).Amount = 1
    End If
    frmEditArrows.scrlAmount.Value = Arrows(EditorIndex).Amount
    
    Call InitLoadPicture(App.Path & "\GFX\Arrows.smbo", frmEditArrows.picArrows)

    frmEditArrows.Show vbModal
End Sub

Public Sub ArrowEditorOk()
    Arrows(EditorIndex).Pic = frmEditArrows.scrlArrow.Value
    Arrows(EditorIndex).Range = frmEditArrows.scrlRange.Value
    Arrows(EditorIndex).Name = frmEditArrows.txtName.Text
    Arrows(EditorIndex).Amount = frmEditArrows.scrlAmount.Value
    Call SendSaveArrow(EditorIndex)
    Call ArrowEditorCancel
End Sub

Public Sub ArrowEditorCancel()
    InArrowEditor = False
    Unload frmEditArrows
End Sub

Public Sub ItemEditorInit()
    Dim i As Long
    
    EditorItemY = (Item(EditorIndex).Pic \ 6)
    EditorItemX = (Item(EditorIndex).Pic - (Item(EditorIndex).Pic \ 6) * 6)

    frmItemEditor.scrlClassReq.Max = MAX_CLASSES
    
    Call InitLoadPicture(App.Path & "\GFX\Items.smbo", frmItemEditor.picItems)
    
    frmItemEditor.txtName.Text = Trim$(Item(EditorIndex).Name)
    frmItemEditor.txtDesc.Text = Trim$(Item(EditorIndex).desc)
    frmItemEditor.cmbType.ListIndex = Item(EditorIndex).Type
    frmItemEditor.txtPrice.Text = Item(EditorIndex).Price
    frmItemEditor.chkStackable.Value = Item(EditorIndex).Stackable
    frmItemEditor.chkBound.Value = Item(EditorIndex).Bound
    If Item(EditorIndex).Cookable = True Then
        frmItemEditor.chkCookable.Value = Checked
    Else
        frmItemEditor.chkCookable.Value = Unchecked
    End If
    
    If (frmItemEditor.cmbType.ListIndex >= ITEM_TYPE_WEAPON) And (frmItemEditor.cmbType.ListIndex <= ITEM_TYPE_MUSHROOMBADGE) Then
        frmItemEditor.fraEquipment.Visible = True
        frmItemEditor.fraAttributes.Visible = True
        If frmItemEditor.cmbType.ListIndex = ITEM_TYPE_WEAPON Then
            frmItemEditor.fraBow.Visible = True
        End If

        frmItemEditor.scrlHPReq.Value = Item(EditorIndex).HPReq
        frmItemEditor.scrlFPReq.Value = Item(EditorIndex).FPReq
        frmItemEditor.scrlStrReq.Value = Item(EditorIndex).StrReq
        frmItemEditor.scrlDefReq.Value = Item(EditorIndex).DefReq
        frmItemEditor.scrlSpeedReq.Value = Item(EditorIndex).SpeedReq
        frmItemEditor.scrlMagicReq.Value = Item(EditorIndex).MagicReq
        frmItemEditor.scrlClassReq.Value = Item(EditorIndex).ClassReq
        frmItemEditor.scrlLevelReq.Value = Item(EditorIndex).LevelReq
        frmItemEditor.scrlAccessReq.Value = Item(EditorIndex).AccessReq
        frmItemEditor.scrlAddHP.Value = Item(EditorIndex).AddHP
        frmItemEditor.scrlAddMP.Value = Item(EditorIndex).AddMP
        frmItemEditor.scrlAddSP.Value = Item(EditorIndex).AddSP
        frmItemEditor.scrlAddStr.Value = Item(EditorIndex).AddSTR
        frmItemEditor.scrlAddDef.Value = Item(EditorIndex).AddDef
        frmItemEditor.scrlAddMagi.Value = Item(EditorIndex).AddMAGI
        frmItemEditor.scrlAddSpeed.Value = Item(EditorIndex).AddSpeed
        frmItemEditor.scrlAddEXP.Value = Item(EditorIndex).AddEXP
        frmItemEditor.scrlAttackSpeed.Value = Item(EditorIndex).AttackSpeed
        frmItemEditor.scrlAddCritHit.Value = (Item(EditorIndex).AddCritChance * 10)
        frmItemEditor.scrlAddBlockChance.Value = (Item(EditorIndex).AddBlockChance * 10)
        If Item(EditorIndex).Ammo < 0 Then
            frmItemEditor.scrlAmmo.Value = 1
        Else
            frmItemEditor.scrlAmmo.Value = Item(EditorIndex).Ammo
        End If

        If Item(EditorIndex).Data3 > 0 Then
            If Item(EditorIndex).Stackable = 2 And Item(EditorIndex).Type <> ITEM_TYPE_AMMO Then
                frmItemEditor.chkBow.Value = Checked
                frmItemEditor.chkGrapple.Value = Checked
                frmItemEditor.chkAmmo.Value = Unchecked
            Else
                frmItemEditor.chkBow.Value = Checked
                frmItemEditor.chkGrapple.Value = Unchecked
            End If
            If Item(EditorIndex).Ammo > -1 Then
                frmItemEditor.chkAmmo.Value = Checked
                frmItemEditor.chkBow.Value = Checked
            End If
        Else
            frmItemEditor.chkBow.Value = Unchecked
        End If

        frmItemEditor.cmbBow.Clear
        If frmItemEditor.chkBow.Value = Checked Then
            For i = 1 To 100
                frmItemEditor.cmbBow.addItem i & ": " & Arrows(i).Name
            Next i
            frmItemEditor.cmbBow.ListIndex = Item(EditorIndex).Data3 - 1
            frmItemEditor.picBow.Top = (Arrows(Item(EditorIndex).Data3).Pic * 32) * -1
            frmItemEditor.cmbBow.Enabled = True
        Else
            frmItemEditor.cmbBow.addItem "None"
            frmItemEditor.cmbBow.ListIndex = 0
            frmItemEditor.cmbBow.Enabled = False
        End If
        frmItemEditor.chkStackable.Visible = False
    Else
        frmItemEditor.fraEquipment.Visible = False
    End If

    If (frmItemEditor.cmbType.ListIndex = ITEM_TYPE_CHANGEHPFPSP) Then
        frmItemEditor.fraVitals.Visible = True
        frmItemEditor.chkStackable.Visible = True
        frmItemEditor.scrlChangeHP.Value = Item(EditorIndex).Data1
        frmItemEditor.scrlChangeFP.Value = Item(EditorIndex).Data2
        frmItemEditor.scrlChangeSP.Value = Item(EditorIndex).Data3
    Else
        frmItemEditor.fraVitals.Visible = False
    End If

    If (frmItemEditor.cmbType.ListIndex = ITEM_TYPE_SPELL) Then
        frmItemEditor.fraSpell.Visible = True
        frmItemEditor.scrlSpell.Value = Item(EditorIndex).Data1
        frmItemEditor.chkStackable.Visible = False
    Else
        frmItemEditor.fraSpell.Visible = False
    End If

    If (frmItemEditor.cmbType.ListIndex >= ITEM_TYPE_SCRIPTED) Then
        frmItemEditor.fraScript.Visible = True
        frmItemEditor.scrlScript.Value = Item(EditorIndex).Data1
        frmItemEditor.lblScript.Caption = Item(EditorIndex).Data1
        
        frmItemEditor.chkStackable.Visible = True
    Else
        frmItemEditor.fraScript.Visible = False
    End If
    If (frmItemEditor.cmbType.ListIndex >= ITEM_TYPE_AMMO) Then
        frmItemEditor.fraScript.Visible = False
        frmItemEditor.chkGrapple.Visible = False
        frmItemEditor.chkAmmo.Visible = False
        frmItemEditor.chkStackable.Visible = False
        frmItemEditor.chkStackable.Value = Checked
        frmItemEditor.fraBow.Visible = True
        frmItemEditor.chkGrapple.Value = Unchecked
        frmItemEditor.chkAmmo.Value = Unchecked
        
        frmItemEditor.cmbBow.Clear
        If frmItemEditor.chkBow.Value = Checked Then
            For i = 1 To 100
                frmItemEditor.cmbBow.addItem i & ": " & Arrows(i).Name
            Next i
            frmItemEditor.cmbBow.ListIndex = Item(EditorIndex).Data3 - 1
            frmItemEditor.picBow.Top = (Arrows(Item(EditorIndex).Data3).Pic * 32) * -1
            frmItemEditor.cmbBow.Enabled = True
        Else
            frmItemEditor.cmbBow.addItem "None"
            frmItemEditor.cmbBow.ListIndex = 0
            frmItemEditor.cmbBow.Enabled = False
        End If
    End If
    If (frmItemEditor.cmbType.ListIndex = ITEM_TYPE_CARD) Then
        frmItemEditor.fraBow.Visible = False
        frmItemEditor.chkStackable.Visible = False
        frmItemEditor.chkStackable.Value = Checked
        frmItemEditor.chkBound.Visible = True
        frmItemEditor.chkBound.Value = Item(EditorIndex).Bound
    End If
    
    frmItemEditor.VScroll1.Value = EditorItemY
    frmItemEditor.picItems.Top = (EditorItemY) * -32
    frmItemEditor.Show vbModal
End Sub

Public Sub ItemEditorOk()
    Item(EditorIndex).Name = frmItemEditor.txtName.Text
    Item(EditorIndex).desc = frmItemEditor.txtDesc.Text
    Item(EditorIndex).Pic = EditorItemY * 6 + EditorItemX
    Item(EditorIndex).Type = frmItemEditor.cmbType.ListIndex
    Item(EditorIndex).Price = Val(frmItemEditor.txtPrice.Text)
    Item(EditorIndex).Bound = frmItemEditor.chkBound.Value
    
    Item(EditorIndex).AddHP = 0
    Item(EditorIndex).AddMP = 0
    Item(EditorIndex).AddSP = 0
    Item(EditorIndex).AddSTR = 0
    Item(EditorIndex).AddDef = 0
    Item(EditorIndex).AddMAGI = 0
    Item(EditorIndex).AddSpeed = 0
    Item(EditorIndex).AddEXP = 0
    Item(EditorIndex).AttackSpeed = 0
    Item(EditorIndex).AddCritChance = 0
    Item(EditorIndex).AddBlockChance = 0
    Item(EditorIndex).Ammo = -1
    
    If frmItemEditor.chkCookable.Value = Checked Then
        Item(EditorIndex).Cookable = True
    Else
        Item(EditorIndex).Cookable = False
    End If
    
    Select Case frmItemEditor.cmbType.ListIndex
        Case ITEM_TYPE_WEAPON, ITEM_TYPE_TWO_HAND, ITEM_TYPE_ARMOR, ITEM_TYPE_HELMET, ITEM_TYPE_SPECIALBADGE, ITEM_TYPE_LEGS, ITEM_TYPE_FLOWERBADGE, ITEM_TYPE_MUSHROOMBADGE
            Item(EditorIndex).HPReq = frmItemEditor.scrlHPReq.Value
            Item(EditorIndex).FPReq = frmItemEditor.scrlFPReq.Value
            If frmItemEditor.chkBow.Value = Checked Then
                If frmItemEditor.chkAmmo.Value = Checked Then
                    Item(EditorIndex).Data3 = frmItemEditor.cmbBow.ListIndex + 1
                    Item(EditorIndex).Stackable = 0
                    Item(EditorIndex).Ammo = frmItemEditor.scrlAmmo.Value
                ElseIf frmItemEditor.chkGrapple.Value = Checked Then
                    Item(EditorIndex).Data3 = frmItemEditor.cmbBow.ListIndex + 1
                    Item(EditorIndex).Stackable = 2
                    Item(EditorIndex).Ammo = -1
                Else
                    Item(EditorIndex).Data3 = frmItemEditor.cmbBow.ListIndex + 1
                    Item(EditorIndex).Stackable = 0
                    Item(EditorIndex).Ammo = -1
                End If
            Else
                Item(EditorIndex).Data3 = 0
                Item(EditorIndex).Stackable = 0
                Item(EditorIndex).Ammo = -1
            End If
            Item(EditorIndex).StrReq = frmItemEditor.scrlStrReq.Value
            Item(EditorIndex).DefReq = frmItemEditor.scrlDefReq.Value
            Item(EditorIndex).SpeedReq = frmItemEditor.scrlSpeedReq.Value
            Item(EditorIndex).MagicReq = frmItemEditor.scrlMagicReq.Value
            Item(EditorIndex).ClassReq = frmItemEditor.scrlClassReq.Value
            Item(EditorIndex).AccessReq = frmItemEditor.scrlAccessReq.Value
            Item(EditorIndex).LevelReq = frmItemEditor.scrlLevelReq.Value

            Item(EditorIndex).AddHP = frmItemEditor.scrlAddHP.Value
            Item(EditorIndex).AddMP = frmItemEditor.scrlAddMP.Value
            Item(EditorIndex).AddSP = frmItemEditor.scrlAddSP.Value
            Item(EditorIndex).AddSTR = frmItemEditor.scrlAddStr.Value
            Item(EditorIndex).AddDef = frmItemEditor.scrlAddDef.Value
            Item(EditorIndex).AddMAGI = frmItemEditor.scrlAddMagi.Value
            Item(EditorIndex).AddSpeed = frmItemEditor.scrlAddSpeed.Value
            Item(EditorIndex).AddEXP = frmItemEditor.scrlAddEXP.Value
            Item(EditorIndex).AttackSpeed = frmItemEditor.scrlAttackSpeed.Value
            Item(EditorIndex).AddCritChance = (frmItemEditor.scrlAddCritHit.Value / 10)
            Item(EditorIndex).AddBlockChance = (frmItemEditor.scrlAddBlockChance.Value / 10)
        Case ITEM_TYPE_CHANGEHPFPSP
            Item(EditorIndex).Data1 = frmItemEditor.scrlChangeHP.Value
            Item(EditorIndex).Data2 = frmItemEditor.scrlChangeFP.Value
            Item(EditorIndex).Data3 = frmItemEditor.scrlChangeSP.Value
            Item(EditorIndex).StrReq = 0
            Item(EditorIndex).DefReq = 0
            Item(EditorIndex).SpeedReq = 0
            Item(EditorIndex).MagicReq = 0
            Item(EditorIndex).ClassReq = -1
            Item(EditorIndex).AccessReq = 0
            Item(EditorIndex).LevelReq = 0
    
            Item(EditorIndex).Stackable = frmItemEditor.chkStackable.Value
        Case ITEM_TYPE_NONE
            Item(EditorIndex).Stackable = frmItemEditor.chkStackable.Value
            Item(EditorIndex).Ammo = -1
        Case ITEM_TYPE_SPELL
            Item(EditorIndex).Data1 = frmItemEditor.scrlSpell.Value
            Item(EditorIndex).Data2 = 0
            Item(EditorIndex).Data3 = 0
            Item(EditorIndex).StrReq = 0
            Item(EditorIndex).DefReq = 0
            Item(EditorIndex).SpeedReq = 0
            Item(EditorIndex).MagicReq = 0
            Item(EditorIndex).ClassReq = -1
            Item(EditorIndex).AccessReq = 0
            Item(EditorIndex).LevelReq = 0
    
            Item(EditorIndex).Stackable = frmItemEditor.chkStackable.Value
        Case ITEM_TYPE_SCRIPTED
            Item(EditorIndex).Data1 = frmItemEditor.scrlScript.Value
            Item(EditorIndex).Data2 = 0
            Item(EditorIndex).Data3 = 0
            Item(EditorIndex).StrReq = 0
            Item(EditorIndex).DefReq = 0
            Item(EditorIndex).SpeedReq = 0
            Item(EditorIndex).MagicReq = 0
            Item(EditorIndex).ClassReq = -1
            Item(EditorIndex).AccessReq = 0
            Item(EditorIndex).LevelReq = 0
            
            Item(EditorIndex).Stackable = frmItemEditor.chkStackable.Value
        Case ITEM_TYPE_AMMO
            Item(EditorIndex).Data1 = 0
            Item(EditorIndex).Data2 = 0
            Item(EditorIndex).Data3 = frmItemEditor.cmbBow.ListIndex + 1
            Item(EditorIndex).StrReq = 0
            Item(EditorIndex).DefReq = 0
            Item(EditorIndex).SpeedReq = 0
            Item(EditorIndex).MagicReq = 0
            Item(EditorIndex).ClassReq = -1
            Item(EditorIndex).AccessReq = 0
            Item(EditorIndex).LevelReq = 0
            Item(EditorIndex).Stackable = 1
            Item(EditorIndex).Cookable = False
        Case ITEM_TYPE_CARD
            Item(EditorIndex).Data1 = 0
            Item(EditorIndex).Data2 = 0
            Item(EditorIndex).Data3 = 0
            Item(EditorIndex).StrReq = 0
            Item(EditorIndex).DefReq = 0
            Item(EditorIndex).SpeedReq = 0
            Item(EditorIndex).MagicReq = 0
            Item(EditorIndex).ClassReq = -1
            Item(EditorIndex).AccessReq = 0
            Item(EditorIndex).LevelReq = 0
            
            Item(EditorIndex).Stackable = 1
            Item(EditorIndex).Bound = frmItemEditor.chkBound.Value
            Item(EditorIndex).Cookable = False
        Case Else
            Item(EditorIndex).Ammo = -1
    End Select
    
    Call SendSaveItem(EditorIndex)
    InItemsEditor = False
    Unload frmItemEditor
End Sub

Public Sub ItemEditorCancel()
    InItemsEditor = False
    Unload frmItemEditor
End Sub

Public Sub NpcEditorInit()
    On Error Resume Next
    
    Call InitLoadPicture(App.Path & "\GFX\Sprites.smbo", frmNpcEditor.picSprites)
    
    frmNpcEditor.txtName.Text = Trim$(Npc(EditorIndex).Name)
    frmNpcEditor.txtAttackSay.Text = Trim$(Npc(EditorIndex).AttackSay)
    frmNpcEditor.txtAttackSay2.Text = Trim$(Npc(EditorIndex).AttackSay2)
    frmNpcEditor.scrlSprite.Value = Npc(EditorIndex).Sprite
    frmNpcEditor.txtSpawnSecs.Text = STR(Npc(EditorIndex).SpawnSecs)
    frmNpcEditor.cmbBehavior.ListIndex = Npc(EditorIndex).Behavior
    frmNpcEditor.scrlRange.Value = Npc(EditorIndex).Range
    frmNpcEditor.scrlSTR.Value = Npc(EditorIndex).STR
    frmNpcEditor.scrlDEF.Value = Npc(EditorIndex).DEF
    frmNpcEditor.scrlSPEED.Value = Npc(EditorIndex).speed
    frmNpcEditor.scrlMAGI.Value = Npc(EditorIndex).MAGI
    frmNpcEditor.BigNpc.Value = Npc(EditorIndex).Big
    frmNpcEditor.StartHP.Value = Npc(EditorIndex).MaxHp
    frmNpcEditor.ExpGive.Value = Npc(EditorIndex).Exp
    frmNpcEditor.scrlChance.Value = Npc(EditorIndex).ItemNPC(1).chance
    frmNpcEditor.scrlNum.Value = Npc(EditorIndex).ItemNPC(1).ItemNum
    frmNpcEditor.scrlValue.Value = Npc(EditorIndex).ItemNPC(1).ItemValue
    frmNpcEditor.scrlLevel.Value = Npc(EditorIndex).Level
    If Npc(EditorIndex).Behavior = NPC_BEHAVIOR_SCRIPTED Then
        frmNpcEditor.scrlScript.Value = Npc(EditorIndex).SpawnSecs
        frmNpcEditor.scrlElement.Value = Npc(EditorIndex).Element
    End If
    
    If Npc(EditorIndex).SpawnTime = 0 Then
        frmNpcEditor.chkDay.Value = Checked
        frmNpcEditor.chkNight.Value = Checked
    End If

    frmNpcEditor.Show vbModal
End Sub

Public Sub NpcEditorOk()
    Npc(EditorIndex).Name = frmNpcEditor.txtName.Text
    Npc(EditorIndex).AttackSay = frmNpcEditor.txtAttackSay.Text
    Npc(EditorIndex).AttackSay2 = frmNpcEditor.txtAttackSay2.Text
    Npc(EditorIndex).Sprite = frmNpcEditor.scrlSprite.Value
    Npc(EditorIndex).Behavior = frmNpcEditor.cmbBehavior.ListIndex
    
    If Npc(EditorIndex).Behavior <> NPC_BEHAVIOR_SCRIPTED Then
        Npc(EditorIndex).SpawnSecs = Val(frmNpcEditor.txtSpawnSecs.Text)
    Else
        Npc(EditorIndex).SpawnSecs = frmNpcEditor.scrlScript.Value
    End If
    
    Npc(EditorIndex).Range = frmNpcEditor.scrlRange.Value
    Npc(EditorIndex).STR = frmNpcEditor.scrlSTR.Value
    Npc(EditorIndex).DEF = frmNpcEditor.scrlDEF.Value
    Npc(EditorIndex).speed = frmNpcEditor.scrlSPEED.Value
    Npc(EditorIndex).MAGI = frmNpcEditor.scrlMAGI.Value
    Npc(EditorIndex).Big = frmNpcEditor.BigNpc.Value
    Npc(EditorIndex).MaxHp = frmNpcEditor.StartHP.Value
    Npc(EditorIndex).Exp = frmNpcEditor.ExpGive.Value
    Npc(EditorIndex).Level = frmNpcEditor.scrlLevel.Value

    Npc(EditorIndex).SpriteSize = 1
    Npc(EditorIndex).SpawnTime = 0

    Call SendSaveNPC(EditorIndex)
    InNpcEditor = False
    Unload frmNpcEditor
End Sub

Public Sub NpcEditorCancel()
    InNpcEditor = False
    Unload frmNpcEditor
End Sub

Public Sub NpcEditorBltSprite()
    Dim drec As RECT, srec As RECT
    
    If frmNpcEditor.BigNpc.Value = Checked Then
        frmNpcEditor.picSprites.Top = frmNpcEditor.scrlSprite.Value * 64
        frmNpcEditor.picSprites.Left = 3360
        
        drec.Top = 0
        drec.Bottom = 64
        drec.Left = 0
        drec.Right = 64
        srec.Top = frmNpcEditor.scrlSprite.Value * 64
        srec.Bottom = srec.Top + 64
        srec.Left = 3 * 64
        srec.Right = srec.Left + 64
    
        Call DD_BigSpriteSurf.BltToDC(frmNpcEditor.picSprite.hDC, srec, drec)
    Else
        frmNpcEditor.picSprites.Left = 3600
        frmNpcEditor.picSprites.Top = frmNpcEditor.scrlSprite.Value * 64
        
        drec.Top = 0
        drec.Bottom = 64
        drec.Left = 0
        drec.Right = 32
        srec.Top = frmNpcEditor.scrlSprite.Value * 64
        srec.Bottom = srec.Top + 64
        srec.Left = 3 * PIC_X ' BitBlt xSrc
        srec.Right = srec.Left + PIC_Y
        
        Call DD_SpriteSurf.BltToDC(frmNpcEditor.picSprite.hDC, srec, drec)
    End If
End Sub

' Initializes the shop editor
Public Sub ShopEditorInit()
    Dim i As Integer
    Dim itemN As Integer
    Dim cItemMade As Boolean

    On Error GoTo ShopEditorInit_Error

    frmShopEditor.txtName.Text = Trim$(Shop(EditorIndex).Name)
    frmShopEditor.chkShow.Value = Shop(EditorIndex).ShowInfo
    frmShopEditor.chkSellsItems.Value = Shop(EditorIndex).BuysItems

    cItemMade = False

    frmShopEditor.cmbCurrency.Clear
    frmShopEditor.lstItems.Clear

    ' Add all the currency items to cmbCurrency
    For i = 1 To MAX_ITEMS
        If Item(i).Type = ITEM_TYPE_CURRENCY Then
            ' It's a currency item, so add it to the list
            frmShopEditor.cmbCurrency.addItem (i & " - " & Trim(Item(i).Name))
            ' Add it to the item data so that we know the number
            frmShopEditor.cmbCurrency.ItemData(frmShopEditor.cmbCurrency.ListCount - 1) = i
            cItemMade = True 'we have at least 1 currency item
            If Shop(EditorIndex).currencyItem = i Then
                frmShopEditor.cmbCurrency.ListIndex = frmShopEditor.cmbCurrency.ListCount - 1
            End If
        End If
    Next i

    If Not cItemMade Then
        Call MsgBox("Please make at least one type of currency first!")
        Call ShopEditorCancel
        Exit Sub
    End If

    ' Add all the items to the list
    For i = 1 To MAX_SHOP_ITEMS
        itemN = Shop(EditorIndex).ShopItem(i).ItemNum

        ' If the item is not empty
        If itemN > 0 Then
            ' Add the item to the shop list
            Call frmShopEditor.AddShopItem(itemN, Shop(EditorIndex).ShopItem(i).Price, Shop(EditorIndex).ShopItem(i).currencyItem, Shop(EditorIndex).ShopItem(i).Amount)
        End If
    Next i

    ' Add all items to the 'add item' list
    For i = 1 To MAX_ITEMS
        frmShopEditor.cmbItemList.addItem (i & " - " & Trim(Item(i).Name))
    Next i

    frmShopEditor.frmAddEditItem.Visible = False

    ' Init shop editor temp array
    frmShopEditor.LoadShopItemData (EditorIndex)

    frmShopEditor.Show vbModal

    On Error GoTo 0
    Exit Sub

ShopEditorInit_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ShopEditorInit of Module modGameLogic"
    ' Close the shop editor
    frmShopEditor.Visible = False
    Call ShopEditorCancel
End Sub

Public Sub ShopEditorOk()
    Dim i As Integer
    Dim currencyItem As Integer

    If frmShopEditor.cmbCurrency.ListIndex < 0 Then
        MsgBox "Please pick a currency item!", vbExclamation
        Exit Sub
    End If
    
    currencyItem = frmShopEditor.cmbCurrency.ItemData(frmShopEditor.cmbCurrency.ListIndex)

    Shop(EditorIndex).Name = frmShopEditor.txtName.Text
    Shop(EditorIndex).FixesItems = 0
    Shop(EditorIndex).BuysItems = frmShopEditor.chkSellsItems.Value
    Shop(EditorIndex).ShowInfo = frmShopEditor.chkShow.Value
    Shop(EditorIndex).currencyItem = currencyItem

    For i = 1 To MAX_SHOP_ITEMS
        Shop(EditorIndex).ShopItem(i).Amount = frmShopEditor.GetShopItemAmt(i)
        Shop(EditorIndex).ShopItem(i).ItemNum = frmShopEditor.GetShopItemNum(i)
        Shop(EditorIndex).ShopItem(i).Price = frmShopEditor.GetShopItemPrice(i)
        Shop(EditorIndex).ShopItem(i).currencyItem = frmShopEditor.GetShopCurrencyItem(i)
    Next i

    Call SendSaveShop(EditorIndex)
    InShopEditor = False
    Unload frmShopEditor
End Sub

Public Sub ShopEditorCancel()
    InShopEditor = False
    Unload frmShopEditor
End Sub

Public Sub SpellEditorInit()
    Dim i As Long
    
    Call InitLoadPicture(App.Path & "\GFX\Icons.smbo", frmSpellEditor.iconn)
    
    With frmSpellEditor
        .cmbClassReq.addItem "All Classes"
    
        For i = 0 To MAX_CLASSES
            .cmbClassReq.addItem Trim$(Class(i).Name)
        Next i

        .txtName.Text = Trim$(Spell(EditorIndex).Name)
        .cmbClassReq.ListIndex = Spell(EditorIndex).ClassReq
        .scrlLevelReq.Value = Spell(EditorIndex).LevelReq

        .cmbType.ListIndex = Spell(EditorIndex).Type
        
        If Spell(EditorIndex).Type = SPELL_TYPE_STATCHANGE Then
            .fraChooseStat.Visible = True
        Else
            .fraChooseStat.Visible = False
        End If
        
        .scrlVitalMod.Value = Spell(EditorIndex).Data1
        .scrlCost.Value = Spell(EditorIndex).MPCost
        
        SpellSoundFileName = Trim$(Spell(EditorIndex).Sound)
        
        If SpellSoundFileName = "." Or SpellSoundFileName = ".." Or SpellSoundFileName = "0" Then
            .lblSound.Caption = "No Sound"
        Else
            .lblSound.Caption = SpellSoundFileName
        End If
    
        If Spell(EditorIndex).Range = 0 Then
            Spell(EditorIndex).Range = 1
        End If
    
        .scrlRange.Value = Spell(EditorIndex).Range
        .scrlSpellAnim.Value = Spell(EditorIndex).SpellAnim
    
        If Spell(EditorIndex).SpellTime < 1 Then
            .scrlSpellTime.Value = 40
        Else
            .scrlSpellTime.Value = Spell(EditorIndex).SpellTime
        End If
    
        If Spell(EditorIndex).SpellDone < 1 Then
            .scrlSpellDone.Value = 1
        Else
            .scrlSpellDone.Value = Spell(EditorIndex).SpellDone
        End If

        .chkArea.Value = Spell(EditorIndex).AE
        
        If Spell(EditorIndex).SelfSpell = True Then
            .chkSelfSpell.Value = Checked
        Else
            .chkSelfSpell.Value = Unchecked
        End If
        
        .chkBig.Value = Spell(EditorIndex).Big

        .scrlElement.Value = Spell(EditorIndex).Element
        .scrlElement.Max = MAX_ELEMENTS
    
        .cmbPassiveStat.ListIndex = Spell(EditorIndex).PassiveStat
        .scrlPassiveStat.Value = Spell(EditorIndex).PassiveStatChange
    
        If Spell(EditorIndex).UsePassiveStat = True Then
            .chkPassive.Value = Checked
        Else
            .chkPassive.Value = Unchecked
        End If
    
        If Spell(EditorIndex).Type = SPELL_TYPE_STATCHANGE Then
            .cmbStat.ListIndex = Spell(EditorIndex).Stat
            .lblStat.Caption = .cmbStat.List(Spell(EditorIndex).Stat)
            .scrlTime.Value = Spell(EditorIndex).StatTime
            
            If .scrlTime.Value <> 1 Then
                .lblStatTime.Caption = Trim$(Spell(EditorIndex).StatTime & " Seconds")
            Else
                .lblStatTime.Caption = Trim$(Spell(EditorIndex).StatTime & " Second")
            End If
                
            If Len(CStr(Spell(EditorIndex).Multiplier)) > 1 Then
                .scrlMult1 = Int(Left(Spell(EditorIndex).Multiplier, 1))
                .scrlMult2 = Int(Right(Spell(EditorIndex).Multiplier, 1))
            Else
                .scrlMult1 = Int(Spell(EditorIndex).Multiplier)
                .scrlMult2 = Int(0)
            End If
            
            .lblMultiplier.Caption = .scrlMult1.Value & "." & .scrlMult2.Value
        End If
        
        .Show vbModal
    End With
End Sub

Public Sub SpellEditorOk()
    With frmSpellEditor
        Spell(EditorIndex).Name = .txtName.Text
        Spell(EditorIndex).ClassReq = .cmbClassReq.ListIndex
        Spell(EditorIndex).LevelReq = .scrlLevelReq.Value
        Spell(EditorIndex).Type = .cmbType.ListIndex
        Spell(EditorIndex).Data1 = .scrlVitalMod.Value
        Spell(EditorIndex).Data3 = 0
        Spell(EditorIndex).MPCost = .scrlCost.Value
        Spell(EditorIndex).Sound = SpellSoundFileName
        Spell(EditorIndex).Range = .scrlRange.Value
    
        Spell(EditorIndex).SpellAnim = .scrlSpellAnim.Value
        Spell(EditorIndex).SpellTime = .scrlSpellTime.Value
        Spell(EditorIndex).SpellDone = .scrlSpellDone.Value
    
        Spell(EditorIndex).AE = .chkArea.Value
        Spell(EditorIndex).Big = .chkBig.Value
        
        Spell(EditorIndex).Element = .scrlElement.Value
        Spell(EditorIndex).Stat = .cmbStat.ListIndex
        Spell(EditorIndex).StatTime = .scrlTime.Value
        Spell(EditorIndex).Multiplier = Val(.scrlMult1 & "." & .scrlMult2)
            
        If .chkPassive.Value = Checked Then
            Spell(EditorIndex).UsePassiveStat = True
            Spell(EditorIndex).PassiveStat = .cmbPassiveStat.ListIndex
            Spell(EditorIndex).PassiveStatChange = .scrlPassiveStat.Value
        Else
            Spell(EditorIndex).UsePassiveStat = False
            Spell(EditorIndex).PassiveStat = 0
            Spell(EditorIndex).PassiveStatChange = 0
        End If
        
        If .chkSelfSpell.Value = Checked Then
            Spell(EditorIndex).SelfSpell = True
        Else
            Spell(EditorIndex).SelfSpell = False
        End If
    End With
        
    Call SendSaveSpell(EditorIndex)
    InSpellEditor = False
    Unload frmSpellEditor
End Sub

Public Sub SpellEditorCancel()
    InSpellEditor = False
    Unload frmSpellEditor
End Sub

Public Sub RecipeEditorInit()
    Dim i As Long

    frmRecipeEditor.lstIngredient1.Clear
    frmRecipeEditor.lstIngredient2.Clear
    frmRecipeEditor.lstResultItem.Clear
    
    frmRecipeEditor.lstIngredient1.addItem "None"
    frmRecipeEditor.lstIngredient2.addItem "None"
    frmRecipeEditor.lstResultItem.addItem "None"
    For i = 1 To MAX_ITEMS
        frmRecipeEditor.lstIngredient1.addItem i & ": " & Trim$(Item(i).Name)
        frmRecipeEditor.lstIngredient2.addItem i & ": " & Trim$(Item(i).Name)
        frmRecipeEditor.lstResultItem.addItem i & ": " & Trim$(Item(i).Name)
    Next i
    
    frmRecipeEditor.lstIngredient1.ListIndex = Recipe(EditorIndex).Ingredient1
    frmRecipeEditor.lstIngredient2.ListIndex = Recipe(EditorIndex).Ingredient2
    frmRecipeEditor.lstResultItem.ListIndex = Recipe(EditorIndex).ResultItem
    
    frmRecipeEditor.Show vbModal
End Sub

Public Sub RecipeEditorOk()
    Recipe(EditorIndex).Ingredient1 = frmRecipeEditor.lstIngredient1.ListIndex
    Recipe(EditorIndex).Ingredient2 = frmRecipeEditor.lstIngredient2.ListIndex
    Recipe(EditorIndex).ResultItem = frmRecipeEditor.lstResultItem.ListIndex
    If Recipe(EditorIndex).ResultItem > 0 Then
        Recipe(EditorIndex).Name = Trim$(Item(frmRecipeEditor.lstResultItem.ListIndex).Name)
    Else
        Recipe(EditorIndex).Name = vbNullString
    End If
    
    Call SendSaveRecipe(EditorIndex)
    InRecipeEditor = False
    Unload frmRecipeEditor
End Sub

Public Sub RecipeEditorCancel()
    InRecipeEditor = False
    Unload frmRecipeEditor
End Sub
