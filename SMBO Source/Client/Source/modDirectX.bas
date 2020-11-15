Attribute VB_Name = "modDirectX"
Option Explicit

Public DX As DirectX7
Public DD As DirectDraw7

Public DD_Clip As DirectDrawClipper

Public DD_PrimarySurf As DirectDrawSurface7
Public DDSD_Primary As DDSURFACEDESC2

Public DD_SpriteSurf As DirectDrawSurface7
Public DDSD_Sprite As DDSURFACEDESC2

Public DD_ItemSurf As DirectDrawSurface7
Public DDSD_Item As DDSURFACEDESC2

Public DD_EmoticonSurf As DirectDrawSurface7
Public DDSD_Emoticon As DDSURFACEDESC2

Public DD_BackBuffer As DirectDrawSurface7
Public DDSD_BackBuffer As DDSURFACEDESC2

Public DD_BigSpriteSurf As DirectDrawSurface7
Public DDSD_BigSprite As DDSURFACEDESC2

Public DD_SpellAnim As DirectDrawSurface7
Public DDSD_SpellAnim As DDSURFACEDESC2

Public DD_BigSpellAnim As DirectDrawSurface7
Public DDSD_BigSpellAnim As DDSURFACEDESC2

Public DD_TileSurf(0 To ExtraSheets) As DirectDrawSurface7
Public DDSD_Tile(0 To ExtraSheets) As DDSURFACEDESC2
Public TileFile(0 To ExtraSheets) As Byte

Public DDSD_ArrowAnim As DDSURFACEDESC2
Public DD_ArrowAnim As DirectDrawSurface7

Public DD_PvPImageSurf As DirectDrawSurface7
Public DDSD_PvPImage As DDSURFACEDESC2

Public DD_HPImageSurf As DirectDrawSurface7
Public DDSD_HPImage As DDSURFACEDESC2

Public DD_FPImageSurf As DirectDrawSurface7
Public DDSD_FPImage As DDSURFACEDESC2

Public DD_ExpImageSurf As DirectDrawSurface7
Public DDSD_ExpImage As DDSURFACEDESC2

Public DD_FadedAtkBattleImageSurf As DirectDrawSurface7
Public DDSD_FadedAtkBattleImage As DDSURFACEDESC2

Public DD_AtkBattleImageSurf As DirectDrawSurface7
Public DDSD_AtkBattleImage As DDSURFACEDESC2

Public DD_FadedRunBattleImageSurf As DirectDrawSurface7
Public DDSD_FadedRunBattleImage As DDSURFACEDESC2

Public DD_RunBattleImageSurf As DirectDrawSurface7
Public DDSD_RunBattleImage As DDSURFACEDESC2

Public DD_FadedItemBattleImageSurf As DirectDrawSurface7
Public DDSD_FadedItemBattleImage As DDSURFACEDESC2

Public DD_ItemBattleImageSurf As DirectDrawSurface7
Public DDSD_ItemBattleImage As DDSURFACEDESC2

Public DD_FadedSpecialBattleImageSurf As DirectDrawSurface7
Public DDSD_FadedSpecialBattleImage As DDSURFACEDESC2

Public DD_SpecialBattleImageSurf As DirectDrawSurface7
Public DDSD_SpecialBattleImage As DDSURFACEDESC2

Public DD_VictoryImageSurf As DirectDrawSurface7
Public DDSD_VictoryImage As DDSURFACEDESC2

Public DD_VictoryAnimSurf As DirectDrawSurface7
Public DDSD_VictoryAnim As DDSURFACEDESC2

Public rec As RECT
Public rec_pos As RECT

Sub InitDirectX()
    On Error GoTo DXErr
    
    ' Initialize DirextX
    Set DX = New DirectX7

    ' Initialize DirectDraw
    Set DD = DX.DirectDrawCreate(vbNullString)

    ' Indicate windows mode application
    Call DD.SetCooperativeLevel(frmMirage.hWnd, DDSCL_NORMAL)
    
    ' Init type and get the primary surface
    DDSD_Primary.lFlags = DDSD_CAPS
    DDSD_Primary.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_SYSTEMMEMORY
    Set DD_PrimarySurf = DD.CreateSurface(DDSD_Primary)

    ' Create the clipper
    Set DD_Clip = DD.CreateClipper(0)

    ' Associate the picture hwnd with the clipper
    DD_Clip.SetHWnd frmMirage.picScreen.hWnd

    ' Have the blits to the screen clipped to the picture box
    DD_PrimarySurf.SetClipper DD_Clip

    ' Initialize all surfaces
    Call InitSurfaces
    Exit Sub

    ' Error handling
DXErr:
    Call MsgBox("Error initializing DirectDraw! Make sure you have DirectX 7 or higher installed and a compatible graphics device. Err: " & Err.Number & ", Desc: " & Err.Description, vbCritical)
    Call GameDestroy
    End
End Sub

Sub InitSurfaces()
    Dim Key As DDCOLORKEY
    Dim i As Long
    Dim DC As Long

    ' Check for files existing
    If Not FileExists("\GFX\Sprites.smbo") Or Not FileExists("\GFX\Items.smbo") Or Not FileExists("\GFX\BigSprites.smbo") Or Not FileExists("\GFX\Emoticons.smbo") Or Not FileExists("\GFX\Arrows.smbo") Then
        Call MsgBox("You're missing some GFX files!", vbOKOnly, "Super Mario Bros. Online")
        Call GameDestroy
    End If
    
    ' Set the key for masks
    Key.Low = 0
    Key.High = 0

    ' Initialize back buffer
    DDSD_BackBuffer.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    DDSD_BackBuffer.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    DDSD_BackBuffer.lWidth = (MAX_MAPX + 1) * PIC_X
    DDSD_BackBuffer.lHeight = (MAX_MAPY + 1) * PIC_Y
    Set DD_BackBuffer = DD.CreateSurface(DDSD_BackBuffer)

    ' Init tiles
    For i = 0 To ExtraSheets
        If Dir$(App.Path & "\GFX\Tiles" & i & ".smbo") <> vbNullString Then
            Call InitDDSurf(App.Path & "\GFX\Tiles" & i, DD_TileSurf(i), DDSD_Tile(i))
            TileFile(i) = 1
        Else
            TileFile(i) = 0
        End If
    Next i
    
    ' Init items
    Call InitDDSurf(App.Path & "\GFX\Items", DD_ItemSurf, DDSD_Item)
    ' Init sprites
    Call InitDDSurf(App.Path & "\GFX\Sprites", DD_SpriteSurf, DDSD_Sprite)
    ' Init big sprites
    Call InitDDSurf(App.Path & "\GFX\BigSprites", DD_BigSpriteSurf, DDSD_BigSprite)
    ' Init emoticons
    Call InitDDSurf(App.Path & "\GFX\Emoticons", DD_EmoticonSurf, DDSD_Emoticon)
    ' Init spells
    Call InitDDSurf(App.Path & "\GFX\Spells", DD_SpellAnim, DDSD_SpellAnim)
    ' Init big spells
    Call InitDDSurf(App.Path & "\GFX\BigSpells", DD_BigSpellAnim, DDSD_BigSpellAnim)
    ' Init arrows
    Call InitDDSurf(App.Path & "\GFX\Arrows", DD_ArrowAnim, DDSD_ArrowAnim)
    ' Init victory animation
    Call InitDDSurf(App.Path & "\GFX\VictoryAnim", DD_VictoryAnimSurf, DDSD_VictoryAnim)
    ' Init pvp sign
    Call InitDDSurf(App.Path & "\GUI\pvpsign", DD_PvPImageSurf, DDSD_PvPImage)
    ' Init hp icon
    Call InitDDSurf(App.Path & "\GUI\hpicon", DD_HPImageSurf, DDSD_HPImage)
    ' Init fp icon
    Call InitDDSurf(App.Path & "\GUI\fpicon", DD_FPImageSurf, DDSD_FPImage)
    ' Init exp icon
    Call InitDDSurf(App.Path & "\GUI\expimage", DD_ExpImageSurf, DDSD_ExpImage)
    ' Init faded attack battle icon
    Call InitDDSurf(App.Path & "\GUI\fadedattackbattleicon", DD_FadedAtkBattleImageSurf, DDSD_FadedAtkBattleImage)
    ' Init attack battle icon
    Call InitDDSurf(App.Path & "\GUI\attackbattleicon", DD_AtkBattleImageSurf, DDSD_AtkBattleImage)
    ' Init faded run battle icon
    Call InitDDSurf(App.Path & "\GUI\fadedrunbattleicon", DD_FadedRunBattleImageSurf, DDSD_FadedRunBattleImage)
    ' Init run battle icon
    Call InitDDSurf(App.Path & "\GUI\runbattleicon", DD_RunBattleImageSurf, DDSD_RunBattleImage)
    ' Init faded item battle icon
    Call InitDDSurf(App.Path & "\GUI\fadeditembattleicon", DD_FadedItemBattleImageSurf, DDSD_FadedItemBattleImage)
    ' Init run battle icon
    Call InitDDSurf(App.Path & "\GUI\itembattleicon", DD_ItemBattleImageSurf, DDSD_ItemBattleImage)
    ' Init faded special battle icon
    Call InitDDSurf(App.Path & "\GUI\fadedspecialattackbattleicon", DD_FadedSpecialBattleImageSurf, DDSD_FadedSpecialBattleImage)
    ' Init special battle icon
    Call InitDDSurf(App.Path & "\GUI\specialattackbattleicon", DD_SpecialBattleImageSurf, DDSD_SpecialBattleImage)
    ' Init victory screen
    Call InitDDSurf(App.Path & "\GUI\Victory", DD_VictoryImageSurf, DDSD_VictoryImage)
End Sub

Public Sub InitDDSurf(FileName As String, ByRef Surf As DirectDrawSurface7, ByRef SurfDesc As DDSURFACEDESC2)
    Dim DC As Long
    Dim Decrypt As BitMapUtils
    
    Set Decrypt = New BitMapUtils
    
    Call Decrypt.LoadByteData(FileName & ".smbo")
    Call Decrypt.DecryptByteData("5006")
    Call Decrypt.DecompressByteData
    
    ' Init ddsd type
    SurfDesc.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    SurfDesc.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    SurfDesc.lWidth = Decrypt.ImageWidth
    SurfDesc.lHeight = Decrypt.ImageHeight
    
    ' set surface
    Set Surf = DD.CreateSurface(SurfDesc)
    
    DC = Surf.GetDC
    
    Call Decrypt.Blt(DC)
    Call Surf.ReleaseDC(DC)
    
    SetMaskColorFromPixel Surf, 0, 0
End Sub

Public Sub InitLoadPicture(FileName As String, ByRef PictureBox As PictureBox)
    Dim Decrypt As BitMapUtils
    Set Decrypt = New BitMapUtils
    
    Call Decrypt.LoadByteData(FileName)
    Call Decrypt.DecryptByteData("5006")
    Call Decrypt.DecompressByteData
    
    PictureBox.Cls
    
    PictureBox.Width = Decrypt.ImageWidth
    PictureBox.Height = Decrypt.ImageHeight
        
    Call Decrypt.Blt(PictureBox.hDC)
End Sub

Sub DestroyDirectX()
    Dim i As Long

    Set DX = Nothing
    Set DD = Nothing

    Set DD_Clip = Nothing

    Set DD_PrimarySurf = Nothing
    Set DD_BackBuffer = Nothing

    Set DD_SpriteSurf = Nothing

    For i = 0 To ExtraSheets
        If TileFile(i) = 1 Then
            Set DD_TileSurf(i) = Nothing
        End If
    Next i

    Set DD_ItemSurf = Nothing
    Set DD_BigSpriteSurf = Nothing
    Set DD_EmoticonSurf = Nothing
    Set DD_SpellAnim = Nothing
    Set DD_BigSpellAnim = Nothing
    Set DD_ArrowAnim = Nothing

    Set DD_PvPImageSurf = Nothing
    Set DD_HPImageSurf = Nothing
    Set DD_FPImageSurf = Nothing
    Set DD_ExpImageSurf = Nothing
    Set DD_FadedAtkBattleImageSurf = Nothing
    Set DD_AtkBattleImageSurf = Nothing
    Set DD_FadedRunBattleImageSurf = Nothing
    Set DD_RunBattleImageSurf = Nothing
    Set DD_FadedItemBattleImageSurf = Nothing
    Set DD_ItemBattleImageSurf = Nothing
    Set DD_FadedSpecialBattleImageSurf = Nothing
    Set DD_SpecialBattleImageSurf = Nothing
End Sub

Function NeedToRestoreSurfaces() As Boolean
    Dim TestCoopRes As Long

    TestCoopRes = DD.TestCooperativeLevel

    If (TestCoopRes = DD_OK) Then
        NeedToRestoreSurfaces = False
    Else
        NeedToRestoreSurfaces = True
    End If
End Function

Public Sub SetMaskColorFromPixel(ByRef TheSurface As DirectDrawSurface7, ByVal x As Long, ByVal y As Long)
    Dim TmpR As RECT
    Dim TmpDDSD As DDSURFACEDESC2
    Dim TmpColorKey As DDCOLORKEY

    With TmpR
        .Left = x
        .Top = y
        .Right = x
        .Bottom = y
    End With

    TheSurface.Lock TmpR, TmpDDSD, DDLOCK_WAIT Or DDLOCK_READONLY, 0

    With TmpColorKey
        .Low = TheSurface.GetLockedPixel(x, y)
        .High = .Low
    End With

    TheSurface.SetColorKey DDCKEY_SRCBLT, TmpColorKey

    TheSurface.Unlock TmpR
End Sub

Sub DisplayFx(ByRef surfDisplay As DirectDrawSurface7, intX As Long, intY As Long, intWidth As Long, intHeight As Long, lngROP As Long, blnFxCap As Boolean, Tile As Long)
    Dim lngSrcDC As Long
    Dim lngDestDC As Long

    lngDestDC = DD_BackBuffer.GetDC
    lngSrcDC = surfDisplay.GetDC
    BitBlt lngDestDC, intX, intY, intWidth, intHeight, lngSrcDC, (Tile - (Tile \ TilesInSheets) * TilesInSheets) * PIC_X, (Tile \ TilesInSheets) * PIC_Y, lngROP
    surfDisplay.ReleaseDC lngSrcDC
    DD_BackBuffer.ReleaseDC lngDestDC
End Sub

Public Function GetScreenLeft(ByVal Index As Long) As Long
    GetScreenLeft = GetPlayerX(Index) - 11
End Function

Public Function GetScreenTop(ByVal Index As Long) As Long
    GetScreenTop = GetPlayerY(Index) - 8
End Function

Public Function GetScreenRight(ByVal Index As Long) As Long
    GetScreenRight = GetPlayerX(Index) + 10
End Function

Public Function GetScreenBottom(ByVal Index As Long) As Long
    GetScreenBottom = GetPlayerY(Index) + 8
End Function

Sub BltTile2(ByVal x As Long, ByVal y As Long, ByVal Tile As Long)
    If TileFile(10) = 0 Then
        Exit Sub
    End If

    rec.Top = (Tile \ TilesInSheets) * PIC_Y
    rec.Bottom = rec.Top + PIC_Y
    rec.Left = (Tile - (Tile \ TilesInSheets) * TilesInSheets) * PIC_X
    rec.Right = rec.Left + PIC_X
    
    Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) + sx - NewXOffset, y - (NewPlayerY * PIC_Y) + sx - NewYOffset, DD_TileSurf(10), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    ' DisplayFx DD_TileSurf(10), (x - NewPlayerX * PIC_X) + sx - NewXOffset, y - (NewPlayerY * PIC_Y) + sx - NewYOffset, 32, 16, vbSrcAnd, DDBLT_ROP Or DDBLT_WAIT, Tile
End Sub

Sub BltPlayerText(ByVal Index As Long)
    Dim TextX As Long, TextY As Long, Color As Long
    Dim intLoop As Long
    Dim bytLineCount As Byte, bytLineLength As Byte
    Dim strLine(0 To MAX_LINES - 1) As String, strWords() As String
    
    If IsPlayingHideNSneak = True Then
        Exit Sub
    End If
    
    If CanBltPlayerGfx(Index) = False Then
        Exit Sub
    End If
    
    If Player(MyIndex).BattleVictory = True Then
        Exit Sub
    End If
    
    strWords() = Split(Bubble(Index).Text, " ")

    TextX = GetPlayerX(Index) * PIC_X + Player(Index).xOffset + Int(PIC_X) - 6
    TextY = GetPlayerY(Index) * PIC_Y + Player(Index).yOffset - Int(PIC_Y) - 25

    ' Loop through all the words.
    For intLoop = 0 To UBound(strWords)
        ' Increment the line length.
        bytLineLength = bytLineLength + Len(strWords(intLoop)) + 1

        ' If we have room on the current line.
        If bytLineLength < MAX_LINE_LENGTH Then
            ' Add the text to the current line.
            strLine(bytLineCount) = strLine(bytLineCount) & strWords(intLoop) & " "
        Else
            bytLineCount = bytLineCount + 1

            If bytLineCount = MAX_LINES Then
                bytLineCount = bytLineCount - 1
                Exit For
            End If

            strLine(bytLineCount) = Trim$(strWords(intLoop)) & " "
            bytLineLength = 0
        End If
    Next intLoop
    
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
    
    For intLoop = 0 To (MAX_LINES - 1)
        If strLine(intLoop) <> vbNullString Then
            Call DrawText(TexthDC, TextX - (NewPlayerX * PIC_X) + sx - NewXOffset - ((Len(strLine(intLoop)) * 8) \ 2) - 4, TextY - (NewPlayerY * PIC_Y) + sx - NewYOffset, strLine(intLoop), Color, GameFont)
            
            TextY = TextY + 16
        End If
    Next intLoop
End Sub

Sub BltPlayerBars(ByVal Index As Long)
    Dim x As Long, y As Long, y2 As Long
    
    If IsPlayingHideNSneak = True Then
        Exit Sub
    End If
    
    If CanBltPlayerGfx(Index) = False Then
        Exit Sub
    End If
    
    If Index <> MyIndex And GetPlayerGuild(Index) = vbNullString Then
        Exit Sub
    End If
    
    If Player(Index).HP = Player(Index).MaxHp Or Player(Index).HP = 0 Then
        Exit Sub
    End If
    
    x = (GetPlayerX(Index) * PIC_X + sx + Player(Index).xOffset) - (NewPlayerX * PIC_X) - NewXOffset
    y = (GetPlayerY(Index) * PIC_Y + sx + Player(Index).yOffset) - (NewPlayerY * PIC_Y) - NewYOffset
    y = y + 30
    y2 = y - 4
    
    ' draws the back bars
    Call DD_BackBuffer.SetFillColor(RGB(255, 0, 0))
    Call DD_BackBuffer.DrawBox(x, y, x + 32, y2)

    ' draws HP
    Call DD_BackBuffer.SetFillColor(RGB(0, 255, 0))
    Call DD_BackBuffer.DrawBox(x, y, x + ((Player(Index).HP / 100) / (Player(Index).MaxHp / 100) * 32), y2)
End Sub

Sub BltPlayerSPBars(ByVal Index As Long)
    Dim x As Long, y As Long
    
    If IsPlayingHideNSneak = True Then
        Exit Sub
    End If
    
    If CanBltPlayerGfx(Index) = False Then
        Exit Sub
    End If

    x = (GetPlayerX(Index) * PIC_X + sx + Player(Index).xOffset) - (NewPlayerX * PIC_X) - NewXOffset
    y = (GetPlayerY(Index) * PIC_Y + sx + Player(Index).yOffset) - (NewPlayerY * PIC_Y) - NewYOffset
    
    If Player(Index).SP = Player(Index).MaxSP Then
        Exit Sub
    End If
    
    ' draws the SP back bars
    Call DD_BackBuffer.SetFillColor(RGB(0, 0, 0))
    Call DD_BackBuffer.DrawBox(x, y + 35, x + 32, y + 31)
        
    ' draws SP
    Call DD_BackBuffer.SetFillColor(RGB(0, 85, 200))
    Call DD_BackBuffer.DrawBox(x, y + 35, x + ((Player(Index).SP / 100) / (Player(Index).MaxSP / 100) * 32), y + 31)
End Sub

Sub BltNpcBars(ByVal Index As Long)
    Dim x As Long, y As Long, y2 As Long

    On Error GoTo BltNpcBars_Error
    
    If CanBltNpcGfx(Index) = False Then
        Exit Sub
    End If
    
    If MapNpc(Index).HP = 0 Then
        Exit Sub
    End If
    If MapNpc(Index).num < 1 Then
        Exit Sub
    End If

    If Npc(MapNpc(Index).num).Big = 1 Then
        x = (MapNpc(Index).x * PIC_X + sx - 9 + MapNpc(Index).xOffset) - (NewPlayerX * PIC_X) - NewXOffset
        y = (MapNpc(Index).y * PIC_Y + sx + MapNpc(Index).yOffset) - (NewPlayerY * PIC_Y) - NewYOffset
        y = y + 32
        y2 = y + 4

        Call DD_BackBuffer.SetFillColor(RGB(255, 0, 0))
        Call DD_BackBuffer.DrawBox(x, y, x + 50, y2)
        Call DD_BackBuffer.SetFillColor(RGB(0, 255, 0))
        
        If MapNpc(Index).MaxHp < 1 Then
            Call DD_BackBuffer.DrawBox(x, y, x + ((MapNpc(Index).HP / 100) / ((MapNpc(Index).MaxHp + 1) / 100) * 50), y2)
        Else
            Call DD_BackBuffer.DrawBox(x, y, x + ((MapNpc(Index).HP / 100) / (MapNpc(Index).MaxHp / 100) * 50), y2)
        End If
    Else
        x = (MapNpc(Index).x * PIC_X + sx + MapNpc(Index).xOffset) - (NewPlayerX * PIC_X) - NewXOffset
        y = (MapNpc(Index).y * PIC_Y + sx + MapNpc(Index).yOffset) - (NewPlayerY * PIC_Y) - NewYOffset
        y = y + 32
        y2 = y + 4
        
        Call DD_BackBuffer.SetFillColor(RGB(255, 0, 0))
        Call DD_BackBuffer.DrawBox(x, y, x + 32, y2)
        Call DD_BackBuffer.SetFillColor(RGB(0, 255, 0))

        If MapNpc(Index).MaxHp < 1 Then
            Call DD_BackBuffer.DrawBox(x, y, x + ((MapNpc(Index).HP / 100) / ((MapNpc(Index).MaxHp + 1) / 100) * 32), y2)
        Else
            Call DD_BackBuffer.DrawBox(x, y, x + ((MapNpc(Index).HP / 100) / (MapNpc(Index).MaxHp / 100) * 32), y2)
        End If

    End If


    On Error GoTo 0
    Exit Sub

BltNpcBars_Error:

    If Err.Number = DDERR_CANTCREATEDC Then

    End If

End Sub

Sub BltWeather()
    Dim i As Long

    Call DD_BackBuffer.SetForeColor(RGB(0, 0, 200))

    If GameWeather = WEATHER_RAINING Or GameWeather = WEATHER_THUNDER Then
        For i = 1 To MAX_RAINDROPS
            If DropRain(i).Randomized = False Then
                If frmMirage.tmrRainDrop.Enabled = False Then
                    BLT_RAIN_DROPS = 1
                    frmMirage.tmrRainDrop.Enabled = True
                    If frmMirage.tmrRainDrop.Tag = vbNullString Then
                        frmMirage.tmrRainDrop.Interval = 100
                        frmMirage.tmrRainDrop.Tag = "123"
                    End If
                End If
            End If
        Next i
    ElseIf GameWeather = WEATHER_SNOWING Then
        For i = 1 To MAX_RAINDROPS
            If DropSnow(i).Randomized = False Then
                If frmMirage.tmrSnowDrop.Enabled = False Then
                    BLT_SNOW_DROPS = 1
                    frmMirage.tmrSnowDrop.Enabled = True
                    If frmMirage.tmrSnowDrop.Tag = vbNullString Then
                        frmMirage.tmrSnowDrop.Interval = 200
                        frmMirage.tmrSnowDrop.Tag = "123"
                    End If
                End If
            End If
        Next i
    Else
        If BLT_RAIN_DROPS > 0 And BLT_RAIN_DROPS <= RainIntensity Then
            Call ClearRainDrop(BLT_RAIN_DROPS)
        End If
        frmMirage.tmrRainDrop.Tag = vbNullString
    End If

    For i = 1 To MAX_RAINDROPS
        If Not ((DropRain(i).x = 0) Or (DropRain(i).y = 0)) Then
            rec.Top = 0
            rec.Bottom = rec.Top + PIC_Y
            rec.Left = 6 * PIC_X
            rec.Right = rec.Left + PIC_X
            DropRain(i).x = DropRain(i).x + DropRain(i).speed
            DropRain(i).y = DropRain(i).y + DropRain(i).speed
            Call DD_BackBuffer.BltFast(DropRain(i).x + DropRain(i).speed, DropRain(i).y + DropRain(i).speed, DD_TileSurf(10), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            If (DropRain(i).x > (MAX_MAPX + 1) * PIC_X) Or (DropRain(i).y > (MAX_MAPY + 1) * PIC_Y) Then
                DropRain(i).Randomized = False
            End If
        End If
    Next i
    If TileFile(10) = 1 Then
        rec.Top = (14 \ TilesInSheets) * PIC_Y
        rec.Bottom = rec.Top + PIC_Y
        rec.Left = (14 - (14 \ TilesInSheets) * TilesInSheets) * PIC_X
        rec.Right = rec.Left + PIC_X
        For i = 1 To MAX_RAINDROPS
            If Not ((DropSnow(i).x = 0) Or (DropSnow(i).y = 0)) Then
                DropSnow(i).x = DropSnow(i).x + DropSnow(i).speed
                DropSnow(i).y = DropSnow(i).y + DropSnow(i).speed
                Call DD_BackBuffer.BltFast(DropSnow(i).x + DropSnow(i).speed, DropSnow(i).y + DropSnow(i).speed, DD_TileSurf(10), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                If (DropSnow(i).x > (MAX_MAPX + 1) * PIC_X) Or (DropSnow(i).y > (MAX_MAPY + 1) * PIC_Y) Then
                    DropSnow(i).Randomized = False
                End If
            End If
        Next i
    End If

    ' If it's thunder, make the screen randomly flash white
    If GameWeather = WEATHER_THUNDER Then
        If Int((100 - 1 + 1) * Rnd2) + 1 = 8 Then
            DD_BackBuffer.SetFillColor RGB(255, 255, 255)

            Call DD_BackBuffer.DrawBox(0, 0, (MAX_MAPX + 1) * PIC_X, (MAX_MAPY + 1) * PIC_Y)
        End If
    End If
End Sub

Sub BltMapWeather()
    Dim i As Long

    Call DD_BackBuffer.SetForeColor(RGB(0, 0, 200))

    If Map(GetPlayerMap(MyIndex)).Weather = 1 Or Map(GetPlayerMap(MyIndex)).Weather = 3 Then
        For i = 1 To MAX_RAINDROPS
            If DropRain(i).Randomized = False Then
                If frmMirage.tmrRainDrop.Enabled = False Then
                    BLT_RAIN_DROPS = 1
                    frmMirage.tmrRainDrop.Enabled = True
                End If
            End If
        Next i
        For i = 1 To MAX_RAINDROPS
            If Not ((DropRain(i).x = 0) Or (DropRain(i).y = 0)) Then
                rec.Top = (14 - (14 \ TilesInSheets)) * PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                rec.Left = 6 * PIC_X
                rec.Right = rec.Left + PIC_X
                DropRain(i).x = DropRain(i).x + DropRain(i).speed
                DropRain(i).y = DropRain(i).y + DropRain(i).speed
                Call DD_BackBuffer.BltFast(DropRain(i).x + DropRain(i).speed, DropRain(i).y + DropRain(i).speed, DD_TileSurf(10), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                If (DropRain(i).x > (MAX_MAPX + 1) * PIC_X) Or (DropRain(i).y > (MAX_MAPY + 1) * PIC_Y) Then
                    DropRain(i).Randomized = False
                End If
            End If
        Next i

        If Map(GetPlayerMap(MyIndex)).Weather = 3 Then
            If Int((100 - 1 + 1) * Rnd2) + 1 < 3 Then
                DD_BackBuffer.SetFillColor RGB(255, 255, 255)

                Call DD_BackBuffer.DrawBox(0, 0, (MAX_MAPX + 1) * PIC_X, (MAX_MAPY + 1) * PIC_Y)
            End If
        End If

    ElseIf Map(GetPlayerMap(MyIndex)).Weather = 2 Then
        For i = 1 To MAX_RAINDROPS
            If DropSnow(i).Randomized = False Then
                If frmMirage.tmrSnowDrop.Enabled = False Then
                    BLT_SNOW_DROPS = 1
                    frmMirage.tmrSnowDrop.Enabled = True
                End If
            End If
        Next i
        If TileFile(10) = 1 Then
            rec.Top = (14 \ TilesInSheets) * PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            rec.Left = (14 - (14 \ TilesInSheets) * TilesInSheets) * PIC_X
            rec.Right = rec.Left + PIC_X

            For i = 1 To MAX_RAINDROPS
                If Not ((DropSnow(i).x = 0) Or (DropSnow(i).y = 0)) Then
                    DropSnow(i).x = DropSnow(i).x + DropSnow(i).speed
                    DropSnow(i).y = DropSnow(i).y + DropSnow(i).speed
                    Call DD_BackBuffer.BltFast(DropSnow(i).x + DropSnow(i).speed, DropSnow(i).y + DropSnow(i).speed, DD_TileSurf(10), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    If (DropSnow(i).x > (MAX_MAPX + 1) * PIC_X) Or (DropSnow(i).y > (MAX_MAPY + 1) * PIC_Y) Then
                        DropSnow(i).Randomized = False
                    End If
                End If
            Next i
        End If
    Else
        If BLT_RAIN_DROPS > 0 And BLT_RAIN_DROPS <= frmMirage.tmrRainDrop.Interval Then
            Call ClearRainDrop(BLT_RAIN_DROPS)
        End If
        frmMirage.tmrRainDrop.Tag = vbNullString
    End If
End Sub

Sub RNDRainDrop(ByVal RDNumber As Long)
Start:
    DropRain(RDNumber).x = Int((((MAX_MAPX + 1) * PIC_X) * Rnd2) + 1)
    DropRain(RDNumber).y = Int((((MAX_MAPY + 1) * PIC_Y) * Rnd2) + 1)
    If (DropRain(RDNumber).y > (MAX_MAPY + 1) * PIC_Y / 4) And (DropRain(RDNumber).x > (MAX_MAPX + 1) * PIC_X / 4) Then
        GoTo Start
    End If
    DropRain(RDNumber).speed = Int((10 * Rnd2) + 6)
    DropRain(RDNumber).Randomized = True
End Sub

Sub ClearRainDrop(ByVal RDNumber As Long)
    On Error Resume Next
    DropRain(RDNumber).x = 0
    DropRain(RDNumber).y = 0
    DropRain(RDNumber).speed = 0
    DropRain(RDNumber).Randomized = False
End Sub

Sub RNDSnowDrop(ByVal RDNumber As Long)
Start:
    DropSnow(RDNumber).x = Int((((MAX_MAPX + 1) * PIC_X) * Rnd2) + 1)
    DropSnow(RDNumber).y = Int((((MAX_MAPY + 1) * PIC_Y) * Rnd2) + 1)
    If (DropSnow(RDNumber).y > (MAX_MAPY + 1) * PIC_Y / 4) And (DropSnow(RDNumber).x > (MAX_MAPX + 1) * PIC_X / 4) Then
        GoTo Start
    End If
    DropSnow(RDNumber).speed = Int((10 * Rnd2) + 6)
    DropSnow(RDNumber).Randomized = True
End Sub

Sub ClearSnowDrop(ByVal RDNumber As Long)
    On Error Resume Next
    DropSnow(RDNumber).x = 0
    DropSnow(RDNumber).y = 0
    DropSnow(RDNumber).speed = 0
    DropSnow(RDNumber).Randomized = False
End Sub

Sub BltSpell(ByVal Index As Long)
    Dim x As Long, y As Long, i As Long

    If Player(Index).SpellNum <= 0 Or Player(Index).SpellNum > MAX_SPELLS Then
        Exit Sub
    End If


    For i = 1 To MAX_SPELL_ANIM
        ' IF SPELL IS NOT BIG
        If Spell(Player(Index).SpellNum).Big = 0 Then
            If Player(Index).SpellAnim(i).CastedSpell = Yes Then
                If Player(Index).SpellAnim(i).SpellDone < Spell(Player(Index).SpellNum).SpellDone Then

                    rec.Top = Spell(Player(Index).SpellNum).SpellAnim * PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    rec.Left = Player(Index).SpellAnim(i).SpellVar * PIC_X
                    rec.Right = rec.Left + PIC_X

                    If Player(Index).SpellAnim(i).TargetType = 0 Then

                        ' SMALL: IF TARGET IS A PLAYER
                        If Player(Index).SpellAnim(i).Target > 0 Then

                            ' SMALL: IF TARGET IS SELF
                            If Player(Index).SpellAnim(i).Target = MyIndex Then
                                x = NewX + sx
                                y = NewY + sx
                                Call DD_BackBuffer.BltFast(x, y, DD_SpellAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)

                            ' SMALL: IF TARGET IS ANOTHER PLAYER
                            Else
                                x = GetPlayerX(Player(Index).SpellAnim(i).Target) * PIC_X + sx + Player(Player(Index).SpellAnim(i).Target).xOffset
                                y = GetPlayerY(Player(Index).SpellAnim(i).Target) * PIC_Y + sx + Player(Player(Index).SpellAnim(i).Target).yOffset
                                Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpellAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                            End If
                        End If

                    ' SMALL: IF TARGET IS AN NPC
                    Else
                        x = MapNpc(Player(Index).SpellAnim(i).Target).x * PIC_X + sx + MapNpc(Player(Index).SpellAnim(i).Target).xOffset
                        y = MapNpc(Player(Index).SpellAnim(i).Target).y * PIC_Y + sx + MapNpc(Player(Index).SpellAnim(i).Target).yOffset
                        Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpellAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If


' SMALL: ADVANCE SPELL ONE CYCLE

                    If GetTickCount > Player(Index).SpellAnim(i).SpellTime + Spell(Player(Index).SpellNum).SpellTime Then
                        Player(Index).SpellAnim(i).SpellTime = GetTickCount
                        Player(Index).SpellAnim(i).SpellVar = Player(Index).SpellAnim(i).SpellVar + 1
                    End If

                    If Player(Index).SpellAnim(i).SpellVar > 12 Then
                        Player(Index).SpellAnim(i).SpellDone = Player(Index).SpellAnim(i).SpellDone + 1
                        Player(Index).SpellAnim(i).SpellVar = 0
                    End If

                Else
                    Player(Index).SpellAnim(i).CastedSpell = No
                End If
            End If
        Else
            If Player(Index).SpellAnim(i).CastedSpell = Yes Then
                If Player(Index).SpellAnim(i).SpellDone < Spell(Player(Index).SpellNum).SpellDone Then

                    rec.Top = Spell(Player(Index).SpellNum).SpellAnim * (PIC_Y * 3)
                    rec.Bottom = rec.Top + PIC_Y + 64
                    rec.Left = Player(Index).SpellAnim(i).SpellVar * PIC_X
                    rec.Right = rec.Left + PIC_X + 64

                    If Player(Index).SpellAnim(i).TargetType = 0 Then

                        ' BIG: IF TARGET IS A PLAYER
                        If Player(Index).SpellAnim(i).Target > 0 Then

                            ' BIG: IF TARGET IS SELF
                            If Player(Index).SpellAnim(i).Target = MyIndex Then
                                x = NewX + sx - 32
                                y = NewY + sx - 32

                                If y < 0 Then
                                    rec.Top = rec.Top + (y * -1)
                                    y = 0
                                End If

                                If x < 0 Then
                                    rec.Left = rec.Left + (x * -1)
                                    x = 0
                                End If

                                If (x + 64) > (MAX_MAPX * 32) Then
                                    rec.Right = rec.Left + 64
                                End If

                                If (y + 64) > (MAX_MAPY * 32) Then
                                    rec.Bottom = rec.Top + 64
                                End If

                                Call DD_BackBuffer.BltFast(x, y, DD_BigSpellAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)

                            ' BIG: IF TARGET IS A DIFFERENT PLAYER
                            Else
                                x = GetPlayerX(Player(Index).SpellAnim(i).Target) * PIC_X + sx - 32 + Player(Player(Index).SpellAnim(i).Target).xOffset
                                y = GetPlayerY(Player(Index).SpellAnim(i).Target) * PIC_Y + sx - 32 + Player(Player(Index).SpellAnim(i).Target).yOffset

                                If y < 0 Then
                                    rec.Top = rec.Top + (y * -1)
                                    y = 0
                                End If

                                If x < 0 Then
                                    rec.Left = rec.Left + (x * -1)
                                    x = 0
                                End If

                                If (x + 64) > (MAX_MAPX * 32) Then
                                    rec.Right = rec.Left + 64
                                End If

                                If (y + 64) > (MAX_MAPY * 32) Then
                                    rec.Bottom = rec.Top + 64
                                End If

                                Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_BigSpellAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                            End If
                        End If

                    ' BIG: IF TARGET IS AN NPC
                    Else
                        x = MapNpc(Player(Index).SpellAnim(i).Target).x * PIC_X + sx - 32 + MapNpc(Player(Index).SpellAnim(i).Target).xOffset
                        y = MapNpc(Player(Index).SpellAnim(i).Target).y * PIC_Y + sx - 32 + MapNpc(Player(Index).SpellAnim(i).Target).yOffset

                        If y < 0 Then
                            rec.Top = rec.Top + (y * -1)
                            y = 0
                        End If

                        If x < 0 Then
                            rec.Left = rec.Left + (x * -1)
                            x = 0
                        End If

                        If (x + 64) > (MAX_MAPX * 32) Then
                            rec.Right = rec.Left + 64
                        End If

                        If (y + 64) > (MAX_MAPY * 32) Then
                            rec.Bottom = rec.Top + 64
                        End If

                        Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_BigSpellAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If

                    ' BIG: ADVANCE SPELL ONE CYCLE
                    If GetTickCount > Player(Index).SpellAnim(i).SpellTime + Spell(Player(Index).SpellNum).SpellTime Then
                        Player(Index).SpellAnim(i).SpellTime = GetTickCount
                        Player(Index).SpellAnim(i).SpellVar = Player(Index).SpellAnim(i).SpellVar + 3
                    End If

                    If Player(Index).SpellAnim(i).SpellVar > 36 Then
                        Player(Index).SpellAnim(i).SpellDone = Player(Index).SpellAnim(i).SpellDone + 1
                        Player(Index).SpellAnim(i).SpellVar = 0
                    End If

                Else
                    Player(Index).SpellAnim(i).CastedSpell = No
                End If
            End If
        End If
    Next i
End Sub

' Scripted Spell
Sub BltScriptSpell(ByVal i As Long)
    Dim rec As RECT
    Dim x As Long, y As Long

    x = ScriptSpell(i).x
    y = ScriptSpell(i).y

    If Spell(ScriptSpell(i).SpellNum).Big = 0 Then
        If ScriptSpell(i).SpellDone < Spell(ScriptSpell(i).SpellNum).SpellDone Then
            rec.Top = Spell(ScriptSpell(i).SpellNum).SpellAnim * PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            rec.Left = ScriptSpell(i).SpellVar * PIC_X
            rec.Right = rec.Left + PIC_X

            x = x * PIC_X + sx
            y = y * PIC_Y + sx

            If ScriptSpell(i).SpellVar > 10 Then
                ScriptSpell(i).SpellDone = ScriptSpell(i).SpellDone + 1
                ScriptSpell(i).SpellVar = 0
            End If

            If GetTickCount > ScriptSpell(i).SpellTime + Spell(ScriptSpell(i).SpellNum).SpellTime Then
                ScriptSpell(i).SpellTime = GetTickCount
                ScriptSpell(i).SpellVar = ScriptSpell(i).SpellVar + 1
            End If

            If ScriptSpell(i).SpellNum = 42 And ScriptSpell(i).Index = MyIndex Then
                If JugemsCloudHolder = "SMBO" Then
                    JugemsCloudHolder = Map(GetPlayerMap(MyIndex)).Tile(ScriptSpell(i).x, ScriptSpell(i).y).String1
                End If
                
                Map(GetPlayerMap(MyIndex)).Tile(ScriptSpell(i).x, ScriptSpell(i).y).String1 = "Jugem's Cloud"
            End If

            Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X), y - (NewPlayerY * PIC_Y), DD_SpellAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        Else ' spell is done
            If ScriptSpell(i).SpellNum = 42 And ScriptSpell(i).Index = MyIndex Then
                Map(GetPlayerMap(MyIndex)).Tile(ScriptSpell(i).x, ScriptSpell(i).y).String1 = JugemsCloudHolder
                
                JugemsCloudHolder = "SMBO"
            End If
        
            ScriptSpell(i).CastedSpell = No
        End If
    Else
        If ScriptSpell(i).SpellDone < Spell(ScriptSpell(i).SpellNum).SpellDone Then
            rec.Top = Spell(ScriptSpell(i).SpellNum).SpellAnim * (PIC_Y * 3)
            rec.Bottom = rec.Top + PIC_Y + 64
            rec.Left = ScriptSpell(i).SpellVar * PIC_X
            rec.Right = rec.Left + PIC_X + 64

            x = x * PIC_X + sx - 32
            y = y * PIC_Y + sx - 32

            If y < 0 Then
                rec.Top = rec.Top + (y * -1)
                y = 0
            End If

            If x < 0 Then
                rec.Left = rec.Left + (x * -1)
                x = 0
            End If

            If (x + 64) > (MAX_MAPX * 32) Then
                rec.Right = rec.Left + 64
            End If

            If (y + 64) > (MAX_MAPY * 32) Then
                rec.Bottom = rec.Top + 64
            End If

            If ScriptSpell(i).SpellVar > 30 Then
                ScriptSpell(i).SpellDone = ScriptSpell(i).SpellDone + 1
                ScriptSpell(i).SpellVar = 0
            End If

            If GetTickCount > ScriptSpell(i).SpellTime + Spell(ScriptSpell(i).SpellNum).SpellTime Then
                ScriptSpell(i).SpellTime = GetTickCount
                ScriptSpell(i).SpellVar = ScriptSpell(i).SpellVar + 3
            End If

            Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X), y - (NewPlayerY * PIC_Y), DD_BigSpellAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        Else 'spell is done
            ScriptSpell(i).CastedSpell = No
        End If
    End If
End Sub

Sub BltEmoticons(ByVal Index As Long)
    Dim x2 As Long, y2 As Long
    Dim ETime As Long
    ETime = 1300

    If Player(Index).EmoticonNum < 0 Then
        Exit Sub
    End If

    If Player(Index).EmoticonTime + ETime > GetTickCount Then
        If GetTickCount < Player(Index).EmoticonTime + ((ETime \ 3) * 1) Then
            Player(Index).EmoticonVar = 0
        ElseIf GetTickCount < Player(Index).EmoticonTime + ((ETime \ 3) * 2) Then
            Player(Index).EmoticonVar = 1
        ElseIf GetTickCount < Player(Index).EmoticonTime + ((ETime \ 3) * 3) Then
            Player(Index).EmoticonVar = 2
        End If

        rec.Top = Player(Index).EmoticonNum * PIC_Y
        rec.Bottom = rec.Top + PIC_Y
        rec.Left = Player(Index).EmoticonVar * PIC_X
        rec.Right = rec.Left + PIC_X

        If Index = MyIndex Then
            x2 = NewX + sx + 16
            y2 = NewY + sx - 32

            If y2 < 0 Then
                Exit Sub
            End If

            Call DD_BackBuffer.BltFast(x2, y2, DD_EmoticonSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        Else
            x2 = GetPlayerX(Index) * PIC_X + sx + Player(Index).xOffset + 16
            y2 = GetPlayerY(Index) * PIC_Y + sx + Player(Index).yOffset - 32

            If y2 < 0 Then
                Exit Sub
            End If

            Call DD_BackBuffer.BltFast(x2 - (NewPlayerX * PIC_X) - NewXOffset, y2 - (NewPlayerY * PIC_Y) - NewYOffset, DD_EmoticonSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    End If
End Sub

Sub BltArrow(ByVal Index As Long)
    Dim x As Long, y As Long, i As Long, z As Long, TempX As Long, TempY As Long

    For z = 1 To MAX_PLAYER_ARROWS
        If Player(Index).Arrow(z).Arrow > 0 Then
            rec.Top = Player(Index).Arrow(z).ArrowAnim * PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            rec.Left = Player(Index).Arrow(z).ArrowPosition * PIC_X
            rec.Right = rec.Left + PIC_X

            If GetTickCount > Player(Index).Arrow(z).ArrowTime + 30 Then
                Player(Index).Arrow(z).ArrowTime = GetTickCount
                Player(Index).Arrow(z).ArrowVarX = Player(Index).Arrow(z).ArrowVarX + 10
                Player(Index).Arrow(z).ArrowVarY = Player(Index).Arrow(z).ArrowVarY + 10
            End If
            
            Select Case Player(Index).Arrow(z).ArrowPosition
                Case 0
                    x = Player(Index).Arrow(z).ArrowX
                    y = Player(Index).Arrow(z).ArrowY + (Player(Index).Arrow(z).ArrowVarY \ 32)

                    If y > Player(Index).Arrow(z).ArrowY + Arrows(Player(Index).Arrow(z).ArrowNum).Range - 2 Then
                        Player(Index).Arrow(z).Arrow = 0
                    End If

                    If y <= MAX_MAPY Then
                        Call DD_BackBuffer.BltFast((Player(Index).Arrow(z).ArrowX - NewPlayerX) * PIC_X + sx - NewXOffset, (Player(Index).Arrow(z).ArrowY - NewPlayerY) * PIC_Y + sx - NewYOffset + Player(Index).Arrow(z).ArrowVarY, DD_ArrowAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If
                Case 1
                    x = Player(Index).Arrow(z).ArrowX
                    y = Player(Index).Arrow(z).ArrowY - (Player(Index).Arrow(z).ArrowVarY \ 32)

                    If y < Player(Index).Arrow(z).ArrowY - Arrows(Player(Index).Arrow(z).ArrowNum).Range + 2 Then
                        Player(Index).Arrow(z).Arrow = 0
                    End If

                    If y >= 0 Then
                        Call DD_BackBuffer.BltFast((Player(Index).Arrow(z).ArrowX - NewPlayerX) * PIC_X + sx - NewXOffset, (Player(Index).Arrow(z).ArrowY - NewPlayerY) * PIC_Y + sx - NewYOffset - Player(Index).Arrow(z).ArrowVarY, DD_ArrowAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If
                Case 2
                    x = Player(Index).Arrow(z).ArrowX + (Player(Index).Arrow(z).ArrowVarX \ 32)
                    y = Player(Index).Arrow(z).ArrowY

                    If x > Player(Index).Arrow(z).ArrowX + Arrows(Player(Index).Arrow(z).ArrowNum).Range - 2 Then
                        Player(Index).Arrow(z).Arrow = 0
                    End If

                    If x <= MAX_MAPX Then
                        Call DD_BackBuffer.BltFast((Player(Index).Arrow(z).ArrowX - NewPlayerX) * PIC_X + sx - NewXOffset + Player(Index).Arrow(z).ArrowVarX, (Player(Index).Arrow(z).ArrowY - NewPlayerY) * PIC_Y + sx - NewYOffset, DD_ArrowAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If
                Case 3
                    x = Player(Index).Arrow(z).ArrowX - (Player(Index).Arrow(z).ArrowVarX \ 32)
                    y = Player(Index).Arrow(z).ArrowY

                    If x < Player(Index).Arrow(z).ArrowX - Arrows(Player(Index).Arrow(z).ArrowNum).Range + 2 Then
                        Player(Index).Arrow(z).Arrow = 0
                    End If

                    If x >= 0 Then
                        Call DD_BackBuffer.BltFast((Player(Index).Arrow(z).ArrowX - NewPlayerX) * PIC_X + sx - NewXOffset - Player(Index).Arrow(z).ArrowVarX, (Player(Index).Arrow(z).ArrowY - NewPlayerY) * PIC_Y + sx - NewYOffset, DD_ArrowAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If
            End Select

            If x >= 0 And x <= MAX_MAPX And y >= 0 And y <= MAX_MAPY And Map(GetPlayerMap(MyIndex)).Tile(x, y).Type = TILE_TYPE_BLOCKED Then
                Player(Index).Arrow(z).Arrow = 0
            End If
            
            If Map(GetPlayerMap(MyIndex)).Tile(x, y).Type = TILE_TYPE_SWITCH Then
                If (Map(GetPlayerMap(MyIndex)).Tile(x, y).Data3 And 2) = 2 Then ' Advanced Bit Logic, ask for help before changing this line.
                    Call SendData(CPackets.Carrowswitch & SEP_CHAR & x & SEP_CHAR & y & END_CHAR)
                    Player(Index).Arrow(z).Arrow = 0
                End If
            End If
            
            ' Respawn the bullet bill if it hits a Dodgebill tile
            If GetPlayerMap(MyIndex) = 188 Then
                If Map(GetPlayerMap(MyIndex)).Tile(x, y).Type = TILE_TYPE_DODGEBILL Then
                    If Index = MyIndex Then
                        ' Find out where to spawn the bullet bill in relation to its direction
                        Select Case Player(Index).Arrow(z).ArrowPosition
                            Case DIR_UP
                                TempX = x
                                TempY = y - 1
                            Case DIR_DOWN
                                TempX = x
                                TempY = y + 1
                            Case DIR_LEFT
                                TempX = x - 1
                                TempY = y
                            Case DIR_RIGHT
                                TempX = x + 1
                                TempY = y
                        End Select

                        Call SendData(CPackets.Cdodgebillspawn & SEP_CHAR & TempX & SEP_CHAR & TempY & END_CHAR)
                        Player(Index).Arrow(z).Arrow = 0
                        Exit Sub
                    End If
                End If
            End If
            
            For i = 1 To MAX_PLAYERS
                If IsPlaying(i) Then
                    If GetPlayerMap(i) = GetPlayerMap(MyIndex) And GetPlayerX(i) = x And GetPlayerY(i) = y Then
                        If Index = MyIndex Then
                            Call SendData(CPackets.Carrowhit & SEP_CHAR & 0 & SEP_CHAR & i & SEP_CHAR & x & SEP_CHAR & y & END_CHAR)
                        End If

                        If Index <> i And Player(i).InBattle = False Then
                            Player(Index).Arrow(z).Arrow = 0
                        End If

                        Exit Sub
                    End If
                End If
            Next i

            For i = 1 To MAX_MAP_NPCS
                If MapNpc(i).num > 0 Then
                    If MapNpc(i).x = x And MapNpc(i).y = y Then
                        If Index = MyIndex Then
                            Call SendData(CPackets.Carrowhit & SEP_CHAR & 1 & SEP_CHAR & i & SEP_CHAR & x & SEP_CHAR & y & END_CHAR)
                        End If
                            
                        If MapNpc(i).InBattle = False Then
                            Player(Index).Arrow(z).Arrow = 0
                        End If
                            
                        Exit Sub
                    End If
                End If
            Next i
        End If
    Next z
End Sub

Sub BltLevelUp(ByVal Index As Long)
    Dim rec As RECT
    Dim x As Integer
    Dim y As Integer

    If Player(Index).LevelUpT + 3000 > GetTickCount Then
        If GetPlayerMap(Index) = GetPlayerMap(MyIndex) Then
            rec.Top = PIC_Y * 2
            rec.Bottom = rec.Top + PIC_Y
            rec.Left = PIC_X * 4
            rec.Right = rec.Left + 96

            x = GetPlayerX(Index) * PIC_X + Player(Index).xOffset + sx
            y = GetPlayerY(Index) * PIC_Y + Player(Index).yOffset + sx

            Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - PIC_X - NewXOffset, y - (NewPlayerY * PIC_Y) - PIC_Y - NewYOffset - 8, DD_TileSurf(10), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)

            If Player(Index).LevelUp >= 3 Then
                Player(Index).LevelUp = Player(Index).LevelUp - 1
            ElseIf Player(Index).LevelUp >= 1 Then
                Player(Index).LevelUp = Player(Index).LevelUp + 1
            End If
        Else
            Player(Index).LevelUpT = 0
        End If
    End If
End Sub

Sub BltPvPSign(ByVal x As Long, ByVal y As Long)
    rec.Top = 0
    rec.Bottom = 33 ' Height of image
    rec.Left = 0
    rec.Right = 54 ' Width of image

    Call DD_BackBuffer.BltFast(x, y, DD_PvPImageSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub

Sub BltHPIcon(ByVal x As Long, ByVal y As Long)
    rec.Top = 0
    rec.Bottom = 40
    rec.Left = 0
    rec.Right = 150

    Call DD_BackBuffer.BltFast(x, y, DD_HPImageSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub

Sub BltFPIcon(ByVal x As Long, ByVal y As Long)
    rec.Top = 0
    rec.Bottom = 40
    rec.Left = 0
    rec.Right = 150

    Call DD_BackBuffer.BltFast(x, y, DD_FPImageSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub

Sub BltExpIcon(ByVal x As Long, ByVal y As Long)
    rec.Top = 0
    rec.Bottom = 40
    rec.Left = 0
    rec.Right = 150

    Call DD_BackBuffer.BltFast(x, y, DD_ExpImageSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub

Sub BltFadedAttackImage(ByVal x As Long, ByVal y As Long)
    rec.Top = 0
    rec.Bottom = 49
    rec.Left = 0
    rec.Right = 49

    Call DD_BackBuffer.BltFast(x, y, DD_FadedAtkBattleImageSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub

Sub BltAttackImage(ByVal x As Long, ByVal y As Long)
    rec.Top = 0
    rec.Bottom = 49
    rec.Left = 0
    rec.Right = 49

    Call DD_BackBuffer.BltFast(x, y, DD_AtkBattleImageSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub

Sub BltFadedRunImage(ByVal x As Long, ByVal y As Long)
    rec.Top = 0
    rec.Bottom = 49
    rec.Left = 0
    rec.Right = 49

    Call DD_BackBuffer.BltFast(x, y, DD_FadedRunBattleImageSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub

Sub BltRunImage(ByVal x As Long, ByVal y As Long)
    rec.Top = 0
    rec.Bottom = 49
    rec.Left = 0
    rec.Right = 49

    Call DD_BackBuffer.BltFast(x, y, DD_RunBattleImageSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub

Sub BltFadedItemImage(ByVal x As Long, ByVal y As Long)
    rec.Top = 0
    rec.Bottom = 49
    rec.Left = 0
    rec.Right = 49

    Call DD_BackBuffer.BltFast(x, y, DD_FadedItemBattleImageSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub

Sub BltItemImage(ByVal x As Long, ByVal y As Long)
    rec.Top = 0
    rec.Bottom = 49
    rec.Left = 0
    rec.Right = 49

    Call DD_BackBuffer.BltFast(x, y, DD_ItemBattleImageSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub

Sub BltFadedSpecialImage(ByVal x As Long, ByVal y As Long)
    rec.Top = 0
    rec.Bottom = 49
    rec.Left = 0
    rec.Right = 49

    Call DD_BackBuffer.BltFast(x, y, DD_FadedSpecialBattleImageSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub

Sub BltSpecialImage(ByVal x As Long, ByVal y As Long)
    rec.Top = 0
    rec.Bottom = 49
    rec.Left = 0
    rec.Right = 49

    Call DD_BackBuffer.BltFast(x, y, DD_SpecialBattleImageSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub

Sub BltVictoryImage(ByVal x As Long, ByVal y As Long)
    rec.Top = 0
    rec.Bottom = DDSD_VictoryImage.lHeight
    rec.Left = 0
    rec.Right = DDSD_VictoryImage.lWidth

    Call DD_BackBuffer.BltFast(x, y, DD_VictoryImageSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub
