Attribute VB_Name = "modGlobals"
Option Explicit

' mouse cursor location
Public CurX As Long
Public CurY As Long

Public snumber As Long

' Game text buffer
Public MyText As String

' Index of actual player
Public MyIndex As Long

' Map animation #, used to keep track of what map animation is currently on
Public MapAnim As Byte
Public MapAnimTimer As Long

' Used to freeze controls when getting a new map
Public GettingMap As Boolean

' For Map editor
Public ScreenMode As Byte
Public GridMode As Byte
Public MapEditorSelectedType As Byte

' Used to check if in editor or not and variables for use in editor
Public InEditor As Boolean
Public EditorTileX As Long
Public EditorTileY As Long
Public EditorWarpMap As Long
Public EditorWarpX As Long
Public EditorWarpY As Long
Public EditorSet As Byte

' Camera globals
Public ScreenX As Long
Public ScreenY As Long
Public ScreenX2 As Long
Public ScreenY2 As Long

' Used for map item editor
Public ItemEditorNum As Long
Public ItemEditorValue As Long

' Used for map key editor
Public KeyEditorNum As Long
Public KeyEditorTake As Long
Public KeyText As String

' Used for map key open editor
Public KeyOpenEditorX As Long
Public KeyOpenEditorY As Long
Public KeyOpenEditorMsg As String

' Map for local use
Public SaveMapItem(1 To MAX_MAP_ITEMS) As MapItemRec
Public SaveMapNpc(1 To MAX_MAP_NPCS) As MapNpcRec

' Used for index based editors
Public InItemsEditor As Boolean
Public InNpcEditor As Boolean
Public InShopEditor As Boolean
Public InSpellEditor As Boolean
Public InElementEditor As Boolean
Public InEmoticonEditor As Boolean
Public InArrowEditor As Boolean
Public InRecipeEditor As Boolean
Public EditorIndex As Long

' Game fps
Public GameFPS As Long
Public BFPS As Boolean

' Used for atmosphere
Public GameWeather As Long
Public RainIntensity As Long

' Scrolling Variables
Public NewPlayerX As Long
Public NewPlayerY As Long
Public NewXOffset As Long
Public NewYOffset As Long
Public NewX As Long
Public NewY As Long

' Damage Variables
Public DmgDamage As Long
Public DmgTime As Long
Public NPCDmgDamage As Long
Public NPCDmgTime As Long
Public NPCWho As Long

Public EditorItemX As Long
Public EditorItemY As Long

Public EditorShopNum As Long

Public EditorItemNum1 As Byte
Public EditorItemNum2 As Byte
Public EditorItemNum3 As Byte

Public Arena1 As Byte
Public Arena2 As Byte
Public Arena3 As Byte

Public ii As Long, iii As Long
Public sx As Long

Public SpritePic As Long

Public SoundFileName As String
Public SpellSoundFileName As String

Public SignLine1 As String
Public SignLine2 As String
Public SignLine3 As String

Public ClassChange As Long
Public ClassChangeReq As Long

Public NoticeTitle As String
Public NoticeText As String
Public NoticeSound As String

Public ScriptNum As Long

Public Connected As Boolean

' Used for NPC spawn
Public NPCSpawnNum As Long

' Used for roof tile
Public RoofId As String

Public AutoLogin As Long

' Used to make sure we have all the data before logging in
Public AllDataReceived As Boolean

' Last Direction
Public LAST_DIR As Long

' Keeps/deletes the player's username when personal messaging
Public KeepUsername As Boolean

' Keep track of time
Public Hours As Integer
Public Minutes As Integer
Public Seconds As Integer
Public Gamespeed As Integer

' Font data
Public Font As String
Public fontsize As Byte
Public Font2 As String

Public SOffsetX As Integer
Public SOffsetY As Integer

Public BLoc As Boolean

Public ServerIP As String
Public PlayerBuffer As String
Public InGame As Boolean

Public TexthDC As Long
Public GameFont As Long
Public GameFont2 As Long
Public GameFont3 As Long
Public GameFont4 As Long

' Game direction vars
Public DirUp As Boolean
Public DirDown As Boolean
Public DirLeft As Boolean
Public DirRight As Boolean
Public ShiftDown As Boolean
Public ControlDown As Boolean

' General constants
Public temp As Long
Public lvl As Integer

Public Anim1Data As Long
Public Anim2Data As Long
Public M2AnimData As Long
Public FAnimData As Long
Public F2AnimData As Long

' OnClick tile info
Public ClickScript As Integer

' Minus Stat values
Public MinusHp As Integer
Public MinusMp As Integer
Public MinusSp As Integer
Public MessageMinus As String

' Switch attribute values
Public SwitchWarpMap As Long
Public SwitchWarpPos As Long
Public SwitchWarpFlags As Long

' Level Block attribute values
Public LevelToBlock As Integer

' Question Block attribute values
Public ItemThing1 As Long
Public ItemThing2 As Long
Public ItemThing3 As Long
Public ItemThing4 As Long
Public ItemThing5 As Long
Public ItemThing6 As Long
Public ChanceThing1 As Long
Public ChanceThing2 As Long
Public ChanceThing3 As Long
Public ChanceThing4 As Long
Public ChanceThing5 As Long
Public ChanceThing6 As Long
Public ValueThing1 As Long
Public ValueThing2 As Long
Public ValueThing3 As Long
Public ValueThing4 As Long
Public ValueThing5 As Long
Public ValueThing6 As Long

' Jump Block attribute values
Public JumpHeight As Byte
Public JumpDecrease As Byte
Public JumpDir(1 To 4) As Byte
Public JumpDirAddHeight(1 To 4) As Byte

' Jugem's Cloud attribute values
Public CloudDir As Byte
Public JugemsCloudHolder As String

' Bean Tile attribute values
Public BeanItemNum As Long
Public BeanItemQuantity As Integer

' Playing sound
Public CurrentSong As String
Public MapMusicStarted As Boolean

' Bubble thing
Public Bubble(1 To MAX_PLAYERS) As ChatBubble

' Used for parsing
Public SEP_CHAR As String * 1
Public END_CHAR As String * 1

Public Map(0 To MAX_MAPS) As MapRec
Public TempTile() As TempTileRec
Public Player(1 To MAX_PLAYERS) As PlayerRec
Public Class(0 To MAX_CLASSES) As ClassRec
Public Item(1 To MAX_ITEMS) As ItemRec
Public Npc(1 To MAX_NPCS) As NpcRec
Public MapItem(1 To MAX_MAP_ITEMS) As MapItemRec
Public MapNpc(1 To MAX_MAP_NPCS) As MapNpcRec
Public Shop(1 To MAX_SHOPS) As ShopRec
Public Spell(1 To MAX_SPELLS) As SpellRec
Public Element(0 To MAX_ELEMENTS) As ElementRec
Public Emoticons(0 To MAX_EMOTICONS) As EmoRec
Public ScriptSpell(1 To MAX_SCRIPTSPELLS) As ScriptSpellAnimRec
Public Recipe(1 To MAX_RECIPES) As RecipeRec

Public MAX_RAINDROPS As Long
Public BLT_RAIN_DROPS As Long
Public DropRain() As DropRainRec

Public BLT_SNOW_DROPS As Long
Public DropSnow() As DropRainRec
Public Arrows(1 To MAX_ARROWS) As ArrowRec

Public BattlePMsg() As BattleMsgRec
Public BattleMMsg() As BattleMsgRec

Public QuestionBlock() As QuestionBlockRec

Public Inventory As Long
Public slot As Long

' Variable for viewing additional inventory slots when the Down button is clicked
Public InventorySlotsIndex As Integer

Public Direct As Long
Public GuildBlock As String

' Used for trading
Public PlayerTrading(1 To MAX_PLAYER_TRADES) As PlayerTradeRec
Public OtherPlayerTrading(1 To MAX_PLAYER_TRADES) As PlayerTradeRec
Public IsTrading As Boolean

' Turn-based battle variables
Public TurnBasedTime As Long
Public TurnBasedTimeToWait As Long
Public TurnBasedTimer As Boolean
Public PlayerTurn As Integer

' Turn-based battle victory variables
Public BattleVictoryTimer As Long
Public BattleFrameCount As Byte
Public BattleNPC As Long
Public StartedVictoryAnim As Boolean
Public CanFinishBattle As Boolean
Public DisplayInfo As Boolean
Public VictoryInfo(1 To 7) As String

' Turn-based battle graphic variables
Public IsPlayerTurn As Boolean
Public ButtonHighlighted As Integer
Public CanUseItem As Boolean
Public CanUseSpecial As Boolean
Public HasAttacked As Boolean

' Cooking variables
Public CookingTime As Long
Public RecipeNumber As Long
Public IsCooking As Boolean
Public CookingTimer As Boolean
Public CookNpcNum As Long

Public IsChefBeanB As Boolean

' Special Badge timer
Public SpecialBadgeTime As Long

' Simultaneous Block variable
Public SimulBlockCoords(1 To 4) As String
Public SimulBlockWarpCoords(1 To 2) As Long

' Other
Public IsBanking As Boolean
Public IsShopping As Boolean

' Hide n' Sneak variables
Public IsHiderFrozen As Boolean
Public IsPlayingHideNSneak As Boolean

Public TimerThing As Long
