Attribute VB_Name = "modConstants"
Option Explicit

' Winsock globals
Public Const GAME_PORT As Integer = 0000

' Global Variables
Public Const MAX_CLASSES As Integer = 5
Public MAX_SERVLINES As Long
Public LEVEL As Integer
Public SAVETIME As Long

' Global Timers.
Public CHATLOG_TIMER As Long
Public SHUTDOWN_TIMER As Long
Public PLYRSAVE_TIMER As Long

' Max Settings
Public Const MAX_PLAYERS As Long = 30
Public Const MAX_ITEMS As Long = 400
Public Const MAX_NPCS As Long = 250
Public Const MAX_SHOPS As Integer = 50
Public Const MAX_SPELLS As Integer = 100
Public Const MAX_MAPS As Long = 350
Public Const MAX_MAP_ITEMS As Long = 50
Public Const MAX_GUILDS As Integer = 20
Public Const MAX_GUILD_MEMBERS As Integer = 15
Public Const MAX_EMOTICONS As Integer = 20
Public Const MAX_ELEMENTS As Integer = 20
Public Const MAX_LEVEL As Integer = 75
Public Const MAX_SCRIPTSPELLS As Integer = 30
Public Const MAX_RECIPES As Integer = 100
Public Const MAX_ARROWS As Integer = 100
Public Const MAX_INV As Long = 24
Public Const MAX_BANK As Integer = 100
Public Const MAX_MAP_NPCS As Long = 15
Public Const MAX_PLAYER_SPELLS As Integer = 20
Public Const MAX_PLAYER_TRADES As Integer = 8
Public Const MAX_NPC_DROPS As Integer = 10
Public Const MAX_SHOP_ITEMS As Integer = 25
Public Const MAX_PARTY_MEMBERS As Byte = 4

' Choice Constants.
Public Const NO As Byte = 0
Public Const YES As Byte = 1

' Account Constants.
Public Const NAME_LENGTH As Integer = 20
Public Const MAX_CHARS As Byte = 3

' Basic Security.
Public Const SEC_CODE As String = "270"

' Map Coords.
Public Const MAX_MAPX As Long = 30
Public Const MAX_MAPY As Long = 30

' Map Morals.
Public Const MAP_MORAL_NONE As Byte = 0
Public Const MAP_MORAL_SAFE As Byte = 1
Public Const MAP_MORAL_NO_PENALTY As Byte = 2
Public Const MAP_MORAL_MINIGAME As Byte = 3

' Tile Consants.
Public Const TILE_TYPE_WALKABLE As Byte = 0
Public Const TILE_TYPE_BLOCKED As Byte = 1
Public Const TILE_TYPE_WARP As Byte = 2
Public Const TILE_TYPE_ITEM As Byte = 3
Public Const TILE_TYPE_NPCAVOID As Byte = 4
Public Const TILE_TYPE_KEY As Byte = 5
Public Const TILE_TYPE_KEYOPEN As Byte = 6
Public Const TILE_TYPE_HEAL As Byte = 7
Public Const TILE_TYPE_KILL As Byte = 8
Public Const TILE_TYPE_SHOP As Byte = 9
Public Const TILE_TYPE_CBLOCK As Byte = 10
Public Const TILE_TYPE_ARENA As Byte = 11
Public Const TILE_TYPE_SOUND As Byte = 12
Public Const TILE_TYPE_SPRITE_CHANGE As Byte = 13
Public Const TILE_TYPE_SIGN As Byte = 14
Public Const TILE_TYPE_DOOR As Byte = 15
Public Const TILE_TYPE_NOTICE As Byte = 16
Public Const TILE_TYPE_CHEST As Byte = 17
Public Const TILE_TYPE_CLASS_CHANGE As Byte = 18
Public Const TILE_TYPE_SCRIPTED As Byte = 19
Public Const TILE_TYPE_HOUSE As Byte = 21
Public Const TILE_TYPE_BANK As Byte = 23
Public Const TILE_TYPE_GUILDBLOCK As Byte = 25
Public Const TILE_TYPE_HOOKSHOT As Byte = 26
Public Const TILE_TYPE_WALKTHRU As Byte = 27
Public Const TILE_TYPE_ROOF As Byte = 28
Public Const TILE_TYPE_ROOFBLOCK As Byte = 29
Public Const TILE_TYPE_ONCLICK As Byte = 30
Public Const TILE_TYPE_LOWER_STAT As Byte = 31
Public Const TILE_TYPE_SWITCH As Byte = 32
Public Const TILE_TYPE_LVLBLOCK As Byte = 33
Public Const TILE_TYPE_QUESTIONBLOCK As Byte = 34
Public Const TILE_TYPE_DRILL As Byte = 35
Public Const TILE_TYPE_JUMPBLOCK As Byte = 36
Public Const TILE_TYPE_DODGEBILL As Byte = 37
Public Const TILE_TYPE_HAMMERBARRAGE As Byte = 38
Public Const TILE_TYPE_JUGEMSCLOUD As Byte = 39
Public Const TILE_TYPE_SIMULBLOCK As Byte = 40
Public Const TILE_TYPE_BEAN As Byte = 41

' Item Constants.
Public Const ITEM_TYPE_NONE As Byte = 0
Public Const ITEM_TYPE_WEAPON As Byte = 1
Public Const ITEM_TYPE_TWO_HAND As Byte = 2
Public Const ITEM_TYPE_ARMOR As Byte = 3
Public Const ITEM_TYPE_HELMET As Byte = 4
Public Const ITEM_TYPE_SPECIALBADGE As Byte = 5
Public Const ITEM_TYPE_LEGS As Byte = 6
Public Const ITEM_TYPE_FLOWERBADGE As Byte = 7
Public Const ITEM_TYPE_MUSHROOMBADGE As Byte = 8
Public Const ITEM_TYPE_CHANGEHPFPSP As Byte = 9
Public Const ITEM_TYPE_KEY As Byte = 10
Public Const ITEM_TYPE_CURRENCY As Byte = 11
Public Const ITEM_TYPE_SPELL As Byte = 12
Public Const ITEM_TYPE_SCRIPTED As Byte = 13
Public Const ITEM_TYPE_AMMO As Byte = 14
Public Const ITEM_TYPE_CARD As Byte = 15

' Direction Constants.
Public Const DIR_UP As Byte = 0
Public Const DIR_DOWN As Byte = 1
Public Const DIR_LEFT As Byte = 2
Public Const DIR_RIGHT As Byte = 3

' Player Movement Constants.
Public Const MOVING_WALKING As Byte = 1
Public Const MOVING_RUNNING As Byte = 2

' Weather Constants.
Public Const WEATHER_NONE As Byte = 0
Public Const WEATHER_RAINING As Byte = 1
Public Const WEATHER_SNOWING As Byte = 2
Public Const WEATHER_THUNDER As Byte = 3

' Admin Constants.
Public Const ADMIN_MONITER As Byte = 1
Public Const ADMIN_MAPPER As Byte = 2
Public Const ADMIN_DEVELOPER As Byte = 3
Public Const ADMIN_CREATOR As Byte = 4

' NPC Constants.
Public Const NPC_BEHAVIOR_ATTACKONSIGHT As Byte = 0
Public Const NPC_BEHAVIOR_ATTACKWHENATTACKED As Byte = 1
Public Const NPC_BEHAVIOR_FRIENDLY As Byte = 2
Public Const NPC_BEHAVIOR_SHOPKEEPER As Byte = 3
Public Const NPC_BEHAVIOR_GUARD As Byte = 4
Public Const NPC_BEHAVIOR_SCRIPTED As Byte = 5

' Spell Constants.
Public Const SPELL_TYPE_ADDHP As Long = 0
Public Const SPELL_TYPE_ADDMP As Long = 1
Public Const SPELL_TYPE_ADDSP As Long = 2
Public Const SPELL_TYPE_SUBHP As Long = 3
Public Const SPELL_TYPE_SUBMP As Long = 4
Public Const SPELL_TYPE_SUBSP As Long = 5
Public Const SPELL_TYPE_SCRIPTED As Long = 6
Public Const SPELL_TYPE_STATCHANGE As Long = 7

' Target Type Constants.
Public Const TARGET_TYPE_PLAYER As Byte = 0
Public Const TARGET_TYPE_NPC As Byte = 1

' Version Constants.
Public Const CLIENT_MAJOR As Byte = 2
Public Const CLIENT_MINOR As Byte = 0
Public Const CLIENT_REVISION As Byte = 0

Public Const ADMIN_LOG As String = "Logs\Admin.txt"
Public Const PLAYER_LOG As String = "Logs\Player.txt"

' Max Lines On Console.
Public Const MAX_LINES = 5000

' Default System Colors.
Public Const BLACK As Byte = 0
Public Const BLUE As Byte = 1
Public Const GREEN As Byte = 2
Public Const CYAN As Byte = 3
Public Const RED As Byte = 4
Public Const MAGENTA As Byte = 5
Public Const BROWN As Byte = 6
Public Const GREY As Byte = 7
Public Const DARKGREY As Byte = 8
Public Const BRIGHTBLUE As Byte = 9
Public Const BRIGHTGREEN As Byte = 10
Public Const BRIGHTCYAN As Byte = 11
Public Const BRIGHTRED As Byte = 12
Public Const PINK As Byte = 13
Public Const YELLOW As Byte = 14
Public Const WHITE As Byte = 15

' Default Message Colors.
Public Const SayColor As Byte = GREY
Public Const GlobalColor As Byte = GREEN
Public Const BroadcastColor As Byte = WHITE
Public Const TellColor As Byte = WHITE
Public Const EmoteColor As Byte = WHITE
Public Const AdminColor As Byte = BRIGHTCYAN
Public Const HelpColor As Byte = WHITE
Public Const WhoColor As Byte = GREY
Public Const JoinLeftColor As Byte = GREY
Public Const NpcColor As Byte = WHITE
Public Const AlertColor As Byte = WHITE
Public Const NewMapColor As Byte = GREY

' Minigame Player Constants
Public Const STSMaxPlayers As Byte = 4
Public Const DodgeBillMaxPlayers As Byte = 5
Public Const MaxSeekers As Byte = 3
Public Const MaxHiders As Byte = 5
