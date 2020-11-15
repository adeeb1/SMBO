Attribute VB_Name = "modGlobals"
Option Explicit

' Used for parsing
Public SEP_CHAR As String * 1
Public END_CHAR As String * 1

Public Map(1 To MAX_MAPS) As MapRec
Public MapCache(1 To MAX_MAPS) As String
Public TempTile(1 To MAX_MAPS) As TempTileRec
Public PlayersOnMap(1 To MAX_MAPS) As Long
Public Player(1 To MAX_PLAYERS) As AccountRec
Public ClassData(0 To MAX_CLASSES) As ClassRec
Public Item(0 To MAX_ITEMS) As ItemRec
Public NPC(0 To MAX_NPCS) As NpcRec
Public MapItem(1 To MAX_MAPS, 1 To MAX_MAP_ITEMS) As MapItemRec
Public MapNPC(1 To MAX_MAPS, 1 To MAX_MAP_NPCS) As MapNpcRec
Public Shop(1 To MAX_SHOPS) As ShopRec
Public Spell(1 To MAX_SPELLS) As SpellRec
Public Guild(1 To MAX_GUILDS) As GuildRec
Public Emoticons(0 To MAX_EMOTICONS) As EmoRec
Public Element(0 To MAX_ELEMENTS) As ElementRec
Public Experience(1 To MAX_LEVEL) As Long
Public Timer(1 To MAX_PLAYERS) As TimerRec
Public QuestionBlock() As QuestionBlockRec
Public Recipe(1 To MAX_RECIPES) As RecipeRec
Public Party(1 To MAX_PLAYERS) As PartyRec

Public Arrows(1 To MAX_ARROWS) As ArrowRec

Public VictoryInfo(1 To MAX_PLAYERS, 1 To 7) As String
Public IsInVictoryAnim(1 To MAX_PLAYERS) As Boolean

Public addSP As StatRec

Public temp As Integer

Public START_MAP As Long
Public START_X As Long
Public START_Y As Long

Global PlayerI As Byte

' Map Control
Public Const IS_SCROLLING As Long = 1

' Used for respawning items
Public SpawnSeconds As Long

' Used for weather effects
Public WeatherType As Long
Public WeatherLevel As Long
Public GameClock As String
Public Gamespeed As Long

Public Hours As Integer
Public Seconds As Long
Public Minutes As Integer

' Used for closing key doors again
Public KeyTimer As Long

' Used for gradually giving back players and npcs hp
Public GiveHPTimer As Long
Public GiveMPTimer As Long
Public GiveSPTimer As Long
Public GiveNPCHPTimer As Long

' Used for Poison Cave
Public LoseHPTimer As Long

' Used for logging
Public ServerLog As Boolean
Public TimeDisable As Boolean

' Minigame Constants
Public STSPath As String
Public DodgeBillPath As String
Public HideNSneakPath As String

' Hide n' Sneak - Left Game
Public HasLeftHideNSneak As Boolean
