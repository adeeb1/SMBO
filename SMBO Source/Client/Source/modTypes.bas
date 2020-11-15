Attribute VB_Name = "modTypes"
Option Explicit

Type ChatBubble
    Text As String
    Created As Long
End Type

Type BankRec
    num As Long
    Value As Long
    Ammo As Long
End Type

Type PlayerInvRec
    num As Long
    Value As Long
    Ammo As Long
End Type

Type ElementRec
    Name As String * 50
    Strong As Long
    Weak As Long
End Type

Type SpellAnimRec
    CastedSpell As Byte

    SpellTime As Long
    SpellVar As Long
    SpellDone As Long

    Target As Long
    TargetType As Long
End Type

Type ScriptSpellAnimRec
    CastedSpell As Byte

    SpellTime As Long
    SpellVar As Long
    SpellDone As Long

    SpellNum As Long
    x As Long
    y As Long
    Index As Long
End Type

Type PlayerArrowRec
    Arrow As Byte
    ArrowNum As Long
    ArrowAnim As Long
    ArrowTime As Long
    ArrowVarX As Long
    ArrowVarY As Long
    ArrowX As Long
    ArrowY As Long
    ArrowPosition As Byte
    ArrowAmount As Long
End Type

Type PlayerRec
    ' General
    Name As String * NAME_LENGTH
    Guild As String
    Guildaccess As Byte
    Class As Long
    Sprite As Long
    Level As Long
    Exp As Long
    Access As Byte
    PK As Byte
    Step As Byte

    ' Vitals
    HP As Long
    MP As Long
    SP As Long

    ' Stats
    STR As Long
    DEF As Long
    speed As Long
    MAGI As Long
    POINTS As Long

    ' Inventory
    Inv(1 To MAX_INV) As PlayerInvRec
    Spell(1 To MAX_PLAYER_SPELLS) As Long
    Bank(1 To MAX_BANK) As BankRec
    
    ' Equipment
    Equipment(1 To 7) As PlayerInvRec
    
    ' Position
    Map As Long
    x As Long
    y As Long
    Dir As Long

    ' Client use only
    MaxHp As Long
    MaxMP As Long
    MaxSP As Long
    xOffset As Integer
    yOffset As Integer
    MovingH As Integer
    MovingV As Integer
    Moving As Byte
    Attacking As Byte
    AttackTimer As Long
    MapGetTimer As Long
    CastedSpell As Byte

    SpellNum As Long
    SpellAnim(1 To MAX_SPELL_ANIM) As SpellAnimRec

    EmoticonNum As Long
    EmoticonTime As Long
    EmoticonVar As Long

    LevelUp As Long
    LevelUpT As Long

    Arrow(1 To MAX_PLAYER_ARROWS) As PlayerArrowRec

    SkilLvl() As Long
    SkilExp() As Long

    Armor As Long
    Helmet As Long
    Shield As Long
    Weapon As Long
    legs As Long
    Ring As Long
    Necklace As Long
    Color As Long

    head As Long
    body As Long
    leg As Long

    NextLvlExp As Long
    InBattle As Boolean
    CritHitChance As Double
    BlockChance As Double
    Height As Integer
    Jumping As Boolean
    JumpAnim As Byte
    JumpDir As Byte
    TempJumpAnim As Byte
    JumpTime As Long
    BattleVictory As Boolean
    MaxInv As Integer
    NewInv() As PlayerInvRec
End Type

Type TileRec
    Ground As Long
    Mask As Long
    Anim As Long
    Mask2 As Long
    M2Anim As Long
    Fringe As Long
    FAnim As Long
    Fringe2 As Long
    F2Anim As Long
    Type As Byte
    Data1 As Long
    Data2 As Long
    Data3 As Long
    String1 As String
    String2 As String
    String3 As String
    light As Long
    GroundSet As Byte
    MaskSet As Byte
    AnimSet As Byte
    Mask2Set As Byte
    M2AnimSet As Byte
    FringeSet As Byte
    FAnimSet As Byte
    Fringe2Set As Byte
    F2AnimSet As Byte
End Type

Type MapRec
    Name As String * 60
    Revision As Integer
    Moral As Byte
    Up As Integer
    Down As Integer
    Left As Integer
    Right As Integer
    music As String
    BootMap As Integer
    BootX As Byte
    BootY As Byte
    Shop As Integer
    Indoors As Byte
    Tile() As TileRec
    Npc(1 To 15) As Integer
    SpawnX(1 To 15) As Byte
    SpawnY(1 To 15) As Byte
    owner As String
    scrolling As Byte
    Weather As Integer
End Type

Type ClassRec
    Name As String * NAME_LENGTH
    MaleSprite As Long
    FemaleSprite As Long
    
    Locked As Long
    
    STR As Long
    DEF As Long
    speed As Long
    MAGI As Long
    
    ' For client use
    HP As Long
    MP As Long
    SP As Long
    
    ' Description
    desc As String
End Type

Type ItemRec
    Name As String * 50
    desc As String * 150
    
    Pic As Long
    Type As Byte
    Data1 As Long
    Data2 As Long
    Data3 As Long
    StrReq As Long
    DefReq As Long
    SpeedReq As Long
    MagicReq As Long
    ClassReq As Long
    AccessReq As Byte
    
    AddHP As Long
    AddMP As Long
    AddSP As Long
    AddSTR As Long
    AddDef As Long
    AddMAGI As Long
    AddSpeed As Long
    AddEXP As Long
    AttackSpeed As Long
    Price As Long
    
    Stackable As Long
    Bound As Long
    LevelReq As Long
    HPReq As Long
    FPReq As Long
    Ammo As Long
    AddCritChance As Double
    AddBlockChance As Double
    Cookable As Boolean
End Type
    
Type MapItemRec
    num As Long
    Value As Long
    
    x As Byte
    y As Byte
    Ammo As Long
End Type

Type NPCEditorRec
    ItemNum As Long
    ItemValue As Long
    chance As Long
End Type

Type NpcRec
    Name As String * 60
    AttackSay As String * 100
    
    Sprite As Long
    SpawnSecs As Long
    Behavior As Byte
    Range As Byte
    SpriteSize As Long
    
    STR  As Long
    DEF As Long
    speed As Long
    MAGI As Long
    Big As Long
    MaxHp As Long
    Exp As Long
    SpawnTime As Long
    Spell As Long
    
    ItemNPC(1 To MAX_NPC_DROPS) As NPCEditorRec
    
    Element As Long
    AttackSay2 As String * 100
    Level As Long
End Type

Type MapNpcRec
    num As Long
    
    Target As Long
    
    HP As Long
    MaxHp As Long
    MP As Long
    SP As Long
    
    Map As Long
    x As Byte
    y As Byte
    Dir As Byte
    Big As Byte
    
    ' Client use only
    xOffset As Integer
    yOffset As Integer
    Moving As Byte
    Attacking As Byte
    AttackTimer As Long
    InBattle As Boolean
End Type

Type ShopItemRec
    ItemNum As Long
    Price As Currency
    Amount As Currency
    currencyItem As Integer
End Type

Type ShopRec
    Name As String * 50
    FixesItems As Byte
    BuysItems As Byte
    ShowInfo As Byte
    ShopItem(1 To MAX_SHOP_ITEMS) As ShopItemRec
    currencyItem As Integer
End Type

Type SpellRec
    Name As String * 50
    ClassReq As Long
    LevelReq As Long
    Sound As String
    MPCost As Long
    Type As Long
    Data1 As Long
    Data2 As Long
    Data3 As Long
    Range As Byte
    
    SpellAnim As Long
    SpellTime As Long
    SpellDone As Long
    
    AE As Long
    Big As Long
    
    Element As Long
    reload As Long
    Stat As Integer
    StatTime As Long
    Multiplier As Double
    PassiveStat As Integer
    PassiveStatChange As Integer
    UsePassiveStat As Boolean
    SelfSpell As Boolean
End Type

Type TempTileRec
    DoorOpen As Byte
End Type

Type PlayerTradeRec
    InvNum As Long
    InvName As String
    InvVal As Long
End Type

Type EmoRec
    Pic As Long
    Command As String
End Type

Type DropRainRec
    x As Long
    y As Long
    Randomized As Boolean
    speed As Byte
End Type

Type ArrowRec
    Name As String
    Pic As Long
    Range As Byte
    Amount As Long
End Type

Type BattleMsgRec
    Msg As String
    Index As Byte
    Color As Byte
    Time As Long
    Done As Byte
    y As Long
End Type

Type QuestionBlockRec
    Item1 As Long
    Item2 As Long
    Item3 As Long
    Item4 As Long
    Item5 As Long
    Item6 As Long
    Chance1 As Long
    Chance2 As Long
    Chance3 As Long
    Chance4 As Long
    Chance5 As Long
    Chance6 As Long
    Value1 As Long
    Value2 As Long
    Value3 As Long
    Value4 As Long
    Value5 As Long
    Value6 As Long
End Type

Type RecipeRec
    Ingredient1 As Long
    Ingredient2 As Long
    ResultItem As Long
    Name As String
End Type
