Attribute VB_Name = "modTypes"
Option Explicit

' ---NOTE to future developers!----------
' When loading, types ARE order-sensitive!
' This means do not change the order of variables in between
' versions, and add new variables to the end. This way, we can
' just load the old files! I learned that the hard way :D
' -Pickle

Type NewPlayerInvRec
    num As Integer
    Value As Long
    Dur As Integer
    Ammo As Integer
End Type

Type PlayerInvRec
    num As Integer
    Value As Long
    Ammo As Integer
End Type

Type NewBankRec
    num As Integer
    Value As Long
    Dur As Integer
    Ammo As Integer
End Type

Type BankRec
    num As Integer
    Value As Long
    Ammo As Integer
End Type

Type ElementRec
    Name As String * 50
    Strong As Integer
    Weak As Integer
End Type

Type V000PlayerRec
    ' General
    Name As String * NAME_LENGTH
    Guild As String
    GuildAccess As Byte
    Sex As Byte
    Class As Integer
    Sprite As Long
    LEVEL As Integer
    Exp As Long
    Access As Byte
    PK As Byte

    ' Vitals
    HP As Long
    MP As Long
    SP As Long

    ' Stats
    STR As Long
    DEF As Long
    Speed As Long
    Magi As Long
    POINTS As Long

    ' Worn equipment
    ArmorSlot As Integer
    WeaponSlot As Integer
    HelmetSlot As Integer
    ShieldSlot As Integer
    LegsSlot As Integer
    RingSlot As Integer
    NecklaceSlot As Integer

    ' Inventory
    Inv() As PlayerInvRec
    Spell(1 To MAX_PLAYER_SPELLS) As Integer
    Bank(1 To MAX_BANK) As BankRec

    ' Position and movement
    Map As Integer
    X As Byte
    Y As Byte
    Dir As Byte

    TargetNPC As Integer

    Head As Integer
    Body As Integer
    Leg As Integer

    PAPERDOLL As Byte

    MAXHP As Long
    MAXMP As Long
    MAXSP As Long
End Type

Public Type PlayerRec
    ' General
 '090829 Scorpious2k
    Vflag As Byte       ' version flag - always > 127
    Ver As Byte
    SubVer As Byte
    Rel As Byte
 '090829 End
    Name As String * NAME_LENGTH
    Guild As String
    GuildAccess As Byte
    Sex As Byte
    Class As Integer
    Sprite As Long
    LEVEL As Integer
    Exp As Long
    Access As Byte
    PK As Byte

    ' Vitals
    HP As Long
    MP As Long
    SP As Long

    ' Stats
    STR As Long
    DEF As Long
    Speed As Long
    Magi As Long
    POINTS As Long

    ' Worn equipment
    ArmorSlot As Integer
    WeaponSlot As Integer
    HelmetSlot As Integer
    ShieldSlot As Integer
    LegsSlot As Integer
    RingSlot As Integer
    NecklaceSlot As Integer

    ' Inventory
    Inv(1 To MAX_INV) As NewPlayerInvRec
    Spell(1 To MAX_PLAYER_SPELLS) As Integer
    Bank(1 To 50) As NewBankRec

    ' Position and movement
    Map As Integer
 '090829    X As Byte
 '090829    Y As Byte
    X As Integer
    Y As Integer
    Dir As Byte

    TargetNPC As Integer

    Head As Integer
    Body As Integer
    Leg As Integer

    PAPERDOLL As Byte

    MAXHP As Long
    MAXMP As Long
    MAXSP As Long
    InBattle As Boolean
    Turn As Boolean
    OldX As Integer
    OldY As Integer
    HasTurnBased As Boolean
    RecoverTime As Long
    CritHitChance As Double
    BlockChance As Double
    ' Max Stats
    MAXSTR As Long
    MAXDEF As Long
    MAXSpeed As Long
    MAXStache As Long
    Equipment(1 To 7) As PlayerInvRec
    PartyNum As Long
    PartyInvitedBy As Long
    Height As Integer
    TempSprite As Long
    MaxInv As Integer
    NewInv() As NewPlayerInvRec
    InvConverted As Integer
    NewBank() As NewBankRec
    BankConverted As Integer
End Type

Type PlayerTradeRec
    InvNum As Long
    InvName As String
    InvVal As Long
End Type

Type PartyRec
    Leader As Long
    Member(1 To MAX_PARTY_MEMBERS) As Long
    ShareExp(1 To MAX_PARTY_MEMBERS) As Boolean
End Type

Public Type AccountRec
    ' Account
    Login As String * NAME_LENGTH
    Password As String * NAME_LENGTH
    Email As String

    ' Some error here that needs to be fixed. [Mellowz]
    Char(0 To MAX_CHARS) As PlayerRec

    ' None saved local vars
    Buffer As String
    IncBuffer As String
    CharNum As Byte
    InGame As Boolean
    AttackTimer As Long
    DataTimer As Long
    DataBytes As Long
    DataPackets As Long

    TargetType As Byte
    Target As Byte
    CastedSpell As Byte

    SpellTime As Long
    SpellVar As Long
    SpellDone As Long
    SpellNum As Long

    GettingMap As Byte

    Emoticon As Long

    InTrade As Boolean
    TradePlayer As Long
    TradeOk As Byte
    Trades(1 To MAX_PLAYER_TRADES) As PlayerTradeRec

    InChat As Byte
    ChatPlayer As Long

    Mute As Boolean
    Locked As Boolean
    LockedSpells As Boolean
    LockedItems As Boolean
    LockedAttack As Boolean
    TargetNPC As Long

    HookShotX As Byte
    HookShotY As Byte
    
    GetsDE As Boolean
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
    Light As Long
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
    NPC(1 To 15) As Integer
    SpawnX(1 To 15) As Byte
    SpawnY(1 To 15) As Byte
    Owner As String
    Scrolling As Byte
    Weather As Integer
End Type

Type ClassRec
    Name As String * NAME_LENGTH

    AdvanceFrom As Long
    LevelReq As Long
    Type As Long
    Locked As Long

    MaleSprite As Long
    FemaleSprite As Long

    STR As Long
    DEF As Long
    Speed As Long
    Magi As Long

    Map As Long
    X As Byte
    Y As Byte

    ' Description
    Desc As String
End Type

Type NewItemRec
    Name As String * 50
    Desc As String * 150

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

    addHP As Long
    addMP As Long
    addSP As Long
    AddStr As Long
    AddDef As Long
    AddMagi As Long
    AddSpeed As Long
    AddEXP As Long
    AttackSpeed As Long
    Price As Long
    Stackable As Byte
    Bound As Byte
    LevelReq As Long
    
    TwoHanded As Long
End Type

Type ItemRec
    Name As String * 50
    Desc As String * 150

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

    addHP As Long
    addMP As Long
    addSP As Long
    AddStr As Long
    AddDef As Long
    AddMagi As Long
    AddSpeed As Long
    AddEXP As Long
    AttackSpeed As Long
    Price As Long
    Stackable As Byte
    Bound As Byte
    LevelReq As Long
    HPReq As Long
    FPReq As Long
    Ammo As Long
    
    TwoHanded As Long
    AddCritChance As Double
    AddBlockChance As Double
    Cookable As Boolean
End Type

Type NewMapItemRec
    num As Long
    Value As Long
    Dur As Long
    
    X As Byte
    Y As Byte
    Ammo As Long
End Type

Type MapItemRec
    num As Long
    Value As Long
    
    X As Byte
    Y As Byte
    Ammo As Long
End Type

Type NPCEditorRec
    ItemNum As Long
    ItemValue As Long
    Chance As Long
End Type

Type NewNpcRec
    Name As String * 60
    AttackSay As String * 100

    Sprite As Long
    SpawnSecs As Long
    Behavior As Byte
    Range As Byte

    STR  As Long
    DEF As Long
    Speed As Long
    Magi As Long
    Big As Long
    MAXHP As Long
    Exp As Long
    SpawnTime As Long

    ItemNPC(1 To MAX_NPC_DROPS) As NPCEditorRec

    Element As Long

    SPRITESIZE As Byte
End Type

Type NpcRec
    Name As String * 60
    AttackSay As String * 100

    Sprite As Long
    SpawnSecs As Long
    Behavior As Byte
    Range As Byte

    STR  As Long
    DEF As Long
    Speed As Long
    Magi As Long
    Big As Long
    MAXHP As Long
    Exp As Long
    SpawnTime As Long

    ItemNPC(1 To MAX_NPC_DROPS) As NPCEditorRec

    Element As Long

    SPRITESIZE As Byte
    AttackSay2 As String * 100
    LEVEL As Long
End Type

Type MapNpcRec
    num As Long

    Target As Long

    HP As Long
    MP As Long
    SP As Long

    X As Byte
    Y As Byte
    Dir As Byte

    ' For server use only
    SpawnWait As Long
    AttackTimer As Long
    InBattle As Boolean
    Turn As Boolean
    OldX As Integer
    OldY As Integer
End Type

Type ShopItemRec
    ItemNum As Long
    Price As Currency
    Amount As Currency
    CurrencyItem As Integer
End Type

Type ShopRec
    Name As String * 50
    FixesItems As Byte
    BuysItems As Byte
    ShowInfo As Byte
    ShopItem(1 To MAX_SHOP_ITEMS) As ShopItemRec
    CurrencyItem As Integer
End Type

Type SpellRec
    Name As String * 50
    ClassReq As Long
    LevelReq As Long
    MPCost As Long
    Sound As String
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
    Stat As Integer
    StatTime As Long
    Multiplier As Double
    PassiveStat As Integer
    PassiveStatChange As Integer
    UsePassiveStat As Boolean
    SelfSpell As Boolean
End Type

Type TempTileRec
    DoorOpen()  As Byte
    DoorTimer() As Long
End Type

Type GuildRec
    Name As String * NAME_LENGTH
    Founder As String * NAME_LENGTH
    Member(1 To MAX_GUILD_MEMBERS) As String * NAME_LENGTH
End Type

Type EmoRec
    Pic As Long
    Command As String
End Type

Type ArrowRec
    Name As String
    Pic As Long
    Range As Byte
    Amount As Integer
End Type

Type StatRec
    LEVEL As Long
    STR As Long
    DEF As Long
    Magi As Long
    Speed As Long
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

Type TimerRec
    Index As Long
    Player As String
    num As Long
    Interval As Long
    WaitTime As Long
    Parameter1 As Long
    Parameter2 As Long
    Parameter3 As Long
End Type
