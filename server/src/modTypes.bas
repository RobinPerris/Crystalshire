Attribute VB_Name = "modTypes"
Option Explicit

' Public data structures
Public Map(1 To MAX_MAPS) As MapRec
Public MapCRC32(1 To MAX_MAPS) As MapCRCStruct
Public MapCache(1 To MAX_MAPS) As Cache
Public TempTile(1 To MAX_MAPS) As TempTileRec
Public PlayersOnMap(1 To MAX_MAPS) As Long
Public ResourceCache(1 To MAX_MAPS) As ResourceCacheRec
Public Player(1 To MAX_PLAYERS) As PlayerRec
Public TempPlayer(1 To MAX_PLAYERS) As TempPlayerRec
Public Class() As ClassRec
Public Item(1 To MAX_ITEMS) As ItemRec
Public Npc(1 To MAX_NPCS) As NpcRec
Public MapItem(1 To MAX_MAPS, 1 To MAX_MAP_ITEMS) As MapItemRec
Public MapNpc(1 To MAX_MAPS) As MapNpcDataRec
Public Shop(1 To MAX_SHOPS) As ShopRec
Public Spell(1 To MAX_SPELLS) As SpellRec
Public Resource(1 To MAX_RESOURCES) As ResourceRec
Public Animation(1 To MAX_ANIMATIONS) As AnimationRec
Public Party(1 To MAX_PARTYS) As PartyRec
Public Conv(1 To MAX_CONVS) As ConvWrapperRec
Public Options As OptionsRec

Public Type MapCRCStruct
    MapDataCRC As Long
    MapTileCRC As Long
End Type

Private Type OptionsRec
    MOTD As String
End Type

Public Type PartyRec
    Leader As Long
    Member(1 To MAX_PARTY_MEMBERS) As Long
    MemberCount As Long
End Type

Public Type PlayerInvRec
    Num As Long
    Value As Long
    Bound As Byte
End Type

Public Type PlayerSpellRec
    Spell As Long
    Uses As Long
End Type

Private Type Cache
    Data() As Byte
End Type

Public Type HotbarRec
    Slot As Long
    sType As Byte
End Type

Public Type PlayerRec
    ' Account
    Login As String
    
    ' General
    Name As String * ACCOUNT_LENGTH
    Sex As Byte
    Class As Long
    Sprite As Long
    Level As Byte
    exp As Long
    Access As Byte
    PK As Byte
    
    ' Vitals
    Vital(1 To Vitals.Vital_Count - 1) As Long
    
    ' Stats
    Stat(1 To Stats.Stat_Count - 1) As Byte
    POINTS As Long
    
    ' Worn equipment
    Equipment(1 To Equipment.Equipment_Count - 1) As Long
    
    ' Inventory
    Inv(1 To MAX_INV) As PlayerInvRec
    Spell(1 To MAX_PLAYER_SPELLS) As PlayerSpellRec
    Bank(1 To MAX_BANK) As PlayerInvRec
    
    ' Hotbar
    Hotbar(1 To MAX_HOTBAR) As HotbarRec
    
    ' Position
    Map As Long
    x As Byte
    y As Byte
    dir As Byte
    
    ' Variables
    Variable(1 To MAX_BYTE) As Long
    
    ' Tutorial
    TutorialState As Byte
    
    ' Banned
    isBanned As Byte
    isMuted As Byte
    
    ' character selection
    charNum As Long
    
    ' usergroup
    Usergroup As Long
End Type

Public Type SpellBufferRec
    Spell As Long
    Timer As Long
    target As Long
    tType As Byte
End Type

Public Type DoTRec
    Used As Boolean
    Spell As Long
    Timer As Long
    Caster As Long
    StartTime As Long
End Type

Public Type TempPlayerRec
    ' Non saved local vars
    Buffer As clsBuffer
    InGame As Boolean
    AttackTimer As Long
    DataTimer As Long
    DataBytes As Long
    DataPackets As Long
    targetType As Byte
    target As Long
    GettingMap As Byte
    SpellCD(1 To MAX_PLAYER_SPELLS) As Long
    InShop As Long
    StunTimer As Long
    StunDuration As Long
    InBank As Boolean
    inEvent As Boolean
    eventNum As Long
    pageNum As Long
    commandNum As Long
    ' trade
    TradeRequest As Long
    InTrade As Long
    TradeOffer(1 To MAX_INV) As PlayerInvRec
    AcceptTrade As Boolean
    ' dot/hot
    DoT(1 To MAX_DOTS) As DoTRec
    HoT(1 To MAX_DOTS) As DoTRec
    ' spell buffer
    spellBuffer As SpellBufferRec
    ' regen
    stopRegen As Boolean
    stopRegenTimer As Long
    ' party
    inParty As Long
    partyInvite As Long
    ' chat
    inChatWith As Long
    curChat As Long
    c_mapNum As Long
    c_mapNpcNum As Long
    ' food
    foodItem(1 To Vitals.Vital_Count - 1) As Long
    foodTick(1 To Vitals.Vital_Count - 1) As Long
    foodTimer(1 To Vitals.Vital_Count - 1) As Long
End Type

Private Type TempEventRec
    x As Long
    y As Long
    SelfSwitch As Byte
End Type

Private Type EventCommandRec
    Type As Byte
    Text As String
    colour As Long
    Channel As Byte
    targetType As Byte
    target As Long
End Type

Private Type EventPageRec
    chkPlayerVar As Byte
    chkSelfSwitch As Byte
    chkHasItem As Byte
    
    PlayerVarNum As Long
    SelfSwitchNum As Long
    HasItemNum As Long
    
    PlayerVariable As Long
    
    GraphicType As Byte
    Graphic As Long
    GraphicX As Long
    GraphicY As Long
    
    MoveType As Byte
    MoveSpeed As Byte
    MoveFreq As Byte
    
    WalkAnim As Byte
    StepAnim As Byte
    DirFix As Byte
    WalkThrough As Byte
    
    Priority As Byte
    Trigger As Byte
    
    CommandCount As Long
    Commands() As EventCommandRec
End Type

Private Type EventRec
    Name As String
    x As Long
    y As Long
    PageCount As Long
    EventPage() As EventPageRec
End Type

Private Type MapDataRec
    Name As String
    Music As String
    Moral As Byte
    
    Up As Long
    Down As Long
    left As Long
    Right As Long
    
    BootMap As Long
    BootX As Byte
    BootY As Byte
    
    MaxX As Byte
    MaxY As Byte
    
    BossNpc As Long
    
    Npc(1 To MAX_MAP_NPCS) As Long
End Type

Private Type TileDataRec
    x As Long
    y As Long
    Tileset As Long
End Type

Public Type TileRec
    Layer(1 To MapLayer.Layer_Count - 1) As TileDataRec
    Autotile(1 To MapLayer.Layer_Count - 1) As Byte

    Type As Byte
    Data1 As Long
    Data2 As Long
    Data3 As Long
    Data4 As Long
    Data5 As Long
    DirBlock As Byte
End Type

Private Type MapTileRec
    EventCount As Long
    Tile() As TileRec
    Events() As EventRec
End Type

Private Type MapRec
    MapData As MapDataRec
    TileData As MapTileRec
End Type

Private Type ClassRec
    Name As String * NAME_LENGTH
    Stat(1 To Stats.Stat_Count - 1) As Byte
    MaleSprite() As Long
    FemaleSprite() As Long
    
    startItemCount As Long
    StartItem() As Long
    StartValue() As Long
    
    startSpellCount As Long
    StartSpell() As Long
End Type

Private Type ItemRec
    Name As String * NAME_LENGTH
    Desc As String * 255
    Sound As String * NAME_LENGTH
    
    Pic As Long

    Type As Byte
    Data1 As Long
    Data2 As Long
    Data3 As Long
    ClassReq As Long
    AccessReq As Long
    LevelReq As Long
    Mastery As Byte
    price As Long
    Add_Stat(1 To Stats.Stat_Count - 1) As Byte
    Rarity As Byte
    Speed As Long
    Handed As Long
    BindType As Byte
    Stat_Req(1 To Stats.Stat_Count - 1) As Byte
    Animation As Long
    Paperdoll As Long
        
    ' consume
    AddHP As Long
    AddMP As Long
    AddEXP As Long
    CastSpell As Long
    instaCast As Byte
    
    ' food
    HPorSP As Long
    FoodPerTick As Long
    FoodTickCount As Long
    FoodInterval As Long
    
    ' requirements
    proficiency As Long
End Type

Private Type MapItemRec
    Num As Long
    Value As Long
    x As Byte
    y As Byte
    ' ownership + despawn
    playerName As String
    playerTimer As Long
    canDespawn As Boolean
    despawnTimer As Long
    Bound As Boolean
End Type

Private Type NpcRec
    Name As String * NAME_LENGTH
    AttackSay As String * 100
    Sound As String * NAME_LENGTH
    
    Sprite As Long
    SpawnSecs As Long
    Behaviour As Byte
    Range As Byte
    Stat(1 To Stats.Stat_Count - 1) As Byte
    HP As Long
    exp As Long
    Animation As Long
    damage As Long
    Level As Long
    Conv As Long
    
    ' Npc drops
    DropChance(1 To MAX_NPC_DROPS) As Double
    DropItem(1 To MAX_NPC_DROPS) As Byte
    DropItemValue(1 To MAX_NPC_DROPS) As Integer
    
    ' Casting
    Spirit As Long
    Spell(1 To MAX_NPC_SPELLS) As Long
End Type

Private Type MapNpcRec
    Num As Long
    target As Long
    targetType As Byte
    Vital(1 To Vitals.Vital_Count - 1) As Long
    x As Byte
    y As Byte
    dir As Byte
    ' For server use only
    SpawnWait As Long
    AttackTimer As Long
    StunDuration As Long
    StunTimer As Long
    ' regen
    stopRegen As Boolean
    stopRegenTimer As Long
    ' dot/hot
    DoT(1 To MAX_DOTS) As DoTRec
    HoT(1 To MAX_DOTS) As DoTRec
    ' chat
    c_lastDir As Byte
    c_inChatWith As Long
    ' spell casting
    spellBuffer As SpellBufferRec
    SpellCD(1 To MAX_NPC_SPELLS) As Long
End Type

Private Type MapNpcDataRec
    Npc(1 To MAX_MAP_NPCS) As MapNpcRec
End Type

Private Type TradeItemRec
    Item As Long
    ItemValue As Long
    costitem As Long
    costvalue As Long
End Type

Private Type ShopRec
    Name As String * NAME_LENGTH
    BuyRate As Long
    TradeItem(1 To MAX_TRADES) As TradeItemRec
End Type

Private Type SpellRec
    Name As String * NAME_LENGTH
    Desc As String * 255
    Sound As String * NAME_LENGTH
    
    Type As Byte
    mpCost As Long
    LevelReq As Long
    AccessReq As Long
    ClassReq As Long
    CastTime As Long
    CDTime As Long
    Icon As Long
    Map As Long
    x As Long
    y As Long
    dir As Byte
    Vital As Long
    Duration As Long
    Interval As Long
    Range As Byte
    IsAoE As Boolean
    AoE As Long
    CastAnim As Long
    SpellAnim As Long
    StunDuration As Long
    
    ' ranking
    UniqueIndex As Long
    NextRank As Long
    NextUses As Long
End Type

Private Type TempTileRec
    DoorOpen() As Byte
    DoorTimer As Long
End Type

Private Type TempMapDataRec
    Npc() As MapNpcRec
End Type

Private Type MapResourceRec
    ResourceState As Byte
    ResourceTimer As Long
    x As Long
    y As Long
    cur_health As Long
End Type

Private Type ResourceCacheRec
    Resource_Count As Long
    ResourceData() As MapResourceRec
End Type

Private Type ResourceRec
    Name As String * NAME_LENGTH
    SuccessMessage As String * NAME_LENGTH
    EmptyMessage As String * NAME_LENGTH
    Sound As String * NAME_LENGTH
    
    ResourceType As Byte
    ResourceImage As Long
    ExhaustedImage As Long
    ItemReward As Long
    ToolRequired As Long
    health As Long
    RespawnTime As Long
    WalkThrough As Boolean
    Animation As Long
End Type

Private Type AnimationRec
    Name As String * NAME_LENGTH
    Sound As String * NAME_LENGTH
    
    Sprite(0 To 1) As Long
    Frames(0 To 1) As Long
    LoopCount(0 To 1) As Long
    LoopTime(0 To 1) As Long
End Type

Private Type ConvRec
    Conv As String
    rText(1 To 4) As String
    rTarget(1 To 4) As Long
    Event As Long
    Data1 As Long
    Data2 As Long
    Data3 As Long
End Type

Private Type ConvWrapperRec
    Name As String * NAME_LENGTH
    chatCount As Long
    Conv() As ConvRec
End Type
