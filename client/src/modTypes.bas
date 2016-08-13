Attribute VB_Name = "modTypes"
Option Explicit

' Public data structures
Public map As MapRec
Public MapCRC32(1 To MAX_MAPS) As MapCRCStruct
Public Bank As BankRec
Public TempTile() As TempTileRec
Public Player(1 To MAX_PLAYERS) As PlayerRec
Public Class() As ClassRec
Public Item(1 To MAX_ITEMS) As ItemRec
Public Npc(1 To MAX_NPCS) As NpcRec
Public MapItem(1 To MAX_MAP_ITEMS) As MapItemRec
Public MapNpc(1 To MAX_MAP_NPCS) As MapNpcRec
Public Shop(1 To MAX_SHOPS) As ShopRec
Public Spell(1 To MAX_SPELLS) As SpellRec
Public Resource(1 To MAX_RESOURCES) As ResourceRec
Public Animation(1 To MAX_ANIMATIONS) As AnimationRec
Public Conv(1 To MAX_CONVS) As ConvWrapperRec
Public ActionMsg(1 To MAX_BYTE) As ActionMsgRec
Public Blood(1 To MAX_BYTE) As BloodRec
Public AnimInstance(1 To MAX_BYTE) As AnimInstanceRec
Public Party As PartyRec
Public Autotile() As AutotileRec
Public Options As OptionsRec

' Type recs
Public Type MapCRCStruct
    MapDataCRC As Long
    MapTileCRC As Long
End Type

Private Type OptionsRec
    Music As Byte
    sound As Byte
    NoAuto As Byte
    Render As Byte
    Username As String
    SaveUser As Long
    channelState(0 To Channel_Count - 1) As Byte
    PlayIntro As Byte
    Resolution As Byte
    Fullscreen As Byte
End Type

Public Type PartyRec
    Leader As Long
    Member(1 To MAX_PARTY_MEMBERS) As Long
    MemberCount As Long
End Type

Public Type PlayerInvRec
    num As Long
    value As Long
    bound As Byte
End Type

Public Type PlayerSpellRec
    Spell As Long
    Uses As Long
End Type

Private Type BankRec
    Item(1 To MAX_BANK) As PlayerInvRec
End Type

Private Type PlayerRec
    ' General
    name As String
    Class As Long
    sprite As Long
    Level As Byte
    EXP As Long
    Access As Byte
    PK As Byte
    ' Vitals
    Vital(1 To Vitals.Vital_Count - 1) As Long
    MaxVital(1 To Vitals.Vital_Count - 1) As Long
    ' Stats
    Stat(1 To Stats.Stat_Count - 1) As Byte
    POINTS As Long
    ' Worn equipment
    Equipment(1 To Equipment.Equipment_Count - 1) As Long
    ' Position
    map As Long
    x As Byte
    y As Byte
    dir As Byte
    ' Variables
    Variable(1 To MAX_BYTE) As Long
    ' Client use only
    xOffset As Integer
    yOffset As Integer
    Moving As Byte
    Attacking As Byte
    AttackTimer As Long
    MapGetTimer As Long
    Step As Byte
    Anim As Long
    AnimTimer As Long
    usergroup As Long
End Type

Private Type EventCommandRec
    Type As Byte
    text As String
    Colour As Long
    channel As Byte
    TargetType As Byte
    target As Long
    x As Long
    y As Long
End Type

Public Type EventPageRec
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

Public Type EventRec
    name As String
    x As Long
    y As Long
    pageCount As Long
    EventPage() As EventPageRec
End Type

Private Type MapDataRec
    name As String
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
    tileSet As Long
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
    name As String * NAME_LENGTH
    Stat(1 To Stats.Stat_Count - 1) As Byte
    MaleSprite() As Long
    FemaleSprite() As Long
    ' For client use
    Vital(1 To Vitals.Vital_Count - 1) As Long
End Type

Public Type ItemRec
    name As String * NAME_LENGTH
    Desc As String * 255
    sound As String * NAME_LENGTH
    Pic As Long

    Type As Byte
    Data1 As Long
    Data2 As Long
    Data3 As Long
    ClassReq As Long
    AccessReq As Long
    LevelReq As Long
    Mastery As Byte
    Price As Long
    Add_Stat(1 To Stats.Stat_Count - 1) As Byte
    Rarity As Byte
    speed As Long
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
    playerName As String
    num As Long
    value As Long
    Frame As Byte
    x As Byte
    y As Byte
    bound As Boolean
End Type

Public Type NpcRec
    name As String * NAME_LENGTH
    AttackSay As String * 100
    sound As String * NAME_LENGTH
    sprite As Long
    SpawnSecs As Long
    Behaviour As Byte
    Range As Byte
    Stat(1 To Stats.Stat_Count - 1) As Byte
    HP As Long
    EXP As Long
    Animation As Long
    Damage As Long
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
    num As Long
    target As Long
    TargetType As Byte
    Vital(1 To Vitals.Vital_Count - 1) As Long
    map As Long
    x As Byte
    y As Byte
    dir As Byte
    ' Client use only
    xOffset As Long
    yOffset As Long
    Moving As Byte
    Attacking As Byte
    AttackTimer As Long
    Step As Byte
    Anim As Long
    AnimTimer As Long
End Type

Public Type TradeItemRec
    Item As Long
    ItemValue As Long
    CostItem As Long
    CostValue As Long
End Type

Private Type ShopRec
    name As String * NAME_LENGTH
    BuyRate As Long
    TradeItem(1 To MAX_TRADES) As TradeItemRec
End Type

Public Type SpellRec
    name As String * NAME_LENGTH
    Desc As String * 255
    sound As String * NAME_LENGTH

    Type As Byte
    MPCost As Long
    LevelReq As Long
    AccessReq As Long
    ClassReq As Long
    CastTime As Long
    CDTime As Long
    icon As Long
    map As Long
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
    ' doors... obviously
    DoorOpen As Byte
    DoorFrame As Byte
    DoorTimer As Long
    DoorAnimate As Byte ' 0 = nothing| 1 = opening | 2 = closing
    ' fading appear tiles
    isFading(1 To MapLayer.Layer_Count - 1) As Boolean
    fadeAlpha(1 To MapLayer.Layer_Count - 1) As Long
    FadeTimer(1 To MapLayer.Layer_Count - 1) As Long
    FadeDir(1 To MapLayer.Layer_Count - 1) As Byte
End Type

Public Type MapResourceRec
    x As Long
    y As Long
    ResourceState As Byte
End Type

Private Type ResourceRec
    name As String * NAME_LENGTH
    SuccessMessage As String * NAME_LENGTH
    EmptyMessage As String * NAME_LENGTH
    sound As String * NAME_LENGTH
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

Private Type ActionMsgRec
    message As String
    Created As Long

    Type As Long
    Color As Long
    Scroll As Long
    x As Long
    y As Long
    timer As Long
    alpha As Long
End Type

Private Type BloodRec
    sprite As Long
    timer As Long
    x As Long
    y As Long
End Type

Private Type AnimationRec
    name As String * NAME_LENGTH
    sound As String * NAME_LENGTH
    sprite(0 To 1) As Long
    Frames(0 To 1) As Long
    LoopCount(0 To 1) As Long
    looptime(0 To 1) As Long
End Type

Private Type AnimInstanceRec
    Animation As Long
    x As Long
    y As Long
    ' used for locking to players/npcs
    lockindex As Long
    LockType As Byte
    isCasting As Byte
    ' timing
    timer(0 To 1) As Long
    ' rendering check
    Used(0 To 1) As Boolean
    ' counting the loop
    LoopIndex(0 To 1) As Long
    FrameIndex(0 To 1) As Long
End Type

Public Type HotbarRec
    Slot As Long
    sType As Byte
End Type

Public Type PointRec
    x As Long
    y As Long
End Type

Public Type QuarterTileRec
    QuarterTile(1 To 4) As PointRec
    renderState As Byte
    srcX(1 To 4) As Long
    srcY(1 To 4) As Long
End Type

Public Type AutotileRec
    Layer(1 To MapLayer.Layer_Count - 1) As QuarterTileRec
End Type

Public Type ConvRec
    Conv As String
    rText(1 To 4) As String
    rTarget(1 To 4) As Long
    Event As Long
    Data1 As Long
    Data2 As Long
    Data3 As Long
End Type

Private Type ConvWrapperRec
    name As String * NAME_LENGTH
    chatCount As Long
    Conv() As ConvRec
End Type

Public Type ChatBubbleRec
    Msg As String
    Colour As Long
    target As Long
    TargetType As Byte
    timer As Long
    active As Boolean
End Type

Public Type TextColourRec
    text As String
    Colour As Long
End Type

Public Type GeomRec
    top As Long
    left As Long
    height As Long
    width As Long
End Type
