Attribute VB_Name = "modTypes"
Option Explicit

' Public data structures
Public Map(1 To MAX_MAPS) As MapRec
Public TempEventMap(1 To MAX_MAPS) As GlobalEventsRec
Public MapCache(1 To MAX_MAPS) As Cache
Public TempTile(1 To MAX_MAPS) As TempTileRec
Public PlayersOnMap(1 To MAX_MAPS) As Long
Public ResourceCache(1 To MAX_MAPS) As ResourceCacheRec
Public Player(1 To MAX_PLAYERS) As PlayerRec
Public Bank(1 To MAX_PLAYERS) As BankRec
Public TempPlayer(1 To MAX_PLAYERS) As TempPlayerRec
Public Class() As ClassRec
Public Item(1 To MAX_ITEMS) As ItemRec
Public NPC(1 To MAX_NPCS) As NpcRec
Public Pet(1 To MAX_PETS) As PetsRec
Public MapItem(1 To MAX_MAPS, 1 To MAX_MAP_ITEMS) As MapItemRec
Public MapNpc(1 To MAX_MAPS) As MapDataRec
Public MapPet(1 To MAX_MAPS) As MapDataPetRec
Public Shop(1 To MAX_SHOPS) As ShopRec
Public Spell(1 To MAX_SPELLS) As SpellRec
Public Resource(1 To MAX_RESOURCES) As ResourceRec
Public Animation(1 To MAX_ANIMATIONS) As AnimationRec
Public Party(1 To MAX_PARTYS) As PartyRec
Public Options As OptionsRec
' Public Conv(1 To MAX_CONVS) As ConvWrapperRec
Public Doors(1 To MAX_DOORS) As DoorRec
Public Switches(1 To MAX_SWITCHES) As String
Public Variables(1 To MAX_VARIABLES) As String

Private Type MoveRouteRec
    index As Long
    Data1 As Long
    Data2 As Long
    Data3 As Long
    data4 As Long
    data5 As Long
    data6 As Long
End Type

Private Type GlobalEventRec
    x As Long
    y As Long
    Dir As Long
    active As Long
    
    WalkingAnim As Long
    FixedDir As Long
    WalkThrough As Long
    Position As Long
    
    GraphicType As Long
    GraphicNum As Long
    GraphicX As Long
    GraphicX2 As Long
    GraphicY As Long
    GraphicY2 As Long
    
    'Server Only Options
    MoveType As Long
    MoveSpeed As Long
    MoveFreq As Long
    MoveRouteCount As Long
    MoveRoute() As MoveRouteRec
    MoveRouteStep As Long
    
    RepeatMoveRoute As Long
    IgnoreIfCannotMove As Long
    
    MoveTimer As Long
End Type

Public Type GlobalEventsRec
    EventCount As Long
    Events() As GlobalEventRec
End Type

Private Type OptionsRec
    Game_Name As String
    MOTD As String
    Port As Long
    Website As String
End Type

Public Type PartyRec
    Leader As Long
    Member(1 To MAX_PARTY_MEMBERS) As Long
    MemberCount As Long
End Type

Public Type PlayerInvRec
    num As Long
    Value As Long
End Type

Private Type Cache
    Data() As Byte
End Type

Private Type BankRec
    Item(1 To MAX_BANK) As PlayerInvRec
End Type

Public Type HotbarRec
    Slot As Long
    sType As Byte
End Type

' project tiles
Public Type ProjectileRec
    TravelTime As Long
    Direction As Long
    x As Long
    y As Long
    Pic As Long
    Range As Long
    Damage As Long
    Speed As Long
End Type

Public Type PetRec
    SpriteNum As Byte
    Name As String * 50
    Owner As Long
    CNum As Long
End Type

Public Type DoorRec
    Name As String * NAME_LENGTH
    DoorType As Long
    
    WarpMap As Long
    WarpX As Long
    WarpY As Long
    
    UnlockType As Long
    key As Long
    Switch As Long
    
    state As Long
End Type

Private Type PlayerRec
    ' Account
    Login As String * ACCOUNT_LENGTH
    Password As String * NAME_LENGTH
    
    ' General
    Name As String * ACCOUNT_LENGTH
    Sex As Byte
    Class As Long
    Sprite As Long
    Level As Byte
    exp As Long
    Access As Byte
    PK As Byte
    GuildFileId As Long
    GuildMemberId As Long
    GuildHome As Byte
    ' Vitals
    Vital(1 To Vitals.Vital_Count - 1) As Long
    
    ' Stats
    stat(1 To Stats.Stat_Count - 1) As Byte
    POINTS As Long
    
    ' Worn equipment
    Equipment(1 To Equipment.Equipment_Count - 1) As Long
    
    ' Inventory
    Inv(1 To MAX_INV) As PlayerInvRec
    Spell(1 To MAX_PLAYER_SPELLS) As Long
    
    ' Hotbar
    Hotbar(1 To MAX_HOTBAR) As HotbarRec
    
    ' Position
    Map As Long
    x As Byte
    y As Byte
    Dir As Byte
    ' quests
    PlayerQuest(1 To MAX_QUESTS) As PlayerQuestRec
    ' pets
    Pet As PetRec
    'doors
    PlayerDoors(1 To MAX_DOORS) As DoorRec
    
    Switches(0 To MAX_SWITCHES) As Byte
    Variables(0 To MAX_VARIABLES) As Long
    
    'Message
    message As String
    
    WieldDagger As Byte
    skillLV(1 To MAX_PLAYER_SPELLS) As Byte
    skillEXP(1 To MAX_PLAYER_SPELLS) As Long
    Killer As Long
    BuffStatus(1 To MAX_BUFF) As Byte
    BuffTime(1 To MAX_BUFF) As Byte
    ElementATK As Byte
    ElementDEF As Byte
    
    OnDeath As Boolean
End Type

Public Type SpellBufferRec
    Spell As Long
    Timer As Long
    Target As Long
    tType As Byte
End Type

Public Type DoTRec
    Used As Boolean
    Spell As Long
    Timer As Long
    Caster As Long
    StartTime As Long
End Type

Public Type ConditionalBranchRec
    Condition As Long
    Data1 As Long
    Data2 As Long
    Data3 As Long
    CommandList As Long
    ElseCommandList As Long
End Type

Private Type EventCommandRec
    index As Byte
    Text1 As String
    Text2 As String
    Text3 As String
    Text4 As String
    Text5 As String
    Data1 As Long
    Data2 As Long
    Data3 As Long
    data4 As Long
    data5 As Long
    data6 As Long
    ConditionalBranch As ConditionalBranchRec
    MoveRouteCount As Long
    MoveRoute() As MoveRouteRec
End Type

Private Type CommandListRec
    CommandCount As Long
    ParentList As Long
    Commands() As EventCommandRec
End Type

Private Type EventPageRec
    'These are condition variables that decide if the event even appears to the player.
    chkVariable As Long
    VariableIndex As Long
    VariableCondition As Long
    VariableCompare As Long
    
    chkSwitch As Long
    SwitchIndex As Long
    SwitchCompare As Long
    
    chkHasItem As Long
    HasItemIndex As Long
    
    chkSelfSwitch As Long
    SelfSwitchIndex As Long
    SelfSwitchCompare As Long
    'End Conditions
    
    'Handles the Event Sprite
    GraphicType As Byte
    Graphic As Long
    GraphicX As Long
    GraphicY As Long
    GraphicX2 As Long
    GraphicY2 As Long
    
    'Handles Movement - Move Routes to come soon.
    MoveType As Byte
    MoveSpeed As Byte
    MoveFreq As Byte
    MoveRouteCount As Long
    MoveRoute() As MoveRouteRec
    IgnoreMoveRoute As Long
    RepeatMoveRoute As Long
    
    'Guidelines for the event
    WalkAnim As Long
    DirFix As Long
    WalkThrough As Long
    ShowName As Long
    
    'Trigger for the event
    Trigger As Byte
    
    'Commands for the event
    CommandListCount As Long
    CommandList() As CommandListRec
    
    Position As Byte
    
    'For EventMap
    x As Long
    y As Long
End Type

Private Type EventRec
    Name As String
    Global As Byte
    PageCount As Long
    Pages() As EventPageRec
    x As Long
    y As Long
    'Self Switches re-set on restart.
    SelfSwitches(0 To 4) As Long
End Type

Public Type GlobalMapEvents
    eventID As Long
    pageID As Long
    x As Long
    y As Long
End Type

Private Type MapEventRec
    Dir As Long
    x As Long
    y As Long
    
    WalkingAnim As Long
    FixedDir As Long
    WalkThrough As Long
    
    GraphicType As Long
    GraphicX As Long
    GraphicY As Long
    GraphicX2 As Long
    GraphicY2 As Long
    GraphicNum As Long
    
    movementspeed As Long
    Position As Long
    Visible As Long
    eventID As Long
    pageID As Long
    
    'Server Only Options
    MoveType As Long
    MoveSpeed As Long
    MoveFreq As Long
    MoveRouteCount As Long
    MoveRoute() As MoveRouteRec
    MoveRouteStep As Long
    
    RepeatMoveRoute As Long
    IgnoreIfCannotMove As Long
    
    MoveTimer As Long
    SelfSwitches(0 To 4) As Long
End Type

Private Type EventMapRec
    CurrentEvents As Long
    EventPages() As MapEventRec
End Type

Private Type EventProcessingRec
    CurList As Long
    CurSlot As Long
    eventID As Long
    pageID As Long
    WaitingForResponse As Long
    ActionTimer As Long
    ListLeftOff() As Long
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
    Target As Long
    GettingMap As Byte
    SpellCD(1 To MAX_PLAYER_SPELLS) As Long
    InShop As Long
    StunTimer As Long
    StunDuration As Long
    InBank As Boolean
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
    ' projectiles
    Projectile(1 To MAX_PLAYER_PROJECTILES) As ProjectileRec
    ' guilds
    tmpGuildSlot As Long
    tmpGuildInviteSlot As Long
    tmpGuildInviteTimer As Long
    tmpGuildInviteId As Long
    ' pets
    TempPetSlot As Byte
    
    ' Oldmap
    OldMap As Long
   havePet As Boolean
   
    EventMap As EventMapRec
    EventProcessingCount As Long
    EventProcessing() As EventProcessingRec
End Type

Private Type TileDataRec
    x As Long
    y As Long
    Tileset As Long
End Type

Private Type TileRec
    Layer(1 To MapLayer.Layer_Count - 1) As TileDataRec
    Type As Byte
    Data1 As Long
    Data2 As Long
    Data3 As Long
    DirBlock As Byte
End Type

Private Type MapRec
    Name As String * NAME_LENGTH
    Music As String * NAME_LENGTH
    
    Weather As Long
    
    Revision As Long
    Moral As Byte
    
    Up As Long
    Down As Long
    Left As Long
    Right As Long
    
    BootMap As Long
    BootX As Byte
    BootY As Byte
    
    MaxX As Byte
    MaxY As Byte
    
    Tile() As TileRec
    NPC(1 To MAX_MAP_NPCS) As Long
    
    EventCount As Long
    Events() As EventRec
End Type

Private Type ClassRec
    Name As String * NAME_LENGTH
    stat(1 To Stats.Stat_Count - 1) As Byte
    MaleSprite() As Long
    FemaleSprite() As Long
    
    startItemCount As Long
    StartItem() As Long
    StartValue() As Long
    
    startSpellCount As Long
    StartSpell() As Long
    Locked As Byte
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
    SpeedLow As Long
    Handed As Long
    BindType As Byte
    Stat_Req(1 To Stats.Stat_Count - 1) As Byte
    Animation As Long
    Paperdoll As Long
    
    AddHP As Long
    AddMP As Long
    AddEXP As Long
    CastSpell As Long
    instaCast As Byte
    
    ' projectile
    Projectile As ProjectileRec
    ' tools
    Toolpower As Long
    ' crafting
    Tool As Long
    ToolReq As Long
    
    isDagger As Boolean
    isTwohanded As Boolean

    Daggerpdoll As Long
    
    ' New Funtion by Allstar
    Kick As Byte
    MATK As Long
    NDEF As Long
    CritRate As Long
    DelayDown As Double
    CritATK As Long
    ReUse As Long
    DelayUse As Long
    
    HP As Long
    MP As Long
    
    Add1 As Long
    Sub1 As Long
    HPCase As Long
    Add2 As Long
    Sub2 As Long
    MPCase As Long
    Vampire As Long
    Dodge As Long
    DropOnDeath As Byte
    RegenHp As Integer
    RegenMp As Integer
    Per1 As Byte
    Per2 As Byte

    ClassR1 As Byte
    ClassR2 As Byte
    ClassR3 As Byte
    ClassR4 As Byte
    ClassR5 As Byte
    ClassR6 As Byte
    ClassR7 As Byte
    ClassR8 As Byte
    ClassR9 As Byte
    ClassR10 As Byte
    ClassR11 As Byte
    
    LHand As Byte
    DmgLow As Byte
    MagicLow As Byte
    DmgHigh As Long
    MagicHigh As Long
    Reflect As Byte
    DmgReflect As Long
    AbsorbMagic As Byte
    
    ' buff item mode
    Buff(1 To MAX_BUFF) As Byte
    BuffTime(1 To MAX_BUFF) As Byte
End Type

Private Type MapItemRec
    num As Long
    Value As Long
    x As Byte
    y As Byte
    ' ownership + despawn
    playerName As String
    playerTimer As Long
    canDespawn As Boolean
    despawnTimer As Long
End Type

Private Type NpcRec
    Name As String * NAME_LENGTH
    AttackSay As String * 100
    Sound As String * NAME_LENGTH
    
    Sprite As Long
    SpawnSecs As Long
    Behaviour As Byte
    Range As Byte
    DropChance(1 To MAX_NPC_DROPS) As Double
    DropItem(1 To MAX_NPC_DROPS) As Integer
    DropItemValue(1 To MAX_NPC_DROPS) As Integer
    stat(1 To Stats.Stat_Count - 1) As Byte
    HP As Long
    exp As Long
    EXP_max As Long
    Animation As Long
    Damage As Long
    Level As Long
    Quest As Byte
    QuestNum As Long
    ' Npc Spells
    Spell(1 To MAX_NPC_SPELLS) As Long
    ' bosses
    BossNum As Integer
    
    AttackSpeed As Integer
    CritRate As Integer
    CritChange As Integer
    
    ' New system by Allstar
    Def As Long
    Dodge As Byte
    Block As Byte
    RegenHp As Long
    RegenMp As Long
    MATK As Long
    ReflectDmg As Long
    AbsorbMagic As Byte
    
    ElementATK As Byte
    ElementDEF As Byte
    
    Alpha As Byte
End Type

Private Type PetsRec
    Name As String * NAME_LENGTH
    AttackSay As String * 100
    Sound As String * NAME_LENGTH
    
    Sprite As Long
    SpawnSecs As Long
    Behaviour As Byte
    Range As Byte
    DropChance(1 To MAX_NPC_DROPS) As Double
    DropItem(1 To MAX_NPC_DROPS) As Integer
    DropItemValue(1 To MAX_NPC_DROPS) As Integer
    stat(1 To Stats.Stat_Count - 1) As Byte
    HP As Long
    exp As Long
    EXP_max As Long
    Animation As Long
    Damage As Long
    Level As Long
    Quest As Byte
    QuestNum As Long
    ' Npc Spells
    Spell(1 To MAX_NPC_SPELLS) As Long
    ' bosses
    BossNum As Integer
    
    AttackSpeed As Integer
    CritRate As Integer
    CritChange As Integer
    
    ' New system by Allstar
    Def As Long
    Dodge As Byte
    Block As Byte
    RegenHp As Long
    RegenMp As Long
    MATK As Long
    ReflectDmg As Long
    AbsorbMagic As Byte
    
    ElementATK As Byte
    ElementDEF As Byte
End Type

Private Type MapNpcRec
    num As Long
    Target As Long
    targetType As Byte
    Vital(1 To Vitals.Vital_Count - 1) As Long
    x As Byte
    y As Byte
    Dir As Byte
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
    c_lastDir As Byte
    c_inChatWith As Long
    'Pet Data
    IsPet As Byte
    PetData As PetRec
    ' Npc spells
    SpellTimer(1 To MAX_NPC_SPELLS) As Long
    Heals As Integer
    GetDamage As Long
End Type

Private Type MapPetRec
    num As Long
    Target As Long
    targetType As Byte
    Vital(1 To Vitals.Vital_Count - 1) As Long
    x As Byte
    y As Byte
    Dir As Byte
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
    c_lastDir As Byte
    c_inChatWith As Long
    'Pet Data
    IsPet As Byte
    PetData As PetRec
    ' Npc spells
    SpellTimer(1 To MAX_NPC_SPELLS) As Long
    Heals As Integer
    GetDamage As Long
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
    MPCost As Long
    HPCost As Long
    LevelReq As Long
    AccessReq As Long
    ClassReq As Long
    CastTime As Long
    CDTime As Long
    Icon As Long
    Map As Long
    x As Long
    y As Long
    Dir As Byte
    Vital As Long
    Duration As Long
    Interval As Long
    Range As Byte
    IsAoE As Boolean
    AoE As Long
    CastAnim As Long
    spellAnim As Long
    StunDuration As Long
    Projectile As ProjectileRec
    
    ' New Funtion by Allstar
    PhysicalDmg As Long
    MagicDmg As Long
    ATKPer As Long
    MagicPer As Long
    Passive As Long
    PATK As Long
    PDEF As Long
    PerSkill As Long
    CanMove As Long
    
    CanCancle As Long
    
    S1 As Long ' Value
    S2 As Long ' ATK
    S3 As Long ' MATK
    S4 As Long ' % Passive
    
    Element As Byte
End Type

Private Type TempTileRec
    DoorOpen() As Byte
    DoorTimer As Long
End Type

Private Type MapDataRec
    NPC() As MapNpcRec
End Type

Private Type MapDataPetRec
    Pet() As MapPetRec
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
    ToolpowerReq As Long
    SuccessRate As Byte
End Type

Private Type AnimationRec
    Name As String * NAME_LENGTH
    Sound As String * NAME_LENGTH
    
    Sprite(0 To 1) As Long
    Frames(0 To 1) As Long
    LoopCount(0 To 1) As Long
    LoopTime(0 To 1) As Long
    
    Alpha As Byte
End Type

