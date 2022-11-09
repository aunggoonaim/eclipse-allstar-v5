Attribute VB_Name = "modTypes"
Option Explicit

' Public data structures
Public Map As MapRec
Public Bank As BankRec
Public TempTile() As TempTileRec
Public Player(1 To MAX_PLAYERS) As PlayerRec
Public Class() As ClassRec
Public Item(1 To MAX_ITEMS) As ItemRec
Public NPC(1 To MAX_NPCS) As NpcRec
Public PET(1 To MAX_PETS) As PetsRec
Public MapItem(1 To MAX_MAP_ITEMS) As MapItemRec
Public MapNpc(1 To MAX_MAP_NPCS) As MapNpcRec
Public MapPet(1 To MAX_MAP_PETS) As MapPetRec
Public Shop(1 To MAX_SHOPS) As ShopRec
Public Spell(1 To MAX_SPELLS) As SpellRec
Public Resource(1 To MAX_RESOURCES) As ResourceRec
Public Animation(1 To MAX_ANIMATIONS) As AnimationRec
' Public Conv(1 To MAX_CONVS) As ConvWrapperRec «èÍ¹
Public Doors(1 To MAX_DOORS) As DoorRec
Public Switches(1 To MAX_SWITCHES) As String
Public Variables(1 To MAX_VARIABLES) As String

' client-side stuff
Public ActionMsg(1 To MAX_BYTE) As ActionMsgRec
Public Blood(1 To MAX_BYTE) As BloodRec
Public AnimInstance(1 To MAX_BYTE) As AnimInstanceRec
Public MenuButton(1 To MAX_MENUBUTTONS) As ButtonRec
Public MainButton(1 To MAX_MAINBUTTONS) As ButtonRec
Public party As PartyRec

' options
Public Options As OptionsRec

'Evilbunnie's DrawnChat system
Public Chat(1 To 20) As ChatRec

'Evilbunnie's DrawnChat system
Private Type ChatRec
    text As String
    Colour As Long
End Type

' Type recs
Private Type OptionsRec
    Game_Name As String
    SavePass As Byte
    Password As String * NAME_LENGTH
    Username As String * ACCOUNT_LENGTH
    IP As String
    Port As Long
    MenuMusic As String
    Music As Byte
    Sound As Byte
    Debug As Byte
    DefaultVolume As Byte
    Minimap As Byte
End Type

Public Type PartyRec
    Leader As Long
    Member(1 To MAX_PARTY_MEMBERS) As Long
    MemberCount As Long
    'Num As Long
End Type

Public Type PlayerInvRec
    Num As Long
    Value As Long
End Type

Private Type BankRec
    Item(1 To MAX_BANK) As PlayerInvRec
End Type

Private Type SpellAnim
    spellnum As Long
    Timer As Long
    FramePointer As Long
End Type

' projectiles
Public Type ProjectileRec
    TravelTime As Long
    Direction As Long
    X As Long
    Y As Long
    Pic As Long
    Range As Long
    Damage As Long
    Speed As Long
End Type

Public Type PetRec
    SpriteNum As Byte
    Name As String * 50
    Owner As Long
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
    ' General
    Name As String
    Class As Long
    Sprite As Long
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
    Map As Long
    X As Byte
    Y As Byte
    Dir As Byte
    
    ' Client use only
    xOffset As Integer
    yOffset As Integer
    Moving As Byte
    Attacking As Byte
    AttackTimer As Long
    MapGetTimer As Long
    Step As Byte
    ' quest
    PlayerQuest(1 To MAX_QUESTS) As PlayerQuestRec
    ' guild
    GuildName As String
    GuildHome As Byte
    ' pet
    PET As PetRec
    ' doors
    PlayerDoors(1 To MAX_DOORS) As DoorRec
    
    ' projectiles
    ProjecTile(1 To MAX_PLAYER_PROJECTILES) As ProjectileRec
    
    EventTimer As Long
    
    'Message
    Message As String
    
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

Private Type TileDataRec
    X As Long
    Y As Long
    Tileset As Long
End Type

Public Type ConditionalBranchRec
    Condition As Long
    Data1 As Long
    Data2 As Long
    Data3 As Long
    CommandList As Long
    ElseCommandList As Long
End Type

Public Type MoveRouteRec
    Index As Long
    Data1 As Long
    Data2 As Long
    Data3 As Long
    Data4 As Long
    Data5 As Long
    Data6 As Long
End Type

Public Type EventCommandRec
    Index As Long
    Text1 As String
    Text2 As String
    Text3 As String
    Text4 As String
    Text5 As String
    Data1 As Long
    Data2 As Long
    Data3 As Long
    Data4 As Long
    Data5 As Long
    Data6 As Long
    ConditionalBranch As ConditionalBranchRec
    MoveRouteCount As Long
    MoveRoute() As MoveRouteRec
End Type

Public Type CommandListRec
    CommandCount As Long
    ParentList As Long
    Commands() As EventCommandRec
End Type

Public Type EventPageRec
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
    WalkAnim As Byte
    DirFix As Byte
    WalkThrough As Byte
    ShowName As Byte
    
    'Trigger for the event
    Trigger As Byte
    
    'Commands for the event
    CommandListCount As Long
    CommandList() As CommandListRec
    
    Position As Byte
    
    'Client Needed Only
    X As Long
    Y As Long
End Type

Public Type EventRec
    Name As String
    Global As Long
    pageCount As Long
    Pages() As EventPageRec
    X As Long
    Y As Long
End Type

Public Type TileRec
    Layer(1 To MapLayer.Layer_Count - 1) As TileDataRec
    Type As Byte
    Data1 As Long
    Data2 As Long
    Data3 As Long
    DirBlock As Byte
End Type

Private Type MapEventRec
    Name As String
    Dir As Long
    X As Long
    Y As Long
    GraphicType As Long
    GraphicX As Long
    GraphicY As Long
    GraphicX2 As Long
    GraphicY2 As Long
    GraphicNum As Long
    Moving As Long
    MovementSpeed As Long
    Position As Long
    xOffset As Long
    yOffset As Long
    Step As Long
    Visible As Long
    WalkAnim As Long
    DirFix As Long
    ShowDir As Long
    WalkThrough As Long
    ShowName As Long
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
    
    maxX As Byte
    maxY As Byte
    
    Tile() As TileRec
    NPC(1 To MAX_MAP_NPCS) As Long
    
    EventCount As Long
    Events() As EventRec
    
    'Client Side Only -- Temporary
    CurrentEvents As Long
    MapEvents() As MapEventRec
End Type

Private Type ClassRec
    Name As String * NAME_LENGTH
    Stat(1 To Stats.Stat_Count - 1) As Byte
    MaleSprite() As Long
    FemaleSprite() As Long
    ' For client use
    Vital(1 To Vitals.Vital_Count - 1) As Long
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
    Price As Long
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
    ProjecTile As ProjectileRec
    ' tools
    Toolpower As Long
    ' crafting
    Tool As Long
    ToolReq As Long
    
    isDagger As Boolean
    isTwoHanded As Boolean
    
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
    playerName As String
    Num As Long
    Value As Long
    Frame As Byte
    X As Byte
    Y As Byte
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
    Stat(1 To Stats.Stat_Count - 1) As Byte
    HP As Long
    EXP As Long
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
    Stat(1 To Stats.Stat_Count - 1) As Byte
    HP As Long
    EXP As Long
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
    Num As Long
    target As Long
    TargetType As Byte
    Vital(1 To Vitals.Vital_Count - 1) As Long
    Map As Long
    X As Byte
    Y As Byte
    Dir As Byte
    ' Client use only
    xOffset As Long
    yOffset As Long
    Moving As Byte
    Attacking As Byte
    AttackTimer As Long
    Step As Byte
    c_lastDir As Byte
    c_inChatWith As Long
    ' pets
    'Pet Data
    IsPet As Byte
    PetData As PetRec
    ' Npc spells
    SpellTimer(1 To MAX_NPC_SPELLS) As Long
    Heals As Integer
End Type

Private Type MapPetRec
    Num As Long
    target As Long
    TargetType As Byte
    Vital(1 To Vitals.Vital_Count - 1) As Long
    Map As Long
    X As Byte
    Y As Byte
    Dir As Byte
    ' Client use only
    xOffset As Long
    yOffset As Long
    Moving As Byte
    Attacking As Byte
    AttackTimer As Long
    Step As Byte
    c_lastDir As Byte
    c_inChatWith As Long
    ' pets
    'Pet Data
    IsPet As Byte
    PetData As PetRec
    ' Npc spells
    SpellTimer(1 To MAX_NPC_SPELLS) As Long
    Heals As Integer
End Type

Private Type TradeItemRec
    Item As Long
    ItemValue As Long
    CostItem As Long
    CostValue As Long
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
    X As Long
    Y As Long
    Dir As Byte
    Vital As Long
    Duration As Long
    Interval As Long
    Range As Byte
    IsAoE As Boolean
    AoE As Long
    CastAnim As Long
    SpellAnim As Long
    StunDuration As Long
    ProjecTile As ProjectileRec
    
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
    DoorOpen As Byte
    DoorFrame As Byte
    DoorTimer As Long
    DoorAnimate As Byte ' 0 = nothing| 1 = opening | 2 = closing
End Type

Public Type MapResourceRec
    X As Long
    Y As Long
    ResourceState As Byte
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

Private Type ActionMsgRec
    Message As String
    Created As Long
    Type As Long
    Color As Long
    Scroll As Long
    X As Long
    Y As Long
    Timer As Long
End Type

Private Type BloodRec
    Sprite As Long
    Timer As Long
    X As Long
    Y As Long
End Type

Private Type AnimationRec
    Name As String * NAME_LENGTH
    Sound As String * NAME_LENGTH
    
    Sprite(0 To 1) As Long
    Frames(0 To 1) As Long
    LoopCount(0 To 1) As Long
    looptime(0 To 1) As Long
    
    Alpha As Byte
End Type

Private Type AnimInstanceRec
    Animation As Long
    X As Long
    Y As Long
    ' used for locking to players/npcs
    lockindex As Long
    LockType As Byte
    ' timing
    Timer(0 To 1) As Long
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

Public Type ButtonRec
    FileName As String
    state As Byte
End Type

Type DropRec
    X As Long
    Y As Long
    ySpeed As Long
    xSpeed As Long
    Init As Boolean
End Type

Public Type EventListRec
    CommandList As Long
    CommandNum As Long
End Type
