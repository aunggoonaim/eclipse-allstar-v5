Attribute VB_Name = "modConstants"
    Option Explicit

' API Declares
Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByRef Msg() As Byte, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long

' animated buttons
Public Const MAX_MENUBUTTONS As Long = 4
Public Const MAX_MAINBUTTONS As Long = 9
Public Const MENUBUTTON_PATH As String = "\Data Files\graphics\gui\menu\buttons\"
Public Const MAINBUTTON_PATH As String = "\Data Files\graphics\gui\main\buttons\"

' Hotbar
Public Const HotbarTop As Byte = 2
Public Const HotbarLeft As Byte = 2
Public Const HotbarOffsetX As Byte = 8

' Inventory constants
Public Const InvTop As Byte = 24
Public Const InvLeft As Byte = 12
Public Const InvOffsetY As Byte = 3
Public Const InvOffsetX As Byte = 3
Public Const InvColumns As Byte = 5

' Bank constants
Public Const BankTop As Byte = 38
Public Const BankLeft As Byte = 42
Public Const BankOffsetY As Byte = 3
Public Const BankOffsetX As Byte = 4
Public Const BankColumns As Byte = 11

' spells constants
Public Const SpellTop As Byte = 24
Public Const SpellLeft As Byte = 12
Public Const SpellOffsetY As Byte = 3
Public Const SpellOffsetX As Byte = 3
Public Const SpellColumns As Byte = 5

' shop constants
Public Const ShopTop As Byte = 6
Public Const ShopLeft As Byte = 8
Public Const ShopOffsetY As Byte = 2
Public Const ShopOffsetX As Byte = 4
Public Const ShopColumns As Byte = 5

' Character consts
Public Const EqTop As Byte = 224
Public Const EqLeft As Byte = 18
Public Const EqOffsetX As Byte = 10
Public Const EqColumns As Byte = 4

' values
Public Const MAX_BYTE As Byte = 255
Public Const MAX_INTEGER As Integer = 32767
Public Const MAX_LONG As Long = 2147483647

' path constants
Public Const SOUND_PATH As String = "\Data Files\sound\"
Public Const MUSIC_PATH As String = "\Data Files\music\"

' Font variables
Public Const FONT_NAME As String = "MS Sans Serif"
Public Const FONT_SIZE As Byte = 10

' Log Path and variables
Public Const LOG_DEBUG As String = "debug.txt"
Public Const LOG_PATH As String = "\Data Files\logs\"

' Map Path and variables
Public Const MAP_PATH As String = "\Data Files\maps\"
Public Const MAP_EXT As String = ".map"

' Gfx Path and variables
Public Const GFX_PATH As String = "\Data Files\graphics\"
' Public Const GFX_EXT As String = ".bmp"
Public Const GFX_EXT As String = ".png"

' Key constants
Public Const VK_UP As Long = &H26
Public Const VK_DOWN As Long = &H28
Public Const VK_LEFT As Long = &H25
Public Const VK_RIGHT As Long = &H27
Public Const VK_SHIFT As Long = &H10
Public Const VK_RETURN As Long = &HD
Public Const VK_CONTROL As Long = &H11

' Menu states
Public Const MENU_STATE_NEWACCOUNT As Byte = 0
Public Const MENU_STATE_DELACCOUNT As Byte = 1
Public Const MENU_STATE_LOGIN As Byte = 2
Public Const MENU_STATE_GETCHARS As Byte = 3
Public Const MENU_STATE_NEWCHAR As Byte = 4
Public Const MENU_STATE_ADDCHAR As Byte = 5
Public Const MENU_STATE_DELCHAR As Byte = 6
Public Const MENU_STATE_USECHAR As Byte = 7
Public Const MENU_STATE_INIT As Byte = 8

' Speed moving vars
Public Const WALK_SPEED As Byte = 4
Public Const RUN_SPEED As Byte = 4

' Tile size constants
Public Const PIC_X As Long = 32
Public Const PIC_Y As Long = 32

' Sprite, item, spell size constants
Public Const SIZE_X As Byte = 32
Public Const SIZE_Y As Byte = 32

' ********************************************************
' * The values below must match with the server's values *
' ********************************************************

' Constants ทั่วไป
Public Const MAX_PLAYERS As Long = 200
Public Const MAX_ITEMS As Long = 255
Public Const MAX_NPCS As Long = 255
Public Const MAX_PETS As Long = 100
Public Const MAX_ANIMATIONS As Long = 255
Public Const MAX_INV As Long = 35
Public Const MAX_MAP_ITEMS As Long = 255
Public Const MAX_MAP_NPCS As Long = 30
Public Const MAX_MAP_PETS As Long = 50
Public Const MAX_SHOPS As Long = 50
Public Const MAX_PLAYER_SPELLS As Long = 35
Public Const MAX_SPELLS As Long = 255
Public Const MAX_TRADES As Long = 30
Public Const MAX_RESOURCES As Long = 100
Public Const MAX_LEVELS As Long = 60
Public Const MAX_BANK As Long = 255
Public Const MAX_HOTBAR As Long = 12
Public Const MAX_PARTYS As Long = 255
Public Const MAX_PARTY_MEMBERS As Long = 4
Public Const MAX_SWITCHES As Long = 200
Public Const MAX_VARIABLES As Long = 200
Public Const MAX_PLAYER_PROJECTILES As Long = 20 ' projectiles
Public Const MAX_NPC_DROPS As Long = 10
Public Const MAX_CLASS As Long = 11

' NPC Spells
Public Const MAX_NPC_SPELLS As Long = 5
' Doors
Public Const MAX_DOORS As Long = 255
' Skill level
Public Const MAX_SKILL_LEVEL As Byte = 10
' Buff
Public Const MAX_BUFF As Byte = 10

' Website
Public Const GAME_WEBSITE As String = "http://monsterwar.zz.mu/"

' text color constants
Public Const Black As Byte = 0
Public Const Blue As Byte = 1
Public Const Green As Byte = 2
Public Const Cyan As Byte = 3
Public Const Red As Byte = 4
Public Const Magenta As Byte = 5
Public Const Brown As Byte = 6
Public Const Grey As Byte = 7
Public Const DarkGrey As Byte = 8
Public Const BrightBlue As Byte = 9
Public Const BrightGreen As Byte = 10
Public Const BrightCyan As Byte = 11
Public Const BrightRed As Byte = 12
Public Const Pink As Byte = 13
Public Const Yellow As Byte = 14
Public Const White As Byte = 15
Public Const SayColor As Byte = White
Public Const GlobalColor As Byte = BrightBlue
Public Const BroadcastColor As Byte = BrightCyan
Public Const TellColor As Byte = Pink
Public Const EmoteColor As Byte = BrightCyan
Public Const AdminColor As Byte = BrightCyan
Public Const HelpColor As Byte = BrightBlue
Public Const WhoColor As Byte = Yellow
Public Const JoinLeftColor As Byte = White
Public Const NpcColor As Byte = Yellow
Public Const AlertColor As Byte = Red
Public Const NewMapColor As Byte = BrightBlue

' Boolean constants
Public Const NO As Byte = 0
Public Const YES As Byte = 1

' String constants
Public Const NAME_LENGTH As Byte = 20
Public Const ACCOUNT_LENGTH As Byte = 20

' Sex constants
Public Const SEX_MALE As Byte = 0
Public Const SEX_FEMALE As Byte = 1

' Map constants
Public Const MAX_MAPS As Byte = 200
Public Const MAX_MAPX As Byte = 22
Public Const MAX_MAPY As Byte = 12
Public Const MAP_MORAL_NONE As Byte = 0
Public Const MAP_MORAL_SAFE As Byte = 1
Public Const MAP_MORAL_PETARENA As Byte = 2
Public Const MAP_MORAL_PARTY_MAP As Byte = 3

' Tile consants
Public Const TILE_TYPE_WALKABLE As Byte = 0
Public Const TILE_TYPE_BLOCKED As Byte = 1
Public Const TILE_TYPE_WARP As Byte = 2
Public Const TILE_TYPE_ITEM As Byte = 3
Public Const TILE_TYPE_NPCAVOID As Byte = 4
Public Const TILE_TYPE_KEY As Byte = 5
Public Const TILE_TYPE_KEYOPEN As Byte = 6
Public Const TILE_TYPE_RESOURCE As Byte = 7
Public Const TILE_TYPE_DOOR As Byte = 8
Public Const TILE_TYPE_NPCSPAWN As Byte = 9
Public Const TILE_TYPE_SHOP As Byte = 10
Public Const TILE_TYPE_BANK As Byte = 11
Public Const TILE_TYPE_HEAL As Byte = 12
Public Const TILE_TYPE_TRAP As Byte = 13
Public Const TILE_TYPE_SLIDE As Byte = 14
Public Const TILE_TYPE_CHEST As Byte = 15
Public Const TILE_TYPE_SPRITE As Byte = 16
Public Const TILE_TYPE_ANIMATION As Byte = 17
Public Const TILE_TYPE_CHECKPOINT As Byte = 18
Public Const TILE_TYPE_CRAFT As Byte = 19
Public Const TILE_TYPE_ONCLICK As Byte = 20


' Item constants
Public Const ITEM_TYPE_NONE As Byte = 0
Public Const ITEM_TYPE_WEAPON As Byte = 1
Public Const ITEM_TYPE_ARMOR As Byte = 2
Public Const ITEM_TYPE_HELMET As Byte = 3
Public Const ITEM_TYPE_SHIELD As Byte = 4
Public Const ITEM_TYPE_CONSUME As Byte = 5
Public Const ITEM_TYPE_KEY As Byte = 6
Public Const ITEM_TYPE_CURRENCY As Byte = 7
Public Const ITEM_TYPE_SPELL As Byte = 8
Public Const ITEM_TYPE_SUMMON As Byte = 9
Public Const ITEM_TYPE_RECIPE As Byte = 10
Public Const ITEM_TYPE_SCRIPT As Byte = 11

' Direction constants
Public Const DIR_UP As Byte = 0
Public Const DIR_DOWN As Byte = 1
Public Const DIR_LEFT As Byte = 2
Public Const DIR_RIGHT As Byte = 3

' Constants for player movement: Tiles per Second
Public Const MOVING_WALKING As Byte = 1
Public Const MOVING_RUNNING As Byte = 2

' Admin constants
Public Const ADMIN_MONITOR As Byte = 1
Public Const ADMIN_MAPPER As Byte = 2
Public Const ADMIN_DEVELOPER As Byte = 3
Public Const ADMIN_CREATOR As Byte = 4

' NPC constants
Public Const NPC_BEHAVIOUR_ATTACKONSIGHT As Byte = 0
Public Const NPC_BEHAVIOUR_ATTACKWHENATTACKED As Byte = 1
Public Const NPC_BEHAVIOUR_FRIENDLY As Byte = 2
Public Const NPC_BEHAVIOUR_SHOPKEEPER As Byte = 3
Public Const NPC_BEHAVIOUR_GUARD As Byte = 4
Public Const NPC_BEHAVIOUR_BOSS As Byte = 5

' Spell constants
Public Const SPELL_TYPE_DAMAGEHP As Byte = 0
Public Const SPELL_TYPE_DAMAGEMP As Byte = 1
Public Const SPELL_TYPE_HEALHP As Byte = 2
Public Const SPELL_TYPE_HEALMP As Byte = 3
Public Const SPELL_TYPE_WARP As Byte = 4
Public Const SPELL_TYPE_PET As Byte = 5
Public Const SPELL_TYPE_PROJECTILE As Byte = 6 ' or next number on list
Public Const SPELL_TYPE_SCRIPT As Byte = 7

' Game editor constants
Public Const EDITOR_ITEM As Byte = 1
Public Const EDITOR_NPC As Byte = 2
Public Const EDITOR_SPELL As Byte = 3
Public Const EDITOR_SHOP As Byte = 4
Public Const EDITOR_RESOURCE As Byte = 5
Public Const EDITOR_ANIMATION As Byte = 6
Public Const EDITOR_DOORS As Byte = 7

' Target type constants
Public Const TARGET_TYPE_NONE As Byte = 0
Public Const TARGET_TYPE_PLAYER As Byte = 1
Public Const TARGET_TYPE_NPC As Byte = 2

' Dialogue box constants
Public Const DIALOGUE_TYPE_NONE As Byte = 0
Public Const DIALOGUE_TYPE_TRADE As Byte = 1
Public Const DIALOGUE_TYPE_FORGET As Byte = 2
Public Const DIALOGUE_TYPE_PARTY As Byte = 3

' Do Events
Public Const nLng As Long = (&H80 Or &H1 Or &H4 Or &H20) + (&H8 Or &H40)

' Scrolling action message constants
Public Const ACTIONMSG_STATIC As Byte = 0
Public Const ACTIONMSG_SCROLL As Byte = 1
Public Const ACTIONMSG_SCREEN As Byte = 2

' stuffs
Public Const HalfX As Integer = ((MAX_MAPX + 1) / 2) * PIC_X
Public Const HalfY As Integer = ((MAX_MAPY + 1) / 2) * PIC_Y
Public Const ScreenX As Integer = (MAX_MAPX + 1) * PIC_X
Public Const ScreenY As Integer = (MAX_MAPY + 1) * PIC_Y
Public Const StartXValue As Integer = ((MAX_MAPX + 1) / 2)
Public Const StartYValue As Integer = ((MAX_MAPY + 1) / 2)
Public Const EndXValue As Integer = (MAX_MAPX + 1) + 1
Public Const EndYValue As Integer = (MAX_MAPY + 1) + 1
Public Const Half_PIC_X As Integer = PIC_X / 2
Public Const Half_PIC_Y As Integer = PIC_Y / 2

' weather สภาพอากาศ และอัตราการเกิดสภาพอากาศ
Public Const WEATHER_RAINING As Byte = 1
Public Const WEATHER_SNOWING As Byte = 2
Public Const WEATHER_BIRD As Byte = 3
Public Const WEATHER_SAND As Byte = 4
Public Const WEATHER_FIRE As Byte = 5


Public Const MAX_RAINDROPS As Byte = 80
Public Const MAX_SNOWDROPS As Byte = 60
Public Const MAX_BIRDDROPS As Byte = 12
Public Const MAX_SANDDROPS As Byte = 20
Public Const MAX_FIREDROPS As Byte = 12

' Sound wave
Public Const LEVEL_SOUND As String = "LevelUp.wav"
Public Const PUNCH_SOUND As String = "Punch.wav"
Public Const DIE_SOUND As String = "Die.wav"

' Full Screen
Public Full_Game As Boolean

' BuffStatus การเช็คสถานะบัฟ
Public Const BUFF_STUN As Byte = 1 ' มึน
Public Const BUFF_FREEZ As Byte = 2 ' แช่แข็ง
Public Const BUFF_NOEYE As Byte = 3 ' ตาบอด
Public Const BUFF_TOXIN As Byte = 4 ' ห้ามฟื้นฟูเลือด
Public Const BUFF_FEAR As Byte = 5 ' หวาดกลัว
Public Const BUFF_NOATK As Byte = 6 ' ปลดอาวุธ
Public Const BUFF_INVISIBLE As Byte = 7 ' หายตัว
Public Const BUFF_SILENT As Byte = 8 ' ใบ้
Public Const BUFF_NODEF As Byte = 9 ' เกราะแตก
Public Const BUFF_TLOCK As Byte = 10 ' ถูกล็อคเป้า
