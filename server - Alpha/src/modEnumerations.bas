Attribute VB_Name = "modEnumerations"
Option Explicit

' The order of the packets must match with the client's packet enumeration

' Packets sent by server to client
Public Enum ServerPackets
    SAlertMsg = 1
    SLoginOk
    SNewCharClasses
    SClassesData
    SInGame
    SPlayerInv
    SPlayerInvUpdate
    SPlayerWornEq
    SPlayerHp
    SPlayerMp
    SPlayerStats
    SPlayerData
    SPlayerBuff
    SPlayerMove
    SNpcMove
    SPlayerDir
    SNpcDir
    SPlayerXY
    SPlayerXYMap
    SAttack
    SNpcAttack
    SCheckForMap
    SMapData
    SMapItemData
    SMapNpcData
    SMapDone
    SGlobalMsg
    SAdminMsg
    SPlayerMsg
    SMapMsg
    SSpawnItem
    SItemEditor
    SUpdateItem
    SREditor
    SSpawnNpc
    SNpcDead
    SNpcEditor
    SUpdateNpc
    SMapKey
    SEditMap
    SShopEditor
    SUpdateShop
    SSpellEditor
    SUpdateSpell
    SSpells
    SLeft
    SResourceCache
    SResourceEditor
    SUpdateResource
    SSendPing
    SDoorAnimation
    SActionMsg
    SPlayerEXP
    SBlood
    SAnimationEditor
    SUpdateAnimation
    SAnimation
    SMapNpcVitals
    SCooldown
    SClearSpellBuffer
    SSayMsg
    SOpenShop
    SResetShopAction
    SStunned
    SMapWornEq
    SBank
    STrade
    SCloseTrade
    STradeUpdate
    STradeStatus
    STarget
    SHotbar
    SHighIndex
    SSound
    STradeRequest
    SPartyInvite
    SPartyUpdate
    SPartyVitals
    ' projectile
    SHandleProjectile
    ' quests
    SQuestEditor
    SUpdateQuest
    SPlayerQuest
    SQuestMessage
    ' guilds
    SSendGuild
    SAdminGuild
    SGuildAdminSwitchTab
    ' pets
    SNPCCache
    ' doors
    SDoorsEditor
    SUpdateDoors
    
    SSpawnEvent
    SEventMove
    SEventDir
    SEventChat
    SEventStart
    SEventEnd
    SPlayBGM
    SPlaySound
    SFadeoutBGM
    SStopSound
    SSwitchesAndVariables
    SMapEventData
    SSpell

    SOnDeath
    ' Make sure SMSG_COUNT is below everything else
    SMSG_COUNT
End Enum

' Packets sent by client to server
Public Enum ClientPackets
    CNewAccount = 1
    CDelAccount
    CLogin
    CAddChar
    CUseChar
    CSayMsg
    CEmoteMsg
    CBroadcastMsg
    CPlayerMsg
    CPlayerMove
    CPlayerDir
    CUseItem
    CAttack
    CUseStatPoint
    CPlayerInfoRequest
    CWarpMeTo
    CWarpToMe
    CWarpTo
    CSetSprite
    CGetStats
    CRequestNewMap
    CMapData
    CNeedMap
    CMapGetItem
    CMapDropItem
    CMapRespawn
    CMapReport
    CKickPlayer
    CBanList
    CBanDestroy
    CBanPlayer
    CRequestEditMap
    CRequestEditItem
    CSaveItem
    CRequestEditNpc
    CSaveNpc
    CRequestEditShop
    CSaveShop
    CRequestEditSpell
    CSaveSpell
    CSetAccess
    CWhosOnline
    CSetMotd
    CSearch
    CSpells
    CCast
    CQuit
    CSwapInvSlots
    CRequestEditResource
    CSaveResource
    CCheckPing
    CUnequip
    CRequestPlayerData
    CRequestItems
    CRequestNPCS
    CRequestResources
    CSpawnItem
    CRequestEditAnimation
    CSaveAnimation
    CRequestAnimations
    CRequestSpells
    CRequestShops
    CRequestLevelUp
    CForgetSpell
    CCloseShop
    CBuyItem
    CSellItem
    CChangeBankSlots
    CDepositItem
    CWithdrawItem
    CCloseBank
    CAdminWarp
    CTradeRequest
    CAcceptTrade
    CDeclineTrade
    CTradeItem
    CUntradeItem
    CHotbarChange
    CHotbarUse
    CSwapSpellSlots
    CAcceptTradeRequest
    CDeclineTradeRequest
    CPartyRequest
    CAcceptParty
    CDeclineParty
    CPartyLeave
    ' projectile
    CProjecTileAttack
    'quests
    CRequestEditQuest
    CSaveQuest
    CRequestQuests
    CPlayerHandleQuest
    CQuestLogUpdate
    ' party chat
    CPartyChatMsg
    ' guilds
    CGuildCommand
    CSayGuild
    CSaveGuild
    CRequestGuildAdminTabSwitch
    ' pets
    CSpawnPet
    CPetFollowOwner
    CPetAttackTarget
    CPetWander
    CPetDisband
    ' doors
    CSaveDoor
    CRequestDoors
    CRequestEditDoors
    
    CEventChatReply
    CEvent
    CSwitchesAndVariables
    CRequestSwitchesAndVariables
    
    COnDeath
    ' Make sure CMSG_COUNT is below everything else
    CMSG_COUNT
End Enum

Public HandleDataSub(CMSG_COUNT) As Long
' Public HandleDataSub(SMSG_COUNT) As Long

' Stats used by Players, Npcs and Classes
Public Enum Stats
    Strength = 1
    Endurance
    intelligence
    Agility
    willpower
    ' Make sure Stat_Count is below everything else
    Stat_Count
End Enum

' Vitals used by Players, Npcs and Classes
Public Enum Vitals
    HP = 1
    MP
    ' Make sure Vital_Count is below everything else
    Vital_Count
End Enum

' Equipment used by Players
Public Enum Equipment
    Weapon = 1
    Armor
    Helmet
    Shield
    ' Make sure Equipment_Count is below everything else
    Equipment_Count
End Enum

' Layers in a map
Public Enum MapLayer
    Ground = 1
    Mask
    Mask2
    Fringe
    Fringe2
' Make sure Layer_Count is below everything else
Layer_Count
End Enum

' Sound entities
Public Enum SoundEntity
    seAnimation = 1
    seItem
    seNpc
    seResource
    seSpell
    seLevelUp
    sePunch
    seDie
    ' Make sure SoundEntity_Count is below everything else
    SoundEntity_Count
End Enum

' Event Types
Public Enum EventType
    ' Message
    evAddText = 1
    evShowText
    evShowChoices
    ' Game Progression
    evPlayerVar
    evPlayerSwitch
    evSelfSwitch
    ' Flow Control
    evCondition
    evExitProcess
    ' Player
    evChangeItems
    evRestoreHP
    evRestoreMP
    evLevelUp
    evChangeLevel
    evChangeSkills
    evChangeClass
    evChangeSprite
    evChangeSex
    evChangePK
    ' Movement
    evWarpPlayer
    evSetMoveRoute
    ' Character
    evPlayAnimation
    ' Music and Sounds
    evPlayBGM
    evFadeoutBGM
    evPlaySound
    evStopSound
    'Etc...
    evCustomScript
    evSetAccess
    'Shop/Bank
    evOpenBank
    evOpenShop
    'New
    evGiveExp
End Enum
