Attribute VB_Name = "modVars"
Option Explicit

Public NumLines As Long
Public Const MAX_LINES As Long = 100

Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Sub CopyMemory Lib "Kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Sub ZeroMemory Lib "Kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByRef Msg() As Byte, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Const CLIENT_MAJOR As Byte = 1
Public Const CLIENT_MINOR As Byte = 8
Public Const CLIENT_REVISION As Byte = 0

Public Const GAME_SERVER_IP As String = "127.0.0.1" ' "46.23.70.66"
Public Const GAME_SERVER_PORT As Long = 7001 ' the port used by the main game server
Public Const AUTH_SERVER_PORT As Long = 7002 ' the port used for people to connect to auth server
Public Const SERVER_AUTH_PORT As Long = 7003 ' the portal used for server to talk to auth server

Public Const GAME_NAME As String = "Crystalshire"
Public Const GAME_WEBSITE As String = "http://www.crystalshire.com"

Public Const MAX_PLAYERS As Byte = 200

Public classMD5 As clsMD5

' Packets sent by authentication server to game server
Public Enum AuthPackets
    ASetPlayerLoginToken
    ASetUsergroup
End Enum

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
    SChatUpdate
    SConvEditor
    SUpdateConv
    SStartTutorial
    SChatBubble
    SSetPlayerLoginToken
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
    CTarget
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
    CChatOption
    CRequestEditConv
    CSaveConv
    CRequestConvs
    CFinishTutorial
    CAuthLogin
    ' Make sure CMSG_COUNT is below everything else
    CMSG_COUNT
End Enum

Public HandleDataSub(CMSG_COUNT) As Long

' dialogue alert strings
Public Const DIALOGUE_MSG_CONNECTION As Byte = 1
Public Const DIALOGUE_MSG_BANNED As Byte = 2
Public Const DIALOGUE_MSG_KICKED As Byte = 3
Public Const DIALOGUE_MSG_OUTDATED As Byte = 4
Public Const DIALOGUE_MSG_USERLENGTH As Byte = 5
Public Const DIALOGUE_MSG_ILLEGALNAME As Byte = 6
Public Const DIALOGUE_MSG_REBOOTING As Byte = 7
Public Const DIALOGUE_MSG_NAMETAKEN As Byte = 8
Public Const DIALOGUE_MSG_NAMELENGTH As Byte = 9
Public Const DIALOGUE_MSG_NAMEILLEGAL As Byte = 10
Public Const DIALOGUE_MSG_MYSQL As Byte = 11
Public Const DIALOGUE_MSG_WRONGPASS As Byte = 12
Public Const DIALOGUE_MSG_ACTIVATED As Byte = 13

' Menu
Public Const MENU_MAIN As Byte = 1
Public Const MENU_LOGIN As Byte = 2
Public Const MENU_REGISTER As Byte = 3
Public Const MENU_CREDITS As Byte = 4
Public Const MENU_CLASS As Byte = 5
Public Const MENU_NEWCHAR As Byte = 6
Public Const MENU_CHARS As Byte = 7
