Attribute VB_Name = "modGlobals"
Option Explicit
' loading screen
Public loadingText As String
' description
Public descType As Byte
Public descItem As Long
Public descLastType As Byte
Public descLastItem As Long
Public descText() As TextColourRec
' chars
Public CharName(1 To MAX_CHARS) As String
Public CharSprite(1 To MAX_CHARS) As Long
Public CharAccess(1 To MAX_CHARS) As Long
Public CharClass(1 To MAX_CHARS) As Long
Public CharNum As Long
Public usergroup As Long
' login
Public loginToken As String
'elastic bars
Public BarWidth_NpcHP(1 To MAX_MAP_NPCS) As Long
Public BarWidth_PlayerHP(1 To MAX_PLAYERS) As Long
Public BarWidth_NpcHP_Max(1 To MAX_MAP_NPCS) As Long
Public BarWidth_PlayerHP_Max(1 To MAX_PLAYERS) As Long
Public BarWidth_GuiHP As Long
Public BarWidth_GuiSP As Long
Public BarWidth_GuiEXP As Long
Public BarWidth_GuiHP_Max As Long
Public BarWidth_GuiSP_Max As Long
Public BarWidth_GuiEXP_Max As Long
' fog
Public fogOffsetX As Long
Public fogOffsetY As Long
' chat bubble
Public chatBubble(1 To MAX_BYTE) As ChatBubbleRec
Public chatBubbleIndex As Long
' Map animations
Public waterfallFrame As Long
Public autoTileFrame As Long
' tutorial
Public inTutorial As Long
Public tutorialState As Byte
' NPC Chat
Public chatNpc As Long
Public chatText As String
Public chatOpt(1 To 4) As String
' gui
Public hideGUI As Boolean
Public chatOn As Boolean
Public chatShowLine As String * 1
' map editor boxes
Public shpSelectedTop As Long
Public shpSelectedLeft As Long
Public shpSelectedHeight As Long
Public shpSelectedWidth As Long
Public shpLocTop As Long
Public shpLocLeft As Long
' autotiling
Public autoInner(1 To 4) As PointRec
Public autoNW(1 To 4) As PointRec
Public autoNE(1 To 4) As PointRec
Public autoSW(1 To 4) As PointRec
Public autoSE(1 To 4) As PointRec
' menu
Public inMenu As Boolean
' Cursor
Public GlobalX As Long
Public GlobalY As Long
Public GlobalX_Map As Long
Public GlobalY_Map As Long
' Paperdoll rendering order
Public PaperdollOrder() As Long
' music & sound list cache
Public musicCache() As String
Public soundCache() As String
Public hasPopulated As Boolean
' global dialogue index
Public diaHeader As String
Public diaBody As String
Public diaBody2 As String
Public diaIndex As Long
Public diaData1 As Long
Public diaDataString As String
Public diaStyle As Byte
' Hotbar
Public Hotbar(1 To MAX_HOTBAR) As HotbarRec
' Amount of blood decals
Public BloodCount As Long
' targetting
Public myTarget As Long
Public myTargetType As Long
' for directional blocking
Public DirArrowX(1 To 4) As Byte
Public DirArrowY(1 To 4) As Byte
' trading
Public InTrade As Long
Public TradeYourOffer(1 To MAX_INV) As PlayerInvRec
Public TradeTheirOffer(1 To MAX_INV) As PlayerInvRec
' Cache the Resources in an array
Public MapResource() As MapResourceRec
Public Resource_Index As Long
Public Resources_Init As Boolean
' drag + drop
Public DragInvSlotNum As Long
' gui
Public LastItemDesc As Long ' Stores the last item we showed in desc
Public tmpCurrencyItem As Long
Public InShop As Long ' is the player in a shop?
Public InBank As Long
Public inChat As Boolean
' Player variables
Public MyIndex As Long ' Index of actual player
Public PlayerInv(1 To MAX_INV) As PlayerInvRec   ' Inventory
Public PlayerSpells(1 To MAX_PLAYER_SPELLS) As PlayerSpellRec
Public InventoryItemSelected As Long
Public SpellBuffer As Long
Public SpellBufferTimer As Long
Public SpellCD(1 To MAX_PLAYER_SPELLS) As Long
Public StunDuration As Long
Public TNL As Long
' Stops movement when updating a map
Public CanMoveNow As Boolean
' Debug mode
Public DEBUG_MODE As Boolean
' TCP variables
Public PlayerBuffer As String
' Controls main gameloop
Public InGame As Boolean
Public isLogging As Boolean
' Game direction vars
Public ShiftDown As Boolean
Public ControlDown As Boolean
Public tabDown As Boolean
Public wDown As Boolean
Public sDown As Boolean
Public aDown As Boolean
Public dDown As Boolean
Public upDown As Boolean
Public downDown As Boolean
Public leftDown As Boolean
Public rightDown As Boolean
' Used to freeze controls when getting a new map
Public GettingMap As Boolean
' Used to check if FPS needs to be drawn
Public BFPS As Boolean
Public BLoc As Boolean
' FPS and Time-based movement vars
Public ElapsedTime As Long
Public GameFPS As Long
' Text vars
Public vbQuote As String
' Mouse cursor tile location
Public CurX As Long
Public CurY As Long
' Game editors
Public Editor As Byte
Public EditorIndex As Long
Public AnimEditorFrame(0 To 1) As Long
Public AnimEditorTimer(0 To 1) As Long
' Used to check if in editor or not and variables for use in editor
Public InMapEditor As Boolean
Public EditorTileX As Long
Public EditorTileY As Long
Public EditorTileWidth As Long
Public EditorTileHeight As Long
Public EditorWarpMap As Long
Public EditorWarpX As Long
Public EditorWarpY As Long
Public EditorWarpFall As Long
Public SpawnNpcNum As Long
Public SpawnNpcDir As Byte
Public EditorShop As Long
Public EditorEvent As Long
' appear
Public EditorAppearRange As Long
Public EditorAppearBottom As Long
' Used for map item editor
Public ItemEditorNum As Long
Public ItemEditorValue As Long
' Used for map key editor
Public KeyEditorNum As Long
Public KeyEditorTake As Long
Public KeyEditorTime As Long
' Used for map key open editor
Public KeyOpenEditorX As Long
Public KeyOpenEditorY As Long
' Map Resources
Public ResourceEditorNum As Long
' Used for map editor heal & trap & slide tiles
Public MapEditorHealType As Long
Public MapEditorHealAmount As Long
Public MapEditorSlideDir As Long
' used for map editor chat
Public MapEditorChatDir As Byte
Public MapEditorChatNpc As Long
' Maximum classes
Public Max_Classes As Long
Public Camera As RECT
Public TileView As RECT
' Pinging
Public PingStart As Long
Public PingEnd As Long
Public Ping As Long
' indexing
Public ActionMsgIndex As Byte
Public BloodIndex As Byte
Public AnimationIndex As Byte
' fps lock
Public FPS_Lock As Boolean
' Editor edited items array
Public Item_Changed(1 To MAX_ITEMS) As Boolean
Public NPC_Changed(1 To MAX_NPCS) As Boolean
Public Resource_Changed(1 To MAX_RESOURCES) As Boolean
Public Animation_Changed(1 To MAX_ANIMATIONS) As Boolean
Public Spell_Changed(1 To MAX_SPELLS) As Boolean
Public Shop_Changed(1 To MAX_SHOPS) As Boolean
Public Conv_Changed(1 To MAX_CONVS) As Boolean
' New char
Public newCharSprite As Long
Public newCharClass As Long
Public newCharGender As Long
' looping saves
Public Player_HighIndex As Long
Public Npc_HighIndex As Long
Public Action_HighIndex As Long
' fading
Public fadeAlpha As Long
' screenshot mode
Public screenshotMode As Long
' shop
Public shopSelectedSlot As Long
Public shopSelectedItem As Long
Public shopIsSelling As Boolean
' conv
Public convOptions As Long
Public optPos(1 To 4) As Long
Public optHeight As Long
' right click menu
Public PlayerMenuIndex As Long
' chat
Public inSmallChat As Boolean
Public actChatHeight As Long
Public actChatWidth As Long
Public ChatButtonUp As Boolean
Public ChatButtonDown As Boolean
' Events
Public selTileX As Long
Public selTileY As Long
Public inEvent As Boolean
Public eventNum As Long
Public eventPageNum As Long
Public eventCommandNum As Long
' Map
Public applyingMap As Boolean
Public MapEditorAppearDistance As Long
