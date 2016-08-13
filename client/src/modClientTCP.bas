Attribute VB_Name = "modClientTCP"
Option Explicit
' ******************************************
' ** Communcation to server, TCP          **
' ** Winsock Control (mswinsck.ocx)       **
' ** String packets (slow and big)        **
' ******************************************
Private PlayerBuffer As clsBuffer

Sub TcpInit(ByVal IP As String, ByVal Port As Long)
    Set PlayerBuffer = Nothing
    Set PlayerBuffer = New clsBuffer
    ' connect
    frmMain.Socket.Close
    frmMain.Socket.RemoteHost = IP
    frmMain.Socket.RemotePort = Port
End Sub

Sub DestroyTCP()
    frmMain.Socket.Close
End Sub

Public Sub IncomingData(ByVal DataLength As Long)
    Dim Buffer() As Byte
    Dim pLength As Long
    frmMain.Socket.GetData Buffer, vbUnicode, DataLength
    PlayerBuffer.WriteBytes Buffer()

    If PlayerBuffer.length >= 4 Then pLength = PlayerBuffer.ReadLong(False)

    Do While pLength > 0 And pLength <= PlayerBuffer.length - 4

        If pLength <= PlayerBuffer.length - 4 Then
            PlayerBuffer.ReadLong
            HandleData PlayerBuffer.ReadBytes(pLength)
        End If

        pLength = 0

        If PlayerBuffer.length >= 4 Then pLength = PlayerBuffer.ReadLong(False)
    Loop

    PlayerBuffer.Trim

    DoEvents
End Sub

Public Function ConnectToServer() As Boolean
    Dim Wait As Long

    ' Check to see if we are already connected, if so just exit
    If IsConnected Then
        ConnectToServer = True
        Exit Function
    End If

    Wait = GetTickCount
    frmMain.Socket.Close
    frmMain.Socket.Connect
    SetStatus "Connecting to server."

    ' Wait until connected or 3 seconds have passed and report the server being down
    Do While (Not IsConnected) And (GetTickCount <= Wait + 3000)
        DoEvents
    Loop

    ConnectToServer = IsConnected
    SetStatus vbNullString
End Function

Function IsConnected() As Boolean

    If frmMain.Socket.state = sckConnected Then
        IsConnected = True
    End If

End Function

Function IsPlaying(ByVal index As Long) As Boolean

    ' if the player doesn't exist, the name will equal 0
    If LenB(GetPlayerName(index)) > 0 Then
        IsPlaying = True
    End If

End Function

Sub SendData(ByRef data() As Byte)
    Dim Buffer As clsBuffer

    If IsConnected Then
        Set Buffer = New clsBuffer
        Buffer.WriteLong (UBound(data) - LBound(data)) + 1
        Buffer.WriteBytes data()
        frmMain.Socket.SendData Buffer.ToArray()
    End If

End Sub

' *****************************
' ** Outgoing Client Packets **
' *****************************
Public Sub SendNewAccount(ByVal name As String, ByVal password As String)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CNewAccount
    Buffer.WriteString name
    Buffer.WriteString password
    Buffer.WriteLong CLIENT_MAJOR
    Buffer.WriteLong CLIENT_MINOR
    Buffer.WriteLong CLIENT_REVISION
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendLogin(ByVal name As String)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CLogin
    Buffer.WriteString name
    Buffer.WriteString loginToken
    Buffer.WriteLong CLIENT_MAJOR
    Buffer.WriteLong CLIENT_MINOR
    Buffer.WriteLong CLIENT_REVISION
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendAuthLogin(ByVal name As String, ByVal password As String)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CAuthLogin
    Buffer.WriteString name
    Buffer.WriteString password
    Buffer.WriteLong CLIENT_MAJOR
    Buffer.WriteLong CLIENT_MINOR
    Buffer.WriteLong CLIENT_REVISION
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendAddChar(ByVal name As String, ByVal sex As Long, ByVal ClassNum As Long, ByVal sprite As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CAddChar
    Buffer.WriteString name
    Buffer.WriteLong sex
    Buffer.WriteLong ClassNum
    Buffer.WriteLong sprite
    Buffer.WriteLong CharNum
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendUseChar(ByVal CharSlot As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CUseChar
    Buffer.WriteLong CharSlot
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendDelChar(ByVal CharSlot As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CDelChar
    Buffer.WriteLong CharSlot
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SayMsg(ByVal text As String)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CSayMsg
    Buffer.WriteString text
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub BroadcastMsg(ByVal text As String)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CBroadcastMsg
    Buffer.WriteString text
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub EmoteMsg(ByVal text As String)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CEmoteMsg
    Buffer.WriteString text
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub PlayerMsg(ByVal text As String, ByVal MsgTo As String)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CSayMsg
    Buffer.WriteString MsgTo
    Buffer.WriteString text
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendPlayerMove()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CPlayerMove
    Buffer.WriteLong GetPlayerDir(MyIndex)
    Buffer.WriteLong Player(MyIndex).Moving
    Buffer.WriteLong Player(MyIndex).x
    Buffer.WriteLong Player(MyIndex).y
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendPlayerDir()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CPlayerDir
    Buffer.WriteLong GetPlayerDir(MyIndex)
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendPlayerRequestNewMap()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestNewMap
    Buffer.WriteLong GetPlayerDir(MyIndex)
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendMap()
    Dim x As Long
    Dim y As Long
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    CanMoveNow = False
    
    Buffer.WriteLong CMapData

    Buffer.WriteString Trim$(map.MapData.name)
    Buffer.WriteString Trim$(map.MapData.Music)
    Buffer.WriteByte map.MapData.Moral
    Buffer.WriteLong map.MapData.Up
    Buffer.WriteLong map.MapData.Down
    Buffer.WriteLong map.MapData.left
    Buffer.WriteLong map.MapData.Right
    Buffer.WriteLong map.MapData.BootMap
    Buffer.WriteByte map.MapData.BootX
    Buffer.WriteByte map.MapData.BootY
    Buffer.WriteByte map.MapData.MaxX
    Buffer.WriteByte map.MapData.MaxY
    Buffer.WriteLong map.MapData.BossNpc
    For i = 1 To MAX_MAP_NPCS
        Buffer.WriteLong map.MapData.Npc(i)
    Next
    
    Buffer.WriteLong map.TileData.EventCount
    If map.TileData.EventCount > 0 Then
        For i = 1 To map.TileData.EventCount
            With map.TileData.Events(i)
                Buffer.WriteString .name
                Buffer.WriteLong .x
                Buffer.WriteLong .y
                Buffer.WriteLong .pageCount
            End With
            If map.TileData.Events(i).pageCount > 0 Then
                For x = 1 To map.TileData.Events(i).pageCount
                    With map.TileData.Events(i).EventPage(x)
                        Buffer.WriteByte .chkPlayerVar
                        Buffer.WriteByte .chkSelfSwitch
                        Buffer.WriteByte .chkHasItem
                        Buffer.WriteLong .PlayerVarNum
                        Buffer.WriteLong .SelfSwitchNum
                        Buffer.WriteLong .HasItemNum
                        Buffer.WriteLong .PlayerVariable
                        Buffer.WriteByte .GraphicType
                        Buffer.WriteLong .Graphic
                        Buffer.WriteLong .GraphicX
                        Buffer.WriteLong .GraphicY
                        Buffer.WriteByte .MoveType
                        Buffer.WriteByte .MoveSpeed
                        Buffer.WriteByte .MoveFreq
                        Buffer.WriteByte .WalkAnim
                        Buffer.WriteByte .StepAnim
                        Buffer.WriteByte .DirFix
                        Buffer.WriteByte .WalkThrough
                        Buffer.WriteByte .Priority
                        Buffer.WriteByte .Trigger
                        Buffer.WriteLong .CommandCount
                    End With
                    If map.TileData.Events(i).EventPage(x).CommandCount > 0 Then
                        For y = 1 To map.TileData.Events(i).EventPage(x).CommandCount
                            With map.TileData.Events(i).EventPage(x).Commands(y)
                                Buffer.WriteByte .Type
                                Buffer.WriteString .text
                                Buffer.WriteLong .Colour
                                Buffer.WriteByte .channel
                                Buffer.WriteByte .TargetType
                                Buffer.WriteLong .target
                                Buffer.WriteLong .x
                                Buffer.WriteLong .y
                            End With
                        Next
                    End If
                Next
            End If
        Next
    End If

    For x = 0 To map.MapData.MaxX
        For y = 0 To map.MapData.MaxY
            With map.TileData.Tile(x, y)
                For i = 1 To MapLayer.Layer_Count - 1
                    Buffer.WriteLong .Layer(i).x
                    Buffer.WriteLong .Layer(i).y
                    Buffer.WriteLong .Layer(i).tileset
                    Buffer.WriteByte .Autotile(i)
                Next
                Buffer.WriteByte .Type
                Buffer.WriteLong .Data1
                Buffer.WriteLong .Data2
                Buffer.WriteLong .Data3
                Buffer.WriteLong .Data4
                Buffer.WriteLong .Data5
                Buffer.WriteByte .DirBlock
            End With
        Next
    Next

    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub WarpMeTo(ByVal name As String)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CWarpMeTo
    Buffer.WriteString name
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub WarpToMe(ByVal name As String)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CWarpToMe
    Buffer.WriteString name
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub WarpTo(ByVal mapNum As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CWarpTo
    Buffer.WriteLong mapNum
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendSetAccess(ByVal name As String, ByVal Access As Byte)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CSetAccess
    Buffer.WriteString name
    Buffer.WriteLong Access
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendSetSprite(ByVal SpriteNum As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CSetSprite
    Buffer.WriteLong SpriteNum
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendKick(ByVal name As String)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CKickPlayer
    Buffer.WriteString name
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendBan(ByVal name As String)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CBanPlayer
    Buffer.WriteString name
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendBanList()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CBanList
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendRequestEditItem()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestEditItem
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendSaveItem(ByVal itemNum As Long)
    Dim Buffer As clsBuffer
    Dim ItemSize As Long
    Dim ItemData() As Byte
    Set Buffer = New clsBuffer
    ItemSize = LenB(Item(itemNum))
    ReDim ItemData(ItemSize - 1)
    CopyMemory ItemData(0), ByVal VarPtr(Item(itemNum)), ItemSize
    Buffer.WriteLong CSaveItem
    Buffer.WriteLong itemNum
    Buffer.WriteBytes ItemData
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendRequestEditAnimation()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestEditAnimation
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendSaveAnimation(ByVal Animationnum As Long)
    Dim Buffer As clsBuffer
    Dim AnimationSize As Long
    Dim AnimationData() As Byte
    Set Buffer = New clsBuffer
    AnimationSize = LenB(Animation(Animationnum))
    ReDim AnimationData(AnimationSize - 1)
    CopyMemory AnimationData(0), ByVal VarPtr(Animation(Animationnum)), AnimationSize
    Buffer.WriteLong CSaveAnimation
    Buffer.WriteLong Animationnum
    Buffer.WriteBytes AnimationData
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendRequestEditNpc()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestEditNpc
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendSaveNpc(ByVal npcNum As Long)
    Dim Buffer As clsBuffer
    Dim NpcSize As Long
    Dim NpcData() As Byte
    Set Buffer = New clsBuffer
    NpcSize = LenB(Npc(npcNum))
    ReDim NpcData(NpcSize - 1)
    CopyMemory NpcData(0), ByVal VarPtr(Npc(npcNum)), NpcSize
    Buffer.WriteLong CSaveNpc
    Buffer.WriteLong npcNum
    Buffer.WriteBytes NpcData
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendRequestEditResource()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestEditResource
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendSaveResource(ByVal ResourceNum As Long)
    Dim Buffer As clsBuffer
    Dim ResourceSize As Long
    Dim ResourceData() As Byte
    Set Buffer = New clsBuffer
    ResourceSize = LenB(Resource(ResourceNum))
    ReDim ResourceData(ResourceSize - 1)
    CopyMemory ResourceData(0), ByVal VarPtr(Resource(ResourceNum)), ResourceSize
    Buffer.WriteLong CSaveResource
    Buffer.WriteLong ResourceNum
    Buffer.WriteBytes ResourceData
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendMapRespawn()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CMapRespawn
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendUseItem(ByVal invNum As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CUseItem
    Buffer.WriteLong invNum
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendDropItem(ByVal invNum As Long, ByVal amount As Long)
    Dim Buffer As clsBuffer

    If InBank Or InShop Then Exit Sub

    ' do basic checks
    If invNum < 1 Or invNum > MAX_INV Then Exit Sub
    If PlayerInv(invNum).num < 1 Or PlayerInv(invNum).num > MAX_ITEMS Then Exit Sub
    If Item(GetPlayerInvItemNum(MyIndex, invNum)).Type = ITEM_TYPE_CURRENCY Then
        If amount < 1 Or amount > PlayerInv(invNum).value Then Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong CMapDropItem
    Buffer.WriteLong invNum
    Buffer.WriteLong amount
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendWhosOnline()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CWhosOnline
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendMOTDChange(ByVal MOTD As String)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CSetMotd
    Buffer.WriteString MOTD
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendRequestEditShop()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestEditShop
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendSaveShop(ByVal shopNum As Long)
    Dim Buffer As clsBuffer
    Dim ShopSize As Long
    Dim ShopData() As Byte
    Set Buffer = New clsBuffer
    ShopSize = LenB(Shop(shopNum))
    ReDim ShopData(ShopSize - 1)
    CopyMemory ShopData(0), ByVal VarPtr(Shop(shopNum)), ShopSize
    Buffer.WriteLong CSaveShop
    Buffer.WriteLong shopNum
    Buffer.WriteBytes ShopData
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendRequestEditSpell()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestEditSpell
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendSaveSpell(ByVal spellnum As Long)
    Dim Buffer As clsBuffer
    Dim SpellSize As Long
    Dim SpellData() As Byte
    Set Buffer = New clsBuffer
    SpellSize = LenB(Spell(spellnum))
    ReDim SpellData(SpellSize - 1)
    CopyMemory SpellData(0), ByVal VarPtr(Spell(spellnum)), SpellSize
    Buffer.WriteLong CSaveSpell
    Buffer.WriteLong spellnum
    Buffer.WriteBytes SpellData
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendRequestEditMap()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestEditMap
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendBanDestroy()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CBanDestroy
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendChangeInvSlots(ByVal oldSlot As Long, ByVal newSlot As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CSwapInvSlots
    Buffer.WriteLong oldSlot
    Buffer.WriteLong newSlot
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    ' buffer it
    PlayerSwitchInvSlots oldSlot, newSlot
End Sub

Sub SendChangeSpellSlots(ByVal oldSlot As Long, ByVal newSlot As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CSwapSpellSlots
    Buffer.WriteLong oldSlot
    Buffer.WriteLong newSlot
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    ' buffer it
    PlayerSwitchSpellSlots oldSlot, newSlot
End Sub

Sub GetPing()
    Dim Buffer As clsBuffer
    PingStart = GetTickCount
    Set Buffer = New clsBuffer
    Buffer.WriteLong CCheckPing
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUnequip(ByVal eqNum As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CUnequip
    Buffer.WriteLong eqNum
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendRequestPlayerData()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestPlayerData
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendRequestItems()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestItems
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendRequestAnimations()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestAnimations
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendRequestNPCS()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestNPCS
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendRequestResources()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestResources
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendRequestSpells()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestSpells
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendRequestShops()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestShops
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendSpawnItem(ByVal tmpItem As Long, ByVal tmpAmount As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CSpawnItem
    Buffer.WriteLong tmpItem
    Buffer.WriteLong tmpAmount
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendTrainStat(ByVal statNum As Byte)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CUseStatPoint
    Buffer.WriteByte statNum
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendRequestLevelUp()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestLevelUp
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub BuyItem(ByVal shopSlot As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CBuyItem
    Buffer.WriteLong shopSlot
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SellItem(ByVal invSlot As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CSellItem
    Buffer.WriteLong invSlot
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub DepositItem(ByVal invSlot As Long, ByVal amount As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CDepositItem
    Buffer.WriteLong invSlot
    Buffer.WriteLong amount
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub WithdrawItem(ByVal bankslot As Long, ByVal amount As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CWithdrawItem
    Buffer.WriteLong bankslot
    Buffer.WriteLong amount
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub CloseBank()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CCloseBank
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    InBank = False
End Sub

Public Sub ChangeBankSlots(ByVal oldSlot As Long, ByVal newSlot As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CChangeBankSlots
    Buffer.WriteLong oldSlot
    Buffer.WriteLong newSlot
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub AdminWarp(ByVal x As Long, ByVal y As Long)
    If x < 0 Or y < 0 Or x > map.MapData.MaxX Or y > map.MapData.MaxY Then Exit Sub
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CAdminWarp
    Buffer.WriteLong x
    Buffer.WriteLong y
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub AcceptTrade()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CAcceptTrade
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub DeclineTrade()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CDeclineTrade
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub TradeItem(ByVal invSlot As Long, ByVal amount As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CTradeItem
    Buffer.WriteLong invSlot
    Buffer.WriteLong amount
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub UntradeItem(ByVal invSlot As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CUntradeItem
    Buffer.WriteLong invSlot
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendHotbarChange(ByVal sType As Long, ByVal Slot As Long, ByVal hotbarNum As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CHotbarChange
    Buffer.WriteLong sType
    Buffer.WriteLong Slot
    Buffer.WriteLong hotbarNum
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendHotbarUse(ByVal Slot As Long)
    Dim Buffer As clsBuffer, x As Long

    ' check if spell
    If Hotbar(Slot).sType = 2 Then ' spell

        For x = 1 To MAX_PLAYER_SPELLS

            ' is the spell matching the hotbar?
            If PlayerSpells(x).Spell = Hotbar(Slot).Slot Then
                ' found it, cast it
                CastSpell x
                Exit Sub
            End If

        Next

        ' can't find the spell, exit out
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong CHotbarUse
    Buffer.WriteLong Slot
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendMapReport()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CMapReport
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub PlayerTarget(ByVal target As Long, ByVal TargetType As Long)
    Dim Buffer As clsBuffer

    If myTargetType = TargetType And myTarget = target Then
        myTargetType = 0
        myTarget = 0
    Else
        myTarget = target
        myTargetType = TargetType
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong CTarget
    Buffer.WriteLong target
    Buffer.WriteLong TargetType
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendTradeRequest(playerIndex As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CTradeRequest
    Buffer.WriteLong playerIndex
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendAcceptTradeRequest()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CAcceptTradeRequest
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendDeclineTradeRequest()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CDeclineTradeRequest
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPartyLeave()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CPartyLeave
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPartyRequest(index As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CPartyRequest
    Buffer.WriteLong index
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendAcceptParty()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CAcceptParty
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendDeclineParty()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CDeclineParty
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendRequestEditConv()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestEditConv
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendSaveConv(ByVal Convnum As Long)
    Dim Buffer As clsBuffer
    Dim i As Long
    Dim x As Long
    Set Buffer = New clsBuffer
    Buffer.WriteLong CSaveConv
    Buffer.WriteLong Convnum

    With Conv(Convnum)
        Buffer.WriteString .name
        Buffer.WriteLong .chatCount

        For i = 1 To .chatCount
            Buffer.WriteString .Conv(i).Conv

            For x = 1 To 4
                Buffer.WriteString .Conv(i).rText(x)
                Buffer.WriteLong .Conv(i).rTarget(x)
            Next

            Buffer.WriteLong .Conv(i).Event
            Buffer.WriteLong .Conv(i).Data1
            Buffer.WriteLong .Conv(i).Data2
            Buffer.WriteLong .Conv(i).Data3
        Next

    End With

    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendRequestConvs()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestConvs
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendChatOption(ByVal index As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CChatOption
    Buffer.WriteLong index
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendFinishTutorial()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CFinishTutorial
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendCloseShop()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CCloseShop
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub
