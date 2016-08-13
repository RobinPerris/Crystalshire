Attribute VB_Name = "modPlayer"
Option Explicit

Public Sub InitChat(ByVal index As Long, ByVal mapnum As Long, ByVal mapNpcNum As Long, Optional ByVal remoteChat As Boolean = False)
Dim npcNum As Long
    npcNum = MapNpc(mapnum).Npc(mapNpcNum).Num
    
    ' check if we can chat
    If Npc(npcNum).Conv = 0 Then Exit Sub
    If Len(Trim$(Conv(Npc(npcNum).Conv).Name)) = 0 Then Exit Sub
    
    If Not remoteChat Then
        With MapNpc(mapnum).Npc(mapNpcNum)
            .c_inChatWith = index
            .c_lastDir = .dir
            If GetPlayerY(index) = .y - 1 Then
                .dir = DIR_UP
            ElseIf GetPlayerY(index) = .y + 1 Then
                .dir = DIR_DOWN
            ElseIf GetPlayerX(index) = .x - 1 Then
                .dir = DIR_LEFT
            ElseIf GetPlayerX(index) = .x + 1 Then
                .dir = DIR_RIGHT
            End If
            ' send NPC's dir to the map
            NpcDir mapnum, mapNpcNum, .dir
        End With
    End If
    
    ' Set chat value to Npc
    TempPlayer(index).inChatWith = npcNum
    TempPlayer(index).c_mapNpcNum = mapNpcNum
    TempPlayer(index).c_mapNum = mapnum
    ' set to the root chat
    TempPlayer(index).curChat = 1
    ' send the root chat
    sendChat index
End Sub

Public Sub chatOption(ByVal index As Long, ByVal chatOption As Long)
    Dim exitChat As Boolean
    Dim convNum As Long
    Dim curChat As Long
    
    If TempPlayer(index).inChatWith = 0 Then Exit Sub
    
    convNum = Npc(TempPlayer(index).inChatWith).Conv
    curChat = TempPlayer(index).curChat
    
    exitChat = False
    
    ' follow route
    If Conv(convNum).Conv(curChat).rTarget(chatOption) = 0 Then
        exitChat = True
    Else
        TempPlayer(index).curChat = Conv(convNum).Conv(curChat).rTarget(chatOption)
    End If
    
    ' if exiting chat, clear temp values
    If exitChat Then
        TempPlayer(index).inChatWith = 0
        TempPlayer(index).curChat = 0
        ' send chat update
        sendChat index
        ' send npc dir
        With MapNpc(TempPlayer(index).c_mapNum).Npc(TempPlayer(index).c_mapNpcNum)
            If .c_inChatWith = index Then
                .c_inChatWith = 0
                .dir = .c_lastDir
                NpcDir TempPlayer(index).c_mapNum, TempPlayer(index).c_mapNpcNum, .dir
            End If
        End With
        ' clear last of data
        TempPlayer(index).c_mapNpcNum = 0
        TempPlayer(index).c_mapNum = 0
        ' exit out early so we don't send chat update twice
        Exit Sub
    End If
    
    ' send update to the client
    sendChat index
End Sub

Public Sub chat_Unique(ByVal index As Long)
Dim convNum As Long
Dim curChat As Long
Dim itemAmount As Long
    
    If TempPlayer(index).inChatWith > 0 Then
        convNum = Npc(TempPlayer(index).inChatWith).Conv
        curChat = TempPlayer(index).curChat
        
        ' is unique?
        If Conv(convNum).Conv(curChat).Event = 4 Then ' unique
            ' which unique event?
            Select Case Conv(convNum).Conv(curChat).Data1
                Case 1 ' Little Boy
                    ' check has the gold
                    itemAmount = HasItem(index, 1)
                    If itemAmount = 0 Or itemAmount < 50 Then
                        PlayerMsg index, "You do not have enough gold.", BrightRed
                        Exit Sub
                    Else
                        PlayerWarp index, 15, 33, 32
                        SetPlayerDir index, DIR_LEFT
                        TakeInvItem index, 1, 50
                        PlayerMsg index, "The boy takes your money then pushes you head first through the hole.", BrightGreen
                        Exit Sub
                    End If
            End Select
        End If
    End If
End Sub

Public Sub sendChat(ByVal index As Long)
    Dim convNum As Long
    Dim curChat As Long
    Dim mainText As String
    Dim optText(1 To 4) As String
    Dim P_GENDER As String
    Dim P_NAME As String
    Dim P_CLASS As String
    Dim i As Long
    
    If TempPlayer(index).inChatWith > 0 Then
        convNum = Npc(TempPlayer(index).inChatWith).Conv
        curChat = TempPlayer(index).curChat
        
        ' check for unique events and trigger them early
        If Conv(convNum).Conv(curChat).Event > 0 Then
            Select Case Conv(convNum).Conv(curChat).Event
                Case 1 ' Open Shop
                    If Conv(convNum).Conv(curChat).Data1 > 0 Then ' shop exists?
                        SendOpenShop index, Conv(convNum).Conv(curChat).Data1
                        TempPlayer(index).InShop = Conv(convNum).Conv(curChat).Data1 ' stops movement and the like
                    End If
                    ' exit out early so we don't send chat update twice
                    ClosePlayerChat index
                    Exit Sub
                Case 2 ' Open Bank
                    SendBank index
                    TempPlayer(index).InBank = True
                    ' exit out early
                    ClosePlayerChat index
                    Exit Sub
                Case 3 ' Give Item
                    ' exit out early
                    ClosePlayerChat index
                    Exit Sub
                Case 4 ' Unique event
                    chat_Unique index
                    ClosePlayerChat index
                    Exit Sub
            End Select
        End If

continue:
        ' cache player's details
        If Player(index).Sex = SEX_MALE Then
            P_GENDER = "man"
        Else
            P_GENDER = "woman"
        End If
        P_NAME = Trim$(Player(index).Name)
        P_CLASS = Trim$(Class(Player(index).Class).Name)
        
        mainText = Conv(convNum).Conv(curChat).Conv
        For i = 1 To 4
            optText(i) = Conv(convNum).Conv(curChat).rText(i)
        Next
    End If
    
    SendChatUpdate index, TempPlayer(index).inChatWith, mainText, optText(1), optText(2), optText(3), optText(4)
    Exit Sub
End Sub

Public Sub ClosePlayerChat(ByVal index As Long)
    ' exit the chat
    TempPlayer(index).inChatWith = 0
    TempPlayer(index).curChat = 0
    ' send chat update
    sendChat index
    ' send npc dir
    With MapNpc(TempPlayer(index).c_mapNum).Npc(TempPlayer(index).c_mapNpcNum)
        If .c_inChatWith = index Then
            .c_inChatWith = 0
            .dir = .c_lastDir
            NpcDir TempPlayer(index).c_mapNum, TempPlayer(index).c_mapNpcNum, .dir
        End If
    End With
    ' clear last of data
    TempPlayer(index).c_mapNpcNum = 0
    TempPlayer(index).c_mapNum = 0
    Exit Sub
End Sub

Sub UseChar(ByVal index As Long, ByVal charNum As Long)
    If Not IsPlaying(index) Then
        Player(index).charNum = charNum
        Call LoadPlayer(index, Trim$(Player(index).Login), charNum)
        Call JoinGame(index)
        Call AddLog(GetPlayerLogin(index) & "/" & GetPlayerName(index) & " has began playing " & GAME_NAME & ".", PLAYER_LOG)
        Call TextAdd(GetPlayerLogin(index) & "/" & GetPlayerName(index) & " has began playing " & GAME_NAME & ".")
        Call UpdateCaption
    End If
End Sub

Sub JoinGame(ByVal index As Long)
    Dim i As Long
    
    ' Set the flag so we know the person is in the game
    TempPlayer(index).InGame = True
    'Update the log
    frmServer.lvwInfo.ListItems(index).SubItems(1) = GetPlayerIP(index)
    frmServer.lvwInfo.ListItems(index).SubItems(2) = GetPlayerLogin(index)
    frmServer.lvwInfo.ListItems(index).SubItems(3) = GetPlayerName(index)
    
    ' send the login ok
    SendLoginOk index
    
    TotalPlayersOnline = TotalPlayersOnline + 1
    
    ' Send some more little goodies, no need to explain these
    Call CheckEquippedItems(index)
    Call SendClasses(index)
    Call SendItems(index)
    Call SendAnimations(index)
    Call SendNpcs(index)
    Call SendShops(index)
    Call SendSpells(index)
    Call SendConvs(index)
    Call SendResources(index)
    Call SendInventory(index)
    Call SendWornEquipment(index)
    Call SendMapEquipment(index)
    Call SendPlayerSpells(index)
    Call SendHotbar(index)
    Call SendPlayerVariables(index)
    
    ' send vitals, exp + stats
    For i = 1 To Vitals.Vital_Count - 1
        Call SendVital(index, i)
    Next
    SendEXP index
    Call SendStats(index)
    
    ' Warp the player to his saved location
    Call PlayerWarp(index, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index))
    
    ' Send a global message that he/she joined
    Call GlobalMsg(GetPlayerName(index) & " has joined " & GAME_NAME & "!", White)
    
    ' Send welcome messages
    Call SendWelcome(index)

    ' Send Resource cache
    If GetPlayerMap(index) > 0 And GetPlayerMap(index) <= MAX_MAPS Then
        For i = 0 To ResourceCache(GetPlayerMap(index)).Resource_Count
            SendResourceCacheTo index, i
        Next
    End If
    
    ' Send the flag so they know they can start doing stuff
    SendInGame index
    
    ' tell them to do the damn tutorial
    If Player(index).TutorialState = 0 Then SendStartTutorial index
End Sub

Sub LeftGame(ByVal index As Long)
    Dim n As Long, i As Long
    Dim tradeTarget As Long
    
    If TempPlayer(index).InGame Then
        TempPlayer(index).InGame = False

        ' Check if player was the only player on the map and stop npc processing if so
        If GetPlayerMap(index) <= 0 Or GetPlayerMap(index) > MAX_MAPS Then Exit Sub
        If GetTotalMapPlayers(GetPlayerMap(index)) < 1 Then
            PlayersOnMap(GetPlayerMap(index)) = NO
        End If
        
        ' cancel any trade they're in
        If TempPlayer(index).InTrade > 0 Then
            tradeTarget = TempPlayer(index).InTrade
            PlayerMsg tradeTarget, Trim$(GetPlayerName(index)) & " has declined the trade.", BrightRed
            ' clear out trade
            For i = 1 To MAX_INV
                TempPlayer(tradeTarget).TradeOffer(i).Num = 0
                TempPlayer(tradeTarget).TradeOffer(i).Value = 0
            Next
            TempPlayer(tradeTarget).InTrade = 0
            SendCloseTrade tradeTarget
        End If
        
        ' leave party.
        Party_PlayerLeave index

        ' save and clear data.
        Call SavePlayer(index)

        ' Send a global message that he/she left
        Call GlobalMsg(GetPlayerName(index) & " has left " & GAME_NAME & "!", White)

        Call TextAdd(GetPlayerName(index) & " has disconnected from " & GAME_NAME & ".")
        Call SendLeftGame(index)
        TotalPlayersOnline = TotalPlayersOnline - 1
    End If

    Call ClearPlayer(index)
End Sub

Function GetPlayerProtection(ByVal index As Long) As Long
    Dim Armor As Long
    Dim Helm As Long
    GetPlayerProtection = 0

    ' Check for subscript out of range
    If IsPlaying(index) = False Or index <= 0 Or index > Player_HighIndex Then
        Exit Function
    End If

    Armor = GetPlayerEquipment(index, Armor)
    Helm = GetPlayerEquipment(index, Helmet)
    GetPlayerProtection = (GetPlayerStat(index, Stats.Endurance) \ 5)

    If Armor > 0 Then
        GetPlayerProtection = GetPlayerProtection + Item(Armor).Data2
    End If

    If Helm > 0 Then
        GetPlayerProtection = GetPlayerProtection + Item(Helm).Data2
    End If

End Function

Function CanPlayerCriticalHit(ByVal index As Long) As Boolean
    On Error Resume Next
    Dim i As Long
    Dim n As Long

    If GetPlayerEquipment(index, Weapon) > 0 Then
        n = (Rnd) * 2

        If n = 1 Then
            i = (GetPlayerStat(index, Stats.Strength) \ 2) + (GetPlayerLevel(index) \ 2)
            n = Int(Rnd * 100) + 1

            If n <= i Then
                CanPlayerCriticalHit = True
            End If
        End If
    End If

End Function

Function CanPlayerBlockHit(ByVal index As Long) As Boolean
    Dim i As Long
    Dim n As Long
    Dim ShieldSlot As Long
    ShieldSlot = GetPlayerEquipment(index, Shield)

    If ShieldSlot > 0 Then
        n = Int(Rnd * 2)

        If n = 1 Then
            i = (GetPlayerStat(index, Stats.Endurance) \ 2) + (GetPlayerLevel(index) \ 2)
            n = Int(Rnd * 100) + 1

            If n <= i Then
                CanPlayerBlockHit = True
            End If
        End If
    End If

End Function

Sub PlayerWarp(ByVal index As Long, ByVal mapnum As Long, ByVal x As Long, ByVal y As Long)
    Dim shopNum As Long
    Dim OldMap As Long
    Dim i As Long
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If IsPlaying(index) = False Or mapnum <= 0 Or mapnum > MAX_MAPS Then
        Exit Sub
    End If

    ' Check if you are out of bounds
    If x > Map(mapnum).MapData.MaxX Then x = Map(mapnum).MapData.MaxX
    If y > Map(mapnum).MapData.MaxY Then y = Map(mapnum).MapData.MaxY
    If x < 0 Then x = 0
    If y < 0 Then y = 0
    
    ' if same map then just send their co-ordinates
    If mapnum = GetPlayerMap(index) Then
        SendPlayerXYToMap index
    End If
    
    ' clear target
    TempPlayer(index).target = 0
    TempPlayer(index).targetType = TARGET_TYPE_NONE
    SendTarget index

    ' Save old map to send erase player data to
    OldMap = GetPlayerMap(index)

    If OldMap <> mapnum Then
        Call SendLeaveMap(index, OldMap)
    End If

    Call SetPlayerMap(index, mapnum)
    Call SetPlayerX(index, x)
    Call SetPlayerY(index, y)
    
    ' send player's equipment to new map
    SendMapEquipment index
    
    ' send equipment of all people on new map
    If GetTotalMapPlayers(mapnum) > 0 Then
        For i = 1 To Player_HighIndex
            If IsPlaying(i) Then
                If GetPlayerMap(i) = mapnum Then
                    SendMapEquipmentTo i, index
                End If
            End If
        Next
    End If

    ' Now we check if there were any players left on the map the player just left, and if not stop processing npcs
    If GetTotalMapPlayers(OldMap) = 0 Then
        PlayersOnMap(OldMap) = NO
        ' Regenerate all NPCs' health
        For i = 1 To MAX_MAP_NPCS
            If MapNpc(OldMap).Npc(i).Num > 0 Then
                MapNpc(OldMap).Npc(i).Vital(Vitals.HP) = GetNpcMaxVital(MapNpc(OldMap).Npc(i).Num, Vitals.HP)
            End If
        Next
    End If

    ' Sets it so we know to process npcs on the map
    PlayersOnMap(mapnum) = YES
    TempPlayer(index).GettingMap = YES
    SendCheckForMap index, mapnum
End Sub

Function CanMove(index As Long, dir As Long) As Byte
Dim warped As Boolean, newMapX As Long, newMapY As Long

    CanMove = 1
    Select Case dir
        Case DIR_UP
            ' Check to see if they are trying to go out of bounds
            If GetPlayerY(index) > 0 Then
                If CheckDirection(index, DIR_UP) Then
                    CanMove = 0
                    Exit Function
                End If
            Else
                ' Check if they can warp to a new map
                If Map(GetPlayerMap(index)).MapData.Up > 0 Then
                    newMapY = Map(Map(GetPlayerMap(index)).MapData.Up).MapData.MaxY
                    Call PlayerWarp(index, Map(GetPlayerMap(index)).MapData.Up, GetPlayerX(index), newMapY)
                    warped = True
                    CanMove = 2
                End If
                CanMove = 0
                Exit Function
            End If
        Case DIR_DOWN
            ' Check to see if they are trying to go out of bounds
            If GetPlayerY(index) < Map(GetPlayerMap(index)).MapData.MaxY Then
                If CheckDirection(index, DIR_DOWN) Then
                    CanMove = 0
                    Exit Function
                End If
            Else
                ' Check if they can warp to a new map
                If Map(GetPlayerMap(index)).MapData.Down > 0 Then
                    Call PlayerWarp(index, Map(GetPlayerMap(index)).MapData.Down, GetPlayerX(index), 0)
                    warped = True
                    CanMove = 2
                End If
                CanMove = False
                Exit Function
            End If
        Case DIR_LEFT
            ' Check to see if they are trying to go out of bounds
            If GetPlayerX(index) > 0 Then
                If CheckDirection(index, DIR_LEFT) Then
                    CanMove = 0
                    Exit Function
                End If
            Else
                ' Check if they can warp to a new map
                If Map(GetPlayerMap(index)).MapData.left > 0 Then
                    newMapX = Map(Map(GetPlayerMap(index)).MapData.left).MapData.MaxX
                    Call PlayerWarp(index, Map(GetPlayerMap(index)).MapData.left, newMapX, GetPlayerY(index))
                    warped = True
                    CanMove = 2
                End If
                CanMove = False
                Exit Function
            End If
        Case DIR_RIGHT
            ' Check to see if they are trying to go out of bounds
            If GetPlayerX(index) < Map(GetPlayerMap(index)).MapData.MaxX Then
                If CheckDirection(index, DIR_RIGHT) Then
                    CanMove = False
                    Exit Function
                End If
            Else
                ' Check if they can warp to a new map
                If Map(GetPlayerMap(index)).MapData.Right > 0 Then
                    Call PlayerWarp(index, Map(GetPlayerMap(index)).MapData.Right, 0, GetPlayerY(index))
                    warped = True
                    CanMove = 2
                End If
                CanMove = False
                Exit Function
            End If
    End Select
    ' check if we've warped
    If warped Then
        ' clear their target
        TempPlayer(index).target = 0
        TempPlayer(index).targetType = TARGET_TYPE_NONE
        SendTarget index
    End If
End Function

Function CheckDirection(index As Long, direction As Long) As Boolean
Dim x As Long, y As Long, i As Long, EventCount As Long, mapnum As Long, page As Long

    CheckDirection = False
    
    Select Case direction
        Case DIR_UP
            x = GetPlayerX(index)
            y = GetPlayerY(index) - 1
        Case DIR_DOWN
            x = GetPlayerX(index)
            y = GetPlayerY(index) + 1
        Case DIR_LEFT
            x = GetPlayerX(index) - 1
            y = GetPlayerY(index)
        Case DIR_RIGHT
            x = GetPlayerX(index) + 1
            y = GetPlayerY(index)
    End Select
    
    ' Check to see if the map tile is blocked or not
    If Map(GetPlayerMap(index)).TileData.Tile(x, y).Type = TILE_TYPE_BLOCKED Then
        CheckDirection = True
        Exit Function
    End If

    ' Check to see if the map tile is tree or not
    If Map(GetPlayerMap(index)).TileData.Tile(x, y).Type = TILE_TYPE_RESOURCE Then
        CheckDirection = True
        Exit Function
    End If
    
    ' Check to make sure that any events on that space aren't blocked
    mapnum = GetPlayerMap(index)
    EventCount = Map(mapnum).TileData.EventCount
    For i = 1 To EventCount
        With Map(mapnum).TileData.Events(i)
            If .x = x And .y = y Then
                ' Get the active event page
                page = ActiveEventPage(index, i)
                If page > 0 Then
                    If Map(mapnum).TileData.Events(i).EventPage(page).WalkThrough = 0 Then
                        CheckDirection = True
                        Exit Function
                    End If
                End If
            End If
        End With
    Next

    ' Check to see if a player is already on that tile
    If Map(GetPlayerMap(index)).MapData.Moral = 0 Then
        For i = 1 To Player_HighIndex
            If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(index) Then
                If GetPlayerX(i) = x Then
                    If GetPlayerY(i) = y Then
                        CheckDirection = True
                        Exit Function
                    End If
                End If
            End If
        Next i
    End If

    ' Check to see if a npc is already on that tile
    For i = 1 To MAX_MAP_NPCS
        If MapNpc(GetPlayerMap(index)).Npc(i).Num > 0 Then
            If MapNpc(GetPlayerMap(index)).Npc(i).x = x Then
                If MapNpc(GetPlayerMap(index)).Npc(i).y = y Then
                    CheckDirection = True
                    Exit Function
                End If
            End If
        End If
    Next
End Function

Sub PlayerMove(ByVal index As Long, ByVal dir As Long, ByVal movement As Long, Optional ByVal sendToSelf As Boolean = False)
    Dim Buffer As clsBuffer, mapnum As Long, x As Long, y As Long, moved As Byte, MovedSoFar As Boolean, newMapX As Byte, newMapY As Byte
    Dim TileType As Long, vitalType As Long, colour As Long, amount As Long, canMoveResult As Long, i As Long

    ' Check for subscript out of range
    If IsPlaying(index) = False Or dir < DIR_UP Or dir > DIR_RIGHT Or movement < 1 Or movement > 2 Then
        Exit Sub
    End If

    Call SetPlayerDir(index, dir)
    moved = NO
    mapnum = GetPlayerMap(index)
    
    If mapnum = 0 Then Exit Sub
    
    ' check if they're casting a spell
    If TempPlayer(index).spellBuffer.Spell > 0 Then
        SendCancelAnimation index
        SendClearSpellBuffer index
        TempPlayer(index).spellBuffer.Spell = 0
        TempPlayer(index).spellBuffer.target = 0
        TempPlayer(index).spellBuffer.Timer = 0
        TempPlayer(index).spellBuffer.tType = 0
    End If
    
    ' check directions
    canMoveResult = CanMove(index, dir)
    If canMoveResult = 1 Then
        Select Case dir
            Case DIR_UP
                Call SetPlayerY(index, GetPlayerY(index) - 1)
                SendPlayerMove index, movement, sendToSelf
                moved = YES
            Case DIR_DOWN
                Call SetPlayerY(index, GetPlayerY(index) + 1)
                SendPlayerMove index, movement, sendToSelf
                moved = YES
            Case DIR_LEFT
                Call SetPlayerX(index, GetPlayerX(index) - 1)
                SendPlayerMove index, movement, sendToSelf
                moved = YES
            Case DIR_RIGHT
                Call SetPlayerX(index, GetPlayerX(index) + 1)
                SendPlayerMove index, movement, sendToSelf
                moved = YES
        End Select
    End If
    
    With Map(GetPlayerMap(index)).TileData.Tile(GetPlayerX(index), GetPlayerY(index))
        ' Check to see if the tile is a warp tile, and if so warp them
        If .Type = TILE_TYPE_WARP Then
            mapnum = .Data1
            x = .Data2
            y = .Data3
            Call PlayerWarp(index, mapnum, x, y)
            moved = YES
        End If
    
        ' Check to see if the tile is a door tile, and if so warp them
        If .Type = TILE_TYPE_DOOR Then
            mapnum = .Data1
            x = .Data2
            y = .Data3
            ' send the animation to the map
            SendDoorAnimation GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index)
            Call PlayerWarp(index, mapnum, x, y)
            moved = YES
        End If
    
        ' Check for key trigger open
        If .Type = TILE_TYPE_KEYOPEN Then
            x = .Data1
            y = .Data2
    
            If Map(GetPlayerMap(index)).TileData.Tile(x, y).Type = TILE_TYPE_KEY And TempTile(GetPlayerMap(index)).DoorOpen(x, y) = NO Then
                TempTile(GetPlayerMap(index)).DoorOpen(x, y) = YES
                TempTile(GetPlayerMap(index)).DoorTimer = GetTickCount
                SendMapKey index, x, y, 1
                'Call MapMsg(GetPlayerMap(index), "A door has been unlocked.", White)
            End If
        End If
        
        ' Check for a shop, and if so open it
        If .Type = TILE_TYPE_SHOP Then
            x = .Data1
            If x > 0 Then ' shop exists?
                If Len(Trim$(Shop(x).Name)) > 0 Then ' name exists?
                    SendOpenShop index, x
                    TempPlayer(index).InShop = x ' stops movement and the like
                End If
            End If
        End If
        
        ' Check to see if the tile is a bank, and if so send bank
        If .Type = TILE_TYPE_BANK Then
            SendBank index
            TempPlayer(index).InBank = True
            moved = YES
        End If
        
        ' Check if it's a heal tile
        If .Type = TILE_TYPE_HEAL Then
            vitalType = .Data1
            amount = .Data2
            If Not GetPlayerVital(index, vitalType) = GetPlayerMaxVital(index, vitalType) Then
                If vitalType = Vitals.HP Then
                    colour = BrightGreen
                Else
                    colour = BrightBlue
                End If
                SendActionMsg GetPlayerMap(index), "+" & amount, colour, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32, 1
                SetPlayerVital index, vitalType, GetPlayerVital(index, vitalType) + amount
                PlayerMsg index, "You feel rejuvinating forces flowing through your boy.", BrightGreen
                Call SendVital(index, vitalType)
                ' send vitals to party if in one
                If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
            End If
            moved = YES
        End If
        
        ' Check if it's a trap tile
        If .Type = TILE_TYPE_TRAP Then
            amount = .Data1
            SendActionMsg GetPlayerMap(index), "-" & amount, BrightRed, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32, 1
            If GetPlayerVital(index, HP) - amount <= 0 Then
                KillPlayer index
                PlayerMsg index, "You're killed by a trap.", BrightRed
            Else
                SetPlayerVital index, HP, GetPlayerVital(index, HP) - amount
                PlayerMsg index, "You're injured by a trap.", BrightRed
                Call SendVital(index, HP)
                ' send vitals to party if in one
                If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
            End If
            moved = YES
        End If
    End With
    
    ' check for events
    If Map(GetPlayerMap(index)).TileData.EventCount > 0 Then
        For i = 1 To Map(GetPlayerMap(index)).TileData.EventCount
            CheckPlayerEvent index, i
        Next
    End If

    ' They tried to hack
    If moved = NO And canMoveResult <> 2 Then
        PlayerWarp index, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index)
    End If
End Sub

Sub CheckPlayerEvent(index As Long, eventNum As Long)
Dim Count As Long, mapnum As Long, i As Long
    ' find the page to process
    mapnum = GetPlayerMap(index)
    ' make sure it's in the same spot
    If Map(mapnum).TileData.Events(eventNum).x <> GetPlayerX(index) Then Exit Sub
    If Map(mapnum).TileData.Events(eventNum).y <> GetPlayerY(index) Then Exit Sub
    ' loop
    Count = Map(mapnum).TileData.Events(eventNum).PageCount
    ' get the active page
    i = ActiveEventPage(index, eventNum)
    ' exit out early
    If i = 0 Then Exit Sub
    ' make sure the page has actual commands
    If Map(mapnum).TileData.Events(eventNum).EventPage(i).CommandCount = 0 Then Exit Sub
    ' set event
    TempPlayer(index).inEvent = True
    TempPlayer(index).eventNum = eventNum
    TempPlayer(index).pageNum = i
    TempPlayer(index).commandNum = 1
    ' send it to the player
    SendEvent index
End Sub

Sub EventLogic(index As Long)
Dim eventNum As Long, pageNum As Long, commandNum As Long
    eventNum = TempPlayer(index).eventNum
    pageNum = TempPlayer(index).pageNum
    commandNum = TempPlayer(index).commandNum
    ' carry out the command
    With Map(GetPlayerMap(index)).TileData.Events(eventNum).EventPage(pageNum)
        ' server-side processing
        Select Case .Commands(commandNum).Type
            Case EventType.evPlayerVar
                If .Commands(commandNum).target > 0 Then Player(index).Variable(.Commands(commandNum).target) = .Commands(commandNum).colour
        End Select
        ' increment commands
        If commandNum < .CommandCount Then
            TempPlayer(index).commandNum = TempPlayer(index).commandNum + 1
            Exit Sub
        End If
    End With
    ' we're done - close event
    TempPlayer(index).eventNum = 0
    TempPlayer(index).pageNum = 0
    TempPlayer(index).commandNum = 0
    TempPlayer(index).inEvent = False
    ' send it to the player
    'SendEvent index
End Sub

Sub ForcePlayerMove(ByVal index As Long, ByVal movement As Long, ByVal direction As Long)
    If direction < DIR_UP Or direction > DIR_RIGHT Then Exit Sub
    If movement < 1 Or movement > 2 Then Exit Sub
    
    Select Case direction
        Case DIR_UP
            If GetPlayerY(index) = 0 Then Exit Sub
        Case DIR_LEFT
            If GetPlayerX(index) = 0 Then Exit Sub
        Case DIR_DOWN
            If GetPlayerY(index) = Map(GetPlayerMap(index)).MapData.MaxY Then Exit Sub
        Case DIR_RIGHT
            If GetPlayerX(index) = Map(GetPlayerMap(index)).MapData.MaxX Then Exit Sub
    End Select
    
    PlayerMove index, direction, movement, True
End Sub

Sub CheckEquippedItems(ByVal index As Long)
    Dim Slot As Long
    Dim itemNum As Long
    Dim i As Long

    ' We want to check incase an admin takes away an object but they had it equipped
    For i = 1 To Equipment.Equipment_Count - 1
        itemNum = GetPlayerEquipment(index, i)

        If itemNum > 0 Then

            Select Case i
                Case Equipment.Weapon

                    If Item(itemNum).Type <> ITEM_TYPE_WEAPON Then SetPlayerEquipment index, 0, i
                Case Equipment.Armor

                    If Item(itemNum).Type <> ITEM_TYPE_ARMOR Then SetPlayerEquipment index, 0, i
                Case Equipment.Helmet

                    If Item(itemNum).Type <> ITEM_TYPE_HELMET Then SetPlayerEquipment index, 0, i
                Case Equipment.Shield

                    If Item(itemNum).Type <> ITEM_TYPE_SHIELD Then SetPlayerEquipment index, 0, i
            End Select

        Else
            SetPlayerEquipment index, 0, i
        End If

    Next

End Sub

Function FindOpenInvSlot(ByVal index As Long, ByVal itemNum As Long) As Long
    Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(index) = False Or itemNum <= 0 Or itemNum > MAX_ITEMS Then
        Exit Function
    End If

    If Item(itemNum).Type = ITEM_TYPE_CURRENCY Then

        ' If currency then check to see if they already have an instance of the item and add it to that
        For i = 1 To MAX_INV

            If GetPlayerInvItemNum(index, i) = itemNum Then
                FindOpenInvSlot = i
                Exit Function
            End If

        Next

    End If

    For i = 1 To MAX_INV

        ' Try to find an open free slot
        If GetPlayerInvItemNum(index, i) = 0 Then
            FindOpenInvSlot = i
            Exit Function
        End If

    Next

End Function

Function FindOpenBankSlot(ByVal index As Long, ByVal itemNum As Long) As Long
    Dim i As Long

    If Not IsPlaying(index) Then Exit Function
    If itemNum <= 0 Or itemNum > MAX_ITEMS Then Exit Function

        For i = 1 To MAX_BANK
            If GetPlayerBankItemNum(index, i) = itemNum Then
                FindOpenBankSlot = i
                Exit Function
            End If
        Next i

    For i = 1 To MAX_BANK
        If GetPlayerBankItemNum(index, i) = 0 Then
            FindOpenBankSlot = i
            Exit Function
        End If
    Next i

End Function

Function HasItem(ByVal index As Long, ByVal itemNum As Long) As Long
    Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(index) = False Or itemNum <= 0 Or itemNum > MAX_ITEMS Then
        Exit Function
    End If

    For i = 1 To MAX_INV

        ' Check to see if the player has the item
        If GetPlayerInvItemNum(index, i) = itemNum Then
            If Item(itemNum).Type = ITEM_TYPE_CURRENCY Then
                HasItem = GetPlayerInvItemValue(index, i)
            Else
                HasItem = 1
            End If

            Exit Function
        End If

    Next

End Function

Function TakeInvItem(ByVal index As Long, ByVal itemNum As Long, ByVal ItemVal As Long) As Boolean
    Dim i As Long
    Dim n As Long
    
    TakeInvItem = False

    ' Check for subscript out of range
    If IsPlaying(index) = False Or itemNum <= 0 Or itemNum > MAX_ITEMS Then
        Exit Function
    End If

    For i = 1 To MAX_INV

        ' Check to see if the player has the item
        If GetPlayerInvItemNum(index, i) = itemNum Then
            If Item(itemNum).Type = ITEM_TYPE_CURRENCY Then

                ' Is what we are trying to take away more then what they have?  If so just set it to zero
                If ItemVal >= GetPlayerInvItemValue(index, i) Then
                    TakeInvItem = True
                Else
                    Call SetPlayerInvItemValue(index, i, GetPlayerInvItemValue(index, i) - ItemVal)
                    Call SendInventoryUpdate(index, i)
                End If
            Else
                TakeInvItem = True
            End If

            If TakeInvItem Then
                Call SetPlayerInvItemNum(index, i, 0)
                Call SetPlayerInvItemValue(index, i, 0)
                Player(index).Inv(i).Bound = 0
                ' Send the inventory update
                Call SendInventoryUpdate(index, i)
                Exit Function
            End If
        End If

    Next

End Function

Function TakeInvSlot(ByVal index As Long, ByVal invSlot As Long, ByVal ItemVal As Long) As Boolean
    Dim i As Long
    Dim n As Long
    Dim itemNum
    
    TakeInvSlot = False

    ' Check for subscript out of range
    If IsPlaying(index) = False Or invSlot <= 0 Or invSlot > MAX_ITEMS Then
        Exit Function
    End If
    
    itemNum = GetPlayerInvItemNum(index, invSlot)

    If Item(itemNum).Type = ITEM_TYPE_CURRENCY Then

        ' Is what we are trying to take away more then what they have?  If so just set it to zero
        If ItemVal >= GetPlayerInvItemValue(index, invSlot) Then
            TakeInvSlot = True
        Else
            Call SetPlayerInvItemValue(index, invSlot, GetPlayerInvItemValue(index, invSlot) - ItemVal)
        End If
    Else
        TakeInvSlot = True
    End If

    If TakeInvSlot Then
        Call SetPlayerInvItemNum(index, invSlot, 0)
        Call SetPlayerInvItemValue(index, invSlot, 0)
        Player(index).Inv(invSlot).Bound = 0
        Exit Function
    End If

End Function

Function GiveInvItem(ByVal index As Long, ByVal itemNum As Long, ByVal ItemVal As Long, Optional ByVal sendUpdate As Boolean = True, Optional ByVal forceBound As Boolean = False) As Boolean
    Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(index) = False Or itemNum <= 0 Or itemNum > MAX_ITEMS Then
        GiveInvItem = False
        Exit Function
    End If

    i = FindOpenInvSlot(index, itemNum)

    ' Check to see if inventory is full
    If i <> 0 Then
        Call SetPlayerInvItemNum(index, i, itemNum)
        Call SetPlayerInvItemValue(index, i, GetPlayerInvItemValue(index, i) + ItemVal)
        ' force bound?
        If Not forceBound Then
            ' bind on pickup?
            If Item(itemNum).BindType = 1 Then ' bind on pickup
                Player(index).Inv(i).Bound = 1
                PlayerMsg index, "This item is now bound to your soul.", BrightRed
            Else
                Player(index).Inv(i).Bound = 0
            End If
        Else
            Player(index).Inv(i).Bound = 1
        End If
        ' send update
        If sendUpdate Then Call SendInventoryUpdate(index, i)
        GiveInvItem = True
    Else
        Call PlayerMsg(index, "Your inventory is full.", BrightRed)
        GiveInvItem = False
    End If

End Function

Function HasSpell(ByVal index As Long, ByVal spellNum As Long) As Boolean
    Dim i As Long

    For i = 1 To MAX_PLAYER_SPELLS

        If Player(index).Spell(i).Spell = spellNum Then
            HasSpell = True
            Exit Function
        End If

    Next

End Function

Function FindOpenSpellSlot(ByVal index As Long) As Long
    Dim i As Long

    For i = 1 To MAX_PLAYER_SPELLS

        If Player(index).Spell(i).Spell = 0 Then
            FindOpenSpellSlot = i
            Exit Function
        End If

    Next

End Function

Sub PlayerMapGetItem(ByVal index As Long)
    Dim i As Long
    Dim n As Long
    Dim mapnum As Long
    Dim Msg As String

    If Not IsPlaying(index) Then Exit Sub
    mapnum = GetPlayerMap(index)
    
    If mapnum = 0 Then Exit Sub

    For i = 1 To MAX_MAP_ITEMS
        ' See if theres even an item here
        If (MapItem(mapnum, i).Num > 0) And (MapItem(mapnum, i).Num <= MAX_ITEMS) Then
            ' our drop?
            If CanPlayerPickupItem(index, i) Then
                ' Check if item is at the same location as the player
                If (MapItem(mapnum, i).x = GetPlayerX(index)) Then
                    If (MapItem(mapnum, i).y = GetPlayerY(index)) Then
                        ' Find open slot
                        n = FindOpenInvSlot(index, MapItem(mapnum, i).Num)
    
                        ' Open slot available?
                        If n <> 0 Then
                            ' Set item in players inventor
                            Call SetPlayerInvItemNum(index, n, MapItem(mapnum, i).Num)
    
                            If Item(GetPlayerInvItemNum(index, n)).Type = ITEM_TYPE_CURRENCY Then
                                Call SetPlayerInvItemValue(index, n, GetPlayerInvItemValue(index, n) + MapItem(mapnum, i).Value)
                                Msg = MapItem(mapnum, i).Value & " " & Trim$(Item(GetPlayerInvItemNum(index, n)).Name)
                            Else
                                Call SetPlayerInvItemValue(index, n, 0)
                                Msg = Trim$(Item(GetPlayerInvItemNum(index, n)).Name)
                            End If
                            
                            ' is it bind on pickup?
                            Player(index).Inv(n).Bound = 0
                            If Item(GetPlayerInvItemNum(index, n)).BindType = 1 Or MapItem(mapnum, i).Bound Then
                                Player(index).Inv(n).Bound = 1
                                If Not Trim$(MapItem(mapnum, i).playerName) = Trim$(GetPlayerName(index)) Then
                                    PlayerMsg index, "This item is now bound to your soul.", BrightRed
                                End If
                            End If

                            ' Erase item from the map
                            ClearMapItem i, mapnum
                            
                            Call SendInventoryUpdate(index, n)
                            Call SpawnItemSlot(i, 0, 0, GetPlayerMap(index), 0, 0)
                            SendActionMsg GetPlayerMap(index), Msg, White, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
                            Exit For
                        Else
                            Call PlayerMsg(index, "Your inventory is full.", BrightRed)
                            Exit For
                        End If
                    End If
                End If
            End If
        End If
    Next
End Sub

Function CanPlayerPickupItem(ByVal index As Long, ByVal mapItemNum As Long)
Dim mapnum As Long, tmpIndex As Long, i As Long

    mapnum = GetPlayerMap(index)
    
    ' no lock or locked to player?
    If MapItem(mapnum, mapItemNum).playerName = vbNullString Or MapItem(mapnum, mapItemNum).playerName = Trim$(GetPlayerName(index)) Then
        CanPlayerPickupItem = True
        Exit Function
    End If
    
    ' if in party show their party member's drops
    If TempPlayer(index).inParty > 0 Then
        For i = 1 To MAX_PARTY_MEMBERS
            tmpIndex = Party(TempPlayer(index).inParty).Member(i)
            If tmpIndex > 0 Then
                If Trim$(GetPlayerName(tmpIndex)) = MapItem(mapnum, mapItemNum).playerName Then
                    If MapItem(mapnum, mapItemNum).Bound = 0 Then
                        CanPlayerPickupItem = True
                        Exit Function
                    End If
                End If
            End If
        Next
    End If
    
    ' exit out
    CanPlayerPickupItem = False
End Function

Sub PlayerMapDropItem(ByVal index As Long, ByVal invNum As Long, ByVal amount As Long)
    Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(index) = False Or invNum <= 0 Or invNum > MAX_INV Then
        Exit Sub
    End If
    
    ' check the player isn't doing something
    If TempPlayer(index).InBank Or TempPlayer(index).InShop Or TempPlayer(index).InTrade > 0 Then Exit Sub

    If (GetPlayerInvItemNum(index, invNum) > 0) Then
        If (GetPlayerInvItemNum(index, invNum) <= MAX_ITEMS) Then
            ' make sure it's not bound
            If Item(GetPlayerInvItemNum(index, invNum)).BindType > 0 Then
                If Player(index).Inv(invNum).Bound = 1 Then
                    PlayerMsg index, "This item is soulbound and cannot be picked up by other players.", BrightRed
                End If
            End If
            
            i = FindOpenMapItemSlot(GetPlayerMap(index))

            If i <> 0 Then
                MapItem(GetPlayerMap(index), i).Num = GetPlayerInvItemNum(index, invNum)
                MapItem(GetPlayerMap(index), i).x = GetPlayerX(index)
                MapItem(GetPlayerMap(index), i).y = GetPlayerY(index)
                MapItem(GetPlayerMap(index), i).playerName = Trim$(GetPlayerName(index))
                MapItem(GetPlayerMap(index), i).playerTimer = GetTickCount + ITEM_SPAWN_TIME
                MapItem(GetPlayerMap(index), i).canDespawn = True
                MapItem(GetPlayerMap(index), i).despawnTimer = GetTickCount + ITEM_DESPAWN_TIME
                If Player(index).Inv(invNum).Bound > 0 Then
                    MapItem(GetPlayerMap(index), i).Bound = True
                Else
                    MapItem(GetPlayerMap(index), i).Bound = False
                End If

                If Item(GetPlayerInvItemNum(index, invNum)).Type = ITEM_TYPE_CURRENCY Then

                    ' Check if its more then they have and if so drop it all
                    If amount >= GetPlayerInvItemValue(index, invNum) Then
                        MapItem(GetPlayerMap(index), i).Value = GetPlayerInvItemValue(index, invNum)
                        Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " drops " & GetPlayerInvItemValue(index, invNum) & " " & Trim$(Item(GetPlayerInvItemNum(index, invNum)).Name) & ".", Yellow)
                        Call SetPlayerInvItemNum(index, invNum, 0)
                        Call SetPlayerInvItemValue(index, invNum, 0)
                        Player(index).Inv(invNum).Bound = 0
                    Else
                        MapItem(GetPlayerMap(index), i).Value = amount
                        Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " drops " & amount & " " & Trim$(Item(GetPlayerInvItemNum(index, invNum)).Name) & ".", Yellow)
                        Call SetPlayerInvItemValue(index, invNum, GetPlayerInvItemValue(index, invNum) - amount)
                    End If

                Else
                    ' Its not a currency object so this is easy
                    MapItem(GetPlayerMap(index), i).Value = 0
                    ' send message
                    Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " drops " & CheckGrammar(Trim$(Item(GetPlayerInvItemNum(index, invNum)).Name)) & ".", Yellow)
                    Call SetPlayerInvItemNum(index, invNum, 0)
                    Call SetPlayerInvItemValue(index, invNum, 0)
                    Player(index).Inv(invNum).Bound = 0
                End If

                ' Send inventory update
                Call SendInventoryUpdate(index, invNum)
                ' Spawn the item before we set the num or we'll get a different free map item slot
                Call SpawnItemSlot(i, MapItem(GetPlayerMap(index), i).Num, amount, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index), Trim$(GetPlayerName(index)), MapItem(GetPlayerMap(index), i).canDespawn, MapItem(GetPlayerMap(index), i).Bound)
            Else
                Call PlayerMsg(index, "Too many items already on the ground.", BrightRed)
            End If
        End If
    End If

End Sub

Sub CheckPlayerLevelUp(ByVal index As Long)
    Dim i As Long
    Dim expRollover As Long
    Dim level_count As Long
    
    level_count = 0
    
    Do While GetPlayerExp(index) >= GetPlayerNextLevel(index)
        expRollover = GetPlayerExp(index) - GetPlayerNextLevel(index)
        
        ' can level up?
        If Not SetPlayerLevel(index, GetPlayerLevel(index) + 1) Then
            Exit Sub
        End If
        
        Call SetPlayerPOINTS(index, GetPlayerPOINTS(index) + 3)
        Call SetPlayerExp(index, expRollover)
        level_count = level_count + 1
    Loop
    
    If level_count > 0 Then
        If level_count = 1 Then
            'singular
            GlobalMsg GetPlayerName(index) & " has gained " & level_count & " level!", Brown
        Else
            'plural
            GlobalMsg GetPlayerName(index) & " has gained " & level_count & " levels!", Brown
        End If
        SendEXP index
        SendPlayerData index
    End If
End Sub

' //////////////////////
' // PLAYER FUNCTIONS //
' //////////////////////
Function GetPlayerLogin(ByVal index As Long) As String
    GetPlayerLogin = Trim$(Player(index).Login)
End Function

Sub SetPlayerLogin(ByVal index As Long, ByVal Login As String)
    Player(index).Login = Login
End Sub

Function GetPlayerName(ByVal index As Long) As String

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerName = Trim$(Player(index).Name)
End Function

Sub SetPlayerName(ByVal index As Long, ByVal Name As String)
    Player(index).Name = Name
End Sub

Function GetPlayerClass(ByVal index As Long) As Long
    GetPlayerClass = Player(index).Class
End Function

Sub SetPlayerClass(ByVal index As Long, ByVal ClassNum As Long)
    Player(index).Class = ClassNum
End Sub

Function GetPlayerSprite(ByVal index As Long) As Long

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerSprite = Player(index).Sprite
End Function

Sub SetPlayerSprite(ByVal index As Long, ByVal Sprite As Long)
    Player(index).Sprite = Sprite
End Sub

Function GetPlayerLevel(ByVal index As Long) As Long
    If index <= 0 Or index > MAX_PLAYERS Then Exit Function
    GetPlayerLevel = Player(index).Level
End Function

Function SetPlayerLevel(ByVal index As Long, ByVal Level As Long) As Boolean
    If index <= 0 Or index > MAX_PLAYERS Then Exit Function
    SetPlayerLevel = False
    If Level > MAX_LEVELS Then
        Player(index).Level = MAX_LEVELS
        Exit Function
    End If
    Player(index).Level = Level
    SetPlayerLevel = True
End Function

Function GetPlayerNextLevel(ByVal index As Long) As Long
    GetPlayerNextLevel = 100 + (((GetPlayerLevel(index) ^ 2) * 10) * 2)
End Function

Function GetPlayerExp(ByVal index As Long) As Long
    If index <= 0 Or index > MAX_PLAYERS Then Exit Function
    GetPlayerExp = Player(index).exp
End Function

Sub SetPlayerExp(ByVal index As Long, ByVal exp As Long)
    If index <= 0 Or index > MAX_PLAYERS Then Exit Sub
    Player(index).exp = exp
End Sub

Function GetPlayerAccess(ByVal index As Long) As Long

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerAccess = Player(index).Access
End Function

Sub SetPlayerAccess(ByVal index As Long, ByVal Access As Long)
    Player(index).Access = Access
End Sub

Function GetPlayerPK(ByVal index As Long) As Long

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerPK = Player(index).PK
End Function

Sub SetPlayerPK(ByVal index As Long, ByVal PK As Long)
    Player(index).PK = PK
End Sub

Function GetPlayerVital(ByVal index As Long, ByVal Vital As Vitals) As Long
    If index > MAX_PLAYERS Then Exit Function
    GetPlayerVital = Player(index).Vital(Vital)
End Function

Sub SetPlayerVital(ByVal index As Long, ByVal Vital As Vitals, ByVal Value As Long)
    Player(index).Vital(Vital) = Value

    If GetPlayerVital(index, Vital) > GetPlayerMaxVital(index, Vital) Then
        Player(index).Vital(Vital) = GetPlayerMaxVital(index, Vital)
    End If

    If GetPlayerVital(index, Vital) < 0 Then
        Player(index).Vital(Vital) = 0
    End If

End Sub

Public Function GetPlayerStat(ByVal index As Long, ByVal Stat As Stats) As Long
    Dim x As Long, i As Long
    If index > MAX_PLAYERS Then Exit Function
    
    x = Player(index).Stat(Stat)
    
    For i = 1 To Equipment.Equipment_Count - 1
        If Player(index).Equipment(i) > 0 Then
            If Item(Player(index).Equipment(i)).Add_Stat(Stat) > 0 Then
                x = x + Item(Player(index).Equipment(i)).Add_Stat(Stat)
            End If
        End If
    Next
    
    GetPlayerStat = x
End Function

Public Function GetPlayerRawStat(ByVal index As Long, ByVal Stat As Stats) As Long
    If index > MAX_PLAYERS Then Exit Function
    
    GetPlayerRawStat = Player(index).Stat(Stat)
End Function

Public Sub SetPlayerStat(ByVal index As Long, ByVal Stat As Stats, ByVal Value As Long)
    Player(index).Stat(Stat) = Value
End Sub

Function GetPlayerPOINTS(ByVal index As Long) As Long

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerPOINTS = Player(index).POINTS
End Function

Sub SetPlayerPOINTS(ByVal index As Long, ByVal POINTS As Long)
    If POINTS <= 0 Then POINTS = 0
    Player(index).POINTS = POINTS
End Sub

Function GetPlayerMap(ByVal index As Long) As Long

    If index <= 0 Or index > MAX_PLAYERS Then Exit Function
    GetPlayerMap = Player(index).Map
End Function

Sub SetPlayerMap(ByVal index As Long, ByVal mapnum As Long)

    If mapnum > 0 And mapnum <= MAX_MAPS Then
        Player(index).Map = mapnum
    End If

End Sub

Function GetPlayerX(ByVal index As Long) As Long
    If index <= 0 Or index > MAX_PLAYERS Then Exit Function
    GetPlayerX = Player(index).x
End Function

Sub SetPlayerX(ByVal index As Long, ByVal x As Long)
    If index <= 0 Or index > MAX_PLAYERS Then Exit Sub
    Player(index).x = x
End Sub

Function GetPlayerY(ByVal index As Long) As Long
    If index <= 0 Or index > MAX_PLAYERS Then Exit Function
    GetPlayerY = Player(index).y
End Function

Sub SetPlayerY(ByVal index As Long, ByVal y As Long)
    If index <= 0 Or index > MAX_PLAYERS Then Exit Sub
    Player(index).y = y
End Sub

Function GetPlayerDir(ByVal index As Long) As Long

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerDir = Player(index).dir
End Function

Sub SetPlayerDir(ByVal index As Long, ByVal dir As Long)
    Player(index).dir = dir
End Sub

Function GetPlayerIP(ByVal index As Long) As String

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerIP = frmServer.Socket(index).RemoteHostIP
End Function

Function GetPlayerInvItemNum(ByVal index As Long, ByVal invSlot As Long) As Long
    If index > MAX_PLAYERS Then Exit Function
    If invSlot = 0 Then Exit Function
    
    GetPlayerInvItemNum = Player(index).Inv(invSlot).Num
End Function

Sub SetPlayerInvItemNum(ByVal index As Long, ByVal invSlot As Long, ByVal itemNum As Long)
    Player(index).Inv(invSlot).Num = itemNum
End Sub

Function GetPlayerInvItemValue(ByVal index As Long, ByVal invSlot As Long) As Long

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerInvItemValue = Player(index).Inv(invSlot).Value
End Function

Sub SetPlayerInvItemValue(ByVal index As Long, ByVal invSlot As Long, ByVal ItemValue As Long)
    Player(index).Inv(invSlot).Value = ItemValue
End Sub

Function GetPlayerEquipment(ByVal index As Long, ByVal EquipmentSlot As Equipment) As Long

    If index > MAX_PLAYERS Then Exit Function
    If EquipmentSlot = 0 Then Exit Function
    
    GetPlayerEquipment = Player(index).Equipment(EquipmentSlot)
End Function

Sub SetPlayerEquipment(ByVal index As Long, ByVal invNum As Long, ByVal EquipmentSlot As Equipment)
    Player(index).Equipment(EquipmentSlot) = invNum
End Sub

' ToDo
Sub OnDeath(ByVal index As Long)
    Dim i As Long
    
    ' Set HP to nothing
    Call SetPlayerVital(index, Vitals.HP, 0)
    SendVital index, HP

    ' Drop all worn items
    For i = 1 To Equipment.Equipment_Count - 1
        If GetPlayerEquipment(index, i) > 0 Then
            PlayerMapDropItem index, GetPlayerEquipment(index, i), 0
        End If
    Next

    ' Warp player away
    Call SetPlayerDir(index, DIR_DOWN)
    
    With Map(GetPlayerMap(index)).MapData
        ' to the bootmap if it is set
        If .BootMap > 0 Then
            PlayerWarp index, .BootMap, .BootX, .BootY
        Else
            Call PlayerWarp(index, START_MAP, START_X, START_Y)
        End If
    End With
    
    ' clear all DoTs and HoTs
    For i = 1 To MAX_DOTS
        With TempPlayer(index).DoT(i)
            .Used = False
            .Spell = 0
            .Timer = 0
            .Caster = 0
            .StartTime = 0
        End With
        
        With TempPlayer(index).HoT(i)
            .Used = False
            .Spell = 0
            .Timer = 0
            .Caster = 0
            .StartTime = 0
        End With
    Next
    
    ' Clear spell casting
    TempPlayer(index).spellBuffer.Spell = 0
    TempPlayer(index).spellBuffer.Timer = 0
    TempPlayer(index).spellBuffer.target = 0
    TempPlayer(index).spellBuffer.tType = 0
    Call SendClearSpellBuffer(index)
    
    ' Restore vitals
    Call SetPlayerVital(index, Vitals.HP, GetPlayerMaxVital(index, Vitals.HP))
    Call SetPlayerVital(index, Vitals.MP, GetPlayerMaxVital(index, Vitals.MP))
    Call SendVital(index, Vitals.HP)
    Call SendVital(index, Vitals.MP)
    
    ' send vitals to party if in one
    If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index

    ' If the player the attacker killed was a pk then take it away
    If GetPlayerPK(index) = YES Then
        Call SetPlayerPK(index, NO)
        Call SendPlayerData(index)
    End If

End Sub

Sub CheckResource(ByVal index As Long, ByVal x As Long, ByVal y As Long)
    Dim Resource_num As Long
    Dim Resource_index As Long
    Dim rX As Long, rY As Long
    Dim i As Long
    Dim damage As Long
    
    If Map(GetPlayerMap(index)).TileData.Tile(x, y).Type = TILE_TYPE_RESOURCE Then
        Resource_num = 0
        Resource_index = Map(GetPlayerMap(index)).TileData.Tile(x, y).Data1

        ' Get the cache number
        For i = 0 To ResourceCache(GetPlayerMap(index)).Resource_Count

            If ResourceCache(GetPlayerMap(index)).ResourceData(i).x = x Then
                If ResourceCache(GetPlayerMap(index)).ResourceData(i).y = y Then
                    Resource_num = i
                End If
            End If

        Next

        If Resource_num > 0 Then
            If GetPlayerEquipment(index, Weapon) > 0 Then
                If Item(GetPlayerEquipment(index, Weapon)).Data3 = Resource(Resource_index).ToolRequired Then

                    ' inv space?
                    If Resource(Resource_index).ItemReward > 0 Then
                        If FindOpenInvSlot(index, Resource(Resource_index).ItemReward) = 0 Then
                            PlayerMsg index, "You have no inventory space.", BrightRed
                            Exit Sub
                        End If
                    End If

                    ' check if already cut down
                    If ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).ResourceState = 0 Then
                    
                        rX = ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).x
                        rY = ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).y
                        
                        damage = Item(GetPlayerEquipment(index, Weapon)).Data2
                    
                        ' check if damage is more than health
                        If damage > 0 Then
                            ' cut it down!
                            If ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).cur_health - damage <= 0 Then
                                SendActionMsg GetPlayerMap(index), "-" & ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).cur_health, BrightRed, 1, (rX * 32), (rY * 32)
                                ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).ResourceState = 1 ' Cut
                                ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).ResourceTimer = GetTickCount
                                SendResourceCacheToMap GetPlayerMap(index), Resource_num
                                ' send message if it exists
                                If Len(Trim$(Resource(Resource_index).SuccessMessage)) > 0 Then
                                    SendActionMsg GetPlayerMap(index), Trim$(Resource(Resource_index).SuccessMessage), BrightGreen, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
                                End If
                                ' carry on
                                GiveInvItem index, Resource(Resource_index).ItemReward, 1
                                SendAnimation GetPlayerMap(index), Resource(Resource_index).Animation, rX, rY
                            Else
                                ' just do the damage
                                ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).cur_health = ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).cur_health - damage
                                SendActionMsg GetPlayerMap(index), "-" & damage, BrightRed, 1, (rX * 32), (rY * 32)
                                SendAnimation GetPlayerMap(index), Resource(Resource_index).Animation, rX, rY
                            End If
                            ' send the sound
                            SendMapSound index, rX, rY, SoundEntity.seResource, Resource_index
                        Else
                            ' too weak
                            SendActionMsg GetPlayerMap(index), "Miss!", BrightRed, 1, (rX * 32), (rY * 32)
                        End If
                    Else
                        ' send message if it exists
                        If Len(Trim$(Resource(Resource_index).EmptyMessage)) > 0 Then
                            SendActionMsg GetPlayerMap(index), Trim$(Resource(Resource_index).EmptyMessage), BrightRed, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
                        End If
                    End If

                Else
                    PlayerMsg index, "You have the wrong type of tool equiped.", BrightRed
                End If

            Else
                PlayerMsg index, "You need a tool to interact with this resource.", BrightRed
            End If
        End If
    End If
End Sub

Function GetPlayerBankItemNum(ByVal index As Long, ByVal BankSlot As Long) As Long
    If BankSlot = 0 Then Exit Function
    GetPlayerBankItemNum = Player(index).Bank(BankSlot).Num
End Function

Sub SetPlayerBankItemNum(ByVal index As Long, ByVal BankSlot As Long, ByVal itemNum As Long)
    If BankSlot = 0 Then Exit Sub
    Player(index).Bank(BankSlot).Num = itemNum
End Sub

Function GetPlayerBankItemValue(ByVal index As Long, ByVal BankSlot As Long) As Long
    If BankSlot = 0 Then Exit Function
    GetPlayerBankItemValue = Player(index).Bank(BankSlot).Value
End Function

Sub SetPlayerBankItemValue(ByVal index As Long, ByVal BankSlot As Long, ByVal ItemValue As Long)
    If BankSlot = 0 Then Exit Sub
    Player(index).Bank(BankSlot).Value = ItemValue
End Sub

Sub GiveBankItem(ByVal index As Long, ByVal invSlot As Long, ByVal amount As Long)
Dim BankSlot

    If invSlot < 0 Or invSlot > MAX_INV Then
        Exit Sub
    End If
    
    If amount < 0 Or amount > GetPlayerInvItemValue(index, invSlot) Then
        Exit Sub
    End If
    
    BankSlot = FindOpenBankSlot(index, GetPlayerInvItemNum(index, invSlot))
        
    If BankSlot > 0 Then
        If Item(GetPlayerInvItemNum(index, invSlot)).Type = ITEM_TYPE_CURRENCY Then
            If GetPlayerBankItemNum(index, BankSlot) = GetPlayerInvItemNum(index, invSlot) Then
                Call SetPlayerBankItemValue(index, BankSlot, GetPlayerBankItemValue(index, BankSlot) + amount)
                Call TakeInvItem(index, GetPlayerInvItemNum(index, invSlot), amount)
            Else
                Call SetPlayerBankItemNum(index, BankSlot, GetPlayerInvItemNum(index, invSlot))
                Call SetPlayerBankItemValue(index, BankSlot, amount)
                Call TakeInvItem(index, GetPlayerInvItemNum(index, invSlot), amount)
            End If
        Else
            If GetPlayerBankItemNum(index, BankSlot) = GetPlayerInvItemNum(index, invSlot) Then
                Call SetPlayerBankItemValue(index, BankSlot, GetPlayerBankItemValue(index, BankSlot) + 1)
                Call TakeInvItem(index, GetPlayerInvItemNum(index, invSlot), 0)
            Else
                Call SetPlayerBankItemNum(index, BankSlot, GetPlayerInvItemNum(index, invSlot))
                Call SetPlayerBankItemValue(index, BankSlot, 1)
                Call TakeInvItem(index, GetPlayerInvItemNum(index, invSlot), 0)
            End If
        End If
    End If
    
    SavePlayer index
    SendBank index

End Sub

Sub TakeBankItem(ByVal index As Long, ByVal BankSlot As Long, ByVal amount As Long)
Dim invSlot

    If BankSlot < 0 Or BankSlot > MAX_BANK Then
        Exit Sub
    End If
    
    If amount < 0 Or amount > GetPlayerBankItemValue(index, BankSlot) Then
        Exit Sub
    End If
    
    invSlot = FindOpenInvSlot(index, GetPlayerBankItemNum(index, BankSlot))
        
    If invSlot > 0 Then
        If Item(GetPlayerBankItemNum(index, BankSlot)).Type = ITEM_TYPE_CURRENCY Then
            Call GiveInvItem(index, GetPlayerBankItemNum(index, BankSlot), amount)
            Call SetPlayerBankItemValue(index, BankSlot, GetPlayerBankItemValue(index, BankSlot) - amount)
            If GetPlayerBankItemValue(index, BankSlot) <= 0 Then
                Call SetPlayerBankItemNum(index, BankSlot, 0)
                Call SetPlayerBankItemValue(index, BankSlot, 0)
            End If
        Else
            If GetPlayerBankItemValue(index, BankSlot) > 1 Then
                Call GiveInvItem(index, GetPlayerBankItemNum(index, BankSlot), 0)
                Call SetPlayerBankItemValue(index, BankSlot, GetPlayerBankItemValue(index, BankSlot) - 1)
            Else
                Call GiveInvItem(index, GetPlayerBankItemNum(index, BankSlot), 0)
                Call SetPlayerBankItemNum(index, BankSlot, 0)
                Call SetPlayerBankItemValue(index, BankSlot, 0)
            End If
        End If
    End If
    
    SavePlayer index
    SendBank index

End Sub

Public Sub KillPlayer(ByVal index As Long)
Dim exp As Long

    ' Calculate exp to give attacker
    exp = GetPlayerExp(index) \ 3

    ' Make sure we dont get less then 0
    If exp < 0 Then exp = 0
    If exp = 0 Then
        Call PlayerMsg(index, "You lost no exp.", BrightRed)
    Else
        Call SetPlayerExp(index, GetPlayerExp(index) - exp)
        SendEXP index
        Call PlayerMsg(index, "You lost " & exp & " exp.", BrightRed)
    End If
    
    Call OnDeath(index)
End Sub

Public Sub UseItem(ByVal index As Long, ByVal invNum As Long)
Dim n As Long, i As Long, tempItem As Long, x As Long, y As Long, itemNum As Long

    ' Prevent hacking
    If invNum < 1 Or invNum > MAX_ITEMS Then
        Exit Sub
    End If

    If (GetPlayerInvItemNum(index, invNum) > 0) And (GetPlayerInvItemNum(index, invNum) <= MAX_ITEMS) Then
        n = Item(GetPlayerInvItemNum(index, invNum)).Data2
        itemNum = GetPlayerInvItemNum(index, invNum)
        
        ' Find out what kind of item it is
        Select Case Item(itemNum).Type
            Case ITEM_TYPE_ARMOR
            
                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(index, i) < Item(itemNum).Stat_Req(i) Then
                        PlayerMsg index, "You do not meet the stat requirements to equip this item.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(index) < Item(itemNum).LevelReq Then
                    PlayerMsg index, "You do not meet the level requirement to equip this item.", BrightRed
                    Exit Sub
                End If
                
                ' class requirement
                If Item(itemNum).ClassReq > 0 Then
                    If Not GetPlayerClass(index) = Item(itemNum).ClassReq Then
                        PlayerMsg index, "You do not meet the class requirement to equip this item.", BrightRed
                        Exit Sub
                    End If
                End If
                
                ' access requirement
                If Not GetPlayerAccess(index) >= Item(itemNum).AccessReq Then
                    PlayerMsg index, "You do not meet the access requirement to equip this item.", BrightRed
                    Exit Sub
                End If
                
                ' prociency requirement
                If Not hasProficiency(index, Item(itemNum).proficiency) Then
                    PlayerMsg index, "You do not have the proficiency this item requires.", BrightRed
                    Exit Sub
                End If

                If GetPlayerEquipment(index, Armor) > 0 Then
                    tempItem = GetPlayerEquipment(index, Armor)
                End If

                SetPlayerEquipment index, itemNum, Armor
                
                PlayerMsg index, "You equip " & CheckGrammar(Item(itemNum).Name), BrightGreen
                
                ' tell them if it's soulbound
                If Item(itemNum).BindType = 2 Then ' BoE
                    If Player(index).Inv(invNum).Bound = 0 Then
                        PlayerMsg index, "This item is now bound to your soul.", BrightRed
                    End If
                End If
                
                TakeInvItem index, itemNum, 0

                If tempItem > 0 Then
                    If Item(tempItem).BindType > 0 Then
                        GiveInvItem index, tempItem, 0, , True ' give back the stored item
                        tempItem = 0
                    Else
                        GiveInvItem index, tempItem, 0
                        tempItem = 0
                    End If
                End If

                Call SendWornEquipment(index)
                Call SendMapEquipment(index)
                
                ' send vitals
                Call SendVital(index, Vitals.HP)
                Call SendVital(index, Vitals.MP)
                ' send vitals to party if in one
                If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
                
                ' send the sound
                SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, itemNum
            Case ITEM_TYPE_WEAPON
            
                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(index, i) < Item(itemNum).Stat_Req(i) Then
                        PlayerMsg index, "You do not meet the stat requirements to equip this item.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(index) < Item(itemNum).LevelReq Then
                    PlayerMsg index, "You do not meet the level requirement to equip this item.", BrightRed
                    Exit Sub
                End If
                
                ' class requirement
                If Item(itemNum).ClassReq > 0 Then
                    If Not GetPlayerClass(index) = Item(itemNum).ClassReq Then
                        PlayerMsg index, "You do not meet the class requirement to equip this item.", BrightRed
                        Exit Sub
                    End If
                End If
                
                ' access requirement
                If Not GetPlayerAccess(index) >= Item(itemNum).AccessReq Then
                    PlayerMsg index, "You do not meet the access requirement to equip this item.", BrightRed
                    Exit Sub
                End If
                
                ' prociency requirement
                If Not hasProficiency(index, Item(itemNum).proficiency) Then
                    PlayerMsg index, "You do not have the proficiency this item requires.", BrightRed
                    Exit Sub
                End If

                If GetPlayerEquipment(index, Weapon) > 0 Then
                    tempItem = GetPlayerEquipment(index, Weapon)
                End If

                SetPlayerEquipment index, itemNum, Weapon
                PlayerMsg index, "You equip " & CheckGrammar(Item(itemNum).Name), BrightGreen
                
                ' tell them if it's soulbound
                If Item(itemNum).BindType = 2 Then ' BoE
                    If Player(index).Inv(invNum).Bound = 0 Then
                        PlayerMsg index, "This item is now bound to your soul.", BrightRed
                    End If
                End If
                
                TakeInvItem index, itemNum, 1
                
                If tempItem > 0 Then
                    If Item(tempItem).BindType > 0 Then
                        GiveInvItem index, tempItem, 0, , True ' give back the stored item
                        tempItem = 0
                    Else
                        GiveInvItem index, tempItem, 0
                        tempItem = 0
                    End If
                End If

                Call SendWornEquipment(index)
                Call SendMapEquipment(index)
                
                ' send vitals
                Call SendVital(index, Vitals.HP)
                Call SendVital(index, Vitals.MP)
                ' send vitals to party if in one
                If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
                
                ' send the sound
                SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, itemNum
            Case ITEM_TYPE_HELMET
            
                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(index, i) < Item(itemNum).Stat_Req(i) Then
                        PlayerMsg index, "You do not meet the stat requirements to equip this item.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(index) < Item(itemNum).LevelReq Then
                    PlayerMsg index, "You do not meet the level requirement to equip this item.", BrightRed
                    Exit Sub
                End If
                
                ' class requirement
                If Item(itemNum).ClassReq > 0 Then
                    If Not GetPlayerClass(index) = Item(itemNum).ClassReq Then
                        PlayerMsg index, "You do not meet the class requirement to equip this item.", BrightRed
                        Exit Sub
                    End If
                End If
                
                ' access requirement
                If Not GetPlayerAccess(index) >= Item(itemNum).AccessReq Then
                    PlayerMsg index, "You do not meet the access requirement to equip this item.", BrightRed
                    Exit Sub
                End If
                
                ' prociency requirement
                If Not hasProficiency(index, Item(itemNum).proficiency) Then
                    PlayerMsg index, "You do not have the proficiency this item requires.", BrightRed
                    Exit Sub
                End If

                If GetPlayerEquipment(index, Helmet) > 0 Then
                    tempItem = GetPlayerEquipment(index, Helmet)
                End If

                SetPlayerEquipment index, itemNum, Helmet
                PlayerMsg index, "You equip " & CheckGrammar(Item(itemNum).Name), BrightGreen
                
                ' tell them if it's soulbound
                If Item(itemNum).BindType = 2 Then ' BoE
                    If Player(index).Inv(invNum).Bound = 0 Then
                        PlayerMsg index, "This item is now bound to your soul.", BrightRed
                    End If
                End If
                
                TakeInvItem index, itemNum, 1

                If tempItem > 0 Then
                    If Item(tempItem).BindType > 0 Then
                        GiveInvItem index, tempItem, 0, , True ' give back the stored item
                        tempItem = 0
                    Else
                        GiveInvItem index, tempItem, 0
                        tempItem = 0
                    End If
                End If

                Call SendWornEquipment(index)
                Call SendMapEquipment(index)
                
                ' send vitals
                Call SendVital(index, Vitals.HP)
                Call SendVital(index, Vitals.MP)
                ' send vitals to party if in one
                If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
                
                ' send the sound
                SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, itemNum
            Case ITEM_TYPE_SHIELD
            
                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(index, i) < Item(itemNum).Stat_Req(i) Then
                        PlayerMsg index, "You do not meet the stat requirements to equip this item.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(index) < Item(itemNum).LevelReq Then
                    PlayerMsg index, "You do not meet the level requirement to equip this item.", BrightRed
                    Exit Sub
                End If
                
                ' class requirement
                If Item(itemNum).ClassReq > 0 Then
                    If Not GetPlayerClass(index) = Item(itemNum).ClassReq Then
                        PlayerMsg index, "You do not meet the class requirement to equip this item.", BrightRed
                        Exit Sub
                    End If
                End If
                
                ' access requirement
                If Not GetPlayerAccess(index) >= Item(itemNum).AccessReq Then
                    PlayerMsg index, "You do not meet the access requirement to equip this item.", BrightRed
                    Exit Sub
                End If
                
                ' prociency requirement
                If Not hasProficiency(index, Item(itemNum).proficiency) Then
                    PlayerMsg index, "You do not have the proficiency this item requires.", BrightRed
                    Exit Sub
                End If

                If GetPlayerEquipment(index, Shield) > 0 Then
                    tempItem = GetPlayerEquipment(index, Shield)
                End If

                SetPlayerEquipment index, itemNum, Shield
                PlayerMsg index, "You equip " & CheckGrammar(Item(itemNum).Name), BrightGreen
                
                ' tell them if it's soulbound
                If Item(itemNum).BindType = 2 Then ' BoE
                    If Player(index).Inv(invNum).Bound = 0 Then
                        PlayerMsg index, "This item is now bound to your soul.", BrightRed
                    End If
                End If
                
                TakeInvItem index, itemNum, 1

                If tempItem > 0 Then
                    If Item(tempItem).BindType > 0 Then
                        GiveInvItem index, tempItem, 0, , True ' give back the stored item
                        tempItem = 0
                    Else
                        GiveInvItem index, tempItem, 0
                        tempItem = 0
                    End If
                End If
                
                ' send vitals
                Call SendVital(index, Vitals.HP)
                Call SendVital(index, Vitals.MP)
                ' send vitals to party if in one
                If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index

                Call SendWornEquipment(index)
                Call SendMapEquipment(index)
                
                ' send the sound
                SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, itemNum
            ' consumable
            Case ITEM_TYPE_CONSUME
                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(index, i) < Item(itemNum).Stat_Req(i) Then
                        PlayerMsg index, "You do not meet the stat requirements to use this item.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(index) < Item(itemNum).LevelReq Then
                    PlayerMsg index, "You do not meet the level requirement to use this item.", BrightRed
                    Exit Sub
                End If
                
                ' class requirement
                If Item(itemNum).ClassReq > 0 Then
                    If Not GetPlayerClass(index) = Item(itemNum).ClassReq Then
                        PlayerMsg index, "You do not meet the class requirement to use this item.", BrightRed
                        Exit Sub
                    End If
                End If
                
                ' access requirement
                If Not GetPlayerAccess(index) >= Item(itemNum).AccessReq Then
                    PlayerMsg index, "You do not meet the access requirement to use this item.", BrightRed
                    Exit Sub
                End If
                
                ' add hp
                If Item(itemNum).AddHP > 0 Then
                    Player(index).Vital(Vitals.HP) = Player(index).Vital(Vitals.HP) + Item(itemNum).AddHP
                    SendActionMsg GetPlayerMap(index), "+" & Item(itemNum).AddHP, BrightGreen, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
                    SendVital index, HP
                    ' send vitals to party if in one
                    If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
                End If
                ' add mp
                If Item(itemNum).AddMP > 0 Then
                    Player(index).Vital(Vitals.MP) = Player(index).Vital(Vitals.MP) + Item(itemNum).AddMP
                    SendActionMsg GetPlayerMap(index), "+" & Item(itemNum).AddMP, BrightBlue, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
                    SendVital index, MP
                    ' send vitals to party if in one
                    If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
                End If
                ' add exp
                If Item(itemNum).AddEXP > 0 Then
                    SetPlayerExp index, GetPlayerExp(index) + Item(itemNum).AddEXP
                    CheckPlayerLevelUp index
                    SendActionMsg GetPlayerMap(index), "+" & Item(itemNum).AddEXP & " EXP", White, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
                    SendEXP index
                End If
                Call SendAnimation(GetPlayerMap(index), Item(itemNum).Animation, 0, 0, TARGET_TYPE_PLAYER, index)
                Call TakeInvItem(index, Player(index).Inv(invNum).Num, 0)
                
                ' send the sound
                SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, itemNum
            Case ITEM_TYPE_KEY
                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(index, i) < Item(itemNum).Stat_Req(i) Then
                        PlayerMsg index, "You do not meet the stat requirements to use this item.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(index) < Item(itemNum).LevelReq Then
                    PlayerMsg index, "You do not meet the level requirement to use this item.", BrightRed
                    Exit Sub
                End If
                
                ' class requirement
                If Item(itemNum).ClassReq > 0 Then
                    If Not GetPlayerClass(index) = Item(itemNum).ClassReq Then
                        PlayerMsg index, "You do not meet the class requirement to use this item.", BrightRed
                        Exit Sub
                    End If
                End If
                
                ' access requirement
                If Not GetPlayerAccess(index) >= Item(itemNum).AccessReq Then
                    PlayerMsg index, "You do not meet the access requirement to use this item.", BrightRed
                    Exit Sub
                End If

                Select Case GetPlayerDir(index)
                    Case DIR_UP

                        If GetPlayerY(index) > 0 Then
                            x = GetPlayerX(index)
                            y = GetPlayerY(index) - 1
                        Else
                            Exit Sub
                        End If

                    Case DIR_DOWN

                        If GetPlayerY(index) < Map(GetPlayerMap(index)).MapData.MaxY Then
                            x = GetPlayerX(index)
                            y = GetPlayerY(index) + 1
                        Else
                            Exit Sub
                        End If

                    Case DIR_LEFT

                        If GetPlayerX(index) > 0 Then
                            x = GetPlayerX(index) - 1
                            y = GetPlayerY(index)
                        Else
                            Exit Sub
                        End If

                    Case DIR_RIGHT

                        If GetPlayerX(index) < Map(GetPlayerMap(index)).MapData.MaxX Then
                            x = GetPlayerX(index) + 1
                            y = GetPlayerY(index)
                        Else
                            Exit Sub
                        End If

                End Select

                ' Check if a key exists
                If Map(GetPlayerMap(index)).TileData.Tile(x, y).Type = TILE_TYPE_KEY Then

                    ' Check if the key they are using matches the map key
                    If itemNum = Map(GetPlayerMap(index)).TileData.Tile(x, y).Data1 Then
                        TempTile(GetPlayerMap(index)).DoorOpen(x, y) = YES
                        TempTile(GetPlayerMap(index)).DoorTimer = GetTickCount
                        SendMapKey index, x, y, 1
                        'Call MapMsg(GetPlayerMap(index), "A door has been unlocked.", White)
                        
                        Call SendAnimation(GetPlayerMap(index), Item(itemNum).Animation, x, y)

                        ' Check if we are supposed to take away the item
                        If Map(GetPlayerMap(index)).TileData.Tile(x, y).Data2 = 1 Then
                            Call TakeInvItem(index, itemNum, 0)
                            Call PlayerMsg(index, "The key is destroyed in the lock.", Yellow)
                        End If
                    End If
                End If
                
                ' send the sound
                SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, itemNum
            Case ITEM_TYPE_UNIQUE
                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(index, i) < Item(itemNum).Stat_Req(i) Then
                        PlayerMsg index, "You do not meet the stat requirements to use this item.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(index) < Item(itemNum).LevelReq Then
                    PlayerMsg index, "You do not meet the level requirement to use this item.", BrightRed
                    Exit Sub
                End If
                
                ' class requirement
                If Item(itemNum).ClassReq > 0 Then
                    If Not GetPlayerClass(index) = Item(itemNum).ClassReq Then
                        PlayerMsg index, "You do not meet the class requirement to use this item.", BrightRed
                        Exit Sub
                    End If
                End If
                
                ' access requirement
                If Not GetPlayerAccess(index) >= Item(itemNum).AccessReq Then
                    PlayerMsg index, "You do not meet the access requirement to use this item.", BrightRed
                    Exit Sub
                End If
                
                ' Go through with it
                Unique_Item index, itemNum
            Case ITEM_TYPE_SPELL
                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(index, i) < Item(itemNum).Stat_Req(i) Then
                        PlayerMsg index, "You do not meet the stat requirements to use this item.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(index) < Item(itemNum).LevelReq Then
                    PlayerMsg index, "You do not meet the level requirement to use this item.", BrightRed
                    Exit Sub
                End If
                
                ' class requirement
                If Item(itemNum).ClassReq > 0 Then
                    If Not GetPlayerClass(index) = Item(itemNum).ClassReq Then
                        PlayerMsg index, "You do not meet the class requirement to use this item.", BrightRed
                        Exit Sub
                    End If
                End If
                
                ' access requirement
                If Not GetPlayerAccess(index) >= Item(itemNum).AccessReq Then
                    PlayerMsg index, "You do not meet the access requirement to use this item.", BrightRed
                    Exit Sub
                End If
                
                ' Get the spell num
                n = Item(itemNum).Data1

                If n > 0 Then

                    ' Make sure they are the right class
                    If Spell(n).ClassReq = GetPlayerClass(index) Or Spell(n).ClassReq = 0 Then
                    
                        ' make sure they don't already know it
                        For i = 1 To MAX_PLAYER_SPELLS
                            If Player(index).Spell(i).Spell > 0 Then
                                If Player(index).Spell(i).Spell = n Then
                                    PlayerMsg index, "You already know this spell.", BrightRed
                                    Exit Sub
                                End If
                                If Spell(Player(index).Spell(i).Spell).UniqueIndex = Spell(n).UniqueIndex Then
                                    PlayerMsg index, "You already know this spell.", BrightRed
                                    Exit Sub
                                End If
                            End If
                        Next
                    
                        ' Make sure they are the right level
                        i = Spell(n).LevelReq


                        If i <= GetPlayerLevel(index) Then
                            i = FindOpenSpellSlot(index)

                            ' Make sure they have an open spell slot
                            If i > 0 Then

                                ' Make sure they dont already have the spell
                                If Not HasSpell(index, n) Then
                                    Player(index).Spell(i).Spell = n
                                    Call SendAnimation(GetPlayerMap(index), Item(itemNum).Animation, 0, 0, TARGET_TYPE_PLAYER, index)
                                    Call TakeInvItem(index, itemNum, 0)
                                    Call PlayerMsg(index, "You feel the rush of knowledge fill your mind. You can now use " & Trim$(Spell(n).Name) & ".", BrightGreen)
                                    SendPlayerSpells index
                                Else
                                    Call PlayerMsg(index, "You already have knowledge of this skill.", BrightRed)
                                End If

                            Else
                                Call PlayerMsg(index, "You cannot learn any more skills.", BrightRed)
                            End If

                        Else
                            Call PlayerMsg(index, "You must be level " & i & " to learn this skill.", BrightRed)
                        End If

                    Else
                        Call PlayerMsg(index, "This spell can only be learned by " & CheckGrammar(GetClassName(Spell(n).ClassReq)) & ".", BrightRed)
                    End If
                End If
                
                ' send the sound
                SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, itemNum
            Case ITEM_TYPE_FOOD
                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(index, i) < Item(itemNum).Stat_Req(i) Then
                        PlayerMsg index, "You do not meet the stat requirements to use this item.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(index) < Item(itemNum).LevelReq Then
                    PlayerMsg index, "You do not meet the level requirement to use this item.", BrightRed
                    Exit Sub
                End If
                
                ' class requirement
                If Item(itemNum).ClassReq > 0 Then
                    If Not GetPlayerClass(index) = Item(itemNum).ClassReq Then
                        PlayerMsg index, "You do not meet the class requirement to use this item.", BrightRed
                        Exit Sub
                    End If
                End If
                
                ' access requirement
                If Not GetPlayerAccess(index) >= Item(itemNum).AccessReq Then
                    PlayerMsg index, "You do not meet the access requirement to use this item.", BrightRed
                    Exit Sub
                End If
                
                ' make sure they're not in combat
                If TempPlayer(index).stopRegen Then
                    PlayerMsg index, "You cannot eat whilst in combat.", BrightRed
                    Exit Sub
                End If
                
                ' make sure not full hp
                x = Item(itemNum).HPorSP
                If Player(index).Vital(x) >= GetPlayerMaxVital(index, x) Then
                    PlayerMsg index, "You don't need to eat this at the moment.", BrightRed
                    Exit Sub
                End If
                
                ' set the player's food
                If Item(itemNum).HPorSP = 2 Then 'mp
                    If Not TempPlayer(index).foodItem(Vitals.MP) = itemNum Then
                        TempPlayer(index).foodItem(Vitals.MP) = itemNum
                        TempPlayer(index).foodTick(Vitals.MP) = 0
                        TempPlayer(index).foodTimer(Vitals.MP) = GetTickCount
                    Else
                        PlayerMsg index, "You are already eating this.", BrightRed
                        Exit Sub
                    End If
                Else ' hp
                    If Not TempPlayer(index).foodItem(Vitals.HP) = itemNum Then
                        TempPlayer(index).foodItem(Vitals.HP) = itemNum
                        TempPlayer(index).foodTick(Vitals.HP) = 0
                        TempPlayer(index).foodTimer(Vitals.HP) = GetTickCount
                    Else
                        PlayerMsg index, "You are already eating this.", BrightRed
                        Exit Sub
                    End If
                End If
                
                ' take the item
                Call TakeInvItem(index, Player(index).Inv(invNum).Num, 0)
        End Select
    End If
End Sub
