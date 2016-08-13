Attribute VB_Name = "modGameLogic"
Option Explicit

Function FindOpenPlayerSlot() As Long
    Dim i As Long
    FindOpenPlayerSlot = 0

    For i = 1 To MAX_PLAYERS

        If Not IsConnected(i) Then
            FindOpenPlayerSlot = i
            Exit Function
        End If

    Next

End Function

Function FindOpenMapItemSlot(ByVal mapnum As Long) As Long
    Dim i As Long
    FindOpenMapItemSlot = 0

    ' Check for subscript out of range
    If mapnum <= 0 Or mapnum > MAX_MAPS Then
        Exit Function
    End If

    For i = 1 To MAX_MAP_ITEMS

        If MapItem(mapnum, i).Num = 0 Then
            FindOpenMapItemSlot = i
            Exit Function
        End If

    Next

End Function

Function TotalOnlinePlayers() As Long
    Dim i As Long
    TotalOnlinePlayers = 0

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            TotalOnlinePlayers = TotalOnlinePlayers + 1
        End If

    Next

End Function

Function FindPlayer(ByVal Name As String) As Long
    Dim i As Long

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then

            ' Make sure we dont try to check a name thats to small
            If Len(GetPlayerName(i)) >= Len(Trim$(Name)) Then
                If UCase$(Mid$(GetPlayerName(i), 1, Len(Trim$(Name)))) = UCase$(Trim$(Name)) Then
                    FindPlayer = i
                    Exit Function
                End If
            End If
        End If

    Next

    FindPlayer = 0
End Function

Sub SpawnItem(ByVal itemNum As Long, ByVal ItemVal As Long, ByVal mapnum As Long, ByVal x As Long, ByVal y As Long, Optional ByVal playerName As String = vbNullString)
    Dim i As Long

    ' Check for subscript out of range
    If itemNum < 1 Or itemNum > MAX_ITEMS Or mapnum <= 0 Or mapnum > MAX_MAPS Then
        Exit Sub
    End If

    ' Find open map item slot
    i = FindOpenMapItemSlot(mapnum)
    Call SpawnItemSlot(i, itemNum, ItemVal, mapnum, x, y, playerName)
End Sub

Sub SpawnItemSlot(ByVal MapItemSlot As Long, ByVal itemNum As Long, ByVal ItemVal As Long, ByVal mapnum As Long, ByVal x As Long, ByVal y As Long, Optional ByVal playerName As String = vbNullString, Optional ByVal canDespawn As Boolean = True, Optional ByVal isSB As Boolean = False)
    Dim packet As String
    Dim i As Long
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If MapItemSlot <= 0 Or MapItemSlot > MAX_MAP_ITEMS Or itemNum < 0 Or itemNum > MAX_ITEMS Or mapnum <= 0 Or mapnum > MAX_MAPS Then
        Exit Sub
    End If

    i = MapItemSlot

    If i <> 0 Then
        If itemNum >= 0 And itemNum <= MAX_ITEMS Then
            MapItem(mapnum, i).playerName = playerName
            MapItem(mapnum, i).playerTimer = GetTickCount + ITEM_SPAWN_TIME
            MapItem(mapnum, i).canDespawn = canDespawn
            MapItem(mapnum, i).despawnTimer = GetTickCount + ITEM_DESPAWN_TIME
            MapItem(mapnum, i).Num = itemNum
            MapItem(mapnum, i).Value = ItemVal
            MapItem(mapnum, i).x = x
            MapItem(mapnum, i).y = y
            MapItem(mapnum, i).Bound = isSB
            ' send to map
            SendSpawnItemToMap mapnum, i
        End If
    End If

End Sub

Sub SpawnAllMapsItems()
    Dim i As Long

    For i = 1 To MAX_MAPS
        Call SpawnMapItems(i)
    Next

End Sub

Sub SpawnMapItems(ByVal mapnum As Long)
    Dim x As Long
    Dim y As Long

    ' Check for subscript out of range
    If mapnum <= 0 Or mapnum > MAX_MAPS Then
        Exit Sub
    End If

    ' Spawn what we have
    For x = 0 To Map(mapnum).MapData.MaxX
        For y = 0 To Map(mapnum).MapData.MaxY

            ' Check if the tile type is an item or a saved tile incase someone drops something
            If (Map(mapnum).TileData.Tile(x, y).Type = TILE_TYPE_ITEM) Then

                ' Check to see if its a currency and if they set the value to 0 set it to 1 automatically
                If Item(Map(mapnum).TileData.Tile(x, y).Data1).Type = ITEM_TYPE_CURRENCY And Map(mapnum).TileData.Tile(x, y).Data2 <= 0 Then
                    Call SpawnItem(Map(mapnum).TileData.Tile(x, y).Data1, 1, mapnum, x, y)
                Else
                    Call SpawnItem(Map(mapnum).TileData.Tile(x, y).Data1, Map(mapnum).TileData.Tile(x, y).Data2, mapnum, x, y)
                End If
            End If

        Next
    Next

End Sub

Function Random(ByVal Low As Long, ByVal High As Long) As Long
    Random = ((High - Low + 1) * Rnd) + Low
End Function

Public Sub SpawnNpc(ByVal mapNpcNum As Long, ByVal mapnum As Long)
    Dim Buffer As clsBuffer
    Dim npcNum As Long
    Dim i As Long
    Dim x As Long
    Dim y As Long
    Dim Spawned As Boolean

    ' Check for subscript out of range
    If mapNpcNum <= 0 Or mapNpcNum > MAX_MAP_NPCS Or mapnum <= 0 Or mapnum > MAX_MAPS Then Exit Sub
    npcNum = Map(mapnum).MapData.Npc(mapNpcNum)

    If npcNum > 0 Then
    
        With MapNpc(mapnum).Npc(mapNpcNum)
            .Num = npcNum
            .target = 0
            .targetType = 0 ' clear
            .Vital(Vitals.HP) = GetNpcMaxVital(npcNum, Vitals.HP)
            .Vital(Vitals.MP) = GetNpcMaxVital(npcNum, Vitals.MP)
            .dir = Int(Rnd * 4)
            .spellBuffer.Spell = 0
            .spellBuffer.Timer = 0
            .spellBuffer.target = 0
            .spellBuffer.tType = 0
        
            'Check if theres a spawn tile for the specific npc
            For x = 0 To Map(mapnum).MapData.MaxX
                For y = 0 To Map(mapnum).MapData.MaxY
                    If Map(mapnum).TileData.Tile(x, y).Type = TILE_TYPE_NPCSPAWN Then
                        If Map(mapnum).TileData.Tile(x, y).Data1 = mapNpcNum Then
                            .x = x
                            .y = y
                            .dir = Map(mapnum).TileData.Tile(x, y).Data2
                            Spawned = True
                            Exit For
                        End If
                    End If
                Next y
            Next x
            
            If Not Spawned Then
        
                ' Well try 100 times to randomly place the sprite
                For i = 1 To 100
                    x = Random(0, Map(mapnum).MapData.MaxX)
                    y = Random(0, Map(mapnum).MapData.MaxY)
        
                    If x > Map(mapnum).MapData.MaxX Then x = Map(mapnum).MapData.MaxX
                    If y > Map(mapnum).MapData.MaxY Then y = Map(mapnum).MapData.MaxY
        
                    ' Check if the tile is walkable
                    If NpcTileIsOpen(mapnum, x, y) Then
                        .x = x
                        .y = y
                        Spawned = True
                        Exit For
                    End If
        
                Next
                
            End If
    
            ' Didn't spawn, so now we'll just try to find a free tile
            If Not Spawned Then
    
                For x = 0 To Map(mapnum).MapData.MaxX
                    For y = 0 To Map(mapnum).MapData.MaxY
    
                        If NpcTileIsOpen(mapnum, x, y) Then
                            .x = x
                            .y = y
                            Spawned = True
                        End If
    
                    Next
                Next
    
            End If
    
            ' If we suceeded in spawning then send it to everyone
            If Spawned Then
                Set Buffer = New clsBuffer
                Buffer.WriteLong SSpawnNpc
                Buffer.WriteLong mapNpcNum
                Buffer.WriteLong .Num
                Buffer.WriteLong .x
                Buffer.WriteLong .y
                Buffer.WriteLong .dir
                SendDataToMap mapnum, Buffer.ToArray()
                Set Buffer = Nothing
            End If
            
            SendMapNpcVitals mapnum, mapNpcNum
        End With
    End If
End Sub

Public Function NpcTileIsOpen(ByVal mapnum As Long, ByVal x As Long, ByVal y As Long) As Boolean
    Dim LoopI As Long
    NpcTileIsOpen = True

    If PlayersOnMap(mapnum) Then

        For LoopI = 1 To Player_HighIndex

            If GetPlayerMap(LoopI) = mapnum Then
                If GetPlayerX(LoopI) = x Then
                    If GetPlayerY(LoopI) = y Then
                        NpcTileIsOpen = False
                        Exit Function
                    End If
                End If
            End If

        Next

    End If

    For LoopI = 1 To MAX_MAP_NPCS

        If MapNpc(mapnum).Npc(LoopI).Num > 0 Then
            If MapNpc(mapnum).Npc(LoopI).x = x Then
                If MapNpc(mapnum).Npc(LoopI).y = y Then
                    NpcTileIsOpen = False
                    Exit Function
                End If
            End If
        End If

    Next

    If Map(mapnum).TileData.Tile(x, y).Type <> TILE_TYPE_WALKABLE Then
        If Map(mapnum).TileData.Tile(x, y).Type <> TILE_TYPE_NPCSPAWN Then
            If Map(mapnum).TileData.Tile(x, y).Type <> TILE_TYPE_ITEM Then
                NpcTileIsOpen = False
            End If
        End If
    End If
End Function

Sub SpawnMapNpcs(ByVal mapnum As Long)
    Dim i As Long

    For i = 1 To MAX_MAP_NPCS
        Call SpawnNpc(i, mapnum)
    Next

End Sub

Sub SpawnAllMapNpcs()
    Dim i As Long

    For i = 1 To MAX_MAPS
        Call SpawnMapNpcs(i)
    Next

End Sub

Function CanNpcMove(ByVal mapnum As Long, ByVal mapNpcNum As Long, ByVal dir As Byte) As Boolean
    Dim i As Long
    Dim n As Long
    Dim x As Long
    Dim y As Long

    ' Check for subscript out of range
    If mapnum <= 0 Or mapnum > MAX_MAPS Or mapNpcNum <= 0 Or mapNpcNum > MAX_MAP_NPCS Or dir < DIR_UP Or dir > DIR_RIGHT Then
        Exit Function
    End If

    x = MapNpc(mapnum).Npc(mapNpcNum).x
    y = MapNpc(mapnum).Npc(mapNpcNum).y
    CanNpcMove = True

    Select Case dir
        Case DIR_UP

            ' Check to make sure not outside of boundries
            If y > 0 Then
                n = Map(mapnum).TileData.Tile(x, y - 1).Type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = mapnum) And (GetPlayerX(i) = MapNpc(mapnum).Npc(mapNpcNum).x) And (GetPlayerY(i) = MapNpc(mapnum).Npc(mapNpcNum).y - 1) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> mapNpcNum) And (MapNpc(mapnum).Npc(i).Num > 0) And (MapNpc(mapnum).Npc(i).x = MapNpc(mapnum).Npc(mapNpcNum).x) And (MapNpc(mapnum).Npc(i).y = MapNpc(mapnum).Npc(mapNpcNum).y - 1) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(Map(mapnum).TileData.Tile(MapNpc(mapnum).Npc(mapNpcNum).x, MapNpc(mapnum).Npc(mapNpcNum).y).DirBlock, DIR_UP + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If

        Case DIR_DOWN

            ' Check to make sure not outside of boundries
            If y < Map(mapnum).MapData.MaxY Then
                n = Map(mapnum).TileData.Tile(x, y + 1).Type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = mapnum) And (GetPlayerX(i) = MapNpc(mapnum).Npc(mapNpcNum).x) And (GetPlayerY(i) = MapNpc(mapnum).Npc(mapNpcNum).y + 1) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> mapNpcNum) And (MapNpc(mapnum).Npc(i).Num > 0) And (MapNpc(mapnum).Npc(i).x = MapNpc(mapnum).Npc(mapNpcNum).x) And (MapNpc(mapnum).Npc(i).y = MapNpc(mapnum).Npc(mapNpcNum).y + 1) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(Map(mapnum).TileData.Tile(MapNpc(mapnum).Npc(mapNpcNum).x, MapNpc(mapnum).Npc(mapNpcNum).y).DirBlock, DIR_DOWN + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If

        Case DIR_LEFT

            ' Check to make sure not outside of boundries
            If x > 0 Then
                n = Map(mapnum).TileData.Tile(x - 1, y).Type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = mapnum) And (GetPlayerX(i) = MapNpc(mapnum).Npc(mapNpcNum).x - 1) And (GetPlayerY(i) = MapNpc(mapnum).Npc(mapNpcNum).y) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> mapNpcNum) And (MapNpc(mapnum).Npc(i).Num > 0) And (MapNpc(mapnum).Npc(i).x = MapNpc(mapnum).Npc(mapNpcNum).x - 1) And (MapNpc(mapnum).Npc(i).y = MapNpc(mapnum).Npc(mapNpcNum).y) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(Map(mapnum).TileData.Tile(MapNpc(mapnum).Npc(mapNpcNum).x, MapNpc(mapnum).Npc(mapNpcNum).y).DirBlock, DIR_LEFT + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If

        Case DIR_RIGHT

            ' Check to make sure not outside of boundries
            If x < Map(mapnum).MapData.MaxX Then
                n = Map(mapnum).TileData.Tile(x + 1, y).Type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = mapnum) And (GetPlayerX(i) = MapNpc(mapnum).Npc(mapNpcNum).x + 1) And (GetPlayerY(i) = MapNpc(mapnum).Npc(mapNpcNum).y) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> mapNpcNum) And (MapNpc(mapnum).Npc(i).Num > 0) And (MapNpc(mapnum).Npc(i).x = MapNpc(mapnum).Npc(mapNpcNum).x + 1) And (MapNpc(mapnum).Npc(i).y = MapNpc(mapnum).Npc(mapNpcNum).y) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(Map(mapnum).TileData.Tile(MapNpc(mapnum).Npc(mapNpcNum).x, MapNpc(mapnum).Npc(mapNpcNum).y).DirBlock, DIR_RIGHT + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If

    End Select

End Function

Sub NpcMove(ByVal mapnum As Long, ByVal mapNpcNum As Long, ByVal dir As Long, ByVal movement As Long)
    Dim packet As String
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If mapnum <= 0 Or mapnum > MAX_MAPS Or mapNpcNum <= 0 Or mapNpcNum > MAX_MAP_NPCS Or dir < DIR_UP Or dir > DIR_RIGHT Or movement < 1 Or movement > 2 Then
        Exit Sub
    End If

    MapNpc(mapnum).Npc(mapNpcNum).dir = dir

    Select Case dir
        Case DIR_UP
            MapNpc(mapnum).Npc(mapNpcNum).y = MapNpc(mapnum).Npc(mapNpcNum).y - 1
            Set Buffer = New clsBuffer
            Buffer.WriteLong SNpcMove
            Buffer.WriteLong mapNpcNum
            Buffer.WriteLong MapNpc(mapnum).Npc(mapNpcNum).x
            Buffer.WriteLong MapNpc(mapnum).Npc(mapNpcNum).y
            Buffer.WriteLong MapNpc(mapnum).Npc(mapNpcNum).dir
            Buffer.WriteLong movement
            SendDataToMap mapnum, Buffer.ToArray()
            Set Buffer = Nothing
        Case DIR_DOWN
            MapNpc(mapnum).Npc(mapNpcNum).y = MapNpc(mapnum).Npc(mapNpcNum).y + 1
            Set Buffer = New clsBuffer
            Buffer.WriteLong SNpcMove
            Buffer.WriteLong mapNpcNum
            Buffer.WriteLong MapNpc(mapnum).Npc(mapNpcNum).x
            Buffer.WriteLong MapNpc(mapnum).Npc(mapNpcNum).y
            Buffer.WriteLong MapNpc(mapnum).Npc(mapNpcNum).dir
            Buffer.WriteLong movement
            SendDataToMap mapnum, Buffer.ToArray()
            Set Buffer = Nothing
        Case DIR_LEFT
            MapNpc(mapnum).Npc(mapNpcNum).x = MapNpc(mapnum).Npc(mapNpcNum).x - 1
            Set Buffer = New clsBuffer
            Buffer.WriteLong SNpcMove
            Buffer.WriteLong mapNpcNum
            Buffer.WriteLong MapNpc(mapnum).Npc(mapNpcNum).x
            Buffer.WriteLong MapNpc(mapnum).Npc(mapNpcNum).y
            Buffer.WriteLong MapNpc(mapnum).Npc(mapNpcNum).dir
            Buffer.WriteLong movement
            SendDataToMap mapnum, Buffer.ToArray()
            Set Buffer = Nothing
        Case DIR_RIGHT
            MapNpc(mapnum).Npc(mapNpcNum).x = MapNpc(mapnum).Npc(mapNpcNum).x + 1
            Set Buffer = New clsBuffer
            Buffer.WriteLong SNpcMove
            Buffer.WriteLong mapNpcNum
            Buffer.WriteLong MapNpc(mapnum).Npc(mapNpcNum).x
            Buffer.WriteLong MapNpc(mapnum).Npc(mapNpcNum).y
            Buffer.WriteLong MapNpc(mapnum).Npc(mapNpcNum).dir
            Buffer.WriteLong movement
            SendDataToMap mapnum, Buffer.ToArray()
            Set Buffer = Nothing
    End Select

End Sub

Sub NpcDir(ByVal mapnum As Long, ByVal mapNpcNum As Long, ByVal dir As Long)
    Dim packet As String
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If mapnum <= 0 Or mapnum > MAX_MAPS Or mapNpcNum <= 0 Or mapNpcNum > MAX_MAP_NPCS Or dir < DIR_UP Or dir > DIR_RIGHT Then
        Exit Sub
    End If

    MapNpc(mapnum).Npc(mapNpcNum).dir = dir
    Set Buffer = New clsBuffer
    Buffer.WriteLong SNpcDir
    Buffer.WriteLong mapNpcNum
    Buffer.WriteLong dir
    SendDataToMap mapnum, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Function GetTotalMapPlayers(ByVal mapnum As Long) As Long
    Dim i As Long
    Dim n As Long
    n = 0

    For i = 1 To Player_HighIndex

        If IsPlaying(i) And GetPlayerMap(i) = mapnum Then
            n = n + 1
        End If

    Next

    GetTotalMapPlayers = n
End Function

Sub ClearTempTiles()
    Dim i As Long

    For i = 1 To MAX_MAPS
        ClearTempTile i
    Next

End Sub

Sub ClearTempTile(ByVal mapnum As Long)
    Dim y As Long
    Dim x As Long
    TempTile(mapnum).DoorTimer = 0
    ReDim TempTile(mapnum).DoorOpen(0 To Map(mapnum).MapData.MaxX, 0 To Map(mapnum).MapData.MaxY)

    For x = 0 To Map(mapnum).MapData.MaxX
        For y = 0 To Map(mapnum).MapData.MaxY
            TempTile(mapnum).DoorOpen(x, y) = NO
        Next
    Next

End Sub

Public Sub CacheResources(ByVal mapnum As Long)
    Dim x As Long, y As Long, Resource_Count As Long
    Resource_Count = 0

    For x = 0 To Map(mapnum).MapData.MaxX
        For y = 0 To Map(mapnum).MapData.MaxY

            If Map(mapnum).TileData.Tile(x, y).Type = TILE_TYPE_RESOURCE Then
                Resource_Count = Resource_Count + 1
                ReDim Preserve ResourceCache(mapnum).ResourceData(0 To Resource_Count)
                ResourceCache(mapnum).ResourceData(Resource_Count).x = x
                ResourceCache(mapnum).ResourceData(Resource_Count).y = y
                ResourceCache(mapnum).ResourceData(Resource_Count).cur_health = Resource(Map(mapnum).TileData.Tile(x, y).Data1).health
            End If

        Next
    Next

    ResourceCache(mapnum).Resource_Count = Resource_Count
End Sub

Sub PlayerSwitchBankSlots(ByVal index As Long, ByVal oldSlot As Long, ByVal newSlot As Long)
Dim OldNum As Long
Dim OldValue As Long
Dim NewNum As Long
Dim NewValue As Long

    If oldSlot = 0 Or newSlot = 0 Then
        Exit Sub
    End If
    
    OldNum = GetPlayerBankItemNum(index, oldSlot)
    OldValue = GetPlayerBankItemValue(index, oldSlot)
    NewNum = GetPlayerBankItemNum(index, newSlot)
    NewValue = GetPlayerBankItemValue(index, newSlot)
    
    SetPlayerBankItemNum index, newSlot, OldNum
    SetPlayerBankItemValue index, newSlot, OldValue
    
    SetPlayerBankItemNum index, oldSlot, NewNum
    SetPlayerBankItemValue index, oldSlot, NewValue
        
    SendBank index
End Sub

Sub PlayerSwitchInvSlots(ByVal index As Long, ByVal oldSlot As Long, ByVal newSlot As Long)
Dim OldNum As Long, OldValue As Long, oldBound As Byte
Dim NewNum As Long, NewValue As Long, newBound As Byte

    If oldSlot = 0 Or newSlot = 0 Then
        Exit Sub
    End If

    OldNum = GetPlayerInvItemNum(index, oldSlot)
    OldValue = GetPlayerInvItemValue(index, oldSlot)
    oldBound = Player(index).Inv(oldSlot).Bound
    NewNum = GetPlayerInvItemNum(index, newSlot)
    NewValue = GetPlayerInvItemValue(index, newSlot)
    newBound = Player(index).Inv(newSlot).Bound
    
    SetPlayerInvItemNum index, newSlot, OldNum
    SetPlayerInvItemValue index, newSlot, OldValue
    Player(index).Inv(newSlot).Bound = oldBound
    
    SetPlayerInvItemNum index, oldSlot, NewNum
    SetPlayerInvItemValue index, oldSlot, NewValue
    Player(index).Inv(oldSlot).Bound = newBound
    
    SendInventory index
End Sub

Sub PlayerSwitchSpellSlots(ByVal index As Long, ByVal oldSlot As Long, ByVal newSlot As Long)
Dim OldNum As Long, NewNum As Long, OldUses As Long, NewUses As Long

    If oldSlot = 0 Or newSlot = 0 Then
        Exit Sub
    End If

    OldNum = Player(index).Spell(oldSlot).Spell
    NewNum = Player(index).Spell(newSlot).Spell
    OldUses = Player(index).Spell(oldSlot).Uses
    NewUses = Player(index).Spell(newSlot).Uses
    
    Player(index).Spell(oldSlot).Spell = NewNum
    Player(index).Spell(oldSlot).Uses = NewUses
    Player(index).Spell(newSlot).Spell = OldNum
    Player(index).Spell(newSlot).Uses = OldUses
    SendPlayerSpells index
End Sub

Sub PlayerUnequipItem(ByVal index As Long, ByVal EqSlot As Long)

    If EqSlot <= 0 Or EqSlot > Equipment.Equipment_Count - 1 Then Exit Sub ' exit out early if error'd
    If FindOpenInvSlot(index, GetPlayerEquipment(index, EqSlot)) > 0 Then
        GiveInvItem index, GetPlayerEquipment(index, EqSlot), 0, , True
        PlayerMsg index, "You unequip " & CheckGrammar(Item(GetPlayerEquipment(index, EqSlot)).Name), Yellow
        ' send the sound
        SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, GetPlayerEquipment(index, EqSlot)
        ' remove equipment
        SetPlayerEquipment index, 0, EqSlot
        SendWornEquipment index
        SendMapEquipment index
        SendStats index
        ' send vitals
        Call SendVital(index, Vitals.HP)
        Call SendVital(index, Vitals.MP)
        ' send vitals to party if in one
        If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
    Else
        PlayerMsg index, "Your inventory is full.", BrightRed
    End If

End Sub

Public Function CheckGrammar(ByVal Word As String, Optional ByVal Caps As Byte = 0) As String
Dim FirstLetter As String * 1
   
    FirstLetter = LCase$(left$(Word, 1))
   
    If FirstLetter = "$" Then
      CheckGrammar = (Mid$(Word, 2, Len(Word) - 1))
      Exit Function
    End If
   
    If FirstLetter Like "*[aeiou]*" Then
        If Caps Then CheckGrammar = "An " & Word Else CheckGrammar = "an " & Word
    Else
        If Caps Then CheckGrammar = "A " & Word Else CheckGrammar = "a " & Word
    End If
End Function

Function isInRange(ByVal Range As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Boolean
Dim nVal As Long
    isInRange = False
    nVal = Sqr((x1 - x2) ^ 2 + (y1 - y2) ^ 2)
    If nVal <= Range Then isInRange = True
End Function

Public Function isDirBlocked(ByRef blockvar As Byte, ByRef dir As Byte) As Boolean
    If Not blockvar And (2 ^ dir) Then
        isDirBlocked = False
    Else
        isDirBlocked = True
    End If
End Function

Public Function RAND(ByVal Low As Long, ByVal High As Long) As Long
    Randomize
    RAND = Int((High - Low + 1) * Rnd) + Low
End Function

' #####################
' ## Party functions ##
' #####################
Public Sub Party_PlayerLeave(ByVal index As Long)
Dim partynum As Long, i As Long

    partynum = TempPlayer(index).inParty
    If partynum > 0 Then
        ' find out how many members we have
        Party_CountMembers partynum
        ' make sure there's more than 2 people
        If Party(partynum).MemberCount > 2 Then
            ' check if leader
            If Party(partynum).Leader = index Then
                ' set next person down as leader
                For i = 1 To MAX_PARTY_MEMBERS
                    If Party(partynum).Member(i) > 0 And Party(partynum).Member(i) <> index Then
                        Party(partynum).Leader = Party(partynum).Member(i)
                        PartyMsg partynum, GetPlayerName(i) & " is now the party leader.", BrightBlue
                        Exit For
                    End If
                Next
                ' leave party
                PartyMsg partynum, GetPlayerName(index) & " has left the party.", BrightRed
                ' remove from array
                For i = 1 To MAX_PARTY_MEMBERS
                    If Party(partynum).Member(i) = index Then
                        Party(partynum).Member(i) = 0
                        Exit For
                    End If
                Next
                ' recount party
                Party_CountMembers partynum
                ' set update to all
                SendPartyUpdate partynum
                ' send clear to player
                SendPartyUpdateTo index
            Else
                ' not the leader, just leave
                PartyMsg partynum, GetPlayerName(index) & " has left the party.", BrightRed
                ' remove from array
                For i = 1 To MAX_PARTY_MEMBERS
                    If Party(partynum).Member(i) = index Then
                        Party(partynum).Member(i) = 0
                        Exit For
                    End If
                Next
                ' recount party
                Party_CountMembers partynum
                ' set update to all
                SendPartyUpdate partynum
                ' send clear to player
                SendPartyUpdateTo index
            End If
        Else
            ' find out how many members we have
            Party_CountMembers partynum
            ' only 2 people, disband
            PartyMsg partynum, "Party disbanded.", BrightRed
            ' clear out everyone's party
            For i = 1 To MAX_PARTY_MEMBERS
                index = Party(partynum).Member(i)
                ' player exist?
                If index > 0 Then
                    ' remove them
                    TempPlayer(index).inParty = 0
                    ' send clear to players
                    SendPartyUpdateTo index
                End If
            Next
            ' clear out the party itself
            ClearParty partynum
        End If
    End If
End Sub

Public Sub Party_Invite(ByVal index As Long, ByVal targetPlayer As Long)
Dim partynum As Long, i As Long

    ' check if the person is a valid target
    If Not IsConnected(targetPlayer) Or Not IsPlaying(targetPlayer) Then Exit Sub
    
    ' make sure they're not busy
    If TempPlayer(targetPlayer).partyInvite > 0 Then
        ' they've already got a request for trade/party
        PlayerMsg index, "This player has an outstanding party invitation already.", BrightRed
        ' exit out early
        Exit Sub
    End If
    ' make syure they're not in a party
    If TempPlayer(targetPlayer).inParty > 0 Then
        ' they're already in a party
        PlayerMsg index, "This player is already in a party.", BrightRed
        'exit out early
        Exit Sub
    End If
    
    ' check if we're in a party
    If TempPlayer(index).inParty > 0 Then
        partynum = TempPlayer(index).inParty
        ' make sure we're the leader
        If Party(partynum).Leader = index Then
            ' got a blank slot?
            For i = 1 To MAX_PARTY_MEMBERS
                If Party(partynum).Member(i) = 0 Then
                    ' send the invitation
                    SendPartyInvite targetPlayer, index
                    ' set the invite target
                    TempPlayer(targetPlayer).partyInvite = index
                    ' let them know
                    PlayerMsg index, "Invitation sent.", Green
                    Exit Sub
                End If
            Next
            ' no room
            PlayerMsg index, "Party is full.", BrightRed
            Exit Sub
        Else
            ' not the leader
            PlayerMsg index, "You are not the party leader.", BrightRed
            Exit Sub
        End If
    Else
        ' not in a party - doesn't matter!
        SendPartyInvite targetPlayer, index
        ' set the invite target
        TempPlayer(targetPlayer).partyInvite = index
        ' let them know
        PlayerMsg index, "Invitation sent.", Green
        Exit Sub
    End If
End Sub

Public Sub Party_InviteAccept(ByVal index As Long, ByVal targetPlayer As Long)
Dim partynum As Long, i As Long, x As Long

    If index = 0 Then Exit Sub
    
    If Not IsConnected(index) Or Not IsPlaying(index) Then
        TempPlayer(targetPlayer).TradeRequest = 0
        TempPlayer(index).TradeRequest = 0
        Exit Sub
    End If
    
    If Not IsConnected(targetPlayer) Or Not IsPlaying(targetPlayer) Then
        TempPlayer(targetPlayer).TradeRequest = 0
        TempPlayer(index).TradeRequest = 0
        Exit Sub
    End If
    
    If TempPlayer(targetPlayer).inParty > 0 Then
        PlayerMsg index, Trim$(GetPlayerName(targetPlayer)) & " is already in a party.", BrightRed
        PlayerMsg targetPlayer, "You're already in a party.", BrightRed
        Exit Sub
    End If

    ' check if already in a party
    If TempPlayer(index).inParty > 0 Then
        ' get the partynumber
        partynum = TempPlayer(index).inParty
        ' got a blank slot?
        For i = 1 To MAX_PARTY_MEMBERS
            If Party(partynum).Member(i) = 0 Then
                'add to the party
                Party(partynum).Member(i) = targetPlayer
                ' recount party
                Party_CountMembers partynum
                ' send everyone's data to everyone
                SendPlayerData_Party partynum
                ' send update to all - including new player
                SendPartyUpdate partynum
                ' Send party vitals to everyone again
                For x = 1 To MAX_PARTY_MEMBERS
                    If Party(partynum).Member(x) > 0 Then
                        SendPartyVitals partynum, Party(partynum).Member(x)
                    End If
                Next
                ' let everyone know they've joined
                PartyMsg partynum, GetPlayerName(targetPlayer) & " has joined the party.", Pink
                ' add them in
                TempPlayer(targetPlayer).inParty = partynum
                Exit Sub
            End If
        Next
        ' no empty slots - let them know
        PlayerMsg index, "Party is full.", BrightRed
        PlayerMsg targetPlayer, "Party is full.", BrightRed
        Exit Sub
    Else
        ' not in a party. Create one with the new person.
        For i = 1 To MAX_PARTYS
            ' find blank party
            If Not Party(i).Leader > 0 Then
                partynum = i
                Exit For
            End If
        Next
        ' create the party
        Party(partynum).MemberCount = 2
        Party(partynum).Leader = index
        Party(partynum).Member(1) = index
        Party(partynum).Member(2) = targetPlayer
        SendPlayerData_Party partynum
        SendPartyUpdate partynum
        SendPartyVitals partynum, index
        SendPartyVitals partynum, targetPlayer
        ' let them know it's created
        PartyMsg partynum, "Party created.", BrightGreen
        PartyMsg partynum, GetPlayerName(index) & " has joined the party.", Pink
        PartyMsg partynum, GetPlayerName(targetPlayer) & " has joined the party.", Pink
        ' clear the invitation
        TempPlayer(targetPlayer).partyInvite = 0
        ' add them to the party
        TempPlayer(index).inParty = partynum
        TempPlayer(targetPlayer).inParty = partynum
        Exit Sub
    End If
End Sub

Public Sub Party_InviteDecline(ByVal index As Long, ByVal targetPlayer As Long)
    If Not IsConnected(index) Or Not IsPlaying(index) Then
        TempPlayer(targetPlayer).TradeRequest = 0
        TempPlayer(index).TradeRequest = 0
        Exit Sub
    End If
    
    If Not IsConnected(targetPlayer) Or Not IsPlaying(targetPlayer) Then
        TempPlayer(targetPlayer).TradeRequest = 0
        TempPlayer(index).TradeRequest = 0
        Exit Sub
    End If
    
    PlayerMsg index, GetPlayerName(targetPlayer) & " has declined to join the party.", BrightRed
    PlayerMsg targetPlayer, "You declined to join the party.", BrightRed
    ' clear the invitation
    TempPlayer(targetPlayer).partyInvite = 0
End Sub

Public Sub Party_CountMembers(ByVal partynum As Long)
Dim i As Long, highIndex As Long, x As Long
    ' find the high index
    For i = MAX_PARTY_MEMBERS To 1 Step -1
        If Party(partynum).Member(i) > 0 Then
            highIndex = i
            Exit For
        End If
    Next
    ' count the members
    For i = 1 To MAX_PARTY_MEMBERS
        ' we've got a blank member
        If Party(partynum).Member(i) = 0 Then
            ' is it lower than the high index?
            If i < highIndex Then
                ' move everyone down a slot
                For x = i To MAX_PARTY_MEMBERS - 1
                    Party(partynum).Member(x) = Party(partynum).Member(x + 1)
                    Party(partynum).Member(x + 1) = 0
                Next
            Else
                ' not lower - highindex is count
                Party(partynum).MemberCount = highIndex
                Exit Sub
            End If
        End If
        ' check if we've reached the max
        If i = MAX_PARTY_MEMBERS Then
            If highIndex = i Then
                Party(partynum).MemberCount = MAX_PARTY_MEMBERS
                Exit Sub
            End If
        End If
    Next
    ' if we're here it means that we need to re-count again
    Party_CountMembers partynum
End Sub

Public Sub Party_ShareExp(ByVal partynum As Long, ByVal exp As Long, ByVal index As Long, Optional ByVal enemyLevel As Long = 0)
Dim expShare As Long, leftOver As Long, i As Long, tmpIndex As Long

    If Party(partynum).MemberCount <= 0 Then Exit Sub

    ' check if it's worth sharing
    If Not exp >= Party(partynum).MemberCount Then
        ' no party - keep exp for self
        GivePlayerEXP index, exp, enemyLevel
        Exit Sub
    End If
    
    ' find out the equal share
    expShare = exp \ Party(partynum).MemberCount
    leftOver = exp Mod Party(partynum).MemberCount
    
    ' loop through and give everyone exp
    For i = 1 To MAX_PARTY_MEMBERS
        tmpIndex = Party(partynum).Member(i)
        ' existing member?Kn
        If tmpIndex > 0 Then
            ' playing?
            If IsConnected(tmpIndex) And IsPlaying(tmpIndex) Then
                ' give them their share
                GivePlayerEXP tmpIndex, expShare, enemyLevel
            End If
        End If
    Next
    
    ' give the remainder to a random member
    tmpIndex = Party(partynum).Member(RAND(1, Party(partynum).MemberCount))
    ' give the exp
    If leftOver > 0 Then GivePlayerEXP tmpIndex, leftOver, enemyLevel
End Sub

Public Sub GivePlayerEXP(ByVal index As Long, ByVal exp As Long, Optional ByVal enemyLevel As Long = 0)
Dim multiplier As Long, partynum As Long, expBonus As Long
    ' no exp
    If exp = 0 Then Exit Sub
    ' rte9
    If index <= 0 Or index > MAX_PLAYERS Then Exit Sub
    ' make sure we're not max level
    If Not GetPlayerLevel(index) >= MAX_LEVELS Then
        ' check for exp deduction
        If enemyLevel > 0 Then
            ' exp deduction
            If enemyLevel <= GetPlayerLevel(index) - 3 Then
                ' 3 levels lower, exit out
                Exit Sub
            ElseIf enemyLevel <= GetPlayerLevel(index) - 2 Then
                ' half exp if enemy is 2 levels lower
                exp = exp / 2
            End If
        End If
        ' check if in party
        partynum = TempPlayer(index).inParty
        If partynum > 0 Then
            If Party(partynum).MemberCount > 1 Then
                multiplier = Party(partynum).MemberCount - 1
                ' multiply the exp
                expBonus = (exp / 100) * (multiplier * 3) ' 3 = 3% per party member
                ' Modify the exp
                exp = exp + expBonus
            End If
        End If
        ' give the exp
        Call SetPlayerExp(index, GetPlayerExp(index) + exp)
        SendEXP index
        SendActionMsg GetPlayerMap(index), "+" & exp & " EXP", White, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
        ' check if we've leveled
        CheckPlayerLevelUp index
    Else
        Call SetPlayerExp(index, 0)
        SendEXP index
    End If
End Sub

Public Sub Unique_Item(ByVal index As Long, ByVal itemNum As Long)
Dim ClassNum As Long, i As Long

    Select Case Item(itemNum).Data1
        Case 1 ' Reset Stats
            ClassNum = GetPlayerClass(index)
            If ClassNum <= 0 Or ClassNum > Max_Classes Then Exit Sub
            ' re-set the actual stats to class defaults
            For i = 1 To Stats.Stat_Count - 1
                SetPlayerStat index, i, Class(ClassNum).Stat(i)
            Next
            ' give player their points back
            SetPlayerPOINTS index, (GetPlayerLevel(index) - 1) * 3
            ' take item
            TakeInvItem index, itemNum, 1
            ' let them know we've done it
            PlayerMsg index, "Your stats have been reset.", BrightGreen
            ' send them their new stats
            SendPlayerData index
        Case Else ' Exit out otherwise
            Exit Sub
    End Select
End Sub

Public Function loginTokenOk(ByVal user As String, ByVal tLoginToken As String) As Boolean
Dim i As Long
    loginTokenOk = False
    For i = 1 To MAX_PLAYERS
        If LoginToken(i).Active Then
            If LoginToken(i).user = user And LoginToken(i).Token = tLoginToken Then
                ' return true
                loginTokenOk = True
                ' clear the token
                LoginToken(i).Active = False
                LoginToken(i).user = vbNullString
                LoginToken(i).Token = vbNullString
                LoginToken(i).TimeCreated = 0
                ' exit out
                Exit Function
            End If
        End If
    Next
End Function

Public Function hasProficiency(ByVal index As Long, ByVal proficiency As Long) As Boolean
    Select Case proficiency
        Case 0 ' None
            hasProficiency = True
            Exit Function
        Case 1 ' Heavy
            If GetPlayerClass(index) = 1 Then
                hasProficiency = True
                Exit Function
            End If
        Case 2 ' Light
            If GetPlayerClass(index) = 2 Or GetPlayerClass(index) = 3 Then
                hasProficiency = True
                Exit Function
            End If
    End Select
    hasProficiency = False
End Function

Function ActiveEventPage(ByVal index As Long, ByVal eventNum As Long) As Long
Dim x As Long, mapnum As Long, process As Boolean
    mapnum = GetPlayerMap(index)
    For x = Map(mapnum).TileData.Events(eventNum).PageCount To 1 Step -1
        ' check if we match
        With Map(mapnum).TileData.Events(eventNum).EventPage(x)
            process = True
            ' player var check
            If .chkPlayerVar Then
                If .PlayerVarNum > 0 Then
                    If Player(index).Variable(.PlayerVarNum) < .PlayerVariable Then
                        process = False
                    End If
                End If
            End If
            ' has item check
            If .chkHasItem Then
                If .HasItemNum > 0 Then
                    If HasItem(index, .HasItemNum) = 0 Then
                        process = False
                    End If
                End If
            End If
            ' this page
            If process = True Then
                ActiveEventPage = x
                Exit Function
            End If
        End With
    Next
End Function
