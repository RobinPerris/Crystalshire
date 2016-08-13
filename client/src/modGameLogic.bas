Attribute VB_Name = "modGameLogic"
Option Explicit

Public Sub GameLoop()
    Dim FrameTime As Long, tick As Long, TickFPS As Long, FPS As Long, i As Long, WalkTimer As Long, x As Long, y As Long
    Dim tmr25 As Long, tmr10000 As Long, tmr100 As Long, mapTimer As Long, chatTmr As Long, targetTmr As Long, fogTmr As Long, barTmr As Long
    Dim barDifference As Long

    ' *** Start GameLoop ***
    Do While InGame
        tick = GetTickCount                            ' Set the inital tick
        ElapsedTime = tick - FrameTime                 ' Set the time difference for time-based movement
        FrameTime = tick                               ' Set the time second loop time to the first.

        ' handle input
        If GetForegroundWindow() = frmMain.hWnd Then
            HandleMouseInput
        End If

        ' * Check surface timers *
        ' Sprites
        If tmr10000 < tick Then
            ' check ping
            Call GetPing
            tmr10000 = tick + 10000
        End If

        If tmr25 < tick Then
            InGame = IsConnected
            Call CheckKeys ' Check to make sure they aren't trying to auto do anything

            If GetForegroundWindow() = frmMain.hWnd Then
                Call CheckInputKeys ' Check which keys were pressed
            End If

            ' check if we need to end the CD icon
            If Count_Spellicon > 0 Then
                For i = 1 To MAX_PLAYER_SPELLS
                    If PlayerSpells(i).Spell > 0 Then
                        If SpellCD(i) > 0 Then
                            If SpellCD(i) + (Spell(PlayerSpells(i).Spell).CDTime * 1000) < tick Then
                                SpellCD(i) = 0
                            End If
                        End If
                    End If
                Next
            End If

            ' check if we need to unlock the player's spell casting restriction
            If SpellBuffer > 0 Then
                If SpellBufferTimer + (Spell(PlayerSpells(SpellBuffer).Spell).CastTime * 1000) < tick Then
                    SpellBuffer = 0
                    SpellBufferTimer = 0
                End If
            End If

            If CanMoveNow Then
                Call CheckMovement ' Check if player is trying to move
                Call CheckAttack   ' Check to see if player is trying to attack
            End If

            For i = 1 To MAX_BYTE
                CheckAnimInstance i
            Next
            
            ' appear tile logic
            AppearTileFadeLogic
            CheckAppearTiles
            
            ' handle events
            If inEvent Then
                If eventNum > 0 Then
                    If eventPageNum > 0 Then
                        If eventCommandNum > 0 Then
                            EventLogic
                        End If
                    End If
                End If
            End If

            tmr25 = tick + 25
        End If

        ' targetting
        If targetTmr < tick Then
            If tabDown Then
                FindNearestTarget
            End If

            targetTmr = tick + 50
        End If
        
        ' chat timer
        If chatTmr < tick Then
            ' scrolling
            If ChatButtonUp Then
                ScrollChatBox 0
            End If

            If ChatButtonDown Then
                ScrollChatBox 1
            End If
            
            ' remove messages
            If chatLastRemove + CHAT_DIFFERENCE_TIMER < GetTickCount Then
                ' remove timed out messages from chat
                For i = Chat_HighIndex To 1 Step -1
                    If Len(Chat(i).text) > 0 Then
                        If Chat(i).visible Then
                            If Chat(i).timer + CHAT_TIMER < tick Then
                                Chat(i).visible = False
                                chatLastRemove = GetTickCount
                                Exit For
                            End If
                        End If
                    End If
                Next
            End If

            chatTmr = tick + 50
        End If

        ' fog scrolling
        If fogTmr < tick Then
            ' move
            fogOffsetX = fogOffsetX - 1
            fogOffsetY = fogOffsetY - 1

            ' reset
            If fogOffsetX < -256 Then fogOffsetX = 0
            If fogOffsetY < -256 Then fogOffsetY = 0
            
            ' reset timer
            fogTmr = tick + 20
        End If

        ' elastic bars
        If barTmr < tick Then
            SetBarWidth BarWidth_GuiHP_Max, BarWidth_GuiHP
            SetBarWidth BarWidth_GuiSP_Max, BarWidth_GuiSP
            SetBarWidth BarWidth_GuiEXP_Max, BarWidth_GuiEXP
            For i = 1 To MAX_MAP_NPCS
                If MapNpc(i).num > 0 Then
                    SetBarWidth BarWidth_NpcHP_Max(i), BarWidth_NpcHP(i)
                End If
            Next

            For i = 1 To MAX_PLAYERS
                If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                    SetBarWidth BarWidth_PlayerHP_Max(i), BarWidth_PlayerHP(i)
                End If
            Next

            ' reset timer
            barTmr = tick + 10
        End If

        ' Animations!
        If mapTimer < tick Then

            ' animate waterfalls
            Select Case waterfallFrame

                Case 0
                    waterfallFrame = 1

                Case 1
                    waterfallFrame = 2

                Case 2
                    waterfallFrame = 0
            End Select

            ' animate autotiles
            Select Case autoTileFrame

                Case 0
                    autoTileFrame = 1

                Case 1
                    autoTileFrame = 2

                Case 2
                    autoTileFrame = 0
            End Select

            ' animate textbox
            If chatShowLine = "|" Then
                chatShowLine = vbNullString
            Else
                chatShowLine = "|"
            End If
            
            ' re-set timer
            mapTimer = tick + 500
        End If

        ' Process input before rendering, otherwise input will be behind by 1 frame
        If WalkTimer < tick Then

            For i = 1 To Player_HighIndex

                If IsPlaying(i) Then
                    Call ProcessMovement(i)
                End If

            Next i

            ' Process npc movements (actually move them)
            For i = 1 To Npc_HighIndex

                If map.MapData.Npc(i) > 0 Then
                    Call ProcessNpcMovement(i)
                End If

            Next i

            WalkTimer = tick + 30 ' edit this value to change WalkTimer
        End If

        ' *********************
        ' ** Render Graphics **
        ' *********************
        Call Render_Graphics

        DoEvents

        ' Lock fps
        If Not FPS_Lock Then

            Do While GetTickCount < tick + 20
                DoEvents
                Sleep 1
            Loop

        End If

        ' Calculate fps
        If TickFPS < tick Then
            GameFPS = FPS
            TickFPS = tick + 1000
            FPS = 0
        Else
            FPS = FPS + 1
        End If

    Loop

    frmMain.visible = False

    If isLogging Then
        isLogging = False
        MenuLoop
        GettingMap = True
        Stop_Music
        Play_Music MenuMusic
    Else
        ' Shutdown the game
        Call SetStatus("Destroying game data.")
        Call DestroyGame
    End If

End Sub

Public Sub MenuLoop()
    Dim FrameTime As Long, tick As Long, TickFPS As Long, FPS As Long, tmr500 As Long, fadeTmr As Long

    ' *** Start GameLoop ***
    Do While inMenu
        tick = GetTickCount                            ' Set the inital tick
        ElapsedTime = tick - FrameTime                 ' Set the time difference for time-based movement
        FrameTime = tick                               ' Set the time second loop time to the first.

        ' handle input
        If GetForegroundWindow() = frmMain.hWnd Then
            HandleMouseInput
        End If
        
        ' Animations!
        If tmr500 < tick Then
            ' animate textbox
            If chatShowLine = "|" Then
                chatShowLine = vbNullString
            Else
                chatShowLine = "|"
            End If

            ' re-set timer
            tmr500 = tick + 500
        End If
        
        ' trailer
        If videoPlaying Then VideoLoop
        
        ' fading
        If fadeTmr < tick Then
            If Not videoPlaying Then
                If fadeAlpha > 5 Then
                    ' lower fade
                    fadeAlpha = fadeAlpha - 5
                Else
                    fadeAlpha = 0
                End If
            End If
            fadeTmr = tick + 1
        End If

        ' *********************
        ' ** Render Graphics **
        ' *********************
        Call Render_Menu
        
        ' do events
        DoEvents

        ' Lock fps
        If Not FPS_Lock Then

            Do While GetTickCount < tick + 20
                DoEvents
                Sleep 1
            Loop

        End If

        ' Calculate fps
        If TickFPS < tick Then
            GameFPS = FPS
            TickFPS = tick + 1000
            FPS = 0
        Else
            FPS = FPS + 1
        End If

    Loop

End Sub

Sub ProcessMovement(ByVal index As Long)
    Dim MovementSpeed As Long

    ' Check if player is walking, and if so process moving them over
    Select Case Player(index).Moving

        Case MOVING_WALKING: MovementSpeed = RUN_SPEED

        Case MOVING_RUNNING: MovementSpeed = WALK_SPEED

        Case Else: Exit Sub
    End Select

    Select Case GetPlayerDir(index)

        Case DIR_UP
            Player(index).yOffset = Player(index).yOffset - MovementSpeed

            If Player(index).yOffset < 0 Then Player(index).yOffset = 0

        Case DIR_DOWN
            Player(index).yOffset = Player(index).yOffset + MovementSpeed

            If Player(index).yOffset > 0 Then Player(index).yOffset = 0

        Case DIR_LEFT
            Player(index).xOffset = Player(index).xOffset - MovementSpeed

            If Player(index).xOffset < 0 Then Player(index).xOffset = 0

        Case DIR_RIGHT
            Player(index).xOffset = Player(index).xOffset + MovementSpeed

            If Player(index).xOffset > 0 Then Player(index).xOffset = 0
    End Select

    ' Check if completed walking over to the next tile
    If Player(index).Moving > 0 Then
        If GetPlayerDir(index) = DIR_RIGHT Or GetPlayerDir(index) = DIR_DOWN Then
            If (Player(index).xOffset >= 0) And (Player(index).yOffset >= 0) Then
                Player(index).Moving = 0

                If Player(index).Step = 0 Then
                    Player(index).Step = 2
                Else
                    Player(index).Step = 0
                End If
            End If

        Else

            If (Player(index).xOffset <= 0) And (Player(index).yOffset <= 0) Then
                Player(index).Moving = 0

                If Player(index).Step = 0 Then
                    Player(index).Step = 2
                Else
                    Player(index).Step = 0
                End If
            End If
        End If
    End If

End Sub

Sub ProcessNpcMovement(ByVal MapNpcNum As Long)
    Dim MovementSpeed As Long

    ' Check if NPC is walking, and if so process moving them over
    If MapNpc(MapNpcNum).Moving = MOVING_WALKING Then
        MovementSpeed = RUN_SPEED
    Else
        Exit Sub
    End If

    Select Case MapNpc(MapNpcNum).dir

        Case DIR_UP
            MapNpc(MapNpcNum).yOffset = MapNpc(MapNpcNum).yOffset - MovementSpeed

            If MapNpc(MapNpcNum).yOffset < 0 Then MapNpc(MapNpcNum).yOffset = 0

        Case DIR_DOWN
            MapNpc(MapNpcNum).yOffset = MapNpc(MapNpcNum).yOffset + MovementSpeed

            If MapNpc(MapNpcNum).yOffset > 0 Then MapNpc(MapNpcNum).yOffset = 0

        Case DIR_LEFT
            MapNpc(MapNpcNum).xOffset = MapNpc(MapNpcNum).xOffset - MovementSpeed

            If MapNpc(MapNpcNum).xOffset < 0 Then MapNpc(MapNpcNum).xOffset = 0

        Case DIR_RIGHT
            MapNpc(MapNpcNum).xOffset = MapNpc(MapNpcNum).xOffset + MovementSpeed

            If MapNpc(MapNpcNum).xOffset > 0 Then MapNpc(MapNpcNum).xOffset = 0
    End Select

    ' Check if completed walking over to the next tile
    If MapNpc(MapNpcNum).Moving > 0 Then
        If MapNpc(MapNpcNum).dir = DIR_RIGHT Or MapNpc(MapNpcNum).dir = DIR_DOWN Then
            If (MapNpc(MapNpcNum).xOffset >= 0) And (MapNpc(MapNpcNum).yOffset >= 0) Then
                MapNpc(MapNpcNum).Moving = 0

                If MapNpc(MapNpcNum).Step = 0 Then
                    MapNpc(MapNpcNum).Step = 2
                Else
                    MapNpc(MapNpcNum).Step = 0
                End If
            End If

        Else

            If (MapNpc(MapNpcNum).xOffset <= 0) And (MapNpc(MapNpcNum).yOffset <= 0) Then
                MapNpc(MapNpcNum).Moving = 0

                If MapNpc(MapNpcNum).Step = 0 Then
                    MapNpc(MapNpcNum).Step = 2
                Else
                    MapNpc(MapNpcNum).Step = 0
                End If
            End If
        End If
    End If

End Sub

Sub CheckMapGetItem()
    Dim Buffer As New clsBuffer, tmpIndex As Long, i As Long, x As Long
    Set Buffer = New clsBuffer

    If GetTickCount > Player(MyIndex).MapGetTimer + 250 Then

        ' find out if we want to pick it up
        For i = 1 To MAX_MAP_ITEMS

            If MapItem(i).x = Player(MyIndex).x And MapItem(i).y = Player(MyIndex).y Then
                If MapItem(i).num > 0 Then
                    If Item(MapItem(i).num).BindType = 1 Then

                        ' make sure it's not a party drop
                        If Party.Leader > 0 Then

                            For x = 1 To MAX_PARTY_MEMBERS
                                tmpIndex = Party.Member(x)

                                If tmpIndex > 0 Then
                                    If Trim$(GetPlayerName(tmpIndex)) = Trim$(MapItem(i).playerName) Then
                                        If Item(MapItem(i).num).ClassReq > 0 Then
                                            If Item(MapItem(i).num).ClassReq <> Player(MyIndex).Class Then
                                                Dialogue "Loot Check", "This item is BoP and is not for your class.", "Are you sure you want to pick it up?", TypeLOOTITEM, StyleYESNO
                                                Exit Sub
                                            End If
                                        End If
                                    End If
                                End If

                            Next

                        End If

                    Else
                        'not bound
                        Exit For
                    End If
                End If
            End If

        Next

        ' nevermind, pick it up
        Player(MyIndex).MapGetTimer = GetTickCount
        Buffer.WriteLong CMapGetItem
        SendData Buffer.ToArray()
    End If

    Set Buffer = Nothing
End Sub

Public Sub CheckAttack()
    Dim Buffer As clsBuffer
    Dim attackspeed As Long

    If ControlDown Then
        If SpellBuffer > 0 Then Exit Sub ' currently casting a spell, can't attack
        If StunDuration > 0 Then Exit Sub ' stunned, can't attack

        ' speed from weapon
        If GetPlayerEquipment(MyIndex, Weapon) > 0 Then
            attackspeed = Item(GetPlayerEquipment(MyIndex, Weapon)).speed
        Else
            attackspeed = 1000
        End If

        If Player(MyIndex).AttackTimer + attackspeed < GetTickCount Then
            If Player(MyIndex).Attacking = 0 Then

                With Player(MyIndex)
                    .Attacking = 1
                    .AttackTimer = GetTickCount
                End With

                Set Buffer = New clsBuffer
                Buffer.WriteLong CAttack
                SendData Buffer.ToArray()
                Set Buffer = Nothing
            End If
        End If
    End If

End Sub

Function IsTryingToMove() As Boolean

    'If DirUp Or DirDown Or DirLeft Or DirRight Then
    If wDown Or sDown Or aDown Or dDown Or upDown Or leftDown Or downDown Or rightDown Then
        IsTryingToMove = True
    End If

End Function

Function CanMove() As Boolean
    Dim d As Long
    CanMove = True

    ' Make sure they aren't trying to move when they are already moving
    If Player(MyIndex).Moving <> 0 Then
        CanMove = False
        Exit Function
    End If

    ' Make sure they haven't just casted a spell
    'If SpellBuffer > 0 Then
    '    CanMove = False
    '    Exit Function
    'End If

    ' make sure they're not stunned
    If StunDuration > 0 Then
        CanMove = False
        Exit Function
    End If

    ' make sure they're not in a shop
    If InShop > 0 Then
        CanMove = False
        Exit Function
    End If

    ' not in bank
    If InBank Then
        CanMove = False
        Exit Function
    End If

    If inTutorial Then
        CanMove = False
        Exit Function
    End If

    d = GetPlayerDir(MyIndex)

    If wDown Or upDown Then
        Call SetPlayerDir(MyIndex, DIR_UP)

        ' Check to see if they are trying to go out of bounds
        If GetPlayerY(MyIndex) > 0 Then
            If CheckDirection(DIR_UP) Then
                CanMove = False

                ' Set the new direction if they weren't facing that direction
                If d <> DIR_UP Then
                    Call SendPlayerDir
                End If

                Exit Function
            End If

        Else

            ' Check if they can warp to a new map
            If map.MapData.Up > 0 Then
                Call MapEditorLeaveMap
                Call SendPlayerRequestNewMap
                GettingMap = True
                CanMoveNow = False
            End If

            CanMove = False
            Exit Function
        End If
    End If

    If downDown Or sDown Then
        Call SetPlayerDir(MyIndex, DIR_DOWN)

        ' Check to see if they are trying to go out of bounds
        If GetPlayerY(MyIndex) < map.MapData.MaxY Then
            If CheckDirection(DIR_DOWN) Then
                CanMove = False

                ' Set the new direction if they weren't facing that direction
                If d <> DIR_DOWN Then
                    Call SendPlayerDir
                End If

                Exit Function
            End If

        Else

            ' Check if they can warp to a new map
            If map.MapData.Down > 0 Then
                Call MapEditorLeaveMap
                Call SendPlayerRequestNewMap
                GettingMap = True
                CanMoveNow = False
            End If

            CanMove = False
            Exit Function
        End If
    End If

    If aDown Or leftDown Then
        Call SetPlayerDir(MyIndex, DIR_LEFT)

        ' Check to see if they are trying to go out of bounds
        If GetPlayerX(MyIndex) > 0 Then
            If CheckDirection(DIR_LEFT) Then
                CanMove = False

                ' Set the new direction if they weren't facing that direction
                If d <> DIR_LEFT Then
                    Call SendPlayerDir
                End If

                Exit Function
            End If

        Else

            ' Check if they can warp to a new map
            If map.MapData.left > 0 Then
                Call MapEditorLeaveMap
                Call SendPlayerRequestNewMap
                GettingMap = True
                CanMoveNow = False
            End If

            CanMove = False
            Exit Function
        End If
    End If

    If dDown Or rightDown Then
        Call SetPlayerDir(MyIndex, DIR_RIGHT)

        ' Check to see if they are trying to go out of bounds
        If GetPlayerX(MyIndex) < map.MapData.MaxX Then
            If CheckDirection(DIR_RIGHT) Then
                CanMove = False

                ' Set the new direction if they weren't facing that direction
                If d <> DIR_RIGHT Then
                    Call SendPlayerDir
                End If

                Exit Function
            End If

        Else

            ' Check if they can warp to a new map
            If map.MapData.Right > 0 Then
                Call MapEditorLeaveMap
                Call SendPlayerRequestNewMap
                GettingMap = True
                CanMoveNow = False
            End If

            CanMove = False
            Exit Function
        End If
    End If

End Function

Function CheckDirection(ByVal direction As Byte) As Boolean
    Dim x As Long, y As Long, i As Long, EventCount As Long, page As Long
    
    CheckDirection = False

    If GettingMap Then Exit Function

    ' check directional blocking
    If isDirBlocked(map.TileData.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex)).DirBlock, direction + 1) Then
        CheckDirection = True
        Exit Function
    End If

    Select Case direction

        Case DIR_UP
            x = GetPlayerX(MyIndex)
            y = GetPlayerY(MyIndex) - 1

        Case DIR_DOWN
            x = GetPlayerX(MyIndex)
            y = GetPlayerY(MyIndex) + 1

        Case DIR_LEFT
            x = GetPlayerX(MyIndex) - 1
            y = GetPlayerY(MyIndex)

        Case DIR_RIGHT
            x = GetPlayerX(MyIndex) + 1
            y = GetPlayerY(MyIndex)
    End Select

    ' Check to see if the map tile is blocked or not
    If map.TileData.Tile(x, y).Type = TILE_TYPE_BLOCKED Then
        CheckDirection = True
        Exit Function
    End If

    ' Check to see if the map tile is tree or not
    If map.TileData.Tile(x, y).Type = TILE_TYPE_RESOURCE Then
        CheckDirection = True
        Exit Function
    End If
    
    ' Check to make sure that any events on that space aren't blocked
    EventCount = map.TileData.EventCount
    For i = 1 To EventCount
        With map.TileData.Events(i)
            If .x = x And .y = y Then
                ' Get the active event page
                page = ActiveEventPage(i)
                If page > 0 Then
                    If map.TileData.Events(i).EventPage(page).WalkThrough = 0 Then
                        CheckDirection = True
                        Exit Function
                    End If
                End If
            End If
        End With
    Next

    ' Check to see if the key door is open or not
    If map.TileData.Tile(x, y).Type = TILE_TYPE_KEY Then
        ' This actually checks if its open or not
        If TempTile(x, y).DoorOpen = 0 Then
            CheckDirection = True
            Exit Function
        End If
    End If

    ' Check to see if a player is already on that tile
    If map.MapData.Moral = 0 Then
        For i = 1 To Player_HighIndex
            If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
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
    For i = 1 To Npc_HighIndex
        If MapNpc(i).num > 0 Then
            If MapNpc(i).x = x Then
                If MapNpc(i).y = y Then
                    CheckDirection = True
                    Exit Function
                End If
            End If
        End If
    Next
    
    ' check if it's a drop warp - avoid if walking
    If ShiftDown Then
        If map.TileData.Tile(x, y).Type = TILE_TYPE_WARP Then
            If map.TileData.Tile(x, y).Data4 Then
                CheckDirection = True
            End If
        End If
    End If

End Function

Sub CheckMovement()

    If Not GettingMap Then
    If IsTryingToMove Then
        If CanMove Then

            ' Check if player has the shift key down for running
            If ShiftDown Then
                Player(MyIndex).Moving = MOVING_RUNNING
            Else
                Player(MyIndex).Moving = MOVING_WALKING
            End If

            Select Case GetPlayerDir(MyIndex)

                Case DIR_UP
                    Call SendPlayerMove
                    Player(MyIndex).yOffset = PIC_Y
                    Call SetPlayerY(MyIndex, GetPlayerY(MyIndex) - 1)

                Case DIR_DOWN
                    Call SendPlayerMove
                    Player(MyIndex).yOffset = PIC_Y * -1
                    Call SetPlayerY(MyIndex, GetPlayerY(MyIndex) + 1)

                Case DIR_LEFT
                    Call SendPlayerMove
                    Player(MyIndex).xOffset = PIC_X
                    Call SetPlayerX(MyIndex, GetPlayerX(MyIndex) - 1)

                Case DIR_RIGHT
                    Call SendPlayerMove
                    Player(MyIndex).xOffset = PIC_X * -1
                    Call SetPlayerX(MyIndex, GetPlayerX(MyIndex) + 1)
            End Select

            If map.TileData.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex)).Type = TILE_TYPE_WARP Then
                GettingMap = True
            End If
        End If
    End If
    End If

End Sub

Public Function isInBounds()

    If (CurX >= 0) Then
        If (CurX <= map.MapData.MaxX) Then
            If (CurY >= 0) Then
                If (CurY <= map.MapData.MaxY) Then
                    isInBounds = True
                End If
            End If
        End If
    End If

End Function

Public Function IsValidMapPoint(ByVal x As Long, ByVal y As Long) As Boolean
    IsValidMapPoint = False

    If x < 0 Then Exit Function
    If y < 0 Then Exit Function
    If x > map.MapData.MaxX Then Exit Function
    If y > map.MapData.MaxY Then Exit Function
    IsValidMapPoint = True
End Function

Public Function IsItem(startX As Long, startY As Long) As Long
Dim tempRec As RECT
Dim i As Long
    For i = 1 To MAX_INV
        If GetPlayerInvItemNum(MyIndex, i) Then
            With tempRec
                .top = startY + InvTop + ((InvOffsetY + 32) * ((i - 1) \ InvColumns))
                .bottom = .top + PIC_Y
                .left = startX + InvLeft + ((InvOffsetX + 32) * (((i - 1) Mod InvColumns)))
                .Right = .left + PIC_X
            End With

            If currMouseX >= tempRec.left And currMouseX <= tempRec.Right Then
                If currMouseY >= tempRec.top And currMouseY <= tempRec.bottom Then
                    IsItem = i
                    Exit Function
                End If
            End If
        End If
    Next
End Function

Public Function IsTrade(startX As Long, startY As Long) As Long
Dim tempRec As RECT
Dim i As Long

    For i = 1 To MAX_INV
        With tempRec
            .top = startY + TradeTop + ((TradeOffsetY + 32) * ((i - 1) \ TradeColumns))
            .bottom = .top + PIC_Y
            .left = startX + TradeLeft + ((TradeOffsetX + 32) * (((i - 1) Mod TradeColumns)))
            .Right = .left + PIC_X
        End With

        If currMouseX >= tempRec.left And currMouseX <= tempRec.Right Then
            If currMouseY >= tempRec.top And currMouseY <= tempRec.bottom Then
                IsTrade = i
                Exit Function
            End If
        End If
    Next
End Function

Public Function IsEqItem(startX As Long, startY As Long) As Long
Dim tempRec As RECT
Dim i As Long
    For i = 1 To Equipment.Equipment_Count - 1
        If GetPlayerEquipment(MyIndex, i) Then
            With tempRec
                .top = startY + EqTop + (32 * ((i - 1) \ EqColumns))
                .bottom = .top + PIC_Y
                .left = startX + EqLeft + ((EqOffsetX + 32) * (((i - 1) Mod EqColumns)))
                .Right = .left + PIC_X
            End With

            If currMouseX >= tempRec.left And currMouseX <= tempRec.Right Then
                If currMouseY >= tempRec.top And currMouseY <= tempRec.bottom Then
                    IsEqItem = i
                    Exit Function
                End If
            End If
        End If
    Next
End Function

Public Function IsSkill(startX As Long, startY As Long) As Long
Dim tempRec As RECT
Dim i As Long

    For i = 1 To MAX_PLAYER_SPELLS
        If PlayerSpells(i).Spell Then
            With tempRec
                .top = startY + SkillTop + ((SkillOffsetY + 32) * ((i - 1) \ SkillColumns))
                .bottom = .top + PIC_Y
                .left = startX + SkillLeft + ((SkillOffsetX + 32) * (((i - 1) Mod SkillColumns)))
                .Right = .left + PIC_X
            End With

            If currMouseX >= tempRec.left And currMouseX <= tempRec.Right Then
                If currMouseY >= tempRec.top And currMouseY <= tempRec.bottom Then
                    IsSkill = i
                    Exit Function
                End If
            End If
        End If
    Next
End Function

Public Function IsHotbar(startX As Long, startY As Long) As Long
Dim tempRec As RECT
Dim i As Long

    For i = 1 To MAX_HOTBAR
        If Hotbar(i).Slot Then
            With tempRec
                .top = startY + HotbarTop
                .bottom = .top + PIC_Y
                .left = startX + HotbarLeft + ((i - 1) * HotbarOffsetX)
                .Right = .left + PIC_X
            End With

            If currMouseX >= tempRec.left And currMouseX <= tempRec.Right Then
                If currMouseY >= tempRec.top And currMouseY <= tempRec.bottom Then
                    IsHotbar = i
                    Exit Function
                End If
            End If
        End If
    Next
End Function

Public Sub UseItem()

    ' Check for subscript out of range
    If InventoryItemSelected < 1 Or InventoryItemSelected > MAX_INV Then
        Exit Sub
    End If

    Call SendUseItem(InventoryItemSelected)
End Sub

Public Sub ForgetSpell(ByVal spellSlot As Long)
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If spellSlot < 1 Or spellSlot > MAX_PLAYER_SPELLS Then
        Exit Sub
    End If

    ' dont let them forget a spell which is in CD
    If SpellCD(spellSlot) > 0 Then
        AddText "Cannot forget a spell which is cooling down!", BrightRed
        Exit Sub
    End If

    ' dont let them forget a spell which is buffered
    If SpellBuffer = spellSlot Then
        AddText "Cannot forget a spell which you are casting!", BrightRed
        Exit Sub
    End If

    If PlayerSpells(spellSlot).Spell > 0 Then
        Set Buffer = New clsBuffer
        Buffer.WriteLong CForgetSpell
        Buffer.WriteLong spellSlot
        SendData Buffer.ToArray()
        Set Buffer = Nothing
    Else
        AddText "No spell here.", BrightRed
    End If

End Sub

Public Sub CastSpell(ByVal spellSlot As Long)
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If spellSlot < 1 Or spellSlot > MAX_PLAYER_SPELLS Then
        Exit Sub
    End If

    If SpellCD(spellSlot) > 0 Then
        AddText "Spell has not cooled down yet!", BrightRed
        Exit Sub
    End If
    
    ' make sure we're not casting same spell
    If SpellBuffer > 0 Then
        If SpellBuffer = spellSlot Then
            ' stop them
            Exit Sub
        End If
    End If

    If PlayerSpells(spellSlot).Spell = 0 Then Exit Sub

    ' Check if player has enough MP
    If GetPlayerVital(MyIndex, Vitals.MP) < Spell(PlayerSpells(spellSlot).Spell).MPCost Then
        Call AddText("Not enough MP to cast " & Trim$(Spell(PlayerSpells(spellSlot).Spell).name) & ".", BrightRed)
        Exit Sub
    End If

    If PlayerSpells(spellSlot).Spell > 0 Then
        If GetTickCount > Player(MyIndex).AttackTimer + 1000 Then
            If Player(MyIndex).Moving = 0 Then
                Set Buffer = New clsBuffer
                Buffer.WriteLong CCast
                Buffer.WriteLong spellSlot
                SendData Buffer.ToArray()
                Set Buffer = Nothing
                SpellBuffer = spellSlot
                SpellBufferTimer = GetTickCount
            Else
                Call AddText("Cannot cast while walking!", BrightRed)
            End If
        End If

    Else
        Call AddText("No spell here.", BrightRed)
    End If

End Sub

Sub ClearTempTile()
    Dim x As Long
    Dim y As Long
    ReDim TempTile(0 To map.MapData.MaxX, 0 To map.MapData.MaxY)

    For x = 0 To map.MapData.MaxX
        For y = 0 To map.MapData.MaxY
            TempTile(x, y).DoorOpen = 0

            If Not GettingMap Then cacheRenderState x, y, MapLayer.Mask
        Next
    Next

End Sub

Public Sub DevMsg(ByVal text As String, ByVal Color As Byte)

    If InGame Then
        If GetPlayerAccess(MyIndex) > ADMIN_DEVELOPER Then
            Call AddText(text, Color)
        End If
    End If

    Debug.Print text
End Sub

Public Function TwipsToPixels(ByVal twip_val As Long, ByVal XorY As Byte) As Long

    If XorY = 0 Then
        TwipsToPixels = twip_val / Screen.TwipsPerPixelX
    ElseIf XorY = 1 Then
        TwipsToPixels = twip_val / Screen.TwipsPerPixelY
    End If

End Function

Public Function PixelsToTwips(ByVal pixel_val As Long, ByVal XorY As Byte) As Long

    If XorY = 0 Then
        PixelsToTwips = pixel_val * Screen.TwipsPerPixelX
    ElseIf XorY = 1 Then
        PixelsToTwips = pixel_val * Screen.TwipsPerPixelY
    End If

End Function

Public Function ConvertCurrency(ByVal amount As Long) As String

    If Int(amount) < 10000 Then
        ConvertCurrency = amount
    ElseIf Int(amount) < 999999 Then
        ConvertCurrency = Int(amount / 1000) & "k"
    ElseIf Int(amount) < 999999999 Then
        ConvertCurrency = Int(amount / 1000000) & "m"
    Else
        ConvertCurrency = Int(amount / 1000000000) & "b"
    End If

End Function

Public Sub CacheResources()
    Dim x As Long, y As Long, Resource_Count As Long
    Resource_Count = 0

    For x = 0 To map.MapData.MaxX
        For y = 0 To map.MapData.MaxY

            If map.TileData.Tile(x, y).Type = TILE_TYPE_RESOURCE Then
                Resource_Count = Resource_Count + 1
                ReDim Preserve MapResource(0 To Resource_Count)
                MapResource(Resource_Count).x = x
                MapResource(Resource_Count).y = y
            End If

        Next
    Next

    Resource_Index = Resource_Count
End Sub

Public Sub CreateActionMsg(ByVal message As String, ByVal Color As Integer, ByVal MsgType As Byte, ByVal x As Long, ByVal y As Long)
    Dim i As Long
    ActionMsgIndex = ActionMsgIndex + 1

    If ActionMsgIndex >= MAX_BYTE Then ActionMsgIndex = 1

    With ActionMsg(ActionMsgIndex)
        .message = message
        .Color = Color
        .Type = MsgType
        .Created = GetTickCount
        .Scroll = 1
        .x = x
        .y = y
        .alpha = 255
    End With

    If ActionMsg(ActionMsgIndex).Type = ACTIONMsgSCROLL Then
        ActionMsg(ActionMsgIndex).y = ActionMsg(ActionMsgIndex).y + Rand(-2, 6)
        ActionMsg(ActionMsgIndex).x = ActionMsg(ActionMsgIndex).x + Rand(-8, 8)
    End If

    ' find the new high index
    For i = MAX_BYTE To 1 Step -1

        If ActionMsg(i).Created > 0 Then
            Action_HighIndex = i + 1
            Exit For
        End If

    Next

    ' make sure we don't overflow
    If Action_HighIndex > MAX_BYTE Then Action_HighIndex = MAX_BYTE
End Sub

Public Sub ClearActionMsg(ByVal index As Byte)
    Dim i As Long
    ActionMsg(index).message = vbNullString
    ActionMsg(index).Created = 0
    ActionMsg(index).Type = 0
    ActionMsg(index).Color = 0
    ActionMsg(index).Scroll = 0
    ActionMsg(index).x = 0
    ActionMsg(index).y = 0

    ' find the new high index
    For i = MAX_BYTE To 1 Step -1

        If ActionMsg(i).Created > 0 Then
            Action_HighIndex = i + 1
            Exit For
        End If

    Next

    ' make sure we don't overflow
    If Action_HighIndex > MAX_BYTE Then Action_HighIndex = MAX_BYTE
End Sub

Public Sub CheckAnimInstance(ByVal index As Long)
    Dim looptime As Long
    Dim Layer As Long
    Dim FrameCount As Long

    ' if doesn't exist then exit sub
    If AnimInstance(index).Animation <= 0 Then Exit Sub
    If AnimInstance(index).Animation >= MAX_ANIMATIONS Then Exit Sub

    For Layer = 0 To 1

        If AnimInstance(index).Used(Layer) Then
            looptime = Animation(AnimInstance(index).Animation).looptime(Layer)

            FrameCount = Animation(AnimInstance(index).Animation).Frames(Layer)

            ' if zero'd then set so we don't have extra loop and/or frame
            If AnimInstance(index).FrameIndex(Layer) = 0 Then AnimInstance(index).FrameIndex(Layer) = 1
            If AnimInstance(index).LoopIndex(Layer) = 0 Then AnimInstance(index).LoopIndex(Layer) = 1

            ' check if frame timer is set, and needs to have a frame change
            If AnimInstance(index).timer(Layer) + looptime <= GetTickCount Then

                ' check if out of range
                If AnimInstance(index).FrameIndex(Layer) >= FrameCount Then
                    AnimInstance(index).LoopIndex(Layer) = AnimInstance(index).LoopIndex(Layer) + 1

                    If AnimInstance(index).LoopIndex(Layer) > Animation(AnimInstance(index).Animation).LoopCount(Layer) Then
                        AnimInstance(index).Used(Layer) = False
                    Else
                        AnimInstance(index).FrameIndex(Layer) = 1
                    End If

                Else
                    AnimInstance(index).FrameIndex(Layer) = AnimInstance(index).FrameIndex(Layer) + 1
                End If

                AnimInstance(index).timer(Layer) = GetTickCount
            End If
        End If

    Next

    ' if neither layer is used, clear
    If AnimInstance(index).Used(0) = False And AnimInstance(index).Used(1) = False Then ClearAnimInstance (index)
End Sub

Public Function GetBankItemNum(ByVal bankslot As Long) As Long

    If bankslot = 0 Then
        GetBankItemNum = 0
        Exit Function
    End If

    If bankslot > MAX_BANK Then
        GetBankItemNum = 0
        Exit Function
    End If

    GetBankItemNum = Bank.Item(bankslot).num
End Function

Public Sub SetBankItemNum(ByVal bankslot As Long, ByVal itemNum As Long)
    Bank.Item(bankslot).num = itemNum
End Sub

Public Function GetBankItemValue(ByVal bankslot As Long) As Long
    GetBankItemValue = Bank.Item(bankslot).value
End Function

Public Sub SetBankItemValue(ByVal bankslot As Long, ByVal ItemValue As Long)
    Bank.Item(bankslot).value = ItemValue
End Sub

' BitWise Operators for directional blocking
Public Sub setDirBlock(ByRef blockvar As Byte, ByRef dir As Byte, ByVal block As Boolean)

    If block Then
        blockvar = blockvar Or (2 ^ dir)
    Else
        blockvar = blockvar And Not (2 ^ dir)
    End If

End Sub

Public Function isDirBlocked(ByRef blockvar As Byte, ByRef dir As Byte) As Boolean

    If Not blockvar And (2 ^ dir) Then
        isDirBlocked = False
    Else
        isDirBlocked = True
    End If

End Function

Public Sub PlayMapSound(ByVal x As Long, ByVal y As Long, ByVal entityType As Long, ByVal entityNum As Long)
    Dim soundName As String

    If entityNum <= 0 Then Exit Sub

    ' find the sound
    Select Case entityType

            ' animations
        Case SoundEntity.seAnimation

            If entityNum > MAX_ANIMATIONS Then Exit Sub
            soundName = Trim$(Animation(entityNum).sound)

            ' items
        Case SoundEntity.seItem

            If entityNum > MAX_ITEMS Then Exit Sub
            soundName = Trim$(Item(entityNum).sound)

            ' npcs
        Case SoundEntity.seNpc

            If entityNum > MAX_NPCS Then Exit Sub
            soundName = Trim$(Npc(entityNum).sound)

            ' resources
        Case SoundEntity.seResource

            If entityNum > MAX_RESOURCES Then Exit Sub
            soundName = Trim$(Resource(entityNum).sound)

            ' spells
        Case SoundEntity.seSpell

            If entityNum > MAX_SPELLS Then Exit Sub
            soundName = Trim$(Spell(entityNum).sound)

            ' other
        Case Else
            Exit Sub
    End Select

    ' exit out if it's not set
    If Trim$(soundName) = "None." Then Exit Sub

    ' play the sound
    If x > 0 And y > 0 Then Play_Sound soundName, x, y
End Sub

Public Sub CloseDialogue()
    diaIndex = 0
    HideWindow GetWindowIndex("winBlank")
    HideWindow GetWindowIndex("winDialogue")
End Sub

Public Sub Dialogue(ByVal header As String, ByVal body As String, ByVal body2 As String, ByVal index As Long, Optional ByVal style As Byte = 1, Optional ByVal Data1 As Long = 0)

    ' exit out if we've already got a dialogue open
    If diaIndex > 0 Then Exit Sub
    
    ' set buttons
    With Windows(GetWindowIndex("winDialogue"))
        If style = StyleYESNO Then
           .Controls(GetControlIndex("winDialogue", "btnYes")).visible = True
            .Controls(GetControlIndex("winDialogue", "btnNo")).visible = True
            .Controls(GetControlIndex("winDialogue", "btnOkay")).visible = False
            .Controls(GetControlIndex("winDialogue", "txtInput")).visible = False
            .Controls(GetControlIndex("winDialogue", "lblBody_2")).visible = True
        ElseIf style = StyleOKAY Then
            .Controls(GetControlIndex("winDialogue", "btnYes")).visible = False
            .Controls(GetControlIndex("winDialogue", "btnNo")).visible = False
            .Controls(GetControlIndex("winDialogue", "btnOkay")).visible = True
            .Controls(GetControlIndex("winDialogue", "txtInput")).visible = False
            .Controls(GetControlIndex("winDialogue", "lblBody_2")).visible = True
        ElseIf style = StyleINPUT Then
            .Controls(GetControlIndex("winDialogue", "btnYes")).visible = False
            .Controls(GetControlIndex("winDialogue", "btnNo")).visible = False
            .Controls(GetControlIndex("winDialogue", "btnOkay")).visible = True
            .Controls(GetControlIndex("winDialogue", "txtInput")).visible = True
            .Controls(GetControlIndex("winDialogue", "lblBody_2")).visible = False
        End If
        
        ' set labels
        .Controls(GetControlIndex("winDialogue", "lblHeader")).text = header
        .Controls(GetControlIndex("winDialogue", "lblBody_1")).text = body
        .Controls(GetControlIndex("winDialogue", "lblBody_2")).text = body2
        .Controls(GetControlIndex("winDialogue", "txtInput")).text = vbNullString
    End With
    
    ' set it all up
    diaIndex = index
    diaData1 = Data1
    diaStyle = style
    
    ' make the windows visible
    ShowWindow GetWindowIndex("winBlank"), True
    ShowWindow GetWindowIndex("winDialogue"), True
End Sub

Public Sub dialogueHandler(ByVal index As Long)
Dim value As Long, diaInput As String

    Dim Buffer As New clsBuffer
    Set Buffer = New clsBuffer
    
    diaInput = Trim$(Windows(GetWindowIndex("winDialogue")).Controls(GetControlIndex("winDialogue", "txtInput")).text)

    ' find out which button
    If index = 1 Then ' okay button

        ' dialogue index
        Select Case diaIndex
                Case TypeTRADEAMOUNT
                    value = Val(diaInput)
                    TradeItem diaData1, value
                Case TypeDROPITEM
                    value = Val(diaInput)
                    SendDropItem diaData1, value
        End Select

    ElseIf index = 2 Then ' yes button

        ' dialogue index
        Select Case diaIndex

            Case TypeTRADE
                SendAcceptTradeRequest

            Case TypeFORGET

                ForgetSpell diaData1

            Case TypePARTY
                SendAcceptParty

            Case TypeLOOTITEM
                ' send the packet
                Player(MyIndex).MapGetTimer = GetTickCount
                Buffer.WriteLong CMapGetItem
                SendData Buffer.ToArray()

            Case TypeDELCHAR
                ' send the deletion
                SendDelChar diaData1
        End Select

    ElseIf index = 3 Then ' no button

        ' dialogue index
        Select Case diaIndex

            Case TypeTRADE
                SendDeclineTradeRequest

            Case TypePARTY
                SendDeclineParty
        End Select
    End If

    CloseDialogue
    diaIndex = 0
    diaInput = vbNullString
End Sub

Public Function ConvertMapX(ByVal x As Long) As Long
    ConvertMapX = x - (TileView.left * PIC_X) - Camera.left
End Function

Public Function ConvertMapY(ByVal y As Long) As Long
    ConvertMapY = y - (TileView.top * PIC_Y) - Camera.top
End Function

Public Sub UpdateCamera()
    Dim offsetX As Long, offsetY As Long, startX As Long, startY As Long, EndX As Long, EndY As Long
    
    offsetX = Player(MyIndex).xOffset + PIC_X
    offsetY = Player(MyIndex).yOffset + PIC_Y
    startX = GetPlayerX(MyIndex) - ((TileWidth + 1) \ 2) - 1
    startY = GetPlayerY(MyIndex) - ((TileHeight + 1) \ 2) - 1

    If TileWidth + 1 <= map.MapData.MaxX Then
        If startX < 0 Then
            offsetX = 0
    
            If startX = -1 Then
                If Player(MyIndex).xOffset > 0 Then
                    offsetX = Player(MyIndex).xOffset
                End If
            End If
    
            startX = 0
        End If
        
        EndX = startX + (TileWidth + 1) + 1
        
        If EndX > map.MapData.MaxX Then
            offsetX = 32
    
            If EndX = map.MapData.MaxX + 1 Then
                If Player(MyIndex).xOffset < 0 Then
                    offsetX = Player(MyIndex).xOffset + PIC_X
                End If
            End If
    
            EndX = map.MapData.MaxX
            startX = EndX - TileWidth - 1
        End If
    Else
        EndX = startX + (TileWidth + 1) + 1
    End If
    
    If TileHeight + 1 <= map.MapData.MaxY Then
        If startY < 0 Then
            offsetY = 0
    
            If startY = -1 Then
                If Player(MyIndex).yOffset > 0 Then
                    offsetY = Player(MyIndex).yOffset
                End If
            End If
    
            startY = 0
        End If
        
        EndY = startY + (TileHeight + 1) + 1
        
        If EndY > map.MapData.MaxY Then
            offsetY = 32
    
            If EndY = map.MapData.MaxY + 1 Then
                If Player(MyIndex).yOffset < 0 Then
                    offsetY = Player(MyIndex).yOffset + PIC_Y
                End If
            End If
    
            EndY = map.MapData.MaxY
            startY = EndY - TileHeight - 1
        End If
    Else
        EndY = startY + (TileHeight + 1) + 1
    End If
    
    If TileWidth + 1 = map.MapData.MaxX Then
        offsetX = 0
    End If
    
    If TileHeight + 1 = map.MapData.MaxY Then
        offsetY = 0
    End If

    With TileView
        .top = startY
        .bottom = EndY
        .left = startX
        .Right = EndX
    End With

    With Camera
        .top = offsetY
        .bottom = .top + ScreenY
        .left = offsetX
        .Right = .left + ScreenX
    End With

    CurX = TileView.left + ((GlobalX + Camera.left) \ PIC_X)
    CurY = TileView.top + ((GlobalY + Camera.top) \ PIC_Y)
    GlobalX_Map = GlobalX + (TileView.left * PIC_X) + Camera.left
    GlobalY_Map = GlobalY + (TileView.top * PIC_Y) + Camera.top
End Sub

Public Function CensorWord(ByVal sString As String) As String
    CensorWord = String$(Len(sString), "*")
End Function

Public Sub placeAutotile(ByVal layernum As Long, ByVal x As Long, ByVal y As Long, ByVal tileQuarter As Byte, ByVal autoTileLetter As String)

    With Autotile(x, y).Layer(layernum).QuarterTile(tileQuarter)

        Select Case autoTileLetter

            Case "a"
                .x = autoInner(1).x
                .y = autoInner(1).y

            Case "b"
                .x = autoInner(2).x
                .y = autoInner(2).y

            Case "c"
                .x = autoInner(3).x
                .y = autoInner(3).y

            Case "d"
                .x = autoInner(4).x
                .y = autoInner(4).y

            Case "e"
                .x = autoNW(1).x
                .y = autoNW(1).y

            Case "f"
                .x = autoNW(2).x
                .y = autoNW(2).y

            Case "g"
                .x = autoNW(3).x
                .y = autoNW(3).y

            Case "h"
                .x = autoNW(4).x
                .y = autoNW(4).y

            Case "i"
                .x = autoNE(1).x
                .y = autoNE(1).y

            Case "j"
                .x = autoNE(2).x
                .y = autoNE(2).y

            Case "k"
                .x = autoNE(3).x
                .y = autoNE(3).y

            Case "l"
                .x = autoNE(4).x
                .y = autoNE(4).y

            Case "m"
                .x = autoSW(1).x
                .y = autoSW(1).y

            Case "n"
                .x = autoSW(2).x
                .y = autoSW(2).y

            Case "o"
                .x = autoSW(3).x
                .y = autoSW(3).y

            Case "p"
                .x = autoSW(4).x
                .y = autoSW(4).y

            Case "q"
                .x = autoSE(1).x
                .y = autoSE(1).y

            Case "r"
                .x = autoSE(2).x
                .y = autoSE(2).y

            Case "s"
                .x = autoSE(3).x
                .y = autoSE(3).y

            Case "t"
                .x = autoSE(4).x
                .y = autoSE(4).y
        End Select

    End With

End Sub

Public Sub initAutotiles()
    Dim x As Long, y As Long, layernum As Long
    ' Procedure used to cache autotile positions. All positioning is
    ' independant from the tileset. Calculations are convoluted and annoying.
    ' Maths is not my strong point. Luckily we're caching them so it's a one-off
    ' thing when the map is originally loaded. As such optimisation isn't an issue.
    ' For simplicity's sake we cache all subtile SOURCE positions in to an array.
    ' We also give letters to each subtile for easy rendering tweaks. ;]
    ' First, we need to re-size the array
    ReDim Autotile(0 To map.MapData.MaxX, 0 To map.MapData.MaxY)
    ' Inner tiles (Top right subtile region)
    ' NW - a
    autoInner(1).x = 32
    autoInner(1).y = 0
    ' NE - b
    autoInner(2).x = 48
    autoInner(2).y = 0
    ' SW - c
    autoInner(3).x = 32
    autoInner(3).y = 16
    ' SE - d
    autoInner(4).x = 48
    autoInner(4).y = 16
    ' Outer Tiles - NW (bottom subtile region)
    ' NW - e
    autoNW(1).x = 0
    autoNW(1).y = 32
    ' NE - f
    autoNW(2).x = 16
    autoNW(2).y = 32
    ' SW - g
    autoNW(3).x = 0
    autoNW(3).y = 48
    ' SE - h
    autoNW(4).x = 16
    autoNW(4).y = 48
    ' Outer Tiles - NE (bottom subtile region)
    ' NW - i
    autoNE(1).x = 32
    autoNE(1).y = 32
    ' NE - g
    autoNE(2).x = 48
    autoNE(2).y = 32
    ' SW - k
    autoNE(3).x = 32
    autoNE(3).y = 48
    ' SE - l
    autoNE(4).x = 48
    autoNE(4).y = 48
    ' Outer Tiles - SW (bottom subtile region)
    ' NW - m
    autoSW(1).x = 0
    autoSW(1).y = 64
    ' NE - n
    autoSW(2).x = 16
    autoSW(2).y = 64
    ' SW - o
    autoSW(3).x = 0
    autoSW(3).y = 80
    ' SE - p
    autoSW(4).x = 16
    autoSW(4).y = 80
    ' Outer Tiles - SE (bottom subtile region)
    ' NW - q
    autoSE(1).x = 32
    autoSE(1).y = 64
    ' NE - r
    autoSE(2).x = 48
    autoSE(2).y = 64
    ' SW - s
    autoSE(3).x = 32
    autoSE(3).y = 80
    ' SE - t
    autoSE(4).x = 48
    autoSE(4).y = 80

    For x = 0 To map.MapData.MaxX
        For y = 0 To map.MapData.MaxY
            For layernum = 1 To MapLayer.Layer_Count - 1
                ' calculate the subtile positions and place them
                calculateAutotile x, y, layernum
                ' cache the rendering state of the tiles and set them
                cacheRenderState x, y, layernum
            Next
        Next
    Next

End Sub

Public Sub cacheRenderState(ByVal x As Long, ByVal y As Long, ByVal layernum As Long)
    Dim quarterNum As Long

    ' exit out early
    If x < 0 Or x > map.MapData.MaxX Or y < 0 Or y > map.MapData.MaxY Then Exit Sub

    With map.TileData.Tile(x, y)

        ' check if the tile can be rendered
        If .Layer(layernum).tileSet <= 0 Or .Layer(layernum).tileSet > Count_Tileset Then
            Autotile(x, y).Layer(layernum).renderState = RENDER_STATE_NONE
            Exit Sub
        End If
        
        ' check if we're a bottom
        If layernum = MapLayer.Ground Then
            ' check if bottom
            If y > 0 Then
                If map.TileData.Tile(x, y - 1).Type = TILE_TYPE_APPEAR Then
                    If map.TileData.Tile(x, y - 1).Data2 Then
                        Autotile(x, y).Layer(layernum).renderState = RENDER_STATE_APPEAR
                        Exit Sub
                    End If
                End If
            End If
        End If

        ' check if it's a key - hide mask if key is closed
        If layernum = MapLayer.Mask Then
            If .Type = TILE_TYPE_KEY Then
                If TempTile(x, y).DoorOpen = 0 Then
                    Autotile(x, y).Layer(layernum).renderState = RENDER_STATE_NONE
                    Exit Sub
                End If
            End If
            If .Type = TILE_TYPE_APPEAR Then
                Autotile(x, y).Layer(layernum).renderState = RENDER_STATE_APPEAR
                Exit Sub
            End If
        End If

        ' check if it needs to be rendered as an autotile
        If .Autotile(layernum) = AUTOTILE_NONE Or .Autotile(layernum) = AUTOTILE_FAKE Or Options.NoAuto = 1 Then
            ' default to... default
            Autotile(x, y).Layer(layernum).renderState = RENDER_STATE_NORMAL
        Else
            Autotile(x, y).Layer(layernum).renderState = RENDER_STATE_AUTOTILE

            ' cache tileset positioning
            For quarterNum = 1 To 4
                Autotile(x, y).Layer(layernum).srcX(quarterNum) = (map.TileData.Tile(x, y).Layer(layernum).x * 32) + Autotile(x, y).Layer(layernum).QuarterTile(quarterNum).x
                Autotile(x, y).Layer(layernum).srcY(quarterNum) = (map.TileData.Tile(x, y).Layer(layernum).y * 32) + Autotile(x, y).Layer(layernum).QuarterTile(quarterNum).y
            Next

        End If

    End With

End Sub

Public Sub calculateAutotile(ByVal x As Long, ByVal y As Long, ByVal layernum As Long)

    ' Right, so we've split the tile block in to an easy to remember
    ' collection of letters. We now need to do the calculations to find
    ' out which little lettered block needs to be rendered. We do this
    ' by reading the surrounding tiles to check for matches.
    ' First we check to make sure an autotile situation is actually there.
    ' Then we calculate exactly which situation has arisen.
    ' The situations are "inner", "outer", "horizontal", "vertical" and "fill".
    ' Exit out if we don't have an auatotile
    If map.TileData.Tile(x, y).Autotile(layernum) = 0 Then Exit Sub

    ' Okay, we have autotiling but which one?
    Select Case map.TileData.Tile(x, y).Autotile(layernum)

            ' Normal or animated - same difference
        Case AUTOTILE_NORMAL, AUTOTILE_ANIM
            ' North West Quarter
            CalculateNW_Normal layernum, x, y
            ' North East Quarter
            CalculateNE_Normal layernum, x, y
            ' South West Quarter
            CalculateSW_Normal layernum, x, y
            ' South East Quarter
            CalculateSE_Normal layernum, x, y

            ' Cliff
        Case AUTOTILE_CLIFF
            ' North West Quarter
            CalculateNW_Cliff layernum, x, y
            ' North East Quarter
            CalculateNE_Cliff layernum, x, y
            ' South West Quarter
            CalculateSW_Cliff layernum, x, y
            ' South East Quarter
            CalculateSE_Cliff layernum, x, y

            ' Waterfalls
        Case AUTOTILE_WATERFALL
            ' North West Quarter
            CalculateNW_Waterfall layernum, x, y
            ' North East Quarter
            CalculateNE_Waterfall layernum, x, y
            ' South West Quarter
            CalculateSW_Waterfall layernum, x, y
            ' South East Quarter
            CalculateSE_Waterfall layernum, x, y

            ' Anything else
        Case Else
            ' Don't need to render anything... it's fake or not an autotile
    End Select

End Sub

' Normal autotiling
Public Sub CalculateNW_Normal(ByVal layernum As Long, ByVal x As Long, ByVal y As Long)
    Dim tmpTile(1 To 3) As Boolean
    Dim situation As Byte

    ' North West
    If checkTileMatch(layernum, x, y, x - 1, y - 1) Then tmpTile(1) = True

    ' North
    If checkTileMatch(layernum, x, y, x, y - 1) Then tmpTile(2) = True

    ' West
    If checkTileMatch(layernum, x, y, x - 1, y) Then tmpTile(3) = True

    ' Calculate Situation - Inner
    If Not tmpTile(2) And Not tmpTile(3) Then situation = AUTO_INNER

    ' Horizontal
    If Not tmpTile(2) And tmpTile(3) Then situation = AUTO_HORIZONTAL

    ' Vertical
    If tmpTile(2) And Not tmpTile(3) Then situation = AUTO_VERTICAL

    ' Outer
    If Not tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_OUTER

    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL

    ' Actually place the subtile
    Select Case situation

        Case AUTO_INNER
            placeAutotile layernum, x, y, 1, "e"

        Case AUTO_OUTER
            placeAutotile layernum, x, y, 1, "a"

        Case AUTO_HORIZONTAL
            placeAutotile layernum, x, y, 1, "i"

        Case AUTO_VERTICAL
            placeAutotile layernum, x, y, 1, "m"

        Case AUTO_FILL
            placeAutotile layernum, x, y, 1, "q"
    End Select

End Sub

Public Sub CalculateNE_Normal(ByVal layernum As Long, ByVal x As Long, ByVal y As Long)
    Dim tmpTile(1 To 3) As Boolean
    Dim situation As Byte

    ' North
    If checkTileMatch(layernum, x, y, x, y - 1) Then tmpTile(1) = True

    ' North East
    If checkTileMatch(layernum, x, y, x + 1, y - 1) Then tmpTile(2) = True

    ' East
    If checkTileMatch(layernum, x, y, x + 1, y) Then tmpTile(3) = True

    ' Calculate Situation - Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = AUTO_INNER

    ' Horizontal
    If Not tmpTile(1) And tmpTile(3) Then situation = AUTO_HORIZONTAL

    ' Vertical
    If tmpTile(1) And Not tmpTile(3) Then situation = AUTO_VERTICAL

    ' Outer
    If tmpTile(1) And Not tmpTile(2) And tmpTile(3) Then situation = AUTO_OUTER

    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL

    ' Actually place the subtile
    Select Case situation

        Case AUTO_INNER
            placeAutotile layernum, x, y, 2, "j"

        Case AUTO_OUTER
            placeAutotile layernum, x, y, 2, "b"

        Case AUTO_HORIZONTAL
            placeAutotile layernum, x, y, 2, "f"

        Case AUTO_VERTICAL
            placeAutotile layernum, x, y, 2, "r"

        Case AUTO_FILL
            placeAutotile layernum, x, y, 2, "n"
    End Select

End Sub

Public Sub CalculateSW_Normal(ByVal layernum As Long, ByVal x As Long, ByVal y As Long)
    Dim tmpTile(1 To 3) As Boolean
    Dim situation As Byte

    ' West
    If checkTileMatch(layernum, x, y, x - 1, y) Then tmpTile(1) = True

    ' South West
    If checkTileMatch(layernum, x, y, x - 1, y + 1) Then tmpTile(2) = True

    ' South
    If checkTileMatch(layernum, x, y, x, y + 1) Then tmpTile(3) = True

    ' Calculate Situation - Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = AUTO_INNER

    ' Horizontal
    If tmpTile(1) And Not tmpTile(3) Then situation = AUTO_HORIZONTAL

    ' Vertical
    If Not tmpTile(1) And tmpTile(3) Then situation = AUTO_VERTICAL

    ' Outer
    If tmpTile(1) And Not tmpTile(2) And tmpTile(3) Then situation = AUTO_OUTER

    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL

    ' Actually place the subtile
    Select Case situation

        Case AUTO_INNER
            placeAutotile layernum, x, y, 3, "o"

        Case AUTO_OUTER
            placeAutotile layernum, x, y, 3, "c"

        Case AUTO_HORIZONTAL
            placeAutotile layernum, x, y, 3, "s"

        Case AUTO_VERTICAL
            placeAutotile layernum, x, y, 3, "g"

        Case AUTO_FILL
            placeAutotile layernum, x, y, 3, "k"
    End Select

End Sub

Public Sub CalculateSE_Normal(ByVal layernum As Long, ByVal x As Long, ByVal y As Long)
    Dim tmpTile(1 To 3) As Boolean
    Dim situation As Byte

    ' South
    If checkTileMatch(layernum, x, y, x, y + 1) Then tmpTile(1) = True

    ' South East
    If checkTileMatch(layernum, x, y, x + 1, y + 1) Then tmpTile(2) = True

    ' East
    If checkTileMatch(layernum, x, y, x + 1, y) Then tmpTile(3) = True

    ' Calculate Situation - Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = AUTO_INNER

    ' Horizontal
    If Not tmpTile(1) And tmpTile(3) Then situation = AUTO_HORIZONTAL

    ' Vertical
    If tmpTile(1) And Not tmpTile(3) Then situation = AUTO_VERTICAL

    ' Outer
    If tmpTile(1) And Not tmpTile(2) And tmpTile(3) Then situation = AUTO_OUTER

    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL

    ' Actually place the subtile
    Select Case situation

        Case AUTO_INNER
            placeAutotile layernum, x, y, 4, "t"

        Case AUTO_OUTER
            placeAutotile layernum, x, y, 4, "d"

        Case AUTO_HORIZONTAL
            placeAutotile layernum, x, y, 4, "p"

        Case AUTO_VERTICAL
            placeAutotile layernum, x, y, 4, "l"

        Case AUTO_FILL
            placeAutotile layernum, x, y, 4, "h"
    End Select

End Sub

' Waterfall autotiling
Public Sub CalculateNW_Waterfall(ByVal layernum As Long, ByVal x As Long, ByVal y As Long)
    Dim tmpTile As Boolean

    ' West
    If checkTileMatch(layernum, x, y, x - 1, y) Then tmpTile = True

    ' Actually place the subtile
    If tmpTile Then
        ' Extended
        placeAutotile layernum, x, y, 1, "i"
    Else
        ' Edge
        placeAutotile layernum, x, y, 1, "e"
    End If

End Sub

Public Sub CalculateNE_Waterfall(ByVal layernum As Long, ByVal x As Long, ByVal y As Long)
    Dim tmpTile As Boolean

    ' East
    If checkTileMatch(layernum, x, y, x + 1, y) Then tmpTile = True

    ' Actually place the subtile
    If tmpTile Then
        ' Extended
        placeAutotile layernum, x, y, 2, "f"
    Else
        ' Edge
        placeAutotile layernum, x, y, 2, "j"
    End If

End Sub

Public Sub CalculateSW_Waterfall(ByVal layernum As Long, ByVal x As Long, ByVal y As Long)
    Dim tmpTile As Boolean

    ' West
    If checkTileMatch(layernum, x, y, x - 1, y) Then tmpTile = True

    ' Actually place the subtile
    If tmpTile Then
        ' Extended
        placeAutotile layernum, x, y, 3, "k"
    Else
        ' Edge
        placeAutotile layernum, x, y, 3, "g"
    End If

End Sub

Public Sub CalculateSE_Waterfall(ByVal layernum As Long, ByVal x As Long, ByVal y As Long)
    Dim tmpTile As Boolean

    ' East
    If checkTileMatch(layernum, x, y, x + 1, y) Then tmpTile = True

    ' Actually place the subtile
    If tmpTile Then
        ' Extended
        placeAutotile layernum, x, y, 4, "h"
    Else
        ' Edge
        placeAutotile layernum, x, y, 4, "l"
    End If

End Sub

' Cliff autotiling
Public Sub CalculateNW_Cliff(ByVal layernum As Long, ByVal x As Long, ByVal y As Long)
    Dim tmpTile(1 To 3) As Boolean
    Dim situation As Byte

    ' North West
    If checkTileMatch(layernum, x, y, x - 1, y - 1) Then tmpTile(1) = True

    ' North
    If checkTileMatch(layernum, x, y, x, y - 1) Then tmpTile(2) = True

    ' West
    If checkTileMatch(layernum, x, y, x - 1, y) Then tmpTile(3) = True

    ' Calculate Situation - Horizontal
    If Not tmpTile(2) And tmpTile(3) Then situation = AUTO_HORIZONTAL

    ' Vertical
    If tmpTile(2) And Not tmpTile(3) Then situation = AUTO_VERTICAL

    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL

    ' Inner
    If Not tmpTile(2) And Not tmpTile(3) Then situation = AUTO_INNER

    ' Actually place the subtile
    Select Case situation

        Case AUTO_INNER
            placeAutotile layernum, x, y, 1, "e"

        Case AUTO_HORIZONTAL
            placeAutotile layernum, x, y, 1, "i"

        Case AUTO_VERTICAL
            placeAutotile layernum, x, y, 1, "m"

        Case AUTO_FILL
            placeAutotile layernum, x, y, 1, "q"
    End Select

End Sub

Public Sub CalculateNE_Cliff(ByVal layernum As Long, ByVal x As Long, ByVal y As Long)
    Dim tmpTile(1 To 3) As Boolean
    Dim situation As Byte

    ' North
    If checkTileMatch(layernum, x, y, x, y - 1) Then tmpTile(1) = True

    ' North East
    If checkTileMatch(layernum, x, y, x + 1, y - 1) Then tmpTile(2) = True

    ' East
    If checkTileMatch(layernum, x, y, x + 1, y) Then tmpTile(3) = True

    ' Calculate Situation - Horizontal
    If Not tmpTile(1) And tmpTile(3) Then situation = AUTO_HORIZONTAL

    ' Vertical
    If tmpTile(1) And Not tmpTile(3) Then situation = AUTO_VERTICAL

    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL

    ' Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = AUTO_INNER

    ' Actually place the subtile
    Select Case situation

        Case AUTO_INNER
            placeAutotile layernum, x, y, 2, "j"

        Case AUTO_HORIZONTAL
            placeAutotile layernum, x, y, 2, "f"

        Case AUTO_VERTICAL
            placeAutotile layernum, x, y, 2, "r"

        Case AUTO_FILL
            placeAutotile layernum, x, y, 2, "n"
    End Select

End Sub

Public Sub CalculateSW_Cliff(ByVal layernum As Long, ByVal x As Long, ByVal y As Long)
    Dim tmpTile(1 To 3) As Boolean
    Dim situation As Byte

    ' West
    If checkTileMatch(layernum, x, y, x - 1, y) Then tmpTile(1) = True

    ' South West
    If checkTileMatch(layernum, x, y, x - 1, y + 1) Then tmpTile(2) = True

    ' South
    If checkTileMatch(layernum, x, y, x, y + 1) Then tmpTile(3) = True

    ' Calculate Situation - Horizontal
    If tmpTile(1) And Not tmpTile(3) Then situation = AUTO_HORIZONTAL

    ' Vertical
    If Not tmpTile(1) And tmpTile(3) Then situation = AUTO_VERTICAL

    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL

    ' Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = AUTO_INNER

    ' Actually place the subtile
    Select Case situation

        Case AUTO_INNER
            placeAutotile layernum, x, y, 3, "o"

        Case AUTO_HORIZONTAL
            placeAutotile layernum, x, y, 3, "s"

        Case AUTO_VERTICAL
            placeAutotile layernum, x, y, 3, "g"

        Case AUTO_FILL
            placeAutotile layernum, x, y, 3, "k"
    End Select

End Sub

Public Sub CalculateSE_Cliff(ByVal layernum As Long, ByVal x As Long, ByVal y As Long)
    Dim tmpTile(1 To 3) As Boolean
    Dim situation As Byte

    ' South
    If checkTileMatch(layernum, x, y, x, y + 1) Then tmpTile(1) = True

    ' South East
    If checkTileMatch(layernum, x, y, x + 1, y + 1) Then tmpTile(2) = True

    ' East
    If checkTileMatch(layernum, x, y, x + 1, y) Then tmpTile(3) = True

    ' Calculate Situation -  Horizontal
    If Not tmpTile(1) And tmpTile(3) Then situation = AUTO_HORIZONTAL

    ' Vertical
    If tmpTile(1) And Not tmpTile(3) Then situation = AUTO_VERTICAL

    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL

    ' Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = AUTO_INNER

    ' Actually place the subtile
    Select Case situation

        Case AUTO_INNER
            placeAutotile layernum, x, y, 4, "t"

        Case AUTO_HORIZONTAL
            placeAutotile layernum, x, y, 4, "p"

        Case AUTO_VERTICAL
            placeAutotile layernum, x, y, 4, "l"

        Case AUTO_FILL
            placeAutotile layernum, x, y, 4, "h"
    End Select

End Sub

Public Function checkTileMatch(ByVal layernum As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Boolean
    ' we'll exit out early if true
    checkTileMatch = True

    ' if it's off the map then set it as autotile and exit out early
    If x2 < 0 Or x2 > map.MapData.MaxX Or y2 < 0 Or y2 > map.MapData.MaxY Then
        checkTileMatch = True
        Exit Function
    End If

    ' fakes ALWAYS return true
    If map.TileData.Tile(x2, y2).Autotile(layernum) = AUTOTILE_FAKE Then
        checkTileMatch = True
        Exit Function
    End If

    ' check neighbour is an autotile
    If map.TileData.Tile(x2, y2).Autotile(layernum) = 0 Then
        checkTileMatch = False
        Exit Function
    End If

    ' check we're a matching
    If map.TileData.Tile(x1, y1).Layer(layernum).tileSet <> map.TileData.Tile(x2, y2).Layer(layernum).tileSet Then
        checkTileMatch = False
        Exit Function
    End If

    ' check tiles match
    If map.TileData.Tile(x1, y1).Layer(layernum).x <> map.TileData.Tile(x2, y2).Layer(layernum).x Then
        checkTileMatch = False
        Exit Function
    End If

    If map.TileData.Tile(x1, y1).Layer(layernum).y <> map.TileData.Tile(x2, y2).Layer(layernum).y Then
        checkTileMatch = False
        Exit Function
    End If

End Function

Public Sub OpenNpcChat(ByVal npcNum As Long, ByVal mT As String, ByRef o() As String)
Dim i As Long, x As Long

    ' find out how many options we have
    convOptions = 0
    For i = 1 To 4
        If Len(o(i)) > 0 Then convOptions = convOptions + 1
    Next
    
    ' gui stuff
    With Windows(GetWindowIndex("winNpcChat"))
        ' set main text
        .Window.text = "Conversation with " & Trim$(Npc(npcNum).name)
        .Controls(GetControlIndex("winNpcChat", "lblChat")).text = mT
        ' make everything visible
        For i = 1 To 4
            .Controls(GetControlIndex("winNpcChat", "btnOpt" & i)).top = optPos(i)
            .Controls(GetControlIndex("winNpcChat", "btnOpt" & i)).visible = True
        Next
        ' set sizes
        .Window.height = optHeight
        .Controls(GetControlIndex("winNpcChat", "picParchment")).height = .Window.height - 30
        ' move options depending on count
        If convOptions < 4 Then
            For i = convOptions + 1 To 4
                .Controls(GetControlIndex("winNpcChat", "btnOpt" & i)).top = optPos(i)
                .Controls(GetControlIndex("winNpcChat", "btnOpt" & i)).visible = False
            Next
            For i = 1 To convOptions
                .Controls(GetControlIndex("winNpcChat", "btnOpt" & i)).top = optPos(i + (4 - convOptions))
                .Controls(GetControlIndex("winNpcChat", "btnOpt" & i)).visible = True
            Next
            .Window.height = optHeight - ((4 - convOptions) * 18)
            .Controls(GetControlIndex("winNpcChat", "picParchment")).height = .Window.height - 32
        End If
        ' set labels
        x = convOptions
        For i = 1 To 4
            .Controls(GetControlIndex("winNpcChat", "btnOpt" & i)).text = x & ". " & o(i)
            x = x - 1
        Next
        For i = 0 To 5
            .Controls(GetControlIndex("winNpcChat", "picFace")).image(i) = Tex_Face(Npc(npcNum).sprite)
        Next
    End With
    
    ' we're in chat now boy
    inChat = True
    
    ' show the window
    ShowWindow GetWindowIndex("winNpcChat")
End Sub

Public Sub SetTutorialState(ByVal stateNum As Byte)
    Dim i As Long

    Select Case stateNum

        Case 1 ' introduction
            chatText = "Ah, so you have appeared at last my dear. Please, listen to what I have to say."
            chatOpt(1) = "*sigh* I suppose I should..."

            For i = 2 To 4
                chatOpt(i) = vbNullString
            Next

        Case 2 ' next
            chatText = "There are some important things you need to know. Here they are. To move, use W, A, S and D. To attack or to talk to someone, press CTRL. To initiate chat press ENTER."
            chatOpt(1) = "Go on..."

            For i = 2 To 4
                chatOpt(i) = vbNullString
            Next

        Case 3 ' chatting
            chatText = "When chatting you can talk in different channels. By default you're talking in the map channel. To talk globally append an apostrophe (') to the start of your message. To perform an emote append a hyphen (-) to the start of your message."
            chatOpt(1) = "Wait, what about combat?"

            For i = 2 To 4
                chatOpt(i) = vbNullString
            Next

        Case 4 ' combat
            chatText = "Combat can be done through melee and skills. You can melee an enemy by facing them and pressing CTRL. To use a skill you can double click it in your skill menu, double click it in the hotbar or use the number keys. (1, 2, 3, etc.)"
            chatOpt(1) = "Oh! What do stats do?"

            For i = 2 To 4
                chatOpt(i) = vbNullString
            Next

        Case 5 ' stats
            chatText = "Strength increases damage and allows you to equip better weaponry. Endurance increases your maximum health. Intelligence increases your maximum spirit. Agility allows you to reduce damage received and also increases critical hit chances. Willpower increase regeneration abilities."
            chatOpt(1) = "Thanks. See you later."

            For i = 2 To 4
                chatOpt(i) = vbNullString
            Next

        Case Else ' goodbye
            chatText = vbNullString

            For i = 1 To 4
                chatOpt(i) = vbNullString
            Next

            SendFinishTutorial
            inTutorial = False
            AddText "Well done, you finished the tutorial.", BrightGreen
            Exit Sub
    End Select

    ' set the state
    tutorialState = stateNum
End Sub

Public Sub ScrollChatBox(ByVal direction As Byte)
    If direction = 0 Then ' up
        If ChatScroll < ChatLines Then
            ChatScroll = ChatScroll + 1
        End If
    Else
        If ChatScroll > 0 Then
            ChatScroll = ChatScroll - 1
        End If
    End If
End Sub

Public Sub ClearMapCache()
    Dim i As Long, filename As String

    For i = 1 To MAX_MAPS
        filename = App.path & "\data files\maps\map" & i & ".map"

        If FileExist(filename) Then
            Kill filename
        End If

    Next

    AddText "Map cache destroyed.", BrightGreen
End Sub

Public Sub AddChatBubble(ByVal target As Long, ByVal TargetType As Byte, ByVal Msg As String, ByVal Colour As Long)
    Dim i As Long, index As Long
    ' set the global index
    chatBubbleIndex = chatBubbleIndex + 1
    
    ' reset to yourself for eventing
    If TargetType = 0 Then
        TargetType = TARGET_TYPE_PLAYER
        If target = 0 Then target = MyIndex
    End If

    If chatBubbleIndex < 1 Or chatBubbleIndex > MAX_BYTE Then chatBubbleIndex = 1
    ' default to new bubble
    index = chatBubbleIndex

    ' loop through and see if that player/npc already has a chat bubble
    For i = 1 To MAX_BYTE
        If chatBubble(i).TargetType = TargetType Then
            If chatBubble(i).target = target Then
                ' reset master index
                If chatBubbleIndex > 1 Then chatBubbleIndex = chatBubbleIndex - 1
                ' we use this one now, yes?
                index = i
                Exit For
            End If
        End If
    Next

    ' set the bubble up
    With chatBubble(index)
        .target = target
        .TargetType = TargetType
        .Msg = Msg
        .Colour = Colour
        .timer = GetTickCount
        .active = True
    End With
End Sub

Public Sub FindNearestTarget()
    Dim i As Long, x As Long, y As Long, x2 As Long, y2 As Long, xDif As Long, yDif As Long
    Dim bestX As Long, bestY As Long, bestIndex As Long
    x2 = GetPlayerX(MyIndex)
    y2 = GetPlayerY(MyIndex)
    bestX = 255
    bestY = 255

    For i = 1 To MAX_MAP_NPCS

        If MapNpc(i).num > 0 Then
            x = MapNpc(i).x
            y = MapNpc(i).y

            ' find the difference - x
            If x < x2 Then
                xDif = x2 - x
            ElseIf x > x2 Then
                xDif = x - x2
            Else
                xDif = 0
            End If

            ' find the difference - y
            If y < y2 Then
                yDif = y2 - y
            ElseIf y > y2 Then
                yDif = y - y2
            Else
                yDif = 0
            End If

            ' best so far?
            If (xDif + yDif) < (bestX + bestY) Then
                bestX = xDif
                bestY = yDif
                bestIndex = i
            End If
        End If

    Next

    ' target the best
    If bestIndex > 0 And bestIndex <> myTarget Then PlayerTarget bestIndex, TARGET_TYPE_NPC
End Sub

Public Sub FindTarget()
    Dim i As Long, x As Long, y As Long

    ' check players
    For i = 1 To MAX_PLAYERS

        If IsPlaying(i) And GetPlayerMap(MyIndex) = GetPlayerMap(i) Then
            x = (GetPlayerX(i) * 32) + Player(i).xOffset + 32
            y = (GetPlayerY(i) * 32) + Player(i).yOffset + 32

            If x >= GlobalX_Map And x <= GlobalX_Map + 32 Then
                If y >= GlobalY_Map And y <= GlobalY_Map + 32 Then
                    ' found our target!
                    PlayerTarget i, TARGET_TYPE_PLAYER
                    Exit Sub
                End If
            End If
        End If

    Next

    ' check npcs
    For i = 1 To MAX_MAP_NPCS

        If MapNpc(i).num > 0 Then
            x = (MapNpc(i).x * 32) + MapNpc(i).xOffset + 32
            y = (MapNpc(i).y * 32) + MapNpc(i).yOffset + 32

            If x >= GlobalX_Map And x <= GlobalX_Map + 32 Then
                If y >= GlobalY_Map And y <= GlobalY_Map + 32 Then
                    ' found our target!
                    PlayerTarget i, TARGET_TYPE_NPC
                    Exit Sub
                End If
            End If
        End If

    Next

End Sub

Public Sub SetBarWidth(ByRef MaxWidth As Long, ByRef width As Long)
    Dim barDifference As Long

    If MaxWidth < width Then
        ' find out the amount to increase per loop
        barDifference = ((width - MaxWidth) / 100) * 10

        ' if it's less than 1 then default to 1
        If barDifference < 1 Then barDifference = 1
        ' set the width
        width = width - barDifference
    ElseIf MaxWidth > width Then
        ' find out the amount to increase per loop
        barDifference = ((MaxWidth - width) / 100) * 10

        ' if it's less than 1 then default to 1
        If barDifference < 1 Then barDifference = 1
        ' set the width
        width = width + barDifference
    End If

End Sub

Public Sub AttemptLogin()
    TcpInit GAME_SERVER_IP, GAME_SERVER_PORT

    ' send login packet
    If ConnectToServer Then
        SendLogin Windows(GetWindowIndex("winLogin")).Controls(GetControlIndex("winLogin", "txtUser")).text
        Exit Sub
    End If

    If Not IsConnected Then
        ShowWindow GetWindowIndex("winLogin")
        Dialogue "Connection Problem", "Cannot connect to game server.", "Please try again later.", TypeALERT
    End If
End Sub

Public Sub DialogueAlert(ByVal index As Long)
    Dim header As String, body As String, body2 As String

    ' find the body/header
    Select Case index

        Case MsgCONNECTION
            header = "Connection Problem"
            body = "You lost connection to the server."
            body2 = "Please try again later."

        Case MsgBANNED
            header = "Banned"
            body = "You have been banned from playing Crystalshire."
            body2 = "Please send all ban appeals to an administrator."

        Case MsgKICKED
            header = "Kicked"
            body = "You have been kicked from Crystalshire."
            body2 = "Please try and behave."

        Case MsgOUTDATED
            header = "Wrong Version"
            body = "Your game client is the wrong version."
            body2 = "Please re-load the game or wait for a patch."

        Case MsgUSERLENGTH
            header = "Invalid Length"
            body = "Your username or password is too short or too long."
            body2 = "Please enter a valid username and password."

        Case MsgILLEGALNAME
            header = "Illegal Characters"
            body = "Your username or password contains illegal characters."
            body2 = "Please enter a valid username and password."

        Case MsgREBOOTING
            header = "Connection Refused"
            body = "The server is currently rebooting."
            body2 = "Please try again soon."

        Case MsgNAMETAKEN
            header = "Invalid Name"
            body = "This name is already in use."
            body2 = "Please try another name."

        Case MsgNAMELENGTH
            header = "Invalid Name"
            body = "This name is too short or too long."
            body2 = "Please try another name."

        Case MsgNAMEILLEGAL
            header = "Invalid Name"
            body = "This name contains illegal characters."
            body2 = "Please try another name."

        Case MsgMYSQL
            header = "Connection Problem"
            body = "Cannot connect to database."
            body2 = "Please try again later."

        Case MsgWRONGPASS
            header = "Invalid Login"
            body = "Invalid username or password."
            body2 = "Please try again."

        Case MsgACTIVATED
            header = "Inactive Account"
            body = "Your account is not activated."
            body2 = "Please activate your account then try again."

        Case MsgMERGE
            header = "Successful Merge"
            body = "Character merged with new account."
            body2 = "Old account permanently destroyed."

        Case MsgMAXCHARS
            header = "Cannot Merge"
            body = "You cannot merge a full account."
            body2 = "Please clear a character slot."

        Case MsgMERGENAME
            header = "Cannot Merge"
            body = "An existing character has this name."
            body2 = "Please contact an administrator."
            
        Case MsgDELCHAR
            header = "Deleted Character"
            body = "Your character was successfully deleted."
            body2 = "Please log on to continue playing."
    End Select

    ' set the dialogue up!
    Dialogue header, body, body2, TypeALERT
End Sub

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

Public Function Clamp(ByVal value As Long, ByVal Min As Long, ByVal Max As Long) As Long
    Clamp = value

    If value < Min Then Clamp = Min
    If value > Max Then Clamp = Max
End Function

Public Sub ShowClasses()
    HideWindows
    newCharClass = 1
    newCharSprite = 1
    newCharGender = SEX_MALE
    Windows(GetWindowIndex("winClasses")).Controls(GetControlIndex("winClasses", "lblClassName")).text = Trim$(Class(newCharClass).name)
    Windows(GetWindowIndex("winNewChar")).Controls(GetControlIndex("winNewChar", "txtName")).text = vbNullString
    Windows(GetWindowIndex("winNewChar")).Controls(GetControlIndex("winNewChar", "chkMale")).value = 1
    Windows(GetWindowIndex("winNewChar")).Controls(GetControlIndex("winNewChar", "chkFemale")).value = 0
    ShowWindow GetWindowIndex("winClasses")
End Sub

Public Sub SetGoldLabel()
Dim i As Long, amount As Long
    amount = 0
    For i = 1 To MAX_INV
        If GetPlayerInvItemNum(MyIndex, i) = 1 Then
            amount = GetPlayerInvItemValue(MyIndex, i)
        End If
    Next
    Windows(GetWindowIndex("winShop")).Controls(GetControlIndex("winShop", "lblGold")).text = Format$(amount, "#,###,###,###") & "g"
    Windows(GetWindowIndex("winInventory")).Controls(GetControlIndex("winInventory", "lblGold")).text = Format$(amount, "#,###,###,###") & "g"
End Sub

Public Sub ShowInvDesc(x As Long, y As Long, invNum As Long)
Dim soulBound As Boolean

    ' rte9
    If invNum <= 0 Or invNum > MAX_INV Then Exit Sub
    
    ' show
    If GetPlayerInvItemNum(MyIndex, invNum) Then
        If Item(GetPlayerInvItemNum(MyIndex, invNum)).BindType > 0 And PlayerInv(invNum).bound > 0 Then soulBound = True
        ShowItemDesc x, y, GetPlayerInvItemNum(MyIndex, invNum), soulBound
    End If
End Sub

Public Sub ShowShopDesc(x As Long, y As Long, itemNum As Long)
    If itemNum <= 0 Or itemNum > MAX_ITEMS Then Exit Sub
    ' show
    ShowItemDesc x, y, itemNum, False
End Sub

Public Sub ShowEqDesc(x As Long, y As Long, eqNum As Long)
Dim soulBound As Boolean

    ' rte9
    If eqNum <= 0 Or eqNum > Equipment.Equipment_Count - 1 Then Exit Sub
    
    ' show
    If Player(MyIndex).Equipment(eqNum) Then
        If Item(Player(MyIndex).Equipment(eqNum)).BindType > 0 Then soulBound = True
        ShowItemDesc x, y, Player(MyIndex).Equipment(eqNum), soulBound
    End If
End Sub

Public Sub ShowPlayerSpellDesc(x As Long, y As Long, slotNum As Long)
    
    ' rte9
    If slotNum <= 0 Or slotNum > MAX_PLAYER_SPELLS Then Exit Sub
    
    ' show
    If PlayerSpells(slotNum).Spell Then
        ShowSpellDesc x, y, PlayerSpells(slotNum).Spell, slotNum
    End If
End Sub

Public Sub ShowSpellDesc(x As Long, y As Long, spellnum As Long, spellSlot As Long)
Dim Colour As Long, theName As String, sUse As String, i As Long, barWidth As Long, tmpWidth As Long

    ' set globals
    descType = 2 ' spell
    descItem = spellnum
    
    ' set position
    Windows(GetWindowIndex("winDescription")).Window.left = x
    Windows(GetWindowIndex("winDescription")).Window.top = y
    
    ' show the window
    ShowWindow GetWindowIndex("winDescription"), , False
    
    ' exit out early if last is same
    If (descLastType = descType) And (descLastItem = descItem) Then Exit Sub
    
    ' clear
    ReDim descText(1 To 1) As TextColourRec
    
    ' hide req. labels
    Windows(GetWindowIndex("winDescription")).Controls(GetControlIndex("winDescription", "lblLevel")).visible = False
    Windows(GetWindowIndex("winDescription")).Controls(GetControlIndex("winDescription", "picBar")).visible = True
    
    ' set variables
    With Windows(GetWindowIndex("winDescription"))
        ' set name
        .Controls(GetControlIndex("winDescription", "lblName")).text = Trim$(Spell(spellnum).name)
        .Controls(GetControlIndex("winDescription", "lblName")).textColour = White
        
        ' find ranks
        If spellSlot > 0 Then
            ' draw the rank bar
            barWidth = 66
            If Spell(spellnum).NextRank > 0 Then
                tmpWidth = ((PlayerSpells(spellSlot).Uses / barWidth) / (Spell(spellnum).NextUses / barWidth)) * barWidth
            Else
                tmpWidth = 66
            End If
            .Controls(GetControlIndex("winDescription", "picBar")).value = tmpWidth
            ' does it rank up?
            If Spell(spellnum).NextRank > 0 Then
                Colour = White
                sUse = "Uses: " & PlayerSpells(spellSlot).Uses & "/" & Spell(spellnum).NextUses
                If PlayerSpells(spellSlot).Uses = Spell(spellnum).NextUses Then
                    If Not GetPlayerLevel(MyIndex) >= Spell(Spell(spellnum).NextRank).LevelReq Then
                        Colour = BrightRed
                        sUse = "Lvl " & Spell(Spell(spellnum).NextRank).LevelReq & " req."
                    End If
                End If
            Else
                Colour = Grey
                sUse = "Max Rank"
            End If
            ' show controls
            .Controls(GetControlIndex("winDescription", "lblClass")).visible = True
            .Controls(GetControlIndex("winDescription", "picBar")).visible = True
             'set vals
            .Controls(GetControlIndex("winDescription", "lblClass")).text = sUse
            .Controls(GetControlIndex("winDescription", "lblClass")).textColour = Colour
        Else
            ' hide some controls
            .Controls(GetControlIndex("winDescription", "lblClass")).visible = False
            .Controls(GetControlIndex("winDescription", "picBar")).visible = False
        End If
    End With
    
    Select Case Spell(spellnum).Type
        Case SPELL_TYPE_DAMAGEHP
            AddDescInfo "Damage HP"
        Case SPELL_TYPE_DAMAGEMP
            AddDescInfo "Damage SP"
        Case SPELL_TYPE_HEALHP
            AddDescInfo "Heal HP"
        Case SPELL_TYPE_HEALMP
            AddDescInfo "Heal SP"
        Case SPELL_TYPE_WARP
            AddDescInfo "Warp"
    End Select
    
    ' more info
    Select Case Spell(spellnum).Type
        Case SPELL_TYPE_DAMAGEHP, SPELL_TYPE_DAMAGEMP, SPELL_TYPE_HEALHP, SPELL_TYPE_HEALMP
            ' damage
            AddDescInfo "Vital: " & Spell(spellnum).Vital
            
            ' mp cost
            AddDescInfo "Cost: " & Spell(spellnum).MPCost & " SP"
            
            ' cast time
            AddDescInfo "Cast Time: " & Spell(spellnum).CastTime & "s"
            
            ' cd time
            AddDescInfo "Cooldown: " & Spell(spellnum).CDTime & "s"
            
            ' aoe
            If Spell(spellnum).AoE > 0 Then
                AddDescInfo "AoE: " & Spell(spellnum).AoE
            End If
            
            ' stun
            If Spell(spellnum).StunDuration > 0 Then
                AddDescInfo "Stun: " & Spell(spellnum).StunDuration & "s"
            End If
            
            ' dot
            If Spell(spellnum).Duration > 0 And Spell(spellnum).Interval > 0 Then
                AddDescInfo "DoT: " & (Spell(spellnum).Duration / Spell(spellnum).Interval) & " tick"
            End If
    End Select
End Sub

Public Sub ShowItemDesc(x As Long, y As Long, itemNum As Long, soulBound As Boolean)
Dim Colour As Long, theName As String, className As String, levelTxt As String, i As Long
    
    ' set globals
    descType = 1 ' inventory
    descItem = itemNum
    
    ' set position
    Windows(GetWindowIndex("winDescription")).Window.left = x
    Windows(GetWindowIndex("winDescription")).Window.top = y
    
    ' show the window
    ShowWindow GetWindowIndex("winDescription"), , False
    
    ' exit out early if last is same
    If (descLastType = descType) And (descLastItem = descItem) Then Exit Sub
    
    ' set last to this
    descLastType = descType
    descLastItem = descItem
    
    ' show req. labels
    Windows(GetWindowIndex("winDescription")).Controls(GetControlIndex("winDescription", "lblClass")).visible = True
    Windows(GetWindowIndex("winDescription")).Controls(GetControlIndex("winDescription", "lblLevel")).visible = True
    Windows(GetWindowIndex("winDescription")).Controls(GetControlIndex("winDescription", "picBar")).visible = False
    
    ' set variables
    With Windows(GetWindowIndex("winDescription"))
        ' name
        If Not soulBound Then
            theName = Trim$(Item(itemNum).name)
        Else
            theName = "(SB) " & Trim$(Item(itemNum).name)
        End If
        .Controls(GetControlIndex("winDescription", "lblName")).text = theName
        Select Case Item(itemNum).Rarity
            Case 0 ' white
                Colour = White
            Case 1 ' green
                Colour = Green
            Case 2 ' blue
                Colour = BrightBlue
            Case 3 ' maroon
                Colour = Red
            Case 4 ' purple
                Colour = Pink
            Case 5 ' orange
                Colour = Brown
        End Select
        .Controls(GetControlIndex("winDescription", "lblName")).textColour = Colour
        ' class req
        If Item(itemNum).ClassReq > 0 Then
            className = Trim$(Class(Item(itemNum).ClassReq).name)
            ' do we match it?
            If GetPlayerClass(MyIndex) = Item(itemNum).ClassReq Then
                Colour = Green
            Else
                Colour = BrightRed
            End If
        ElseIf Item(itemNum).proficiency > 0 Then
            Select Case Item(itemNum).proficiency
                Case 1 ' Sword/Armour
                    If Item(itemNum).Type >= ITEM_TYPE_ARMOR And Item(itemNum).Type <= ITEM_TYPE_SHIELD Then
                        className = "Heavy Armour"
                    ElseIf Item(itemNum).Type = ITEM_TYPE_WEAPON Then
                        className = "Heavy Weapon"
                    End If
                    If hasProficiency(MyIndex, Item(itemNum).proficiency) Then
                        Colour = Green
                    Else
                        Colour = BrightRed
                    End If
                Case 2 ' Staff/Cloth
                    If Item(itemNum).Type >= ITEM_TYPE_ARMOR And Item(itemNum).Type <= ITEM_TYPE_SHIELD Then
                        className = "Cloth Armour"
                    ElseIf Item(itemNum).Type = ITEM_TYPE_WEAPON Then
                        className = "Light Weapon"
                    End If
                    If hasProficiency(MyIndex, Item(itemNum).proficiency) Then
                        Colour = Green
                    Else
                        Colour = BrightRed
                    End If
            End Select
        Else
            className = "No class req."
            Colour = Green
        End If
        .Controls(GetControlIndex("winDescription", "lblClass")).text = className
        .Controls(GetControlIndex("winDescription", "lblClass")).textColour = Colour
        ' level
        If Item(itemNum).LevelReq > 0 Then
            levelTxt = "Level " & Item(itemNum).LevelReq
            ' do we match it?
            If GetPlayerLevel(MyIndex) >= Item(itemNum).LevelReq Then
                Colour = Green
            Else
                Colour = BrightRed
            End If
        Else
            levelTxt = "No level req."
            Colour = Green
        End If
        .Controls(GetControlIndex("winDescription", "lblLevel")).text = levelTxt
        .Controls(GetControlIndex("winDescription", "lblLevel")).textColour = Colour
    End With
    
    ' clear
    ReDim descText(1 To 1) As TextColourRec
    
    ' go through the rest of the text
    Select Case Item(itemNum).Type
        Case ITEM_TYPE_NONE
            AddDescInfo "No type"
        Case ITEM_TYPE_WEAPON
            AddDescInfo "Weapon"
        Case ITEM_TYPE_ARMOR
            AddDescInfo "Armour"
        Case ITEM_TYPE_HELMET
            AddDescInfo "Helmet"
        Case ITEM_TYPE_SHIELD
            AddDescInfo "Shield"
        Case ITEM_TYPE_CONSUME
            AddDescInfo "Consume"
        Case ITEM_TYPE_KEY
            AddDescInfo "Key"
        Case ITEM_TYPE_CURRENCY
            AddDescInfo "Currency"
        Case ITEM_TYPE_SPELL
            AddDescInfo "Spell"
        Case ITEM_TYPE_FOOD
            AddDescInfo "Food"
    End Select
    
    ' more info
    Select Case Item(itemNum).Type
        Case ITEM_TYPE_NONE, ITEM_TYPE_KEY, ITEM_TYPE_CURRENCY
            ' binding
            If Item(itemNum).BindType = 1 Then
                AddDescInfo "Bind on Pickup"
            ElseIf Item(itemNum).BindType = 2 Then
                AddDescInfo "Bind on Equip"
            End If
            ' price
            AddDescInfo "Value: " & Item(itemNum).Price & "g"
        Case ITEM_TYPE_WEAPON, ITEM_TYPE_ARMOR, ITEM_TYPE_HELMET, ITEM_TYPE_SHIELD
            ' damage/defence
            If Item(itemNum).Type = ITEM_TYPE_WEAPON Then
                AddDescInfo "Damage: " & Item(itemNum).Data2
                ' speed
                AddDescInfo "Speed: " & (Item(itemNum).speed / 1000) & "s"
            Else
                If Item(itemNum).Data2 > 0 Then
                    AddDescInfo "Defence: " & Item(itemNum).Data2
                End If
            End If
            ' binding
            If Item(itemNum).BindType = 1 Then
                AddDescInfo "Bind on Pickup"
            ElseIf Item(itemNum).BindType = 2 Then
                AddDescInfo "Bind on Equip"
            End If
            ' price
            AddDescInfo "Value: " & Item(itemNum).Price & "g"
            ' stat bonuses
            If Item(itemNum).Add_Stat(Stats.Strength) > 0 Then
                AddDescInfo "+" & Item(itemNum).Add_Stat(Stats.Strength) & " Str"
            End If
            If Item(itemNum).Add_Stat(Stats.Endurance) > 0 Then
                AddDescInfo "+" & Item(itemNum).Add_Stat(Stats.Endurance) & " End"
            End If
            If Item(itemNum).Add_Stat(Stats.Intelligence) > 0 Then
                AddDescInfo "+" & Item(itemNum).Add_Stat(Stats.Intelligence) & " Int"
            End If
            If Item(itemNum).Add_Stat(Stats.Agility) > 0 Then
                AddDescInfo "+" & Item(itemNum).Add_Stat(Stats.Agility) & " Agi"
            End If
            If Item(itemNum).Add_Stat(Stats.Willpower) > 0 Then
                AddDescInfo "+" & Item(itemNum).Add_Stat(Stats.Willpower) & " Will"
            End If
        Case ITEM_TYPE_CONSUME
            If Item(itemNum).CastSpell > 0 Then
                AddDescInfo "Casts Spell"
            End If
            If Item(itemNum).AddHP > 0 Then
                AddDescInfo "+" & Item(itemNum).AddHP & " HP"
            End If
            If Item(itemNum).AddMP > 0 Then
                AddDescInfo "+" & Item(itemNum).AddMP & " SP"
            End If
            If Item(itemNum).AddEXP > 0 Then
                AddDescInfo "+" & Item(itemNum).AddEXP & " EXP"
            End If
            ' price
            AddDescInfo "Value: " & Item(itemNum).Price & "g"
        Case ITEM_TYPE_SPELL
            ' price
            AddDescInfo "Value: " & Item(itemNum).Price & "g"
        Case ITEM_TYPE_FOOD
            If Item(itemNum).HPorSP = 2 Then
                AddDescInfo "Heal: " & (Item(itemNum).FoodPerTick * Item(itemNum).FoodTickCount) & " SP"
            Else
                AddDescInfo "Heal: " & (Item(itemNum).FoodPerTick * Item(itemNum).FoodTickCount) & " HP"
            End If
            ' time
            AddDescInfo "Time: " & (Item(itemNum).FoodInterval * (Item(itemNum).FoodTickCount / 1000)) & "s"
            ' price
            AddDescInfo "Value: " & Item(itemNum).Price & "g"
    End Select
End Sub

Public Sub AddDescInfo(text As String, Optional Colour As Long = White)
Dim count As Long
    count = UBound(descText)
    ReDim Preserve descText(1 To count + 1) As TextColourRec
    descText(count + 1).text = text
    descText(count + 1).Colour = Colour
End Sub

Public Sub SwitchHotbar(oldSlot As Long, newSlot As Long)
Dim oldSlot_type As Long, oldSlot_value As Long, newSlot_type As Long, newSlot_value As Long

    oldSlot_type = Hotbar(oldSlot).sType
    newSlot_type = Hotbar(newSlot).sType
    oldSlot_value = Hotbar(oldSlot).Slot
    newSlot_value = Hotbar(newSlot).Slot
    
    ' send the changes
    SendHotbarChange oldSlot_type, oldSlot_value, newSlot
    SendHotbarChange newSlot_type, newSlot_value, oldSlot
End Sub

Public Sub ShowChat()
    ShowWindow GetWindowIndex("winChat"), , False
    HideWindow GetWindowIndex("winChatSmall")
    ' Set the active control
    activeWindow = GetWindowIndex("winChat")
    SetActiveControl GetWindowIndex("winChat"), GetControlIndex("winChat", "txtChat")
    inSmallChat = False
    ChatScroll = 0
End Sub

Public Sub HideChat()
    ShowWindow GetWindowIndex("winChatSmall"), , False
    HideWindow GetWindowIndex("winChat")
    inSmallChat = True
    ChatScroll = 0
End Sub

Public Sub SetChatHeight(height As Long)
    actChatHeight = height
End Sub

Public Sub SetChatWidth(width As Long)
    actChatWidth = width
End Sub

Public Sub UpdateChat()
    SaveOptions
End Sub

Sub OpenShop(shopNum As Long)
    ' set globals
    InShop = shopNum
    shopSelectedSlot = 1
    shopSelectedItem = Shop(InShop).TradeItem(1).Item
    Windows(GetWindowIndex("winShop")).Controls(GetControlIndex("winShop", "chkSelling")).value = 0
    Windows(GetWindowIndex("winShop")).Controls(GetControlIndex("winShop", "chkBuying")).value = 1
    Windows(GetWindowIndex("winShop")).Controls(GetControlIndex("winShop", "btnSell")).visible = False
    Windows(GetWindowIndex("winShop")).Controls(GetControlIndex("winShop", "btnBuy")).visible = True
    shopIsSelling = False
    ' set the current item
    UpdateShop
    ' show the window
    ShowWindow GetWindowIndex("winShop")
End Sub

Sub CloseShop()
    SendCloseShop
    HideWindow GetWindowIndex("winShop")
    shopSelectedSlot = 0
    shopSelectedItem = 0
    shopIsSelling = False
    InShop = 0
End Sub

Sub UpdateShop()
Dim i As Long, CostValue As Long

    If InShop = 0 Then Exit Sub
    
    ' make sure we have an item selected
    If shopSelectedSlot = 0 Then shopSelectedSlot = 1
    
    With Windows(GetWindowIndex("winShop"))
        ' buying items
        If Not shopIsSelling Then
            shopSelectedItem = Shop(InShop).TradeItem(shopSelectedSlot).Item
            ' labels
            If shopSelectedItem > 0 Then
                .Controls(GetControlIndex("winShop", "lblName")).text = Trim$(Item(shopSelectedItem).name)
                ' check if it's gold
                If Shop(InShop).TradeItem(shopSelectedSlot).CostItem = 1 Then
                    ' it's gold
                    .Controls(GetControlIndex("winShop", "lblCost")).text = Shop(InShop).TradeItem(shopSelectedSlot).CostValue & "g"
                Else
                    ' if it's one then just print the name
                    If Shop(InShop).TradeItem(shopSelectedSlot).CostValue = 1 Then
                        .Controls(GetControlIndex("winShop", "lblCost")).text = Trim$(Item(Shop(InShop).TradeItem(shopSelectedSlot).CostItem).name)
                    Else
                        .Controls(GetControlIndex("winShop", "lblCost")).text = Shop(InShop).TradeItem(shopSelectedSlot).CostValue & " " & Trim$(Item(Shop(InShop).TradeItem(shopSelectedSlot).CostItem).name)
                    End If
                End If
                ' draw the item
                For i = 0 To 5
                    .Controls(GetControlIndex("winShop", "picItem")).image(i) = Tex_Item(Item(shopSelectedItem).Pic)
                Next
            Else
                .Controls(GetControlIndex("winShop", "lblName")).text = "Empty Slot"
                .Controls(GetControlIndex("winShop", "lblCost")).text = vbNullString
                ' draw the item
                For i = 0 To 5
                    .Controls(GetControlIndex("winShop", "picItem")).image(i) = 0
                Next
            End If
        Else
            shopSelectedItem = GetPlayerInvItemNum(MyIndex, shopSelectedSlot)
            ' labels
            If shopSelectedItem > 0 Then
                .Controls(GetControlIndex("winShop", "lblName")).text = Trim$(Item(shopSelectedItem).name)
                ' calc cost
                CostValue = (Item(shopSelectedItem).Price / 100) * Shop(InShop).BuyRate
                .Controls(GetControlIndex("winShop", "lblCost")).text = CostValue & "g"
                ' draw the item
                For i = 0 To 5
                    .Controls(GetControlIndex("winShop", "picItem")).image(i) = Tex_Item(Item(shopSelectedItem).Pic)
                Next
            Else
                .Controls(GetControlIndex("winShop", "lblName")).text = "Empty Slot"
                .Controls(GetControlIndex("winShop", "lblCost")).text = vbNullString
                ' draw the item
                For i = 0 To 5
                    .Controls(GetControlIndex("winShop", "picItem")).image(i) = 0
                Next
            End If
        End If
    End With
End Sub

Public Function IsShopSlot(startX As Long, startY As Long) As Long
Dim tempRec As RECT
Dim i As Long

    For i = 1 To MAX_TRADES
        With tempRec
            .top = startY + ShopTop + ((ShopOffsetY + 32) * ((i - 1) \ ShopColumns))
            .bottom = .top + PIC_Y
            .left = startX + ShopLeft + ((ShopOffsetX + 32) * (((i - 1) Mod ShopColumns)))
            .Right = .left + PIC_X
        End With

        If currMouseX >= tempRec.left And currMouseX <= tempRec.Right Then
            If currMouseY >= tempRec.top And currMouseY <= tempRec.bottom Then
                IsShopSlot = i
                Exit Function
            End If
        End If
    Next
End Function

Sub ShowPlayerMenu(index As Long, x As Long, y As Long)
    PlayerMenuIndex = index
    If PlayerMenuIndex = 0 Then Exit Sub
    Windows(GetWindowIndex("winPlayerMenu")).Window.left = x - 5
    Windows(GetWindowIndex("winPlayerMenu")).Window.top = y - 5
    Windows(GetWindowIndex("winPlayerMenu")).Controls(GetControlIndex("winPlayerMenu", "btnName")).text = Trim$(GetPlayerName(PlayerMenuIndex))
    ShowWindow GetWindowIndex("winRightClickBG")
    ShowWindow GetWindowIndex("winPlayerMenu"), , False
End Sub

Public Function AryCount(ByRef Ary() As Byte) As Long
On Error Resume Next

    AryCount = UBound(Ary) + 1
End Function

Public Function ByteToInt(ByVal B1 As Long, ByVal B2 As Long) As Long
    ByteToInt = B1 * 256 + B2
End Function

Sub UpdateStats_UI()
    ' set the bar labels
    With Windows(GetWindowIndex("winBars"))
        .Controls(GetControlIndex("winBars", "lblHP")).text = GetPlayerVital(MyIndex, HP) & "/" & GetPlayerMaxVital(MyIndex, HP)
        .Controls(GetControlIndex("winBars", "lblMP")).text = GetPlayerVital(MyIndex, MP) & "/" & GetPlayerMaxVital(MyIndex, MP)
        .Controls(GetControlIndex("winBars", "lblEXP")).text = GetPlayerExp(MyIndex) & "/" & TNL
    End With
    ' update character screen
    With Windows(GetWindowIndex("winCharacter"))
        .Controls(GetControlIndex("winCharacter", "lblHealth")).text = "Health: " & GetPlayerVital(MyIndex, HP) & "/" & GetPlayerMaxVital(MyIndex, HP)
        .Controls(GetControlIndex("winCharacter", "lblSpirit")).text = "Spirit: " & GetPlayerVital(MyIndex, MP) & "/" & GetPlayerMaxVital(MyIndex, MP)
        .Controls(GetControlIndex("winCharacter", "lblExperience")).text = "Experience: " & Player(MyIndex).EXP & "/" & TNL
    End With
End Sub

Sub UpdatePartyInterface()
Dim i As Long, image(0 To 5) As Long, x As Long, pIndex As Long, height As Long, cIn As Long

    ' unload it if we're not in a party
    If Party.Leader = 0 Then
        HideWindow GetWindowIndex("winParty")
        Exit Sub
    End If
    
    ' load the window
    ShowWindow GetWindowIndex("winParty")
    ' fill the controls
    With Windows(GetWindowIndex("winParty"))
        ' clear controls first
        For i = 1 To 3
            .Controls(GetControlIndex("winParty", "lblName" & i)).text = vbNullString
            .Controls(GetControlIndex("winParty", "picEmptyBar_HP" & i)).visible = False
            .Controls(GetControlIndex("winParty", "picEmptyBar_SP" & i)).visible = False
            .Controls(GetControlIndex("winParty", "picBar_HP" & i)).visible = False
            .Controls(GetControlIndex("winParty", "picBar_SP" & i)).visible = False
            .Controls(GetControlIndex("winParty", "picShadow" & i)).visible = False
            .Controls(GetControlIndex("winParty", "picChar" & i)).visible = False
            .Controls(GetControlIndex("winParty", "picChar" & i)).value = 0
        Next
        ' labels
        cIn = 1
        For i = 1 To Party.MemberCount
            ' cache the index
            pIndex = Party.Member(i)
            If pIndex > 0 Then
                If pIndex <> MyIndex Then
                    If IsPlaying(pIndex) Then
                        ' name and level
                        .Controls(GetControlIndex("winParty", "lblName" & cIn)).visible = True
                        .Controls(GetControlIndex("winParty", "lblName" & cIn)).text = Trim$(GetPlayerName(pIndex))
                        ' picture
                        .Controls(GetControlIndex("winParty", "picShadow" & cIn)).visible = True
                        .Controls(GetControlIndex("winParty", "picChar" & cIn)).visible = True
                        ' store the player's index as a value for later use
                        .Controls(GetControlIndex("winParty", "picChar" & cIn)).value = pIndex
                        For x = 0 To 5
                            .Controls(GetControlIndex("winParty", "picChar" & cIn)).image(x) = Tex_Char(GetPlayerSprite(pIndex))
                        Next
                        ' bars
                        .Controls(GetControlIndex("winParty", "picEmptyBar_HP" & cIn)).visible = True
                        .Controls(GetControlIndex("winParty", "picEmptyBar_SP" & cIn)).visible = True
                        .Controls(GetControlIndex("winParty", "picBar_HP" & cIn)).visible = True
                        .Controls(GetControlIndex("winParty", "picBar_SP" & cIn)).visible = True
                        ' increment control usage
                        cIn = cIn + 1
                    End If
                End If
            End If
        Next
        ' update the bars
        UpdatePartyBars
        ' set the window size
        Select Case Party.MemberCount
            Case 2: height = 78
            Case 3: height = 118
            Case 4: height = 158
        End Select
        .Window.height = height
    End With
End Sub

Sub UpdatePartyBars()
Dim i As Long, pIndex As Long, barWidth As Long, width As Long

    ' unload it if we're not in a party
    If Party.Leader = 0 Then
        Exit Sub
    End If
    
    ' max bar width
    barWidth = 173
    
    ' make sure we're in a party
    With Windows(GetWindowIndex("winParty"))
        For i = 1 To 3
            ' get the pIndex from the control
            If .Controls(GetControlIndex("winParty", "picChar" & i)).visible = True Then
                pIndex = .Controls(GetControlIndex("winParty", "picChar" & i)).value
                ' make sure they exist
                If pIndex > 0 Then
                    If IsPlaying(pIndex) Then
                        ' get their health
                        If GetPlayerVital(pIndex, HP) > 0 And GetPlayerMaxVital(pIndex, HP) > 0 Then
                            width = ((GetPlayerVital(pIndex, Vitals.HP) / barWidth) / (GetPlayerMaxVital(pIndex, Vitals.HP) / barWidth)) * barWidth
                            .Controls(GetControlIndex("winParty", "picBar_HP" & i)).width = width
                        Else
                            .Controls(GetControlIndex("winParty", "picBar_HP" & i)).width = 0
                        End If
                        ' get their spirit
                        If GetPlayerVital(pIndex, MP) > 0 And GetPlayerMaxVital(pIndex, MP) > 0 Then
                            width = ((GetPlayerVital(pIndex, Vitals.MP) / barWidth) / (GetPlayerMaxVital(pIndex, Vitals.MP) / barWidth)) * barWidth
                            .Controls(GetControlIndex("winParty", "picBar_SP" & i)).width = width
                        Else
                            .Controls(GetControlIndex("winParty", "picBar_SP" & i)).width = 0
                        End If
                    End If
                End If
            End If
        Next
    End With
End Sub

Sub ShowTrade()
    ' show the window
    ShowWindow GetWindowIndex("winTrade")
    ' set the controls up
    With Windows(GetWindowIndex("winTrade"))
        .Window.text = "Trading with " & Trim$(GetPlayerName(InTrade))
        .Controls(GetControlIndex("winTrade", "lblYourTrade")).text = Trim$(GetPlayerName(MyIndex)) & "'s Offer"
        .Controls(GetControlIndex("winTrade", "lblTheirTrade")).text = Trim$(GetPlayerName(InTrade)) & "'s Offer"
        .Controls(GetControlIndex("winTrade", "lblYourValue")).text = "0g"
        .Controls(GetControlIndex("winTrade", "lblTheirValue")).text = "0g"
        .Controls(GetControlIndex("winTrade", "lblStatus")).text = "Choose items to offer."
    End With
End Sub

Sub CheckResolution()
Dim Resolution As Byte, width As Long, height As Long
    ' find the selected resolution
    Resolution = Options.Resolution
    ' reset
    If Resolution = 0 Then
        Resolution = 12
        ' loop through till we find one which fits
        Do Until ScreenFit(Resolution) Or Resolution > RES_COUNT
            ScreenFit Resolution
            Resolution = Resolution + 1
            DoEvents
        Loop
        ' right resolution
        If Resolution > RES_COUNT Then Resolution = RES_COUNT
        Options.Resolution = Resolution
    End If
    
    ' size the window
    GetResolutionSize Options.Resolution, width, height
    Resize width, height
    
    ' save it
    curResolution = Options.Resolution
    
    SaveOptions
End Sub

Function ScreenFit(Resolution As Byte) As Boolean
Dim sWidth As Long, sHeight As Long, width As Long, height As Long

    ' exit out early
    If Resolution = 0 Then
        ScreenFit = False
        Exit Function
    End If

    ' get screen size
    sWidth = Screen.width / Screen.TwipsPerPixelX
    sHeight = Screen.height / Screen.TwipsPerPixelY
    
    GetResolutionSize Resolution, width, height
    
    ' check if match
    If width > sWidth Or height > sHeight Then
        ScreenFit = False
    Else
        ScreenFit = True
    End If
End Function

Function GetResolutionSize(Resolution As Byte, ByRef width As Long, ByRef height As Long)
    Select Case Resolution
        Case 1
            width = 1920
            height = 1080
        Case 2
            width = 1680
            height = 1050
        Case 3
            width = 1600
            height = 900
        Case 4
            width = 1440
            height = 900
        Case 5
            width = 1440
            height = 1050
        Case 6
            width = 1366
            height = 768
        Case 7
            width = 1360
            height = 1024
        Case 8
            width = 1360
            height = 768
        Case 9
            width = 1280
            height = 1024
        Case 10
            width = 1280
            height = 800
        Case 11
            width = 1280
            height = 768
        Case 12
            width = 1280
            height = 720
        Case 13
            width = 1024
            height = 768
        Case 14
            width = 1024
            height = 576
        Case 15
            width = 800
            height = 600
        Case 16
            width = 800
            height = 450
    End Select
End Function

Sub Resize(ByVal width As Long, ByVal height As Long)
    frmMain.width = (frmMain.width \ 15 - frmMain.ScaleWidth + width) * 15
    frmMain.height = (frmMain.height \ 15 - frmMain.ScaleHeight + height) * 15
    frmMain.left = (Screen.width - frmMain.width) \ 2
    frmMain.top = (Screen.height - frmMain.height) \ 2
    DoEvents
End Sub

Sub ResizeGUI()
Dim top As Long

    ' move hotbar
    Windows(GetWindowIndex("winHotbar")).Window.left = ScreenWidth - 430
    ' move chat
    Windows(GetWindowIndex("winChat")).Window.top = ScreenHeight - 178
    Windows(GetWindowIndex("winChatSmall")).Window.top = ScreenHeight - 162
    ' move menu
    Windows(GetWindowIndex("winMenu")).Window.left = ScreenWidth - 236
    Windows(GetWindowIndex("winMenu")).Window.top = ScreenHeight - 37
    ' move invitations
    Windows(GetWindowIndex("winInvite_Party")).Window.left = ScreenWidth - 234
    Windows(GetWindowIndex("winInvite_Party")).Window.top = ScreenHeight - 80
    ' loop through
    top = ScreenHeight - 80
    If Windows(GetWindowIndex("winInvite_Party")).Window.visible Then
        top = top - 37
    End If
    Windows(GetWindowIndex("winInvite_Trade")).Window.left = ScreenWidth - 234
    Windows(GetWindowIndex("winInvite_Trade")).Window.top = top
    ' re-size right-click background
    Windows(GetWindowIndex("winRightClickBG")).Window.width = ScreenWidth
    Windows(GetWindowIndex("winRightClickBG")).Window.height = ScreenHeight
    ' re-size black background
    Windows(GetWindowIndex("winBlank")).Window.width = ScreenWidth
    Windows(GetWindowIndex("winBlank")).Window.height = ScreenHeight
    ' re-size combo background
    Windows(GetWindowIndex("winComboMenuBG")).Window.width = ScreenWidth
    Windows(GetWindowIndex("winComboMenuBG")).Window.height = ScreenHeight
    ' centralise windows
    CentraliseWindow GetWindowIndex("winLogin")
    CentraliseWindow GetWindowIndex("winCharacters")
    CentraliseWindow GetWindowIndex("winLoading")
    CentraliseWindow GetWindowIndex("winDialogue")
    CentraliseWindow GetWindowIndex("winClasses")
    CentraliseWindow GetWindowIndex("winNewChar")
    CentraliseWindow GetWindowIndex("winEscMenu")
    CentraliseWindow GetWindowIndex("winInventory")
    CentraliseWindow GetWindowIndex("winCharacter")
    CentraliseWindow GetWindowIndex("winSkills")
    CentraliseWindow GetWindowIndex("winOptions")
    CentraliseWindow GetWindowIndex("winShop")
    CentraliseWindow GetWindowIndex("winNpcChat")
    CentraliseWindow GetWindowIndex("winTrade")
    CentraliseWindow GetWindowIndex("winGuild")
End Sub

Sub SetResolution()
Dim width As Long, height As Long
    curResolution = Options.Resolution
    GetResolutionSize curResolution, width, height
    Resize width, height
    ScreenWidth = width
    ScreenHeight = height
    TileWidth = (width / 32) - 1
    TileHeight = (height / 32) - 1
    ScreenX = (TileWidth) * PIC_X
    ScreenY = (TileHeight) * PIC_Y
    ResetGFX
    ResizeGUI
End Sub

Sub ShowComboMenu(curWindow As Long, curControl As Long)
Dim top As Long
    With Windows(curWindow).Controls(curControl)
        ' linked to
        Windows(GetWindowIndex("winComboMenu")).Window.linkedToWin = curWindow
        Windows(GetWindowIndex("winComboMenu")).Window.linkedToCon = curControl
        ' set the size
        Windows(GetWindowIndex("winComboMenu")).Window.height = 2 + (UBound(.list) * 16)
        Windows(GetWindowIndex("winComboMenu")).Window.left = Windows(curWindow).Window.left + .left + 2
        top = Windows(curWindow).Window.top + .top + .height
        If top + Windows(GetWindowIndex("winComboMenu")).Window.height > ScreenHeight Then top = ScreenHeight - Windows(GetWindowIndex("winComboMenu")).Window.height
        Windows(GetWindowIndex("winComboMenu")).Window.top = top
        Windows(GetWindowIndex("winComboMenu")).Window.width = .width - 4
        ' set the values
        Windows(GetWindowIndex("winComboMenu")).Window.list() = .list()
        Windows(GetWindowIndex("winComboMenu")).Window.value = .value
        Windows(GetWindowIndex("winComboMenu")).Window.group = 0
        ' load the menu
        ShowWindow GetWindowIndex("winComboMenuBG"), True, False
        ShowWindow GetWindowIndex("winComboMenu"), True, False
    End With
End Sub

Sub ComboMenu_MouseMove(curWindow As Long)
Dim y As Long, i As Long
    With Windows(curWindow).Window
        y = currMouseY - .top
        ' find the option we're hovering over
        If UBound(.list) > 0 Then
            For i = 1 To UBound(.list)
                If y >= (16 * (i - 1)) And y <= (16 * (i)) Then
                    .group = i
                End If
            Next
        End If
    End With
End Sub

Sub ComboMenu_MouseDown(curWindow As Long)
Dim y As Long, i As Long
    With Windows(curWindow).Window
        y = currMouseY - .top
        ' find the option we're hovering over
        If UBound(.list) > 0 Then
            For i = 1 To UBound(.list)
                If y >= (16 * (i - 1)) And y <= (16 * (i)) Then
                    Windows(.linkedToWin).Controls(.linkedToCon).value = i
                    CloseComboMenu
                End If
            Next
        End If
    End With
End Sub

Sub SetOptionsScreen()
    ' clear the combolists
    Erase Windows(GetWindowIndex("winOptions")).Controls(GetControlIndex("winOptions", "cmbRes")).list
    ReDim Windows(GetWindowIndex("winOptions")).Controls(GetControlIndex("winOptions", "cmbRes")).list(0)
    Erase Windows(GetWindowIndex("winOptions")).Controls(GetControlIndex("winOptions", "cmbRender")).list
    ReDim Windows(GetWindowIndex("winOptions")).Controls(GetControlIndex("winOptions", "cmbRender")).list(0)
    
    ' Resolutions
    Combobox_AddItem GetWindowIndex("winOptions"), GetControlIndex("winOptions", "cmbRes"), "1920x1080"
    Combobox_AddItem GetWindowIndex("winOptions"), GetControlIndex("winOptions", "cmbRes"), "1680x1050"
    Combobox_AddItem GetWindowIndex("winOptions"), GetControlIndex("winOptions", "cmbRes"), "1600x900"
    Combobox_AddItem GetWindowIndex("winOptions"), GetControlIndex("winOptions", "cmbRes"), "1440x900"
    Combobox_AddItem GetWindowIndex("winOptions"), GetControlIndex("winOptions", "cmbRes"), "1440x1050"
    Combobox_AddItem GetWindowIndex("winOptions"), GetControlIndex("winOptions", "cmbRes"), "1366x768"
    Combobox_AddItem GetWindowIndex("winOptions"), GetControlIndex("winOptions", "cmbRes"), "1360x1024"
    Combobox_AddItem GetWindowIndex("winOptions"), GetControlIndex("winOptions", "cmbRes"), "1360x768"
    Combobox_AddItem GetWindowIndex("winOptions"), GetControlIndex("winOptions", "cmbRes"), "1280x1024"
    Combobox_AddItem GetWindowIndex("winOptions"), GetControlIndex("winOptions", "cmbRes"), "1280x800"
    Combobox_AddItem GetWindowIndex("winOptions"), GetControlIndex("winOptions", "cmbRes"), "1280x768"
    Combobox_AddItem GetWindowIndex("winOptions"), GetControlIndex("winOptions", "cmbRes"), "1280x720"
    Combobox_AddItem GetWindowIndex("winOptions"), GetControlIndex("winOptions", "cmbRes"), "1024x768"
    Combobox_AddItem GetWindowIndex("winOptions"), GetControlIndex("winOptions", "cmbRes"), "1024x576"
    Combobox_AddItem GetWindowIndex("winOptions"), GetControlIndex("winOptions", "cmbRes"), "800x600"
    Combobox_AddItem GetWindowIndex("winOptions"), GetControlIndex("winOptions", "cmbRes"), "800x450"
    
    ' Render Options
    Combobox_AddItem GetWindowIndex("winOptions"), GetControlIndex("winOptions", "cmbRender"), "Automatic"
    Combobox_AddItem GetWindowIndex("winOptions"), GetControlIndex("winOptions", "cmbRender"), "Hardware"
    Combobox_AddItem GetWindowIndex("winOptions"), GetControlIndex("winOptions", "cmbRender"), "Mixed"
    Combobox_AddItem GetWindowIndex("winOptions"), GetControlIndex("winOptions", "cmbRender"), "Software"
    
    ' fill the options screen
    With Windows(GetWindowIndex("winOptions"))
        .Controls(GetControlIndex("winOptions", "chkMusic")).value = Options.Music
        .Controls(GetControlIndex("winOptions", "chkSound")).value = Options.sound
        If Options.NoAuto = 1 Then
            .Controls(GetControlIndex("winOptions", "chkAutotiles")).value = 0
        Else
            .Controls(GetControlIndex("winOptions", "chkAutotiles")).value = 1
        End If
        .Controls(GetControlIndex("winOptions", "chkFullscreen")).value = Options.Fullscreen
        .Controls(GetControlIndex("winOptions", "cmbRes")).value = Options.Resolution
        .Controls(GetControlIndex("winOptions", "cmbRender")).value = Options.Render + 1
    End With
End Sub

Sub EventLogic()
Dim target As Long
    ' carry out the command
    With map.TileData.Events(eventNum).EventPage(eventPageNum)
        Select Case .Commands(eventCommandNum).Type
            Case EventType.evAddText
                AddText .Commands(eventCommandNum).text, .Commands(eventCommandNum).Colour, , .Commands(eventCommandNum).channel
            Case EventType.evShowChatBubble
                If .Commands(eventCommandNum).TargetType = TARGET_TYPE_PLAYER Then target = MyIndex Else target = .Commands(eventCommandNum).target
                AddChatBubble target, .Commands(eventCommandNum).TargetType, .Commands(eventCommandNum).text, .Commands(eventCommandNum).Colour
            Case EventType.evPlayerVar
                If .Commands(eventCommandNum).target > 0 Then Player(MyIndex).Variable(.Commands(eventCommandNum).target) = .Commands(eventCommandNum).Colour
        End Select
        ' increment commands
        If eventCommandNum < .CommandCount Then
            eventCommandNum = eventCommandNum + 1
            Exit Sub
        End If
    End With
    ' we're done - close event
    eventNum = 0
    eventPageNum = 0
    eventCommandNum = 0
    inEvent = False
End Sub

Function HasItem(ByVal itemNum As Long) As Long
    Dim i As Long

    For i = 1 To MAX_INV
        ' Check to see if the player has the item
        If GetPlayerInvItemNum(MyIndex, i) = itemNum Then
            If Item(itemNum).Type = ITEM_TYPE_CURRENCY Then
                HasItem = GetPlayerInvItemValue(MyIndex, i)
            Else
                HasItem = 1
            End If
            Exit Function
        End If
    Next
End Function

Function ActiveEventPage(ByVal eventNum As Long) As Long
Dim x As Long, process As Boolean
    For x = map.TileData.Events(eventNum).pageCount To 1 Step -1
        ' check if we match
        With map.TileData.Events(eventNum).EventPage(x)
            process = True
            ' player var check
            If .chkPlayerVar Then
                If .PlayerVarNum > 0 Then
                    If Player(MyIndex).Variable(.PlayerVarNum) < .PlayerVariable Then
                        process = False
                    End If
                End If
            End If
            ' has item check
            If .chkHasItem Then
                If .HasItemNum > 0 Then
                    If HasItem(.HasItemNum) = 0 Then
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

Sub PlayerSwitchInvSlots(ByVal oldSlot As Long, ByVal newSlot As Long)
Dim OldNum As Long, OldValue As Long, oldBound As Byte
Dim NewNum As Long, NewValue As Long, newBound As Byte

    If oldSlot = 0 Or newSlot = 0 Then
        Exit Sub
    End If

    OldNum = GetPlayerInvItemNum(MyIndex, oldSlot)
    OldValue = GetPlayerInvItemValue(MyIndex, oldSlot)
    oldBound = PlayerInv(oldSlot).bound
    NewNum = GetPlayerInvItemNum(MyIndex, newSlot)
    NewValue = GetPlayerInvItemValue(MyIndex, newSlot)
    newBound = PlayerInv(newSlot).bound
    
    SetPlayerInvItemNum MyIndex, newSlot, OldNum
    SetPlayerInvItemValue MyIndex, newSlot, OldValue
    PlayerInv(newSlot).bound = oldBound
    
    SetPlayerInvItemNum MyIndex, oldSlot, NewNum
    SetPlayerInvItemValue MyIndex, oldSlot, NewValue
    PlayerInv(oldSlot).bound = newBound
End Sub

Sub PlayerSwitchSpellSlots(ByVal oldSlot As Long, ByVal newSlot As Long)
Dim OldNum As Long, NewNum As Long, OldUses As Long, NewUses As Long

    If oldSlot = 0 Or newSlot = 0 Then
        Exit Sub
    End If

    OldNum = PlayerSpells(oldSlot).Spell
    NewNum = PlayerSpells(newSlot).Spell
    OldUses = PlayerSpells(oldSlot).Uses
    NewUses = PlayerSpells(newSlot).Uses
    
    PlayerSpells(oldSlot).Spell = NewNum
    PlayerSpells(oldSlot).Uses = NewUses
    PlayerSpells(newSlot).Spell = OldNum
    PlayerSpells(newSlot).Uses = OldUses
End Sub

Sub CheckAppearTiles()
Dim x As Long, y As Long, i As Long
    If GettingMap Then Exit Sub
    
    ' clear
    For x = 0 To map.MapData.MaxX
        For y = 0 To map.MapData.MaxY
            If map.TileData.Tile(x, y).Type = TILE_TYPE_APPEAR Then
                TempTile(x, y).DoorOpen = 0
            End If
        Next
    Next
    
    ' set
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            If GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                x = GetPlayerX(i)
                y = GetPlayerY(i)
                CheckAppearTile x, y
                If y - 1 >= 0 Then CheckAppearTile x, y - 1
                If y + 1 <= map.MapData.MaxY Then CheckAppearTile x, y + 1
                If x - 1 >= 0 Then CheckAppearTile x - 1, y
                If x + 1 <= map.MapData.MaxX Then CheckAppearTile x + 1, y
            End If
        End If
    Next
    For i = 1 To MAX_MAP_NPCS
        If MapNpc(i).num > 0 Then
            If MapNpc(i).Vital(Vitals.HP) > 0 Then
                x = MapNpc(i).x
                y = MapNpc(i).y
                CheckAppearTile x, y
                If y - 1 >= 0 Then CheckAppearTile x, y - 1
                If y + 1 <= map.MapData.MaxY Then CheckAppearTile x, y + 1
                If x - 1 >= 0 Then CheckAppearTile x - 1, y
                If x + 1 <= map.MapData.MaxX Then CheckAppearTile x + 1, y
            End If
        End If
    Next
    
    ' fade out old
    For x = 0 To map.MapData.MaxX
        For y = 0 To map.MapData.MaxY
            If TempTile(x, y).DoorOpen = 0 Then
                ' exit if our mother is a bottom
                If y > 0 Then
                    If map.TileData.Tile(x, y - 1).Data2 Then
                        If TempTile(x, y - 1).DoorOpen = 1 Then GoTo continueLoop
                    End If
                End If
                ' not open - fade them out
                For i = 1 To MapLayer.Layer_Count - 1
                    If TempTile(x, y).fadeAlpha(i) > 0 Then
                        TempTile(x, y).isFading(i) = True
                        TempTile(x, y).fadeAlpha(i) = TempTile(x, y).fadeAlpha(i) - 1
                        TempTile(x, y).FadeDir(i) = DIR_DOWN
                    End If
                Next
            End If
continueLoop:
        Next
    Next
End Sub

Sub CheckAppearTile(ByVal x As Long, ByVal y As Long)
    If y < 0 Or x < 0 Or y > map.MapData.MaxY Or x > map.MapData.MaxX Then Exit Sub
    
    If map.TileData.Tile(x, y).Type = TILE_TYPE_APPEAR Then
        TempTile(x, y).DoorOpen = 1
        
        If TempTile(x, y).fadeAlpha(MapLayer.Mask) = 255 Then Exit Sub
        If TempTile(x, y).isFading(MapLayer.Mask) Then
            If TempTile(x, y).FadeDir(MapLayer.Mask) = DIR_DOWN Then
                TempTile(x, y).FadeDir(MapLayer.Mask) = DIR_UP
                ' check if bottom
                If y < map.MapData.MaxY Then
                    If map.TileData.Tile(x, y).Data2 Then
                        TempTile(x, y + 1).FadeDir(MapLayer.Ground) = DIR_UP
                    End If
                End If
                ' / bottom
            End If
            Exit Sub
        End If
        
        TempTile(x, y).FadeDir(MapLayer.Mask) = DIR_UP
        TempTile(x, y).isFading(MapLayer.Mask) = True
        TempTile(x, y).fadeAlpha(MapLayer.Mask) = TempTile(x, y).fadeAlpha(MapLayer.Mask) + 1
        
        ' check if bottom
        If y < map.MapData.MaxY Then
            If map.TileData.Tile(x, y).Data2 Then
                TempTile(x, y + 1).FadeDir(MapLayer.Ground) = DIR_UP
                TempTile(x, y + 1).isFading(MapLayer.Ground) = True
                TempTile(x, y + 1).fadeAlpha(MapLayer.Ground) = TempTile(x, y + 1).fadeAlpha(MapLayer.Ground) + 1
            End If
        End If
        ' / bottom
    End If
End Sub

Public Sub AppearTileFadeLogic()
Dim x As Long, y As Long
    For x = 0 To map.MapData.MaxX
        For y = 0 To map.MapData.MaxY
            If map.TileData.Tile(x, y).Type = TILE_TYPE_APPEAR Then
                ' check if it's fading
                If TempTile(x, y).isFading(MapLayer.Mask) Then
                    ' fading in
                    If TempTile(x, y).FadeDir(MapLayer.Mask) = DIR_UP Then
                        If TempTile(x, y).fadeAlpha(MapLayer.Mask) < 255 Then
                            TempTile(x, y).fadeAlpha(MapLayer.Mask) = TempTile(x, y).fadeAlpha(MapLayer.Mask) + 20
                            ' check if bottom
                            If y < map.MapData.MaxY Then
                                If map.TileData.Tile(x, y).Data2 Then
                                    TempTile(x, y + 1).fadeAlpha(MapLayer.Ground) = TempTile(x, y + 1).fadeAlpha(MapLayer.Ground) + 20
                                End If
                            End If
                            ' / bottom
                        End If
                        If TempTile(x, y).fadeAlpha(MapLayer.Mask) >= 255 Then
                            TempTile(x, y).fadeAlpha(MapLayer.Mask) = 255
                            TempTile(x, y).isFading(MapLayer.Mask) = False
                            ' check if bottom
                            If y < map.MapData.MaxY Then
                                If map.TileData.Tile(x, y).Data2 Then
                                    TempTile(x, y + 1).fadeAlpha(MapLayer.Ground) = 255
                                    TempTile(x, y + 1).isFading(MapLayer.Ground) = False
                                End If
                            End If
                            ' / bottom
                        End If
                    ElseIf TempTile(x, y).FadeDir(MapLayer.Mask) = DIR_DOWN Then
                        If TempTile(x, y).fadeAlpha(MapLayer.Mask) > 0 Then
                            TempTile(x, y).fadeAlpha(MapLayer.Mask) = TempTile(x, y).fadeAlpha(MapLayer.Mask) - 20
                            ' check if bottom
                            If y < map.MapData.MaxY Then
                                If map.TileData.Tile(x, y).Data2 Then
                                    TempTile(x, y + 1).fadeAlpha(MapLayer.Ground) = TempTile(x, y + 1).fadeAlpha(MapLayer.Ground) - 20
                                End If
                            End If
                            ' / bottom
                        End If
                        If TempTile(x, y).fadeAlpha(MapLayer.Mask) <= 0 Then
                            TempTile(x, y).fadeAlpha(MapLayer.Mask) = 0
                            TempTile(x, y).isFading(MapLayer.Mask) = False
                            ' check if bottom
                            If y < map.MapData.MaxY Then
                                If map.TileData.Tile(x, y).Data2 Then
                                    TempTile(x, y + 1).fadeAlpha(MapLayer.Ground) = 0
                                    TempTile(x, y + 1).isFading(MapLayer.Ground) = False
                                End If
                            End If
                            ' / bottom
                        End If
                    End If
                End If
            End If
        Next
    Next
End Sub
