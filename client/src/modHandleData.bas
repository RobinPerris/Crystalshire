Attribute VB_Name = "modHandleData"
Option Explicit

' ******************************************
' ** Parses and handles String packets    **
' ******************************************
Public Function GetAddress(FunAddr As Long) As Long
    GetAddress = FunAddr
End Function

Public Sub InitMessages()
    HandleDataSub(SAlertMsg) = GetAddress(AddressOf HandleAlertMsg)
    HandleDataSub(SLoginOk) = GetAddress(AddressOf HandleLoginOk)
    HandleDataSub(SNewCharClasses) = GetAddress(AddressOf HandleNewCharClasses)
    HandleDataSub(SClassesData) = GetAddress(AddressOf HandleClassesData)
    HandleDataSub(SInGame) = GetAddress(AddressOf HandleInGame)
    HandleDataSub(SPlayerInv) = GetAddress(AddressOf HandlePlayerInv)
    HandleDataSub(SPlayerInvUpdate) = GetAddress(AddressOf HandlePlayerInvUpdate)
    HandleDataSub(SPlayerWornEq) = GetAddress(AddressOf HandlePlayerWornEq)
    HandleDataSub(SPlayerHp) = GetAddress(AddressOf HandlePlayerHp)
    HandleDataSub(SPlayerMp) = GetAddress(AddressOf HandlePlayerMp)
    HandleDataSub(SPlayerStats) = GetAddress(AddressOf HandlePlayerStats)
    HandleDataSub(SPlayerData) = GetAddress(AddressOf HandlePlayerData)
    HandleDataSub(SPlayerMove) = GetAddress(AddressOf HandlePlayerMove)
    HandleDataSub(SNpcMove) = GetAddress(AddressOf HandleNpcMove)
    HandleDataSub(SPlayerDir) = GetAddress(AddressOf HandlePlayerDir)
    HandleDataSub(SNpcDir) = GetAddress(AddressOf HandleNpcDir)
    HandleDataSub(SPlayerXY) = GetAddress(AddressOf HandlePlayerXY)
    HandleDataSub(SPlayerXYMap) = GetAddress(AddressOf HandlePlayerXYMap)
    HandleDataSub(SAttack) = GetAddress(AddressOf HandleAttack)
    HandleDataSub(SNpcAttack) = GetAddress(AddressOf HandleNpcAttack)
    HandleDataSub(SCheckForMap) = GetAddress(AddressOf HandleCheckForMap)
    HandleDataSub(SMapData) = GetAddress(AddressOf HandleMapData)
    HandleDataSub(SMapItemData) = GetAddress(AddressOf HandleMapItemData)
    HandleDataSub(SMapNpcData) = GetAddress(AddressOf HandleMapNpcData)
    HandleDataSub(SMapDone) = GetAddress(AddressOf HandleMapDone)
    HandleDataSub(SGlobalMsg) = GetAddress(AddressOf HandleGlobalMsg)
    HandleDataSub(SAdminMsg) = GetAddress(AddressOf HandleAdminMsg)
    HandleDataSub(SPlayerMsg) = GetAddress(AddressOf HandlePlayerMsg)
    HandleDataSub(SMapMsg) = GetAddress(AddressOf HandleMapMsg)
    HandleDataSub(SSpawnItem) = GetAddress(AddressOf HandleSpawnItem)
    HandleDataSub(SItemEditor) = GetAddress(AddressOf HandleItemEditor)
    HandleDataSub(SUpdateItem) = GetAddress(AddressOf HandleUpdateItem)
    HandleDataSub(SSpawnNpc) = GetAddress(AddressOf HandleSpawnNpc)
    HandleDataSub(SNpcDead) = GetAddress(AddressOf HandleNpcDead)
    HandleDataSub(SNpcEditor) = GetAddress(AddressOf HandleNpcEditor)
    HandleDataSub(SUpdateNpc) = GetAddress(AddressOf HandleUpdateNpc)
    HandleDataSub(SMapKey) = GetAddress(AddressOf HandleMapKey)
    HandleDataSub(SEditMap) = GetAddress(AddressOf HandleEditMap)
    HandleDataSub(SShopEditor) = GetAddress(AddressOf HandleShopEditor)
    HandleDataSub(SUpdateShop) = GetAddress(AddressOf HandleUpdateShop)
    HandleDataSub(SSpellEditor) = GetAddress(AddressOf HandleSpellEditor)
    HandleDataSub(SUpdateSpell) = GetAddress(AddressOf HandleUpdateSpell)
    HandleDataSub(SSpells) = GetAddress(AddressOf HandleSpells)
    HandleDataSub(SLeft) = GetAddress(AddressOf HandleLeft)
    HandleDataSub(SResourceCache) = GetAddress(AddressOf HandleResourceCache)
    HandleDataSub(SResourceEditor) = GetAddress(AddressOf HandleResourceEditor)
    HandleDataSub(SUpdateResource) = GetAddress(AddressOf HandleUpdateResource)
    HandleDataSub(SSendPing) = GetAddress(AddressOf HandleSendPing)
    HandleDataSub(SDoorAnimation) = GetAddress(AddressOf HandleDoorAnimation)
    HandleDataSub(SActionMsg) = GetAddress(AddressOf HandleActionMsg)
    HandleDataSub(SPlayerEXP) = GetAddress(AddressOf HandlePlayerExp)
    HandleDataSub(SBlood) = GetAddress(AddressOf HandleBlood)
    HandleDataSub(SAnimationEditor) = GetAddress(AddressOf HandleAnimationEditor)
    HandleDataSub(SUpdateAnimation) = GetAddress(AddressOf HandleUpdateAnimation)
    HandleDataSub(SAnimation) = GetAddress(AddressOf HandleAnimation)
    HandleDataSub(SMapNpcVitals) = GetAddress(AddressOf HandleMapNpcVitals)
    HandleDataSub(SCooldown) = GetAddress(AddressOf HandleCooldown)
    HandleDataSub(SClearSpellBuffer) = GetAddress(AddressOf HandleClearSpellBuffer)
    HandleDataSub(SSayMsg) = GetAddress(AddressOf HandleSayMsg)
    HandleDataSub(SOpenShop) = GetAddress(AddressOf HandleOpenShop)
    HandleDataSub(SResetShopAction) = GetAddress(AddressOf HandleResetShopAction)
    HandleDataSub(SStunned) = GetAddress(AddressOf HandleStunned)
    HandleDataSub(SMapWornEq) = GetAddress(AddressOf HandleMapWornEq)
    HandleDataSub(SBank) = GetAddress(AddressOf HandleBank)
    HandleDataSub(STrade) = GetAddress(AddressOf HandleTrade)
    HandleDataSub(SCloseTrade) = GetAddress(AddressOf HandleCloseTrade)
    HandleDataSub(STradeUpdate) = GetAddress(AddressOf HandleTradeUpdate)
    HandleDataSub(STradeStatus) = GetAddress(AddressOf HandleTradeStatus)
    HandleDataSub(STarget) = GetAddress(AddressOf HandleTarget)
    HandleDataSub(SHotbar) = GetAddress(AddressOf HandleHotbar)
    HandleDataSub(SHighIndex) = GetAddress(AddressOf HandleHighIndex)
    HandleDataSub(SSound) = GetAddress(AddressOf HandleSound)
    HandleDataSub(STradeRequest) = GetAddress(AddressOf HandleTradeRequest)
    HandleDataSub(SPartyInvite) = GetAddress(AddressOf HandlePartyInvite)
    HandleDataSub(SPartyUpdate) = GetAddress(AddressOf HandlePartyUpdate)
    HandleDataSub(SPartyVitals) = GetAddress(AddressOf HandlePartyVitals)
    HandleDataSub(SChatUpdate) = GetAddress(AddressOf HandleChatUpdate)
    HandleDataSub(SConvEditor) = GetAddress(AddressOf HandleConvEditor)
    HandleDataSub(SUpdateConv) = GetAddress(AddressOf HandleUpdateConv)
    HandleDataSub(SStartTutorial) = GetAddress(AddressOf HandleStartTutorial)
    HandleDataSub(SChatBubble) = GetAddress(AddressOf HandleChatBubble)
    HandleDataSub(SSetPlayerLoginToken) = GetAddress(AddressOf HandleSetPlayerLoginToken)
    HandleDataSub(SPlayerChars) = GetAddress(AddressOf HandlePlayerChars)
    HandleDataSub(SCancelAnimation) = GetAddress(AddressOf HandleCancelAnimation)
    HandleDataSub(SPlayerVariables) = GetAddress(AddressOf HandlePlayerVariables)
    HandleDataSub(SEvent) = GetAddress(AddressOf HandleEvent)
End Sub

Sub HandleData(ByRef data() As Byte)
    Dim Buffer As clsBuffer
    Dim MsgType As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    MsgType = Buffer.ReadLong

    If MsgType < 0 Then
        DestroyGame
        Exit Sub
    End If

    If MsgType >= SMsgCOUNT Then
        DestroyGame
        Exit Sub
    End If

    CallWindowProc HandleDataSub(MsgType), 1, Buffer.ReadBytes(Buffer.length), 0, 0
End Sub

Sub HandleAlertMsg(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer, dialogue_index As Long, menuReset As Long, kick As Long
    
    SetStatus vbNullString
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes data()
    dialogue_index = Buffer.ReadLong
    menuReset = Buffer.ReadLong
    kick = Buffer.ReadLong
    
    Set Buffer = Nothing
    
    If menuReset > 0 Then
        HideWindows
        Select Case menuReset
            Case MenuCount.menuLogin
                ShowWindow GetWindowIndex("winLogin")
            Case MenuCount.menuChars
                ShowWindow GetWindowIndex("winCharacters")
            Case MenuCount.menuClass
                ShowWindow GetWindowIndex("winClasses")
            Case MenuCount.menuNewChar
                ShowWindow GetWindowIndex("winNewChar")
            Case MenuCount.menuMain
                ShowWindow GetWindowIndex("winLogin")
        End Select
    Else
        If kick > 0 Or inMenu = True Then
            ShowWindow GetWindowIndex("winLogin")
            DialogueAlert dialogue_index
            logoutGame
            Exit Sub
        End If
    End If
    
    DialogueAlert dialogue_index
End Sub

Sub HandleLoginOk(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    ' Now we can receive game data
    MyIndex = Buffer.ReadLong
    ' player high index
    Player_HighIndex = MAX_PLAYERS 'Buffer.ReadLong
    Set Buffer = Nothing
    Call SetStatus("Receiving game data.")
End Sub

Sub HandleNewCharClasses(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim i As Long
    Dim z As Long, x As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    n = 1
    ' Max classes
    Max_Classes = Buffer.ReadLong
    ReDim Class(1 To Max_Classes)
    n = n + 1

    For i = 1 To Max_Classes

        With Class(i)
            .name = Buffer.ReadString
            .Vital(Vitals.HP) = Buffer.ReadLong
            .Vital(Vitals.MP) = Buffer.ReadLong
            ' get array size
            z = Buffer.ReadLong
            ' redim array
            ReDim .MaleSprite(0 To z)

            ' loop-receive data
            For x = 0 To z
                .MaleSprite(x) = Buffer.ReadLong
            Next

            ' get array size
            z = Buffer.ReadLong
            ' redim array
            ReDim .FemaleSprite(0 To z)

            ' loop-receive data
            For x = 0 To z
                .FemaleSprite(x) = Buffer.ReadLong
            Next

            For x = 1 To Stats.Stat_Count - 1
                .Stat(x) = Buffer.ReadLong
            Next

        End With

        n = n + 10
    Next

    Set Buffer = Nothing
End Sub

Sub HandleClassesData(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim i As Long
    Dim z As Long, x As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    n = 1
    ' Max classes
    Max_Classes = Buffer.ReadLong 'CByte(Parse(n))
    ReDim Class(1 To Max_Classes)
    n = n + 1

    For i = 1 To Max_Classes

        With Class(i)
            .name = Buffer.ReadString 'Trim$(Parse(n))
            .Vital(Vitals.HP) = Buffer.ReadLong 'CLng(Parse(n + 1))
            .Vital(Vitals.MP) = Buffer.ReadLong 'CLng(Parse(n + 2))
            ' get array size
            z = Buffer.ReadLong
            ' redim array
            ReDim .MaleSprite(0 To z)

            ' loop-receive data
            For x = 0 To z
                .MaleSprite(x) = Buffer.ReadLong
            Next

            ' get array size
            z = Buffer.ReadLong
            ' redim array
            ReDim .FemaleSprite(0 To z)

            ' loop-receive data
            For x = 0 To z
                .FemaleSprite(x) = Buffer.ReadLong
            Next

            For x = 1 To Stats.Stat_Count - 1
                .Stat(x) = Buffer.ReadLong
            Next

        End With

        n = n + 10
    Next

    Set Buffer = Nothing
End Sub

Sub HandleInGame(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    InGame = True
    inMenu = False
    SetStatus vbNullString
    ' show gui
    ShowWindow GetWindowIndex("winBars"), , False
    ShowWindow GetWindowIndex("winMenu"), , False
    ShowWindow GetWindowIndex("winHotbar"), , False
    ShowWindow GetWindowIndex("winChatSmall"), , False
    ' enter loop
    GameLoop
End Sub

Sub HandlePlayerInv(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()

    For i = 1 To MAX_INV
        Call SetPlayerInvItemNum(MyIndex, i, Buffer.ReadLong)
        Call SetPlayerInvItemValue(MyIndex, i, Buffer.ReadLong)
        PlayerInv(i).bound = Buffer.ReadByte
    Next
    
    SetGoldLabel

    Set Buffer = Nothing
End Sub

Sub HandlePlayerInvUpdate(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    n = Buffer.ReadLong 'CLng(Parse(1))
    Call SetPlayerInvItemNum(MyIndex, n, Buffer.ReadLong) 'CLng(Parse(2)))
    Call SetPlayerInvItemValue(MyIndex, n, Buffer.ReadLong) 'CLng(Parse(3)))
    PlayerInv(n).bound = Buffer.ReadByte
    Set Buffer = Nothing
    SetGoldLabel
End Sub

Sub HandlePlayerWornEq(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    Call SetPlayerEquipment(MyIndex, Buffer.ReadLong, Armor)
    Call SetPlayerEquipment(MyIndex, Buffer.ReadLong, Weapon)
    Call SetPlayerEquipment(MyIndex, Buffer.ReadLong, Helmet)
    Call SetPlayerEquipment(MyIndex, Buffer.ReadLong, Shield)
    Set Buffer = Nothing
End Sub

Sub HandleMapWornEq(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim playerNum As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    playerNum = Buffer.ReadLong
    Call SetPlayerEquipment(playerNum, Buffer.ReadLong, Armor)
    Call SetPlayerEquipment(playerNum, Buffer.ReadLong, Weapon)
    Call SetPlayerEquipment(playerNum, Buffer.ReadLong, Helmet)
    Call SetPlayerEquipment(playerNum, Buffer.ReadLong, Shield)
    Set Buffer = Nothing
End Sub

Private Sub HandlePlayerHp(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    If MyIndex = 0 Then Exit Sub
    Buffer.WriteBytes data()
    Player(MyIndex).MaxVital(Vitals.HP) = Buffer.ReadLong
    Call SetPlayerVital(MyIndex, Vitals.HP, Buffer.ReadLong)
    ' set max width
    If GetPlayerVital(MyIndex, Vitals.HP) > 0 Then
        BarWidth_GuiHP_Max = ((GetPlayerVital(MyIndex, Vitals.HP) / 209) / (GetPlayerMaxVital(MyIndex, Vitals.HP) / 209)) * 209
    Else
        BarWidth_GuiHP_Max = 0
    End If
    ' Update GUI
    UpdateStats_UI
End Sub

Private Sub HandlePlayerMp(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    Player(MyIndex).MaxVital(Vitals.MP) = Buffer.ReadLong
    Call SetPlayerVital(MyIndex, Vitals.MP, Buffer.ReadLong)
    ' set max width
    If GetPlayerVital(MyIndex, Vitals.MP) > 0 Then
        BarWidth_GuiSP_Max = ((GetPlayerVital(MyIndex, Vitals.MP) / 209) / (GetPlayerMaxVital(MyIndex, Vitals.MP) / 209)) * 209
    Else
        BarWidth_GuiSP_Max = 0
    End If
    ' Update GUI
    UpdateStats_UI
End Sub

Private Sub HandlePlayerStats(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim i As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()

    For i = 1 To Stats.Stat_Count - 1
        SetPlayerStat index, i, Buffer.ReadLong
    Next
End Sub

Private Sub HandlePlayerExp(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    SetPlayerExp MyIndex, Buffer.ReadLong
    TNL = Buffer.ReadLong
    ' set max width
    If GetPlayerLevel(MyIndex) <= MAX_LEVELS Then
        If GetPlayerExp(MyIndex) > 0 Then
            BarWidth_GuiEXP_Max = ((GetPlayerExp(MyIndex) / 209) / (TNL / 209)) * 209
        Else
            BarWidth_GuiEXP_Max = 0
        End If
    Else
        BarWidth_GuiEXP_Max = 209
    End If
    ' Update GUI
    UpdateStats_UI
End Sub

Private Sub HandlePlayerData(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long, x As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    i = Buffer.ReadLong
    Call SetPlayerName(i, Buffer.ReadString)
    Player(i).usergroup = Buffer.ReadLong
    Call SetPlayerLevel(i, Buffer.ReadLong)
    Call SetPlayerPOINTS(i, Buffer.ReadLong)
    Call SetPlayerSprite(i, Buffer.ReadLong)
    Call SetPlayerMap(i, Buffer.ReadLong)
    Call SetPlayerX(i, Buffer.ReadLong)
    Call SetPlayerY(i, Buffer.ReadLong)
    Call SetPlayerDir(i, Buffer.ReadLong)
    Call SetPlayerAccess(i, Buffer.ReadLong)
    Call SetPlayerPK(i, Buffer.ReadLong)
    Call SetPlayerClass(i, Buffer.ReadLong)

    For x = 1 To Stats.Stat_Count - 1
        SetPlayerStat i, x, Buffer.ReadLong
    Next

    ' Check if the player is the client player
    If i = MyIndex Then
        ' Reset directions
        wDown = False
        aDown = False
        sDown = False
        dDown = False
        upDown = False
        leftDown = False
        downDown = False
        rightDown = False
        ' set form
        With Windows(GetWindowIndex("winCharacter"))
            .Controls(GetControlIndex("winCharacter", "lblName")).text = "Name: " & Trim$(GetPlayerName(MyIndex))
            .Controls(GetControlIndex("winCharacter", "lblClass")).text = "Class: " & Trim$(Class(GetPlayerClass(MyIndex)).name)
            .Controls(GetControlIndex("winCharacter", "lblLevel")).text = "Level: " & GetPlayerLevel(MyIndex)
            .Controls(GetControlIndex("winCharacter", "lblGuild")).text = "Guild: " & "None"
            .Controls(GetControlIndex("winCharacter", "lblHealth")).text = "Health: " & GetPlayerVital(MyIndex, HP) & "/" & GetPlayerMaxVital(MyIndex, HP)
            .Controls(GetControlIndex("winCharacter", "lblSpirit")).text = "Spirit: " & GetPlayerVital(MyIndex, MP) & "/" & GetPlayerMaxVital(MyIndex, MP)
            .Controls(GetControlIndex("winCharacter", "lblExperience")).text = "Experience: " & Player(MyIndex).EXP & "/" & TNL
            ' stats
            For x = 1 To Stats.Stat_Count - 1
                .Controls(GetControlIndex("winCharacter", "lblStat_" & x)).text = GetPlayerStat(MyIndex, x)
            Next
            ' points
            .Controls(GetControlIndex("winCharacter", "lblPoints")).text = GetPlayerPOINTS(MyIndex)
            ' grey out buttons
            If GetPlayerPOINTS(MyIndex) = 0 Then
                For x = 1 To Stats.Stat_Count - 1
                    .Controls(GetControlIndex("winCharacter", "btnGreyStat_" & x)).visible = True
                Next
            Else
                For x = 1 To Stats.Stat_Count - 1
                    .Controls(GetControlIndex("winCharacter", "btnGreyStat_" & x)).visible = False
                Next
            End If
        End With
    End If

    ' Make sure they aren't walking
    Player(i).Moving = 0
    Player(i).xOffset = 0
    Player(i).yOffset = 0
End Sub

Private Sub HandlePlayerMove(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim x As Long
    Dim y As Long
    Dim dir As Long
    Dim n As Byte
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    i = Buffer.ReadLong
    x = Buffer.ReadLong
    y = Buffer.ReadLong
    dir = Buffer.ReadLong
    n = Buffer.ReadLong
    Call SetPlayerX(i, x)
    Call SetPlayerY(i, y)
    Call SetPlayerDir(i, dir)
    Player(i).xOffset = 0
    Player(i).yOffset = 0
    Player(i).Moving = n

    Select Case GetPlayerDir(i)

        Case DIR_UP
            Player(i).yOffset = PIC_Y

        Case DIR_DOWN
            Player(i).yOffset = PIC_Y * -1

        Case DIR_LEFT
            Player(i).xOffset = PIC_X

        Case DIR_RIGHT
            Player(i).xOffset = PIC_X * -1
    End Select
End Sub

Private Sub HandleNpcMove(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim MapNpcNum As Long
    Dim x As Long
    Dim y As Long
    Dim dir As Long
    Dim Movement As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    MapNpcNum = Buffer.ReadLong
    x = Buffer.ReadLong
    y = Buffer.ReadLong
    dir = Buffer.ReadLong
    Movement = Buffer.ReadLong

    With MapNpc(MapNpcNum)
        .x = x
        .y = y
        .dir = dir
        .xOffset = 0
        .yOffset = 0
        .Moving = Movement

        Select Case .dir

            Case DIR_UP
                .yOffset = PIC_Y

            Case DIR_DOWN
                .yOffset = PIC_Y * -1

            Case DIR_LEFT
                .xOffset = PIC_X

            Case DIR_RIGHT
                .xOffset = PIC_X * -1
        End Select

    End With

End Sub

Private Sub HandlePlayerDir(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim dir As Byte
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    i = Buffer.ReadLong
    dir = Buffer.ReadLong
    Call SetPlayerDir(i, dir)

    With Player(i)
        .xOffset = 0
        .yOffset = 0
        .Moving = 0
    End With

End Sub

Private Sub HandleNpcDir(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim dir As Byte
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    i = Buffer.ReadLong
    dir = Buffer.ReadLong

    With MapNpc(i)
        .dir = dir
        .xOffset = 0
        .yOffset = 0
        .Moving = 0
    End With

End Sub

Private Sub HandlePlayerXY(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim x As Long
    Dim y As Long
    Dim dir As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    x = Buffer.ReadLong
    y = Buffer.ReadLong
    dir = Buffer.ReadLong
    Call SetPlayerX(MyIndex, x)
    Call SetPlayerY(MyIndex, y)
    Call SetPlayerDir(MyIndex, dir)
    ' Make sure they aren't walking
    Player(MyIndex).Moving = 0
    Player(MyIndex).xOffset = 0
    Player(MyIndex).yOffset = 0
End Sub

Private Sub HandlePlayerXYMap(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim x As Long
    Dim y As Long
    Dim dir As Long
    Dim Buffer As clsBuffer
    Dim thePlayer As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    thePlayer = Buffer.ReadLong
    x = Buffer.ReadLong
    y = Buffer.ReadLong
    dir = Buffer.ReadLong
    Call SetPlayerX(thePlayer, x)
    Call SetPlayerY(thePlayer, y)
    Call SetPlayerDir(thePlayer, dir)
    ' Make sure they aren't walking
    Player(thePlayer).Moving = 0
    Player(thePlayer).xOffset = 0
    Player(thePlayer).yOffset = 0
End Sub

Private Sub HandleAttack(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    i = Buffer.ReadLong
    ' Set player to attacking
    Player(i).Attacking = 1
    Player(i).AttackTimer = GetTickCount
End Sub

Private Sub HandleNpcAttack(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    i = Buffer.ReadLong
    ' Set player to attacking
    MapNpc(i).Attacking = 1
    MapNpc(i).AttackTimer = GetTickCount
End Sub

Private Sub HandleCheckForMap(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long, NeedMap As Byte, Buffer As clsBuffer, MapDataCRC As Long, MapTileCRC As Long, mapNum As Long
    
    GettingMap = True
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()

    ' Erase all players except self
    For i = 1 To MAX_PLAYERS
        If i <> MyIndex Then
            Call SetPlayerMap(i, 0)
        End If
    Next

    ' Erase all temporary tile values
    Call ClearTempTile
    Call ClearMapNpcs
    Call ClearMapItems
    Call ClearMap

    ' clear the blood
    For i = 1 To MAX_BYTE
        Blood(i).x = 0
        Blood(i).y = 0
        Blood(i).sprite = 0
        Blood(i).timer = 0
    Next

    ' Get map num
    mapNum = Buffer.ReadLong
    MapDataCRC = Buffer.ReadLong
    MapTileCRC = Buffer.ReadLong
    
    ' check against our own CRC32s
    NeedMap = 0
    If MapDataCRC <> MapCRC32(mapNum).MapDataCRC Then
        NeedMap = 1
    End If
    If MapTileCRC <> MapCRC32(mapNum).MapTileCRC Then
        NeedMap = 1
    End If

    ' Either the revisions didn't match or we dont have the map, so we need it
    Set Buffer = New clsBuffer
    Buffer.WriteLong CNeedMap
    Buffer.WriteLong NeedMap
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Check if we get a map from someone else and if we were editing a map cancel it out
    If Not applyingMap Then
        If InMapEditor Then
            InMapEditor = False
            frmEditor_Map.visible = False
            ClearAttributeDialogue
    
            If frmEditor_MapProperties.visible Then
                frmEditor_MapProperties.visible = False
            End If
        End If
    End If
    
    ' load the map if we don't need it
    If NeedMap = 0 Then
        LoadMap mapNum
        applyingMap = False
    End If
End Sub

Sub HandleMapData(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer, mapNum As Long, i As Long, x As Long, y As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    
    mapNum = Buffer.ReadLong
    
    With map.MapData
        .name = Buffer.ReadString
        .Music = Buffer.ReadString
        .Moral = Buffer.ReadByte
        .Up = Buffer.ReadLong
        .Down = Buffer.ReadLong
        .left = Buffer.ReadLong
        .Right = Buffer.ReadLong
        .BootMap = Buffer.ReadLong
        .BootX = Buffer.ReadByte
        .BootY = Buffer.ReadByte
        .MaxX = Buffer.ReadByte
        .MaxY = Buffer.ReadByte
        .BossNpc = Buffer.ReadLong
        For i = 1 To MAX_MAP_NPCS
            .Npc(i) = Buffer.ReadLong
        Next
    End With
    
    map.TileData.EventCount = Buffer.ReadLong
    If map.TileData.EventCount > 0 Then
        ReDim Preserve map.TileData.Events(1 To map.TileData.EventCount)
        For i = 1 To map.TileData.EventCount
            With map.TileData.Events(i)
                .name = Buffer.ReadString
                .x = Buffer.ReadLong
                .y = Buffer.ReadLong
                .pageCount = Buffer.ReadLong
            End With
            If map.TileData.Events(i).pageCount > 0 Then
                ReDim Preserve map.TileData.Events(i).EventPage(1 To map.TileData.Events(i).pageCount)
                For x = 1 To map.TileData.Events(i).pageCount
                    With map.TileData.Events(i).EventPage(x)
                        .chkPlayerVar = Buffer.ReadByte
                        .chkSelfSwitch = Buffer.ReadByte
                        .chkHasItem = Buffer.ReadByte
                        .PlayerVarNum = Buffer.ReadLong
                        .SelfSwitchNum = Buffer.ReadLong
                        .HasItemNum = Buffer.ReadLong
                        .PlayerVariable = Buffer.ReadLong
                        .GraphicType = Buffer.ReadByte
                        .Graphic = Buffer.ReadLong
                        .GraphicX = Buffer.ReadLong
                        .GraphicY = Buffer.ReadLong
                        .MoveType = Buffer.ReadByte
                        .MoveSpeed = Buffer.ReadByte
                        .MoveFreq = Buffer.ReadByte
                        .WalkAnim = Buffer.ReadByte
                        .StepAnim = Buffer.ReadByte
                        .DirFix = Buffer.ReadByte
                        .WalkThrough = Buffer.ReadByte
                        .Priority = Buffer.ReadByte
                        .Trigger = Buffer.ReadByte
                        .CommandCount = Buffer.ReadLong
                    End With
                    If map.TileData.Events(i).EventPage(x).CommandCount > 0 Then
                        ReDim Preserve map.TileData.Events(i).EventPage(x).Commands(1 To map.TileData.Events(i).EventPage(x).CommandCount)
                        For y = 1 To map.TileData.Events(i).EventPage(x).CommandCount
                            With map.TileData.Events(i).EventPage(x).Commands(y)
                                .Type = Buffer.ReadByte
                                .text = Buffer.ReadString
                                .Colour = Buffer.ReadLong
                                .channel = Buffer.ReadByte
                                .TargetType = Buffer.ReadByte
                                .target = Buffer.ReadLong
                                .x = Buffer.ReadLong
                                .y = Buffer.ReadLong
                            End With
                        Next
                    End If
                Next
            End If
        Next
    End If
    
    ReDim map.TileData.Tile(0 To map.MapData.MaxX, 0 To map.MapData.MaxY)

    For x = 0 To map.MapData.MaxX
        For y = 0 To map.MapData.MaxY
            For i = 1 To MapLayer.Layer_Count - 1
                map.TileData.Tile(x, y).Layer(i).x = Buffer.ReadLong
                map.TileData.Tile(x, y).Layer(i).y = Buffer.ReadLong
                map.TileData.Tile(x, y).Layer(i).tileset = Buffer.ReadLong
                map.TileData.Tile(x, y).Autotile(i) = Buffer.ReadByte
            Next
            map.TileData.Tile(x, y).Type = Buffer.ReadByte
            map.TileData.Tile(x, y).Data1 = Buffer.ReadLong
            map.TileData.Tile(x, y).Data2 = Buffer.ReadLong
            map.TileData.Tile(x, y).Data3 = Buffer.ReadLong
            map.TileData.Tile(x, y).Data4 = Buffer.ReadLong
            map.TileData.Tile(x, y).Data5 = Buffer.ReadLong
            map.TileData.Tile(x, y).DirBlock = Buffer.ReadByte
        Next
    Next

    ClearTempTile
    initAutotiles
    Set Buffer = Nothing
    ' Save the map
    Call SaveMap(mapNum)
    GetMapCRC32 mapNum
    AddText "Downloaded new map.", BrightGreen

    ' Check if we get a map from someone else and if we were editing a map cancel it out
    If Not applyingMap Then
        If InMapEditor Then
            InMapEditor = False
            frmEditor_Map.visible = False
            ClearAttributeDialogue
            If frmEditor_MapProperties.visible Then
                frmEditor_MapProperties.visible = False
            End If
        End If
    End If
    applyingMap = False

End Sub

Private Sub HandleMapItemData(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim Buffer As clsBuffer, tmpLong As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()

    For i = 1 To MAX_MAP_ITEMS

        With MapItem(i)
            .playerName = Buffer.ReadString
            .num = Buffer.ReadLong
            .value = Buffer.ReadLong
            .x = Buffer.ReadLong
            .y = Buffer.ReadLong
            tmpLong = Buffer.ReadLong

            If tmpLong = 0 Then
                .bound = False
            Else
                .bound = True
            End If

        End With

    Next

End Sub

Private Sub HandleMapNpcData(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()

    For i = 1 To MAX_MAP_NPCS

        With MapNpc(i)
            .num = Buffer.ReadLong
            .x = Buffer.ReadLong
            .y = Buffer.ReadLong
            .dir = Buffer.ReadLong
            .Vital(HP) = Buffer.ReadLong
        End With

    Next

End Sub

Private Sub HandleMapDone()
    Dim i As Long
    Dim musicFile As String

    ' clear the action msgs
    For i = 1 To MAX_BYTE
        ClearActionMsg (i)
    Next i

    Action_HighIndex = 1

    ' player music
    If InGame Then
        musicFile = Trim$(map.MapData.Music)

        If Not musicFile = "None." Then
            Play_Music musicFile
        Else
            Stop_Music
        End If
    End If

    ' get the npc high index
    For i = MAX_MAP_NPCS To 1 Step -1

        If MapNpc(i).num > 0 Then
            Npc_HighIndex = i + 1
            Exit For
        End If

    Next

    ' make sure we're not overflowing
    If Npc_HighIndex > MAX_MAP_NPCS Then Npc_HighIndex = MAX_MAP_NPCS
    ' now cache the positions
    initAutotiles
    GettingMap = False
    CanMoveNow = True
End Sub

Private Sub HandleBroadcastMsg(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Msg As String
    Dim Color As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    Msg = Buffer.ReadString
    Color = Buffer.ReadLong
    Call AddText(Msg, Color)
End Sub

Private Sub HandleGlobalMsg(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Msg As String
    Dim Color As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    Msg = Buffer.ReadString
    Color = Buffer.ReadLong
    Call AddText(Msg, Color)
End Sub

Private Sub HandlePlayerMsg(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Msg As String
    Dim Color As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    Msg = Buffer.ReadString
    Color = Buffer.ReadLong
    Call AddText(Msg, Color)
End Sub

Private Sub HandleMapMsg(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Msg As String
    Dim Color As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    Msg = Buffer.ReadString
    Color = Buffer.ReadLong
    Call AddText(Msg, Color)
End Sub

Private Sub HandleAdminMsg(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Msg As String
    Dim Color As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    Msg = Buffer.ReadString
    Color = Buffer.ReadLong
    Call AddText(Msg, Color)
End Sub

Private Sub HandleSpawnItem(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer, tmpLong As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    n = Buffer.ReadLong

    With MapItem(n)
        .playerName = Buffer.ReadString
        .num = Buffer.ReadLong
        .value = Buffer.ReadLong
        .x = Buffer.ReadLong
        .y = Buffer.ReadLong
        tmpLong = Buffer.ReadLong

        If tmpLong = 0 Then
            .bound = False
        Else
            .bound = True
        End If

    End With

End Sub

Private Sub HandleItemEditor()
    Dim i As Long

    With frmEditor_Item
        Editor = EDITOR_ITEM
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_ITEMS
            .lstIndex.AddItem i & ": " & Trim$(Item(i).name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        ItemEditorInit
    End With

End Sub

Private Sub HandleAnimationEditor()
    Dim i As Long

    With frmEditor_Animation
        Editor = EDITOR_ANIMATION
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_ANIMATIONS
            .lstIndex.AddItem i & ": " & Trim$(Animation(i).name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        AnimationEditorInit
    End With

End Sub

Private Sub HandleUpdateItem(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Dim ItemSize As Long
    Dim ItemData() As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    n = Buffer.ReadLong
    ' Update the item
    ItemSize = LenB(Item(n))
    ReDim ItemData(ItemSize - 1)
    ItemData = Buffer.ReadBytes(ItemSize)
    CopyMemory ByVal VarPtr(Item(n)), ByVal VarPtr(ItemData(0)), ItemSize
    Set Buffer = Nothing
End Sub

Private Sub HandleUpdateAnimation(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Dim AnimationSize As Long
    Dim AnimationData() As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    n = Buffer.ReadLong
    ' Update the Animation
    AnimationSize = LenB(Animation(n))
    ReDim AnimationData(AnimationSize - 1)
    AnimationData = Buffer.ReadBytes(AnimationSize)
    CopyMemory ByVal VarPtr(Animation(n)), ByVal VarPtr(AnimationData(0)), AnimationSize
    Set Buffer = Nothing
End Sub

Private Sub HandleSpawnNpc(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    n = Buffer.ReadLong

    With MapNpc(n)
        .num = Buffer.ReadLong
        .x = Buffer.ReadLong
        .y = Buffer.ReadLong
        .dir = Buffer.ReadLong
        ' Client use only
        .xOffset = 0
        .yOffset = 0
        .Moving = 0
    End With

End Sub

Private Sub HandleNpcDead(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    n = Buffer.ReadLong
    Call ClearMapNpc(n)
End Sub

Private Sub HandleNpcEditor()
    Dim i As Long

    With frmEditor_NPC
        Editor = EDITOR_NPC
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_NPCS
            .lstIndex.AddItem i & ": " & Trim$(Npc(i).name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        NpcEditorInit
    End With

End Sub

Private Sub HandleUpdateNpc(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Dim NpcSize As Long
    Dim NpcData() As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    n = Buffer.ReadLong
    NpcSize = LenB(Npc(n))
    ReDim NpcData(NpcSize - 1)
    NpcData = Buffer.ReadBytes(NpcSize)
    CopyMemory ByVal VarPtr(Npc(n)), ByVal VarPtr(NpcData(0)), NpcSize
    Set Buffer = Nothing
End Sub

Private Sub HandleResourceEditor()
    Dim i As Long

    With frmEditor_Resource
        Editor = EDITOR_RESOURCE
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_RESOURCES
            .lstIndex.AddItem i & ": " & Trim$(Resource(i).name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        ResourceEditorInit
    End With

End Sub

Private Sub HandleUpdateResource(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim ResourceNum As Long
    Dim Buffer As clsBuffer
    Dim ResourceSize As Long
    Dim ResourceData() As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    ResourceNum = Buffer.ReadLong
    ResourceSize = LenB(Resource(ResourceNum))
    ReDim ResourceData(ResourceSize - 1)
    ResourceData = Buffer.ReadBytes(ResourceSize)
    ClearResource ResourceNum
    CopyMemory ByVal VarPtr(Resource(ResourceNum)), ByVal VarPtr(ResourceData(0)), ResourceSize
    Set Buffer = Nothing
End Sub

Private Sub HandleMapKey(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim x As Long
    Dim y As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    x = Buffer.ReadLong
    y = Buffer.ReadLong
    n = Buffer.ReadByte
    TempTile(x, y).DoorOpen = n

    ' re-cache rendering
    If Not GettingMap Then cacheRenderState x, y, MapLayer.Mask
End Sub

Private Sub HandleEditMap()
    Call MapEditorInit
End Sub

Private Sub HandleShopEditor()
    Dim i As Long

    With frmEditor_Shop
        Editor = EDITOR_SHOP
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_SHOPS
            .lstIndex.AddItem i & ": " & Trim$(Shop(i).name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        ShopEditorInit
    End With

End Sub

Private Sub HandleUpdateShop(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim shopNum As Long
    Dim Buffer As clsBuffer
    Dim ShopSize As Long
    Dim ShopData() As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    shopNum = Buffer.ReadLong
    ShopSize = LenB(Shop(shopNum))
    ReDim ShopData(ShopSize - 1)
    ShopData = Buffer.ReadBytes(ShopSize)
    CopyMemory ByVal VarPtr(Shop(shopNum)), ByVal VarPtr(ShopData(0)), ShopSize
    Set Buffer = Nothing
End Sub

Private Sub HandleSpellEditor()
    Dim i As Long

    With frmEditor_Spell
        Editor = EDITOR_SPELL
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_SPELLS
            .lstIndex.AddItem i & ": " & Trim$(Spell(i).name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        SpellEditorInit
    End With

End Sub

Private Sub HandleUpdateSpell(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim spellnum As Long
    Dim Buffer As clsBuffer
    Dim SpellSize As Long
    Dim SpellData() As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    spellnum = Buffer.ReadLong
    SpellSize = LenB(Spell(spellnum))
    ReDim SpellData(SpellSize - 1)
    SpellData = Buffer.ReadBytes(SpellSize)
    CopyMemory ByVal VarPtr(Spell(spellnum)), ByVal VarPtr(SpellData(0)), SpellSize
    Set Buffer = Nothing
End Sub

Sub HandleSpells(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()

    For i = 1 To MAX_PLAYER_SPELLS
        PlayerSpells(i).Spell = Buffer.ReadLong
        PlayerSpells(i).Uses = Buffer.ReadLong
    Next

    Set Buffer = Nothing
End Sub

Private Sub HandleLeft(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    Call ClearPlayer(Buffer.ReadLong)
    Set Buffer = Nothing
End Sub

Private Sub HandleResourceCache(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim i As Long

    ' if in map editor, we cache shit ourselves
    If InMapEditor Then Exit Sub
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    Resource_Index = Buffer.ReadLong
    Resources_Init = False

    If Resource_Index > 0 Then
        ReDim Preserve MapResource(0 To Resource_Index)

        For i = 0 To Resource_Index
            MapResource(i).ResourceState = Buffer.ReadByte
            MapResource(i).x = Buffer.ReadLong
            MapResource(i).y = Buffer.ReadLong
        Next

        Resources_Init = True
    Else
        ReDim MapResource(0 To 1)
    End If

    Set Buffer = Nothing
End Sub

Private Sub HandleSendPing(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    PingEnd = GetTickCount
    Ping = PingEnd - PingStart
End Sub

Private Sub HandleDoorAnimation(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim x As Long, y As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    x = Buffer.ReadLong
    y = Buffer.ReadLong

    With TempTile(x, y)
        .DoorFrame = 1
        .DoorAnimate = 1 ' 0 = nothing| 1 = opening | 2 = closing
        .DoorTimer = GetTickCount
    End With

    Set Buffer = Nothing
End Sub

Private Sub HandleActionMsg(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim x As Long, y As Long, message As String, Color As Long, tmpType As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    message = Buffer.ReadString
    Color = Buffer.ReadLong
    tmpType = Buffer.ReadLong
    x = Buffer.ReadLong
    y = Buffer.ReadLong
    Set Buffer = Nothing
    CreateActionMsg message, Color, tmpType, x, y
End Sub

Private Sub HandleBlood(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim x As Long, y As Long, sprite As Long, i As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    x = Buffer.ReadLong
    y = Buffer.ReadLong
    Set Buffer = Nothing
    ' randomise sprite
    sprite = Rand(1, BloodCount)

    ' make sure tile doesn't already have blood
    For i = 1 To MAX_BYTE

        If Blood(i).x = x And Blood(i).y = y Then
            ' already have blood :(
            Exit Sub
        End If

    Next

    ' carry on with the set
    BloodIndex = BloodIndex + 1

    If BloodIndex >= MAX_BYTE Then BloodIndex = 1

    With Blood(BloodIndex)
        .x = x
        .y = y
        .sprite = sprite
        .timer = GetTickCount
    End With

End Sub

Private Sub HandleAnimation(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer, x As Long, y As Long, isCasting As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    AnimationIndex = AnimationIndex + 1

    If AnimationIndex >= MAX_BYTE Then AnimationIndex = 1

    With AnimInstance(AnimationIndex)
        .Animation = Buffer.ReadLong
        .x = Buffer.ReadLong
        .y = Buffer.ReadLong
        .LockType = Buffer.ReadByte
        .lockindex = Buffer.ReadLong
        .isCasting = Buffer.ReadByte
        .Used(0) = True
        .Used(1) = True
    End With

    Set Buffer = Nothing

    ' play the sound if we've got one
    With AnimInstance(AnimationIndex)

        If .LockType = 0 Then
            x = AnimInstance(AnimationIndex).x
            y = AnimInstance(AnimationIndex).y
        ElseIf .LockType = TARGET_TYPE_PLAYER Then
            x = GetPlayerX(.lockindex)
            y = GetPlayerY(.lockindex)
        ElseIf .LockType = TARGET_TYPE_NPC Then
            x = MapNpc(.lockindex).x
            y = MapNpc(.lockindex).y
        End If

    End With

    PlayMapSound x, y, SoundEntity.seAnimation, AnimInstance(AnimationIndex).Animation
End Sub

Private Sub HandleMapNpcVitals(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim i As Long
    Dim MapNpcNum As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    MapNpcNum = Buffer.ReadLong

    For i = 1 To Vitals.Vital_Count - 1
        MapNpc(MapNpcNum).Vital(i) = Buffer.ReadLong
    Next

    Set Buffer = Nothing
End Sub

Private Sub HandleCooldown(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Slot As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    Slot = Buffer.ReadLong
    SpellCD(Slot) = GetTickCount
    Set Buffer = Nothing
End Sub

Private Sub HandleClearSpellBuffer(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SpellBuffer = 0
    SpellBufferTimer = 0
End Sub

Private Sub HandleSayMsg(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer, Access As Long, name As String, message As String, Colour As Long, header As String, PK As Long, saycolour As Long, usergroup As Long
    Dim channel As Byte, colStr As String
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    name = Buffer.ReadString
    Access = Buffer.ReadLong
    usergroup = Buffer.ReadLong
    PK = Buffer.ReadLong
    message = Buffer.ReadString
    header = Buffer.ReadString
    saycolour = Buffer.ReadLong
    Set Buffer = Nothing
    
    ' Check access level
    Colour = White

    If usergroup = 10 Or usergroup = 11 Then Colour = Gold
    If Access > 0 Then Colour = Pink
    If PK > 0 Then Colour = BrightRed
    
    ' find channel
    channel = 0
    Select Case header
        Case "[Map] "
            channel = ChatChannel.chMap
        Case "[Global] "
            channel = ChatChannel.chGlobal
    End Select
    
    ' remove the colour char from the message
    message = Replace$(message, ColourChar, vbNullString)
    ' add to the chat box
    AddText ColourChar & GetColStr(Colour) & header & name & ": " & ColourChar & GetColStr(Grey) & message, Grey, , channel
End Sub

Private Sub HandleOpenShop(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim shopNum As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    shopNum = Buffer.ReadLong
    OpenShop shopNum
    Set Buffer = Nothing
End Sub

Private Sub HandleStunned(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    StunDuration = Buffer.ReadLong
    Set Buffer = Nothing
End Sub

Private Sub HandleBank(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim i As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()

    For i = 1 To MAX_BANK
        Bank.Item(i).num = Buffer.ReadLong
        Bank.Item(i).value = Buffer.ReadLong
    Next

    InBank = True
    Set Buffer = Nothing
End Sub

Private Sub HandleTrade(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    InTrade = Buffer.ReadLong
    Set Buffer = Nothing
    
    ShowTrade
End Sub

Private Sub HandleCloseTrade(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    InTrade = 0
    HideWindow GetWindowIndex("winTrade")
End Sub

Private Sub HandleTradeUpdate(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer, dataType As Byte, i As Long, yourWorth As Long, theirWorth As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    dataType = Buffer.ReadByte

    If dataType = 0 Then ' ours!
        For i = 1 To MAX_INV
            TradeYourOffer(i).num = Buffer.ReadLong
            TradeYourOffer(i).value = Buffer.ReadLong
        Next
        yourWorth = Buffer.ReadLong
        Windows(GetWindowIndex("winTrade")).Controls(GetControlIndex("winTrade", "lblYourValue")).text = yourWorth & "g"
    ElseIf dataType = 1 Then 'theirs
        For i = 1 To MAX_INV
            TradeTheirOffer(i).num = Buffer.ReadLong
            TradeTheirOffer(i).value = Buffer.ReadLong
        Next
        theirWorth = Buffer.ReadLong
        Windows(GetWindowIndex("winTrade")).Controls(GetControlIndex("winTrade", "lblTheirValue")).text = theirWorth & "g"
    End If

    Set Buffer = Nothing
End Sub

Private Sub HandleTradeStatus(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim tradeStatus As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    tradeStatus = Buffer.ReadByte
    Set Buffer = Nothing

    Select Case tradeStatus
        Case 0 ' clear
            Windows(GetWindowIndex("winTrade")).Controls(GetControlIndex("winTrade", "lblStatus")).text = "Choose items to offer."
        Case 1 ' they've accepted
            Windows(GetWindowIndex("winTrade")).Controls(GetControlIndex("winTrade", "lblStatus")).text = "Other player has accepted."
        Case 2 ' you've accepted
            Windows(GetWindowIndex("winTrade")).Controls(GetControlIndex("winTrade", "lblStatus")).text = "Waiting for other player to accept."
        Case 3 ' no room
            Windows(GetWindowIndex("winTrade")).Controls(GetControlIndex("winTrade", "lblStatus")).text = "Not enough inventory space."
    End Select
End Sub

Private Sub HandleTarget(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    myTarget = Buffer.ReadLong
    myTargetType = Buffer.ReadLong
    Set Buffer = Nothing
End Sub

Private Sub HandleHotbar(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim i As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()

    For i = 1 To MAX_HOTBAR
        Hotbar(i).Slot = Buffer.ReadLong
        Hotbar(i).sType = Buffer.ReadByte
    Next
End Sub

Private Sub HandleHighIndex(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    Player_HighIndex = MAX_PLAYERS 'Buffer.ReadLong
End Sub

Private Sub HandleResetShopAction(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    UpdateShop
End Sub

Private Sub HandleSound(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim x As Long, y As Long, entityType As Long, entityNum As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    x = Buffer.ReadLong
    y = Buffer.ReadLong
    entityType = Buffer.ReadLong
    entityNum = Buffer.ReadLong
    PlayMapSound x, y, entityType, entityNum
End Sub

Private Sub HandleTradeRequest(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer, theName As String, top As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    theName = Buffer.ReadString
    ' cache name and show invitation
    diaDataString = theName
    ShowWindow GetWindowIndex("winInvite_Trade")
    Windows(GetWindowIndex("winInvite_Trade")).Controls(GetControlIndex("winInvite_Trade", "btnInvite")).text = ColourChar & White & theName & ColourChar & "-1" & " has invited you to trade."
    AddText Trim$(theName) & " has invited you to trade.", White
    ' loop through
    top = ScreenHeight - 80
    If Windows(GetWindowIndex("winInvite_Party")).Window.visible Then
        top = top - 37
    End If
    Windows(GetWindowIndex("winInvite_Trade")).Window.top = top
End Sub

Private Sub HandlePartyInvite(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer, theName As String, top As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    theName = Buffer.ReadString
    ' cache name and show invitation popup
    diaDataString = theName
    ShowWindow GetWindowIndex("winInvite_Party")
    Windows(GetWindowIndex("winInvite_Party")).Controls(GetControlIndex("winInvite_Party", "btnInvite")).text = ColourChar & White & theName & ColourChar & "-1" & " has invited you to a party."
    AddText Trim$(theName) & " has invited you to a party.", White
    ' loop through
    top = ScreenHeight - 80
    If Windows(GetWindowIndex("winInvite_Trade")).Window.visible Then
        top = top - 37
    End If
    Windows(GetWindowIndex("winInvite_Party")).Window.top = top
End Sub

Private Sub HandlePartyUpdate(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer, i As Long, inParty As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    inParty = Buffer.ReadByte

    ' exit out if we're not in a party
    If inParty = 0 Then
        Call ZeroMemory(ByVal VarPtr(Party), LenB(Party))
        UpdatePartyInterface
        ' exit out early
        Exit Sub
    End If

    ' carry on otherwise
    Party.Leader = Buffer.ReadLong

    For i = 1 To MAX_PARTY_MEMBERS
        Party.Member(i) = Buffer.ReadLong
    Next

    Party.MemberCount = Buffer.ReadLong
    
    ' update the party interface
    UpdatePartyInterface
End Sub

Private Sub HandlePartyVitals(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim playerNum As Long
    Dim Buffer As clsBuffer, i As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    ' which player?
    playerNum = Buffer.ReadLong

    ' set vitals
    For i = 1 To Vitals.Vital_Count - 1
        Player(playerNum).MaxVital(i) = Buffer.ReadLong
        Player(playerNum).Vital(i) = Buffer.ReadLong
    Next

    ' update the party interface
    UpdatePartyBars
End Sub

Private Sub HandleConvEditor(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long

    With frmEditor_Conv
        Editor = EDITOR_CONV
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_CONVS
            .lstIndex.AddItem i & ": " & Trim$(Conv(i).name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        ConvEditorInit
    End With

End Sub

Private Sub HandleUpdateConv(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Convnum As Long
    Dim Buffer As clsBuffer
    Dim i As Long
    Dim x As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    Convnum = Buffer.ReadLong

    With Conv(Convnum)
        .name = Buffer.ReadString
        .chatCount = Buffer.ReadLong
        ReDim Conv(Convnum).Conv(1 To .chatCount)

        For i = 1 To .chatCount
            .Conv(i).Conv = Buffer.ReadString

            For x = 1 To 4
                .Conv(i).rText(x) = Buffer.ReadString
                .Conv(i).rTarget(x) = Buffer.ReadLong
            Next

            .Conv(i).Event = Buffer.ReadLong
            .Conv(i).Data1 = Buffer.ReadLong
            .Conv(i).Data2 = Buffer.ReadLong
            .Conv(i).Data3 = Buffer.ReadLong
        Next

    End With

    Set Buffer = Nothing
End Sub

Private Sub HandleChatUpdate(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer, npcNum As Long, mT As String, o(1 To 4) As String, i As Long
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes data()
    
    npcNum = Buffer.ReadLong
    mT = Buffer.ReadString
    For i = 1 To 4
        o(i) = Buffer.ReadString
    Next

    Set Buffer = Nothing

    ' if npcNum is 0, exit the chat system
    If npcNum = 0 Then
        inChat = False
        HideWindow GetWindowIndex("winNpcChat")
        Exit Sub
    End If

    ' set chat going
    OpenNpcChat npcNum, mT, o
End Sub

Private Sub HandleStartTutorial(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    'inTutorial = True
    ' set the first message
    'SetTutorialState 1
End Sub

Private Sub HandleChatBubble(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer, TargetType As Long, target As Long, message As String, Colour As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    target = Buffer.ReadLong
    TargetType = Buffer.ReadLong
    message = Buffer.ReadString
    Colour = Buffer.ReadLong
    AddChatBubble target, TargetType, message, Colour
    Set Buffer = Nothing
End Sub

Private Sub HandleSetPlayerLoginToken(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    loginToken = Buffer.ReadString
    Set Buffer = Nothing
    ' try and login to game server
    AttemptLogin
End Sub

Private Sub HandlePlayerChars(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer, i As Long, winNum As Long, conNum As Long, isSlotEmpty(1 To MAX_CHARS) As Boolean, x As Long
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes data()
    usergroup = Buffer.ReadLong

    For i = 1 To MAX_CHARS
        CharName(i) = Trim$(Buffer.ReadString)
        CharSprite(i) = Buffer.ReadLong
        CharAccess(i) = Buffer.ReadLong
        CharClass(i) = Buffer.ReadLong
        ' set as empty or not
        If Not Len(Trim$(CharName(i))) > 0 Then isSlotEmpty(i) = True
    Next

    Set Buffer = Nothing
    
    HideWindows
    ShowWindow GetWindowIndex("winCharacters")
    
    ' set GUI window up
    winNum = GetWindowIndex("winCharacters")
    For i = 1 To MAX_CHARS
        conNum = GetControlIndex("winCharacters", "lblCharName_" & i)
        With Windows(winNum).Controls(conNum)
            If Not isSlotEmpty(i) Then
                .text = CharName(i)
            Else
                .text = "Blank Slot"
            End If
        End With
        ' hide/show buttons
        If isSlotEmpty(i) Then
            ' create button
            conNum = GetControlIndex("winCharacters", "btnCreateChar_" & i)
            Windows(winNum).Controls(conNum).visible = True
            ' select button
            conNum = GetControlIndex("winCharacters", "btnSelectChar_" & i)
            Windows(winNum).Controls(conNum).visible = False
            ' delete button
            conNum = GetControlIndex("winCharacters", "btnDelChar_" & i)
            Windows(winNum).Controls(conNum).visible = False
        Else
            ' create button
            conNum = GetControlIndex("winCharacters", "btnCreateChar_" & i)
            Windows(winNum).Controls(conNum).visible = False
            ' select button
            conNum = GetControlIndex("winCharacters", "btnSelectChar_" & i)
            Windows(winNum).Controls(conNum).visible = True
            ' delete button
            conNum = GetControlIndex("winCharacters", "btnDelChar_" & i)
            Windows(winNum).Controls(conNum).visible = True
        End If
    Next
End Sub

Private Sub HandleCancelAnimation(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim theIndex As Long, Buffer As clsBuffer, i As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    theIndex = Buffer.ReadLong
    Set Buffer = Nothing
    ' find the casting animation
    For i = 1 To MAX_BYTE
        If AnimInstance(i).LockType = TARGET_TYPE_PLAYER Then
            If AnimInstance(i).lockindex = theIndex Then
                If AnimInstance(i).isCasting = 1 Then
                    ' clear it
                    ClearAnimInstance i
                End If
            End If
        End If
    Next
End Sub

Private Sub HandlePlayerVariables(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer, i As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    
    For i = 1 To MAX_BYTE
        Player(MyIndex).Variable(i) = Buffer.ReadLong
    Next
    
    Set Buffer = Nothing
End Sub

Private Sub HandleEvent(ByVal index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()
    
    If Buffer.ReadLong = 1 Then
        inEvent = True
    Else
        inEvent = False
    End If
    eventNum = Buffer.ReadLong
    eventPageNum = Buffer.ReadLong
    eventCommandNum = Buffer.ReadLong
    
    Set Buffer = Nothing
End Sub
