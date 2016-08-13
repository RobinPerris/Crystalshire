Attribute VB_Name = "modInput"
Option Explicit
' keyboard input
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

' Actual input
Public Sub CheckKeys()

    ' exit out if dialogue
    If diaIndex > 0 Then Exit Sub
    If GetAsyncKeyState(VK_W) >= 0 Then wDown = False
    If GetAsyncKeyState(VK_S) >= 0 Then sDown = False
    If GetAsyncKeyState(VK_A) >= 0 Then aDown = False
    If GetAsyncKeyState(VK_D) >= 0 Then dDown = False
    If GetAsyncKeyState(VK_UP) >= 0 Then upDown = False
    If GetAsyncKeyState(VK_DOWN) >= 0 Then downDown = False
    If GetAsyncKeyState(VK_LEFT) >= 0 Then leftDown = False
    If GetAsyncKeyState(VK_RIGHT) >= 0 Then rightDown = False
    If GetAsyncKeyState(VK_CONTROL) >= 0 Then ControlDown = False
    If GetAsyncKeyState(VK_SHIFT) >= 0 Then ShiftDown = False
    If GetAsyncKeyState(VK_TAB) >= 0 Then tabDown = False
End Sub

Public Sub CheckInputKeys()

    ' exit out if dialogue
    If diaIndex > 0 Then Exit Sub
    
    ' exit out if talking
    If Windows(GetWindowIndex("winChat")).Window.visible Then Exit Sub
    
    ' continue
    If GetKeyState(vbKeyShift) < 0 Then
        ShiftDown = True
    Else
        ShiftDown = False
    End If

    If GetKeyState(vbKeyControl) < 0 Then
        ControlDown = True
    Else
        ControlDown = False
    End If

    If GetKeyState(vbKeyTab) < 0 Then
        tabDown = True
    Else
        tabDown = False
    End If

    'Move Up
    If Not chatOn Then
        If GetKeyState(vbKeySpace) < 0 Then
            CheckMapGetItem
        End If

        ' move up
        If GetKeyState(vbKeyW) < 0 Then
            wDown = True
            sDown = False
            aDown = False
            dDown = False
            Exit Sub
        Else
            wDown = False
        End If

        'Move Right
        If GetKeyState(vbKeyD) < 0 Then
            wDown = False
            sDown = False
            aDown = False
            dDown = True
            Exit Sub
        Else
            dDown = False
        End If

        'Move down
        If GetKeyState(vbKeyS) < 0 Then
            wDown = False
            sDown = True
            aDown = False
            dDown = False
            Exit Sub
        Else
            sDown = False
        End If

        'Move left
        If GetKeyState(vbKeyA) < 0 Then
            wDown = False
            sDown = False
            aDown = True
            dDown = False
            Exit Sub
        Else
            aDown = False
        End If

        ' move up
        If GetKeyState(vbKeyUp) < 0 Then
            upDown = True
            leftDown = False

            downDown = False
            rightDown = False
            Exit Sub
        Else
            upDown = False
        End If

        'Move Right
        If GetKeyState(vbKeyRight) < 0 Then
            upDown = False
            leftDown = False

            downDown = False
            rightDown = True
            Exit Sub
        Else
            rightDown = False
        End If

        'Move down
        If GetKeyState(vbKeyDown) < 0 Then
            upDown = False
            leftDown = False

            downDown = True
            rightDown = False
            Exit Sub
        Else

            downDown = False
        End If

        'Move left
        If GetKeyState(vbKeyLeft) < 0 Then
            upDown = False
            leftDown = True

            downDown = False
            rightDown = False
            Exit Sub
        Else
            leftDown = False
        End If

    Else
        wDown = False
        sDown = False
        aDown = False
        dDown = False
        upDown = False
        leftDown = False

        downDown = False
        rightDown = False
    End If

End Sub

Public Sub HandleKeyPresses(ByVal KeyAscii As Integer)
    Dim chatText As String, name As String, i As Long, n As Long, Command() As String, Buffer As clsBuffer, tmpNum As Long
    
    ' check if we're skipping video
    If KeyAscii = vbKeyEscape Then
        ' hide options screen
        HideWindow GetWindowIndex("winOptions")
        CloseComboMenu
        ' handle the video
        If videoPlaying Then
            videoPlaying = False
            FadeAlpha = 0
            frmMain.picIntro.visible = False
            StopIntro
            Exit Sub
        End If
        If Windows(GetWindowIndex("winEscMenu")).Window.visible Then
            ' hide it
            HideWindow GetWindowIndex("winBlank")
            HideWindow GetWindowIndex("winEscMenu")
            Exit Sub
        Else
            ' show them
            ShowWindow GetWindowIndex("winBlank"), True
            ShowWindow GetWindowIndex("winEscMenu"), True
            Exit Sub
        End If
    End If
    
    If InGame Then
    chatText = Windows(GetWindowIndex("winChat")).Controls(GetControlIndex("winChat", "txtChat")).text
    End If
    
    ' Do we have an active window
    If activeWindow > 0 Then
        ' make sure it's visible
        If Windows(activeWindow).Window.visible Then
            ' Do we have an active control
            If Windows(activeWindow).activeControl > 0 Then
                ' Do our thing
                With Windows(activeWindow).Controls(Windows(activeWindow).activeControl)
                    ' Handle input
                    Select Case KeyAscii
                        Case vbKeyBack
                            If LenB(.text) > 0 Then
                                .text = left$(.text, Len(.text) - 1)
                            End If
                        Case vbKeyReturn
                            ' override for function callbacks
                            If .entCallBack(entStates.Enter) > 0 Then
                                entCallBack .entCallBack(entStates.Enter), activeWindow, Windows(activeWindow).activeControl, 0, 0
                                Exit Sub
                            Else
                                n = 0
                                For i = Windows(activeWindow).ControlCount To 1 Step -1
                                    If i > Windows(activeWindow).activeControl Then
                                        If SetActiveControl(activeWindow, i) Then n = i
                                    End If
                                Next
                                If n = 0 Then
                                    For i = Windows(activeWindow).ControlCount To 1 Step -1
                                        SetActiveControl activeWindow, i
                                    Next
                                End If
                            End If
                        Case vbKeyTab
                            n = 0
                            For i = Windows(activeWindow).ControlCount To 1 Step -1
                                If i > Windows(activeWindow).activeControl Then
                                    If SetActiveControl(activeWindow, i) Then n = i
                                End If
                            Next
                            If n = 0 Then
                                For i = Windows(activeWindow).ControlCount To 1 Step -1
                                    SetActiveControl activeWindow, i
                                Next
                            End If
                        Case Else
                            .text = .text & ChrW$(KeyAscii)
                    End Select
                    ' exit out early - if not chatting
                    If Windows(activeWindow).Window.name <> "winChat" Then Exit Sub
                End With
            End If
        End If
    End If

    ' exit out early if we're not ingame
    If Not InGame Then Exit Sub
    
    Select Case KeyAscii
        Case vbKeyEscape
            ' hide options screen
            HideWindow GetWindowIndex("winOptions")
            CloseComboMenu
            ' hide/show chat window
            If Windows(GetWindowIndex("winChat")).Window.visible Then
                Windows(GetWindowIndex("winChat")).Controls(GetControlIndex("winChat", "txtChat")).text = vbNullString
                HideChat
                inSmallChat = True
                Exit Sub
            End If
            
            If Windows(GetWindowIndex("winEscMenu")).Window.visible Then
                ' hide it
                HideWindow GetWindowIndex("winBlank")
                HideWindow GetWindowIndex("winEscMenu")
            Else
                ' show them
                ShowWindow GetWindowIndex("winBlank"), True
                ShowWindow GetWindowIndex("winEscMenu"), True
            End If
            ' exit out early
            Exit Sub
        Case 105
            ' hide/show inventory
            If Not Windows(GetWindowIndex("winChat")).Window.visible Then btnMenu_Inv
        Case 99
            ' hide/show inventory
            If Not Windows(GetWindowIndex("winChat")).Window.visible Then btnMenu_Char
        Case 109
            ' hide/show skills
            If Not Windows(GetWindowIndex("winChat")).Window.visible Then btnMenu_Skills
    End Select
    
    ' handles hotbar
    If inSmallChat Then
        For i = 1 To 9
            If KeyAscii = 48 + i Then
                SendHotbarUse i
            End If
            If KeyAscii = 48 Then SendHotbarUse 10
        Next
    End If

    ' Handle when the player presses the return key
    If KeyAscii = vbKeyReturn Then
        If Windows(GetWindowIndex("winChatSmall")).Window.visible Then
            ShowChat
            inSmallChat = False
            Exit Sub
        End If
    
        ' Broadcast message
        If left$(chatText, 1) = "'" Then
            chatText = Mid$(chatText, 2, Len(chatText) - 1)

            If Len(chatText) > 0 Then
                Call BroadcastMsg(chatText)
            End If

            Windows(GetWindowIndex("winChat")).Controls(GetControlIndex("winChat", "txtChat")).text = vbNullString
            HideChat
            Exit Sub
        End If

        ' Emote message
        If left$(chatText, 1) = "-" Then
            chatText = Mid$(chatText, 2, Len(chatText) - 1)

            If Len(chatText) > 0 Then
                Call EmoteMsg(chatText)
            End If

            Windows(GetWindowIndex("winChat")).Controls(GetControlIndex("winChat", "txtChat")).text = vbNullString
            HideChat
            Exit Sub
        End If

        ' Player message
        If left$(chatText, 1) = "!" Then
            Exit Sub
            chatText = Mid$(chatText, 2, Len(chatText) - 1)
            name = vbNullString
            ' Get the desired player from the user text
            tmpNum = Len(chatText)

            For i = 1 To tmpNum

                If Mid$(chatText, i, 1) <> Space$(1) Then
                    name = name & Mid$(chatText, i, 1)
                Else
                    Exit For
                End If

            Next

            chatText = Mid$(chatText, i, Len(chatText) - 1)

            ' Make sure they are actually sending something
            If Len(chatText) - i > 0 Then
                chatText = Mid$(chatText, i + 1, Len(chatText) - i)
                ' Send the message to the player
                Call PlayerMsg(chatText, name)
            Else
                Call AddText("Usage: !playername (message)", AlertColor)
            End If

            Windows(GetWindowIndex("winChat")).Controls(GetControlIndex("winChat", "txtChat")).text = vbNullString
            HideChat
            Exit Sub
        End If

        If left$(chatText, 1) = "/" Then
            Command = Split(chatText, Space$(1))

            Select Case Command(0)

                Case "/help"
                    Call AddText("Social Commands:", HelpColor)
                    Call AddText("'msghere = Global Message", HelpColor)
                    Call AddText("-msghere = Emote Message", HelpColor)
                    Call AddText("!namehere msghere = Player Message", HelpColor)
                    Call AddText("Available Commands: /who, /fps, /fpslock, /gui, /maps", HelpColor)

                Case "/maps"
                    ClearMapCache

                Case "/gui"
                    hideGUI = Not hideGUI

                Case "/info"

                    ' Checks to make sure we have more than one string in the array
                    If UBound(Command) < 1 Then
                        AddText "Usage: /info (name)", AlertColor
                        GoTo continue
                    End If

                    If IsNumeric(Command(1)) Then
                        AddText "Usage: /info (name)", AlertColor
                        GoTo continue
                    End If

                    Set Buffer = New clsBuffer
                    Buffer.WriteLong CPlayerInfoRequest
                    Buffer.WriteString Command(1)
                    SendData Buffer.ToArray()
                    Set Buffer = Nothing

                    ' Whos Online
                Case "/who"
                    SendWhosOnline

                    ' Checking fps
                Case "/fps"
                    BFPS = Not BFPS

                    ' toggle fps lock
                Case "/fpslock"
                    FPS_Lock = Not FPS_Lock

                    ' Request stats
                Case "/stats"
                    Set Buffer = New clsBuffer
                    Buffer.WriteLong CGetStats
                    SendData Buffer.ToArray()
                    Set Buffer = Nothing

                    ' // Monitor Admin Commands //
                    ' Kicking a player
                Case "/kick"

                    If GetPlayerAccess(MyIndex) < ADMIN_MONITOR Then GoTo continue
                    If UBound(Command) < 1 Then
                        AddText "Usage: /kick (name)", AlertColor
                        GoTo continue
                    End If

                    If IsNumeric(Command(1)) Then
                        AddText "Usage: /kick (name)", AlertColor
                        GoTo continue
                    End If

                    SendKick Command(1)

                    ' // Mapper Admin Commands //
                    ' Location
                Case "/loc"

                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo continue
                    BLoc = Not BLoc

                    ' Map Editor
                Case "/editmap"

                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo continue
                    SendRequestEditMap

                    ' Warping to a player
                Case "/warpmeto"

                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo continue
                    If UBound(Command) < 1 Then
                        AddText "Usage: /warpmeto (name)", AlertColor
                        GoTo continue
                    End If

                    If IsNumeric(Command(1)) Then
                        AddText "Usage: /warpmeto (name)", AlertColor
                        GoTo continue
                    End If

                    GettingMap = True
                    WarpMeTo Command(1)

                    ' Warping a player to you
                Case "/warptome"

                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo continue
                    If UBound(Command) < 1 Then
                        AddText "Usage: /warptome (name)", AlertColor
                        GoTo continue
                    End If

                    If IsNumeric(Command(1)) Then
                        AddText "Usage: /warptome (name)", AlertColor
                        GoTo continue
                    End If

                    WarpToMe Command(1)

                    ' Warping to a map
                Case "/warpto"

                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo continue
                    If UBound(Command) < 1 Then
                        AddText "Usage: /warpto (map #)", AlertColor
                        GoTo continue
                    End If

                    If Not IsNumeric(Command(1)) Then
                        AddText "Usage: /warpto (map #)", AlertColor
                        GoTo continue
                    End If

                    n = CLng(Command(1))

                    ' Check to make sure its a valid map #
                    If n > 0 And n <= MAX_MAPS Then
                        GettingMap = True
                        Call WarpTo(n)
                    Else
                        Call AddText("Invalid map number.", Red)
                    End If

                    ' Setting sprite
                Case "/setsprite"

                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo continue
                    If UBound(Command) < 1 Then
                        AddText "Usage: /setsprite (sprite #)", AlertColor
                        GoTo continue
                    End If

                    If Not IsNumeric(Command(1)) Then
                        AddText "Usage: /setsprite (sprite #)", AlertColor
                        GoTo continue
                    End If

                    SendSetSprite CLng(Command(1))

                    ' Map report
                Case "/mapreport"

                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo continue
                    SendMapReport

                    ' Respawn request
                Case "/respawn"

                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo continue
                    SendMapRespawn

                    ' MOTD change
                Case "/motd"

                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo continue
                    If UBound(Command) < 1 Then
                        AddText "Usage: /motd (new motd)", AlertColor
                        GoTo continue
                    End If

                    SendMOTDChange Right$(chatText, Len(chatText) - 5)

                    ' Check the ban list
                Case "/banlist"

                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo continue
                    SendBanList

                    ' Banning a player
                Case "/ban"

                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo continue
                    If UBound(Command) < 1 Then
                        AddText "Usage: /ban (name)", AlertColor
                        GoTo continue
                    End If

                    SendBan Command(1)

                    ' // Developer Admin Commands //
                    ' Editing item request
                Case "/edititem"

                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo continue
                    SendRequestEditItem

                    ' editing conv request
                Case "/editconv"

                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo continue
                    SendRequestEditConv

                    ' Editing animation request
                Case "/editanimation"

                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo continue
                    SendRequestEditAnimation

                    ' Editing npc request
                Case "/editnpc"

                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo continue
                    SendRequestEditNpc

                Case "/editresource"

                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo continue
                    SendRequestEditResource

                    ' Editing shop request
                Case "/editshop"

                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo continue
                    SendRequestEditShop

                    ' Editing spell request
                Case "/editspell"

                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo continue
                    SendRequestEditSpell

                    ' // Creator Admin Commands //
                    ' Giving another player access
                Case "/setaccess"

                    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then GoTo continue
                    If UBound(Command) < 2 Then
                        AddText "Usage: /setaccess (name) (access)", AlertColor
                        GoTo continue
                    End If

                    If IsNumeric(Command(1)) Or Not IsNumeric(Command(2)) Then
                        AddText "Usage: /setaccess (name) (access)", AlertColor
                        GoTo continue
                    End If

                    SendSetAccess Command(1), CLng(Command(2))

                    ' Ban destroy
                Case "/destroybanlist"

                    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then GoTo continue
                    SendBanDestroy

                    ' Packet debug mode
                Case "/debug"

                    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then GoTo continue
                    DEBUG_MODE = (Not DEBUG_MODE)

                Case Else
                    AddText "Not a valid command!", HelpColor
            End Select

            'continue label where we go instead of exiting the sub
continue:
            Windows(GetWindowIndex("winChat")).Controls(GetControlIndex("winChat", "txtChat")).text = vbNullString
            HideChat
            Exit Sub
        End If

        ' Say message
        If Len(chatText) > 0 Then
            Call SayMsg(chatText)
        End If

        Windows(GetWindowIndex("winChat")).Controls(GetControlIndex("winChat", "txtChat")).text = vbNullString
        
        ' hide/show chat window
        If Windows(GetWindowIndex("winChat")).Window.visible Then HideChat
        Exit Sub
    End If
    
    ' hide/show chat window
    If Windows(GetWindowIndex("winChatSmall")).Window.visible Then
        Exit Sub
    End If
End Sub
