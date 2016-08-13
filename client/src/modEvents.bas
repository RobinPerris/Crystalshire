Attribute VB_Name = "modEvents"
Option Explicit

Private Declare Function SendMessageByNum Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Const LB_SETHORIZONTALEXTENT = &H194

' temporary event
Public cpEvent As EventRec

Sub CopyEvent_Map(x As Long, y As Long)
Dim count As Long, i As Long
    count = map.TileData.EventCount
    If count = 0 Then Exit Sub
    
    For i = 1 To count
        If map.TileData.Events(i).x = x And map.TileData.Events(i).y = y Then
            ' copy it
            CopyMemory ByVal VarPtr(cpEvent), ByVal VarPtr(map.TileData.Events(i)), LenB(map.TileData.Events(i))
            ' exit
            Exit Sub
        End If
    Next
End Sub

Sub PasteEvent_Map(x As Long, y As Long)
Dim count As Long, i As Long, eventNum As Long
    count = map.TileData.EventCount
    
    If count > 0 Then
        For i = 1 To count
            If map.TileData.Events(i).x = x And map.TileData.Events(i).y = y Then
                ' already an event - paste over it
                eventNum = i
            End If
        Next
    End If
    
    ' couldn't find one - create one
    If eventNum = 0 Then
        ' increment count
        AddEvent x, y, True
        eventNum = count + 1
    End If
    
    ' copy it
    CopyMemory ByVal VarPtr(map.TileData.Events(eventNum)), ByVal VarPtr(cpEvent), LenB(cpEvent)
    
    ' set position
    map.TileData.Events(eventNum).x = x
    map.TileData.Events(eventNum).y = y
End Sub

Sub AddEvent(x As Long, y As Long, Optional ByVal cancelLoad As Boolean = False)
Dim count As Long, pageCount As Long, i As Long
    count = map.TileData.EventCount + 1
    ' make sure there's not already an event
    If count - 1 > 0 Then
        For i = 1 To count - 1
            If map.TileData.Events(i).x = x And map.TileData.Events(i).y = y Then
                ' already an event - edit it
                If Not cancelLoad Then EventEditorInit i
                Exit Sub
            End If
        Next
    End If
    ' increment count
    map.TileData.EventCount = count
    ReDim Preserve map.TileData.Events(1 To count)
    ' set the new event
    map.TileData.Events(count).x = x
    map.TileData.Events(count).y = y
    ' give it a new page
    pageCount = map.TileData.Events(count).pageCount + 1
    map.TileData.Events(count).pageCount = pageCount
    ReDim Preserve map.TileData.Events(count).EventPage(1 To pageCount)
    ' load the editor
    If Not cancelLoad Then EventEditorInit count
End Sub

Sub DeleteEvent(x As Long, y As Long)
Dim count As Long, i As Long, lowIndex As Long
    If Not InMapEditor Then Exit Sub
    
    count = map.TileData.EventCount
    For i = 1 To count
        If map.TileData.Events(i).x = x And map.TileData.Events(i).y = y Then
            ' delete it
            ClearEvent i
            lowIndex = i
            Exit For
        End If
    Next
    
    ' not found anything
    If lowIndex = 0 Then Exit Sub
    
    ' move everything down an index
    For i = lowIndex To count - 1
        CopyEvent i + 1, i
    Next
    ' delete the last index
    ClearEvent count
    ' set the new count
    map.TileData.EventCount = count - 1
End Sub

Sub ClearEvent(eventNum As Long)
    Call ZeroMemory(ByVal VarPtr(map.TileData.Events(eventNum)), LenB(map.TileData.Events(eventNum)))
End Sub

Sub CopyEvent(original As Long, newone As Long)
    CopyMemory ByVal VarPtr(map.TileData.Events(newone)), ByVal VarPtr(map.TileData.Events(original)), LenB(map.TileData.Events(original))
End Sub

Sub EventEditorInit(eventNum As Long)
Dim i As Long
    EditorEvent = eventNum
    ' copy the event data to the temp event
    CopyMemory ByVal VarPtr(tmpEvent), ByVal VarPtr(map.TileData.Events(eventNum)), LenB(map.TileData.Events(eventNum))
    ' populate form
    With frmEditor_Events
        ' set the tabs
        .tabPages.Tabs.Clear
        For i = 1 To tmpEvent.pageCount
            .tabPages.Tabs.Add , , Str(i)
        Next
        ' items
        .cmbHasItem.Clear
        .cmbHasItem.AddItem "None"
        For i = 1 To MAX_ITEMS
            .cmbHasItem.AddItem i & ": " & Trim$(Item(i).name)
        Next
        ' variables
        .cmbPlayerVar.Clear
        .cmbPlayerVar.AddItem "None"
        For i = 1 To MAX_BYTE
            .cmbPlayerVar.AddItem i
        Next
        ' name
        .txtName.text = tmpEvent.name
        ' enable delete button
        If tmpEvent.pageCount > 1 Then
            .cmdDeletePage.enabled = True
        Else
            .cmdDeletePage.enabled = False
        End If
        .cmdPastePage.enabled = False
        ' set the commands frame
        .fraCommands.width = 417
        .fraCommands.height = 497
        ' set the dialogue frame
        .fraDialogue.width = 417
        .fraDialogue.height = 497
        ' Load page 1 to start off with
        curPageNum = 1
        EventEditorLoadPage curPageNum
    End With
    ' show the editor
    frmEditor_Events.Show
End Sub

Sub AddCommand(theType As EventType)
Dim count As Long
    ' update the array
    With tmpEvent.EventPage(curPageNum)
        count = .CommandCount + 1
        ReDim Preserve .Commands(1 To count)
        .CommandCount = count
        ' set the shit
        Select Case theType
            Case EventType.evAddText
                ' set the values
                .Commands(count).Type = EventType.evAddText
                .Commands(count).text = frmEditor_Events.txtAddText_Text.text
                .Commands(count).colour = frmEditor_Events.scrlAddText_Colour.value
                If frmEditor_Events.optAddText_Game.value Then
                    .Commands(count).channel = 0
                ElseIf frmEditor_Events.optAddText_Map.value Then
                    .Commands(count).channel = 1
                ElseIf frmEditor_Events.optAddText_Global.value Then
                    .Commands(count).channel = 2
                End If
            Case EventType.evShowChatBubble
                .Commands(count).Type = EventType.evShowChatBubble
                .Commands(count).text = frmEditor_Events.txtChatBubble.text
                .Commands(count).colour = frmEditor_Events.scrlChatBubble.value
                .Commands(count).TargetType = frmEditor_Events.cmbChatBubbleType.ListIndex
                .Commands(count).target = frmEditor_Events.cmbChatBubble.ListIndex
            Case EventType.evPlayerVar
                .Commands(count).Type = EventType.evPlayerVar
                .Commands(count).target = frmEditor_Events.cmbVariable.ListIndex
                .Commands(count).colour = Val(frmEditor_Events.txtVariable.text)
            Case EventType.evWarpPlayer
                .Commands(count).Type = EventType.evWarpPlayer
                .Commands(count).x = frmEditor_Events.scrlWPX.value
                .Commands(count).y = frmEditor_Events.scrlWPY.value
                .Commands(count).target = frmEditor_Events.scrlWPMap.value
        End Select
    End With
    ' re-list the commands
    EventListCommands
End Sub

Sub EditCommand()
    With tmpEvent.EventPage(curPageNum).Commands(curCommand)
        Select Case .Type
            Case EventType.evAddText
                .text = frmEditor_Events.txtAddText_Text.text
                .colour = frmEditor_Events.scrlAddText_Colour.value
                If frmEditor_Events.optAddText_Game.value Then
                    .channel = 0
                ElseIf frmEditor_Events.optAddText_Map.value Then
                    .channel = 1
                ElseIf frmEditor_Events.optAddText_Global.value Then
                    .channel = 2
                End If
            Case EventType.evShowChatBubble
                .text = frmEditor_Events.txtChatBubble.text
                .colour = frmEditor_Events.scrlChatBubble.value
                .TargetType = frmEditor_Events.cmbChatBubbleType.ListIndex
                .target = frmEditor_Events.cmbChatBubble.ListIndex
            Case EventType.evPlayerVar
                .target = frmEditor_Events.cmbVariable.ListIndex
                .colour = Val(frmEditor_Events.txtVariable.text)
            Case EventType.evWarpPlayer
                .x = frmEditor_Events.scrlWPX.value
                .y = frmEditor_Events.scrlWPY.value
        End Select
    End With
    ' re-list the commands
    EventListCommands
End Sub

Sub EventListCommands()
Dim i As Long, count As Long
    frmEditor_Events.lstCommands.Clear
    ' check if there are any
    count = tmpEvent.EventPage(curPageNum).CommandCount
    If count > 0 Then
        ' list them
        For i = 1 To count
            With tmpEvent.EventPage(curPageNum).Commands(i)
                Select Case .Type
                    Case EventType.evAddText
                        ListCommandAdd "@>Add Text: " & .text & " - Colour: " & GetColourString(.colour) & " - Channel: " & .channel
                    Case EventType.evShowChatBubble
                        ListCommandAdd "@>Show Chat Bubble: " & .text & " - Colour: " & GetColourString(.colour) & " - Target Type: " & .TargetType & " - Target: " & .target
                    Case EventType.evPlayerVar
                        ListCommandAdd "@>Change variable #" & .target & " to " & .colour
                    Case EventType.evWarpPlayer
                        ListCommandAdd "@>Warp Player to Map #" & .target & ", X: " & .x & ", Y: " & .y
                    Case Else
                        ListCommandAdd "@>Unknown"
                End Select
            End With
        Next
    Else
        frmEditor_Events.lstCommands.AddItem "@>"
    End If
    frmEditor_Events.lstCommands.ListIndex = 0
    curCommand = 1
End Sub

Sub ListCommandAdd(s As String)
Static x As Long
    frmEditor_Events.lstCommands.AddItem s
    ' scrollbar
    If x < frmEditor_Events.TextWidth(s & "  ") Then
       x = frmEditor_Events.TextWidth(s & "  ")
      If frmEditor_Events.ScaleMode = vbTwips Then x = x / Screen.TwipsPerPixelX ' if twips change to pixels
      SendMessageByNum frmEditor_Events.lstCommands.hWnd, LB_SETHORIZONTALEXTENT, x, 0
    End If
End Sub

Sub EventEditorLoadPage(pageNum As Long)
    ' populate form
    With tmpEvent.EventPage(pageNum)
        GraphicSelX = .GraphicX
        GraphicSelY = .GraphicY
        frmEditor_Events.cmbGraphic.ListIndex = .GraphicType
        frmEditor_Events.cmbHasItem.ListIndex = .HasItemNum
        frmEditor_Events.cmbMoveFreq.ListIndex = .MoveFreq
        frmEditor_Events.cmbMoveSpeed.ListIndex = .MoveSpeed
        frmEditor_Events.cmbMoveType.ListIndex = .MoveType
        frmEditor_Events.cmbPlayerVar.ListIndex = .PlayerVarNum
        frmEditor_Events.cmbPriority.ListIndex = .Priority
        frmEditor_Events.cmbSelfSwitch.ListIndex = .SelfSwitchNum
        frmEditor_Events.cmbTrigger.ListIndex = .Trigger
        frmEditor_Events.chkDirFix.value = .DirFix
        frmEditor_Events.chkHasItem.value = .chkHasItem
        frmEditor_Events.chkPlayerVar.value = .chkPlayerVar
        frmEditor_Events.chkSelfSwitch.value = .chkSelfSwitch
        frmEditor_Events.chkStepAnim.value = .StepAnim
        frmEditor_Events.chkWalkAnim.value = .WalkAnim
        frmEditor_Events.chkWalkThrough.value = .WalkThrough
        frmEditor_Events.txtPlayerVariable = .PlayerVariable
        frmEditor_Events.scrlGraphic.value = .Graphic
        If .chkHasItem = 0 Then frmEditor_Events.cmbHasItem.enabled = False Else frmEditor_Events.cmbHasItem.enabled = True
        If .chkSelfSwitch = 0 Then frmEditor_Events.cmbSelfSwitch.enabled = False Else frmEditor_Events.cmbSelfSwitch.enabled = True
        If .chkPlayerVar = 0 Then
            frmEditor_Events.cmbPlayerVar.enabled = False
            frmEditor_Events.txtPlayerVariable.enabled = False
        Else
            frmEditor_Events.cmbPlayerVar.enabled = True
            frmEditor_Events.txtPlayerVariable.enabled = True
        End If
        ' show the commands
        EventListCommands
    End With
End Sub

Sub EventEditorOK()
    ' copy the event data from the temp event
    CopyMemory ByVal VarPtr(map.TileData.Events(EditorEvent)), ByVal VarPtr(tmpEvent), LenB(tmpEvent)
    ' unload the form
    Unload frmEditor_Events
End Sub
