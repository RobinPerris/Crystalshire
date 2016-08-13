Attribute VB_Name = "modText"
Option Explicit

'The size of a FVF vertex
Public Const FVF_Size As Long = 28

'Point API
Public Type POINTAPI
    x As Long
    Y As Long
End Type

Private Type CharVA
    Vertex(0 To 3) As Vertex
End Type

Private Type VFH
    BitmapWidth As Long
    BitmapHeight As Long
    CellWidth As Long
    CellHeight As Long
    BaseCharOffset As Byte
    CharWidth(0 To 255) As Byte
    CharVA(0 To 255) As CharVA
End Type

Private Type CustomFont
    HeaderInfo As VFH
    Texture As Direct3DTexture8
    RowPitch As Integer
    RowFactor As Single
    ColFactor As Single
    CharHeight As Byte
    TextureSize As POINTAPI
    xOffset As Long
    yOffset As Long
End Type

' Fonts
Public Enum Fonts
    ' Georgia
    georgia_16 = 1
    georgiaBold_16
    georgiaDec_16
    ' Rockwell
    rockwellDec_15
    rockwell_15
    rockwellDec_10
    ' Verdana
    verdana_12
    verdanaBold_12
    verdana_13
    ' count value
    Fonts_Count
End Enum

' Store the fonts
Public font() As CustomFont

' Chatbox
Public Type ChatStruct
    Text As String
    Color As Long
    visible As Boolean
    timer As Long
    Channel As Byte
End Type
Public Const ColourChar As String * 1 = "½"
Public Const ChatLines As Long = 200
Public Const ChatWidth As Long = 316
Public Chat(1 To ChatLines) As ChatStruct
Public chatLastRemove As Long
Public Const CHAT_DIFFERENCE_TIMER As Long = 500
Public Chat_HighIndex As Long
Public ChatScroll As Long

Sub LoadFonts()
    'Check if we have the device
    If D3DDevice.TestCooperativeLevel <> D3D_OK Then Exit Sub
    ' re-dim the fonts
    ReDim font(1 To Fonts.Fonts_Count - 1)
    ' load the fonts
    SetFont Fonts.georgia_16, "georgia_16", 256
    SetFont Fonts.georgiaBold_16, "georgiaBold_16", 256
    SetFont Fonts.georgiaDec_16, "georgiaDec_16", 256
    SetFont Fonts.rockwellDec_15, "rockwellDec_15", 256, 2, 2
    SetFont Fonts.rockwell_15, "rockwell_15", 256, 2, 2
    SetFont Fonts.verdana_12, "verdana_12", 256
    SetFont Fonts.verdanaBold_12, "verdanaBold_12", 256
    SetFont Fonts.rockwellDec_10, "rockwellDec_10", 256, 2, 2
End Sub

Sub SetFont(ByVal fontNum As Long, ByVal texName As String, ByVal size As Long, Optional ByVal xOffset As Long, Optional ByVal yOffset As Long)
Dim data() As Byte, f As Long, w As Long, h As Long, path As String
    ' set the path
    path = App.path & Path_Font & texName & ".png"
    ' load the texture
    f = FreeFile
    Open path For Binary As #f
        ReDim data(0 To LOF(f) - 1)
        Get #f, , data
    Close #f
    ' get size
    font(fontNum).TextureSize.x = ByteToInt(data(18), data(19))
    font(fontNum).TextureSize.Y = ByteToInt(data(22), data(23))
    ' set to struct
    Set font(fontNum).Texture = D3DX.CreateTextureFromFileInMemoryEx(D3DDevice, data(0), AryCount(data), font(fontNum).TextureSize.x, font(fontNum).TextureSize.Y, D3DX_DEFAULT, 0, D3DFMT_A8R8G8B8, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, 0, ByVal 0, ByVal 0)
    font(fontNum).xOffset = xOffset
    font(fontNum).yOffset = yOffset
    LoadFontHeader font(fontNum), texName & ".dat"
End Sub

Public Function GetColourString(ByVal colourNum As Long) As String
    Select Case colourNum
        Case 0 ' Black
            GetColourString = "Black"
        Case 1 ' Blue
            GetColourString = "Blue"
        Case 2 ' Green
            GetColourString = "Green"
        Case 3 ' Cyan
            GetColourString = "Cyan"
        Case 4 ' Red
            GetColourString = "Red"
        Case 5 ' Magenta
            GetColourString = "Magenta"
        Case 6 ' Brown
            GetColourString = "Brown"
        Case 7 ' Grey
            GetColourString = "Grey"
        Case 8 ' DarkGrey
            GetColourString = "Dark Grey"
        Case 9 ' BrightBlue
            GetColourString = "Bright Blue"
        Case 10 ' BrightGreen
            GetColourString = "Bright Green"
        Case 11 ' BrightCyan
            GetColourString = "Bright Cyan"
        Case 12 ' BrightRed
            GetColourString = "Bright Red"
        Case 13 ' Pink
            GetColourString = "Pink"
        Case 14 ' Yellow
            GetColourString = "Yellow"
        Case 15 ' White
            GetColourString = "White"
        Case 16 ' dark brown
            GetColourString = "Dark Brown"
        Case 17 ' gold
            GetColourString = "Gold"
        Case 18 ' light green
            GetColourString = "Light Green"
    End Select
End Function

Public Function DX8Colour(ByVal colourNum As Long, ByVal alpha As Long) As Long
    Select Case colourNum
        Case 0 ' Black
            DX8Colour = D3DColorARGB(alpha, 0, 0, 0)
        Case 1 ' Blue
            DX8Colour = D3DColorARGB(alpha, 16, 104, 237)
        Case 2 ' Green
            DX8Colour = D3DColorARGB(alpha, 119, 188, 84)
        Case 3 ' Cyan
            DX8Colour = D3DColorARGB(alpha, 16, 224, 237)
        Case 4 ' Red
            DX8Colour = D3DColorARGB(alpha, 201, 0, 0)
        Case 5 ' Magenta
            DX8Colour = D3DColorARGB(alpha, 255, 0, 255)
        Case 6 ' Brown
            DX8Colour = D3DColorARGB(alpha, 175, 149, 92)
        Case 7 ' Grey
            DX8Colour = D3DColorARGB(alpha, 192, 192, 192)
        Case 8 ' DarkGrey
            DX8Colour = D3DColorARGB(alpha, 128, 128, 128)
        Case 9 ' BrightBlue
            DX8Colour = D3DColorARGB(alpha, 126, 182, 240)
        Case 10 ' BrightGreen
            DX8Colour = D3DColorARGB(alpha, 126, 240, 137)
        Case 11 ' BrightCyan
            DX8Colour = D3DColorARGB(alpha, 157, 242, 242)
        Case 12 ' BrightRed
            DX8Colour = D3DColorARGB(alpha, 255, 0, 0)
        Case 13 ' Pink
            DX8Colour = D3DColorARGB(alpha, 255, 118, 221)
        Case 14 ' Yellow
            DX8Colour = D3DColorARGB(alpha, 255, 255, 0)
        Case 15 ' White
            DX8Colour = D3DColorARGB(alpha, 255, 255, 255)
        Case 16 ' dark brown
            DX8Colour = D3DColorARGB(alpha, 98, 84, 52)
        Case 17 ' gold
            DX8Colour = D3DColorARGB(alpha, 255, 215, 0)
        Case 18 ' light green
            DX8Colour = D3DColorARGB(alpha, 124, 205, 80)
    End Select
End Function

Sub LoadFontHeader(ByRef theFont As CustomFont, ByVal filename As String)
Dim FileNum As Byte
Dim LoopChar As Long
Dim Row As Single
Dim u As Single
Dim v As Single

    'Load the header information
    FileNum = FreeFile
    Open App.path & Path_Font & filename For Binary As #FileNum
    Get #FileNum, , theFont.HeaderInfo
    Close #FileNum
    
    'Calculate some common values
    theFont.CharHeight = theFont.HeaderInfo.CellHeight - 4
    theFont.RowPitch = theFont.HeaderInfo.BitmapWidth \ theFont.HeaderInfo.CellWidth
    theFont.ColFactor = theFont.HeaderInfo.CellWidth / theFont.HeaderInfo.BitmapWidth
    theFont.RowFactor = theFont.HeaderInfo.CellHeight / theFont.HeaderInfo.BitmapHeight

    'Cache the verticies used to draw the character (only requires setting the color and adding to the X/Y values)
    For LoopChar = 0 To 255
        'tU and tV value (basically tU = BitmapXPosition / BitmapWidth, and height for tV)
        Row = (LoopChar - theFont.HeaderInfo.BaseCharOffset) \ theFont.RowPitch
        u = ((LoopChar - theFont.HeaderInfo.BaseCharOffset) - (Row * theFont.RowPitch)) * theFont.ColFactor
        v = Row * theFont.RowFactor

        'Set the verticies
        With theFont.HeaderInfo.CharVA(LoopChar)
            .Vertex(0).Colour = D3DColorARGB(255, 0, 0, 0)   'Black is the most common color
            .Vertex(0).RHW = 1
            .Vertex(0).tu = u
            .Vertex(0).tv = v
            .Vertex(0).x = 0
            .Vertex(0).Y = 0
            .Vertex(0).z = 0
            .Vertex(1).Colour = D3DColorARGB(255, 0, 0, 0)
            .Vertex(1).RHW = 1
            .Vertex(1).tu = u + theFont.ColFactor
            .Vertex(1).tv = v
            .Vertex(1).x = theFont.HeaderInfo.CellWidth
            .Vertex(1).Y = 0
            .Vertex(1).z = 0
            .Vertex(2).Colour = D3DColorARGB(255, 0, 0, 0)
            .Vertex(2).RHW = 1
            .Vertex(2).tu = u
            .Vertex(2).tv = v + theFont.RowFactor
            .Vertex(2).x = 0
            .Vertex(2).Y = theFont.HeaderInfo.CellHeight
            .Vertex(2).z = 0
            .Vertex(3).Colour = D3DColorARGB(255, 0, 0, 0)
            .Vertex(3).RHW = 1
            .Vertex(3).tu = u + theFont.ColFactor
            .Vertex(3).tv = v + theFont.RowFactor
            .Vertex(3).x = theFont.HeaderInfo.CellWidth
            .Vertex(3).Y = theFont.HeaderInfo.CellHeight
            .Vertex(3).z = 0
        End With
    Next LoopChar
End Sub

Public Sub RenderText(ByRef UseFont As CustomFont, ByVal Text As String, ByVal x As Long, ByVal Y As Long, ByVal Color As Long, Optional ByVal alpha As Long = 255, Optional Shadow As Boolean = True)
Dim TempVA(0 To 3) As Vertex, TempStr() As String, count As Long, Ascii() As Byte, i As Long, j As Long, TempColor As Long, yOffset As Single, ignoreChar As Long, resetColor As Long
Dim tmpNum As Long

    ' set the color
    Color = DX8Colour(Color, alpha)

    'Check for valid text to render
    If LenB(Text) = 0 Then Exit Sub
    'Get the text into arrays (split by vbCrLf)
    TempStr = Split(Text, vbCrLf)
    'Set the temp color (or else the first character has no color)
    TempColor = Color
    resetColor = TempColor
    'Set the texture
    D3DDevice.SetTexture 0, UseFont.Texture
    CurrentTexture = -1
    ' set the position
    x = x - UseFont.xOffset
    Y = Y - UseFont.yOffset
    'Loop through each line if there are line breaks (vbCrLf)
    tmpNum = UBound(TempStr)

    For i = 0 To tmpNum
        If Len(TempStr(i)) > 0 Then
            yOffset = (i * UseFont.CharHeight) + (i * 3)
            count = 0
            'Convert the characters to the ascii value
            Ascii() = StrConv(TempStr(i), vbFromUnicode)
            'Loop through the characters
            tmpNum = Len(TempStr(i))
            For j = 1 To tmpNum
                ' check for colour change
                If Mid$(TempStr(i), j, 1) = ColourChar Then
                    Color = Val(Mid$(TempStr(i), j + 1, 2))
                    ' make sure the colour exists
                    If Color = -1 Then
                        TempColor = resetColor
                    Else
                        TempColor = DX8Colour(Color, alpha)
                    End If
                    ignoreChar = 3
                End If
                ' check if we're ignoring this character
                If ignoreChar > 0 Then
                    ignoreChar = ignoreChar - 1
                Else
                    'Copy from the cached vertex array to the temp vertex array
                    Call CopyMemory(TempVA(0), UseFont.HeaderInfo.CharVA(Ascii(j - 1)).Vertex(0), FVF_Size * 4)
                    'Set up the verticies
                    TempVA(0).x = x + count
                    TempVA(0).Y = Y + yOffset
                    TempVA(1).x = TempVA(1).x + x + count
                    TempVA(1).Y = TempVA(0).Y
                    TempVA(2).x = TempVA(0).x
                    TempVA(2).Y = TempVA(2).Y + TempVA(0).Y
                    TempVA(3).x = TempVA(1).x
                    TempVA(3).Y = TempVA(2).Y
                    'Set the colors
                    TempVA(0).Colour = TempColor
                    TempVA(1).Colour = TempColor
                    TempVA(2).Colour = TempColor
                    TempVA(3).Colour = TempColor
                    'Draw the verticies
                    Call D3DDevice.DrawPrimitiveUP(D3DPT_TRIANGLESTRIP, 2, TempVA(0), FVF_Size)
                    'Shift over the the position to render the next character
                    count = count + UseFont.HeaderInfo.CharWidth(Ascii(j - 1))
                End If
            Next j
        End If
    Next i
End Sub

Public Function TextWidth(ByRef UseFont As CustomFont, ByVal Text As String) As Long
Dim LoopI As Integer, tmpNum As Long, skipCount As Long

    'Make sure we have text
    If LenB(Text) = 0 Then Exit Function
    
    'Loop through the text
    tmpNum = Len(Text)
    For LoopI = 1 To tmpNum
        If Mid$(Text, LoopI, 1) = ColourChar Then skipCount = 3
        If skipCount > 0 Then
            skipCount = skipCount - 1
        Else
            TextWidth = TextWidth + UseFont.HeaderInfo.CharWidth(Asc(Mid$(Text, LoopI, 1)))
        End If
    Next LoopI
End Function

Public Function TextHeight(ByRef UseFont As CustomFont) As Long
    TextHeight = UseFont.HeaderInfo.CellHeight
End Function

Sub DrawActionMsg(ByVal index As Integer)
        Dim x As Long, Y As Long, i As Long, Time As Long
    Dim LenMsg As Long

    If ActionMsg(index).message = vbNullString Then Exit Sub

    ' how long we want each message to appear
    Select Case ActionMsg(index).Type

        Case ACTIONMsgSTATIC
            Time = 1500
            LenMsg = TextWidth(font(Fonts.rockwell_15), Trim$(ActionMsg(index).message))

            If ActionMsg(index).Y > 0 Then
                x = ActionMsg(index).x + Int(PIC_X \ 2) - (LenMsg / 2)
                Y = ActionMsg(index).Y + PIC_Y
            Else
                x = ActionMsg(index).x + Int(PIC_X \ 2) - (LenMsg / 2)
                Y = ActionMsg(index).Y - Int(PIC_Y \ 2) + 18
            End If

        Case ACTIONMsgSCROLL
            Time = 1500

            If ActionMsg(index).Y > 0 Then
                x = ActionMsg(index).x + Int(PIC_X \ 2) - ((Len(Trim$(ActionMsg(index).message)) \ 2) * 8)
                Y = ActionMsg(index).Y - Int(PIC_Y \ 2) - 2 - (ActionMsg(index).Scroll * 0.6)
                ActionMsg(index).Scroll = ActionMsg(index).Scroll + 1
            Else
                x = ActionMsg(index).x + Int(PIC_X \ 2) - ((Len(Trim$(ActionMsg(index).message)) \ 2) * 8)
                Y = ActionMsg(index).Y - Int(PIC_Y \ 2) + 18 + (ActionMsg(index).Scroll * 0.001)
                ActionMsg(index).Scroll = ActionMsg(index).Scroll + 1
            End If

            ActionMsg(index).alpha = ActionMsg(index).alpha - 5

            If ActionMsg(index).alpha <= 0 Then ClearActionMsg index: Exit Sub

        Case ACTIONMsgSCREEN
            Time = 3000

            ' This will kill any action screen messages that there in the system
            For i = MAX_BYTE To 1 Step -1

                If ActionMsg(i).Type = ACTIONMsgSCREEN Then
                    If i <> index Then
                        ClearActionMsg index
                        index = i
                    End If
                End If

            Next

            x = (400) - ((TextWidth(font(Fonts.rockwell_15), Trim$(ActionMsg(index).message)) \ 2))
            Y = 24
    End Select

    x = ConvertMapX(x)
    Y = ConvertMapY(Y)

    If ActionMsg(index).Created > 0 Then
        RenderText font(Fonts.rockwell_15), ActionMsg(index).message, x, Y, ActionMsg(index).Color, ActionMsg(index).alpha
    End If
End Sub

Public Function DrawMapEvents()
Dim x As Long, Y As Long, i As Long
    If frmEditor_Map.optEvents.value Then
        If Map.TileData.EventCount > 0 Then
            For i = 1 To Map.TileData.EventCount
                With Map.TileData.Events(i)
                    x = ((ConvertMapX(.x * PIC_X)) - 4) + (PIC_X * 0.5)
                    Y = ((ConvertMapY(.Y * PIC_Y)) - 7) + (PIC_Y * 0.5)
                End With
                RenderTexture Tex_Event, ConvertMapX(Map.TileData.Events(i).x * PIC_X), ConvertMapY(Map.TileData.Events(i).Y * PIC_Y), 0, 0, 32, 32, 32, 32
                RenderText font(Fonts.rockwellDec_10), "E", x, Y, BrightBlue
            Next
        End If
    End If
End Function

Public Function DrawMapAttributes()
Dim x As Long, Y As Long, tx As Long, ty As Long, theFont As Long

    theFont = Fonts.rockwellDec_10

    If frmEditor_Map.optAttribs.value Then
        For x = TileView.left To TileView.Right
            For Y = TileView.top To TileView.bottom
                If IsValidMapPoint(x, Y) Then
                    With Map.TileData.Tile(x, Y)
                        tx = ((ConvertMapX(x * PIC_X)) - 4) + (PIC_X * 0.5)
                        ty = ((ConvertMapY(Y * PIC_Y)) - 7) + (PIC_Y * 0.5)
                        If .Type > 0 Then RenderTexture Tex_Event, ConvertMapX(x * PIC_X), ConvertMapY(Y * PIC_Y), 0, 0, 32, 32, 32, 32
                        Select Case .Type
                            Case TILE_TYPE_BLOCKED
                                RenderText font(theFont), "B", tx, ty, BrightRed
                            Case TILE_TYPE_WARP
                                RenderText font(theFont), "W", tx, ty, BrightBlue
                            Case TILE_TYPE_ITEM
                                RenderText font(theFont), "I", tx, ty, White
                            Case TILE_TYPE_NPCAVOID
                                RenderText font(theFont), "N", tx, ty, White
                            Case TILE_TYPE_KEY
                                RenderText font(theFont), "K", tx, ty, White
                            Case TILE_TYPE_KEYOPEN
                                RenderText font(theFont), "O", tx, ty, White
                            Case TILE_TYPE_RESOURCE
                                RenderText font(theFont), "R", tx, ty, Green
                            Case TILE_TYPE_DOOR
                                RenderText font(theFont), "D", tx, ty, Brown
                            Case TILE_TYPE_NPCSPAWN
                                RenderText font(theFont), "S", tx, ty, Yellow
                            Case TILE_TYPE_SHOP
                                RenderText font(theFont), "S", tx, ty, BrightBlue
                            Case TILE_TYPE_SLIDE
                                RenderText font(theFont), "S", tx, ty, Pink
                            Case TILE_TYPE_CHAT
                                RenderText font(theFont), "C", tx, ty, Blue
                        End Select
                    End With
                End If
            Next
        Next
    End If
End Function

Public Sub AddText(ByVal Text As String, ByVal Color As Long, Optional ByVal alpha As Long = 255, Optional Channel As Byte = 0)
Dim i As Long

    Chat_HighIndex = 0
    ' Move the rest of it up
    For i = (ChatLines - 1) To 1 Step -1
        If Len(Chat(i).Text) > 0 Then
            If i > Chat_HighIndex Then Chat_HighIndex = i + 1
        End If
        Chat(i + 1) = Chat(i)
    Next
    
    Chat(1).Text = Text
    Chat(1).Color = Color
    Chat(1).visible = True
    Chat(1).timer = GetTickCount
    Chat(1).Channel = Channel
End Sub

Sub RenderChat()
Dim xO As Long, yO As Long, Colour As Long, yOffset As Long, rLines As Long, lineCount As Long
Dim tmpText As String, i As Long, isVisible As Boolean, topWidth As Long, tmpArray() As String, x As Long
    
    ' set the position
    xO = 19
    yO = ScreenHeight - 41 '545 + 14
    
    ' loop through chat
    rLines = 1
    i = 1 + ChatScroll
    Do While rLines <= 8
        If i > ChatLines Then Exit Do
        lineCount = 0
        ' exit out early if we come to a blank string
        If Len(Chat(i).Text) = 0 Then Exit Do
        ' get visible state
        isVisible = True
        If inSmallChat Then
            If Not Chat(i).visible Then isVisible = False
        End If
        If Options.channelState(Chat(i).Channel) = 0 Then isVisible = False
        ' make sure it's visible
        If isVisible Then
            ' render line
            Colour = Chat(i).Color
            ' check if we need to word wrap
            If TextWidth(font(Fonts.verdana_12), Chat(i).Text) > ChatWidth Then
                ' word wrap
                tmpText = WordWrap(font(Fonts.verdana_12), Chat(i).Text, ChatWidth, lineCount)
                ' can't have it going offscreen.
                If rLines + lineCount > 9 Then Exit Do
                ' continue on
                yOffset = yOffset - (14 * lineCount)
                RenderText font(Fonts.verdana_12), tmpText, xO, yO + yOffset, Colour
                rLines = rLines + lineCount
                ' set the top width
                tmpArray = Split(tmpText, vbNewLine)
                For x = 0 To UBound(tmpArray)
                    If TextWidth(font(Fonts.verdana_12), tmpArray(x)) > topWidth Then topWidth = TextWidth(font(Fonts.verdana_12), tmpArray(x))
                Next
            Else
                ' normal
                yOffset = yOffset - 14
                RenderText font(Fonts.verdana_12), Chat(i).Text, xO, yO + yOffset, Colour
                rLines = rLines + 1
                ' set the top width
                If TextWidth(font(Fonts.verdana_12), Chat(i).Text) > topWidth Then topWidth = TextWidth(font(Fonts.verdana_12), Chat(i).Text)
            End If
        End If
        ' increment chat pointer
        i = i + 1
    Loop
    
    ' get the height of the small chat box
    SetChatHeight rLines * 14
    SetChatWidth topWidth
End Sub

Public Sub WordWrap_Array(ByVal Text As String, ByVal MaxLineLen As Long, ByRef theArray() As String)
    Dim lineCount As Long, i As Long, size As Long, lastSpace As Long, b As Long, tmpNum As Long

    'Too small of text
    If Len(Text) < 2 Then
        ReDim theArray(1 To 1) As String
        theArray(1) = Text
        Exit Sub
    End If

    ' default values
    b = 1
    lastSpace = 1
    size = 0
    tmpNum = Len(Text)

    For i = 1 To tmpNum

        ' if it's a space, store it
        Select Case Mid$(Text, i, 1)
            Case " ": lastSpace = i
        End Select

        'Add up the size
        size = size + font(Fonts.georgiaDec_16).HeaderInfo.CharWidth(Asc(Mid$(Text, i, 1)))

        'Check for too large of a size
        If size > MaxLineLen Then
            'Check if the last space was too far back
            If i - lastSpace > 12 Then
                'Too far away to the last space, so break at the last character
                lineCount = lineCount + 1
                ReDim Preserve theArray(1 To lineCount) As String
                theArray(lineCount) = Trim$(Mid$(Text, b, (i - 1) - b))
                b = i - 1
                size = 0
            Else
                'Break at the last space to preserve the word
                lineCount = lineCount + 1
                ReDim Preserve theArray(1 To lineCount) As String
                theArray(lineCount) = Trim$(Mid$(Text, b, lastSpace - b))
                b = lastSpace + 1
                'Count all the words we ignored (the ones that weren't printed, but are before "i")
                size = TextWidth(font(Fonts.georgiaDec_16), Mid$(Text, lastSpace, i - lastSpace))
            End If
        End If

        ' Remainder
        If i = Len(Text) Then
            If b <> i Then
                lineCount = lineCount + 1
                ReDim Preserve theArray(1 To lineCount) As String
                theArray(lineCount) = theArray(lineCount) & Mid$(Text, b, i)
            End If
        End If
    Next
End Sub

Public Function WordWrap(theFont As CustomFont, ByVal Text As String, ByVal MaxLineLen As Integer, Optional ByRef lineCount As Long) As String
    Dim TempSplit() As String, TSLoop As Long, lastSpace As Long, size As Long, i As Long, b As Long, tmpNum As Long, skipCount As Long

    'Too small of text
    If Len(Text) < 2 Then
        WordWrap = Text
        Exit Function
    End If

    'Check if there are any line breaks - if so, we will support them
    TempSplit = Split(Text, vbNewLine)
    tmpNum = UBound(TempSplit)

    For TSLoop = 0 To tmpNum
        'Clear the values for the new line
        size = 0
        b = 1
        lastSpace = 1

        'Add back in the vbNewLines
        If TSLoop < UBound(TempSplit()) Then TempSplit(TSLoop) = TempSplit(TSLoop) & vbNewLine

        'Only check lines with a space
        If InStr(1, TempSplit(TSLoop), " ") Then
            'Loop through all the characters
            tmpNum = Len(TempSplit(TSLoop))

            For i = 1 To tmpNum
                'If it is a space, store it so we can easily break at it
                Select Case Mid$(TempSplit(TSLoop), i, 1)
                    Case " "
                        lastSpace = i
                    Case ColourChar
                        skipCount = 3
                End Select
                
                If skipCount > 0 Then
                    skipCount = skipCount - 1
                Else
                    'Add up the size
                    size = size + theFont.HeaderInfo.CharWidth(Asc(Mid$(TempSplit(TSLoop), i, 1)))
                    'Check for too large of a size
                    If size > MaxLineLen Then
                        'Check if the last space was too far back
                        If i - lastSpace > 12 Then
                            'Too far away to the last space, so break at the last character
                            WordWrap = WordWrap & Trim$(Mid$(TempSplit(TSLoop), b, (i - 1) - b)) & vbNewLine
                            lineCount = lineCount + 1
                            b = i - 1
                            size = 0
                        Else
                            'Break at the last space to preserve the word
                            WordWrap = WordWrap & Trim$(Mid$(TempSplit(TSLoop), b, lastSpace - b)) & vbNewLine
                            lineCount = lineCount + 1
                            b = lastSpace + 1
                            'Count all the words we ignored (the ones that weren't printed, but are before "i")
                            size = TextWidth(theFont, Mid$(TempSplit(TSLoop), lastSpace, i - lastSpace))
                        End If
                    End If
    
                    'This handles the remainder
                    If i = Len(TempSplit(TSLoop)) Then
                        If b <> i Then
                            WordWrap = WordWrap & Mid$(TempSplit(TSLoop), b, i)
                            lineCount = lineCount + 1
                        End If
                    End If
                End If
            Next i
        Else
            WordWrap = WordWrap & TempSplit(TSLoop)
        End If
    Next TSLoop
End Function

Public Sub DrawPlayerName(ByVal index As Long)
    Dim textX As Long, textY As Long, Text As String, textSize As Long, Colour As Long
    
    Text = Trim$(GetPlayerName(index))
    textSize = TextWidth(font(Fonts.rockwell_15), Text)
    ' get the colour
    Colour = White

    If Player(index).usergroup = 10 Or Player(index).usergroup = 11 Then Colour = Gold
    If GetPlayerAccess(index) > 0 Then Colour = Pink
    If GetPlayerPK(index) > 0 Then Colour = BrightRed
    textX = Player(index).x * PIC_X + Player(index).xOffset + (PIC_X \ 2) - (textSize \ 2)
    textY = Player(index).Y * PIC_Y + Player(index).yOffset - 32

    If GetPlayerSprite(index) >= 1 And GetPlayerSprite(index) <= Count_Char Then
        textY = GetPlayerY(index) * PIC_Y + Player(index).yOffset - (mTexture(Tex_Char(GetPlayerSprite(index))).h / 4) + 12
    End If

    Call RenderText(font(Fonts.rockwell_15), Text, ConvertMapX(textX), ConvertMapY(textY), Colour)
End Sub

Public Sub DrawNpcName(ByVal index As Long)
    Dim textX As Long, textY As Long, Text As String, textSize As Long, npcNum As Long, Colour As Long
    npcNum = MapNpc(index).num
    Text = Trim$(Npc(npcNum).name)
    textSize = TextWidth(font(Fonts.rockwell_15), Text)

    If Npc(npcNum).Behaviour = NPC_BEHAVIOUR_ATTACKONSIGHT Or Npc(npcNum).Behaviour = NPC_BEHAVIOUR_ATTACKWHENATTACKED Then
        ' get the colour
        If Npc(npcNum).Level <= GetPlayerLevel(MyIndex) - 3 Then
            Colour = Grey
        ElseIf Npc(npcNum).Level <= GetPlayerLevel(MyIndex) - 2 Then
            Colour = Green
        ElseIf Npc(npcNum).Level > GetPlayerLevel(MyIndex) Then
            Colour = Red
        Else
            Colour = White
        End If
    Else
        Colour = White
    End If

    textX = MapNpc(index).x * PIC_X + MapNpc(index).xOffset + (PIC_X \ 2) - (textSize \ 2)
    textY = MapNpc(index).Y * PIC_Y + MapNpc(index).yOffset - 32

    If Npc(npcNum).sprite >= 1 And Npc(npcNum).sprite <= Count_Char Then
        textY = MapNpc(index).Y * PIC_Y + MapNpc(index).yOffset - (mTexture(Tex_Char(Npc(npcNum).sprite)).h / 4) + 12
    End If

    Call RenderText(font(Fonts.rockwell_15), Text, ConvertMapX(textX), ConvertMapY(textY), Colour)
End Sub

Function GetColStr(Colour As Long)
    If Colour < 10 Then
        GetColStr = "0" & Colour
    Else
        GetColStr = Colour
    End If
End Function
