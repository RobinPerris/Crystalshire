Attribute VB_Name = "modInterface"
Option Explicit

' Entity Types
Public Enum EntityTypes
    entLabel = 1
    entWindow
    entButton
    entTextBox
    entScrollbar
    entPictureBox
    entCheckbox
    entCombobox
    entCombomenu
End Enum

' Design Types
Public Enum DesignTypes
    ' Boxes
    desWood = 1
    desWood_Small
    desWood_Empty
    desGreen
    desGreen_Hover
    desGreen_Click
    desRed
    desRed_Hover
    desRed_Click
    desBlue
    desBlue_Hover
    desBlue_Click
    desOrange
    desOrange_Hover
    desOrange_Click
    desGrey
    desDescPic
    ' Windows
    desWin_Black
    desWin_Norm
    desWin_NoBar
    desWin_Empty
    desWin_Desc
    desWin_Shadow
    desWin_Party
    ' Pictures
    desParchment
    desBlackOval
    ' Textboxes
    desTextBlack
    desTextWhite
    desTextBlack_Sq
    ' Checkboxes
    desChkNorm
    desChkChat
    desChkCustom_Buying
    desChkCustom_Selling
    ' Right-click Menu
    desMenuHeader
    desMenuOption
    ' Comboboxes
    desComboNorm
    desComboMenuNorm
    ' tile Selection
    desTileBox
End Enum

' Button States
Public Enum entStates
    Normal = 0
    Hover
    MouseDown
    MouseMove
    MouseUp
    DblClick
    Enter
    ' Count
    state_Count
End Enum

' Alignment
Public Enum Alignment
    alignLeft = 0
    alignRight
    alignCentre
End Enum

' Part Types
Public Enum PartType
    part_None = 0
    Part_Item
    Part_spell
End Enum

' Origins
Public Enum PartTypeOrigins
    origin_None = 0
    origin_Inventory
    origin_Hotbar
    origin_Spells
End Enum

' Entity UDT
Public Type EntityRec
    ' constants
    name As String
    ' values
    Type As Byte
    top As Long
    left As Long
    width As Long
    height As Long
    enabled As Boolean
    visible As Boolean
    canDrag As Boolean
    Max As Long
    Min As Long
    value As Long
    text As String
    image(0 To entStates.state_Count - 1) As Long
    design(0 To entStates.state_Count - 1) As Long
    entCallBack(0 To entStates.state_Count - 1) As Long
    alpha As Long
    clickThrough As Boolean
    xOffset As Long
    yOffset As Long
    align As Byte
    font As Long
    textColour As Long
    textColour_Hover As Long
    textColour_Click As Long
    zChange As Byte
    onDraw As Long
    origLeft As Long
    origTop As Long
    tooltip As String
    group As Long
    list() As String
    activated As Boolean
    linkedToWin As Long
    linkedToCon As Long
    ' window
    icon As Long
    ' textbox
    isCensor As Boolean
    ' temp
    state As entStates
    movedX As Long
    movedY As Long
    zOrder As Long
End Type

' For small parts
Public Type EntityPartRec
    Type As PartType
    Origin As PartTypeOrigins
    value As Long
    Slot As Long
End Type

' Window UDT
Public Type WindowRec
    Window As EntityRec
    Controls() As EntityRec
    ControlCount As Long
    activeControl As Long
End Type

' actual GUI
Public Windows() As WindowRec
Public WindowCount As Long
Public activeWindow As Long
' GUI parts
Public DragBox As EntityPartRec
' Used for automatically the zOrder
Private zOrder_Win As Long
Private zOrder_Con As Long

Public Sub CreateEntity(winNum As Long, zOrder As Long, name As String, tType As EntityTypes, ByRef design() As Long, ByRef image() As Long, ByRef entCallBack() As Long, _
   Optional left As Long, Optional top As Long, Optional width As Long, Optional height As Long, Optional visible As Boolean = True, Optional canDrag As Boolean, Optional Max As Long, _
   Optional Min As Long, Optional value As Long, Optional text As String, Optional align As Byte, Optional font As Long = Fonts.georgia_16, Optional textColour As Long = White, _
   Optional alpha As Long = 255, Optional clickThrough As Boolean, Optional xOffset As Long, Optional yOffset As Long, Optional zChange As Byte, Optional ByVal icon As Long, _
   Optional ByVal onDraw As Long, Optional isActive As Boolean, Optional isCensor As Boolean, Optional textColour_Hover As Long, Optional textColour_Click As Long, _
   Optional tooltip As String, Optional group As Long)
    Dim i As Long

    ' check if it's a legal number
    If winNum <= 0 Or winNum > WindowCount Then
        Exit Sub
    End If

    ' re-dim the control array
    With Windows(winNum)
        .ControlCount = .ControlCount + 1
        ReDim Preserve .Controls(1 To .ControlCount) As EntityRec
    End With

    ' Set the new control values
    With Windows(winNum).Controls(Windows(winNum).ControlCount)
        .name = name
        .Type = tType

        ' loop through states
        For i = 0 To entStates.state_Count - 1
            .design(i) = design(i)
            .image(i) = image(i)
            .entCallBack(i) = entCallBack(i)
        Next

        .left = left
        .top = top
        .origLeft = left
        .origTop = top
        .width = width
        .height = height
        .visible = visible
        .canDrag = canDrag
        .Max = Max
        .Min = Min
        .value = value
        .text = text
        .align = align
        .font = font
        .textColour = textColour
        .textColour_Hover = textColour_Hover
        .textColour_Click = textColour_Click
        .alpha = alpha
        .clickThrough = clickThrough
        .xOffset = xOffset
        .yOffset = yOffset
        .zChange = zChange
        .zOrder = zOrder
        .enabled = True
        .icon = icon
        .onDraw = onDraw
        .isCensor = isCensor
        .tooltip = tooltip
        .group = group
        ReDim .list(0 To 0) As String
    End With
    
    ' set the active control
    If isActive Then Windows(winNum).activeControl = Windows(winNum).ControlCount

    ' set the zOrder
    zOrder_Con = zOrder_Con + 1
End Sub

Public Sub UpdateZOrder(winNum As Long, Optional forced As Boolean = False)
    Dim i As Long
    Dim oldZOrder As Long

    With Windows(winNum).Window

        If Not forced Then If .zChange = 0 Then Exit Sub
        If .zOrder = WindowCount Then Exit Sub
        oldZOrder = .zOrder

        For i = 1 To WindowCount

            If Windows(i).Window.zOrder > oldZOrder Then
                Windows(i).Window.zOrder = Windows(i).Window.zOrder - 1
            End If

        Next

        .zOrder = WindowCount
    End With

End Sub

Public Sub SortWindows()
    Dim tempWindow As WindowRec
    Dim i As Long, x As Long
    x = 1

    While x <> 0
        x = 0

        For i = 1 To WindowCount - 1

            If Windows(i).Window.zOrder > Windows(i + 1).Window.zOrder Then
                tempWindow = Windows(i)
                Windows(i) = Windows(i + 1)
                Windows(i + 1) = tempWindow
                x = 1
            End If

        Next

    Wend

End Sub

Public Sub RenderEntities()
    Dim i As Long, x As Long, curZOrder As Long

    ' don't render anything if we don't have any containers
    If WindowCount = 0 Then Exit Sub
    ' reset zOrder
    curZOrder = 1

    ' loop through windows
    Do While curZOrder <= WindowCount
        For i = 1 To WindowCount
            If curZOrder = Windows(i).Window.zOrder Then
                ' increment
                curZOrder = curZOrder + 1
                ' make sure it's visible
                If Windows(i).Window.visible Then
                    ' render container
                    RenderWindow i
                    ' render controls
                    For x = 1 To Windows(i).ControlCount
                        If Windows(i).Controls(x).visible Then RenderEntity i, x
                    Next
                End If
            End If
        Next
    Loop
End Sub

Public Sub RenderEntity(winNum As Long, entNum As Long)
    Dim xO As Long, yO As Long, hor_centre As Long, ver_centre As Long, height As Long, width As Long, left As Long, texNum As Long, xOffset As Long
    Dim callBack As Long, taddText As String, Colour As Long, textArray() As String, count As Long, yOffset As Long, i As Long, y As Long, x As Long

    ' check if the window exists
    If winNum <= 0 Or winNum > WindowCount Then
        Exit Sub
    End If

    ' check if the entity exists
    If entNum <= 0 Or entNum > Windows(winNum).ControlCount Then
        Exit Sub
    End If

    ' check the container's position
    xO = Windows(winNum).Window.left
    yO = Windows(winNum).Window.top

    With Windows(winNum).Controls(entNum)

        ' find the control type
        Select Case .Type
            ' picture box
            Case EntityTypes.entPictureBox
                ' render specific designs
                If .design(.state) > 0 Then RenderDesign .design(.state), .left + xO, .top + yO, .width, .height, .alpha
                ' render image
                If .image(.state) > 0 Then RenderTexture .image(.state), .left + xO, .top + yO, 0, 0, .width, .height, .width, .height, DX8Colour(White, .alpha)
                
            ' textbox
            Case EntityTypes.entTextBox
                ' render specific designs
                If .design(.state) > 0 Then RenderDesign .design(.state), .left + xO, .top + yO, .width, .height, .alpha
                ' render image
                If .image(.state) > 0 Then RenderTexture .image(.state), .left + xO, .top + yO, 0, 0, .width, .height, .width, .height, DX8Colour(White, .alpha)
                ' render text
                If activeWindow = winNum And Windows(winNum).activeControl = entNum Then taddText = chatShowLine
                ' if it's censored then render censored
                If Not .isCensor Then
                    RenderText font(.font), .text & taddText, .left + xO + .xOffset, .top + yO + .yOffset, .textColour
                Else
                    RenderText font(.font), CensorWord(.text) & taddText, .left + xO + .xOffset, .top + yO + .yOffset, .textColour
                End If

            ' buttons
            Case EntityTypes.entButton
                ' render specific designs
                If .design(.state) > 0 Then
                    If .design(.state) > 0 Then
                        RenderDesign .design(.state), .left + xO, .top + yO, .width, .height
                    End If
                End If
                ' render image
                If .image(.state) > 0 Then
                    If .image(.state) > 0 Then
                        RenderTexture .image(.state), .left + xO, .top + yO, 0, 0, .width, .height, .width, .height
                    End If
                End If
                ' render icon
                If .icon > 0 Then
                    width = mTexture(.icon).w
                    height = mTexture(.icon).h
                    RenderTexture .icon, .left + xO + .xOffset, .top + yO + .yOffset, 0, 0, width, height, width, height
                End If
                ' for changing the text space
                xOffset = width
                ' calculate the vertical centre
                height = TextHeight(font(Fonts.georgiaDec_16))
                If height > .height Then
                    ver_centre = .top + yO
                Else
                    ver_centre = .top + yO + ((.height - height) \ 2) + 1
                End If
                ' calculate the horizontal centre
                width = TextWidth(font(.font), .text)
                If width > .width Then
                    hor_centre = .left + xO + xOffset
                Else
                    hor_centre = .left + xO + xOffset + ((.width - width - xOffset) \ 2)
                End If
                ' get the colour
                If .state = Hover Then
                    Colour = .textColour_Hover
                ElseIf .state = MouseDown Then
                    Colour = .textColour_Click
                Else
                    Colour = .textColour
                End If
                RenderText font(.font), .text, hor_centre, ver_centre, Colour

            ' labels
            Case EntityTypes.entLabel
                If Len(.text) > 0 Then
                    Select Case .align
                        Case Alignment.alignLeft
                            ' check if need to word wrap
                            If TextWidth(font(.font), .text) > .width Then
                                ' wrap text
                                WordWrap_Array .text, .width, textArray()
                                ' render text
                                count = UBound(textArray)
                                For i = 1 To count
                                    RenderText font(.font), textArray(i), .left + xO, .top + yO + yOffset, .textColour, .alpha
                                    yOffset = yOffset + 14
                                Next
                            Else
                                ' just one line
                                RenderText font(.font), .text, .left + xO, .top + yO, .textColour, .alpha
                            End If
                        Case Alignment.alignRight
                            ' check if need to word wrap
                            If TextWidth(font(.font), .text) > .width Then
                                ' wrap text
                                WordWrap_Array .text, .width, textArray()
                                ' render text
                                count = UBound(textArray)
                                For i = 1 To count
                                    left = .left + .width - TextWidth(font(.font), textArray(i))
                                    RenderText font(.font), textArray(i), left + xO, .top + yO + yOffset, .textColour, .alpha
                                    yOffset = yOffset + 14
                                Next
                            Else
                                ' just one line
                                left = .left + .width - TextWidth(font(.font), .text)
                                RenderText font(.font), .text, left + xO, .top + yO, .textColour, .alpha
                            End If
                        Case Alignment.alignCentre
                            ' check if need to word wrap
                            If TextWidth(font(.font), .text) > .width Then
                                ' wrap text
                                WordWrap_Array .text, .width, textArray()
                                ' render text
                                count = UBound(textArray)
                                For i = 1 To count
                                    left = .left + (.width \ 2) - (TextWidth(font(.font), textArray(i)) \ 2)
                                    RenderText font(.font), textArray(i), left + xO, .top + yO + yOffset, .textColour, .alpha
                                    yOffset = yOffset + 14
                                Next
                            Else
                                ' just one line
                                left = .left + (.width \ 2) - (TextWidth(font(.font), .text) \ 2)
                                RenderText font(.font), .text, left + xO, .top + yO, .textColour, .alpha
                            End If
                    End Select
                End If

            ' checkboxes
            Case EntityTypes.entCheckbox
                
                Select Case .design(0)
                    Case DesignTypes.desChkNorm
                        ' empty?
                        If .value = 0 Then texNum = Tex_GUI(2) Else texNum = Tex_GUI(3)
                        ' render box
                        RenderTexture texNum, .left + xO, .top + yO, 0, 0, 14, 14, 14, 14
                        ' find text position
                        Select Case .align
                            Case Alignment.alignLeft
                                left = .left + 18 + xO
                            Case Alignment.alignRight
                                left = .left + 18 + (.width - 18) - TextWidth(font(.font), .text) + xO
                            Case Alignment.alignCentre
                                left = .left + 18 + ((.width - 18) / 2) - (TextWidth(font(.font), .text) / 2) + xO
                        End Select
                        ' render text
                        RenderText font(.font), .text, left, .top + yO, .textColour, .alpha
                    Case DesignTypes.desChkChat
                        If .value = 0 Then .alpha = 150 Else .alpha = 255
                        ' render box
                        RenderTexture Tex_GUI(51), .left + xO, .top + yO, 0, 0, 49, 23, 49, 23, DX8Colour(White, .alpha)
                        ' render text
                        left = .left + (49 / 2) - (TextWidth(font(.font), .text) / 2) + xO
                        ' render text
                        RenderText font(.font), .text, left, .top + yO + 4, .textColour, .alpha
                    Case DesignTypes.desChkCustom_Buying
                        If .value = 0 Then texNum = Tex_GUI(58) Else texNum = Tex_GUI(56)
                        RenderTexture texNum, .left + xO, .top + yO, 0, 0, 49, 20, 49, 20
                    Case DesignTypes.desChkCustom_Selling
                        If .value = 0 Then texNum = Tex_GUI(59) Else texNum = Tex_GUI(57)
                        RenderTexture texNum, .left + xO, .top + yO, 0, 0, 49, 20, 49, 20
                End Select
                
            ' comboboxes
            Case EntityTypes.entCombobox
                Select Case .design(0)
                    Case DesignTypes.desComboNorm
                        ' draw the background
                        RenderDesign DesignTypes.desTextBlack, .left + xO, .top + yO, .width, .height
                        ' render the text
                        If .value > 0 Then
                            If .value <= UBound(.list) Then
                                RenderText font(.font), .list(.value), .left + xO + 5, .top + yO + 3, White
                            End If
                        End If
                        ' draw the little arow
                        RenderTexture Tex_GUI(66), .left + xO + .width - 11, .top + yO + 7, 0, 0, 5, 4, 5, 4
                End Select
        End Select

        ' callback draw
        callBack = .onDraw

        If callBack <> 0 Then entCallBack callBack, winNum, entNum, 0, 0
    End With

End Sub

Public Sub RenderWindow(winNum As Long)
    Dim width As Long, height As Long, callBack As Long, x As Long, y As Long, i As Long, left As Long

    ' check if the window exists
    If winNum <= 0 Or winNum > WindowCount Then
        Exit Sub
    End If

    With Windows(winNum).Window
    
        Select Case .design(0)
            Case DesignTypes.desComboMenuNorm
                RenderTexture Tex_Blank, .left, .top, 0, 0, .width, .height, 1, 1, DX8Colour(Black, 157)
                ' text
                If UBound(.list) > 0 Then
                    y = .top + 2
                    x = .left
                    For i = 1 To UBound(.list)
                        ' render select
                        If i = .value Or i = .group Then RenderTexture Tex_Blank, x, y - 1, 0, 0, .width, 15, 1, 1, DX8Colour(Black, 255)
                        ' render text
                        left = x + (.width \ 2) - (TextWidth(font(.font), .list(i)) \ 2)
                        If i = .value Or i = .group Then
                            RenderText font(.font), .list(i), left, y, Yellow
                        Else
                            RenderText font(.font), .list(i), left, y, White
                        End If
                        y = y + 16
                    Next
                End If
                Exit Sub
        End Select

        Select Case .design(.state)

            Case DesignTypes.desWin_Black
                RenderTexture Tex_Fader, .left, .top, 0, 0, .width, .height, 32, 32, DX8Colour(Black, 190)
                
            Case DesignTypes.desWin_Norm
                ' render window
                RenderDesign DesignTypes.desWood, .left, .top, .width, .height
                RenderDesign DesignTypes.desGreen, .left + 2, .top + 2, .width - 4, 21
                ' render the icon
                width = mTexture(.icon).w
                height = mTexture(.icon).h
                RenderTexture .icon, .left + .xOffset, .top - (width - 18) + .yOffset, 0, 0, width, height, width, height
                ' render the caption
                RenderText font(.font), Trim$(.text), .left + height + 2, .top + 5, .textColour
                
            Case DesignTypes.desWin_NoBar
                ' render window
                RenderDesign DesignTypes.desWood, .left, .top, .width, .height
                
            Case DesignTypes.desWin_Empty
                ' render window
                RenderDesign DesignTypes.desWood_Empty, .left, .top, .width, .height
                RenderDesign DesignTypes.desGreen, .left + 2, .top + 2, .width - 4, 21
                ' render the icon
                width = mTexture(.icon).w
                height = mTexture(.icon).h
                RenderTexture .icon, .left + .xOffset, .top - (width - 18) + .yOffset, 0, 0, width, height, width, height
                ' render the caption
                RenderText font(.font), Trim$(.text), .left + height + 2, .top + 5, .textColour
                
            Case DesignTypes.desWin_Desc
                ' render window
                RenderDesign DesignTypes.desWin_Desc, .left, .top, .width, .height
                
            Case desWin_Shadow
                ' render window
                RenderDesign DesignTypes.desWin_Shadow, .left, .top, .width, .height
                
            Case desWin_Party
                ' render window
                RenderDesign DesignTypes.desWin_Party, .left, .top, .width, .height
        End Select

        ' OnDraw call back
        callBack = .onDraw

        If callBack <> 0 Then entCallBack callBack, winNum, 0, 0, 0
    End With

End Sub

Public Sub RenderDesign(design As Long, left As Long, top As Long, width As Long, height As Long, Optional alpha As Long = 255)
    Dim bs As Long, Colour As Long
    ' change colour for alpha
    Colour = DX8Colour(White, alpha)

    Select Case design
    
        Case DesignTypes.desMenuHeader
            ' render the header
            RenderTexture Tex_Blank, left, top, 0, 0, width, height, 32, 32, D3DColorARGB(200, 47, 77, 29)
            
        Case DesignTypes.desMenuOption
            ' render the option
            RenderTexture Tex_Blank, left, top, 0, 0, width, height, 32, 32, D3DColorARGB(200, 98, 98, 98)

        Case DesignTypes.desWood
            bs = 4
            ' render the wood box
            RenderEntity_Square Tex_Design(1), left, top, width, height, bs, alpha
            ' render wood texture
            RenderTexture Tex_GUI(1), left + bs, top + bs, 100, 100, width - (bs * 2), height - (bs * 2), width - (bs * 2), height - (bs * 2), Colour

        Case DesignTypes.desWood_Small
            bs = 2
            ' render the wood box
            RenderEntity_Square Tex_Design(8), left, top, width, height, bs, alpha
            ' render wood texture
            RenderTexture Tex_GUI(1), left + bs, top + bs, 100, 100, width - (bs * 2), height - (bs * 2), width - (bs * 2), height - (bs * 2), Colour
            
        Case DesignTypes.desWood_Empty
            bs = 4
            ' render the wood box
            RenderEntity_Square Tex_Design(9), left, top, width, height, bs, alpha

        Case DesignTypes.desGreen
            bs = 2
            ' render the green box
            RenderEntity_Square Tex_Design(2), left, top, width, height, bs, alpha
            ' render green gradient overlay
            RenderTexture Tex_Gradient(1), left + bs, top + bs, 0, 0, width - (bs * 2), height - (bs * 2), 128, 128, Colour

        Case DesignTypes.desGreen_Hover
            bs = 2
            ' render the green box
            RenderEntity_Square Tex_Design(2), left, top, width, height, bs, alpha
            ' render green gradient overlay
            RenderTexture Tex_Gradient(2), left + bs, top + bs, 0, 0, width - (bs * 2), height - (bs * 2), 128, 128, Colour

        Case DesignTypes.desGreen_Click
            bs = 2
            ' render the green box
            RenderEntity_Square Tex_Design(2), left, top, width, height, bs, alpha
            ' render green gradient overlay
            RenderTexture Tex_Gradient(3), left + bs, top + bs, 0, 0, width - (bs * 2), height - (bs * 2), 128, 128, Colour

        Case DesignTypes.desRed
            bs = 2
            ' render the red box
            RenderEntity_Square Tex_Design(3), left, top, width, height, bs, alpha
            ' render red gradient overlay
            RenderTexture Tex_Gradient(4), left + bs, top + bs, 0, 0, width - (bs * 2), height - (bs * 2), 128, 128, Colour

        Case DesignTypes.desRed_Hover
            bs = 2
            ' render the red box
            RenderEntity_Square Tex_Design(3), left, top, width, height, bs, alpha
            ' render red gradient overlay
            RenderTexture Tex_Gradient(5), left + bs, top + bs, 0, 0, width - (bs * 2), height - (bs * 2), 128, 128, Colour

        Case DesignTypes.desRed_Click
            bs = 2
            ' render the red box
            RenderEntity_Square Tex_Design(3), left, top, width, height, bs, alpha
            ' render red gradient overlay
            RenderTexture Tex_Gradient(6), left + bs, top + bs, 0, 0, width - (bs * 2), height - (bs * 2), 128, 128, Colour
            
        Case DesignTypes.desBlue
            bs = 2
            ' render the Blue box
            RenderEntity_Square Tex_Design(14), left, top, width, height, bs, alpha
            ' render Blue gradient overlay
            RenderTexture Tex_Gradient(8), left + bs, top + bs, 0, 0, width - (bs * 2), height - (bs * 2), 128, 128, Colour

        Case DesignTypes.desBlue_Hover
            bs = 2
            ' render the Blue box
            RenderEntity_Square Tex_Design(14), left, top, width, height, bs, alpha
            ' render Blue gradient overlay
            RenderTexture Tex_Gradient(9), left + bs, top + bs, 0, 0, width - (bs * 2), height - (bs * 2), 128, 128, Colour

        Case DesignTypes.desBlue_Click
            bs = 2
            ' render the Blue box
            RenderEntity_Square Tex_Design(14), left, top, width, height, bs, alpha
            ' render Blue gradient overlay
            RenderTexture Tex_Gradient(10), left + bs, top + bs, 0, 0, width - (bs * 2), height - (bs * 2), 128, 128, Colour
            
        Case DesignTypes.desOrange
            bs = 2
            ' render the Orange box
            RenderEntity_Square Tex_Design(15), left, top, width, height, bs, alpha
            ' render Orange gradient overlay
            RenderTexture Tex_Gradient(11), left + bs, top + bs, 0, 0, width - (bs * 2), height - (bs * 2), 128, 128, Colour

        Case DesignTypes.desOrange_Hover
            bs = 2
            ' render the Orange box
            RenderEntity_Square Tex_Design(15), left, top, width, height, bs, alpha
            ' render Orange gradient overlay
            RenderTexture Tex_Gradient(12), left + bs, top + bs, 0, 0, width - (bs * 2), height - (bs * 2), 128, 128, Colour

        Case DesignTypes.desOrange_Click
            bs = 2
            ' render the Orange box
            RenderEntity_Square Tex_Design(15), left, top, width, height, bs, alpha
            ' render Orange gradient overlay
            RenderTexture Tex_Gradient(13), left + bs, top + bs, 0, 0, width - (bs * 2), height - (bs * 2), 128, 128, Colour
            
        Case DesignTypes.desGrey
            bs = 2
            ' render the Orange box
            RenderEntity_Square Tex_Design(17), left, top, width, height, bs, alpha
            ' render Orange gradient overlay
            RenderTexture Tex_Gradient(14), left + bs, top + bs, 0, 0, width - (bs * 2), height - (bs * 2), 128, 128, Colour
    
        Case DesignTypes.desParchment
            bs = 20
            ' render the parchment box
            RenderEntity_Square Tex_Design(4), left, top, width, height, bs, alpha

        Case DesignTypes.desBlackOval
            bs = 4
            ' render the black oval
            RenderEntity_Square Tex_Design(5), left, top, width, height, bs, alpha

        Case DesignTypes.desTextBlack
            bs = 5
            ' render the black oval
            RenderEntity_Square Tex_Design(6), left, top, width, height, bs, alpha

        Case DesignTypes.desTextWhite
            bs = 5
            ' render the black oval
            RenderEntity_Square Tex_Design(7), left, top, width, height, bs, alpha
            
        Case DesignTypes.desTextBlack_Sq
            bs = 4
            ' render the black oval
            RenderEntity_Square Tex_Design(10), left, top, width, height, bs, alpha
            
        Case DesignTypes.desWin_Desc
            bs = 8
            ' render black square
            RenderEntity_Square Tex_Design(11), left, top, width, height, bs, alpha
            
        Case DesignTypes.desDescPic
            bs = 3
            ' render the green box
            RenderEntity_Square Tex_Design(12), left, top, width, height, bs, alpha
            ' render green gradient overlay
            RenderTexture Tex_Gradient(7), left + bs, top + bs, 0, 0, width - (bs * 2), height - (bs * 2), 128, 128, Colour
            
        Case DesignTypes.desWin_Shadow
            bs = 35
            ' render the green box
            RenderEntity_Square Tex_Design(13), left - bs, top - bs, width + (bs * 2), height + (bs * 2), bs, alpha

        Case DesignTypes.desWin_Party
            bs = 12
            ' render black square
            RenderEntity_Square Tex_Design(16), left, top, width, height, bs, alpha
            
        Case DesignTypes.desTileBox
            bs = 4
            ' render box
            RenderEntity_Square Tex_Design(18), left, top, width, height, bs, alpha
    End Select

End Sub

Public Sub RenderEntity_Square(texNum As Long, x As Long, y As Long, width As Long, height As Long, borderSize As Long, Optional alpha As Long = 255)
    Dim bs As Long, Colour As Long
    ' change colour for alpha
    Colour = DX8Colour(White, alpha)
    ' Set the border size
    bs = borderSize
    ' Draw centre
    RenderTexture texNum, x + bs, y + bs, bs + 1, bs + 1, width - (bs * 2), height - (bs * 2), 1, 1, Colour
    ' Draw top side
    RenderTexture texNum, x + bs, y, bs, 0, width - (bs * 2), bs, 1, bs, Colour
    ' Draw left side
    RenderTexture texNum, x, y + bs, 0, bs, bs, height - (bs * 2), bs, 1, Colour
    ' Draw right side
    RenderTexture texNum, x + width - bs, y + bs, bs + 3, bs, bs, height - (bs * 2), bs, 1, Colour
    ' Draw bottom side
    RenderTexture texNum, x + bs, y + height - bs, bs, bs + 3, width - (bs * 2), bs, 1, bs, Colour
    ' Draw top left corner
    RenderTexture texNum, x, y, 0, 0, bs, bs, bs, bs, Colour
    ' Draw top right corner
    RenderTexture texNum, x + width - bs, y, bs + 3, 0, bs, bs, bs, bs, Colour
    ' Draw bottom left corner
    RenderTexture texNum, x, y + height - bs, 0, bs + 3, bs, bs, bs, bs, Colour
    ' Draw bottom right corner
    RenderTexture texNum, x + width - bs, y + height - bs, bs + 3, bs + 3, bs, bs, bs, bs, Colour
End Sub

Sub Combobox_AddItem(winIndex As Long, controlIndex As Long, text As String)
Dim count As Long
    count = UBound(Windows(winIndex).Controls(controlIndex).list)
    ReDim Preserve Windows(winIndex).Controls(controlIndex).list(0 To count + 1)
    Windows(winIndex).Controls(controlIndex).list(count + 1) = text
End Sub

Public Sub CreateWindow(name As String, caption As String, zOrder As Long, left As Long, top As Long, width As Long, height As Long, icon As Long, _
   Optional visible As Boolean = True, Optional font As Long = Fonts.georgia_16, Optional textColour As Long = White, Optional xOffset As Long, _
   Optional yOffset As Long, Optional design_norm As Long, Optional design_hover As Long, Optional design_mousedown As Long, Optional image_norm As Long, _
   Optional image_hover As Long, Optional image_mousedown As Long, Optional entCallBack_norm As Long, Optional entCallBack_hover As Long, Optional entCallBack_mousedown As Long, _
   Optional entCallBack_mousemove As Long, Optional entCallBack_dblclick As Long, Optional canDrag As Boolean = True, Optional zChange As Byte = True, Optional ByVal onDraw As Long, _
   Optional isActive As Boolean, Optional clickThrough As Boolean)
   
    Dim i As Long
    Dim design(0 To entStates.state_Count - 1) As Long
    Dim image(0 To entStates.state_Count - 1) As Long
    Dim entCallBack(0 To entStates.state_Count - 1) As Long
    
    ' fill temp arrays
    design(entStates.Normal) = design_norm
    design(entStates.Hover) = design_hover
    design(entStates.MouseDown) = design_mousedown
    design(entStates.DblClick) = design_norm
    design(entStates.MouseUp) = design_norm
    image(entStates.Normal) = image_norm
    image(entStates.Hover) = image_hover
    image(entStates.MouseDown) = image_mousedown
    image(entStates.DblClick) = image_norm
    image(entStates.MouseUp) = image_norm
    entCallBack(entStates.Normal) = entCallBack_norm
    entCallBack(entStates.Hover) = entCallBack_hover
    entCallBack(entStates.MouseDown) = entCallBack_mousedown
    entCallBack(entStates.MouseMove) = entCallBack_mousemove
    entCallBack(entStates.DblClick) = entCallBack_dblclick
    ' redim the windows
    WindowCount = WindowCount + 1
    ReDim Preserve Windows(1 To WindowCount) As WindowRec

    ' set the properties
    With Windows(WindowCount).Window
        .name = name
        .Type = EntityTypes.entWindow

        ' loop through states
        For i = 0 To entStates.state_Count - 1
            .design(i) = design(i)
            .image(i) = image(i)
            .entCallBack(i) = entCallBack(i)
        Next

        .left = left
        .top = top
        .origLeft = left
        .origTop = top
        .width = width
        .height = height
        .visible = visible
        .canDrag = canDrag
        .text = caption
        .font = font
        .textColour = textColour
        .xOffset = xOffset
        .yOffset = yOffset
        .icon = icon
        .enabled = True
        .zChange = zChange
        .zOrder = zOrder
        .onDraw = onDraw
        .clickThrough = clickThrough
        ' set active
        If .visible Then activeWindow = WindowCount
    End With

    ' set the zOrder
    zOrder_Win = zOrder_Win + 1
End Sub

Public Sub CreateTextbox(winNum As Long, name As String, left As Long, top As Long, width As Long, height As Long, Optional text As String, Optional font As Long = Fonts.georgia_16, _
    Optional textColour As Long = White, Optional align As Byte = Alignment.alignLeft, Optional visible As Boolean = True, Optional alpha As Long = 255, Optional image_norm As Long, _
    Optional image_hover As Long, Optional image_mousedown As Long, Optional design_norm As Long, Optional design_hover As Long, Optional design_mousedown As Long, _
    Optional entCallBack_norm As Long, Optional entCallBack_hover As Long, Optional entCallBack_mousedown As Long, Optional entCallBack_mousemove As Long, Optional entCallBack_dblclick As Long, _
    Optional isActive As Boolean, Optional xOffset As Long, Optional yOffset As Long, Optional isCensor As Boolean, Optional entCallBack_enter As Long)
    Dim design(0 To entStates.state_Count - 1) As Long
    Dim image(0 To entStates.state_Count - 1) As Long
    Dim entCallBack(0 To entStates.state_Count - 1) As Long
    ' fill temp arrays
    design(entStates.Normal) = design_norm
    design(entStates.Hover) = design_hover
    design(entStates.MouseDown) = design_mousedown
    image(entStates.Normal) = image_norm
    image(entStates.Hover) = image_hover
    image(entStates.MouseDown) = image_mousedown
    entCallBack(entStates.Normal) = entCallBack_norm
    entCallBack(entStates.Hover) = entCallBack_hover
    entCallBack(entStates.MouseDown) = entCallBack_mousedown
    entCallBack(entStates.MouseMove) = entCallBack_mousemove
    entCallBack(entStates.DblClick) = entCallBack_dblclick
    entCallBack(entStates.Enter) = entCallBack_enter
    ' create the textbox
    CreateEntity winNum, zOrder_Con, name, entTextBox, design(), image(), entCallBack(), left, top, width, height, visible, , , , , text, align, font, textColour, alpha, , xOffset, yOffset, , , , isActive, isCensor
    End Sub

Public Sub CreatePictureBox(winNum As Long, name As String, left As Long, top As Long, width As Long, height As Long, Optional visible As Boolean = True, Optional canDrag As Boolean, _
   Optional alpha As Long = 255, Optional clickThrough As Boolean, Optional image_norm As Long, Optional image_hover As Long, Optional image_mousedown As Long, Optional design_norm As Long, _
   Optional design_hover As Long, Optional design_mousedown As Long, Optional entCallBack_norm As Long, Optional entCallBack_hover As Long, Optional entCallBack_mousedown As Long, _
   Optional entCallBack_mousemove As Long, Optional entCallBack_dblclick As Long, Optional onDraw As Long)
    Dim design(0 To entStates.state_Count - 1) As Long
    Dim image(0 To entStates.state_Count - 1) As Long
    Dim entCallBack(0 To entStates.state_Count - 1) As Long
    ' fill temp arrays
    design(entStates.Normal) = design_norm
    design(entStates.Hover) = design_hover
    design(entStates.MouseDown) = design_mousedown
    image(entStates.Normal) = image_norm
    image(entStates.Hover) = image_hover
    image(entStates.MouseDown) = image_mousedown
    entCallBack(entStates.Normal) = entCallBack_norm
    entCallBack(entStates.Hover) = entCallBack_hover
    entCallBack(entStates.MouseDown) = entCallBack_mousedown
    entCallBack(entStates.MouseMove) = entCallBack_mousemove
    entCallBack(entStates.DblClick) = entCallBack_dblclick
    ' create the box
    CreateEntity winNum, zOrder_Con, name, entPictureBox, design(), image(), entCallBack(), left, top, width, height, visible, canDrag, , , , , , , , alpha, clickThrough, , , , , onDraw
End Sub

Public Sub CreateButton(winNum As Long, name As String, left As Long, top As Long, width As Long, height As Long, Optional text As String, Optional font As Fonts = Fonts.georgia_16, _
   Optional textColour As Long = White, Optional icon As Long, Optional visible As Boolean = True, Optional alpha As Long = 255, Optional image_norm As Long, Optional image_hover As Long, _
   Optional image_mousedown As Long, Optional design_norm As Long, Optional design_hover As Long, Optional design_mousedown As Long, Optional entCallBack_norm As Long, _
   Optional entCallBack_hover As Long, Optional entCallBack_mousedown As Long, Optional entCallBack_mousemove As Long, Optional entCallBack_dblclick As Long, Optional xOffset As Long, _
   Optional yOffset As Long, Optional textColour_Hover As Long = -1, Optional textColour_Click As Long = -1, Optional tooltip As String)
    Dim design(0 To entStates.state_Count - 1) As Long
    Dim image(0 To entStates.state_Count - 1) As Long
    Dim entCallBack(0 To entStates.state_Count - 1) As Long
    ' default the colours
    If textColour_Hover = -1 Then textColour_Hover = textColour
    If textColour_Click = -1 Then textColour_Click = textColour
    ' fill temp arrays
    design(entStates.Normal) = design_norm
    design(entStates.Hover) = design_hover
    design(entStates.MouseDown) = design_mousedown
    image(entStates.Normal) = image_norm
    image(entStates.Hover) = image_hover
    image(entStates.MouseDown) = image_mousedown
    entCallBack(entStates.Normal) = entCallBack_norm
    entCallBack(entStates.Hover) = entCallBack_hover
    entCallBack(entStates.MouseDown) = entCallBack_mousedown
    entCallBack(entStates.MouseMove) = entCallBack_mousemove
    entCallBack(entStates.DblClick) = entCallBack_dblclick
    ' create the box
    CreateEntity winNum, zOrder_Con, name, entButton, design(), image(), entCallBack(), left, top, width, height, visible, , , , , text, , font, textColour, alpha, , xOffset, yOffset, , icon, , , , textColour_Hover, textColour_Click, tooltip
End Sub

Public Sub CreateLabel(winNum As Long, name As String, left As Long, top As Long, width As Long, Optional height As Long, Optional text As String, Optional font As Fonts = Fonts.georgia_16, _
   Optional textColour As Long = White, Optional align As Byte = Alignment.alignLeft, Optional visible As Boolean = True, Optional alpha As Long = 255, Optional clickThrough As Boolean, _
   Optional entCallBack_norm As Long, Optional entCallBack_hover As Long, Optional entCallBack_mousedown As Long, Optional entCallBack_mousemove As Long, Optional entCallBack_dblclick As Long)
    Dim design(0 To entStates.state_Count - 1) As Long
    Dim image(0 To entStates.state_Count - 1) As Long
    Dim entCallBack(0 To entStates.state_Count - 1) As Long
    ' fill temp arrays
    entCallBack(entStates.Normal) = entCallBack_norm
    entCallBack(entStates.Hover) = entCallBack_hover
    entCallBack(entStates.MouseDown) = entCallBack_mousedown
    entCallBack(entStates.MouseMove) = entCallBack_mousemove
    entCallBack(entStates.DblClick) = entCallBack_dblclick
    ' create the box
    CreateEntity winNum, zOrder_Con, name, entLabel, design(), image(), entCallBack(), left, top, width, height, visible, , , , , text, align, font, textColour, alpha, clickThrough
End Sub

Public Sub CreateCheckbox(winNum As Long, name As String, left As Long, top As Long, width As Long, Optional height As Long = 15, Optional value As Long, Optional text As String, _
    Optional font As Fonts = Fonts.georgia_16, Optional textColour As Long = White, Optional align As Byte = Alignment.alignLeft, Optional visible As Boolean = True, Optional alpha As Long = 255, _
    Optional theDesign As Long, Optional entCallBack_norm As Long, Optional entCallBack_hover As Long, Optional entCallBack_mousedown As Long, Optional entCallBack_mousemove As Long, _
    Optional entCallBack_dblclick As Long, Optional group As Long)
    Dim design(0 To entStates.state_Count - 1) As Long
    Dim image(0 To entStates.state_Count - 1) As Long
    Dim entCallBack(0 To entStates.state_Count - 1) As Long
    ' fill temp arrays
    entCallBack(entStates.Normal) = entCallBack_norm
    entCallBack(entStates.Hover) = entCallBack_hover
    entCallBack(entStates.MouseDown) = entCallBack_mousedown
    entCallBack(entStates.MouseMove) = entCallBack_mousemove
    entCallBack(entStates.DblClick) = entCallBack_dblclick
    ' fill temp array
    design(0) = theDesign
    ' create the box
    CreateEntity winNum, zOrder_Con, name, entCheckbox, design(), image(), entCallBack(), left, top, width, height, visible, , , , value, text, align, font, textColour, alpha, , , , , , , , , , , , group
End Sub

Public Sub CreateComboBox(winNum As Long, name As String, left As Long, top As Long, width As Long, height As Long, design As Long, Optional font As Fonts = Fonts.georgia_16)
    Dim theDesign(0 To entStates.state_Count - 1) As Long
    Dim image(0 To entStates.state_Count - 1) As Long
    Dim entCallBack(0 To entStates.state_Count - 1) As Long
    theDesign(0) = design
    ' create the box
    CreateEntity winNum, zOrder_Con, name, entCombobox, theDesign(), image(), entCallBack(), left, top, width, height, , , , , , , , font
End Sub

Public Function GetWindowIndex(winName As String) As Long
    Dim i As Long

    For i = 1 To WindowCount

        If LCase$(Windows(i).Window.name) = LCase$(winName) Then
            GetWindowIndex = i
            Exit Function
        End If

    Next

    GetWindowIndex = 0
End Function

Public Function GetControlIndex(winName As String, controlName As String) As Long
    Dim i As Long, winIndex As Long
    
    winIndex = GetWindowIndex(winName)

    If Not winIndex > 0 Or Not winIndex <= WindowCount Then Exit Function

    For i = 1 To Windows(winIndex).ControlCount

        If LCase$(Windows(winIndex).Controls(i).name) = LCase$(controlName) Then
            GetControlIndex = i
            Exit Function
        End If

    Next

    GetControlIndex = 0
End Function

Public Function SetActiveControl(curWindow As Long, curControl As Long) As Boolean
    ' make sure it's something which CAN be active
    Select Case Windows(curWindow).Controls(curControl).Type
        Case EntityTypes.entTextBox
            Windows(curWindow).activeControl = curControl
            SetActiveControl = True
    End Select
End Function

Public Sub CentraliseWindow(curWindow As Long)
    With Windows(curWindow).Window
        .left = (ScreenWidth / 2) - (.width / 2)
        .top = (ScreenHeight / 2) - (.height / 2)
        .origLeft = .left
        .origTop = .top
    End With
End Sub

Public Sub HideWindows()
Dim i As Long
    For i = 1 To WindowCount
        HideWindow i
    Next
End Sub

Public Sub ShowWindow(curWindow As Long, Optional forced As Boolean, Optional resetPosition As Boolean = True)
    Windows(curWindow).Window.visible = True
    If forced Then
        UpdateZOrder curWindow, forced
        activeWindow = curWindow
    ElseIf Windows(curWindow).Window.zChange Then
        UpdateZOrder curWindow
        activeWindow = curWindow
    End If
    If resetPosition Then
        With Windows(curWindow).Window
            .left = .origLeft
            .top = .origTop
        End With
    End If
End Sub

Public Sub HideWindow(curWindow As Long)
Dim i As Long
    Windows(curWindow).Window.visible = False
    ' find next window to set as active
    For i = WindowCount To 1 Step -1
        If Windows(i).Window.visible And Windows(i).Window.zChange Then
            'UpdateZOrder i
            activeWindow = i
            Exit Sub
        End If
    Next
End Sub

Public Sub CreateWindow_Login()
    ' Create the window
    CreateWindow "winLogin", "Login", zOrder_Win, 0, 0, 276, 182, Tex_Item(45), , Fonts.rockwellDec_15, , 3, 5, DesignTypes.desWin_Norm, DesignTypes.desWin_Norm, DesignTypes.desWin_Norm
    ' Centralise it
    CentraliseWindow WindowCount
    
    ' Set the index for spawning controls
    zOrder_Con = 1
    
    ' Close button
    CreateButton WindowCount, "btnClose", Windows(WindowCount).Window.width - 19, 6, 13, 13, , , , , , , Tex_GUI(8), Tex_GUI(9), Tex_GUI(10), , , , , , GetAddress(AddressOf DestroyGame)
    ' Parchment
    CreatePictureBox WindowCount, "picParchment", 6, 26, 264, 150, , , , , , , , DesignTypes.desParchment, DesignTypes.desParchment, DesignTypes.desParchment
    ' Shadows
    CreatePictureBox WindowCount, "picShadow_1", 67, 43, 142, 9, , , , , , , , DesignTypes.desBlackOval, DesignTypes.desBlackOval, DesignTypes.desBlackOval
    CreatePictureBox WindowCount, "picShadow_2", 67, 79, 142, 9, , , , , , , , DesignTypes.desBlackOval, DesignTypes.desBlackOval, DesignTypes.desBlackOval
    ' Buttons
    CreateButton WindowCount, "btnAccept", 68, 134, 67, 22, "Accept", rockwellDec_15, White, , , , , , , DesignTypes.desGreen, DesignTypes.desGreen_Hover, DesignTypes.desGreen_Click, , , GetAddress(AddressOf btnLogin_Click)
    CreateButton WindowCount, "btnExit", 142, 134, 67, 22, "Exit", rockwellDec_15, White, , , , , , , DesignTypes.desRed, DesignTypes.desRed_Hover, DesignTypes.desRed_Click, , , GetAddress(AddressOf DestroyGame)
    ' Labels
    CreateLabel WindowCount, "lblUsername", 66, 39, 142, , "Username", rockwellDec_15, White, Alignment.alignCentre
    CreateLabel WindowCount, "lblPassword", 66, 75, 142, , "Password", rockwellDec_15, White, Alignment.alignCentre
    ' Textboxes
    CreateTextbox WindowCount, "txtUser", 67, 55, 142, 19, Options.Username, Fonts.rockwell_15, , Alignment.alignLeft, , , , , , DesignTypes.desTextWhite, DesignTypes.desTextWhite, DesignTypes.desTextWhite, , , , , , , 5, 3
    CreateTextbox WindowCount, "txtPass", 67, 91, 142, 19, vbNullString, Fonts.rockwell_15, , Alignment.alignLeft, , , , , , DesignTypes.desTextWhite, DesignTypes.desTextWhite, DesignTypes.desTextWhite, , , , , , , 5, 3, True, GetAddress(AddressOf btnLogin_Click)
    ' Checkbox
    CreateCheckbox WindowCount, "chkSaveUser", 67, 114, 142, , Options.SaveUser, "Save Username?", rockwell_15, , , , , DesignTypes.desChkNorm, , , GetAddress(AddressOf chkSaveUser_Click)
    
    ' Set the active control
    If Not Len(Windows(GetWindowIndex("winLogin")).Controls(GetControlIndex("winLogin", "txtUser")).text) > 0 Then
        SetActiveControl GetWindowIndex("winLogin"), GetControlIndex("winLogin", "txtUser")
    Else
        SetActiveControl GetWindowIndex("winLogin"), GetControlIndex("winLogin", "txtPass")
    End If
End Sub

Public Sub CreateWindow_Characters()
    ' Create the window
    CreateWindow "winCharacters", "Characters", zOrder_Win, 0, 0, 364, 229, Tex_Item(62), False, Fonts.rockwellDec_15, , 3, 5, DesignTypes.desWin_Norm, DesignTypes.desWin_Norm, DesignTypes.desWin_Norm
    ' Centralise it
    CentraliseWindow WindowCount
    
    ' Set the index for spawning controls
    zOrder_Con = 1
    
    ' Close button
    CreateButton WindowCount, "btnClose", Windows(WindowCount).Window.width - 19, 6, 13, 13, , , , , , , Tex_GUI(8), Tex_GUI(9), Tex_GUI(10), , , , , , GetAddress(AddressOf btnCharacters_Close)
    ' Parchment
    CreatePictureBox WindowCount, "picParchment", 6, 26, 352, 197, , , , , , , , DesignTypes.desParchment, DesignTypes.desParchment, DesignTypes.desParchment
    ' Names
    CreatePictureBox WindowCount, "picShadow_1", 22, 41, 98, 9, , , , , , , , DesignTypes.desBlackOval, DesignTypes.desBlackOval, DesignTypes.desBlackOval
    CreateLabel WindowCount, "lblCharName_1", 22, 37, 98, , "Blank Slot", rockwellDec_15, White, Alignment.alignCentre
    CreatePictureBox WindowCount, "picShadow_2", 132, 41, 98, 9, , , , , , , , DesignTypes.desBlackOval, DesignTypes.desBlackOval, DesignTypes.desBlackOval
    CreateLabel WindowCount, "lblCharName_2", 132, 37, 98, , "Blank Slot", rockwellDec_15, White, Alignment.alignCentre
    CreatePictureBox WindowCount, "picShadow_3", 242, 41, 98, 9, , , , , , , , DesignTypes.desBlackOval, DesignTypes.desBlackOval, DesignTypes.desBlackOval
    CreateLabel WindowCount, "lblCharName_3", 242, 37, 98, , "Blank Slot", rockwellDec_15, White, Alignment.alignCentre
    ' Scenery Boxes
    CreatePictureBox WindowCount, "picScene_1", 23, 55, 96, 96, , , , , Tex_GUI(11), Tex_GUI(11), Tex_GUI(11)
    CreatePictureBox WindowCount, "picScene_2", 133, 55, 96, 96, , , , , Tex_GUI(11), Tex_GUI(11), Tex_GUI(11)
    CreatePictureBox WindowCount, "picScene_3", 243, 55, 96, 96, , , , , Tex_GUI(11), Tex_GUI(11), Tex_GUI(11), , , , , , , , , GetAddress(AddressOf Chars_DrawFace)
    ' Create Buttons
    CreateButton WindowCount, "btnSelectChar_1", 22, 155, 98, 24, "Select", rockwellDec_15, , , , , , , , DesignTypes.desGreen, DesignTypes.desGreen_Hover, DesignTypes.desGreen_Click, , , GetAddress(AddressOf btnAcceptChar_1)
    CreateButton WindowCount, "btnCreateChar_1", 22, 155, 98, 24, "Create", rockwellDec_15, , , , , , , , DesignTypes.desGreen, DesignTypes.desGreen_Hover, DesignTypes.desGreen_Click, , , GetAddress(AddressOf btnCreateChar_1)
    CreateButton WindowCount, "btnDelChar_1", 22, 183, 98, 24, "Delete", rockwellDec_15, , , , , , , , DesignTypes.desRed, DesignTypes.desRed_Hover, DesignTypes.desRed_Click, , , GetAddress(AddressOf btnDelChar_1)
    CreateButton WindowCount, "btnSelectChar_2", 132, 155, 98, 24, "Select", rockwellDec_15, , , , , , , , DesignTypes.desGreen, DesignTypes.desGreen_Hover, DesignTypes.desGreen_Click, , , GetAddress(AddressOf btnAcceptChar_2)
    CreateButton WindowCount, "btnCreateChar_2", 132, 155, 98, 24, "Create", rockwellDec_15, , , , , , , , DesignTypes.desGreen, DesignTypes.desGreen_Hover, DesignTypes.desGreen_Click, , , GetAddress(AddressOf btnCreateChar_2)
    CreateButton WindowCount, "btnDelChar_2", 132, 183, 98, 24, "Delete", rockwellDec_15, , , , , , , , DesignTypes.desRed, DesignTypes.desRed_Hover, DesignTypes.desRed_Click, , , GetAddress(AddressOf btnDelChar_2)
    CreateButton WindowCount, "btnSelectChar_3", 242, 155, 98, 24, "Select", rockwellDec_15, , , , , , , , DesignTypes.desGreen, DesignTypes.desGreen_Hover, DesignTypes.desGreen_Click, , , GetAddress(AddressOf btnAcceptChar_3)
    CreateButton WindowCount, "btnCreateChar_3", 242, 155, 98, 24, "Create", rockwellDec_15, , , , , , , , DesignTypes.desGreen, DesignTypes.desGreen_Hover, DesignTypes.desGreen_Click, , , GetAddress(AddressOf btnCreateChar_3)
    CreateButton WindowCount, "btnDelChar_3", 242, 183, 98, 24, "Delete", rockwellDec_15, , , , , , , , DesignTypes.desRed, DesignTypes.desRed_Hover, DesignTypes.desRed_Click, , , GetAddress(AddressOf btnDelChar_3)
End Sub

Public Sub CreateWindow_Loading()
    ' Create the window
    CreateWindow "winLoading", "Loading", zOrder_Win, 0, 0, 278, 79, Tex_Item(104), True, Fonts.rockwellDec_15, , 2, 7, DesignTypes.desWin_Norm, DesignTypes.desWin_Norm, DesignTypes.desWin_Norm
    ' Centralise it
    CentraliseWindow WindowCount
    
    ' Set the index for spawning controls
    zOrder_Con = 1
    
    ' Parchment
    CreatePictureBox WindowCount, "picParchment", 6, 26, 266, 47, , , , , , , , DesignTypes.desParchment, DesignTypes.desParchment, DesignTypes.desParchment
    ' Text background
    CreatePictureBox WindowCount, "picRecess", 26, 39, 226, 22, , , , , , , , DesignTypes.desTextBlack, DesignTypes.desTextBlack, DesignTypes.desTextBlack
    ' Label
    CreateLabel WindowCount, "lblLoading", 6, 43, 266, , "Loading Game Data...", rockwell_15, , Alignment.alignCentre
End Sub

Public Sub CreateWindow_Dialogue()
    ' Create black background
    CreateWindow "winBlank", "", zOrder_Win, 0, 0, 800, 600, 0, , , , , , DesignTypes.desWin_Black, DesignTypes.desWin_Black, DesignTypes.desWin_Black, , , , , , , , , False, False
    ' Create dialogue window
    CreateWindow "winDialogue", "Warning", zOrder_Win, 0, 0, 348, 145, Tex_Item(38), , Fonts.rockwellDec_15, , 3, 5, DesignTypes.desWin_Norm, DesignTypes.desWin_Norm, DesignTypes.desWin_Norm, , , , , , , , , , False
    ' Centralise it
    CentraliseWindow WindowCount
    
    ' Set the index for spawning controls
    zOrder_Con = 1
    
    ' Close button
    CreateButton WindowCount, "btnClose", Windows(WindowCount).Window.width - 19, 6, 13, 13, , , , , , , Tex_GUI(8), Tex_GUI(9), Tex_GUI(10), , , , , , GetAddress(AddressOf btnDialogue_Close)
    ' Parchment
    CreatePictureBox WindowCount, "picParchment", 6, 26, 335, 113, , , , , , , , DesignTypes.desParchment, DesignTypes.desParchment, DesignTypes.desParchment
    ' Header
    CreatePictureBox WindowCount, "picShadow", 103, 44, 144, 9, , , , , , , , DesignTypes.desBlackOval, DesignTypes.desBlackOval, DesignTypes.desBlackOval
    CreateLabel WindowCount, "lblHeader", 103, 41, 144, , "Header", rockwellDec_15, White, Alignment.alignCentre
    ' Labels
    CreateLabel WindowCount, "lblBody_1", 15, 60, 314, , "Invalid username or password.", rockwell_15, , Alignment.alignCentre
    CreateLabel WindowCount, "lblBody_2", 15, 75, 314, , "Please try again.", rockwell_15, , Alignment.alignCentre
    ' Buttons
    CreateButton WindowCount, "btnYes", 104, 98, 68, 24, "Yes", rockwellDec_15, , , False, , , , , DesignTypes.desGreen, DesignTypes.desGreen_Hover, DesignTypes.desGreen_Click, , , GetAddress(AddressOf Dialogue_Yes)
    CreateButton WindowCount, "btnNo", 180, 98, 68, 24, "No", rockwellDec_15, , , False, , , , , DesignTypes.desRed, DesignTypes.desRed_Hover, DesignTypes.desRed_Click, , , GetAddress(AddressOf Dialogue_No)
    CreateButton WindowCount, "btnOkay", 140, 98, 68, 24, "Okay", rockwellDec_15, , , , , , , , DesignTypes.desGreen, DesignTypes.desGreen_Hover, DesignTypes.desGreen_Click, , , GetAddress(AddressOf Dialogue_Okay)
    ' Input
    CreateTextbox WindowCount, "txtInput", 93, 75, 162, 18, , rockwell_15, White, Alignment.alignCentre, , , , , , DesignTypes.desTextBlack, DesignTypes.desTextBlack, DesignTypes.desTextBlack, , , , , , , 4, 2
    ' set active control
    SetActiveControl WindowCount, GetControlIndex("winDialogue", "txtInput")
End Sub

Public Sub CreateWindow_Classes()
    ' Create window
    CreateWindow "winClasses", "Select Class", zOrder_Win, 0, 0, 364, 229, Tex_Item(17), False, Fonts.rockwellDec_15, , 2, 6, DesignTypes.desWin_Norm, DesignTypes.desWin_Norm, DesignTypes.desWin_Norm
    ' Centralise it
    CentraliseWindow WindowCount
    
    ' Set the index for spawning controls
    zOrder_Con = 1
    
    ' Close button
    CreateButton WindowCount, "btnClose", Windows(WindowCount).Window.width - 19, 6, 13, 13, , , , , , , Tex_GUI(8), Tex_GUI(9), Tex_GUI(10), , , , , , GetAddress(AddressOf btnClasses_Close)
    ' Parchment
    CreatePictureBox WindowCount, "picParchment", 6, 26, 352, 197, , , , , , , , DesignTypes.desParchment, DesignTypes.desParchment, DesignTypes.desParchment, , , , , , GetAddress(AddressOf Classes_DrawFace)
    ' Class Name
    CreatePictureBox WindowCount, "picShadow", 183, 42, 98, 9, , , , , , , , DesignTypes.desBlackOval, DesignTypes.desBlackOval, DesignTypes.desBlackOval
    CreateLabel WindowCount, "lblClassName", 183, 39, 98, , "Warrior", rockwellDec_15, White, Alignment.alignCentre
    ' Select Buttons
    CreateButton WindowCount, "btnLeft", 171, 40, 11, 13, , , , , , , Tex_GUI(12), Tex_GUI(14), Tex_GUI(16), , , , , , GetAddress(AddressOf btnClasses_Left)
    CreateButton WindowCount, "btnRight", 282, 40, 11, 13, , , , , , , Tex_GUI(13), Tex_GUI(15), Tex_GUI(17), , , , , , GetAddress(AddressOf btnClasses_Right)
    ' Accept Button
    CreateButton WindowCount, "btnAccept", 183, 185, 98, 22, "Accept", rockwellDec_15, , , , , , , , DesignTypes.desGreen, DesignTypes.desGreen_Hover, DesignTypes.desGreen_Click, , , GetAddress(AddressOf btnClasses_Accept)
    ' Text background
    CreatePictureBox WindowCount, "picBackground", 127, 55, 210, 124, , , , , , , , DesignTypes.desTextBlack, DesignTypes.desTextBlack, DesignTypes.desTextBlack
    ' Overlay
    CreatePictureBox WindowCount, "picOverlay", 6, 26, 0, 0, , , , , , , , , , , , , , , , GetAddress(AddressOf Classes_DrawText)
End Sub

Public Sub CreateWindow_NewChar()
    ' Create window
    CreateWindow "winNewChar", "Create Character", zOrder_Win, 0, 0, 291, 172, Tex_Item(17), False, Fonts.rockwellDec_15, , 2, 6, DesignTypes.desWin_Norm, DesignTypes.desWin_Norm, DesignTypes.desWin_Norm
    ' Centralise it
    CentraliseWindow WindowCount
    
    ' Set the index for spawning controls
    zOrder_Con = 1
    
    ' Close button
    CreateButton WindowCount, "btnClose", Windows(WindowCount).Window.width - 19, 6, 13, 13, , , , , , , Tex_GUI(8), Tex_GUI(9), Tex_GUI(10), , , , , , GetAddress(AddressOf btnNewChar_Cancel)
    ' Parchment
    CreatePictureBox WindowCount, "picParchment", 6, 26, 278, 140, , , , , , , , DesignTypes.desParchment, DesignTypes.desParchment, DesignTypes.desParchment
    ' Name
    CreatePictureBox WindowCount, "picShadow_1", 29, 42, 124, 9, , , , , , , , DesignTypes.desBlackOval, DesignTypes.desBlackOval, DesignTypes.desBlackOval
    CreateLabel WindowCount, "lblName", 29, 39, 124, , "Name", rockwellDec_15, White, Alignment.alignCentre
    ' Textbox
    CreateTextbox WindowCount, "txtName", 29, 55, 124, 19, , Fonts.rockwell_15, , Alignment.alignLeft, , , , , , DesignTypes.desTextWhite, DesignTypes.desTextWhite, DesignTypes.desTextWhite, , , , , , , 5, 3
    ' Gender
    CreatePictureBox WindowCount, "picShadow_2", 29, 85, 124, 9, , , , , , , , DesignTypes.desBlackOval, DesignTypes.desBlackOval, DesignTypes.desBlackOval
    CreateLabel WindowCount, "lblGender", 29, 82, 124, , "Gender", rockwellDec_15, White, Alignment.alignCentre
    ' Checkboxes
    CreateCheckbox WindowCount, "chkMale", 29, 103, 55, , 1, "Male", rockwell_15, , Alignment.alignCentre, , , DesignTypes.desChkNorm, , , GetAddress(AddressOf chkNewChar_Male), , , 1
    CreateCheckbox WindowCount, "chkFemale", 90, 103, 62, , 0, "Female", rockwell_15, , Alignment.alignCentre, , , DesignTypes.desChkNorm, , , GetAddress(AddressOf chkNewChar_Female), , , 1
    ' Buttons
    CreateButton WindowCount, "btnAccept", 29, 127, 60, 24, "Accept", rockwellDec_15, , , , , , , , DesignTypes.desGreen, DesignTypes.desGreen_Hover, DesignTypes.desGreen_Click, , , GetAddress(AddressOf btnNewChar_Accept)
    CreateButton WindowCount, "btnCancel", 93, 127, 60, 24, "Cancel", rockwellDec_15, , , , , , , , DesignTypes.desRed, DesignTypes.desRed_Hover, DesignTypes.desRed_Click, , , GetAddress(AddressOf btnNewChar_Cancel)
    ' Sprite
    CreatePictureBox WindowCount, "picShadow_3", 175, 42, 76, 9, , , , , , , , DesignTypes.desBlackOval, DesignTypes.desBlackOval, DesignTypes.desBlackOval
    CreateLabel WindowCount, "lblSprite", 175, 39, 76, , "Sprite", rockwellDec_15, White, Alignment.alignCentre
    ' Scene
    CreatePictureBox WindowCount, "picScene", 165, 55, 96, 96, , , , , Tex_GUI(11), Tex_GUI(11), Tex_GUI(11), , , , , , , , , GetAddress(AddressOf NewChar_OnDraw)
    ' Buttons
    CreateButton WindowCount, "btnLeft", 163, 40, 11, 13, , , , , , , Tex_GUI(12), Tex_GUI(14), Tex_GUI(16), , , , , , GetAddress(AddressOf btnNewChar_Left)
    CreateButton WindowCount, "btnRight", 252, 40, 11, 13, , , , , , , Tex_GUI(13), Tex_GUI(15), Tex_GUI(17), , , , , , GetAddress(AddressOf btnNewChar_Right)
    
    ' Set the active control
    SetActiveControl GetWindowIndex("winNewChar"), GetControlIndex("winNewChar", "txtName")
End Sub

Public Sub CreateWindow_EscMenu()
    ' Create window
    CreateWindow "winEscMenu", "", zOrder_Win, 0, 0, 210, 156, 0, , , , , , DesignTypes.desWin_NoBar, DesignTypes.desWin_NoBar, DesignTypes.desWin_NoBar, , , , , , , , , False, False
    ' Centralise it
    CentraliseWindow WindowCount
    
    ' Set the index for spawning controls
    zOrder_Con = 1
    
    ' Parchment
    CreatePictureBox WindowCount, "picParchment", 6, 6, 198, 144, , , , , , , , DesignTypes.desParchment, DesignTypes.desParchment, DesignTypes.desParchment
    ' Buttons
    CreateButton WindowCount, "btnReturn", 16, 16, 178, 28, "Return to Game (Esc)", rockwellDec_15, , , , , , , , DesignTypes.desGreen, DesignTypes.desGreen_Hover, DesignTypes.desGreen_Click, , , GetAddress(AddressOf btnEscMenu_Return)
    CreateButton WindowCount, "btnOptions", 16, 48, 178, 28, "Options", rockwellDec_15, , , , , , , , DesignTypes.desOrange, DesignTypes.desOrange_Hover, DesignTypes.desOrange_Click, , , GetAddress(AddressOf btnEscMenu_Options)
    CreateButton WindowCount, "btnMainMenu", 16, 80, 178, 28, "Back to Main Menu", rockwellDec_15, , , , , , , , DesignTypes.desBlue, DesignTypes.desBlue_Hover, DesignTypes.desBlue_Click, , , GetAddress(AddressOf btnEscMenu_MainMenu)
    CreateButton WindowCount, "btnExit", 16, 112, 178, 28, "Exit the Game", rockwellDec_15, , , , , , , , DesignTypes.desRed, DesignTypes.desRed_Hover, DesignTypes.desRed_Click, , , GetAddress(AddressOf btnEscMenu_Exit)
End Sub

Public Sub CreateWindow_Bars()
    ' Create window
    CreateWindow "winBars", "", zOrder_Win, 10, 10, 239, 77, 0, , , , , , DesignTypes.desWin_NoBar, DesignTypes.desWin_NoBar, DesignTypes.desWin_NoBar, , , , , , , , , False, False
    
    ' Set the index for spawning controls
    zOrder_Con = 1
    
    ' Parchment
    CreatePictureBox WindowCount, "picParchment", 6, 6, 227, 65, , , , , , , , DesignTypes.desParchment, DesignTypes.desParchment, DesignTypes.desParchment
    ' Blank Bars
    CreatePictureBox WindowCount, "picHP_Blank", 15, 15, 209, 13, , , , , Tex_GUI(24), Tex_GUI(24), Tex_GUI(24)
    CreatePictureBox WindowCount, "picSP_Blank", 15, 32, 209, 13, , , , , Tex_GUI(25), Tex_GUI(25), Tex_GUI(25)
    CreatePictureBox WindowCount, "picEXP_Blank", 15, 49, 209, 13, , , , , Tex_GUI(26), Tex_GUI(26), Tex_GUI(26)
    ' Draw the bars
    CreatePictureBox WindowCount, "picBlank", 0, 0, 0, 0, , , , , , , , , , , , , , , , GetAddress(AddressOf Bars_OnDraw)
    ' Bar Labels
    CreatePictureBox WindowCount, "picHealth", 16, 11, 44, 14, , , , , Tex_GUI(21), Tex_GUI(21), Tex_GUI(21)
    CreatePictureBox WindowCount, "picSpirit", 16, 28, 44, 14, , , , , Tex_GUI(22), Tex_GUI(22), Tex_GUI(22)
    CreatePictureBox WindowCount, "picExperience", 16, 45, 74, 14, , , , , Tex_GUI(23), Tex_GUI(23), Tex_GUI(23)
    ' Labels
    CreateLabel WindowCount, "lblHP", 15, 14, 209, 12, "999/999", rockwellDec_10, White, Alignment.alignCentre
    CreateLabel WindowCount, "lblMP", 15, 31, 209, 12, "999/999", rockwellDec_10, White, Alignment.alignCentre
    CreateLabel WindowCount, "lblEXP", 15, 48, 209, 12, "999/999", rockwellDec_10, White, Alignment.alignCentre
End Sub

Public Sub CreateWindow_Menu()
    ' Create window
    CreateWindow "winMenu", "", zOrder_Win, 564, 563, 229, 31, 0, , , , , , , , , , , , , , , , , False, False
    
    ' Set the index for spawning controls
    zOrder_Con = 1
    
    ' Wood part
    CreatePictureBox WindowCount, "picWood", 0, 5, 228, 21, , , , , , , , DesignTypes.desWood, DesignTypes.desWood, DesignTypes.desWood
    ' Buttons
    CreateButton WindowCount, "btnChar", 8, 1, 29, 29, , , , Tex_Item(108), , , , , , DesignTypes.desGreen, DesignTypes.desGreen_Hover, DesignTypes.desGreen_Click, , , GetAddress(AddressOf btnMenu_Char), , , -1, -2, , , "Character (C)"
    CreateButton WindowCount, "btnInv", 44, 1, 29, 29, , , , Tex_Item(1), , , , , , DesignTypes.desGreen, DesignTypes.desGreen_Hover, DesignTypes.desGreen_Click, , , GetAddress(AddressOf btnMenu_Inv), , , -1, -2, , , "Inventory (I)"
    CreateButton WindowCount, "btnSkills", 82, 1, 29, 29, , , , Tex_Item(109), , , , , , DesignTypes.desGreen, DesignTypes.desGreen_Hover, DesignTypes.desGreen_Click, , , GetAddress(AddressOf btnMenu_Skills), , , -1, -2, , , "Skills (M)"
    'CreateButton WindowCount, "btnMap", 119, 1, 29, 29, , , , Tex_Item(106), , , , , , DesignTypes.desGreen, DesignTypes.desGreen_Hover, DesignTypes.desGreen_Click, , , GetAddress(AddressOf btnMenu_Map), , , -1, -2
    'CreateButton WindowCount, "btnGuild", 155, 1, 29, 29, , , , Tex_Item(107), , , , , , DesignTypes.desGreen, DesignTypes.desGreen_Hover, DesignTypes.desGreen_Click, , , GetAddress(AddressOf btnMenu_Guild), , , -1, -1
    'CreateButton WindowCount, "btnQuest", 191, 1, 29, 29, , , , Tex_Item(23), , , , , , DesignTypes.desGreen, DesignTypes.desGreen_Hover, DesignTypes.desGreen_Click, , , GetAddress(AddressOf btnMenu_Quest), , , -1, -2
    CreateButton WindowCount, "btnMap", 119, 1, 29, 29, , , , Tex_Item(106), , , , , , DesignTypes.desGrey, DesignTypes.desGrey, DesignTypes.desGrey, , , GetAddress(AddressOf btnMenu_Map), , , -1, -2
    CreateButton WindowCount, "btnGuild", 155, 1, 29, 29, , , , Tex_Item(107), , , , , , DesignTypes.desGrey, DesignTypes.desGrey, DesignTypes.desGrey, , , GetAddress(AddressOf btnMenu_Guild), , , -1, -1
    CreateButton WindowCount, "btnQuest", 191, 1, 29, 29, , , , Tex_Item(23), , , , , , DesignTypes.desGrey, DesignTypes.desGrey, DesignTypes.desGrey, , , GetAddress(AddressOf btnMenu_Quest), , , -1, -2
End Sub

Public Sub CreateWindow_Hotbar()
    ' Create window
    CreateWindow "winHotbar", "", zOrder_Win, 372, 10, 418, 36, 0, , , , , , , , , , , , , GetAddress(AddressOf Hotbar_MouseMove), GetAddress(AddressOf Hotbar_MouseDown), GetAddress(AddressOf Hotbar_MouseMove), GetAddress(AddressOf Hotbar_DblClick), False, False, GetAddress(AddressOf DrawHotbar)
End Sub

Public Sub CreateWindow_Inventory()
    ' Create window
    CreateWindow "winInventory", "Inventory", zOrder_Win, 0, 0, 202, 319, Tex_Item(1), False, Fonts.rockwellDec_15, , 2, 7, DesignTypes.desWin_Empty, DesignTypes.desWin_Empty, DesignTypes.desWin_Empty, , , , , GetAddress(AddressOf Inventory_MouseMove), GetAddress(AddressOf Inventory_MouseDown), GetAddress(AddressOf Inventory_MouseMove), GetAddress(AddressOf Inventory_DblClick), , , GetAddress(AddressOf DrawInventory)
    ' Centralise it
    CentraliseWindow WindowCount
    
    ' Set the index for spawning controls
    zOrder_Con = 1
    
    ' Close button
    CreateButton WindowCount, "btnClose", Windows(WindowCount).Window.width - 19, 6, 13, 13, , , , , , , Tex_GUI(8), Tex_GUI(9), Tex_GUI(10), , , , , , GetAddress(AddressOf btnMenu_Inv)
    ' Gold amount
    CreatePictureBox WindowCount, "picBlank", 8, 293, 186, 18, , , , , Tex_GUI(67), Tex_GUI(67), Tex_GUI(67)
    CreateLabel WindowCount, "lblGold", 42, 296, 100, , "0g", verdana_12
    ' Drop
    CreateButton WindowCount, "btnDrop", 155, 294, 38, 16, , , , Tex_GUI(36), , , , , , DesignTypes.desGreen, DesignTypes.desGreen_Hover, DesignTypes.desGreen_Click, , , , , , 5, 3
End Sub

Public Sub CreateWindow_Character()
    ' Create window
    CreateWindow "winCharacter", "Character Status", zOrder_Win, 0, 0, 174, 356, Tex_Item(62), False, Fonts.rockwellDec_15, , 2, 6, DesignTypes.desWin_Empty, DesignTypes.desWin_Empty, DesignTypes.desWin_Empty, , , , , GetAddress(AddressOf Character_MouseMove), GetAddress(AddressOf Character_MouseDown), GetAddress(AddressOf Character_MouseMove), GetAddress(AddressOf Character_MouseMove), , , GetAddress(AddressOf DrawCharacter)
    ' Centralise it
    CentraliseWindow WindowCount
    
    ' Set the index for spawning controls
    zOrder_Con = 1
    
    ' Close button
    CreateButton WindowCount, "btnClose", Windows(WindowCount).Window.width - 19, 6, 13, 13, , , , , , , Tex_GUI(8), Tex_GUI(9), Tex_GUI(10), , , , , , GetAddress(AddressOf btnMenu_Char)
    ' Parchment
    CreatePictureBox WindowCount, "picParchment", 6, 26, 162, 287, , , , , , , , DesignTypes.desParchment, DesignTypes.desParchment, DesignTypes.desParchment
    ' White boxes
    CreatePictureBox WindowCount, "picWhiteBox", 13, 34, 148, 19, , , , , , , , DesignTypes.desTextWhite, DesignTypes.desTextWhite, DesignTypes.desTextWhite
    CreatePictureBox WindowCount, "picWhiteBox", 13, 54, 148, 19, , , , , , , , DesignTypes.desTextWhite, DesignTypes.desTextWhite, DesignTypes.desTextWhite
    CreatePictureBox WindowCount, "picWhiteBox", 13, 74, 148, 19, , , , , , , , DesignTypes.desTextWhite, DesignTypes.desTextWhite, DesignTypes.desTextWhite
    CreatePictureBox WindowCount, "picWhiteBox", 13, 94, 148, 19, , , , , , , , DesignTypes.desTextWhite, DesignTypes.desTextWhite, DesignTypes.desTextWhite
    CreatePictureBox WindowCount, "picWhiteBox", 13, 114, 148, 19, , , , , , , , DesignTypes.desTextWhite, DesignTypes.desTextWhite, DesignTypes.desTextWhite
    CreatePictureBox WindowCount, "picWhiteBox", 13, 134, 148, 19, , , , , , , , DesignTypes.desTextWhite, DesignTypes.desTextWhite, DesignTypes.desTextWhite
    CreatePictureBox WindowCount, "picWhiteBox", 13, 154, 148, 19, , , , , , , , DesignTypes.desTextWhite, DesignTypes.desTextWhite, DesignTypes.desTextWhite
    ' Labels
    CreateLabel WindowCount, "lblName", 18, 36, 147, 16, "Name", rockwellDec_10
    CreateLabel WindowCount, "lblClass", 18, 56, 147, 16, "Class", rockwellDec_10
    CreateLabel WindowCount, "lblLevel", 18, 76, 147, 16, "Level", rockwellDec_10
    CreateLabel WindowCount, "lblGuild", 18, 96, 147, 16, "Guild", rockwellDec_10
    CreateLabel WindowCount, "lblHealth", 18, 116, 147, 16, "Health", rockwellDec_10
    CreateLabel WindowCount, "lblSpirit", 18, 136, 147, 16, "Spirit", rockwellDec_10
    CreateLabel WindowCount, "lblExperience", 18, 156, 147, 16, "Experience", rockwellDec_10
    ' Attributes
    CreatePictureBox WindowCount, "picShadow", 18, 176, 138, 9, , , , , , , , DesignTypes.desBlackOval, DesignTypes.desBlackOval, DesignTypes.desBlackOval
    CreateLabel WindowCount, "lblLabel", 18, 173, 138, , "Character Attributes", rockwellDec_15, , Alignment.alignCentre
    ' Black boxes
    CreatePictureBox WindowCount, "picBlackBox", 13, 186, 148, 19, , , , , , , , DesignTypes.desTextBlack, DesignTypes.desTextBlack, DesignTypes.desTextBlack
    CreatePictureBox WindowCount, "picBlackBox", 13, 206, 148, 19, , , , , , , , DesignTypes.desTextBlack, DesignTypes.desTextBlack, DesignTypes.desTextBlack
    CreatePictureBox WindowCount, "picBlackBox", 13, 226, 148, 19, , , , , , , , DesignTypes.desTextBlack, DesignTypes.desTextBlack, DesignTypes.desTextBlack
    CreatePictureBox WindowCount, "picBlackBox", 13, 246, 148, 19, , , , , , , , DesignTypes.desTextBlack, DesignTypes.desTextBlack, DesignTypes.desTextBlack
    CreatePictureBox WindowCount, "picBlackBox", 13, 266, 148, 19, , , , , , , , DesignTypes.desTextBlack, DesignTypes.desTextBlack, DesignTypes.desTextBlack
    CreatePictureBox WindowCount, "picBlackBox", 13, 286, 148, 19, , , , , , , , DesignTypes.desTextBlack, DesignTypes.desTextBlack, DesignTypes.desTextBlack
    ' Labels
    CreateLabel WindowCount, "lblLabel", 18, 188, 138, , "Strength", rockwellDec_10, Gold, Alignment.alignRight
    CreateLabel WindowCount, "lblLabel", 18, 208, 138, , "Endurance", rockwellDec_10, Gold, Alignment.alignRight
    CreateLabel WindowCount, "lblLabel", 18, 228, 138, , "Intelligence", rockwellDec_10, Gold, Alignment.alignRight
    CreateLabel WindowCount, "lblLabel", 18, 248, 138, , "Agility", rockwellDec_10, Gold, Alignment.alignRight
    CreateLabel WindowCount, "lblLabel", 18, 268, 138, , "Willpower", rockwellDec_10, Gold, Alignment.alignRight
    CreateLabel WindowCount, "lblLabel", 18, 288, 138, , "Unused Stat Points", rockwellDec_10, LightGreen, Alignment.alignRight
    ' Buttons
    CreateButton WindowCount, "btnStat_1", 15, 188, 15, 15, , , , , , , Tex_GUI(48), Tex_GUI(49), Tex_GUI(50), , , , , , GetAddress(AddressOf Character_SpendPoint1)
    CreateButton WindowCount, "btnStat_2", 15, 208, 15, 15, , , , , , , Tex_GUI(48), Tex_GUI(49), Tex_GUI(50), , , , , , GetAddress(AddressOf Character_SpendPoint2)
    CreateButton WindowCount, "btnStat_3", 15, 228, 15, 15, , , , , , , Tex_GUI(48), Tex_GUI(49), Tex_GUI(50), , , , , , GetAddress(AddressOf Character_SpendPoint3)
    CreateButton WindowCount, "btnStat_4", 15, 248, 15, 15, , , , , , , Tex_GUI(48), Tex_GUI(49), Tex_GUI(50), , , , , , GetAddress(AddressOf Character_SpendPoint4)
    CreateButton WindowCount, "btnStat_5", 15, 268, 15, 15, , , , , , , Tex_GUI(48), Tex_GUI(49), Tex_GUI(50), , , , , , GetAddress(AddressOf Character_SpendPoint5)
    ' fake buttons
    CreatePictureBox WindowCount, "btnGreyStat_1", 15, 188, 15, 15, , , , , Tex_GUI(47), Tex_GUI(47), Tex_GUI(47)
    CreatePictureBox WindowCount, "btnGreyStat_2", 15, 208, 15, 15, , , , , Tex_GUI(47), Tex_GUI(47), Tex_GUI(47)
    CreatePictureBox WindowCount, "btnGreyStat_3", 15, 228, 15, 15, , , , , Tex_GUI(47), Tex_GUI(47), Tex_GUI(47)
    CreatePictureBox WindowCount, "btnGreyStat_4", 15, 248, 15, 15, , , , , Tex_GUI(47), Tex_GUI(47), Tex_GUI(47)
    CreatePictureBox WindowCount, "btnGreyStat_5", 15, 268, 15, 15, , , , , Tex_GUI(47), Tex_GUI(47), Tex_GUI(47)
    ' Labels
    CreateLabel WindowCount, "lblStat_1", 32, 188, 100, , "255", rockwellDec_10
    CreateLabel WindowCount, "lblStat_2", 32, 208, 100, , "255", rockwellDec_10
    CreateLabel WindowCount, "lblStat_3", 32, 228, 100, , "255", rockwellDec_10
    CreateLabel WindowCount, "lblStat_4", 32, 248, 100, , "255", rockwellDec_10
    CreateLabel WindowCount, "lblStat_5", 32, 268, 100, , "255", rockwellDec_10
    CreateLabel WindowCount, "lblPoints", 18, 288, 100, , "255", rockwellDec_10
End Sub

Public Sub CreateWindow_Description()
    ' Create window
    CreateWindow "winDescription", "", zOrder_Win, 0, 0, 193, 142, 0, , , , , , DesignTypes.desWin_Desc, DesignTypes.desWin_Desc, DesignTypes.desWin_Desc, , , , , , , , , False
    
    ' Set the index for spawning controls
    zOrder_Con = 1
    
    ' Name
    CreateLabel WindowCount, "lblName", 8, 12, 177, , "(SB) Flame Sword", rockwellDec_15, BrightBlue, Alignment.alignCentre
    ' Sprite box
    CreatePictureBox WindowCount, "picSprite", 18, 32, 68, 68, , , , , , , , DesignTypes.desDescPic, DesignTypes.desDescPic, DesignTypes.desDescPic, , , , , , GetAddress(AddressOf Description_OnDraw)
    ' Sep
    CreatePictureBox WindowCount, "picSep", 96, 28, 1, 92, , , , , Tex_GUI(44), Tex_GUI(44), Tex_GUI(44)
    ' Requirements
    CreateLabel WindowCount, "lblClass", 5, 102, 92, , "Warrior", verdana_12, LightGreen, Alignment.alignCentre
    CreateLabel WindowCount, "lblLevel", 5, 114, 92, , "Level 20", verdana_12, BrightRed, Alignment.alignCentre
    ' Bar
    CreatePictureBox WindowCount, "picBar", 19, 114, 66, 12, False, , , , Tex_GUI(45), Tex_GUI(45), Tex_GUI(45)
End Sub

Public Sub CreateWindow_DragBox()
    ' Create window
    CreateWindow "winDragBox", "", zOrder_Win, 0, 0, 32, 32, 0, , , , , , , , , , , , GetAddress(AddressOf DragBox_Check), , , , , , , GetAddress(AddressOf DragBox_OnDraw)
    ' Need to set up unique mouseup event
    Windows(WindowCount).Window.entCallBack(entStates.MouseUp) = GetAddress(AddressOf DragBox_Check)
End Sub

Public Sub CreateWindow_Skills()
    ' Create window
    CreateWindow "winSkills", "Skills", zOrder_Win, 0, 0, 202, 297, Tex_Item(109), False, Fonts.rockwellDec_15, , 2, 7, DesignTypes.desWin_Empty, DesignTypes.desWin_Empty, DesignTypes.desWin_Empty, , , , , GetAddress(AddressOf Skills_MouseMove), GetAddress(AddressOf Skills_MouseDown), GetAddress(AddressOf Skills_MouseMove), GetAddress(AddressOf Skills_DblClick), , , GetAddress(AddressOf DrawSkills)
    ' Centralise it
    CentraliseWindow WindowCount
    
    ' Set the index for spawning controls
    zOrder_Con = 1
    
    ' Close button
    CreateButton WindowCount, "btnClose", Windows(WindowCount).Window.width - 19, 6, 13, 13, , , , , , , Tex_GUI(8), Tex_GUI(9), Tex_GUI(10), , , , , , GetAddress(AddressOf btnMenu_Skills)
End Sub

Public Sub CreateWindow_Chat()
    ' Create window
    CreateWindow "winChat", "", zOrder_Win, 8, 422, 352, 152, 0, False, , , , , , , , , , , , , , , , False
    
    ' Set the index for spawning controls
    zOrder_Con = 1
    
    ' Channel boxes
    CreateCheckbox WindowCount, "chkGame", 10, 2, 49, 23, 1, "Game", rockwellDec_10, , , , , DesignTypes.desChkChat, , , GetAddress(AddressOf chkChat_Game)
    CreateCheckbox WindowCount, "chkMap", 60, 2, 49, 23, 1, "Map", rockwellDec_10, , , , , DesignTypes.desChkChat, , , GetAddress(AddressOf chkChat_Map)
    CreateCheckbox WindowCount, "chkGlobal", 110, 2, 49, 23, 1, "Global", rockwellDec_10, , , , , DesignTypes.desChkChat, , , GetAddress(AddressOf chkChat_Global)
    CreateCheckbox WindowCount, "chkParty", 160, 2, 49, 23, 1, "Party", rockwellDec_10, , , , , DesignTypes.desChkChat, , , GetAddress(AddressOf chkChat_Party)
    CreateCheckbox WindowCount, "chkGuild", 210, 2, 49, 23, 1, "Guild", rockwellDec_10, , , , , DesignTypes.desChkChat, , , GetAddress(AddressOf chkChat_Guild)
    CreateCheckbox WindowCount, "chkPrivate", 260, 2, 49, 23, 1, "Private", rockwellDec_10, , , , , DesignTypes.desChkChat, , , GetAddress(AddressOf chkChat_Private)
    ' Blank picturebox - ondraw wrapper
    CreatePictureBox WindowCount, "picNull", 0, 0, 0, 0, , , , , , , , , , , , , , , , GetAddress(AddressOf OnDraw_Chat)
    ' Chat button
    CreateButton WindowCount, "btnChat", 296, 124 + 16, 48, 20, "Say", rockwellDec_15, , , , , , , , DesignTypes.desGreen, DesignTypes.desGreen_Hover, DesignTypes.desGreen_Click, , , GetAddress(AddressOf btnSay_Click)
    ' Chat Textbox
    CreateTextbox WindowCount, "txtChat", 12, 127 + 16, 286, 25, , Fonts.verdana_12
    ' buttons
    CreateButton WindowCount, "btnUp", 328, 28, 11, 13, , , , , , , Tex_GUI(4), Tex_GUI(52), Tex_GUI(4), , , , , , GetAddress(AddressOf btnChat_Up)
    CreateButton WindowCount, "btnDown", 327, 122, 11, 13, , , , , , , Tex_GUI(5), Tex_GUI(53), Tex_GUI(5), , , , , , GetAddress(AddressOf btnChat_Down)
    
    ' Custom Handlers for mouse up
    Windows(WindowCount).Controls(GetControlIndex("winChat", "btnUp")).entCallBack(entStates.MouseUp) = GetAddress(AddressOf btnChat_Up_MouseUp)
    Windows(WindowCount).Controls(GetControlIndex("winChat", "btnDown")).entCallBack(entStates.MouseUp) = GetAddress(AddressOf btnChat_Down_MouseUp)

    ' Set the active control
    SetActiveControl GetWindowIndex("winChat"), GetControlIndex("winChat", "txtChat")
    
    ' sort out the tabs
    With Windows(GetWindowIndex("winChat"))
        .Controls(GetControlIndex("winChat", "chkGame")).value = Options.channelState(ChatChannel.chGame)
        .Controls(GetControlIndex("winChat", "chkMap")).value = Options.channelState(ChatChannel.chMap)
        .Controls(GetControlIndex("winChat", "chkGlobal")).value = Options.channelState(ChatChannel.chGlobal)
        .Controls(GetControlIndex("winChat", "chkParty")).value = Options.channelState(ChatChannel.chParty)
        .Controls(GetControlIndex("winChat", "chkGuild")).value = Options.channelState(ChatChannel.chGuild)
        .Controls(GetControlIndex("winChat", "chkPrivate")).value = Options.channelState(ChatChannel.chPrivate)
    End With
End Sub

Public Sub CreateWindow_ChatSmall()
    ' Create window
    CreateWindow "winChatSmall", "", zOrder_Win, 8, 438, 0, 0, 0, False, , , , , , , , , , , , , , , , False, , GetAddress(AddressOf OnDraw_ChatSmall), , True
    
    ' Set the index for spawning controls
    zOrder_Con = 1
    
    ' Chat Label
    CreateLabel WindowCount, "lblMsg", 12, 127, 286, 25, "Press 'Enter' to open chatbox.", verdana_12, Grey
End Sub

Public Sub CreateWindow_Options()
    ' Create window
    CreateWindow "winOptions", "", zOrder_Win, 0, 0, 210, 212, 0, , , , , , DesignTypes.desWin_NoBar, DesignTypes.desWin_NoBar, DesignTypes.desWin_NoBar, , , , , , , , , False, False
    ' Centralise it
    CentraliseWindow WindowCount
    
    ' Set the index for spawning controls
    zOrder_Con = 1
    
    ' Parchment
    CreatePictureBox WindowCount, "picParchment", 6, 6, 198, 200, , , , , , , , DesignTypes.desParchment, DesignTypes.desParchment, DesignTypes.desParchment
    ' General
    CreatePictureBox WindowCount, "picBlank", 35, 25, 140, 9, , , , , , , , DesignTypes.desBlackOval, DesignTypes.desBlackOval, DesignTypes.desBlackOval
    CreateLabel WindowCount, "lblBlank", 35, 22, 140, , "General Options", rockwellDec_15, White, Alignment.alignCentre
    ' Check boxes
    CreateCheckbox WindowCount, "chkMusic", 35, 40, 80, , , "Music", rockwellDec_10, , , , , DesignTypes.desChkNorm
    CreateCheckbox WindowCount, "chkSound", 115, 40, 80, , , "Sound", rockwellDec_10, , , , , DesignTypes.desChkNorm
    CreateCheckbox WindowCount, "chkAutotiles", 35, 60, 80, , , "Autotiles", rockwellDec_10, , , , , DesignTypes.desChkNorm
    CreateCheckbox WindowCount, "chkFullscreen", 115, 60, 80, , , "Fullscreen", rockwellDec_10, , , , , DesignTypes.desChkNorm
    ' Resolution
    CreatePictureBox WindowCount, "picBlank", 35, 85, 140, 9, , , , , , , , DesignTypes.desBlackOval, DesignTypes.desBlackOval, DesignTypes.desBlackOval
    CreateLabel WindowCount, "lblBlank", 35, 82, 140, , "Select Resolution", rockwellDec_15, White, Alignment.alignCentre
    ' combobox
    CreateComboBox WindowCount, "cmbRes", 30, 100, 150, 18, DesignTypes.desComboNorm, verdana_12
    ' Renderer
    CreatePictureBox WindowCount, "picBlank", 35, 125, 140, 9, , , , , , , , DesignTypes.desBlackOval, DesignTypes.desBlackOval, DesignTypes.desBlackOval
    CreateLabel WindowCount, "lblBlank", 35, 122, 140, , "DirectX Mode", rockwellDec_15, White, Alignment.alignCentre
    ' Check boxes
    CreateComboBox WindowCount, "cmbRender", 30, 140, 150, 18, DesignTypes.desComboNorm, verdana_12
    ' Button
    CreateButton WindowCount, "btnConfirm", 65, 168, 80, 22, "Confirm", rockwellDec_15, , , , , , , , DesignTypes.desGreen, DesignTypes.desGreen_Hover, DesignTypes.desGreen_Click, , , GetAddress(AddressOf btnOptions_Confirm)
    
    ' Populate the options screen
    SetOptionsScreen
End Sub

Public Sub CreateWindow_Shop()
    ' Create window
    CreateWindow "winShop", "Shop", zOrder_Win, 0, 0, 278, 293, Tex_Item(17), False, Fonts.rockwellDec_15, , 2, 5, DesignTypes.desWin_Empty, DesignTypes.desWin_Empty, DesignTypes.desWin_Empty, , , , , GetAddress(AddressOf Shop_MouseMove), GetAddress(AddressOf Shop_MouseDown), GetAddress(AddressOf Shop_MouseMove), GetAddress(AddressOf Shop_MouseMove), , , GetAddress(AddressOf DrawShopBackground)
    ' additional mouse event
    Windows(WindowCount).Window.entCallBack(entStates.MouseUp) = GetAddress(AddressOf Shop_MouseMove)
    ' Centralise it
    CentraliseWindow WindowCount
    
    ' Close button
    CreateButton WindowCount, "btnClose", Windows(WindowCount).Window.width - 19, 6, 13, 13, , , , , , , Tex_GUI(8), Tex_GUI(9), Tex_GUI(10), , , , , , GetAddress(AddressOf btnShop_Close)
    ' Parchment
    CreatePictureBox WindowCount, "picParchment", 6, 215, 266, 50, , , , , , , , DesignTypes.desParchment, DesignTypes.desParchment, DesignTypes.desParchment, , , , , , GetAddress(AddressOf DrawShop)
    ' Picture Box
    CreatePictureBox WindowCount, "picItemBG", 13, 222, 36, 36, , , , , Tex_GUI(54), Tex_GUI(54), Tex_GUI(54)
    CreatePictureBox WindowCount, "picItem", 15, 224, 32, 32
    ' Buttons
    CreateButton WindowCount, "btnBuy", 190, 228, 70, 24, "Buy", rockwellDec_15, White, , , , , , , DesignTypes.desGreen, DesignTypes.desGreen_Hover, DesignTypes.desGreen_Click, , , GetAddress(AddressOf btnShopBuy)
    CreateButton WindowCount, "btnSell", 190, 228, 70, 24, "Sell", rockwellDec_15, White, , False, , , , , DesignTypes.desRed, DesignTypes.desRed_Hover, DesignTypes.desRed_Click, , , GetAddress(AddressOf btnShopSell)
    ' Gold
    CreatePictureBox WindowCount, "picBlank", 9, 266, 162, 18, , , , , Tex_GUI(55), Tex_GUI(55), Tex_GUI(55)
    ' Buying/Selling
    CreateCheckbox WindowCount, "chkBuying", 173, 265, 49, 20, 1, , , , , , , DesignTypes.desChkCustom_Buying, , , GetAddress(AddressOf chkShopBuying)
    CreateCheckbox WindowCount, "chkSelling", 222, 265, 49, 20, 0, , , , , , , DesignTypes.desChkCustom_Selling, , , GetAddress(AddressOf chkShopSelling)
    ' Labels
    CreateLabel WindowCount, "lblName", 56, 226, 300, , "Test Item", verdanaBold_12, Black, Alignment.alignLeft
    CreateLabel WindowCount, "lblCost", 56, 240, 300, , "1000g", verdana_12, Black, Alignment.alignLeft
    ' Gold
    CreateLabel WindowCount, "lblGold", 44, 269, 300, , "0g", verdana_12
End Sub

Public Sub CreateWindow_NpcChat()
    ' Create window
    CreateWindow "winNpcChat", "Conversation with [Name]", zOrder_Win, 0, 0, 480, 228, Tex_Item(111), False, Fonts.rockwellDec_15, , 2, 11, DesignTypes.desWin_Norm, DesignTypes.desWin_Norm, DesignTypes.desWin_Norm
    ' Centralise it
    CentraliseWindow WindowCount
    
    ' Close Button
    CreateButton WindowCount, "btnClose", Windows(WindowCount).Window.width - 19, 6, 13, 13, , , , , , , Tex_GUI(8), Tex_GUI(9), Tex_GUI(10), , , , , , GetAddress(AddressOf btnNpcChat_Close)
    ' Parchment
    CreatePictureBox WindowCount, "picParchment", 6, 26, 468, 198, , , , , , , , DesignTypes.desParchment, DesignTypes.desParchment, DesignTypes.desParchment
    ' Face background
    CreatePictureBox WindowCount, "picFaceBG", 20, 40, 102, 102, , , , , Tex_GUI(60), Tex_GUI(60), Tex_GUI(60)
    ' Actual Face
    CreatePictureBox WindowCount, "picFace", 23, 43, 96, 96, , , , , Tex_Face(1), Tex_Face(1), Tex_Face(1)
    ' Chat BG
    CreatePictureBox WindowCount, "picChatBG", 128, 39, 334, 104, , , , , , , , DesignTypes.desTextBlack, DesignTypes.desTextBlack, DesignTypes.desTextBlack
    ' Chat
    CreateLabel WindowCount, "lblChat", 136, 44, 318, 102, "[Text]", rockwellDec_15, White, Alignment.alignCentre
    ' Reply buttons
    CreateButton WindowCount, "btnOpt4", 69, 145, 343, 15, "[Text]", verdana_12, Black, , , , , , , , , , , , GetAddress(AddressOf btnOpt4), , , , , DarkGrey
    CreateButton WindowCount, "btnOpt3", 69, 162, 343, 15, "[Text]", verdana_12, Black, , , , , , , , , , , , GetAddress(AddressOf btnOpt3), , , , , DarkGrey
    CreateButton WindowCount, "btnOpt2", 69, 179, 343, 15, "[Text]", verdana_12, Black, , , , , , , , , , , , GetAddress(AddressOf btnOpt2), , , , , DarkGrey
    CreateButton WindowCount, "btnOpt1", 69, 196, 343, 15, "[Text]", verdana_12, Black, , , , , , , , , , , , GetAddress(AddressOf btnOpt1), , , , , DarkGrey
    
    ' Cache positions
    optPos(1) = 196
    optPos(2) = 179
    optPos(3) = 162
    optPos(4) = 145
    optHeight = 228
End Sub

Public Sub CreateWindow_RightClick()
    ' Create window
    CreateWindow "winRightClickBG", "", zOrder_Win, 0, 0, 800, 600, 0, , , , , , , , , , , , , , GetAddress(AddressOf RightClick_Close), , , False
    ' Centralise it
    CentraliseWindow WindowCount
End Sub

Public Sub CreateWindow_PlayerMenu()
    ' Create window
    CreateWindow "winPlayerMenu", "", zOrder_Win, 0, 0, 110, 106, 0, , , , , , DesignTypes.desWin_Desc, DesignTypes.desWin_Desc, DesignTypes.desWin_Desc, , , , , , GetAddress(AddressOf RightClick_Close), , , False
    ' Centralise it
    CentraliseWindow WindowCount
    
    ' Name
    CreateButton WindowCount, "btnName", 8, 8, 94, 18, "[Name]", verdanaBold_12, White, , , , , , , DesignTypes.desMenuHeader, DesignTypes.desMenuHeader, DesignTypes.desMenuHeader, , , GetAddress(AddressOf RightClick_Close)
    ' Options
    CreateButton WindowCount, "btnParty", 8, 26, 94, 18, "Invite to Party", verdana_12, White, , , , , , , , DesignTypes.desMenuOption, , , , GetAddress(AddressOf PlayerMenu_Party)
    CreateButton WindowCount, "btnTrade", 8, 44, 94, 18, "Request Trade", verdana_12, White, , , , , , , , DesignTypes.desMenuOption, , , , GetAddress(AddressOf PlayerMenu_Trade)
    CreateButton WindowCount, "btnGuild", 8, 62, 94, 18, "Invite to Guild", verdana_12, White, , , , , , , , DesignTypes.desMenuOption, , , , GetAddress(AddressOf PlayerMenu_Guild)
    CreateButton WindowCount, "btnPM", 8, 80, 94, 18, "Private Message", verdana_12, White, , , , , , , , DesignTypes.desMenuOption, , , , GetAddress(AddressOf PlayerMenu_PM)
End Sub

Public Sub CreateWindow_Party()
    ' Create window
    CreateWindow "winParty", "", zOrder_Win, 4, 78, 252, 158, 0, , , , , , DesignTypes.desWin_Party, DesignTypes.desWin_Party, DesignTypes.desWin_Party, , , , , , , , , False
    
    ' Name labels
    CreateLabel WindowCount, "lblName1", 60, 20, 173, , "Richard - Level 10", rockwellDec_10
    CreateLabel WindowCount, "lblName2", 60, 60, 173, , "Anna - Level 18", rockwellDec_10
    CreateLabel WindowCount, "lblName3", 60, 100, 173, , "Doleo - Level 25", rockwellDec_10
    ' Empty Bars - HP
    CreatePictureBox WindowCount, "picEmptyBar_HP1", 58, 34, 173, 9, , , , , Tex_GUI(62), Tex_GUI(62), Tex_GUI(62)
    CreatePictureBox WindowCount, "picEmptyBar_HP2", 58, 74, 173, 9, , , , , Tex_GUI(62), Tex_GUI(62), Tex_GUI(62)
    CreatePictureBox WindowCount, "picEmptyBar_HP3", 58, 114, 173, 9, , , , , Tex_GUI(62), Tex_GUI(62), Tex_GUI(62)
    ' Empty Bars - SP
    CreatePictureBox WindowCount, "picEmptyBar_SP1", 58, 44, 173, 9, , , , , Tex_GUI(63), Tex_GUI(63), Tex_GUI(63)
    CreatePictureBox WindowCount, "picEmptyBar_SP2", 58, 84, 173, 9, , , , , Tex_GUI(63), Tex_GUI(63), Tex_GUI(63)
    CreatePictureBox WindowCount, "picEmptyBar_SP3", 58, 124, 173, 9, , , , , Tex_GUI(63), Tex_GUI(63), Tex_GUI(63)
    ' Filled bars - HP
    CreatePictureBox WindowCount, "picBar_HP1", 58, 34, 173, 9, , , , , Tex_GUI(64), Tex_GUI(64), Tex_GUI(64)
    CreatePictureBox WindowCount, "picBar_HP2", 58, 74, 173, 9, , , , , Tex_GUI(64), Tex_GUI(64), Tex_GUI(64)
    CreatePictureBox WindowCount, "picBar_HP3", 58, 114, 173, 9, , , , , Tex_GUI(64), Tex_GUI(64), Tex_GUI(64)
    ' Filled bars - SP
    CreatePictureBox WindowCount, "picBar_SP1", 58, 44, 173, 9, , , , , Tex_GUI(65), Tex_GUI(65), Tex_GUI(65)
    CreatePictureBox WindowCount, "picBar_SP2", 58, 84, 173, 9, , , , , Tex_GUI(65), Tex_GUI(65), Tex_GUI(65)
    CreatePictureBox WindowCount, "picBar_SP3", 58, 124, 173, 9, , , , , Tex_GUI(65), Tex_GUI(65), Tex_GUI(65)
    ' Shadows
    CreatePictureBox WindowCount, "picShadow1", 20, 24, 32, 32, , , , , Tex_Shadow, Tex_Shadow, Tex_Shadow
    CreatePictureBox WindowCount, "picShadow2", 20, 64, 32, 32, , , , , Tex_Shadow, Tex_Shadow, Tex_Shadow
    CreatePictureBox WindowCount, "picShadow3", 20, 104, 32, 32, , , , , Tex_Shadow, Tex_Shadow, Tex_Shadow
    ' Characters
    CreatePictureBox WindowCount, "picChar1", 20, 20, 32, 32, , , , , Tex_Char(1), Tex_Char(1), Tex_Char(1)
    CreatePictureBox WindowCount, "picChar2", 20, 60, 32, 32, , , , , Tex_Char(1), Tex_Char(1), Tex_Char(1)
    CreatePictureBox WindowCount, "picChar3", 20, 100, 32, 32, , , , , Tex_Char(1), Tex_Char(1), Tex_Char(1)
End Sub

Public Sub CreateWindow_Invitations()
    ' Create window
    CreateWindow "winInvite_Party", "", zOrder_Win, ScreenWidth - 234, ScreenHeight - 80, 223, 37, 0, , , , , , DesignTypes.desWin_Desc, DesignTypes.desWin_Desc, DesignTypes.desWin_Desc, , , , , , , , , False
    ' Button
    CreateButton WindowCount, "btnInvite", 11, 12, 201, 14, ColourChar & White & "Richard " & ColourChar & "-1" & "has invited you to a party.", verdana_12, Grey, , , , , , , , , , , , GetAddress(AddressOf btnInvite_Party), , , , , Green
    
    ' Create window
    CreateWindow "winInvite_Trade", "", zOrder_Win, ScreenWidth - 234, ScreenHeight - 80, 223, 37, 0, , , , , , DesignTypes.desWin_Desc, DesignTypes.desWin_Desc, DesignTypes.desWin_Desc, , , , , , , , , False
    ' Button
    CreateButton WindowCount, "btnInvite", 11, 12, 201, 14, ColourChar & White & "Richard " & ColourChar & "-1" & "has invited you to a party.", verdana_12, Grey, , , , , , , , , , , , GetAddress(AddressOf btnInvite_Trade), , , , , Green
End Sub

Public Sub CreateWindow_Trade()
    ' Create window
    CreateWindow "winTrade", "Trading with [Name]", zOrder_Win, 0, 0, 412, 386, Tex_Item(112), False, Fonts.rockwellDec_15, , 2, 5, DesignTypes.desWin_Empty, DesignTypes.desWin_Empty, DesignTypes.desWin_Empty, , , , , , , , , , , GetAddress(AddressOf DrawTrade)
    ' Centralise it
    CentraliseWindow WindowCount
    
    ' Close Button
    CreateButton WindowCount, "btnClose", Windows(WindowCount).Window.width - 19, 6, 13, 13, , , , , , , Tex_GUI(8), Tex_GUI(9), Tex_GUI(10), , , , , , GetAddress(AddressOf btnTrade_Close)
    ' Parchment
    CreatePictureBox WindowCount, "picParchment", 10, 312, 392, 66, , , , , , , , DesignTypes.desParchment, DesignTypes.desParchment, DesignTypes.desParchment
    ' Labels
    CreatePictureBox WindowCount, "picShadow", 36, 30, 142, 9, , , , , , , , DesignTypes.desBlackOval, DesignTypes.desBlackOval, DesignTypes.desBlackOval
    CreateLabel WindowCount, "lblYourTrade", 36, 27, 142, 9, "Robin's Offer", rockwellDec_15, White, Alignment.alignCentre
    CreatePictureBox WindowCount, "picShadow", 36 + 200, 30, 142, 9, , , , , , , , DesignTypes.desBlackOval, DesignTypes.desBlackOval, DesignTypes.desBlackOval
    CreateLabel WindowCount, "lblTheirTrade", 36 + 200, 27, 142, 9, "Richard's Offer", rockwellDec_15, White, Alignment.alignCentre
    ' Buttons
    CreateButton WindowCount, "btnAccept", 134, 340, 68, 24, "Accept", rockwellDec_15, White, , , , , , , DesignTypes.desGreen, DesignTypes.desGreen_Hover, DesignTypes.desGreen_Click, , , GetAddress(AddressOf btnTrade_Accept)
    CreateButton WindowCount, "btnDecline", 210, 340, 68, 24, "Decline", rockwellDec_15, White, , , , , , , DesignTypes.desRed, DesignTypes.desRed_Hover, DesignTypes.desRed_Click, , , GetAddress(AddressOf btnTrade_Close)
    ' Labels
    CreateLabel WindowCount, "lblStatus", 114, 322, 184, , "", verdanaBold_12, Red, Alignment.alignCentre
    ' Amounts
    CreateLabel WindowCount, "lblBlank", 25, 330, 100, , "Total Value", verdanaBold_12, Black, Alignment.alignCentre
    CreateLabel WindowCount, "lblBlank", 285, 330, 100, , "Total Value", verdanaBold_12, Black, Alignment.alignCentre
    CreateLabel WindowCount, "lblYourValue", 25, 344, 100, , "52,812g", verdana_12, Black, Alignment.alignCentre
    CreateLabel WindowCount, "lblTheirValue", 285, 344, 100, , "12,531g", verdana_12, Black, Alignment.alignCentre
    ' Item Containers
    CreatePictureBox WindowCount, "picYour", 14, 46, 184, 260, , , , , , , , , , , , GetAddress(AddressOf TradeMouseMove_Your), GetAddress(AddressOf TradeMouseDown_Your), GetAddress(AddressOf TradeMouseMove_Your), , GetAddress(AddressOf DrawYourTrade)
    CreatePictureBox WindowCount, "picTheir", 214, 46, 184, 260, , , , , , , , , , , , GetAddress(AddressOf TradeMouseMove_Their), GetAddress(AddressOf TradeMouseMove_Their), GetAddress(AddressOf TradeMouseMove_Their), , GetAddress(AddressOf DrawTheirTrade)
End Sub

Public Sub CreateWindow_Combobox()
    ' background window
    CreateWindow "winComboMenuBG", "ComboMenuBG", zOrder_Win, 0, 0, 800, 600, 0, , , , , , , , , , , , , , GetAddress(AddressOf CloseComboMenu), , , False, False
    
    ' window
    CreateWindow "winComboMenu", "ComboMenu", zOrder_Win, 0, 0, 100, 100, 0, , Fonts.verdana_12, , , , DesignTypes.desComboMenuNorm, , , , , , , , , , , False, False
    ' centralise it
    CentraliseWindow WindowCount
End Sub

Public Sub CreateWindow_Guild()
    ' Create window
    CreateWindow "winGuild", "Guild", zOrder_Win, 0, 0, 174, 320, Tex_Item(107), False, Fonts.rockwellDec_15, , 2, 6, DesignTypes.desWin_Norm, DesignTypes.desWin_Norm, DesignTypes.desWin_Norm
    ' Centralise it
    CentraliseWindow WindowCount
    
    ' Set the index for spawning controls
    zOrder_Con = 1
    
    ' Close button
    CreateButton WindowCount, "btnClose", Windows(WindowCount).Window.width - 19, 6, 13, 13, , , , , , , Tex_GUI(8), Tex_GUI(9), Tex_GUI(10), , , , , , GetAddress(AddressOf btnMenu_Guild)
    ' Parchment
    CreatePictureBox WindowCount, "picParchment", 6, 26, 162, 287, , , , , , , , DesignTypes.desParchment, DesignTypes.desParchment, DesignTypes.desParchment
    ' Attributes
    CreatePictureBox WindowCount, "picShadow", 18, 38, 138, 9, , , , , , , , DesignTypes.desBlackOval, DesignTypes.desBlackOval, DesignTypes.desBlackOval
    CreateLabel WindowCount, "lblGuild", 18, 35, 138, , "Guild Name", rockwellDec_15, , Alignment.alignCentre
    ' White boxes
    CreatePictureBox WindowCount, "picWhiteBox", 13, 51, 148, 19, , , , , , , , DesignTypes.desTextWhite, DesignTypes.desTextWhite, DesignTypes.desTextWhite
    CreatePictureBox WindowCount, "picWhiteBox", 13, 71, 148, 19, , , , , , , , DesignTypes.desTextWhite, DesignTypes.desTextWhite, DesignTypes.desTextWhite
    CreatePictureBox WindowCount, "picWhiteBox", 13, 91, 148, 19, , , , , , , , DesignTypes.desTextWhite, DesignTypes.desTextWhite, DesignTypes.desTextWhite
    CreatePictureBox WindowCount, "picWhiteBox", 13, 111, 148, 19, , , , , , , , DesignTypes.desTextWhite, DesignTypes.desTextWhite, DesignTypes.desTextWhite
    ' Labels
    CreateLabel WindowCount, "lblRank", 18, 53, 147, 16, "Guild Rank: None", rockwellDec_10
    CreateLabel WindowCount, "lblKills", 18, 73, 147, 16, "Enemy Kills: 0", rockwellDec_10
    CreateLabel WindowCount, "lblGold", 18, 93, 147, 16, "Bank Gold: 0g", rockwellDec_10
    CreateLabel WindowCount, "lblMembers", 18, 113, 147, 16, "Guild Members: 0", rockwellDec_10
End Sub

' Rendering & Initialisation
Public Sub InitGUI()

    ' Starter values
    zOrder_Win = 1
    zOrder_Con = 1
    
    ' Menu
    CreateWindow_Login
    CreateWindow_Characters
    CreateWindow_Loading
    CreateWindow_Dialogue
    CreateWindow_Classes
    CreateWindow_NewChar
    
    ' Game
    CreateWindow_Combobox
    CreateWindow_EscMenu
    CreateWindow_Bars
    CreateWindow_Menu
    CreateWindow_Hotbar
    CreateWindow_Inventory
    CreateWindow_Character
    CreateWindow_Description
    CreateWindow_DragBox
    CreateWindow_Skills
    CreateWindow_Chat
    CreateWindow_ChatSmall
    CreateWindow_Options
    CreateWindow_Shop
    CreateWindow_NpcChat
    CreateWindow_Party
    CreateWindow_Invitations
    CreateWindow_Trade
    CreateWindow_Guild
    
    ' Menus
    CreateWindow_RightClick
    CreateWindow_PlayerMenu
End Sub
