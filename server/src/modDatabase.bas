Attribute VB_Name = "modDatabase"
Option Explicit

' Text API
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long

'For Clear functions
Private Declare Sub ZeroMemory Lib "Kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

Private crcTable(0 To 255) As Long

Public Sub InitCRC32()
Dim i As Long, n As Long, CRC As Long

    For i = 0 To 255
        CRC = i
        For n = 0 To 7
            If CRC And 1 Then
                CRC = (((CRC And &HFFFFFFFE) \ 2) And &H7FFFFFFF) Xor &HEDB88320
            Else
                CRC = ((CRC And &HFFFFFFFE) \ 2) And &H7FFFFFFF
            End If
        Next
        crcTable(i) = CRC
    Next
End Sub

Public Function CRC32(ByRef Data() As Byte) As Long
Dim lCurPos As Long
Dim lLen As Long

    lLen = AryCount(Data) - 1
    CRC32 = &HFFFFFFFF
    
    For lCurPos = 0 To lLen
        CRC32 = (((CRC32 And &HFFFFFF00) \ &H100) And &HFFFFFF) Xor (crcTable((CRC32 And 255) Xor Data(lCurPos)))
    Next
    
    CRC32 = CRC32 Xor &HFFFFFFFF
End Function

Public Sub HandleError(ByVal procName As String, ByVal contName As String, ByVal erNumber, ByVal erDesc, ByVal erSource, ByVal erHelpContext)
Dim filename As String
    filename = App.Path & "\data files\logs\errors.txt"
    Open filename For Append As #1
        Print #1, "The following error occured at '" & procName & "' in '" & contName & "'."
        Print #1, "Run-time error '" & erNumber & "': " & erDesc & "."
        Print #1, ""
    Close #1
End Sub

Public Sub ChkDir(ByVal tDir As String, ByVal tName As String)
    If LCase$(dir(tDir & tName, vbDirectory)) <> tName Then Call MkDir(tDir & tName)
End Sub

' Outputs string to text file
Sub AddLog(ByVal Text As String, ByVal FN As String)
    Dim filename As String
    Dim f As Long

    If ServerLog Then
        filename = App.Path & "\data\logs\" & FN

        If Not FileExist(filename, True) Then
            f = FreeFile
            Open filename For Output As #f
            Close #f
        End If

        f = FreeFile
        Open filename For Append As #f
        Print #f, DateValue(Now) & " " & Time & ": " & Text
        Close #f
    End If

End Sub

' gets a string from a text file
Public Function GetVar(File As String, Header As String, Var As String) As String
    Dim sSpaces As String   ' Max string length
    Dim szReturn As String  ' Return default value if not found
    szReturn = vbNullString
    sSpaces = Space$(5000)
    Call GetPrivateProfileString$(Header, Var, szReturn, sSpaces, Len(sSpaces), File)
    GetVar = RTrim$(sSpaces)
    GetVar = left$(GetVar, Len(GetVar) - 1)
End Function

' writes a variable to a text file
Public Sub PutVar(File As String, Header As String, Var As String, Value As String)
    Call WritePrivateProfileString$(Header, Var, Value, File)
End Sub

Public Function FileExist(ByVal filename As String, Optional RAW As Boolean = False) As Boolean

    If Not RAW Then
        If LenB(dir(App.Path & "\" & filename)) > 0 Then
            FileExist = True
        End If

    Else

        If LenB(dir(filename)) > 0 Then
            FileExist = True
        End If
    End If

End Function

Public Sub SaveOptions()
    PutVar App.Path & "\data\options.ini", "OPTIONS", "MOTD", Options.MOTD
End Sub

Public Sub LoadOptions()
    Options.MOTD = GetVar(App.Path & "\data\options.ini", "OPTIONS", "MOTD")
End Sub

Public Sub ToggleMute(ByVal index As Long)
    ' exit out for rte9
    If index <= 0 Or index > MAX_PLAYERS Then Exit Sub

    ' toggle the player's mute
    If Player(index).isMuted = 1 Then
        Player(index).isMuted = 0
        ' Let them know
        PlayerMsg index, "You have been unmuted and can now talk in global.", BrightGreen
        TextAdd GetPlayerName(index) & " has been unmuted."
    Else
        Player(index).isMuted = 1
        ' Let them know
        PlayerMsg index, "You have been muted and can no longer talk in global.", BrightRed
        TextAdd GetPlayerName(index) & " has been muted."
    End If
    
    ' save the player
    SavePlayer index
End Sub

Public Sub BanIndex(ByVal BanPlayerIndex As Long)
Dim filename As String, IP As String, f As Long, i As Long

    ' Add banned to the player's index
    Player(BanPlayerIndex).isBanned = 1
    SavePlayer BanPlayerIndex

    ' IP banning
    filename = App.Path & "\data\banlist_ip.txt"

    ' Make sure the file exists
    If Not FileExist(filename, True) Then
        f = FreeFile
        Open filename For Output As #f
        Close #f
    End If

    ' Print the IP in the ip ban list
    IP = GetPlayerIP(BanPlayerIndex)
    f = FreeFile
    Open filename For Append As #f
        Print #f, IP
    Close #f
    
    ' Tell them they're banned
    Call GlobalMsg(GetPlayerName(BanPlayerIndex) & " has been banned from " & GAME_NAME & ".", White)
    Call AddLog(GetPlayerName(BanPlayerIndex) & " has been banned.", ADMIN_LOG)
    Call AlertMsg(BanPlayerIndex, DIALOGUE_MSG_BANNED)
End Sub

Public Function isBanned_IP(ByVal IP As String) As Boolean
Dim filename As String, fIP As String, f As Long
    
    filename = App.Path & "\data\banlist_ip.txt"

    ' Check if file exists
    If Not FileExist(filename, True) Then
        f = FreeFile
        Open filename For Output As #f
        Close #f
    End If

    f = FreeFile
    Open filename For Input As #f

    Do While Not EOF(f)
        Input #f, fIP

        ' Is banned?
        If Trim$(LCase$(fIP)) = Trim$(LCase$(Mid$(IP, 1, Len(fIP)))) Then
            isBanned_IP = True
            Close #f
            Exit Function
        End If
    Loop

    Close #f
End Function

Public Function isBanned_Account(ByVal index As Long) As Boolean
    If Player(index).isBanned = 1 Then
        isBanned_Account = True
    Else
        isBanned_Account = False
    End If
End Function

' **************
' ** Accounts **
' **************
Function AccountExist(ByVal Name As String) As Boolean
    Dim filename As String
    filename = App.Path & "\data\accounts\" & SanitiseString(Trim(Name)) & ".ini"

    If FileExist(filename, True) Then
        AccountExist = True
    End If

End Function

Function PasswordOK(ByVal Name As String, ByVal password As String) As Boolean
Dim filename As String
Dim RightPassword As String

    If AccountExist(Name) Then
        filename = App.Path & "\data\accounts\" & SanitiseString(Trim$(Name)) & ".ini"
        
        RightPassword = GetVar(filename, "ACCOUNT", "Password")

        If UCase$(Trim$(password)) = UCase$(Trim$(RightPassword)) Then
            PasswordOK = True
        End If
    End If

End Function

Sub AddAccount(ByVal index As Long, ByVal Name As String)
    Dim i As Long
    
    ClearPlayer index
    
    Player(index).Login = Name

    For i = 1 To MAX_CHARS
        Player(index).charNum = i
        Call SavePlayer(index)
    Next
End Sub

Sub DeleteName(ByVal Name As String)
    Dim f1 As Long
    Dim f2 As Long
    Dim s As String
    Call FileCopy(App.Path & "\data\accounts\_charlist.txt", App.Path & "\data\accounts\_chartemp.txt")
    ' Destroy name from charlist
    f1 = FreeFile
    Open App.Path & "\data\accounts\_chartemp.txt" For Input As #f1
    f2 = FreeFile
    Open App.Path & "\data\accounts\_charlist.txt" For Output As #f2

    Do While Not EOF(f1)
        Input #f1, s

        If Trim$(LCase$(s)) <> Trim$(LCase$(Name)) Then
            Print #f2, s
        End If

    Loop

    Close #f1
    Close #f2
    Call Kill(App.Path & "\data\accounts\_chartemp.txt")
End Sub

' ****************
' ** Characters **
' ****************
Function CharExist(ByVal index As Long, ByVal charNum As Long) As Boolean
Dim theName As String
    theName = GetVar(App.Path & "\data\accounts\" & SanitiseString(Trim$(Player(index).Login)) & ".ini", "CHAR" & charNum, "Name")
    'If LenB(Trim$(Player(index).Name)) > 0 Then
    If LenB(theName) > 0 Then
        CharExist = True
    End If
End Function

Sub AddChar(ByVal index As Long, ByVal Name As String, ByVal Sex As Byte, ByVal ClassNum As Long, ByVal Sprite As Long, ByVal charNum As Long)
    Dim f As Long
    Dim n As Long
    Dim spritecheck As Boolean

    If LenB(Trim$(Player(index).Name)) = 0 Then
        
        spritecheck = False
        
        If charNum < 1 Or charNum > MAX_CHARS Then Exit Sub
        Player(index).charNum = charNum
        
        Player(index).Name = Name
        Player(index).Sex = Sex
        Player(index).Class = ClassNum
        
        If Player(index).Sex = SEX_MALE Then
            Player(index).Sprite = Class(ClassNum).MaleSprite(Sprite)
        Else
            Player(index).Sprite = Class(ClassNum).FemaleSprite(Sprite)
        End If

        Player(index).Level = 1

        For n = 1 To Stats.Stat_Count - 1
            Player(index).Stat(n) = Class(ClassNum).Stat(n)
        Next n

        Player(index).dir = DIR_DOWN
        Player(index).Map = START_MAP
        Player(index).x = START_X
        Player(index).y = START_Y
        Player(index).dir = DIR_DOWN
        Player(index).Vital(Vitals.HP) = GetPlayerMaxVital(index, Vitals.HP)
        Player(index).Vital(Vitals.MP) = GetPlayerMaxVital(index, Vitals.MP)
        
        ' set starter equipment
        If Class(ClassNum).startItemCount > 0 Then
            For n = 1 To Class(ClassNum).startItemCount
                If Class(ClassNum).StartItem(n) > 0 Then
                    ' item exist?
                    If Len(Trim$(Item(Class(ClassNum).StartItem(n)).Name)) > 0 Then
                        Player(index).Inv(n).Num = Class(ClassNum).StartItem(n)
                        Player(index).Inv(n).Value = Class(ClassNum).StartValue(n)
                    End If
                End If
            Next
        End If
        
        ' set start spells
        If Class(ClassNum).startSpellCount > 0 Then
            For n = 1 To Class(ClassNum).startSpellCount
                If Class(ClassNum).StartSpell(n) > 0 Then
                    ' spell exist?
                    If Len(Trim$(Spell(Class(ClassNum).StartItem(n)).Name)) > 0 Then
                        Player(index).Spell(n).Spell = Class(ClassNum).StartSpell(n)
                        Player(index).Hotbar(n).Slot = Class(ClassNum).StartSpell(n)
                        Player(index).Hotbar(n).sType = 2 ' spells
                    End If
                End If
            Next
        End If
        
        ' Append name to file
        f = FreeFile
        Open App.Path & "\data\accounts\_charlist.txt" For Append As #f
        Print #f, Name
        Close #f
        Call SavePlayer(index)
        Exit Sub
    End If

End Sub

Function FindChar(ByVal Name As String) As Boolean
    Dim f As Long
    Dim s As String
    f = FreeFile
    Open App.Path & "\data\accounts\_charlist.txt" For Input As #f

    Do While Not EOF(f)
        Input #f, s

        If Trim$(LCase$(s)) = Trim$(LCase$(Name)) Then
            FindChar = True
            Close #f
            Exit Function
        End If

    Loop

    Close #f
End Function

' *************
' ** Players **
' *************
Sub SaveAllPlayersOnline()
    Dim i As Long

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            Call SavePlayer(i)
        End If
    Next
End Sub

Sub SavePlayer(ByVal index As Long)
Dim filename As String, i As Long, charHeader As String

    If index <= 0 Or index > MAX_PLAYERS Then Exit Sub
    
    ' the file
    filename = App.Path & "\data\accounts\" & SanitiseString(Trim$(Player(index).Login)) & ".ini"
    
    ' General
    PutVar filename, "ACCOUNT", "Login", Trim$(Player(index).Login)
    ' Banned
    PutVar filename, "ACCOUNT", "isBanned", Val(Player(index).isBanned)
    PutVar filename, "ACCOUNT", "isMuted", Val(Player(index).isMuted)
    PutVar filename, "ACCOUNT", "Usergroup", Val(Player(index).Usergroup)
    
    ' exit out early if invalid char
    If Player(index).charNum < 1 Or Player(index).charNum > MAX_CHARS Then Exit Sub
    
    ' the char header
    charHeader = "CHAR" & Player(index).charNum
    
    ' character
    PutVar filename, charHeader, "Name", Trim$(Player(index).Name)
    PutVar filename, charHeader, "Sex", Val(Player(index).Sex)
    PutVar filename, charHeader, "Class", Val(Player(index).Class)
    PutVar filename, charHeader, "Sprite", Val(Player(index).Sprite)
    PutVar filename, charHeader, "Level", Val(Player(index).Level)
    PutVar filename, charHeader, "exp", Val(Player(index).exp)
    PutVar filename, charHeader, "Access", Val(Player(index).Access)
    PutVar filename, charHeader, "PK", Val(Player(index).PK)
    
    ' Vitals
    For i = 1 To Vitals.Vital_Count - 1
        PutVar filename, charHeader, "Vital" & i, Val(Player(index).Vital(i))
    Next
    
    ' Stats
    For i = 1 To Stats.Stat_Count - 1
        PutVar filename, charHeader, "Stat" & i, Val(Player(index).Stat(i))
    Next
    PutVar filename, charHeader, "Points", Val(Player(index).POINTS)

    ' Equipment
    For i = 1 To Equipment.Equipment_Count - 1
        PutVar filename, charHeader, "Equipment" & i, Val(Player(index).Equipment(i))
    Next
    
    ' Inventory
    For i = 1 To MAX_INV
        PutVar filename, charHeader, "InvNum" & i, Val(Player(index).Inv(i).Num)
        PutVar filename, charHeader, "InvValue" & i, Val(Player(index).Inv(i).Value)
        PutVar filename, charHeader, "InvBound" & i, Val(Player(index).Inv(i).Bound)
    Next
    
    ' Spells
    For i = 1 To MAX_PLAYER_SPELLS
        PutVar filename, charHeader, "Spell" & i, Val(Player(index).Spell(i).Spell)
        PutVar filename, charHeader, "SpellUses" & i, Val(Player(index).Spell(i).Uses)
    Next
    
    ' Hotbar
    For i = 1 To MAX_HOTBAR
        PutVar filename, charHeader, "HotbarSlot" & i, Val(Player(index).Hotbar(i).Slot)
        PutVar filename, charHeader, "HotbarType" & i, Val(Player(index).Hotbar(i).sType)
    Next
    
    ' Position
    PutVar filename, charHeader, "Map", Val(Player(index).Map)
    PutVar filename, charHeader, "X", Val(Player(index).x)
    PutVar filename, charHeader, "Y", Val(Player(index).y)
    PutVar filename, charHeader, "Dir", Val(Player(index).dir)
    
    ' Tutorial
    PutVar filename, charHeader, "TutorialState", Val(Player(index).TutorialState)
    
    ' Bank
    For i = 1 To MAX_BANK
        PutVar filename, charHeader, "BankNum" & i, Val(Player(index).Bank(i).Num)
        PutVar filename, charHeader, "BankValue" & i, Val(Player(index).Bank(i).Value)
        PutVar filename, charHeader, "BankBound" & i, Val(Player(index).Bank(i).Bound)
    Next
    
    ' variables
    For i = 1 To MAX_BYTE
        PutVar filename, charHeader, "Var" & i, Val(Player(index).Variable(i))
    Next
End Sub

Sub LoadPlayer(ByVal index As Long, ByVal Name As String, ByVal charNum As Long)
Dim filename As String, i As Long, charHeader As String

    If Trim$(Name) = vbNullString Then Exit Sub
    ' clear player
    Call ClearPlayer(index)
    
    ' the file
    filename = App.Path & "\data\accounts\" & SanitiseString(Trim$(Name)) & ".ini"

    ' General
    Player(index).Login = Name
    ' Banned
    Player(index).isBanned = Val(GetVar(filename, "ACCOUNT", "isBanned"))
    Player(index).isMuted = Val(GetVar(filename, "ACCOUNT", "isMuted"))
    Player(index).Usergroup = Val(GetVar(filename, "ACCOUNT", "Usergroup"))
    
    ' exit out early if not a valid char num
    If charNum < 1 Or charNum > MAX_CHARS Then Exit Sub
    
    ' the char header
    charNum = charNum
    charHeader = "CHAR" & charNum
    
    ' character
    Player(index).Name = GetVar(filename, charHeader, "Name")
    Player(index).Sex = Val(GetVar(filename, charHeader, "Sex"))
    Player(index).Class = Val(GetVar(filename, charHeader, "Class"))
    Player(index).Sprite = Val(GetVar(filename, charHeader, "Sprite"))
    Player(index).Level = Val(GetVar(filename, charHeader, "Level"))
    Player(index).exp = Val(GetVar(filename, charHeader, "Exp"))
    Player(index).Access = Val(GetVar(filename, charHeader, "Access"))
    Player(index).PK = Val(GetVar(filename, charHeader, "PK"))
    
    ' Vitals
    For i = 1 To Vitals.Vital_Count - 1
        Player(index).Vital(i) = Val(GetVar(filename, charHeader, "Vital" & i))
    Next
    
    ' Stats
    For i = 1 To Stats.Stat_Count - 1
        Player(index).Stat(i) = Val(GetVar(filename, charHeader, "Stat" & i))
    Next
    Player(index).POINTS = Val(GetVar(filename, charHeader, "Points"))

    ' Equipment
    For i = 1 To Equipment.Equipment_Count - 1
        Player(index).Equipment(i) = Val(GetVar(filename, charHeader, "Equipment" & i))
    Next
    
    ' Inventory
    For i = 1 To MAX_INV
        Player(index).Inv(i).Num = Val(GetVar(filename, charHeader, "InvNum" & i))
        Player(index).Inv(i).Value = Val(GetVar(filename, charHeader, "InvValue" & i))
        Player(index).Inv(i).Bound = Val(GetVar(filename, charHeader, "InvBound" & i))
    Next
    
    ' Spells
    For i = 1 To MAX_PLAYER_SPELLS
        Player(index).Spell(i).Spell = Val(GetVar(filename, charHeader, "Spell" & i))
        Player(index).Spell(i).Uses = Val(GetVar(filename, charHeader, "SpellUses" & i))
    Next
    
    ' Hotbar
    For i = 1 To MAX_HOTBAR
        Player(index).Hotbar(i).Slot = Val(GetVar(filename, charHeader, "HotbarSlot" & i))
        Player(index).Hotbar(i).sType = Val(GetVar(filename, charHeader, "HotbarType" & i))
    Next
    
    ' Position
    Player(index).Map = Val(GetVar(filename, charHeader, "Map"))
    Player(index).x = Val(GetVar(filename, charHeader, "X"))
    Player(index).y = Val(GetVar(filename, charHeader, "Y"))
    Player(index).dir = Val(GetVar(filename, charHeader, "Dir"))
    
    ' Tutorial
    Player(index).TutorialState = Val(GetVar(filename, charHeader, "TutorialState"))
    
    ' Bank
    For i = 1 To MAX_BANK
        Player(index).Bank(i).Num = Val(GetVar(filename, charHeader, "BankNum" & i))
        Player(index).Bank(i).Value = Val(GetVar(filename, charHeader, "BankValue" & i))
        Player(index).Bank(i).Bound = Val(GetVar(filename, charHeader, "BankBound" & i))
    Next
    
    ' variables
    For i = 1 To MAX_BYTE
        Player(index).Variable(i) = Val(GetVar(filename, charHeader, "Var" & i))
    Next
    
    ' set the character number
    Player(index).charNum = charNum
End Sub

Sub DeleteCharacter(Login As String, charNum As Long)
Dim filename As String, charHeader As String, i As Long

    Login = Trim$(Login)
    If Login = vbNullString Then Exit Sub
    
    ' the file
    filename = App.Path & "\data\accounts\" & SanitiseString(Login) & ".ini"
    
    ' exit out early if invalid char
    If charNum < 1 Or charNum > MAX_CHARS Then Exit Sub
    
    ' the char header
    charHeader = "CHAR" & charNum
    
    ' character
    PutVar filename, charHeader, "Name", vbNullString
    PutVar filename, charHeader, "Sex", 0
    PutVar filename, charHeader, "Class", 0
    PutVar filename, charHeader, "Sprite", 0
    PutVar filename, charHeader, "Level", 0
    PutVar filename, charHeader, "exp", 0
    PutVar filename, charHeader, "Access", 0
    PutVar filename, charHeader, "PK", 0
    
    ' Vitals
    For i = 1 To Vitals.Vital_Count - 1
        PutVar filename, charHeader, "Vital" & i, 0
    Next
    
    ' Stats
    For i = 1 To Stats.Stat_Count - 1
        PutVar filename, charHeader, "Stat" & i, 0
    Next
    PutVar filename, charHeader, "Points", 0

    ' Equipment
    For i = 1 To Equipment.Equipment_Count - 1
        PutVar filename, charHeader, "Equipment" & i, 0
    Next
    
    ' Inventory
    For i = 1 To MAX_INV
        PutVar filename, charHeader, "InvNum" & i, 0
        PutVar filename, charHeader, "InvValue" & i, 0
        PutVar filename, charHeader, "InvBound" & i, 0
    Next
    
    ' Spells
    For i = 1 To MAX_PLAYER_SPELLS
        PutVar filename, charHeader, "Spell" & i, 0
        PutVar filename, charHeader, "SpellUses" & i, 0
    Next
    
    ' Hotbar
    For i = 1 To MAX_HOTBAR
        PutVar filename, charHeader, "HotbarSlot" & i, 0
        PutVar filename, charHeader, "HotbarType" & i, 0
    Next
    
    ' Position
    PutVar filename, charHeader, "Map", 0
    PutVar filename, charHeader, "X", 0
    PutVar filename, charHeader, "Y", 0
    PutVar filename, charHeader, "Dir", 0
    
    ' Tutorial
    PutVar filename, charHeader, "TutorialState", 0
    
    ' Bank
    For i = 1 To MAX_BANK
        PutVar filename, charHeader, "BankNum" & i, 0
        PutVar filename, charHeader, "BankValue" & i, 0
        PutVar filename, charHeader, "BankBound" & i, 0
    Next
End Sub

Sub ClearPlayer(ByVal index As Long)
    Dim i As Long
    
    Call ZeroMemory(ByVal VarPtr(TempPlayer(index)), LenB(TempPlayer(index)))
    Set TempPlayer(index).Buffer = New clsBuffer
    
    Call ZeroMemory(ByVal VarPtr(Player(index)), LenB(Player(index)))
    Player(index).Login = vbNullString
    Player(index).Name = vbNullString
    Player(index).Class = 1

    frmServer.lvwInfo.ListItems(index).SubItems(1) = vbNullString
    frmServer.lvwInfo.ListItems(index).SubItems(2) = vbNullString
    frmServer.lvwInfo.ListItems(index).SubItems(3) = vbNullString
End Sub

Sub ClearChar(ByVal index As Long)
Dim tmpName As String, tmpChar As Long
    
    tmpName = Player(index).Login
    tmpChar = Player(index).charNum
    
    Call ZeroMemory(ByVal VarPtr(Player(index)), LenB(Player(index)))
    
    Player(index).Login = tmpName
    Player(index).charNum = tmpChar
End Sub

' *************
' ** Classes **
' *************
Public Sub CreateClassesINI()
    Dim filename As String
    Dim File As String
    filename = App.Path & "\data\classes.ini"
    Max_Classes = 2

    If Not FileExist(filename, True) Then
        File = FreeFile
        Open filename For Output As File
        Print #File, "[INIT]"
        Print #File, "MaxClasses=" & Max_Classes
        Close File
    End If

End Sub

Sub LoadClasses()
    Dim filename As String
    Dim i As Long, n As Long
    Dim tmpSprite As String
    Dim tmpArray() As String
    Dim startItemCount As Long, startSpellCount As Long
    Dim x As Long

    If CheckClasses Then
        ReDim Class(1 To Max_Classes)
        Call SaveClasses
    Else
        filename = App.Path & "\data\classes.ini"
        Max_Classes = Val(GetVar(filename, "INIT", "MaxClasses"))
        ReDim Class(1 To Max_Classes)
    End If

    Call ClearClasses

    For i = 1 To Max_Classes
        Class(i).Name = GetVar(filename, "CLASS" & i, "Name")
        
        ' read string of sprites
        tmpSprite = GetVar(filename, "CLASS" & i, "MaleSprite")
        ' split into an array of strings
        tmpArray() = Split(tmpSprite, ",")
        ' redim the class sprite array
        ReDim Class(i).MaleSprite(0 To UBound(tmpArray))
        ' loop through converting strings to values and store in the sprite array
        For n = 0 To UBound(tmpArray)
            Class(i).MaleSprite(n) = Val(tmpArray(n))
        Next
        
        ' read string of sprites
        tmpSprite = GetVar(filename, "CLASS" & i, "FemaleSprite")
        ' split into an array of strings
        tmpArray() = Split(tmpSprite, ",")
        ' redim the class sprite array
        ReDim Class(i).FemaleSprite(0 To UBound(tmpArray))
        ' loop through converting strings to values and store in the sprite array
        For n = 0 To UBound(tmpArray)
            Class(i).FemaleSprite(n) = Val(tmpArray(n))
        Next
        
        ' continue
        Class(i).Stat(Stats.Strength) = Val(GetVar(filename, "CLASS" & i, "Strength"))
        Class(i).Stat(Stats.Endurance) = Val(GetVar(filename, "CLASS" & i, "Endurance"))
        Class(i).Stat(Stats.Intelligence) = Val(GetVar(filename, "CLASS" & i, "Intelligence"))
        Class(i).Stat(Stats.Agility) = Val(GetVar(filename, "CLASS" & i, "Agility"))
        Class(i).Stat(Stats.Willpower) = Val(GetVar(filename, "CLASS" & i, "Willpower"))
        
        ' how many starting items?
        startItemCount = Val(GetVar(filename, "CLASS" & i, "StartItemCount"))
        If startItemCount > 0 Then ReDim Class(i).StartItem(1 To startItemCount)
        If startItemCount > 0 Then ReDim Class(i).StartValue(1 To startItemCount)
        
        ' loop for items & values
        Class(i).startItemCount = startItemCount
        If startItemCount >= 1 And startItemCount <= MAX_INV Then
            For x = 1 To startItemCount
                Class(i).StartItem(x) = Val(GetVar(filename, "CLASS" & i, "StartItem" & x))
                Class(i).StartValue(x) = Val(GetVar(filename, "CLASS" & i, "StartValue" & x))
            Next
        End If
        
        ' how many starting spells?
        startSpellCount = Val(GetVar(filename, "CLASS" & i, "StartSpellCount"))
        If startSpellCount > 0 Then ReDim Class(i).StartSpell(1 To startSpellCount)
        
        ' loop for spells
        Class(i).startSpellCount = startSpellCount
        If startSpellCount >= 1 And startSpellCount <= MAX_INV Then
            For x = 1 To startSpellCount
                Class(i).StartSpell(x) = Val(GetVar(filename, "CLASS" & i, "StartSpell" & x))
            Next
        End If
    Next

End Sub

Sub SaveClasses()
    Dim filename As String
    Dim i As Long
    Dim x As Long
    
    filename = App.Path & "\data\classes.ini"

    For i = 1 To Max_Classes
        Call PutVar(filename, "CLASS" & i, "Name", Trim$(Class(i).Name))
        Call PutVar(filename, "CLASS" & i, "Maleprite", "1")
        Call PutVar(filename, "CLASS" & i, "Femaleprite", "1")
        Call PutVar(filename, "CLASS" & i, "Strength", STR(Class(i).Stat(Stats.Strength)))
        Call PutVar(filename, "CLASS" & i, "Endurance", STR(Class(i).Stat(Stats.Endurance)))
        Call PutVar(filename, "CLASS" & i, "Intelligence", STR(Class(i).Stat(Stats.Intelligence)))
        Call PutVar(filename, "CLASS" & i, "Agility", STR(Class(i).Stat(Stats.Agility)))
        Call PutVar(filename, "CLASS" & i, "Willpower", STR(Class(i).Stat(Stats.Willpower)))
        ' loop for items & values
        For x = 1 To UBound(Class(i).StartItem)
            Call PutVar(filename, "CLASS" & i, "StartItem" & x, STR(Class(i).StartItem(x)))
            Call PutVar(filename, "CLASS" & i, "StartValue" & x, STR(Class(i).StartValue(x)))
        Next
        ' loop for spells
        For x = 1 To UBound(Class(i).StartSpell)
            Call PutVar(filename, "CLASS" & i, "StartSpell" & x, STR(Class(i).StartSpell(x)))
        Next
    Next

End Sub

Function CheckClasses() As Boolean
    Dim filename As String
    filename = App.Path & "\data\classes.ini"

    If Not FileExist(filename, True) Then
        Call CreateClassesINI
        CheckClasses = True
    End If

End Function

Sub ClearClasses()
    Dim i As Long

    For i = 1 To Max_Classes
        Call ZeroMemory(ByVal VarPtr(Class(i)), LenB(Class(i)))
        Class(i).Name = vbNullString
    Next

End Sub

' ***********
' ** Items **
' ***********
Sub SaveItems()
    Dim i As Long

    For i = 1 To MAX_ITEMS
        Call SaveItem(i)
    Next

End Sub

Sub SaveItem(ByVal itemNum As Long)
    Dim filename As String
    Dim f  As Long
    filename = App.Path & "\data\items\item" & itemNum & ".dat"
    f = FreeFile
    Open filename For Binary As #f
    Put #f, , Item(itemNum)
    Close #f
End Sub

Sub LoadItems()
    Dim filename As String
    Dim i As Long
    Dim f As Long
    Call CheckItems

    For i = 1 To MAX_ITEMS
        filename = App.Path & "\data\Items\Item" & i & ".dat"
        f = FreeFile
        Open filename For Binary As #f
        Get #f, , Item(i)
        Close #f
    Next

End Sub

Sub CheckItems()
    Dim i As Long

    For i = 1 To MAX_ITEMS

        If Not FileExist("\Data\Items\Item" & i & ".dat") Then
            Call SaveItem(i)
        End If

    Next

End Sub

Sub ClearItem(ByVal index As Long)
    Call ZeroMemory(ByVal VarPtr(Item(index)), LenB(Item(index)))
    Item(index).Name = vbNullString
    Item(index).Desc = vbNullString
    Item(index).Sound = "None."
End Sub

Sub ClearItems()
    Dim i As Long

    For i = 1 To MAX_ITEMS
        Call ClearItem(i)
    Next

End Sub

' ***********
' ** Shops **
' ***********
Sub SaveShops()
    Dim i As Long

    For i = 1 To MAX_SHOPS
        Call SaveShop(i)
    Next

End Sub

Sub SaveShop(ByVal shopNum As Long)
    Dim filename As String
    Dim f As Long
    filename = App.Path & "\data\shops\shop" & shopNum & ".dat"
    f = FreeFile
    Open filename For Binary As #f
    Put #f, , Shop(shopNum)
    Close #f
End Sub

Sub LoadShops()
    Dim filename As String
    Dim i As Long
    Dim f As Long
    Call CheckShops

    For i = 1 To MAX_SHOPS
        filename = App.Path & "\data\shops\shop" & i & ".dat"
        f = FreeFile
        Open filename For Binary As #f
        Get #f, , Shop(i)
        Close #f
    Next

End Sub

Sub CheckShops()
    Dim i As Long

    For i = 1 To MAX_SHOPS

        If Not FileExist("\Data\shops\shop" & i & ".dat") Then
            Call SaveShop(i)
        End If

    Next

End Sub

Sub ClearShop(ByVal index As Long)
    Call ZeroMemory(ByVal VarPtr(Shop(index)), LenB(Shop(index)))
    Shop(index).Name = vbNullString
End Sub

Sub ClearShops()
    Dim i As Long

    For i = 1 To MAX_SHOPS
        Call ClearShop(i)
    Next

End Sub

' ************
' ** Spells **
' ************
Sub SaveSpell(ByVal spellNum As Long)
    Dim filename As String
    Dim f As Long
    filename = App.Path & "\data\spells\spells" & spellNum & ".dat"
    f = FreeFile
    Open filename For Binary As #f
    Put #f, , Spell(spellNum)
    Close #f
End Sub

Sub SaveSpells()
    Dim i As Long
    Call SetStatus("Saving spells... ")

    For i = 1 To MAX_SPELLS
        Call SaveSpell(i)
    Next

End Sub

Sub LoadSpells()
    Dim filename As String
    Dim i As Long
    Dim f As Long
    Call CheckSpells

    For i = 1 To MAX_SPELLS
        filename = App.Path & "\data\spells\spells" & i & ".dat"
        f = FreeFile
        Open filename For Binary As #f
        Get #f, , Spell(i)
        Close #f
    Next

End Sub

Sub CheckSpells()
    Dim i As Long

    For i = 1 To MAX_SPELLS

        If Not FileExist("\Data\spells\spells" & i & ".dat") Then
            Call SaveSpell(i)
        End If

    Next

End Sub

Sub ClearSpell(ByVal index As Long)
    Call ZeroMemory(ByVal VarPtr(Spell(index)), LenB(Spell(index)))
    Spell(index).Name = vbNullString
    Spell(index).LevelReq = 1 'Needs to be 1 for the spell editor
    Spell(index).Desc = vbNullString
    Spell(index).Sound = "None."
End Sub

Sub ClearSpells()
    Dim i As Long

    For i = 1 To MAX_SPELLS
        Call ClearSpell(i)
    Next

End Sub

' **********
' ** NPCs **
' **********
Sub SaveNpcs()
    Dim i As Long

    For i = 1 To MAX_NPCS
        Call SaveNpc(i)
    Next

End Sub

Sub SaveNpc(ByVal npcNum As Long)
    Dim filename As String
    Dim f As Long
    filename = App.Path & "\data\npcs\npc" & npcNum & ".dat"
    f = FreeFile
    Open filename For Binary As #f
    Put #f, , Npc(npcNum)
    Close #f
End Sub

Sub LoadNpcs()
    Dim filename As String
    Dim i As Long
    Dim f As Long
    Call CheckNpcs

    For i = 1 To MAX_NPCS
        filename = App.Path & "\data\npcs\npc" & i & ".dat"
        f = FreeFile
        Open filename For Binary As #f
        Get #f, , Npc(i)
        Close #f
    Next

End Sub

Sub CheckNpcs()
    Dim i As Long

    For i = 1 To MAX_NPCS

        If Not FileExist("\Data\npcs\npc" & i & ".dat") Then
            Call SaveNpc(i)
        End If

    Next

End Sub

Sub ClearNpc(ByVal index As Long)
    Call ZeroMemory(ByVal VarPtr(Npc(index)), LenB(Npc(index)))
    Npc(index).Name = vbNullString
    Npc(index).AttackSay = vbNullString
    Npc(index).Sound = "None."
End Sub

Sub ClearNpcs()
    Dim i As Long

    For i = 1 To MAX_NPCS
        Call ClearNpc(i)
    Next

End Sub

' **********
' ** Resources **
' **********
Sub SaveResources()
    Dim i As Long

    For i = 1 To MAX_RESOURCES
        Call SaveResource(i)
    Next

End Sub

Sub SaveResource(ByVal ResourceNum As Long)
    Dim filename As String
    Dim f As Long
    filename = App.Path & "\data\resources\resource" & ResourceNum & ".dat"
    f = FreeFile
    Open filename For Binary As #f
        Put #f, , Resource(ResourceNum)
    Close #f
End Sub

Sub LoadResources()
    Dim filename As String
    Dim i As Long
    Dim f As Long
    Dim sLen As Long
    
    Call CheckResources

    For i = 1 To MAX_RESOURCES
        filename = App.Path & "\data\resources\resource" & i & ".dat"
        f = FreeFile
        Open filename For Binary As #f
            Get #f, , Resource(i)
        Close #f
    Next

End Sub

Sub CheckResources()
    Dim i As Long

    For i = 1 To MAX_RESOURCES
        If Not FileExist("\Data\Resources\Resource" & i & ".dat") Then
            Call SaveResource(i)
        End If
    Next

End Sub

Sub ClearResource(ByVal index As Long)
    Call ZeroMemory(ByVal VarPtr(Resource(index)), LenB(Resource(index)))
    Resource(index).Name = vbNullString
    Resource(index).SuccessMessage = vbNullString
    Resource(index).EmptyMessage = vbNullString
    Resource(index).Sound = "None."
End Sub

Sub ClearResources()
    Dim i As Long

    For i = 1 To MAX_RESOURCES
        Call ClearResource(i)
    Next
End Sub

' **********
' ** animations **
' **********
Sub SaveAnimations()
    Dim i As Long

    For i = 1 To MAX_ANIMATIONS
        Call SaveAnimation(i)
    Next

End Sub

Sub SaveAnimation(ByVal AnimationNum As Long)
    Dim filename As String
    Dim f As Long
    filename = App.Path & "\data\animations\animation" & AnimationNum & ".dat"
    f = FreeFile
    Open filename For Binary As #f
        Put #f, , Animation(AnimationNum)
    Close #f
End Sub

Sub LoadAnimations()
    Dim filename As String
    Dim i As Long
    Dim f As Long
    Dim sLen As Long
    
    Call CheckAnimations

    For i = 1 To MAX_ANIMATIONS
        filename = App.Path & "\data\animations\animation" & i & ".dat"
        f = FreeFile
        Open filename For Binary As #f
            Get #f, , Animation(i)
        Close #f
    Next

End Sub

Sub CheckAnimations()
    Dim i As Long

    For i = 1 To MAX_ANIMATIONS

        If Not FileExist("\Data\animations\animation" & i & ".dat") Then
            Call SaveAnimation(i)
        End If

    Next

End Sub

Sub ClearAnimation(ByVal index As Long)
    Call ZeroMemory(ByVal VarPtr(Animation(index)), LenB(Animation(index)))
    Animation(index).Name = vbNullString
    Animation(index).Sound = "None."
End Sub

Sub ClearAnimations()
    Dim i As Long

    For i = 1 To MAX_ANIMATIONS
        Call ClearAnimation(i)
    Next
End Sub

' **********
' ** Maps **
' **********
Sub SaveMap(ByVal mapnum As Long)
    Dim filename As String, f As Long, x As Long, y As Long, i As Long
    
    ' save map data
    filename = App.Path & "\data\maps\map" & mapnum & ".ini"
    
    ' if it exists then kill the ini
    If FileExist(filename, True) Then Kill filename
    
    ' General
    With Map(mapnum).MapData
        PutVar filename, "General", "Name", .Name
        PutVar filename, "General", "Music", .Music
        PutVar filename, "General", "Moral", Val(.Moral)
        PutVar filename, "General", "Up", Val(.Up)
        PutVar filename, "General", "Down", Val(.Down)
        PutVar filename, "General", "Left", Val(.left)
        PutVar filename, "General", "Right", Val(.Right)
        PutVar filename, "General", "BootMap", Val(.BootMap)
        PutVar filename, "General", "BootX", Val(.BootX)
        PutVar filename, "General", "BootY", Val(.BootY)
        PutVar filename, "General", "MaxX", Val(.MaxX)
        PutVar filename, "General", "MaxY", Val(.MaxY)
        PutVar filename, "General", "BossNpc", Val(.BossNpc)
        For i = 1 To MAX_MAP_NPCS
            PutVar filename, "General", "Npc" & i, Val(.Npc(i))
        Next
    End With
    
    ' Events
    PutVar filename, "Events", "EventCount", Val(Map(mapnum).TileData.EventCount)
    
    If Map(mapnum).TileData.EventCount > 0 Then
        For i = 1 To Map(mapnum).TileData.EventCount
            With Map(mapnum).TileData.Events(i)
                PutVar filename, "Event" & i, "Name", .Name
                PutVar filename, "Event" & i, "x", Val(.x)
                PutVar filename, "Event" & i, "y", Val(.y)
                PutVar filename, "Event" & i, "PageCount", Val(.PageCount)
            End With
            If Map(mapnum).TileData.Events(i).PageCount > 0 Then
                For x = 1 To Map(mapnum).TileData.Events(i).PageCount
                    With Map(mapnum).TileData.Events(i).EventPage(x)
                        PutVar filename, "Event" & i & "Page" & x, "chkPlayerVar", Val(.chkPlayerVar)
                        PutVar filename, "Event" & i & "Page" & x, "chkSelfSwitch", Val(.chkSelfSwitch)
                        PutVar filename, "Event" & i & "Page" & x, "chkHasItem", Val(.chkHasItem)
                        PutVar filename, "Event" & i & "Page" & x, "PlayerVarNum", Val(.PlayerVarNum)
                        PutVar filename, "Event" & i & "Page" & x, "SelfSwitchNum", Val(.SelfSwitchNum)
                        PutVar filename, "Event" & i & "Page" & x, "HasItemNum", Val(.HasItemNum)
                        PutVar filename, "Event" & i & "Page" & x, "PlayerVariable", Val(.PlayerVariable)
                        PutVar filename, "Event" & i & "Page" & x, "GraphicType", Val(.GraphicType)
                        PutVar filename, "Event" & i & "Page" & x, "Graphic", Val(.Graphic)
                        PutVar filename, "Event" & i & "Page" & x, "GraphicX", Val(.GraphicX)
                        PutVar filename, "Event" & i & "Page" & x, "GraphicY", Val(.GraphicY)
                        PutVar filename, "Event" & i & "Page" & x, "MoveType", Val(.MoveType)
                        PutVar filename, "Event" & i & "Page" & x, "MoveSpeed", Val(.MoveSpeed)
                        PutVar filename, "Event" & i & "Page" & x, "MoveFreq", Val(.MoveFreq)
                        PutVar filename, "Event" & i & "Page" & x, "WalkAnim", Val(.WalkAnim)
                        PutVar filename, "Event" & i & "Page" & x, "StepAnim", Val(.StepAnim)
                        PutVar filename, "Event" & i & "Page" & x, "DirFix", Val(.DirFix)
                        PutVar filename, "Event" & i & "Page" & x, "WalkThrough", Val(.WalkThrough)
                        PutVar filename, "Event" & i & "Page" & x, "Priority", Val(.Priority)
                        PutVar filename, "Event" & i & "Page" & x, "Trigger", Val(.Trigger)
                        PutVar filename, "Event" & i & "Page" & x, "CommandCount", Val(.CommandCount)
                    End With
                    If Map(mapnum).TileData.Events(i).EventPage(x).CommandCount > 0 Then
                        For y = 1 To Map(mapnum).TileData.Events(i).EventPage(x).CommandCount
                            With Map(mapnum).TileData.Events(i).EventPage(x).Commands(y)
                                PutVar filename, "Event" & i & "Page" & x & "Command" & y, "Type", Val(.Type)
                                PutVar filename, "Event" & i & "Page" & x & "Command" & y, "Text", .Text
                                PutVar filename, "Event" & i & "Page" & x & "Command" & y, "Colour", Val(.colour)
                                PutVar filename, "Event" & i & "Page" & x & "Command" & y, "Channel", Val(.Channel)
                                PutVar filename, "Event" & i & "Page" & x & "Command" & y, "TargetType", Val(.targetType)
                                PutVar filename, "Event" & i & "Page" & x & "Command" & y, "Target", Val(.target)
                            End With
                        Next
                    End If
                Next
            End If
        Next
    End If
    
    ' dump tile data
    filename = App.Path & "\data\maps\map" & mapnum & ".dat"
    f = FreeFile
    
    With Map(mapnum)
        Open filename For Binary As #f
            For x = 0 To .MapData.MaxX
                For y = 0 To .MapData.MaxY
                    Put #f, , .TileData.Tile(x, y).Type
                    Put #f, , .TileData.Tile(x, y).Data1
                    Put #f, , .TileData.Tile(x, y).Data2
                    Put #f, , .TileData.Tile(x, y).Data3
                    Put #f, , .TileData.Tile(x, y).Data4
                    Put #f, , .TileData.Tile(x, y).Data5
                    Put #f, , .TileData.Tile(x, y).Autotile
                    Put #f, , .TileData.Tile(x, y).DirBlock
                    For i = 1 To MapLayer.Layer_Count - 1
                        Put #f, , .TileData.Tile(x, y).Layer(i).Tileset
                        Put #f, , .TileData.Tile(x, y).Layer(i).x
                        Put #f, , .TileData.Tile(x, y).Layer(i).y
                    Next
                Next
            Next
        Close #f
    End With
    
    DoEvents
End Sub

Sub SaveMaps()
    Dim i As Long

    For i = 1 To MAX_MAPS
        Call SaveMap(i)
    Next

End Sub

Sub LoadMaps()
    Dim filename As String, mapnum As Long
    
    Call CheckMaps

    For mapnum = 1 To MAX_MAPS
        LoadMap mapnum
        ClearTempTile mapnum
        CacheResources mapnum
        DoEvents
    Next
End Sub

Sub GetMapCRC32(mapnum As Long)
Dim Data() As Byte, filename As String, f As Long
    ' map data
    filename = App.Path & "\data\maps\map" & mapnum & ".ini"
    If FileExist(filename, True) Then
        f = FreeFile
        Open filename For Binary As #f
            Data = Space$(LOF(f))
            Get #f, , Data
        Close #f
        MapCRC32(mapnum).MapDataCRC = CRC32(Data)
    Else
        MapCRC32(mapnum).MapDataCRC = 0
    End If
    ' clear
    Erase Data
    ' tile data
    filename = App.Path & "\data\maps\map" & mapnum & ".dat"
    If FileExist(filename, True) Then
        f = FreeFile
        Open filename For Binary As #f
            Data = Space$(LOF(f))
            Get #f, , Data
        Close #f
        MapCRC32(mapnum).MapTileCRC = CRC32(Data)
    Else
        MapCRC32(mapnum).MapTileCRC = 0
    End If
End Sub

Sub LoadMap(mapnum As Long)
    Dim filename As String, i As Long, f As Long, x As Long, y As Long
    
    ' load map data
    filename = App.Path & "\data\maps\map" & mapnum & ".ini"
    
    ' General
    With Map(mapnum).MapData
        .Name = GetVar(filename, "General", "Name")
        .Music = GetVar(filename, "General", "Music")
        .Moral = Val(GetVar(filename, "General", "Moral"))
        .Up = Val(GetVar(filename, "General", "Up"))
        .Down = Val(GetVar(filename, "General", "Down"))
        .left = Val(GetVar(filename, "General", "Left"))
        .Right = Val(GetVar(filename, "General", "Right"))
        .BootMap = Val(GetVar(filename, "General", "BootMap"))
        .BootX = Val(GetVar(filename, "General", "BootX"))
        .BootY = Val(GetVar(filename, "General", "BootY"))
        .MaxX = Val(GetVar(filename, "General", "MaxX"))
        .MaxY = Val(GetVar(filename, "General", "MaxY"))
        .BossNpc = Val(GetVar(filename, "General", "BossNpc"))
        For i = 1 To MAX_MAP_NPCS
            .Npc(i) = Val(GetVar(filename, "General", "Npc" & i))
        Next
    End With
    
    ' Events
    Map(mapnum).TileData.EventCount = Val(GetVar(filename, "Events", "EventCount"))
    
    If Map(mapnum).TileData.EventCount > 0 Then
        ReDim Preserve Map(mapnum).TileData.Events(1 To Map(mapnum).TileData.EventCount)
        For i = 1 To Map(mapnum).TileData.EventCount
            With Map(mapnum).TileData.Events(i)
                .Name = GetVar(filename, "Event" & i, "Name")
                .x = Val(GetVar(filename, "Event" & i, "x"))
                .y = Val(GetVar(filename, "Event" & i, "y"))
                .PageCount = Val(GetVar(filename, "Event" & i, "PageCount"))
            End With
            If Map(mapnum).TileData.Events(i).PageCount > 0 Then
                ReDim Preserve Map(mapnum).TileData.Events(i).EventPage(1 To Map(mapnum).TileData.Events(i).PageCount)
                For x = 1 To Map(mapnum).TileData.Events(i).PageCount
                    With Map(mapnum).TileData.Events(i).EventPage(x)
                        .chkPlayerVar = Val(GetVar(filename, "Event" & i & "Page" & x, "chkPlayerVar"))
                        .chkSelfSwitch = Val(GetVar(filename, "Event" & i & "Page" & x, "chkSelfSwitch"))
                        .chkHasItem = Val(GetVar(filename, "Event" & i & "Page" & x, "chkHasItem"))
                        .PlayerVarNum = Val(GetVar(filename, "Event" & i & "Page" & x, "PlayerVarNum"))
                        .SelfSwitchNum = Val(GetVar(filename, "Event" & i & "Page" & x, "SelfSwitchNum"))
                        .HasItemNum = Val(GetVar(filename, "Event" & i & "Page" & x, "HasItemNum"))
                        .PlayerVariable = Val(GetVar(filename, "Event" & i & "Page" & x, "PlayerVariable"))
                        .GraphicType = Val(GetVar(filename, "Event" & i & "Page" & x, "GraphicType"))
                        .Graphic = Val(GetVar(filename, "Event" & i & "Page" & x, "Graphic"))
                        .GraphicX = Val(GetVar(filename, "Event" & i & "Page" & x, "GraphicX"))
                        .GraphicY = Val(GetVar(filename, "Event" & i & "Page" & x, "GraphicY"))
                        .MoveType = Val(GetVar(filename, "Event" & i & "Page" & x, "MoveType"))
                        .MoveSpeed = Val(GetVar(filename, "Event" & i & "Page" & x, "MoveSpeed"))
                        .MoveFreq = Val(GetVar(filename, "Event" & i & "Page" & x, "MoveFreq"))
                        .WalkAnim = Val(GetVar(filename, "Event" & i & "Page" & x, "WalkAnim"))
                        .StepAnim = Val(GetVar(filename, "Event" & i & "Page" & x, "StepAnim"))
                        .DirFix = Val(GetVar(filename, "Event" & i & "Page" & x, "DirFix"))
                        .WalkThrough = Val(GetVar(filename, "Event" & i & "Page" & x, "WalkThrough"))
                        .Priority = Val(GetVar(filename, "Event" & i & "Page" & x, "Priority"))
                        .Trigger = Val(GetVar(filename, "Event" & i & "Page" & x, "Trigger"))
                        .CommandCount = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandCount"))
                    End With
                    If Map(mapnum).TileData.Events(i).EventPage(x).CommandCount > 0 Then
                        ReDim Preserve Map(mapnum).TileData.Events(i).EventPage(x).Commands(1 To Map(mapnum).TileData.Events(i).EventPage(x).CommandCount)
                        For y = 1 To Map(mapnum).TileData.Events(i).EventPage(x).CommandCount
                            With Map(mapnum).TileData.Events(i).EventPage(x).Commands(y)
                                .Type = Val(GetVar(filename, "Event" & i & "Page" & x & "Command" & y, "Type"))
                                .Text = GetVar(filename, "Event" & i & "Page" & x & "Command" & y, "Text")
                                .colour = Val(GetVar(filename, "Event" & i & "Page" & x & "Command" & y, "Colour"))
                                .Channel = Val(GetVar(filename, "Event" & i & "Page" & x & "Command" & y, "Channel"))
                                .targetType = Val(GetVar(filename, "Event" & i & "Page" & x & "Command" & y, "TargetType"))
                                .target = Val(GetVar(filename, "Event" & i & "Page" & x & "Command" & y, "Target"))
                            End With
                        Next
                    End If
                Next
            End If
        Next
    End If
    
    ' dump tile data
    filename = App.Path & "\data\maps\map" & mapnum & ".dat"
    f = FreeFile
    
    ' redim the map
    ReDim Map(mapnum).TileData.Tile(0 To Map(mapnum).MapData.MaxX, 0 To Map(mapnum).MapData.MaxY) As TileRec
    
    With Map(mapnum)
        Open filename For Binary As #f
            For x = 0 To .MapData.MaxX
                For y = 0 To .MapData.MaxY
                    Get #f, , .TileData.Tile(x, y).Type
                    Get #f, , .TileData.Tile(x, y).Data1
                    Get #f, , .TileData.Tile(x, y).Data2
                    Get #f, , .TileData.Tile(x, y).Data3
                    Get #f, , .TileData.Tile(x, y).Data4
                    Get #f, , .TileData.Tile(x, y).Data5
                    Get #f, , .TileData.Tile(x, y).Autotile
                    Get #f, , .TileData.Tile(x, y).DirBlock
                    For i = 1 To MapLayer.Layer_Count - 1
                        Get #f, , .TileData.Tile(x, y).Layer(i).Tileset
                        Get #f, , .TileData.Tile(x, y).Layer(i).x
                        Get #f, , .TileData.Tile(x, y).Layer(i).y
                    Next
                Next
            Next
        Close #f
    End With
End Sub

Sub CheckMaps()
    Dim i As Long

    For i = 1 To MAX_MAPS

        If Not FileExist("\Data\maps\map" & i & ".dat") Or Not FileExist("\Data\maps\map" & i & ".ini") Then
            Call SaveMap(i)
        End If

    Next

End Sub

Sub ClearMapItem(ByVal index As Long, ByVal mapnum As Long)
    Call ZeroMemory(ByVal VarPtr(MapItem(mapnum, index)), LenB(MapItem(mapnum, index)))
    MapItem(mapnum, index).playerName = vbNullString
End Sub

Sub ClearMapItems()
    Dim x As Long
    Dim y As Long

    For y = 1 To MAX_MAPS
        For x = 1 To MAX_MAP_ITEMS
            Call ClearMapItem(x, y)
        Next
    Next

End Sub

Sub ClearMapNpc(ByVal index As Long, ByVal mapnum As Long)
    'ReDim MapNpc(mapnum).Npc(1 To MAX_MAP_NPCS)
    Call ZeroMemory(ByVal VarPtr(MapNpc(mapnum).Npc(index)), LenB(MapNpc(mapnum).Npc(index)))
End Sub

Sub ClearMapNpcs()
    Dim x As Long
    Dim y As Long

    For y = 1 To MAX_MAPS
        For x = 1 To MAX_MAP_NPCS
            Call ClearMapNpc(x, y)
        Next
    Next

End Sub

Sub ClearMap(ByVal mapnum As Long)
    Call ZeroMemory(ByVal VarPtr(Map(mapnum)), LenB(Map(mapnum)))
    Map(mapnum).MapData.Name = vbNullString
    Map(mapnum).MapData.MaxX = MAX_MAPX
    Map(mapnum).MapData.MaxY = MAX_MAPY
    ReDim Map(mapnum).TileData.Tile(0 To Map(mapnum).MapData.MaxX, 0 To Map(mapnum).MapData.MaxY)
    ' Reset the values for if a player is on the map or not
    PlayersOnMap(mapnum) = NO
    ' Reset the map cache array for this map.
    MapCache(mapnum).Data = vbNullString
End Sub

Sub ClearMaps()
    Dim i As Long

    For i = 1 To MAX_MAPS
        Call ClearMap(i)
    Next

End Sub

Function GetClassName(ByVal ClassNum As Long) As String
    GetClassName = Trim$(Class(ClassNum).Name)
End Function

Function GetClassMaxVital(ByVal ClassNum As Long, ByVal Vital As Vitals) As Long
    Select Case Vital
        Case HP
            With Class(ClassNum)
                GetClassMaxVital = 100 + (.Stat(Endurance) * 5) + 2
            End With
        Case MP
            With Class(ClassNum)
                GetClassMaxVital = 30 + (.Stat(Intelligence) * 10) + 2
            End With
    End Select
End Function

Function GetClassStat(ByVal ClassNum As Long, ByVal Stat As Stats) As Long
    GetClassStat = Class(ClassNum).Stat(Stat)
End Function

Sub ClearParty(ByVal partynum As Long)
    Call ZeroMemory(ByVal VarPtr(Party(partynum)), LenB(Party(partynum)))
End Sub

' ***********
' ** Convs **
' ***********
Sub SaveConvs()
Dim i As Long

    For i = 1 To MAX_CONVS
        Call SaveConv(i)
    Next
End Sub

Sub SaveConv(ByVal convNum As Long)
Dim filename As String
Dim i As Long, x As Long, f As Long
    
    filename = App.Path & "\data\convs\conv" & convNum & ".dat"
    f = FreeFile
    
    Open filename For Binary As #f
        With Conv(convNum)
            Put #f, , .Name
            Put #f, , .chatCount
            For i = 1 To .chatCount
                Put #f, , CLng(Len(.Conv(i).Conv))
                Put #f, , .Conv(i).Conv
                For x = 1 To 4
                    Put #f, , CLng(Len(.Conv(i).rText(x)))
                    Put #f, , .Conv(i).rText(x)
                    Put #f, , .Conv(i).rTarget(x)
                Next
                Put #f, , .Conv(i).Event
                Put #f, , .Conv(i).Data1
                Put #f, , .Conv(i).Data2
                Put #f, , .Conv(i).Data3
            Next
        End With
    Close #f
End Sub

Sub LoadConvs()
Dim filename As String
Dim i As Long, n As Long, x As Long, f As Long
Dim sLen As Long
    
    Call CheckConvs

    For i = 1 To MAX_CONVS
        filename = App.Path & "\data\convs\conv" & i & ".dat"
        f = FreeFile
        Open filename For Binary As #f
            With Conv(i)
                Get #f, , .Name
                Get #f, , .chatCount
                If .chatCount > 0 Then ReDim .Conv(1 To .chatCount)
                For n = 1 To .chatCount
                    Get #f, , sLen
                    .Conv(n).Conv = Space$(sLen)
                    Get #f, , .Conv(n).Conv
                    For x = 1 To 4
                        Get #f, , sLen
                        .Conv(n).rText(x) = Space$(sLen)
                        Get #f, , .Conv(n).rText(x)
                        Get #f, , .Conv(n).rTarget(x)
                    Next
                    Get #f, , .Conv(n).Event
                    Get #f, , .Conv(n).Data1
                    Get #f, , .Conv(n).Data2
                    Get #f, , .Conv(n).Data3
                Next
            End With
        Close #f
    Next
End Sub

Sub CheckConvs()
Dim i As Long

    For i = 1 To MAX_CONVS
        If Not FileExist("\data\convs\conv" & i & ".dat") Then
            Call SaveConv(i)
        End If
    Next
End Sub

Sub ClearConv(ByVal index As Long)
    Call ZeroMemory(ByVal VarPtr(Conv(index)), LenB(Conv(index)))
    Conv(index).Name = vbNullString
    ReDim Conv(index).Conv(1)
End Sub

Sub ClearConvs()
Dim i As Long

    For i = 1 To MAX_CONVS
        Call ClearConv(i)
    Next

End Sub

Function OldAccount_Exist(ByVal username As String) As Boolean
Dim filename As String
    
    filename = App.Path & "\data\accounts\old\" & SanitiseString(username) & ".ini"
    If FileExist(filename, True) Then
        If LenB(Trim$(GetVar(filename, "ACCOUNT", "Name"))) > 0 Then
            OldAccount_Exist = True
        End If
    End If
End Function

Public Sub MergeAccount(ByVal index As Long, ByVal charNum As Long, ByVal oldAccount As String)
Dim tempChar As PlayerRec, charHeader As String, filename As String, i As Long

    ' set the filename
    filename = App.Path & "\data\accounts\old\" & SanitiseString(oldAccount) & ".ini"
    charHeader = "ACCOUNT"
    
    ' load the old account shit
    With tempChar
        .Name = Trim$(GetVar(filename, charHeader, "Name"))
        .Sex = Val(GetVar(filename, charHeader, "Sex"))
        .Class = Val(GetVar(filename, charHeader, "Class"))
        .Sprite = Val(GetVar(filename, charHeader, "Sprite"))
        .Level = Val(GetVar(filename, charHeader, "Level"))
        .exp = Val(GetVar(filename, charHeader, "Exp"))
        .Access = Val(GetVar(filename, charHeader, "Access"))
        .PK = Val(GetVar(filename, charHeader, "PK"))
        
        ' Vitals
        For i = 1 To Vitals.Vital_Count - 1
            .Vital(i) = Val(GetVar(filename, charHeader, "Vital" & i))
        Next
        
        ' Stats
        For i = 1 To Stats.Stat_Count - 1
            .Stat(i) = Val(GetVar(filename, charHeader, "Stat" & i))
        Next
        .POINTS = Val(GetVar(filename, charHeader, "Points"))
    
        ' Equipment
        For i = 1 To Equipment.Equipment_Count - 1
            .Equipment(i) = Val(GetVar(filename, charHeader, "Equipment" & i))
        Next
        
        ' Inventory
        For i = 1 To MAX_INV
            .Inv(i).Num = Val(GetVar(filename, charHeader, "InvNum" & i))
            .Inv(i).Value = Val(GetVar(filename, charHeader, "InvValue" & i))
            .Inv(i).Bound = Val(GetVar(filename, charHeader, "InvBound" & i))
        Next
        
        ' Spells
        For i = 1 To MAX_PLAYER_SPELLS
            .Spell(i).Spell = Val(GetVar(filename, charHeader, "Spell" & i))
            .Spell(i).Uses = Val(GetVar(filename, charHeader, "SpellUses" & i))
        Next
        
        ' Hotbar
        For i = 1 To MAX_HOTBAR
            .Hotbar(i).Slot = Val(GetVar(filename, charHeader, "HotbarSlot" & i))
            .Hotbar(i).sType = Val(GetVar(filename, charHeader, "HotbarType" & i))
        Next
        
        ' Position
        .Map = Val(GetVar(filename, charHeader, "Map"))
        .x = Val(GetVar(filename, charHeader, "X"))
        .y = Val(GetVar(filename, charHeader, "Y"))
        .dir = Val(GetVar(filename, charHeader, "Dir"))
        
        ' Tutorial
        .TutorialState = Val(GetVar(filename, charHeader, "TutorialState"))
    End With
    
    ' set the filename
    filename = App.Path & "\data\accounts\" & SanitiseString(Trim$(Player(index).Login)) & ".ini"
    charHeader = "CHAR" & charNum
    
    ' save it in the new account's character slot
    With tempChar
        PutVar filename, charHeader, "Name", Trim$(.Name)
        PutVar filename, charHeader, "Sex", Val(.Sex)
        PutVar filename, charHeader, "Class", Val(.Class)
        PutVar filename, charHeader, "Sprite", Val(.Sprite)
        PutVar filename, charHeader, "Level", Val(.Level)
        PutVar filename, charHeader, "exp", Val(.exp)
        PutVar filename, charHeader, "Access", Val(.Access)
        PutVar filename, charHeader, "PK", Val(.PK)
        
        ' Vitals
        For i = 1 To Vitals.Vital_Count - 1
            PutVar filename, charHeader, "Vital" & i, Val(.Vital(i))
        Next
        
        ' Stats
        For i = 1 To Stats.Stat_Count - 1
            PutVar filename, charHeader, "Stat" & i, Val(.Stat(i))
        Next
        PutVar filename, charHeader, "Points", Val(.POINTS)
    
        ' Equipment
        For i = 1 To Equipment.Equipment_Count - 1
            PutVar filename, charHeader, "Equipment" & i, Val(.Equipment(i))
        Next
        
        ' Inventory
        For i = 1 To MAX_INV
            PutVar filename, charHeader, "InvNum" & i, Val(.Inv(i).Num)
            PutVar filename, charHeader, "InvValue" & i, Val(.Inv(i).Value)
            PutVar filename, charHeader, "InvBound" & i, Val(.Inv(i).Bound)
        Next
        
        ' Spells
        For i = 1 To MAX_PLAYER_SPELLS
            PutVar filename, charHeader, "Spell" & i, Val(.Spell(i).Spell)
            PutVar filename, charHeader, "SpellUses" & i, Val(.Spell(i).Uses)
        Next
        
        ' Hotbar
        For i = 1 To MAX_HOTBAR
            PutVar filename, charHeader, "HotbarSlot" & i, Val(.Hotbar(i).Slot)
            PutVar filename, charHeader, "HotbarType" & i, Val(.Hotbar(i).sType)
        Next
        
        ' Position
        PutVar filename, charHeader, "Map", Val(.Map)
        PutVar filename, charHeader, "X", Val(.x)
        PutVar filename, charHeader, "Y", Val(.y)
        PutVar filename, charHeader, "Dir", Val(.dir)
        
        ' Tutorial
        PutVar filename, charHeader, "TutorialState", Val(.TutorialState)
    End With
    
    ' kill the old account - permanently
    Kill App.Path & "\data\accounts\old\" & SanitiseString(oldAccount) & ".ini"
    
    ' send to portal again
    SendPlayerChars index
    
    ' confirmation message
    AlertMsg index, DIALOGUE_MSG_MERGE, MENU_CHARS, False
End Sub
