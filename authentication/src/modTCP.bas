Attribute VB_Name = "modTCP"
Option Explicit

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Function Current_IP(ByVal Index As Long) As String
    Current_IP = frmMain.Socket(Index).RemoteHostIP
End Function

Function ConnectToGameServer() As Boolean
Dim Wait As Long
    
    ' Check to see if we are already connected, if so just exit
    If IsConnectedGameServer Then
        ConnectToGameServer = True
        Exit Function
    End If
    
    Wait = GetTickCount
    frmMain.ServerSocket.Close
    frmMain.ServerSocket.RemoteHost = GAME_SERVER_IP
    frmMain.ServerSocket.RemotePort = SERVER_AUTH_PORT
    frmMain.ServerSocket.Connect
    
    ' Wait until connected or 3 seconds have passed and report the server being down
    'Do While (Not IsConnectedGameServer) And (GetTickCount <= Wait + 3000)
    '    Sleep 1
    '    DoEvents
    'Loop
    
    ConnectToGameServer = IsConnectedGameServer
End Function

Sub AcceptConnection(ByVal Index As Long, ByVal SocketId As Long)
Dim I As Long

    If Index = 0 Then
        I = FindOpenPlayerSlot
        
        If I <> 0 Then
            ' Whoho, we can connect them
            frmMain.Socket(I).Close
            frmMain.Socket(I).Accept SocketId
            SocketConnected I
        End If
    End If
End Sub

Sub SocketConnected(ByVal Index As Long)
    AddText frmMain.txtLog, "Received connection from " & Current_IP(Index) & "."
End Sub

Sub IncomingData(ByVal Index As Long, ByVal DataLength As Long)
Dim Buffer() As Byte
Dim pLength As Long

    ' Check for data flooding
    If Player(Index).DataBytes > 1000 Then
        Exit Sub
    End If

    ' Check for packet flooding
    If Player(Index).DataPackets > 25 Then
        Exit Sub
    End If
            
    ' Check if elapsed time has passed
    Player(Index).DataBytes = Player(Index).DataBytes + DataLength
    If GetTickCount >= Player(Index).DataTimer Then
        Player(Index).DataTimer = GetTickCount + 1000
        Player(Index).DataBytes = 0
        Player(Index).DataPackets = 0
    End If
    
    ' Get the data from the socket now
    frmMain.Socket(Index).GetData Buffer(), vbUnicode, DataLength
    Player(Index).Buffer.WriteBytes Buffer()
    
    If Player(Index).Buffer.Length >= 4 Then
        pLength = Player(Index).Buffer.ReadLong(False)
    
        If pLength < 0 Then
            Exit Sub
        End If
    End If
    
    Do While pLength > 0 And pLength <= Player(Index).Buffer.Length - 4
        If pLength <= Player(Index).Buffer.Length - 4 Then
            Player(Index).DataPackets = Player(Index).DataPackets + 1
            Player(Index).Buffer.ReadLong
            HandleData Index, Player(Index).Buffer.ReadBytes(pLength)
        End If
        
        pLength = 0
        If Player(Index).Buffer.Length >= 4 Then
            pLength = Player(Index).Buffer.ReadLong(False)
        
            If pLength < 0 Then
                Exit Sub
            End If
        End If
    Loop
            
    Player(Index).Buffer.Trim
End Sub

Sub CloseSocket(ByVal Index As Long)
    ClearPlayer Index
    AddText frmMain.txtLog, "Connection from " & Current_IP(Index) & " has been terminated."
    frmMain.Socket(Index).Close
End Sub

Function FindOpenPlayerSlot() As Long
Dim I As Long
    
    For I = 1 To MAX_PLAYERS
        If Not IsConnected(I) Then
            FindOpenPlayerSlot = I
            Exit Function
        End If
    Next
End Function

Function IsConnected(ByVal Index As Long) As Boolean
    If frmMain.Socket(Index).State = sckConnected Then IsConnected = True
End Function

Function IsConnectedGameServer() As Boolean
    IsConnectedGameServer = frmMain.ServerSocket.State = sckConnected
End Function

Sub SendDataTo(ByVal Index As Long, ByRef Data() As Byte)
Dim Buffer As clsBuffer
Dim tempData() As Byte

    If IsConnected(Index) Then
        Set Buffer = New clsBuffer
        tempData = Data
        
        Buffer.PreAllocate 4 + (UBound(tempData) - LBound(tempData)) + 1
        Buffer.WriteLong (UBound(tempData) - LBound(tempData)) + 1
        Buffer.WriteBytes tempData()
        
        frmMain.Socket(Index).SendData Buffer.ToArray()
        
    End If
End Sub

Sub SendDataToGameServer(ByRef Data() As Byte)
Dim Buffer As clsBuffer
Dim tempData() As Byte

    Set Buffer = New clsBuffer
    tempData = Data
    
    If Not ConnectToGameServer Then Exit Sub
       
    Buffer.PreAllocate 4 + (UBound(tempData) - LBound(tempData)) + 1
    Buffer.WriteLong (UBound(tempData) - LBound(tempData)) + 1
    Buffer.WriteBytes tempData()
    
    frmMain.ServerSocket.SendData Buffer.ToArray()
End Sub

Sub HackingAttempt(ByVal Index As Long)
    SendAlertMsg Index, DIALOGUE_MSG_CONNECTION
End Sub

Sub SendAlertMsg(ByVal Index As Long, ByVal Msg As Long, Optional ByVal menuReset As Long = 0, Optional ByVal kick As Boolean = True)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SAlertMsg
    Buffer.WriteLong Msg
    Buffer.WriteLong menuReset
    If kick Then Buffer.WriteLong 1 Else Buffer.WriteLong 0
    
    SendDataTo Index, Buffer.ToArray()
    
    DoEvents
    
    CloseSocket Index
End Sub

Public Sub SendLoginTokenToPlayer(Index As Long, loginToken As String)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SSetPlayerLoginToken
    Buffer.WriteString loginToken
    
    SendDataTo Index, Buffer.ToArray()
End Sub

Public Sub SendLoginTokenToGameServer(Username As String, loginToken As String)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong ASetPlayerLoginToken
    Buffer.WriteString Username
    Buffer.WriteString loginToken
    SendDataToGameServer Buffer.ToArray()
End Sub

Public Sub SendUsergroup(Username As String, usergroup As Long)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong ASetUsergroup
    Buffer.WriteString Username
    Buffer.WriteLong usergroup
    SendDataToGameServer Buffer.ToArray()
End Sub
