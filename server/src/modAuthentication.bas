Attribute VB_Name = "modAuthentication"
Option Explicit

Private Auth_Buffer As New clsBuffer
Private Auth_DataTimer As Long
Private Auth_DataBytes As Long
Private Auth_DataPackets As Long

Private Type LoginTokenRec
    user As String
    Token As String
    TimeCreated As Long
    Active As Boolean
End Type

Public LoginToken(1 To MAX_PLAYERS) As LoginTokenRec
Public Const LoginTimer As Long = 60000 ' 60 seconds

Private Function Auth_GetAddress(FunAddr As Long) As Long
    Auth_GetAddress = FunAddr
End Function

Public Sub Auth_InitMessages()
    Auth_HandleDataSub(ASetPlayerLoginToken) = Auth_GetAddress(AddressOf HandleSetPlayerLoginToken)
    Auth_HandleDataSub(ASetUsergroup) = Auth_GetAddress(AddressOf HandleSetUsergroup)
End Sub

Sub HandleSetPlayerLoginToken(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer, user As String, tLoginToken As String, i As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    user = Buffer.ReadString
    tLoginToken = Buffer.ReadString
    
    Set Buffer = Nothing
    
    ' find an inactive slot
    For i = 1 To MAX_PLAYERS
        If Not LoginToken(i).Active Then
            ' timed out
            LoginToken(i).user = user
            LoginToken(i).Token = tLoginToken
            LoginToken(i).TimeCreated = GetTickCount
            LoginToken(i).Active = True
        End If
    Next
End Sub

Sub HandleSetUsergroup(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer, user As String, Usergroup As Long, filename As String
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    user = Buffer.ReadString
    Usergroup = Buffer.ReadLong
    
    Set Buffer = Nothing
    
    ' find the file
    filename = App.Path & "\data\accounts\" & SanitiseString(Trim$(user)) & ".ini"
    If FileExist(filename, True) Then
        PutVar filename, "ACCOUNT", "Usergroup", STR(Usergroup)
    End If
End Sub

Sub Auth_HandleData(ByRef Data() As Byte)
Dim Buffer As clsBuffer
Dim MsgType As Long
        
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    MsgType = Buffer.ReadLong
    
    If MsgType < 0 Then
        Exit Sub
    End If
    
    If MsgType >= AMSG_COUNT Then
        Exit Sub
    End If
    
    CallWindowProc Auth_HandleDataSub(MsgType), 0, Buffer.ReadBytes(Buffer.Length), 0, 0
End Sub

Sub Auth_IncomingData(ByVal DataLength As Long)
Dim Buffer() As Byte
Dim pLength As Long

    ' Check if elapsed time has passed
    Auth_DataBytes = Auth_DataBytes + DataLength
    If GetTickCount >= Auth_DataTimer Then
        Auth_DataTimer = GetTickCount + 1000
        Auth_DataBytes = 0
        Auth_DataPackets = 0
    End If
    
    ' Get the data from the socket now
    frmServer.AuthSocket.GetData Buffer(), vbUnicode, DataLength
    Auth_Buffer.WriteBytes Buffer()
    
    If Auth_Buffer.Length >= 4 Then
        pLength = Auth_Buffer.ReadLong(False)
    
        If pLength < 0 Then
            Exit Sub
        End If
    End If
    
    Do While pLength > 0 And pLength <= Auth_Buffer.Length - 4
        If pLength <= Auth_Buffer.Length - 4 Then
            Auth_DataPackets = Auth_DataPackets + 1
            Auth_Buffer.ReadLong
            Auth_HandleData Auth_Buffer.ReadBytes(pLength)
        End If
        
        pLength = 0
        If Auth_Buffer.Length >= 4 Then
            pLength = Auth_Buffer.ReadLong(False)
        
            If pLength < 0 Then
                Exit Sub
            End If
        End If
    Loop
            
    Auth_Buffer.Trim
End Sub

Public Sub Auth_AcceptConnection(ByVal SocketId As Long)
    frmServer.AuthSocket.Close
    frmServer.AuthSocket.Accept SocketId
End Sub
