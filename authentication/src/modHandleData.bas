Attribute VB_Name = "modHandleData"
Option Explicit

Public Function GetAddress(FunAddr As Long) As Long
    GetAddress = FunAddr
End Function

Public Sub InitMessages()
    HandleDataSub(CAuthLogin) = GetAddress(AddressOf HandleLogin)
End Sub

Sub HandleData(ByVal Index As Long, ByRef Data() As Byte)
Dim Buffer As clsBuffer, MsgType As Long, packetCallback As Long
            
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    MsgType = Buffer.ReadLong
    
    If (MsgType < 0) Or (MsgType >= CMSG_COUNT) Then
        HackingAttempt Index
        Exit Sub
    End If
    
    packetCallback = HandleDataSub(MsgType)
    If packetCallback <> 0 Then
        CallWindowProc HandleDataSub(MsgType), Index, Buffer.ReadBytes(Buffer.Length), 0, 0
    End If
End Sub

Private Sub HandleLogin(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim Username As String, password As String
Dim loginToken As String
Dim vMAJOR As Long, vMINOR As Long, vREVISION As Long

Dim userInfo As ADODB.Recordset
Dim actPass As String, actSalt As String, actUserID As String, tmpPass As String, actUserGroup As String

    On Error GoTo errorhandler
   
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    Username = Buffer.ReadString
    password = Buffer.ReadString
    vMAJOR = Buffer.ReadLong
    vMINOR = Buffer.ReadLong
    vREVISION = Buffer.ReadLong
    
    ' right version
    If vMAJOR <> CLIENT_MAJOR Or vMINOR <> CLIENT_MINOR Or vREVISION <> CLIENT_REVISION Then
        SendAlertMsg Index, DIALOGUE_MSG_OUTDATED
        Exit Sub
    End If
        
    If Len(Username) < 3 Or Len(password) < 3 Then
        SendAlertMsg Index, DIALOGUE_MSG_USERLENGTH, MENU_LOGIN
        Exit Sub
    End If
    
    ' try and connect to database
    If Not ConnectToSqlServer Then
        SendAlertMsg Index, DIALOGUE_MSG_MYSQL
        Exit Sub
    End If
    
    ' get the recordset
    Set userInfo = GetUser(Username, password)
    
    If Not userInfo.EOF Then
        ' username found
        actUserID = userInfo.Fields("uid").value
        actPass = userInfo.Fields("password").value
        actSalt = userInfo.Fields("salt").value
        
        ' check password
        tmpPass = SaltPassword(MD5(password), actSalt)
        If tmpPass <> actPass Then
            SendAlertMsg Index, DIALOGUE_MSG_WRONGPASS, MENU_LOGIN
            Exit Sub
        End If
        
        ' check usergroups
        actUserGroup = userInfo.Fields("usergroup").value
        
        ' Need activating
        Select Case actUserGroup
            Case 5
                SendAlertMsg Index, DIALOGUE_MSG_ACTIVATED, MENU_LOGIN
                Exit Sub
            Case 7
                SendAlertMsg Index, DIALOGUE_MSG_BANNED
                Exit Sub
        End Select
    Else
        SendAlertMsg Index, DIALOGUE_MSG_WRONGPASS, MENU_LOGIN
        Exit Sub
    End If
    
    ' Send new account group to the game server
    SendUsergroup Username, Val(actUserGroup)
    
    ' Everything passed, create the token and send it off
    loginToken = RandomString("AN-##AA-ANHHAN-H")
    
    SendLoginTokenToGameServer Username, loginToken
    SendLoginTokenToPlayer Index, loginToken
    
    DoEvents
    CloseSocket Index
    
errorhandler:
    Exit Sub
End Sub
