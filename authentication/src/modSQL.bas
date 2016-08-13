Attribute VB_Name = "modSQL"
Option Explicit

'this is for connection and recordset
Public CN As New ADODB.Connection
Public RS_USER As New ADODB.Recordset

'variables to connect to mysqlserver
Public strServer As String
Public strUsername As String
Public strPassword As String
Public strPort As String
Public strDatabase As String
Public strSQL As String
Public strCOMMAND As ADODB.Command
Public xlsFilename As Variant

Public Function ConnectToSqlServer() As Boolean
On Error GoTo errhandler

    'Set CN = Nothing

    strServer = "SERVER"
    strUsername = "USER"
    strPassword = "PASSWORD"
    strPort = "PORT"
    strDatabase = "DBNAME"
    
    Set CN = New ADODB.Connection
    CN.CursorLocation = adUseClient
    CN.ConnectionString = "Driver={MySQL ODBC 3.51 Driver};Server=" & strServer & ";PORT=" & strPort & ";Database=" & strDatabase & ";Uid=" & strUsername & ";Pwd=" & strPassword & ";"
    CN.Open
    
    ConnectToSqlServer = True
    Exit Function
errhandler:
    'Debug.Print Err.Number, Err.Description
    ConnectToSqlServer = False
End Function

Function GetUser(Username As String, password As String) As ADODB.Recordset
Dim oRS As ADODB.Recordset
Dim SqlStr As String
Dim rsPassword As String
Dim rsPasswordSalt As String

    On Error GoTo errorhandler
    
    ' DblChcek if connected to sql
    If CN.State <> adStateOpen Then
        ' Try to reconnect
        If Not ConnectToSqlServer Then Exit Function
    End If
    
    'Define your SQL String
    SqlStr = "Select * From mybb_users Where username = '" & Username & "'"
    
    'Open the recordset
    Set oRS = New ADODB.Recordset
    oRS.Open SqlStr, CN, adOpenStatic, adLockReadOnly
    Set oRS.ActiveConnection = Nothing
    
    Set GetUser = oRS.Clone
    
    'Destroy ADO objects
    oRS.Close
    Set oRS = Nothing
    Exit Function

errorhandler:
    
End Function

Sub DBSetStatus()
Dim sqlcmd As String

    ' We want to only set the server status if both auth and game server are online
    ' as such we do a server connection check first.
    
    If Not ConnectToGameServer Then
        ' We can't connect to game server - make it so we try more often til it's back up
        frmMain.tmrStatus.Interval = 5000
        Debug.Print "Cannot connect to game server - Will not send status to database."
        Exit Sub
    End If
    
    ' If we can now connect to the game server then we want to increase the call
    frmMain.tmrStatus.Interval = 60000

    SimpleQuery "SET TIME_ZONE='+00:00'"
    SimpleQuery "delete from crystalshire.mybb_serverstatus where servername = 'london'"
    SimpleQuery "insert into mybb_serverstatus (servername, connected) values ('london', Now())"
End Sub

Function SimpleQuery(SqlStr) As Boolean
Dim cmd As New ADODB.Command
Dim oRS As ADODB.Recordset

    ' default to failed
    SimpleQuery = False

    ' DblChcek if connected to sql
    If CN.State <> adStateOpen Then
        ' Try to reconnect
        If Not ConnectToSqlServer Then Exit Function
    End If
    
    ' Build the command
    With cmd
        .ActiveConnection = CN
        .CommandText = SqlStr
        .CommandType = adCmdText
    End With
    
    Debug.Print cmd.CommandText
    
    ' Query
    Set oRS = New ADODB.Recordset
    oRS.Open cmd.CommandText, CN
    Set oRS.ActiveConnection = Nothing
    Set oRS = Nothing
End Function

Public Sub DisconnectFromSqlServer()
    CN.Close
    Set CN = Nothing
End Sub
