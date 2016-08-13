Attribute VB_Name = "modMain"
Option Explicit

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long

Public Const GAMENAME As String = "Crystalshire"
Public Const GAMEURL As String = "http://www.crystalshire.com/updater/"

Public DownloadComplete As Boolean
Public failedDownload As Boolean
Public gettingCRC As Boolean
Public CRCDump As String

Public downloadBytes As Long, currentBytes As Long, tempBytes As Long

Public clientCRC() As String, serverCRC() As String
Public clientCount As Long, serverCount As Long
Public downloadFiles() As String, downloadCount As Long

Public maxBarWidth As Long

Public updateCount As Long
Public update() As UpdateUDT
Public curUpdate As Long
Type UpdateUDT
    header As String
    lineCount As Long
    strLine() As String
End Type

Sub Main()
Dim tmpString As String, i As Long, checkSum As String, strOffset As Long, fileSize As Long

    ' set the form
    frmMain.Caption = GAMENAME & " - v1.8.0"
    frmMain.Show
    
    ' set the bar width
    maxBarWidth = frmMain.imgBar.Width
    frmMain.imgBar.Width = 0
    
    ' load the changelog
    LoadChangeLog
    ShowUpdate updateCount
    
    ' download the update ini
    frmMain.NetGrab.DownloadStart GAMEURL & "files/bin/changelog.ini"
    ' loop around until the file is downloaded
    Do Until DownloadComplete
        Sleep 25
        DoEvents
    Loop
    DownloadComplete = False
    
    ' check if the file could download
    If failedDownload Then
        SetProgress "Connection failed.", "Try the game anyway."
        frmMain.NetGrab.DownloadCancel
        frmMain.imgPlay.Visible = True
        frmMain.imgBar.Width = maxBarWidth
        Exit Sub
    End If
    
    ' create folder if it doesn't exist
    ChkDir App.Path & "/", "bin"
    ' save file
    frmMain.NetGrab.SaveAs App.Path & "/bin/changelog.ini"
    
    ' we've downloaded
    SetProgress "Found update.", "Parsing data."
    
    ' load the changelog
    LoadChangeLog
    ShowUpdate updateCount
    
    ' update form
    DoEvents
    
    ' list all files
    tmpString = FindFiles(App.Path, App.Path & "\")
    clientCRC = Split(tmpString, ",")
    clientCount = UBound(clientCRC)
    
    ' set the file downloading whilst we calculate our own checksums
    gettingCRC = True
    frmMain.NetGrab.DownloadStart GAMEURL & "crc.txt"
    
    ' loop through and get the checksums + file sizes
    strOffset = Len(App.Path) + 1
    For i = 0 To clientCount
        If Len(clientCRC(i)) > 0 Then
            checkSum = GetFileCRC(clientCRC(i), fileSize)
            clientCRC(i) = Replace$(clientCRC(i), "\", "/")
            clientCRC(i) = Mid$(clientCRC(i), strOffset) & "," & checkSum & "," & fileSize
            Sleep 1
            DoEvents
        End If
    Next
    
    ' loop around until the file is downloaded
    Do Until DownloadComplete
        Sleep 25
        DoEvents
    Loop
    DownloadComplete = False

    ' no longer getting CRC
    gettingCRC = False
    
    ' split string
    serverCRC() = Split(CRCDump, "|")
    strOffset = 6
    ' loop through and trim
    serverCount = UBound(serverCRC)
    For i = 0 To serverCount
        serverCRC(i) = Mid$(serverCRC(i), strOffset)
    Next
    
    ' check for the update.exe
    If CheckSingleFile("update.exe") Then
        SetProgress "Updating updater.", "Downloading now."
        DownloadFile GAMEURL & "files/bin/update.exe", App.Path & "\bin\update.exe"
    End If
    
    ' check for new crystalshire.exe
    If CheckSingleFile("crystalshire.dat") Then
        SetProgress "Updating updater.", "Downloading now."
        DownloadFile GAMEURL & "files/bin/crystalshire.dat", App.Path & "\bin\crystalshire.dat"
        ' close down and let the update.exe do its work
        Shell "bin\update.exe", vbNormalFocus
        End
    End If
    
    ' compare file CRCs - make a list of files needed to download
    For i = 0 To serverCount
        CompareCRC i
    Next
    
    ' update
    If downloadCount > 0 Then
        SetProgress "Found updated files.", "Attempting download."
    Else
        SetProgress "No updates found.", "Enjoy the game!"
        frmMain.imgPlay.Visible = True
        frmMain.imgBar.Width = maxBarWidth
        Exit Sub
    End If
    
    ' start downloading the updated files!
    SetProgress "Found file update.", "Downloading now."
    DownloadUpdates
End Sub

Sub DownloadFile(URL As String, TargetFile As String)
    ' download it
    frmMain.NetGrab.DownloadStart URL
    ' loop around until the file is downloaded
    Do Until DownloadComplete
        Sleep 25
        DoEvents
    Loop
    DownloadComplete = False
    ' save it
    frmMain.NetGrab.SaveAs TargetFile
End Sub

Function CheckSingleFile(FileName As String) As Boolean
Dim updateString() As String, i As Long, clientStr() As String, serverStr() As String, x As Long

    For i = 0 To serverCount
        If Len(serverCRC(i)) > 0 Then
            updateString() = Split(serverCRC(i), ",")
            updateString() = Split(updateString(0), "/")
            If updateString(UBound(updateString)) = FileName Then
                ' make sure the file exists
                serverStr() = Split(serverCRC(i), ",")
                If Not FileExist(App.Path & serverStr(0)) Then
                    CheckSingleFile = True
                    Exit Function
                End If
                ' loop through and find matching client file
                For x = 0 To clientCount
                    If Len(clientCRC(x)) > 0 Then
                        clientStr() = Split(clientCRC(x), ",")
                        ' compare names
                        If clientStr(0) = serverStr(0) Then
                            ' compare checksums
                            If LCase$(clientStr(1)) <> LCase$(serverStr(1)) Then
                                CheckSingleFile = True
                                Exit Function
                            End If
                        End If
                    End If
                Next
            End If
        End If
    Next
    
    CheckSingleFile = False
End Function

Sub DownloadUpdates()
Dim i As Long, x As Long, updateString() As String, count As Long, localPath As String, builtPath As String
    ' Set the string
    UpdateProgressBar
    
    ' go through the files needed
    For i = 1 To downloadCount
        ' set it downloading
        frmMain.NetGrab.DownloadStart GAMEURL & "files" & downloadFiles(i)
        ' make all the directories we need
        updateString() = Split(downloadFiles(i), "/")
        count = UBound(updateString)
        builtPath = vbNullString
        For x = 1 To count - 1
            ChkDir App.Path & "\" & builtPath, updateString(x)
            builtPath = builtPath & updateString(x) & "\"
        Next
        ' update
        SetProgress "Downloading file.", updateString(count)
        ' update bar
        UpdateProgressBar
        ' loop through until the download completes
        DownloadComplete = False
        Do While Not DownloadComplete
            UpdateProgressBar
            DoEvents
        Loop
        ' download complete
        localPath = Replace$(downloadFiles(i), "/", "\")
        frmMain.NetGrab.SaveAs App.Path & localPath
    Next
    
    ' all downloads complete - max out bar if not already
    frmMain.imgBar.Width = maxBarWidth
    frmMain.lblTransfer = "Transfer completed."
    ' let them know
    SetProgress "Update complete.", "Enjoy the game!"
    frmMain.imgPlay.Visible = True
End Sub

Public Sub ChkDir(ByVal tDir As String, ByVal tName As String)
    If LCase$(Dir$(tDir & tName, vbDirectory)) <> tName Then Call MkDir(tDir & tName)
End Sub

Sub UpdateProgressBar()
Dim percent As Long, progressP As Long, value As Long, sString As String
    If downloadBytes = 0 Then Exit Sub
    ' label
    sString = GetByteString(currentBytes + tempBytes) & "/" & GetByteString(downloadBytes)
    If frmMain.lblTransfer.Caption <> sString Then frmMain.lblTransfer.Caption = sString
    ' bar
    value = ((currentBytes + tempBytes) / downloadBytes) * maxBarWidth
    With frmMain.imgBar
        If .Width <> value Then .Width = value
    End With
End Sub

Function GetByteString(Bytes As Long) As String
    If Bytes >= 1000000 Then
        GetByteString = Format$(Bytes / 1000000, "0.0") & "mB"
    Else
        GetByteString = Format$(Bytes / 1000, "0") & "kB"
    End If
End Function

Sub CompareCRC(serverIndex As Long)
Dim clientStr() As String, serverStr() As String, i As Long, updateString() As String, count As Long
    ' exit out early if the file doesn't exist
    If Len(serverCRC(serverIndex)) = 0 Then Exit Sub
    ' find the file path
    serverStr() = Split(serverCRC(serverIndex), ",")
    ' make sure the file exists
    If Not FileExist(App.Path & serverStr(0)) Then
        AddDownloadQueue serverStr(0), serverStr(2)
        Exit Sub
    End If
    ' update
    updateString() = Split(serverStr(0), "/")
    count = UBound(updateString)
    SetProgress "Comparing CRC32.", updateString(count)
    ' loop through and find matching client file
    For i = 0 To clientCount
        If Len(clientCRC(i)) > 0 Then
            clientStr() = Split(clientCRC(i), ",")
            ' compare names
            If clientStr(0) = serverStr(0) Then
                ' compare checksums
                If LCase$(clientStr(1)) <> LCase$(serverStr(1)) Then AddDownloadQueue serverStr(0), serverStr(2)
                Exit Sub
            End If
        End If
    Next
End Sub

Sub AddDownloadQueue(FileName As String, Bytes As String)
Dim index As Long
    downloadCount = downloadCount + 1
    ReDim Preserve downloadFiles(1 To downloadCount)
    index = UBound(downloadFiles)
    downloadFiles(index) = FileName
    ' add bytes to the max count
    downloadBytes = downloadBytes + Val(Bytes)
End Sub

Function ReadTxtFile(strPath As String) As String
    On Error GoTo ErrTrap
    Dim intFileNumber As Integer
   
    If Dir(strPath) = "" Then Exit Function
    intFileNumber = FreeFile
    Open strPath For Input As #intFileNumber
   
    ReadTxtFile = Input(LOF(intFileNumber), #intFileNumber)
ErrTrap:
    Close #intFileNumber
End Function

Sub SetProgress(string1 As String, string2 As String)
    frmMain.lblProgress.Caption = string1
    frmMain.lblProgress2.Caption = string2
End Sub

Function GetFileCRC(FileName As String, Optional ByRef fileSize As Long = 0) As String
    Dim cStream As New cBinaryFileStream
    Dim cCRC32 As New cCRC32
    Dim lCRC32 As Long
   
    cStream.File = FileName
    lCRC32 = cCRC32.GetFileCrc32(cStream)
    GetFileCRC = Hex(lCRC32)
    fileSize = cStream.Length
End Function

Function FindFiles(ByVal sSearchDir As String, ByVal Path As String) As String
Dim aList() As String
Dim nDir As Integer
Dim i As Integer
Dim s As String
Dim fn As String

If Right$(Path, 1) <> "\" Then Path = Path & "\"

fn = Dir$(Path & "*.*", vbDirectory)
While Len(fn)
    If (fn <> "..") And (fn <> ".") Then
        If GetAttr(Path & fn) And vbDirectory Then
            ReDim Preserve aList(nDir)
            aList(nDir) = fn
            nDir = nDir + 1
        Else
            s = s & Path & fn & ","
        End If
    End If
    fn = Dir$
Wend

For i = 0 To nDir - 1
    s = s & FindFiles(aList(i), Path & aList(i) & "\") & ","
Next i
FindFiles = s
Erase aList
End Function

Public Function FileExist(ByVal FileName As String) As Boolean
    If LenB(Dir$(FileName)) > 0 Then
        FileExist = True
    End If
End Function

Public Function GetVar(File As String, header As String, Var As String) As String
    Dim sSpaces As String   ' Max string length
    Dim szReturn As String  ' Return default value if not found
    szReturn = vbNullString
    sSpaces = Space$(5000)
    Call GetPrivateProfileString$(header, Var, szReturn, sSpaces, Len(sSpaces), File)
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
End Function

Sub LoadChangeLog()
Dim FileName As String, i As Long, x As Long
    FileName = App.Path & "\bin\changelog.ini"
    If Not FileExist(FileName) Then Exit Sub
    updateCount = GetVar(FileName, "General", "UpdateCount")
    ReDim update(1 To updateCount) As UpdateUDT
    For i = 1 To updateCount
        update(i).header = GetVar(FileName, "Update" & i, "Header")
        update(i).lineCount = GetVar(FileName, "Update" & i, "Lines")
        ReDim update(i).strLine(1 To update(i).lineCount) As String
        For x = 1 To update(i).lineCount
            update(i).strLine(x) = GetVar(FileName, "Update" & i, "String" & x)
        Next
    Next
End Sub

Sub ShowUpdate(ByVal updateNum As Long)
Dim i As Long
    frmMain.lblChanges.Caption = vbNullString
    frmMain.lblHeader.Caption = vbNullString
    If updateNum = 0 Then
        frmMain.lblHeader.Caption = "Version Unknown"
        frmMain.lblChanges.Caption = "Cannot find changelog."
        Exit Sub
    End If
    frmMain.lblHeader.Caption = update(updateNum).header
    For i = 1 To update(updateNum).lineCount
        frmMain.lblChanges.Caption = frmMain.lblChanges.Caption & update(updateNum).strLine(i) & vbNewLine
    Next
    curUpdate = updateNum
    With frmMain
        If curUpdate > 1 Then
            .imgLeft.Visible = True
            .imgLeft.Left = .lblHeader.Left - 12
        Else
            .imgLeft.Visible = False
        End If
        If curUpdate < updateCount Then
            .imgRight.Visible = True
            .imgRight.Left = .lblHeader.Left + .lblHeader.Width + 5
        Else
            .imgRight.Visible = False
        End If
    End With
End Sub
