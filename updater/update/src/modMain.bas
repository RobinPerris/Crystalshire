Attribute VB_Name = "modMain"
Option Explicit

Private filePath As String

Sub Main()
    ' generate the file path
    filePath = Replace$(App.Path, "\bin", vbNullString)
    ' check if there's a new updater
    If FileExist(App.Path & "\crystalshire.dat") Then
        ' found the file - delete the normal updater then rename and move this one
        Do Until Delete(filePath & "\Crystalshire.exe")
            DoEvents
        Loop
        ' copy data file
        Delete App.Path & "\tmp.dat"
        Copy App.Path & "\crystalshire.dat", App.Path & "\tmp.dat"
        ' rename the tmp data file
        Rename App.Path & "\tmp.dat", App.Path & "\Crystalshire.exe"
        ' move it and kill the data file
        Copy App.Path & "\Crystalshire.exe", filePath & "\Crystalshire.exe"
        Delete App.Path & "\Crystalshire.exe"
    End If
    ' load updater and end
    If FileExist(filePath & "\Crystalshire.exe") Then Shell filePath & "\Crystalshire.exe", vbNormalFocus
    End
End Sub

Function Delete(theName As String) As Boolean
On Error GoTo errorhandler
    If FileExist(theName) Then Kill theName
    Delete = True
    Exit Function
errorhandler:
    Delete = False
End Function

Function Copy(oldName As String, newName As String)
On Error GoTo errorhandler
    FileCopy oldName, newName
    Copy = True
    Exit Function
errorhandler:
    Copy = False
End Function

Function Rename(oldName As String, newName As String) As Boolean
On Error GoTo errorhandler
    Name oldName As newName
    Rename = True
    Exit Function
errorhandler:
    Rename = False
End Function

Function FileExist(ByVal FileName As String) As Boolean
    If LenB(Dir$(FileName)) > 0 Then
        FileExist = True
    End If
End Function
