Attribute VB_Name = "modSound"
Option Explicit

' Hardcoded sound effects
Public Const Sound_ButtonHover As String = "Cursor1.wav"
Public Const Sound_ButtonClick As String = "Decision1.wav"

Public bInit_Sound As Boolean
Public lastButtonSound As Long
Public lastNpcChatsound As Long

Public Function Init_Sound() As Boolean
    On Error GoTo errorhandler
    
    ' exit out early if we have the system turned off
    If Options.sound = 0 Then Exit Function
    
    ' exit out early if we've already loaded
    If bInit_Sound Then Exit Function
    
    ' init sound engine
    
    ' return positive
    Init_Sound = True
    bInit_Sound = True
    Exit Function
    
errorhandler:
    Init_Sound = False
    bInit_Sound = False
End Function

Public Sub Destroy_Sound()
    If Not bInit_Sound Then Exit Sub
End Sub

Public Sub Play_Sound(ByVal sound As String)
    If Not bInit_Sound Then Exit Sub
End Sub

Public Sub Stop_Sound()
    If Not bInit_Sound Then Exit Sub
End Sub
