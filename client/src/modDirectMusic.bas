Attribute VB_Name = "modMusic"
Option Explicit

Public bInit_Music As Boolean

Public Function Init_Music() As Boolean
    On Error GoTo errorhandler
    
    ' exit out early if we have the system turned off
    If Options.Music = 0 Then Exit Function
    
    ' exit out early if we've already loaded
    If bInit_Music Then Exit Function
    
    ' init music engine
    If Not FSOUND_Init(44100, 32, 0) Then GoTo errorhandler
    
    ' return positive
    Init_Music = True
    bInit_Music = True
    Exit Function
    
errorhandler:
    Init_Music = False
    bInit_Music = False
End Function

Public Sub Destroy_Music()
    ' destroy music engine
    FSOUND_Close
End Sub

Public Sub Play_Music(ByVal song As String)
    If Not bInit_Music Then Exit Sub
End Sub

Public Sub Stop_Music()
    If Not bInit_Music Then Exit Sub
End Sub
