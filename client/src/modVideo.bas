Attribute VB_Name = "modVideo"
Option Explicit

Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

Public BasicAudio As IBasicAudio
Public BasicVideo As IBasicVideo
Public MediaEvent As IMediaEvent
Public MediaPosition As IMediaPosition
Public VideoWindow As IVideoWindow
Public MediaControl As IMediaControl
Public BasicVideo2 As IBasicVideo2

Public videoPlaying As Boolean

Public Sub VideoLoop()

    ' if it's finished then set it finished
    If MediaPosition.CurrentPosition >= MediaPosition.Duration - 1 Then
        ' stop the video playing
        videoPlaying = False
        
        ' the fade alpha
        fadeAlpha = 255
        
        ' set menu loop going
        frmMain.picIntro.visible = False
        
        ' exited out of playing video - shut down
        StopIntro
    End If
End Sub

Public Sub PlayIntro()
Dim handle As Long

    Exit Sub

    On Error GoTo errorhandler

    ' late binding
    Set MediaControl = New FilgraphManager

    ' set the size
    frmMain.picIntro.width = 800
    frmMain.picIntro.height = 600

    ' render the file
    MediaControl.RenderFile App.path & "\data files\video\intro.mp4"
    
    ' bind
    Set BasicAudio = MediaControl
    Set BasicVideo = MediaControl
    Set VideoWindow = MediaControl
    Set MediaPosition = MediaControl
    Set MediaEvent = MediaControl
    Set BasicVideo2 = MediaControl
    
    ' hack the window
    VideoWindow.WindowStyle = &H6000000
    handle = frmMain.picIntro.hWnd
    VideoWindow.Owner = handle
    
    ' turn off music if need be
    If Options.Music = False Then
        BasicAudio.volume = -10000
    Else
        BasicAudio.volume = 0
    End If
    
    ' resize
    VideoWindow.left = 0
    VideoWindow.top = 0
    VideoWindow.width = 800
    VideoWindow.height = 600
    
    ' run the video
    MediaControl.Run

    ' set the loop going
    videoPlaying = True
    VideoLoop
    
    Exit Sub
errorhandler:
    Exit Sub
End Sub

Public Sub StopIntro()
    If MediaControl Is Nothing Then Exit Sub
    
    MediaControl.Stop
    
    Set BasicAudio = Nothing
    Set BasicVideo = Nothing
    Set MediaEvent = Nothing
    Set MediaPosition = Nothing
    Set VideoWindow = Nothing
    Set MediaControl = Nothing
    Set BasicVideo2 = Nothing
    
    ' play the menu music
    If Len(Trim$(MenuMusic)) > 0 Then Play_Music Trim$(MenuMusic)
End Sub
