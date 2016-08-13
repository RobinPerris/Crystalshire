Attribute VB_Name = "modMusic"
Option Explicit

' FMOD
Public Enum FSOUND_INITMODES
    FSOUND_INIT_USEDEFAULTMIDISYNTH = &H1
End Enum

Public Enum FSOUND_MODES
    FSOUND_LOOP_OFF = &H1
    FSOUND_LOOP_NORMAL = &H2
    FSOUND_16BITS = &H10
    FSOUND_MONO = &H20
    FSOUND_SIGNED = &H100
    FSOUND_NORMAL = FSOUND_16BITS Or FSOUND_SIGNED Or FSOUND_MONO
End Enum

Public Enum FSOUND_CHANNELSAMPLEMODE
    FSOUND_FREE = -1
    FSOUND_STEREOPAN = -1
End Enum

Public Declare Function FSOUND_Init Lib "fmod.dll" Alias "_FSOUND_Init@12" (ByVal mixrate As Long, ByVal maxchannels As Long, ByVal flags As FSOUND_INITMODES) As Byte
Public Declare Function FSOUND_Close Lib "fmod.dll" Alias "_FSOUND_Close@0" () As Long
Public Declare Function FMUSIC_LoadSong Lib "fmod.dll" Alias "_FMUSIC_LoadSong@4" (ByVal name As String) As Long
Public Declare Function FMUSIC_PlaySong Lib "fmod.dll" Alias "_FMUSIC_PlaySong@4" (ByVal module As Long) As Byte
Public Declare Function FMUSIC_SetMasterVolume Lib "fmod.dll" Alias "_FMUSIC_SetMasterVolume@8" (ByVal module As Long, ByVal volume As Long) As Byte
Public Declare Function FSOUND_Stream_Open Lib "fmod.dll" Alias "_FSOUND_Stream_Open@16" (ByVal filename As String, ByVal mode As FSOUND_MODES, ByVal offset As Long, ByVal length As Long) As Long
Public Declare Function FSOUND_Stream_Play Lib "fmod.dll" Alias "_FSOUND_Stream_Play@8" (ByVal channel As Long, ByVal stream As Long) As Long
Public Declare Function FSOUND_SetVolume Lib "fmod.dll" Alias "_FSOUND_SetVolume@8" (ByVal channel As Long, ByVal Vol As Long) As Byte
Public Declare Function FSOUND_Stream_Stop Lib "fmod.dll" Alias "_FSOUND_Stream_Stop@4" (ByVal stream As Long) As Byte
Public Declare Function FSOUND_Stream_Close Lib "fmod.dll" Alias "_FSOUND_Stream_Close@4" (ByVal stream As Long) As Byte
Public Declare Function FMUSIC_StopSong Lib "fmod.dll" Alias "_FMUSIC_StopSong@4" (ByVal module As Long) As Byte
Public Declare Function FMUSIC_FreeSong Lib "fmod.dll" Alias "_FMUSIC_FreeSong@4" (ByVal module As Long) As Byte
Public Declare Function FSOUND_Sample_SetDefaults Lib "fmod.dll" Alias "_FSOUND_Sample_SetDefaults@20" (ByVal sptr As Long, ByVal deffreq As Long, ByVal defvol As Long, ByVal defpan As Long, ByVal defpri As Long) As Byte
Public Declare Function FSOUND_PlaySound Lib "fmod.dll" Alias "_FSOUND_PlaySound@8" (ByVal channel As Long, ByVal sptr As Long) As Long
Public Declare Function FSOUND_Sample_Load Lib "fmod.dll" Alias "_FSOUND_Sample_Load@20" (ByVal index As Long, ByVal name As String, ByVal mode As FSOUND_MODES, ByVal offset As Long, ByVal length As Long) As Long
' Maximum sounds
Private Const MAX_SOUNDS = 32
' Hardcoded sound effects
Public Const Sound_ButtonHover As String = "Cursor1.wav"
Public Const Sound_ButtonClick As String = "Decision1.wav"
' Last sounds played
Public lastNpcChatsound As Long
' Init status
Public bInit_Music As Boolean
Public curSong As String
' Music Handlers
Private songHandle As Long
Private streamHandle As Long
' Sound pointer array
Private soundHandle(1 To MAX_SOUNDS) As Long
Private soundIndex As Long

Public Function Init_Music() As Boolean
    Dim result As Boolean

    If inDevelopment Then Exit Function

    On Error GoTo errorhandler

    ' init music engine
    result = FSOUND_Init(44100, 32, FSOUND_INIT_USEDEFAULTMIDISYNTH)

    If Not result Then GoTo errorhandler
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
    Stop_Music
    FSOUND_Close
    bInit_Music = False
    curSong = vbNullString
End Sub

Public Sub Play_Music(ByVal song As String)

    On Error GoTo errorhandler

    If Not bInit_Music Then Exit Sub

    ' exit out early if we have the system turned off
    If Options.Music = 0 Then Exit Sub

    ' does it exist?
    If Not FileExist(App.path & MUSIC_PATH & song) Then Exit Sub

    ' don't re-start currently playing songs
    If curSong = song Then Exit Sub
    ' stop the existing music
    Stop_Music

    ' find the extension
    Select Case Right$(song, 4)

        Case ".mid", ".s3m", ".mod"
            ' open the song
            songHandle = FMUSIC_LoadSong(App.path & MUSIC_PATH & song)
            ' play it
            FMUSIC_PlaySong songHandle
            ' set volume
            FMUSIC_SetMasterVolume songHandle, 150

        Case ".wav", ".mp3", ".ogg", ".wma"
            ' open the stream
            streamHandle = FSOUND_Stream_Open(App.path & MUSIC_PATH & song, FSOUND_LOOP_NORMAL, 0, 0)
            ' play it
            FSOUND_Stream_Play FSOUND_FREE, streamHandle
            ' set volume
            FSOUND_SetVolume streamHandle, 150

        Case Else
            Exit Sub
    End Select

    ' new current song
    curSong = song
    Exit Sub
errorhandler:
    Destroy_Music
End Sub

Public Sub Stop_Music()

    On Error GoTo errorhandler

    If Not streamHandle = 0 Then
        ' stop stream
        FSOUND_Stream_Stop streamHandle
        ' destroy
        FSOUND_Stream_Close streamHandle
        streamHandle = 0
    End If

    If Not songHandle = 0 Then
        ' stop song
        FMUSIC_StopSong songHandle
        ' destroy
        FMUSIC_FreeSong songHandle
        songHandle = 0
    End If

    ' no music
    curSong = vbNullString
    Exit Sub
errorhandler:
    Destroy_Music
End Sub

Public Sub Play_Sound(ByVal sound As String, Optional ByVal x As Long = -1, Optional ByVal y As Long = -1)
    Dim dX As Long, dY As Long, volume As Long, distance As Long

    On Error GoTo errorhandler

    If Not bInit_Music Then Exit Sub

    ' exit out early if we have the system turned off
    If Options.sound = 0 Then Exit Sub
    If x > -1 And y > -1 Then

        ' x
        If x < GetPlayerX(MyIndex) Then
            dX = GetPlayerX(MyIndex) - x
        ElseIf x > GetPlayerX(MyIndex) Then
            dX = x - GetPlayerX(MyIndex)
        End If

        ' y
        If y < GetPlayerY(MyIndex) Then
            dY = GetPlayerY(MyIndex) - y
        ElseIf y > GetPlayerY(MyIndex) Then
            dY = y - GetPlayerY(MyIndex)
        End If

        ' distance
        distance = dX ^ 2 + dY ^ 2
        volume = 150 - (distance / 2)
    Else
        volume = 150
    End If

    ' cap the volume
    If volume < 0 Then volume = 0
    If volume > 256 Then volume = 256
    ' load the sound
    Load_Sound sound
    FSOUND_Sample_SetDefaults soundHandle(soundIndex), -1, volume, FSOUND_STEREOPAN, -1
    ' play it
    FSOUND_PlaySound FSOUND_FREE, soundHandle(soundIndex)
    Exit Sub
errorhandler:
    Destroy_Music
End Sub

Public Sub Load_Sound(ByVal sound As String)
    Dim bRestart As Boolean

    On Error GoTo errorhandler

    ' next sound buffer
    soundIndex = soundIndex + 1

    ' reset if we run out
    If soundIndex > MAX_SOUNDS Or soundIndex < 1 Then
        bRestart = True
        soundIndex = 1
    End If

    ' load the sound
    soundHandle(soundIndex) = FSOUND_Sample_Load(FSOUND_FREE, App.path & SOUND_PATH & sound, FSOUND_NORMAL, 0, 0)
    Exit Sub
errorhandler:
    Destroy_Music
End Sub
