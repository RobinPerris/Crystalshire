VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GameName"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5445
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMain.frx":5F32
   ScaleHeight     =   375
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   363
   StartUpPosition =   2  'CenterScreen
   Begin Autoupdater.NetGrab NetGrab 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.Image imgPlay_Hover 
      Height          =   360
      Left            =   4560
      Picture         =   "frmMain.frx":69F10
      Top             =   6360
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Image imgPlay_Norm 
      Height          =   360
      Left            =   4560
      Picture         =   "frmMain.frx":6B332
      Top             =   6840
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Image imgRight 
      Height          =   105
      Left            =   2760
      Picture         =   "frmMain.frx":6C754
      Top             =   1200
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Image imgLeft 
      Height          =   105
      Left            =   2640
      Picture         =   "frmMain.frx":6C83E
      Top             =   1200
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label lblChanges 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2250
      Left            =   930
      TabIndex        =   4
      Top             =   1410
      Width           =   3600
   End
   Begin VB.Label lblHeader 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.5.0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2070
      TabIndex        =   3
      Top             =   1140
      Width           =   1305
   End
   Begin VB.Label lblTransfer 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   4770
      Width           =   3780
   End
   Begin VB.Label lblProgress2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Attempting connection."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   960
      TabIndex        =   1
      Top             =   4290
      Width           =   1965
   End
   Begin VB.Label lblProgress 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Connecting to server."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   960
      TabIndex        =   0
      Top             =   4050
      Width           =   2100
   End
   Begin VB.Image imgPlay 
      Height          =   360
      Left            =   3495
      Picture         =   "frmMain.frx":6C928
      Top             =   4110
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Image imgBar 
      Height          =   240
      Left            =   750
      Picture         =   "frmMain.frx":6DD4A
      Top             =   4740
      Width           =   3960
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private buttonHover As Boolean

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub imgLeft_Click()
    ShowUpdate curUpdate - 1
End Sub

Private Sub imgPlay_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Shell "bin\client.exe", vbNormalFocus
    End
End Sub

Private Sub imgPlay_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not buttonHover Then
        imgPlay.Picture = imgPlay_Hover.Picture
        buttonHover = True
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If buttonHover Then
        imgPlay.Picture = imgPlay_Norm.Picture
        buttonHover = False
    End If
End Sub

Private Sub imgRight_Click()
    ShowUpdate curUpdate + 1
End Sub

Private Sub NetGrab_downloadComplete(ByVal nBytes As Long)
    DownloadComplete = True
    If gettingCRC Then
        CRCDump = StrConv(NetGrab.Bytes, vbUnicode)
    Else
        currentBytes = currentBytes + tempBytes
        tempBytes = 0
    End If
End Sub

Private Sub NetGrab_DownloadFailed(ByVal ErrNum As Long, ByVal ErrDesc As String)
    failedDownload = True
    DownloadComplete = True
End Sub

Private Sub NetGrab_DownloadProgress(ByVal nBytes As Long)
    tempBytes = nBytes
End Sub
