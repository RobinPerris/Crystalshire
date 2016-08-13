VERSION 5.00
Begin VB.Form frmEditor_Animation 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Animation Editor"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9975
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   9975
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   5760
      TabIndex        =   21
      Top             =   6840
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7440
      TabIndex        =   20
      Top             =   6840
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   4080
      TabIndex        =   19
      Top             =   6840
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Animation Properties"
      Height          =   6615
      Left            =   3360
      TabIndex        =   3
      Top             =   120
      Width           =   6495
      Begin VB.PictureBox picSprite 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2940
         Index           =   0
         Left            =   120
         ScaleHeight     =   196
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   196
         TabIndex        =   30
         Top             =   3480
         Width           =   2940
      End
      Begin VB.ComboBox cmbSound 
         Height          =   300
         Left            =   4080
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   240
         Width           =   2295
      End
      Begin VB.HScrollBar scrlLoopTime 
         Height          =   255
         Index           =   1
         Left            =   3360
         TabIndex        =   27
         Top             =   3120
         Width           =   2895
      End
      Begin VB.HScrollBar scrlLoopTime 
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   25
         Top             =   3120
         Width           =   3015
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   840
         TabIndex        =   23
         Top             =   240
         Width           =   2535
      End
      Begin VB.HScrollBar scrlFrameCount 
         Height          =   255
         Index           =   1
         Left            =   3360
         TabIndex        =   17
         Top             =   2520
         Width           =   2895
      End
      Begin VB.HScrollBar scrlLoopCount 
         Height          =   255
         Index           =   1
         Left            =   3360
         TabIndex        =   15
         Top             =   1920
         Width           =   2895
      End
      Begin VB.PictureBox picSprite 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2940
         Index           =   1
         Left            =   3360
         ScaleHeight     =   196
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   196
         TabIndex        =   13
         Top             =   3480
         Width           =   2940
      End
      Begin VB.HScrollBar scrlSprite 
         Height          =   255
         Index           =   1
         Left            =   3360
         TabIndex        =   12
         Top             =   1320
         Width           =   2895
      End
      Begin VB.HScrollBar scrlFrameCount 
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   2520
         Width           =   3015
      End
      Begin VB.HScrollBar scrlLoopCount 
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   1920
         Width           =   3015
      End
      Begin VB.HScrollBar scrlSprite 
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   1320
         Width           =   3015
      End
      Begin VB.Label Label3 
         Caption         =   "Sound:"
         Height          =   255
         Left            =   3480
         TabIndex        =   28
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblLoopTime 
         Caption         =   "Loop Time: 0"
         Height          =   255
         Index           =   1
         Left            =   3360
         TabIndex        =   26
         Top             =   2880
         Width           =   2655
      End
      Begin VB.Label lblLoopTime 
         Caption         =   "Loop Time: 0"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   24
         Top             =   2880
         Width           =   2655
      End
      Begin VB.Label Label1 
         Caption         =   "Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Layer 1 (Above Player)"
         Height          =   180
         Left            =   3360
         TabIndex        =   18
         Top             =   720
         Width           =   1740
      End
      Begin VB.Label lblFrameCount 
         AutoSize        =   -1  'True
         Caption         =   "Frame Count: 0"
         Height          =   180
         Index           =   1
         Left            =   3360
         TabIndex        =   16
         Top             =   2280
         Width           =   1170
      End
      Begin VB.Label lblLoopCount 
         AutoSize        =   -1  'True
         Caption         =   "Loop Count: 0"
         Height          =   180
         Index           =   1
         Left            =   3360
         TabIndex        =   14
         Top             =   1680
         Width           =   1065
      End
      Begin VB.Label lblSprite 
         AutoSize        =   -1  'True
         Caption         =   "Sprite: 0"
         Height          =   180
         Index           =   1
         Left            =   3360
         TabIndex        =   11
         Top             =   1080
         Width           =   660
      End
      Begin VB.Label lblFrameCount 
         AutoSize        =   -1  'True
         Caption         =   "Frame Count: 0"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   2280
         Width           =   1170
      End
      Begin VB.Label lblLoopCount 
         AutoSize        =   -1  'True
         Caption         =   "Loop Count: 0"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   1680
         Width           =   1065
      End
      Begin VB.Label lblSprite 
         AutoSize        =   -1  'True
         Caption         =   "Sprite: 0"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   660
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Layer 0 (Below Player)"
         Height          =   180
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   1695
      End
   End
   Begin VB.CommandButton cmdArray 
      Caption         =   "Change Array Size"
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   6840
      Width           =   2895
   End
   Begin VB.Frame Frame3 
      Caption         =   "Animation List"
      Height          =   6615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
      Begin VB.ListBox lstIndex 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6300
         ItemData        =   "frmEditor_Animation.frx":0000
         Left            =   120
         List            =   "frmEditor_Animation.frx":0002
         TabIndex        =   1
         Top             =   240
         Width           =   2895
      End
   End
End
Attribute VB_Name = "frmEditor_Animation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbSound_Click()

    If cmbSound.ListIndex >= 0 Then
        Animation(EditorIndex).sound = cmbSound.List(cmbSound.ListIndex)
    Else
        Animation(EditorIndex).sound = "None."
    End If

End Sub

Private Sub cmdCancel_Click()
    AnimationEditorCancel
End Sub

Private Sub cmdDelete_Click()
    Dim tmpIndex As Long

    If EditorIndex = 0 Or EditorIndex > MAX_ANIMATIONS Then Exit Sub
    ClearAnimation EditorIndex
    tmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Animation(EditorIndex).name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    AnimationEditorInit
End Sub

Private Sub cmdSave_Click()
    AnimationEditorOk
End Sub

Private Sub Form_Load()
    Dim i As Long

    For i = 0 To 1
        scrlSprite(i).Max = Count_Anim
        scrlLoopCount(i).Max = 100
        scrlFrameCount(i).Max = 100
        scrlLoopTime(i).Max = 1000
    Next

End Sub

Private Sub lstIndex_Click()
    AnimationEditorInit
End Sub

Private Sub scrlFrameCount_Change(index As Integer)
    lblFrameCount(index).caption = "Frame Count: " & scrlFrameCount(index).value
    Animation(EditorIndex).Frames(index) = scrlFrameCount(index).value
End Sub

Private Sub scrlFrameCount_Scroll(index As Integer)
    scrlFrameCount_Change index
End Sub

Private Sub scrlLoopCount_Change(index As Integer)
    lblLoopCount(index).caption = "Loop Count: " & scrlLoopCount(index).value
    Animation(EditorIndex).LoopCount(index) = scrlLoopCount(index).value
End Sub

Private Sub scrlLoopCount_Scroll(index As Integer)
    scrlLoopCount_Change index
End Sub

Private Sub scrlLoopTime_Change(index As Integer)
    lblLoopTime(index).caption = "Loop Time: " & scrlLoopTime(index).value
    Animation(EditorIndex).looptime(index) = scrlLoopTime(index).value
End Sub

Private Sub scrlLoopTime_Scroll(index As Integer)
    scrlLoopTime_Change index
End Sub

Private Sub scrlSprite_Change(index As Integer)
    lblSprite(index).caption = "Sprite: " & scrlSprite(index).value
    Animation(EditorIndex).sprite(index) = scrlSprite(index).value
End Sub

Private Sub scrlSprite_Scroll(index As Integer)
    scrlSprite_Change index
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
    Dim tmpIndex As Long

    If EditorIndex = 0 Or EditorIndex > MAX_ANIMATIONS Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    Animation(EditorIndex).name = Trim$(txtName.text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Animation(EditorIndex).name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
End Sub
