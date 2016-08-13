VERSION 5.00
Begin VB.Form frmEditor_Spell 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Spell Editor"
   ClientHeight    =   8055
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10335
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
   ScaleHeight     =   537
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   689
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Copy"
      Height          =   375
      Left            =   6480
      TabIndex        =   64
      Top             =   7560
      Width           =   975
   End
   Begin VB.CommandButton cmdPaste 
      Caption         =   "Paste"
      Height          =   375
      Left            =   7560
      TabIndex        =   63
      Top             =   7560
      Width           =   975
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   3360
      TabIndex        =   6
      Top             =   7560
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   8640
      TabIndex        =   5
      Top             =   7560
      Width           =   1575
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   4920
      TabIndex        =   4
      Top             =   7560
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Spell Properties"
      Height          =   7335
      Left            =   3360
      TabIndex        =   3
      Top             =   120
      Width           =   6855
      Begin VB.HScrollBar scrlUses 
         Height          =   255
         LargeChange     =   10
         Left            =   5280
         TabIndex        =   62
         Top             =   6960
         Width           =   1455
      End
      Begin VB.HScrollBar scrlNext 
         Height          =   255
         Left            =   5280
         TabIndex        =   60
         Top             =   6600
         Width           =   1455
      End
      Begin VB.HScrollBar scrlIndex 
         Height          =   255
         Left            =   5280
         TabIndex        =   58
         Top             =   6240
         Width           =   1455
      End
      Begin VB.ComboBox cmbSound 
         Height          =   300
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   56
         Top             =   6840
         Width           =   1215
      End
      Begin VB.TextBox txtDesc 
         Height          =   975
         Left            =   1440
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   54
         Top             =   6240
         Width           =   975
      End
      Begin VB.Frame Frame6 
         Caption         =   "Data"
         Height          =   5895
         Left            =   3480
         TabIndex        =   14
         Top             =   240
         Width           =   3255
         Begin VB.HScrollBar scrlStun 
            Height          =   255
            Left            =   120
            TabIndex        =   51
            Top             =   5520
            Width           =   2895
         End
         Begin VB.HScrollBar scrlAnim 
            Height          =   255
            Left            =   120
            TabIndex        =   49
            Top             =   4920
            Width           =   2895
         End
         Begin VB.HScrollBar scrlAnimCast 
            Height          =   255
            Left            =   120
            TabIndex        =   47
            Top             =   4320
            Width           =   2895
         End
         Begin VB.CheckBox chkAOE 
            Caption         =   "Area of Effect spell?"
            Height          =   255
            Left            =   120
            TabIndex        =   43
            Top             =   3240
            Width           =   3015
         End
         Begin VB.HScrollBar scrlAOE 
            Height          =   255
            Left            =   120
            TabIndex        =   42
            Top             =   3720
            Width           =   3015
         End
         Begin VB.HScrollBar scrlRange 
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   2880
            Width           =   3015
         End
         Begin VB.HScrollBar scrlInterval 
            Height          =   255
            Left            =   1680
            Max             =   60
            TabIndex        =   38
            Top             =   2280
            Width           =   1455
         End
         Begin VB.HScrollBar scrlDuration 
            Height          =   255
            Left            =   120
            Max             =   60
            TabIndex        =   36
            Top             =   2280
            Width           =   1455
         End
         Begin VB.HScrollBar scrlVital 
            Height          =   255
            LargeChange     =   10
            Left            =   120
            Max             =   1000
            TabIndex        =   34
            Top             =   1680
            Width           =   3015
         End
         Begin VB.HScrollBar scrlDir 
            Height          =   255
            Left            =   1680
            TabIndex        =   22
            Top             =   480
            Width           =   1455
         End
         Begin VB.HScrollBar scrlY 
            Height          =   255
            Left            =   1680
            TabIndex        =   20
            Top             =   1080
            Width           =   1455
         End
         Begin VB.HScrollBar scrlX 
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   1080
            Width           =   1455
         End
         Begin VB.HScrollBar scrlMap 
            Height          =   255
            Left            =   120
            Max             =   100
            TabIndex        =   16
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label lblStun 
            Caption         =   "Stun Duration: None"
            Height          =   255
            Left            =   120
            TabIndex        =   52
            Top             =   5280
            Width           =   2895
         End
         Begin VB.Label lblAnim 
            Caption         =   "Animation: None"
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   4680
            Width           =   2895
         End
         Begin VB.Label lblAnimCast 
            Caption         =   "Cast Anim: None"
            Height          =   255
            Left            =   120
            TabIndex        =   46
            Top             =   4080
            Width           =   2895
         End
         Begin VB.Label lblAOE 
            Caption         =   "AoE: Self-cast"
            Height          =   255
            Left            =   120
            TabIndex        =   41
            Top             =   3480
            Width           =   3015
         End
         Begin VB.Label lblRange 
            Caption         =   "Range: Self-cast"
            Height          =   255
            Left            =   120
            TabIndex        =   39
            Top             =   2640
            Width           =   3015
         End
         Begin VB.Label lblInterval 
            Caption         =   "Interval: 0s"
            Height          =   255
            Left            =   1680
            TabIndex        =   37
            Top             =   2040
            Width           =   1455
         End
         Begin VB.Label lblDuration 
            Caption         =   "Duration: 0s"
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   2040
            Width           =   1455
         End
         Begin VB.Label lblVital 
            Caption         =   "Vital: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   1440
            Width           =   3015
         End
         Begin VB.Label lblDir 
            Caption         =   "Dir: Down"
            Height          =   255
            Left            =   1680
            TabIndex        =   21
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label lblY 
            Caption         =   "Y: 0"
            Height          =   255
            Left            =   1680
            TabIndex        =   19
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label lblX 
            Caption         =   "X: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label lblMap 
            Caption         =   "Map: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Basic Information"
         Height          =   5895
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   3255
         Begin VB.PictureBox picSprite 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   2640
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   50
            Top             =   5160
            Width           =   480
         End
         Begin VB.HScrollBar scrlIcon 
            Height          =   255
            Left            =   120
            TabIndex        =   45
            Top             =   5400
            Width           =   2415
         End
         Begin VB.HScrollBar scrlCool 
            Height          =   255
            Left            =   120
            Max             =   60
            TabIndex        =   32
            Top             =   4680
            Width           =   3015
         End
         Begin VB.HScrollBar scrlCast 
            Height          =   255
            Left            =   120
            Max             =   60
            TabIndex        =   30
            Top             =   4080
            Width           =   3015
         End
         Begin VB.ComboBox cmbClass 
            Height          =   300
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   3480
            Width           =   3015
         End
         Begin VB.HScrollBar scrlAccess 
            Height          =   255
            Left            =   120
            Max             =   5
            TabIndex        =   26
            Top             =   2880
            Width           =   3015
         End
         Begin VB.HScrollBar scrlLevel 
            Height          =   255
            Left            =   120
            Max             =   100
            TabIndex        =   24
            Top             =   2280
            Width           =   3015
         End
         Begin VB.HScrollBar scrlMP 
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   1680
            Width           =   3015
         End
         Begin VB.ComboBox cmbType 
            Height          =   300
            ItemData        =   "frmEditor_Spell.frx":0000
            Left            =   120
            List            =   "frmEditor_Spell.frx":0013
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   1080
            Width           =   3015
         End
         Begin VB.TextBox txtName 
            Height          =   270
            Left            =   120
            TabIndex        =   9
            Top             =   480
            Width           =   3015
         End
         Begin VB.Label lblIcon 
            Caption         =   "Icon: None"
            Height          =   255
            Left            =   120
            TabIndex        =   44
            Top             =   5160
            Width           =   3015
         End
         Begin VB.Label lblCool 
            Caption         =   "Cooldown Time: 0s"
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   4440
            Width           =   2535
         End
         Begin VB.Label lblCast 
            Caption         =   "Casting Time: 0s"
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   3840
            Width           =   1695
         End
         Begin VB.Label Label5 
            Caption         =   "Class Required:"
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   3240
            Width           =   1815
         End
         Begin VB.Label lblAccess 
            Caption         =   "Access Required: None"
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   2640
            Width           =   1815
         End
         Begin VB.Label lblLevel 
            Caption         =   "Level Required: None"
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   2040
            Width           =   1815
         End
         Begin VB.Label lblMP 
            Caption         =   "MP Cost: None"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   1440
            Width           =   1815
         End
         Begin VB.Label Label2 
            Caption         =   "Type:"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   840
            Width           =   1815
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Name:"
            Height          =   180
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.Label lblUses 
         Caption         =   "Uses: 0"
         Height          =   255
         Left            =   2520
         TabIndex        =   61
         Top             =   6960
         Width           =   2655
      End
      Begin VB.Label lblNext 
         Caption         =   "Next: None"
         Height          =   255
         Left            =   2520
         TabIndex        =   59
         Top             =   6600
         Width           =   2655
      End
      Begin VB.Label lblIndex 
         Caption         =   "Unique Index: 0"
         Height          =   255
         Left            =   2520
         TabIndex        =   57
         Top             =   6240
         Width           =   2655
      End
      Begin VB.Label Label4 
         Caption         =   "Sound:"
         Height          =   255
         Left            =   120
         TabIndex        =   55
         Top             =   6600
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Description:"
         Height          =   255
         Left            =   120
         TabIndex        =   53
         Top             =   6240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Spell List"
      Height          =   7335
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3135
      Begin VB.ListBox lstIndex 
         Height          =   6900
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.CommandButton cmdArray 
      Caption         =   "Change Array Size"
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   7560
      Width           =   2895
   End
End
Attribute VB_Name = "frmEditor_Spell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkAOE_Click()

    If chkAOE.value = 0 Then
        Spell(EditorIndex).IsAoE = False
    Else
        Spell(EditorIndex).IsAoE = True
    End If

End Sub

Private Sub cmbClass_Click()
    Spell(EditorIndex).ClassReq = cmbClass.ListIndex
End Sub

Private Sub cmbType_Click()
    Spell(EditorIndex).Type = cmbType.ListIndex
End Sub

Private Sub cmdCopy_Click()
    SpellEditorCopy
End Sub

Private Sub cmdDelete_Click()
    Dim tmpIndex As Long
    ClearSpell EditorIndex
    tmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Spell(EditorIndex).name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    SpellEditorInit
End Sub

Private Sub cmdPaste_Click()
    SpellEditorPaste
End Sub

Private Sub cmdSave_Click()
    SpellEditorOk
End Sub

Private Sub lstIndex_Click()
    SpellEditorInit
End Sub

Private Sub cmdCancel_Click()
    SpellEditorCancel
End Sub

Private Sub scrlAccess_Change()

    If scrlAccess.value > 0 Then
        lblAccess.caption = "Access Required: " & scrlAccess.value
    Else
        lblAccess.caption = "Access Required: None"
    End If

    Spell(EditorIndex).AccessReq = scrlAccess.value
End Sub

Private Sub scrlAnim_Change()

    If scrlAnim.value > 0 Then
        lblAnim.caption = "Animation: " & Trim$(Animation(scrlAnim.value).name)
    Else
        lblAnim.caption = "Animation: None"
    End If

    Spell(EditorIndex).SpellAnim = scrlAnim.value
End Sub

Private Sub scrlAnimCast_Change()

    If scrlAnimCast.value > 0 Then
        lblAnimCast.caption = "Cast Anim: " & Trim$(Animation(scrlAnimCast.value).name)
    Else
        lblAnimCast.caption = "Cast Anim: None"
    End If

    Spell(EditorIndex).CastAnim = scrlAnimCast.value
End Sub

Private Sub scrlAOE_Change()

    If scrlAOE.value > 0 Then
        lblAOE.caption = "AoE: " & scrlAOE.value & " tiles."
    Else
        lblAOE.caption = "AoE: Self-cast"
    End If

    Spell(EditorIndex).AoE = scrlAOE.value
End Sub

Private Sub scrlCast_Change()
    lblCast.caption = "Casting Time: " & scrlCast.value & "s"
    Spell(EditorIndex).CastTime = scrlCast.value
End Sub

Private Sub scrlCool_Change()
    lblCool.caption = "Cooldown Time: " & scrlCool.value & "s"
    Spell(EditorIndex).CDTime = scrlCool.value
End Sub

Private Sub scrlDir_Change()
    Dim sDir As String

    Select Case scrlDir.value

        Case DIR_UP
            sDir = "Up"

        Case DIR_DOWN
            sDir = "Down"

        Case DIR_RIGHT
            sDir = "Right"

        Case DIR_LEFT
            sDir = "Left"
    End Select

    lblDir.caption = "Dir: " & sDir
    Spell(EditorIndex).dir = scrlDir.value
End Sub

Private Sub scrlDuration_Change()
    lblDuration.caption = "Duration: " & scrlDuration.value & "s"
    Spell(EditorIndex).Duration = scrlDuration.value
End Sub

Private Sub scrlIcon_Change()

    If scrlIcon.value > 0 Then
        lblIcon.caption = "Icon: " & scrlIcon.value
    Else
        lblIcon.caption = "Icon: None"
    End If

    Spell(EditorIndex).icon = scrlIcon.value
End Sub

Private Sub scrlIndex_Change()
    lblIndex.caption = "Unique Index: " & scrlIndex.value
    Spell(EditorIndex).UniqueIndex = scrlIndex.value
End Sub

Private Sub scrlInterval_Change()
    lblInterval.caption = "Interval: " & scrlInterval.value & "s"
    Spell(EditorIndex).Interval = scrlInterval.value
End Sub

Private Sub scrlLevel_Change()

    If scrlLevel.value > 0 Then
        lblLevel.caption = "Level Required: " & scrlLevel.value
    Else
        lblLevel.caption = "Level Required: None"
    End If

    Spell(EditorIndex).LevelReq = scrlLevel.value
End Sub

Private Sub scrlMap_Change()
    lblMap.caption = "Map: " & scrlMap.value
    Spell(EditorIndex).map = scrlMap.value
End Sub

Private Sub scrlMP_Change()

    If scrlMP.value > 0 Then
        lblMP.caption = "MP Cost: " & scrlMP.value
    Else
        lblMP.caption = "MP Cost: None"
    End If

    Spell(EditorIndex).MPCost = scrlMP.value
End Sub

Private Sub scrlNext_Change()

    If scrlNext.value > 0 Then
        lblNext.caption = "Next: " & scrlNext.value & " - " & Trim$(Spell(scrlNext.value).name)
    Else
        lblNext.caption = "Next: None"
    End If

    Spell(EditorIndex).NextRank = scrlNext.value
End Sub

Private Sub scrlRange_Change()

    If scrlRange.value > 0 Then
        lblRange.caption = "Range: " & scrlRange.value & " tiles."
    Else
        lblRange.caption = "Range: Self-cast"
    End If

    Spell(EditorIndex).Range = scrlRange.value
End Sub

Private Sub scrlStun_Change()

    If scrlStun.value > 0 Then
        lblStun.caption = "Stun Duration: " & scrlStun.value & "s"
    Else
        lblStun.caption = "Stun Duration: None"
    End If

    Spell(EditorIndex).StunDuration = scrlStun.value
End Sub

Private Sub scrlUses_Change()
    lblUses.caption = "Uses: " & scrlUses.value
    Spell(EditorIndex).NextUses = scrlUses.value
End Sub

Private Sub scrlVital_Change()
    lblVital.caption = "Vital: " & scrlVital.value
    Spell(EditorIndex).Vital = scrlVital.value
End Sub

Private Sub scrlX_Change()
    lblX.caption = "X: " & scrlX.value
    Spell(EditorIndex).x = scrlX.value
End Sub

Private Sub scrlY_Change()
    lblY.caption = "Y: " & scrlY.value
    Spell(EditorIndex).y = scrlY.value
End Sub

Private Sub txtDesc_Change()
    Spell(EditorIndex).Desc = txtDesc.text
End Sub

Public Sub txtName_Validate(Cancel As Boolean)
    Dim tmpIndex As Long

    If EditorIndex = 0 Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    Spell(EditorIndex).name = Trim$(txtName.text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Spell(EditorIndex).name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
End Sub

Private Sub cmbSound_Click()

    If cmbSound.ListIndex >= 0 Then
        Spell(EditorIndex).sound = cmbSound.list(cmbSound.ListIndex)
    Else
        Spell(EditorIndex).sound = "None."
    End If

End Sub
