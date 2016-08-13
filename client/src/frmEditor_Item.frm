VERSION 5.00
Begin VB.Form frmEditor_Item 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item Editor"
   ClientHeight    =   8415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9735
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
   Icon            =   "frmEditor_Item.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   561
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   649
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton cmdPaste 
      Caption         =   "Paste"
      Height          =   375
      Left            =   7320
      TabIndex        =   92
      Top             =   7920
      Width           =   735
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Copy"
      Height          =   375
      Left            =   6480
      TabIndex        =   91
      Top             =   7920
      Width           =   735
   End
   Begin VB.Frame fraEquipment 
      Caption         =   "Equipment Data"
      Height          =   3135
      Left            =   3360
      TabIndex        =   32
      Top             =   4680
      Visible         =   0   'False
      Width           =   6255
      Begin VB.HScrollBar scrlProf 
         Height          =   255
         Left            =   120
         Max             =   2
         TabIndex        =   90
         Top             =   2160
         Width           =   2055
      End
      Begin VB.PictureBox picPaperdoll 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
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
         Height          =   1080
         Left            =   3960
         ScaleHeight     =   72
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   144
         TabIndex        =   58
         Top             =   1920
         Width           =   2160
      End
      Begin VB.HScrollBar scrlPaperdoll 
         Height          =   255
         Left            =   5040
         TabIndex        =   57
         Top             =   1560
         Width           =   1095
      End
      Begin VB.HScrollBar scrlSpeed 
         Height          =   255
         LargeChange     =   100
         Left            =   4560
         Max             =   3000
         Min             =   100
         SmallChange     =   100
         TabIndex        =   40
         Top             =   840
         Value           =   100
         Width           =   1575
      End
      Begin VB.HScrollBar scrlStatBonus 
         Height          =   255
         Index           =   5
         LargeChange     =   10
         Left            =   3000
         Max             =   255
         TabIndex        =   39
         Top             =   1560
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatBonus 
         Height          =   255
         Index           =   4
         LargeChange     =   10
         Left            =   960
         Max             =   255
         TabIndex        =   38
         Top             =   1560
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatBonus 
         Height          =   255
         Index           =   3
         LargeChange     =   10
         Left            =   5280
         Max             =   255
         TabIndex        =   37
         Top             =   1200
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatBonus 
         Height          =   255
         Index           =   2
         LargeChange     =   10
         Left            =   3000
         Max             =   255
         TabIndex        =   36
         Top             =   1200
         Width           =   855
      End
      Begin VB.HScrollBar scrlDamage 
         Height          =   255
         LargeChange     =   10
         Left            =   1320
         Max             =   255
         TabIndex        =   35
         Top             =   840
         Width           =   1815
      End
      Begin VB.ComboBox cmbTool 
         Height          =   300
         ItemData        =   "frmEditor_Item.frx":3332
         Left            =   1320
         List            =   "frmEditor_Item.frx":3342
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   360
         Width           =   4815
      End
      Begin VB.HScrollBar scrlStatBonus 
         Height          =   255
         Index           =   1
         LargeChange     =   10
         Left            =   960
         Max             =   255
         TabIndex        =   33
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label lblProf 
         Caption         =   "Proficiency: None"
         Height          =   255
         Left            =   120
         TabIndex        =   89
         Top             =   1920
         Width           =   2055
      End
      Begin VB.Label lblPaperdoll 
         AutoSize        =   -1  'True
         Caption         =   "Paperdoll: 0"
         Height          =   180
         Left            =   3960
         TabIndex        =   56
         Top             =   1560
         Width           =   915
      End
      Begin VB.Label lblSpeed 
         AutoSize        =   -1  'True
         Caption         =   "Speed: 0.1 sec"
         Height          =   180
         Left            =   3240
         TabIndex        =   48
         Top             =   840
         UseMnemonic     =   0   'False
         Width           =   1140
      End
      Begin VB.Label lblStatBonus 
         AutoSize        =   -1  'True
         Caption         =   "+ Will: 0"
         Height          =   180
         Index           =   5
         Left            =   2160
         TabIndex        =   47
         Top             =   1560
         UseMnemonic     =   0   'False
         Width           =   630
      End
      Begin VB.Label lblStatBonus 
         AutoSize        =   -1  'True
         Caption         =   "+ Agi: 0"
         Height          =   180
         Index           =   4
         Left            =   120
         TabIndex        =   46
         Top             =   1560
         UseMnemonic     =   0   'False
         Width           =   615
      End
      Begin VB.Label lblStatBonus 
         AutoSize        =   -1  'True
         Caption         =   "+ Int: 0"
         Height          =   180
         Index           =   3
         Left            =   4440
         TabIndex        =   45
         Top             =   1200
         UseMnemonic     =   0   'False
         Width           =   585
      End
      Begin VB.Label lblStatBonus 
         AutoSize        =   -1  'True
         Caption         =   "+ End: 0"
         Height          =   180
         Index           =   2
         Left            =   2160
         TabIndex        =   44
         Top             =   1200
         UseMnemonic     =   0   'False
         Width           =   645
      End
      Begin VB.Label lblDamage 
         AutoSize        =   -1  'True
         Caption         =   "Damage: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   43
         Top             =   840
         UseMnemonic     =   0   'False
         Width           =   825
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Object Tool:"
         Height          =   180
         Left            =   120
         TabIndex        =   42
         Top             =   360
         Width           =   945
      End
      Begin VB.Label lblStatBonus 
         AutoSize        =   -1  'True
         Caption         =   "+ Str: 0"
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   41
         Top             =   1200
         UseMnemonic     =   0   'False
         Width           =   585
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Info"
      Height          =   3375
      Left            =   3360
      TabIndex        =   17
      Top             =   120
      Width           =   6255
      Begin VB.HScrollBar scrlLevelReq 
         Height          =   255
         LargeChange     =   10
         Left            =   4200
         Max             =   99
         TabIndex        =   74
         Top             =   2760
         Width           =   1935
      End
      Begin VB.HScrollBar scrlAccessReq 
         Height          =   255
         Left            =   4200
         Max             =   5
         TabIndex        =   72
         Top             =   2400
         Width           =   1935
      End
      Begin VB.ComboBox cmbClassReq 
         Height          =   300
         Left            =   3840
         Style           =   2  'Dropdown List
         TabIndex        =   70
         Top             =   2040
         Width           =   2295
      End
      Begin VB.ComboBox cmbSound 
         Height          =   300
         Left            =   3720
         Style           =   2  'Dropdown List
         TabIndex        =   69
         Top             =   1680
         Width           =   2415
      End
      Begin VB.TextBox txtDesc 
         Height          =   1455
         Left            =   120
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   60
         Top             =   1800
         Width           =   2655
      End
      Begin VB.HScrollBar scrlRarity 
         Height          =   255
         Left            =   4200
         Max             =   5
         TabIndex        =   25
         Top             =   960
         Width           =   1935
      End
      Begin VB.ComboBox cmbBind 
         Height          =   300
         ItemData        =   "frmEditor_Item.frx":3363
         Left            =   4200
         List            =   "frmEditor_Item.frx":3370
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   600
         Width           =   1935
      End
      Begin VB.HScrollBar scrlPrice 
         Height          =   255
         LargeChange     =   100
         Left            =   4200
         Max             =   30000
         TabIndex        =   23
         Top             =   240
         Width           =   1935
      End
      Begin VB.HScrollBar scrlAnim 
         Height          =   255
         Left            =   5040
         Max             =   5
         TabIndex        =   22
         Top             =   1320
         Width           =   1095
      End
      Begin VB.ComboBox cmbType 
         Height          =   300
         ItemData        =   "frmEditor_Item.frx":3399
         Left            =   120
         List            =   "frmEditor_Item.frx":33BE
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   1200
         Width           =   2655
      End
      Begin VB.TextBox txtName 
         Height          =   255
         Left            =   720
         TabIndex        =   20
         Top             =   240
         Width           =   2055
      End
      Begin VB.HScrollBar scrlPic 
         Height          =   255
         Left            =   840
         Max             =   255
         TabIndex        =   19
         Top             =   600
         Width           =   1335
      End
      Begin VB.PictureBox picItem 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
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
         Height          =   480
         Left            =   2280
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   18
         Top             =   600
         Width           =   480
      End
      Begin VB.Label lblLevelReq 
         AutoSize        =   -1  'True
         Caption         =   "Level req: 0"
         Height          =   180
         Left            =   2880
         TabIndex        =   75
         Top             =   2760
         Width           =   900
      End
      Begin VB.Label lblAccessReq 
         AutoSize        =   -1  'True
         Caption         =   "Access Req: 0"
         Height          =   180
         Left            =   2880
         TabIndex        =   73
         Top             =   2400
         Width           =   1110
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Class Req:"
         Height          =   180
         Left            =   2880
         TabIndex        =   71
         Top             =   2040
         Width           =   825
      End
      Begin VB.Label Label4 
         Caption         =   "Sound:"
         Height          =   255
         Left            =   2880
         TabIndex        =   68
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Description:"
         Height          =   255
         Left            =   120
         TabIndex        =   59
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label lblRarity 
         AutoSize        =   -1  'True
         Caption         =   "Rarity: 0"
         Height          =   180
         Left            =   2880
         TabIndex        =   31
         Top             =   960
         Width           =   660
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Bind Type:"
         Height          =   180
         Left            =   2880
         TabIndex        =   30
         Top             =   600
         Width           =   810
      End
      Begin VB.Label lblPrice 
         AutoSize        =   -1  'True
         Caption         =   "Price: 0"
         Height          =   180
         Left            =   2880
         TabIndex        =   29
         Top             =   240
         Width           =   600
      End
      Begin VB.Label lblAnim 
         AutoSize        =   -1  'True
         Caption         =   "Anim: None"
         Height          =   180
         Left            =   2880
         TabIndex        =   28
         Top             =   1320
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   180
         Left            =   120
         TabIndex        =   27
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   495
      End
      Begin VB.Label lblPic 
         AutoSize        =   -1  'True
         Caption         =   "Pic: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   26
         Top             =   600
         UseMnemonic     =   0   'False
         Width           =   450
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Requirements"
      Height          =   975
      Left            =   3360
      TabIndex        =   6
      Top             =   3600
      Width           =   6255
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   1
         LargeChange     =   10
         Left            =   720
         Max             =   255
         TabIndex        =   11
         Top             =   240
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   2
         LargeChange     =   10
         Left            =   2880
         Max             =   255
         TabIndex        =   10
         Top             =   240
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   3
         LargeChange     =   10
         Left            =   5160
         Max             =   255
         TabIndex        =   9
         Top             =   240
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   4
         LargeChange     =   10
         Left            =   720
         Max             =   255
         TabIndex        =   8
         Top             =   600
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   5
         LargeChange     =   10
         Left            =   2880
         Max             =   255
         TabIndex        =   7
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         Caption         =   "Str: 0"
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   16
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   435
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         Caption         =   "End: 0"
         Height          =   180
         Index           =   2
         Left            =   2280
         TabIndex        =   15
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   495
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         Caption         =   "Int: 0"
         Height          =   180
         Index           =   3
         Left            =   4560
         TabIndex        =   14
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   435
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         Caption         =   "Agi: 0"
         Height          =   180
         Index           =   4
         Left            =   120
         TabIndex        =   13
         Top             =   600
         UseMnemonic     =   0   'False
         Width           =   465
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         Caption         =   "Will: 0"
         Height          =   180
         Index           =   5
         Left            =   2280
         TabIndex        =   12
         Top             =   600
         UseMnemonic     =   0   'False
         Width           =   480
      End
   End
   Begin VB.CommandButton cmdArray 
      Caption         =   "Change Array Size"
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   7920
      Width           =   2895
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   7920
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   8160
      TabIndex        =   3
      Top             =   7920
      Width           =   1455
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   7920
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      Caption         =   "Item List"
      Height          =   7695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
      Begin VB.ListBox lstIndex 
         Height          =   7260
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.Frame fraUnique 
      Caption         =   "Unique"
      Height          =   615
      Left            =   3360
      TabIndex        =   76
      Top             =   4680
      Visible         =   0   'False
      Width           =   3735
      Begin VB.HScrollBar scrlUnique 
         Height          =   255
         Left            =   1080
         Max             =   255
         Min             =   1
         TabIndex        =   77
         Top             =   240
         Value           =   1
         Width           =   2415
      End
      Begin VB.Label lblUnique 
         AutoSize        =   -1  'True
         Caption         =   "Num: 0"
         Height          =   180
         Left            =   240
         TabIndex        =   78
         Top             =   240
         Width           =   555
      End
   End
   Begin VB.Frame fraSpell 
      Caption         =   "Spell Data"
      Height          =   1215
      Left            =   3360
      TabIndex        =   52
      Top             =   4680
      Visible         =   0   'False
      Width           =   3735
      Begin VB.HScrollBar scrlSpell 
         Height          =   255
         Left            =   1080
         Max             =   255
         Min             =   1
         TabIndex        =   53
         Top             =   720
         Value           =   1
         Width           =   2415
      End
      Begin VB.Label lblSpellName 
         AutoSize        =   -1  'True
         Caption         =   "Name: None"
         Height          =   180
         Left            =   240
         TabIndex        =   55
         Top             =   360
         Width           =   930
      End
      Begin VB.Label lblSpell 
         AutoSize        =   -1  'True
         Caption         =   "Num: 0"
         Height          =   180
         Left            =   240
         TabIndex        =   54
         Top             =   720
         Width           =   555
      End
   End
   Begin VB.Frame fraFood 
      Caption         =   "Food"
      Height          =   3135
      Left            =   3360
      TabIndex        =   79
      Top             =   4680
      Visible         =   0   'False
      Width           =   3735
      Begin VB.HScrollBar scrlFoodInterval 
         Height          =   255
         LargeChange     =   100
         Left            =   120
         Max             =   30000
         TabIndex        =   88
         Top             =   2280
         Width           =   3375
      End
      Begin VB.HScrollBar scrlFoodTick 
         Height          =   255
         Left            =   120
         TabIndex        =   86
         Top             =   1680
         Width           =   3375
      End
      Begin VB.HScrollBar scrlFoodHeal 
         Height          =   255
         Left            =   120
         TabIndex        =   84
         Top             =   1080
         Width           =   3375
      End
      Begin VB.OptionButton optSP 
         Caption         =   "SP"
         Height          =   255
         Left            =   840
         TabIndex        =   82
         Top             =   480
         Width           =   735
      End
      Begin VB.OptionButton optHP 
         Caption         =   "HP"
         Height          =   255
         Left            =   120
         TabIndex        =   81
         Top             =   480
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.Label lblFoodInterval 
         Caption         =   "Interval: 0(ms)"
         Height          =   255
         Left            =   120
         TabIndex        =   87
         Top             =   2040
         Width           =   3375
      End
      Begin VB.Label lblFoodTick 
         Caption         =   "Tick Count: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   85
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label lblFoodHeal 
         Caption         =   "Heal per Tick: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   83
         Top             =   840
         Width           =   3015
      End
      Begin VB.Label Label5 
         Caption         =   "Heals HP or SP"
         Height          =   255
         Left            =   120
         TabIndex        =   80
         Top             =   240
         Width           =   3495
      End
   End
   Begin VB.Frame fraVitals 
      Caption         =   "Consume Data"
      Height          =   3135
      Left            =   3360
      TabIndex        =   49
      Top             =   4680
      Visible         =   0   'False
      Width           =   3735
      Begin VB.CheckBox chkInstant 
         Caption         =   "Instant Cast?"
         Height          =   255
         Left            =   120
         TabIndex        =   67
         Top             =   2760
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.HScrollBar scrlCastSpell 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   65
         Top             =   2400
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.HScrollBar scrlAddExp 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   63
         Top             =   1800
         Width           =   3495
      End
      Begin VB.HScrollBar scrlAddMP 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   61
         Top             =   1200
         Width           =   3495
      End
      Begin VB.HScrollBar scrlAddHp 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   50
         Top             =   600
         Width           =   3495
      End
      Begin VB.Label lblCastSpell 
         AutoSize        =   -1  'True
         Caption         =   "Cast Spell: None"
         Height          =   180
         Left            =   120
         TabIndex        =   66
         Top             =   2160
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Label lblAddExp 
         AutoSize        =   -1  'True
         Caption         =   "Add Exp: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   64
         Top             =   1560
         UseMnemonic     =   0   'False
         Width           =   840
      End
      Begin VB.Label lblAddMP 
         AutoSize        =   -1  'True
         Caption         =   "Add MP: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   62
         Top             =   960
         UseMnemonic     =   0   'False
         Width           =   795
      End
      Begin VB.Label lblAddHP 
         AutoSize        =   -1  'True
         Caption         =   "Add HP: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   51
         Top             =   360
         UseMnemonic     =   0   'False
         Width           =   780
      End
   End
End
Attribute VB_Name = "frmEditor_Item"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbBind_Click()

    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    Item(EditorIndex).BindType = cmbBind.ListIndex
End Sub

Private Sub cmbClassReq_Click()

    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    Item(EditorIndex).ClassReq = cmbClassReq.ListIndex
End Sub

Private Sub cmbSound_Click()

    If cmbSound.ListIndex >= 0 Then
        Item(EditorIndex).sound = cmbSound.list(cmbSound.ListIndex)
    Else
        Item(EditorIndex).sound = "None."
    End If

End Sub

Private Sub cmbTool_Click()

    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    Item(EditorIndex).Data3 = cmbTool.ListIndex
End Sub

Private Sub cmdCopy_Click()
    ItemEditorCopy
End Sub

Private Sub cmdDelete_Click()
    Dim tmpIndex As Long

    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    ClearItem EditorIndex
    tmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Item(EditorIndex).name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    ItemEditorInit
End Sub

Private Sub cmdPaste_Click()
    ItemEditorPaste
End Sub

Private Sub Form_Load()
    scrlPic.Max = Count_Item
    scrlAnim.Max = MAX_ANIMATIONS
    scrlPaperdoll.Max = Count_Paperdoll
End Sub

Private Sub cmdSave_Click()
    Call ItemEditorOk
End Sub

Private Sub cmdCancel_Click()
    Call ItemEditorCancel
End Sub

Private Sub cmbType_Click()

    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    If (cmbType.ListIndex >= ITEM_TYPE_WEAPON) And (cmbType.ListIndex <= ITEM_TYPE_SHIELD) Then
        fraEquipment.visible = True
        'scrlDamage_Change
    Else
        fraEquipment.visible = False
    End If

    If cmbType.ListIndex = ITEM_TYPE_CONSUME Then
        fraVitals.visible = True
        'scrlVitalMod_Change
    Else
        fraVitals.visible = False
    End If

    If (cmbType.ListIndex = ITEM_TYPE_SPELL) Then
        fraSpell.visible = True
    Else
        fraSpell.visible = False
    End If

    If cmbType.ListIndex = ITEM_TYPE_UNIQUE Then
        fraUnique.visible = True
    Else
        fraUnique.visible = False
    End If

    If cmbType.ListIndex = ITEM_TYPE_FOOD Then
        fraFood.visible = True
    Else
        fraFood.visible = False
    End If

    Item(EditorIndex).Type = cmbType.ListIndex
End Sub

Private Sub lstIndex_Click()
    ItemEditorInit
End Sub

Private Sub optHP_Click()
    Item(EditorIndex).HPorSP = 1 ' hp
End Sub

Private Sub optSP_Click()
    Item(EditorIndex).HPorSP = 2 ' sp
End Sub

Private Sub scrlAccessReq_Change()

    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblAccessReq.caption = "Access Req: " & scrlAccessReq.value
    Item(EditorIndex).AccessReq = scrlAccessReq.value
End Sub

Private Sub scrlAddHp_Change()
    lblAddHP.caption = "Add HP: " & scrlAddHp.value
    Item(EditorIndex).AddHP = scrlAddHp.value
End Sub

Private Sub scrlAddMp_Change()
    lblAddMP.caption = "Add MP: " & scrlAddMP.value
    Item(EditorIndex).AddMP = scrlAddMP.value
End Sub

Private Sub scrlAddExp_Change()
    lblAddExp.caption = "Add Exp: " & scrlAddExp.value
    Item(EditorIndex).AddEXP = scrlAddExp.value
End Sub

Private Sub scrlAnim_Change()
    Dim sString As String

    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    If scrlAnim.value = 0 Then
        sString = "None"
    Else
        sString = Trim$(Animation(scrlAnim.value).name)
    End If

    lblAnim.caption = "Anim: " & sString
    Item(EditorIndex).Animation = scrlAnim.value
End Sub

Private Sub scrlDamage_Change()

    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblDamage.caption = "Damage: " & scrlDamage.value
    Item(EditorIndex).Data2 = scrlDamage.value
End Sub

Private Sub scrlFoodHeal_Change()
    lblFoodHeal.caption = "Heal Per Tick: " & scrlFoodHeal.value
    Item(EditorIndex).FoodPerTick = scrlFoodHeal.value
End Sub

Private Sub scrlFoodInterval_Change()
    lblFoodInterval.caption = "Interval: " & scrlFoodInterval.value & "(ms)"
    Item(EditorIndex).FoodInterval = scrlFoodInterval.value
End Sub

Private Sub scrlFoodTick_Change()
    lblFoodTick.caption = "Tick Count: " & scrlFoodTick.value
    Item(EditorIndex).FoodTickCount = scrlFoodTick.value
End Sub

Private Sub scrlLevelReq_Change()

    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblLevelReq.caption = "Level req: " & scrlLevelReq
    Item(EditorIndex).LevelReq = scrlLevelReq.value
End Sub

Private Sub scrlPaperdoll_Change()

    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblPaperdoll.caption = "Paperdoll: " & scrlPaperdoll.value
    Item(EditorIndex).Paperdoll = scrlPaperdoll.value
End Sub

Private Sub scrlPic_Change()

    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblPic.caption = "Pic: " & scrlPic.value
    Item(EditorIndex).Pic = scrlPic.value
End Sub

Private Sub scrlPrice_Change()

    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblPrice.caption = "Price: " & scrlPrice.value
    Item(EditorIndex).Price = scrlPrice.value
End Sub

Private Sub scrlProf_Change()
    Dim theProf As String

    Select Case scrlProf.value

        Case 0 ' None
            theProf = "None"

        Case 1 ' Sword/Armour
            theProf = "Sword/Armour"

        Case 2 ' Staff/Cloth
            theProf = "Staff/Cloth"
    End Select

    lblProf.caption = "Proficiency: " & theProf
    Item(EditorIndex).proficiency = scrlProf.value
End Sub

Private Sub scrlRarity_Change()

    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblRarity.caption = "Rarity: " & scrlRarity.value
    Item(EditorIndex).Rarity = scrlRarity.value
End Sub

Private Sub scrlSpeed_Change()

    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblSpeed.caption = "Speed: " & scrlSpeed.value / 1000 & " sec"
    Item(EditorIndex).speed = scrlSpeed.value
End Sub

Private Sub scrlStatBonus_Change(index As Integer)
    Dim text As String

    Select Case index

        Case 1
            text = "+ Str: "

        Case 2
            text = "+ End: "

        Case 3
            text = "+ Int: "

        Case 4
            text = "+ Agi: "

        Case 5
            text = "+ Will: "
    End Select

    lblStatBonus(index).caption = text & scrlStatBonus(index).value
    Item(EditorIndex).Add_Stat(index) = scrlStatBonus(index).value
End Sub

Private Sub scrlStatReq_Change(index As Integer)
    Dim text As String

    Select Case index

        Case 1
            text = "Str: "

        Case 2
            text = "End: "

        Case 3
            text = "Int: "

        Case 4
            text = "Agi: "

        Case 5
            text = "Will: "
    End Select

    lblStatReq(index).caption = text & scrlStatReq(index).value
    Item(EditorIndex).Stat_Req(index) = scrlStatReq(index).value
End Sub

Private Sub scrlSpell_Change()

    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    If Len(Trim$(Spell(scrlSpell.value).name)) > 0 Then
        lblSpellName.caption = "Name: " & Trim$(Spell(scrlSpell.value).name)
    Else
        lblSpellName.caption = "Name: None"
    End If

    lblSpell.caption = "Spell: " & scrlSpell.value
    Item(EditorIndex).Data1 = scrlSpell.value
End Sub

Private Sub scrlUnique_Change()
    lblUnique.caption = "Num: " & scrlUnique.value
    Item(EditorIndex).Data1 = scrlUnique.value
End Sub

Private Sub txtDesc_Change()

    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    Item(EditorIndex).Desc = txtDesc.text
End Sub

Public Sub txtName_Validate(Cancel As Boolean)
    Dim tmpIndex As Long

    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    Item(EditorIndex).name = Trim$(txtName.text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Item(EditorIndex).name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
End Sub
