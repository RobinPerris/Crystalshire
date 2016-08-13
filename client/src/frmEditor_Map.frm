VERSION 5.00
Begin VB.Form frmEditor_Map 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Map Editor"
   ClientHeight    =   9255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   19470
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
   Icon            =   "frmEditor_Map.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   617
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1298
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   255
      Left            =   6480
      TabIndex        =   107
      Top             =   8160
      Width           =   1335
   End
   Begin VB.PictureBox picAttributes 
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
      Height          =   9495
      Left            =   9600
      ScaleHeight     =   9495
      ScaleWidth      =   9735
      TabIndex        =   29
      Top             =   0
      Visible         =   0   'False
      Width           =   9735
      Begin VB.Frame fraAppear 
         Caption         =   "Appear"
         Height          =   1815
         Left            =   3480
         TabIndex        =   110
         Top             =   3240
         Visible         =   0   'False
         Width           =   3375
         Begin VB.HScrollBar scrlAppearRange 
            Height          =   255
            Left            =   120
            Max             =   10
            TabIndex        =   113
            Top             =   480
            Value           =   1
            Width           =   3135
         End
         Begin VB.CommandButton cmdAppearOkay 
            Caption         =   "Accept"
            Height          =   375
            Left            =   1080
            TabIndex        =   112
            Top             =   1320
            Width           =   1215
         End
         Begin VB.CheckBox chkAppearBottom 
            Caption         =   "Bottom?"
            Height          =   255
            Left            =   2280
            TabIndex        =   111
            Top             =   840
            Width           =   975
         End
         Begin VB.Label lblAppearRange 
            Caption         =   "Range: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   114
            Top             =   240
            Width           =   3135
         End
      End
      Begin VB.Frame fraMapWarp 
         Caption         =   "Map Warp"
         Height          =   3015
         Left            =   3480
         TabIndex        =   59
         Top             =   2640
         Visible         =   0   'False
         Width           =   3375
         Begin VB.CheckBox chkWarpFall 
            Caption         =   "Fall?"
            Height          =   255
            Left            =   2520
            TabIndex        =   109
            Top             =   2040
            Width           =   735
         End
         Begin VB.CommandButton cmdMapWarp 
            Caption         =   "Accept"
            Height          =   375
            Left            =   1080
            TabIndex        =   66
            Top             =   2400
            Width           =   1215
         End
         Begin VB.HScrollBar scrlMapWarpY 
            Height          =   255
            Left            =   120
            TabIndex        =   65
            Top             =   1680
            Width           =   3135
         End
         Begin VB.HScrollBar scrlMapWarpX 
            Height          =   255
            Left            =   120
            TabIndex        =   63
            Top             =   1080
            Width           =   3135
         End
         Begin VB.HScrollBar scrlMapWarp 
            Height          =   255
            Left            =   120
            Min             =   1
            TabIndex        =   61
            Top             =   480
            Value           =   1
            Width           =   3135
         End
         Begin VB.Label lblMapWarpY 
            Caption         =   "Y: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   64
            Top             =   1440
            Width           =   3135
         End
         Begin VB.Label lblMapWarpX 
            Caption         =   "X: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   62
            Top             =   840
            Width           =   3135
         End
         Begin VB.Label lblMapWarp 
            Caption         =   "Map: 1"
            Height          =   255
            Left            =   120
            TabIndex        =   60
            Top             =   240
            Width           =   3135
         End
      End
      Begin VB.Frame fraMapItem 
         Caption         =   "Map Item"
         Height          =   1815
         Left            =   3480
         TabIndex        =   41
         Top             =   3480
         Visible         =   0   'False
         Width           =   3375
         Begin VB.CommandButton cmdMapItem 
            Caption         =   "Accept"
            Height          =   375
            Left            =   1200
            TabIndex        =   46
            Top             =   1200
            Width           =   1215
         End
         Begin VB.HScrollBar scrlMapItemValue 
            Height          =   255
            Left            =   120
            Min             =   1
            TabIndex        =   45
            Top             =   840
            Value           =   1
            Width           =   2535
         End
         Begin VB.HScrollBar scrlMapItem 
            Height          =   255
            Left            =   120
            Max             =   10
            Min             =   1
            TabIndex        =   44
            Top             =   480
            Value           =   1
            Width           =   2535
         End
         Begin VB.PictureBox picMapItem 
            Appearance      =   0  'Flat
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
            ForeColor       =   &H80000008&
            Height          =   480
            Left            =   2760
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   43
            Top             =   600
            Width           =   480
         End
         Begin VB.Label lblMapItem 
            Caption         =   "Item: None x0"
            Height          =   255
            Left            =   120
            TabIndex        =   42
            Top             =   240
            Width           =   3135
         End
      End
      Begin VB.Frame fraNudge 
         Caption         =   "Nudge"
         Height          =   2775
         Left            =   3480
         TabIndex        =   99
         Top             =   3000
         Visible         =   0   'False
         Width           =   3255
         Begin VB.CommandButton cmdNDone 
            Caption         =   "Done"
            Height          =   375
            Left            =   1080
            TabIndex        =   104
            Top             =   2280
            Width           =   1215
         End
         Begin VB.CommandButton cmdNDown 
            Caption         =   "Down"
            Height          =   375
            Left            =   1200
            TabIndex        =   103
            Top             =   1440
            Width           =   855
         End
         Begin VB.CommandButton cmdNLeft 
            Caption         =   "Left"
            Height          =   375
            Left            =   360
            TabIndex        =   102
            Top             =   960
            Width           =   855
         End
         Begin VB.CommandButton cmdNRight 
            Caption         =   "Right"
            Height          =   375
            Left            =   2040
            TabIndex        =   101
            Top             =   960
            Width           =   855
         End
         Begin VB.CommandButton cmdNUp 
            Caption         =   "Up"
            Height          =   375
            Left            =   1200
            TabIndex        =   100
            Top             =   480
            Width           =   855
         End
      End
      Begin VB.Frame fraChat 
         Caption         =   "Chat"
         Height          =   2655
         Left            =   3480
         TabIndex        =   90
         Top             =   3120
         Visible         =   0   'False
         Width           =   3375
         Begin VB.ListBox lstChat 
            Height          =   780
            Left            =   240
            TabIndex        =   93
            Top             =   360
            Width           =   2895
         End
         Begin VB.HScrollBar scrlChat 
            Height          =   255
            Left            =   240
            Max             =   3
            TabIndex        =   92
            Top             =   1560
            Width           =   2895
         End
         Begin VB.CommandButton cmdChat 
            Caption         =   "Okay"
            Height          =   375
            Left            =   960
            TabIndex        =   91
            Top             =   2040
            Width           =   1455
         End
         Begin VB.Label lblChat 
            Caption         =   "Direction Req: Up"
            Height          =   255
            Left            =   240
            TabIndex        =   94
            Top             =   1320
            Width           =   2535
         End
      End
      Begin VB.Frame fraTrap 
         Caption         =   "Trap"
         Height          =   1575
         Left            =   3480
         TabIndex        =   82
         Top             =   3480
         Visible         =   0   'False
         Width           =   3375
         Begin VB.HScrollBar scrlTrap 
            Height          =   255
            Left            =   240
            Max             =   10000
            TabIndex        =   84
            Top             =   600
            Width           =   2895
         End
         Begin VB.CommandButton cmdTrap 
            Caption         =   "Okay"
            Height          =   375
            Left            =   960
            TabIndex        =   83
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label lblTrap 
            Caption         =   "Amount: 0"
            Height          =   255
            Left            =   240
            TabIndex        =   85
            Top             =   360
            Width           =   2535
         End
      End
      Begin VB.Frame fraSlide 
         Caption         =   "Slide"
         Height          =   1455
         Left            =   3480
         TabIndex        =   86
         Top             =   3480
         Visible         =   0   'False
         Width           =   3375
         Begin VB.ComboBox cmbSlide 
            Height          =   300
            ItemData        =   "frmEditor_Map.frx":3332
            Left            =   240
            List            =   "frmEditor_Map.frx":3342
            Style           =   2  'Dropdown List
            TabIndex        =   88
            Top             =   360
            Width           =   2895
         End
         Begin VB.CommandButton cmdSlide 
            Caption         =   "Okay"
            Height          =   375
            Left            =   960
            TabIndex        =   87
            Top             =   840
            Width           =   1455
         End
      End
      Begin VB.Frame fraHeal 
         Caption         =   "Heal"
         Height          =   1815
         Left            =   3480
         TabIndex        =   77
         Top             =   3480
         Visible         =   0   'False
         Width           =   3375
         Begin VB.ComboBox cmbHeal 
            Height          =   300
            ItemData        =   "frmEditor_Map.frx":335D
            Left            =   240
            List            =   "frmEditor_Map.frx":3367
            Style           =   2  'Dropdown List
            TabIndex        =   81
            Top             =   240
            Width           =   2895
         End
         Begin VB.CommandButton cmdHeal 
            Caption         =   "Okay"
            Height          =   375
            Left            =   960
            TabIndex        =   79
            Top             =   1200
            Width           =   1455
         End
         Begin VB.HScrollBar scrlHeal 
            Height          =   255
            Left            =   240
            Max             =   10000
            TabIndex        =   78
            Top             =   840
            Width           =   2895
         End
         Begin VB.Label lblHeal 
            Caption         =   "Amount: 0"
            Height          =   255
            Left            =   240
            TabIndex        =   80
            Top             =   600
            Width           =   2535
         End
      End
      Begin VB.Frame fraNpcSpawn 
         Caption         =   "Npc Spawn"
         Height          =   2655
         Left            =   3480
         TabIndex        =   36
         Top             =   3120
         Visible         =   0   'False
         Width           =   3375
         Begin VB.ListBox lstNpc 
            Height          =   780
            Left            =   240
            TabIndex        =   40
            Top             =   360
            Width           =   2895
         End
         Begin VB.HScrollBar scrlNpcDir 
            Height          =   255
            Left            =   240
            Max             =   3
            TabIndex        =   38
            Top             =   1560
            Width           =   2895
         End
         Begin VB.CommandButton cmdNpcSpawn 
            Caption         =   "Okay"
            Height          =   375
            Left            =   960
            TabIndex        =   37
            Top             =   2040
            Width           =   1455
         End
         Begin VB.Label lblNpcDir 
            Caption         =   "Direction: Up"
            Height          =   255
            Left            =   240
            TabIndex        =   39
            Top             =   1320
            Width           =   2535
         End
      End
      Begin VB.Frame fraResource 
         Caption         =   "Object"
         Height          =   1695
         Left            =   3480
         TabIndex        =   30
         Top             =   3480
         Visible         =   0   'False
         Width           =   3375
         Begin VB.CommandButton cmdResourceOk 
            Caption         =   "Okay"
            Height          =   375
            Left            =   960
            TabIndex        =   33
            Top             =   1080
            Width           =   1455
         End
         Begin VB.HScrollBar scrlResource 
            Height          =   255
            Left            =   240
            Max             =   100
            Min             =   1
            TabIndex        =   32
            Top             =   600
            Value           =   1
            Width           =   2895
         End
         Begin VB.Label lblResource 
            Caption         =   "Object:"
            Height          =   255
            Left            =   240
            TabIndex        =   31
            Top             =   360
            Width           =   2535
         End
      End
      Begin VB.Frame fraShop 
         Caption         =   "Shop"
         Height          =   1335
         Left            =   3480
         TabIndex        =   67
         Top             =   3720
         Visible         =   0   'False
         Width           =   3135
         Begin VB.CommandButton cmdShop 
            Caption         =   "Accept"
            Height          =   375
            Left            =   960
            TabIndex        =   69
            Top             =   720
            Width           =   1215
         End
         Begin VB.ComboBox cmbShop 
            Height          =   300
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   68
            Top             =   240
            Width           =   2895
         End
      End
      Begin VB.Frame fraKeyOpen 
         Caption         =   "Key Open"
         Height          =   2295
         Left            =   3480
         TabIndex        =   53
         Top             =   3240
         Visible         =   0   'False
         Width           =   3375
         Begin VB.CommandButton cmdKeyOpen 
            Caption         =   "Accept"
            Height          =   375
            Left            =   1080
            TabIndex        =   58
            Top             =   1680
            Width           =   1215
         End
         Begin VB.HScrollBar scrlKeyY 
            Height          =   255
            Left            =   120
            TabIndex        =   57
            Top             =   1080
            Width           =   3015
         End
         Begin VB.HScrollBar scrlKeyX 
            Height          =   255
            Left            =   120
            TabIndex        =   55
            Top             =   480
            Width           =   3015
         End
         Begin VB.Label lblKeyY 
            Caption         =   "Y: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   56
            Top             =   840
            Width           =   3015
         End
         Begin VB.Label lblKeyX 
            Caption         =   "X: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   54
            Top             =   240
            Width           =   3015
         End
      End
      Begin VB.Frame fraMapKey 
         Caption         =   "Map Key"
         Height          =   2655
         Left            =   3480
         TabIndex        =   47
         Top             =   3120
         Visible         =   0   'False
         Width           =   3375
         Begin VB.HScrollBar scrlKeyTime 
            Height          =   255
            Left            =   120
            Max             =   120
            Min             =   -1
            SmallChange     =   10
            TabIndex        =   105
            Top             =   1560
            Width           =   3015
         End
         Begin VB.PictureBox picMapKey 
            Appearance      =   0  'Flat
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
            ForeColor       =   &H80000008&
            Height          =   480
            Left            =   2760
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   52
            Top             =   600
            Width           =   480
         End
         Begin VB.CommandButton cmdMapKey 
            Caption         =   "Accept"
            Height          =   375
            Left            =   1080
            TabIndex        =   51
            Top             =   2040
            Width           =   1215
         End
         Begin VB.CheckBox chkMapKey 
            Caption         =   "Take key away upon use."
            Height          =   255
            Left            =   120
            TabIndex        =   50
            Top             =   960
            Value           =   1  'Checked
            Width           =   2535
         End
         Begin VB.HScrollBar scrlMapKey 
            Height          =   255
            Left            =   120
            Max             =   5
            Min             =   1
            TabIndex        =   49
            Top             =   600
            Value           =   1
            Width           =   2535
         End
         Begin VB.Label lblKeyTime 
            Caption         =   "Time: 0s"
            Height          =   255
            Left            =   120
            TabIndex        =   106
            Top             =   1320
            Width           =   3015
         End
         Begin VB.Label lblMapKey 
            Caption         =   "Item: None"
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   240
            Width           =   3135
         End
      End
   End
   Begin VB.CommandButton cmdNudge 
      Caption         =   "Nudge"
      Height          =   255
      Left            =   5040
      TabIndex        =   98
      Top             =   8520
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   6480
      TabIndex        =   10
      Top             =   8880
      Width           =   1335
   End
   Begin VB.CommandButton cmdProperties 
      Caption         =   "Properties"
      Height          =   255
      Left            =   6480
      TabIndex        =   12
      Top             =   8520
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "Type"
      Height          =   1335
      Left            =   7920
      TabIndex        =   25
      Top             =   7800
      Width           =   1455
      Begin VB.OptionButton optEvents 
         Alignment       =   1  'Right Justify
         Caption         =   "Events"
         Height          =   255
         Left            =   360
         TabIndex        =   97
         Top             =   960
         Width           =   855
      End
      Begin VB.OptionButton optBlock 
         Alignment       =   1  'Right Justify
         Caption         =   "Block"
         Height          =   255
         Left            =   480
         TabIndex        =   72
         Top             =   720
         Width           =   735
      End
      Begin VB.OptionButton optAttribs 
         Alignment       =   1  'Right Justify
         Caption         =   "Attributes"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   480
         Width           =   1095
      End
      Begin VB.OptionButton optLayers 
         Alignment       =   1  'Right Justify
         Caption         =   "Layers"
         Height          =   255
         Left            =   360
         TabIndex        =   26
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.HScrollBar scrlPictureX 
      Height          =   255
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox picBack 
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
      Height          =   7680
      Left            =   120
      ScaleHeight     =   512
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   512
      TabIndex        =   14
      Top             =   120
      Width           =   7680
      Begin VB.PictureBox picBackSelect 
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
         Height          =   960
         Left            =   0
         ScaleHeight     =   64
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   64
         TabIndex        =   15
         Top             =   0
         Width           =   960
      End
   End
   Begin VB.VScrollBar scrlPictureY 
      Height          =   375
      Left            =   0
      Max             =   255
      TabIndex        =   13
      Top             =   240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Frame fraTileSet 
      Caption         =   "Tileset: 0"
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   8160
      Width           =   4815
      Begin VB.HScrollBar scrlTileSet 
         Height          =   255
         Left            =   120
         Max             =   10
         Min             =   1
         TabIndex        =   1
         Top             =   360
         Value           =   1
         Width           =   4455
      End
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Okay"
      Height          =   255
      Left            =   5040
      TabIndex        =   11
      Top             =   8160
      Width           =   1335
   End
   Begin VB.Frame fraAttribs 
      Caption         =   "Attributes"
      Height          =   7575
      Left            =   7920
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   1455
      Begin VB.OptionButton optAppear 
         Caption         =   "Appear"
         Height          =   270
         Left            =   120
         TabIndex        =   108
         Top             =   3840
         Width           =   1215
      End
      Begin VB.OptionButton optChat 
         Caption         =   "Chat"
         Height          =   270
         Left            =   120
         TabIndex        =   89
         Top             =   3600
         Width           =   1215
      End
      Begin VB.OptionButton optSlide 
         Caption         =   "Slide"
         Height          =   270
         Left            =   120
         TabIndex        =   76
         Top             =   3360
         Width           =   1215
      End
      Begin VB.OptionButton optTrap 
         Caption         =   "Trap"
         Height          =   270
         Left            =   120
         TabIndex        =   75
         Top             =   3120
         Width           =   1215
      End
      Begin VB.OptionButton optHeal 
         Caption         =   "Heal"
         Height          =   270
         Left            =   120
         TabIndex        =   74
         Top             =   2880
         Width           =   1215
      End
      Begin VB.OptionButton optBank 
         Caption         =   "Bank"
         Height          =   270
         Left            =   120
         TabIndex        =   73
         Top             =   2640
         Width           =   1215
      End
      Begin VB.OptionButton optShop 
         Caption         =   "Shop"
         Height          =   270
         Left            =   120
         TabIndex        =   70
         Top             =   2400
         Width           =   1215
      End
      Begin VB.OptionButton optNpcSpawn 
         Caption         =   "Npc Spawn"
         Height          =   270
         Left            =   120
         TabIndex        =   35
         Top             =   2160
         Width           =   1215
      End
      Begin VB.OptionButton optDoor 
         Caption         =   "Door"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   1920
         Width           =   1215
      End
      Begin VB.OptionButton optResource 
         Caption         =   "Resource"
         Height          =   240
         Left            =   120
         TabIndex        =   28
         Top             =   1680
         Width           =   1215
      End
      Begin VB.OptionButton optKeyOpen 
         Caption         =   "Key Open"
         Height          =   240
         Left            =   120
         TabIndex        =   9
         Top             =   1440
         Width           =   1215
      End
      Begin VB.OptionButton optBlocked 
         Caption         =   "Blocked"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton optWarp 
         Caption         =   "Warp"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton cmdClear2 
         Caption         =   "Clear"
         Height          =   390
         Left            =   120
         TabIndex        =   6
         Top             =   7080
         Width           =   1215
      End
      Begin VB.OptionButton optItem 
         Caption         =   "Item"
         Height          =   270
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton optNpcAvoid 
         Caption         =   "Npc Avoid"
         Height          =   270
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   1215
      End
      Begin VB.OptionButton optKey 
         Caption         =   "Key"
         Height          =   270
         Left            =   120
         TabIndex        =   3
         Top             =   1200
         Width           =   1215
      End
   End
   Begin VB.Frame fraLayers 
      Caption         =   "Layers"
      Height          =   7575
      Left            =   7920
      TabIndex        =   16
      Top             =   120
      Width           =   1455
      Begin VB.HScrollBar scrlAutotile 
         Height          =   255
         Left            =   120
         Max             =   5
         TabIndex        =   95
         Top             =   6240
         Width           =   1215
      End
      Begin VB.CommandButton cmdFill 
         Caption         =   "Fill"
         Height          =   390
         Left            =   120
         TabIndex        =   20
         Top             =   7080
         Width           =   1215
      End
      Begin VB.OptionButton optLayer 
         Caption         =   "Fringe"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   23
         Top             =   960
         Width           =   1215
      End
      Begin VB.OptionButton optLayer 
         Caption         =   "Mask"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   22
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton optLayer 
         Caption         =   "Ground"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton optLayer 
         Caption         =   "Fringe2"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   19
         Top             =   1200
         Width           =   1215
      End
      Begin VB.OptionButton optLayer 
         Caption         =   "Mask2"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   18
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   6600
         Width           =   1215
      End
      Begin VB.Label lblAutotile 
         Alignment       =   2  'Center
         Caption         =   "Normal"
         Height          =   255
         Left            =   120
         TabIndex        =   96
         Top             =   6000
         Width           =   1215
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Drag mouse to select multiple tiles"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   71
      Top             =   7800
      Width           =   7695
   End
End
Attribute VB_Name = "frmEditor_Map"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAppearOkay_Click()
    EditorAppearRange = scrlAppearRange.value
    EditorAppearBottom = chkAppearBottom.value
    picAttributes.visible = False
    fraAppear.visible = False
End Sub

Private Sub cmdApply_Click()
    applyingMap = True
    SendMap
End Sub

Private Sub cmdHeal_Click()
    MapEditorHealType = cmbHeal.ListIndex + 1
    MapEditorHealAmount = scrlHeal.value
    picAttributes.visible = False
    fraHeal.visible = False
End Sub

Private Sub cmdKeyOpen_Click()
    KeyOpenEditorX = scrlKeyX.value
    KeyOpenEditorY = scrlKeyY.value
    picAttributes.visible = False
    fraKeyOpen.visible = False
End Sub

Private Sub cmdMapItem_Click()
    ItemEditorNum = scrlMapItem.value
    ItemEditorValue = scrlMapItemValue.value
    picAttributes.visible = False
    fraMapItem.visible = False
End Sub

Private Sub cmdMapKey_Click()
    KeyEditorNum = scrlMapKey.value
    KeyEditorTake = chkMapKey.value
    KeyEditorTime = scrlKeyTime.value
    If KeyEditorTime = 0 Then KeyEditorTime = -1
    picAttributes.visible = False
    fraMapKey.visible = False
End Sub

Private Sub cmdMapWarp_Click()
    EditorWarpMap = scrlMapWarp.value
    EditorWarpX = scrlMapWarpX.value
    EditorWarpY = scrlMapWarpY.value
    EditorWarpFall = chkWarpFall.value
    picAttributes.visible = False
    fraMapWarp.visible = False
End Sub

Private Sub cmdNCache_Click()
    initAutotiles
End Sub

Private Sub cmdNDone_Click()
    fraNudge.visible = False
    picAttributes.visible = False
End Sub

Private Sub cmdNDown_Click()
    NudgeMap DIR_DOWN
End Sub

Private Sub cmdNLeft_Click()
    NudgeMap DIR_LEFT
End Sub

Private Sub cmdNpcSpawn_Click()
    SpawnNpcNum = lstNpc.ListIndex + 1
    SpawnNpcDir = scrlNpcDir.value
    picAttributes.visible = False
    fraNpcSpawn.visible = False
End Sub

Private Sub cmdNRight_Click()
    NudgeMap DIR_RIGHT
End Sub

Private Sub cmdNudge_Click()
    picAttributes.visible = True
    fraNudge.visible = True
End Sub

Private Sub cmdNUp_Click()
    NudgeMap DIR_UP
End Sub

Private Sub cmdResourceOk_Click()
    ResourceEditorNum = scrlResource.value
    picAttributes.visible = False
    fraResource.visible = False
End Sub

Private Sub cmdShop_Click()
    EditorShop = cmbShop.ListIndex
    picAttributes.visible = False
    fraShop.visible = False
End Sub

Private Sub cmdSlide_Click()
    MapEditorSlideDir = cmbSlide.ListIndex
    picAttributes.visible = False
    fraSlide.visible = False
End Sub

Private Sub cmdTrap_Click()
    MapEditorHealAmount = scrlTrap.value
    picAttributes.visible = False
    fraTrap.visible = False
End Sub

Private Sub Form_Load()
    ' move the entire attributes box on screen
    picAttributes.left = 8
    picAttributes.top = 8
End Sub

Private Sub optAppear_Click()
    ClearAttributeDialogue
    picAttributes.visible = True
    fraAppear.visible = True
End Sub

Private Sub optDoor_Click()
    ClearAttributeDialogue
    picAttributes.visible = True
    fraMapWarp.visible = True
    scrlMapWarp.Max = MAX_MAPS
    scrlMapWarpX.Max = MAX_BYTE
    scrlMapWarpY.Max = MAX_BYTE
End Sub

Private Sub optEvents_Click()
    selTileX = 0
    selTileY = 0
End Sub

Private Sub optHeal_Click()
    ClearAttributeDialogue
    picAttributes.visible = True
    fraHeal.visible = True
End Sub

Private Sub optLayers_Click()

    If optLayers.value Then
        fraLayers.visible = True
        fraAttribs.visible = False
    End If

End Sub

Private Sub optAttribs_Click()

    If optAttribs.value Then
        fraLayers.visible = False
        fraAttribs.visible = True
    End If

End Sub

Private Sub optNpcSpawn_Click()
    Dim n As Long
    
    If lstNpc.ListCount <= 0 Then
        lstNpc.Clear
        For n = 1 To MAX_MAP_NPCS
            If map.MapData.Npc(n) > 0 Then
                lstNpc.AddItem n & ": " & Npc(map.MapData.Npc(n)).name
            Else
                lstNpc.AddItem n & ": No Npc"
            End If
        Next n
        lstNpc.ListIndex = 0
    End If
    
    ClearAttributeDialogue
    picAttributes.visible = True
    fraNpcSpawn.visible = True
End Sub

Private Sub optResource_Click()
    ClearAttributeDialogue
    picAttributes.visible = True
    fraResource.visible = True
End Sub

Private Sub optShop_Click()
    ClearAttributeDialogue
    picAttributes.visible = True
    fraShop.visible = True
End Sub

Private Sub optSlide_Click()
    ClearAttributeDialogue
    picAttributes.visible = True
    fraSlide.visible = True
End Sub

Private Sub optTrap_Click()
    ClearAttributeDialogue
    picAttributes.visible = True
    fraTrap.visible = True
End Sub

Private Sub picBackSelect_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call MapEditorChooseTile(Button, x, y)
End Sub
 
Private Sub picBackSelect_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    shpLocTop = (y \ PIC_Y) * PIC_Y
    shpLocLeft = (x \ PIC_X) * PIC_X
    Call MapEditorDrag(Button, x, y)
End Sub

Private Sub cmdSend_Click()
    Call MapEditorSend
End Sub

Private Sub cmdCancel_Click()
    Call MapEditorCancel
End Sub

Private Sub cmdProperties_Click()
    Load frmEditor_MapProperties
    MapEditorProperties
    frmEditor_MapProperties.Show vbModal
End Sub

Private Sub optWarp_Click()
    ClearAttributeDialogue
    picAttributes.visible = True
    fraMapWarp.visible = True
    scrlMapWarp.Max = MAX_MAPS
    scrlMapWarpX.Max = MAX_BYTE
    scrlMapWarpY.Max = MAX_BYTE
End Sub

Private Sub optItem_Click()
    ClearAttributeDialogue
    picAttributes.visible = True
    fraMapItem.visible = True
    scrlMapItem.Max = MAX_ITEMS
    lblMapItem.caption = Trim$(Item(scrlMapItem.value).name) & " x" & scrlMapItemValue.value
End Sub

Private Sub optKey_Click()
    ClearAttributeDialogue
    picAttributes.visible = True
    fraMapKey.visible = True
    scrlMapKey.Max = MAX_ITEMS
    lblMapKey.caption = "Item: " & Trim$(Item(scrlMapKey.value).name)
End Sub

Private Sub optKeyOpen_Click()
    ClearAttributeDialogue
    fraKeyOpen.visible = True
    picAttributes.visible = True
    scrlKeyX.Max = map.MapData.MaxX
    scrlKeyY.Max = map.MapData.MaxY
End Sub

Private Sub cmdFill_Click()
    MapEditorFillLayer
End Sub

Private Sub cmdClear_Click()
    Call MapEditorClearLayer
End Sub

Private Sub cmdClear2_Click()
    Call MapEditorClearAttribs
End Sub

Private Sub scrlAutotile_Change()

    Select Case scrlAutotile.value

        Case 0 ' normal
            lblAutotile.caption = "Normal"

        Case 1 ' autotile
            lblAutotile.caption = "Autotile"

        Case 2 ' fake autotile
            lblAutotile.caption = "Fake"

        Case 3 ' animated
            lblAutotile.caption = "Animated"

        Case 4 ' cliff
            lblAutotile.caption = "Cliff"

        Case 5 ' waterfall
            lblAutotile.caption = "Waterfall"
    End Select

End Sub

Private Sub scrlHeal_Change()
    lblHeal.caption = "Amount: " & scrlHeal.value
End Sub



Private Sub scrlKeyTime_Change()
    lblKeyTime.caption = "Time: " & scrlKeyTime.value & "s"
End Sub

Private Sub scrlKeyX_Change()
    lblKeyX.caption = "X: " & scrlKeyX.value
End Sub

Private Sub scrlKeyX_Scroll()
    scrlKeyX_Change
End Sub

Private Sub scrlKeyY_Change()
    lblKeyY.caption = "Y: " & scrlKeyY.value
End Sub

Private Sub scrlKeyY_Scroll()
    scrlKeyY_Change
End Sub

Private Sub scrlTrap_Change()
    lblTrap.caption = "Amount: " & scrlTrap.value
End Sub

Private Sub scrlMapItem_Change()

    If Item(scrlMapItem.value).Type = ITEM_TYPE_CURRENCY Then
        scrlMapItemValue.enabled = True
    Else
        scrlMapItemValue.value = 1
        scrlMapItemValue.enabled = False
    End If

    lblMapItem.caption = Trim$(Item(scrlMapItem.value).name) & " x" & scrlMapItemValue.value
End Sub

Private Sub scrlMapItem_Scroll()
    scrlMapItem_Change
End Sub

Private Sub scrlMapItemValue_Change()
    lblMapItem.caption = Trim$(Item(scrlMapItem.value).name) & " x" & scrlMapItemValue.value
End Sub

Private Sub scrlMapItemValue_Scroll()
    scrlMapItemValue_Change
End Sub

Private Sub scrlMapKey_Change()
    lblMapKey.caption = "Item: " & Trim$(Item(scrlMapKey.value).name)
End Sub

Private Sub scrlMapKey_Scroll()
    scrlMapKey_Change
End Sub

Private Sub scrlMapWarp_Change()
    lblMapWarp.caption = "Map: " & scrlMapWarp.value
End Sub

Private Sub scrlMapWarp_Scroll()
    scrlMapWarp_Change
End Sub

Private Sub scrlMapWarpX_Change()
    lblMapWarpX.caption = "X: " & scrlMapWarpX.value
End Sub

Private Sub scrlMapWarpX_Scroll()
    scrlMapWarpX_Change
End Sub

Private Sub scrlMapWarpY_Change()
    lblMapWarpY.caption = "Y: " & scrlMapWarpY.value
End Sub

Private Sub scrlMapWarpY_Scroll()
    scrlMapWarpY_Change
End Sub

Private Sub scrlNpcDir_Change()

    Select Case scrlNpcDir.value

        Case DIR_DOWN
            lblNpcDir = "Direction: Down"

        Case DIR_UP
            lblNpcDir = "Direction: Up"

        Case DIR_LEFT
            lblNpcDir = "Direction: Left"

        Case DIR_RIGHT
            lblNpcDir = "Direction: Right"
    End Select

End Sub

Private Sub scrlNpcDir_Scroll()
    scrlNpcDir_Change
End Sub

Private Sub scrlResource_Change()
    lblResource.caption = "Resource: " & Resource(scrlResource.value).name
End Sub

Private Sub scrlResource_Scroll()
    scrlResource_Change
End Sub

Private Sub scrlPictureX_Change()
    Call MapEditorTileScroll
End Sub

Private Sub scrlPictureY_Change()
    Call MapEditorTileScroll
End Sub

Private Sub scrlPictureX_Scroll()
    scrlPictureY_Change
End Sub

Private Sub scrlPictureY_Scroll()
    scrlPictureY_Change
End Sub

Private Sub scrlTileSet_Change()
    fraTileSet.caption = "Tileset: " & scrlTileSet.value
    frmEditor_Map.scrlPictureX.value = 0
    frmEditor_Map.scrlPictureY.value = 0
    frmEditor_Map.picBackSelect.left = 0
    frmEditor_Map.picBackSelect.top = 0
    GDIRenderTileset
    frmEditor_Map.scrlPictureY.Max = (frmEditor_Map.picBackSelect.height \ PIC_Y) - (frmEditor_Map.picBack.height \ PIC_Y)
    frmEditor_Map.scrlPictureX.Max = (frmEditor_Map.picBackSelect.width \ PIC_X) - (frmEditor_Map.picBack.width \ PIC_X)
    MapEditorTileScroll
End Sub

Private Sub scrlTileSet_Scroll()
    scrlTileSet_Change
End Sub

Private Sub cmdChat_Click()
    MapEditorChatNpc = lstChat.ListIndex + 1
    MapEditorChatDir = scrlChat.value
    picAttributes.visible = False
    fraChat.visible = False
End Sub

Private Sub optChat_Click()
    Dim n As Long
    If lstChat.ListCount <= 0 Then
        lstChat.Clear
        For n = 1 To MAX_MAP_NPCS
            If map.MapData.Npc(n) > 0 Then
                lstChat.AddItem n & ": " & Npc(map.MapData.Npc(n)).name
            Else
                lstChat.AddItem n & ": No Npc"
            End If
        Next n
        scrlChat.value = 0
        lstChat.ListIndex = 0
    End If
    
    ClearAttributeDialogue
    picAttributes.visible = True
    fraChat.visible = True
End Sub

Private Sub scrlChat_Change()
    Dim sAppend As String

    Select Case scrlChat.value

        Case DIR_UP
            sAppend = "Up"

        Case DIR_DOWN
            sAppend = "Down"

        Case DIR_RIGHT
            sAppend = "Right"

        Case DIR_LEFT
            sAppend = "Left"
    End Select

    lblChat.caption = "Direction Req: " & sAppend
End Sub
