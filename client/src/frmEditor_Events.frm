VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL32.OCX"
Begin VB.Form frmEditor_Events 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Event Editor"
   ClientHeight    =   9870
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12855
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   658
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   857
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraGraphic 
      Caption         =   "Graphic Selection"
      Height          =   9135
      Left            =   120
      TabIndex        =   119
      Top             =   120
      Visible         =   0   'False
      Width           =   12615
      Begin VB.PictureBox picGraphicSel 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   7800
         Left            =   240
         ScaleHeight     =   520
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   808
         TabIndex        =   126
         Top             =   720
         Width           =   12120
      End
      Begin VB.CommandButton cmdGraphicCancel 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   11040
         TabIndex        =   125
         Top             =   8640
         Width           =   1455
      End
      Begin VB.CommandButton cmdGraphicOK 
         Caption         =   "OK"
         Height          =   375
         Left            =   9480
         TabIndex        =   124
         Top             =   8640
         Width           =   1455
      End
      Begin VB.ComboBox cmbGraphic 
         Height          =   315
         ItemData        =   "frmEditor_Events.frx":0000
         Left            =   720
         List            =   "frmEditor_Events.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   121
         Top             =   240
         Width           =   2175
      End
      Begin VB.HScrollBar scrlGraphic 
         Height          =   255
         Left            =   4440
         TabIndex        =   120
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label Label6 
         Caption         =   "Type:"
         Height          =   255
         Left            =   120
         TabIndex        =   123
         Top             =   270
         Width           =   855
      End
      Begin VB.Label lblGraphic 
         Caption         =   "Number: 1"
         Height          =   255
         Left            =   3000
         TabIndex        =   122
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame fraDialogue 
      Height          =   7455
      Left            =   6240
      TabIndex        =   91
      Top             =   1560
      Visible         =   0   'False
      Width           =   6375
      Begin VB.Frame fraWarpPlayer 
         Caption         =   "Warp Player"
         Height          =   3015
         Left            =   1200
         TabIndex        =   134
         Top             =   2040
         Width           =   4095
         Begin VB.CommandButton cmdWPCancel 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   2760
            TabIndex        =   142
            Top             =   2520
            Width           =   1215
         End
         Begin VB.CommandButton cmdWPOkay 
            Caption         =   "Ok"
            Height          =   375
            Left            =   1440
            TabIndex        =   141
            Top             =   2520
            Width           =   1215
         End
         Begin VB.HScrollBar scrlWPY 
            Height          =   255
            Left            =   120
            Max             =   255
            TabIndex        =   140
            Top             =   1800
            Width           =   3855
         End
         Begin VB.HScrollBar scrlWPX 
            Height          =   255
            Left            =   120
            Max             =   255
            TabIndex        =   138
            Top             =   1200
            Width           =   3855
         End
         Begin VB.HScrollBar scrlWPMap 
            Height          =   255
            Left            =   120
            TabIndex        =   136
            Top             =   600
            Width           =   3855
         End
         Begin VB.Label lblWPY 
            Caption         =   "Y: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   139
            Top             =   1560
            Width           =   3135
         End
         Begin VB.Label lblWPX 
            Caption         =   "X: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   137
            Top             =   960
            Width           =   3135
         End
         Begin VB.Label lblWPMap 
            Caption         =   "Map: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   135
            Top             =   360
            Width           =   3135
         End
      End
      Begin VB.Frame fraChatBubble 
         Caption         =   "Chat Bubble"
         Height          =   5295
         Left            =   1200
         TabIndex        =   109
         Top             =   960
         Visible         =   0   'False
         Width           =   4095
         Begin VB.ComboBox cmbChatBubble 
            Height          =   315
            ItemData        =   "frmEditor_Events.frx":002B
            Left            =   120
            List            =   "frmEditor_Events.frx":002D
            Style           =   2  'Dropdown List
            TabIndex        =   118
            Top             =   4320
            Width           =   3855
         End
         Begin VB.ComboBox cmbChatBubbleType 
            Height          =   315
            ItemData        =   "frmEditor_Events.frx":002F
            Left            =   120
            List            =   "frmEditor_Events.frx":003F
            Style           =   2  'Dropdown List
            TabIndex        =   116
            Top             =   3720
            Width           =   3855
         End
         Begin VB.CommandButton cmdChatBubbleOk 
            Caption         =   "Ok"
            Height          =   375
            Left            =   1440
            TabIndex        =   114
            Top             =   4800
            Width           =   1215
         End
         Begin VB.CommandButton cmdChatBubbleCancel 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   2760
            TabIndex        =   113
            Top             =   4800
            Width           =   1215
         End
         Begin VB.HScrollBar scrlChatBubble 
            Height          =   255
            Left            =   120
            Max             =   18
            TabIndex        =   111
            Top             =   3120
            Value           =   1
            Width           =   3855
         End
         Begin VB.TextBox txtChatBubble 
            Height          =   2535
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   110
            Top             =   240
            Width           =   3855
         End
         Begin VB.Label Label13 
            Caption         =   "Target:"
            Height          =   255
            Left            =   120
            TabIndex        =   117
            Top             =   4080
            Width           =   3735
         End
         Begin VB.Label Label11 
            Caption         =   "Target Type:"
            Height          =   255
            Left            =   120
            TabIndex        =   115
            Top             =   3480
            Width           =   3255
         End
         Begin VB.Label lblChatBubble 
            Caption         =   "Colour: Black"
            Height          =   255
            Left            =   120
            TabIndex        =   112
            Top             =   2880
            Width           =   3255
         End
      End
      Begin VB.Frame fraPlayerVar 
         Caption         =   "Player Variable"
         Height          =   1695
         Left            =   1200
         TabIndex        =   127
         Top             =   2760
         Visible         =   0   'False
         Width           =   4095
         Begin VB.CommandButton cmdVariableCancel 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   2760
            TabIndex        =   133
            Top             =   1200
            Width           =   1215
         End
         Begin VB.CommandButton cmdVariableOK 
            Caption         =   "Ok"
            Height          =   375
            Left            =   1440
            TabIndex        =   132
            Top             =   1200
            Width           =   1215
         End
         Begin VB.TextBox txtVariable 
            Height          =   285
            Left            =   960
            TabIndex        =   131
            Top             =   840
            Width           =   3015
         End
         Begin VB.ComboBox cmbVariable 
            Height          =   315
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   129
            Top             =   360
            Width           =   3015
         End
         Begin VB.Label Label15 
            Caption         =   "Set to:"
            Height          =   255
            Left            =   120
            TabIndex        =   130
            Top             =   840
            Width           =   1815
         End
         Begin VB.Label Label14 
            Caption         =   "Variable:"
            Height          =   255
            Left            =   120
            TabIndex        =   128
            Top             =   360
            Width           =   3855
         End
      End
      Begin VB.Frame fraAddText 
         Caption         =   "Add Text"
         Height          =   4095
         Left            =   1200
         TabIndex        =   92
         Top             =   1560
         Visible         =   0   'False
         Width           =   4095
         Begin VB.CommandButton cmdAddText_Cancel 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   2760
            TabIndex        =   102
            Top             =   3600
            Width           =   1215
         End
         Begin VB.CommandButton cmdAddText_Ok 
            Caption         =   "Ok"
            Height          =   375
            Left            =   1440
            TabIndex        =   101
            Top             =   3600
            Width           =   1215
         End
         Begin VB.OptionButton optAddText_Global 
            Caption         =   "Global"
            Height          =   255
            Left            =   1920
            TabIndex        =   100
            Top             =   3240
            Width           =   855
         End
         Begin VB.OptionButton optAddText_Map 
            Caption         =   "Map"
            Height          =   255
            Left            =   1080
            TabIndex        =   99
            Top             =   3240
            Width           =   735
         End
         Begin VB.OptionButton optAddText_Game 
            Caption         =   "Game"
            Height          =   255
            Left            =   120
            TabIndex        =   98
            Top             =   3240
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.HScrollBar scrlAddText_Colour 
            Height          =   255
            Left            =   120
            Max             =   18
            TabIndex        =   96
            Top             =   2640
            Value           =   1
            Width           =   3855
         End
         Begin VB.TextBox txtAddText_Text 
            Height          =   1815
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   94
            Top             =   480
            Width           =   3855
         End
         Begin VB.Label Label12 
            Caption         =   "Channel:"
            Height          =   255
            Left            =   120
            TabIndex        =   97
            Top             =   3000
            Width           =   1575
         End
         Begin VB.Label lblAddText_Colour 
            Caption         =   "Colour: Black"
            Height          =   255
            Left            =   120
            TabIndex        =   95
            Top             =   2400
            Width           =   3255
         End
         Begin VB.Label Label7 
            Caption         =   "Text:"
            Height          =   255
            Left            =   120
            TabIndex        =   93
            Top             =   240
            Width           =   1935
         End
      End
   End
   Begin VB.Frame fraCommands 
      Caption         =   "Commands"
      Height          =   7455
      Left            =   6240
      TabIndex        =   43
      Top             =   1560
      Visible         =   0   'False
      Width           =   6375
      Begin VB.PictureBox picCommands 
         BorderStyle     =   0  'None
         Height          =   6015
         Index           =   1
         Left            =   240
         ScaleHeight     =   6015
         ScaleWidth      =   5775
         TabIndex        =   45
         Top             =   720
         Width           =   5775
         Begin VB.Frame Frame12 
            Caption         =   "Player Control"
            Height          =   5055
            Left            =   3000
            TabIndex        =   59
            Top             =   0
            Width           =   2775
            Begin VB.CommandButton Command19 
               Caption         =   "Change Sex"
               Enabled         =   0   'False
               Height          =   375
               Left            =   120
               TabIndex        =   69
               Top             =   4560
               Width           =   2535
            End
            Begin VB.CommandButton Command18 
               Caption         =   "Change Sprite"
               Enabled         =   0   'False
               Height          =   375
               Left            =   120
               TabIndex        =   68
               Top             =   4080
               Width           =   2535
            End
            Begin VB.CommandButton Command17 
               Caption         =   "Change Class"
               Enabled         =   0   'False
               Height          =   375
               Left            =   120
               TabIndex        =   67
               Top             =   3600
               Width           =   2535
            End
            Begin VB.CommandButton Command16 
               Caption         =   "Change Skills"
               Enabled         =   0   'False
               Height          =   375
               Left            =   120
               TabIndex        =   66
               Top             =   3120
               Width           =   2535
            End
            Begin VB.CommandButton Command15 
               Caption         =   "Change Level"
               Enabled         =   0   'False
               Height          =   375
               Left            =   120
               TabIndex        =   65
               Top             =   2640
               Width           =   2535
            End
            Begin VB.CommandButton Command14 
               Caption         =   "Change EXP"
               Enabled         =   0   'False
               Height          =   375
               Left            =   120
               TabIndex        =   64
               Top             =   2160
               Width           =   2535
            End
            Begin VB.CommandButton Command13 
               Caption         =   "Change SP"
               Enabled         =   0   'False
               Height          =   375
               Left            =   120
               TabIndex        =   63
               Top             =   1680
               Width           =   2535
            End
            Begin VB.CommandButton Command12 
               Caption         =   "Change HP"
               Enabled         =   0   'False
               Height          =   375
               Left            =   120
               TabIndex        =   62
               Top             =   1200
               Width           =   2535
            End
            Begin VB.CommandButton Command11 
               Caption         =   "Change Items"
               Enabled         =   0   'False
               Height          =   375
               Left            =   120
               TabIndex        =   61
               Top             =   720
               Width           =   2535
            End
            Begin VB.CommandButton Command8 
               Caption         =   "Change Gold"
               Enabled         =   0   'False
               Height          =   375
               Left            =   120
               TabIndex        =   60
               Top             =   240
               Width           =   2535
            End
         End
         Begin VB.Frame Frame11 
            Caption         =   "Flow Control"
            Height          =   1215
            Left            =   0
            TabIndex        =   56
            Top             =   4560
            Width           =   2775
            Begin VB.CommandButton Command10 
               Caption         =   "Conditional Branch"
               Enabled         =   0   'False
               Height          =   375
               Left            =   120
               TabIndex        =   58
               Top             =   240
               Width           =   2535
            End
            Begin VB.CommandButton Command9 
               Caption         =   "Exit Process"
               Enabled         =   0   'False
               Height          =   375
               Left            =   120
               TabIndex        =   57
               Top             =   720
               Width           =   2535
            End
         End
         Begin VB.Frame Frame10 
            Caption         =   "Event Progression"
            Height          =   1695
            Left            =   0
            TabIndex        =   52
            Top             =   2760
            Width           =   2775
            Begin VB.CommandButton Command7 
               Caption         =   "Self Switch"
               Enabled         =   0   'False
               Height          =   375
               Left            =   120
               TabIndex        =   55
               Top             =   1200
               Width           =   2535
            End
            Begin VB.CommandButton Command6 
               Caption         =   "Event Switch"
               Enabled         =   0   'False
               Height          =   375
               Left            =   120
               TabIndex        =   54
               Top             =   720
               Width           =   2535
            End
            Begin VB.CommandButton cmdPlayerVar 
               Caption         =   "Player Variable"
               Height          =   375
               Left            =   120
               TabIndex        =   53
               Top             =   240
               Width           =   2535
            End
         End
         Begin VB.Frame Frame9 
            Caption         =   "Message"
            Height          =   2655
            Left            =   0
            TabIndex        =   47
            Top             =   0
            Width           =   2775
            Begin VB.CommandButton cmdChatBubble 
               Caption         =   "Show Chat Bubble"
               Height          =   375
               Left            =   120
               TabIndex        =   108
               Top             =   1200
               Width           =   2535
            End
            Begin VB.CommandButton Command4 
               Caption         =   "Input Number"
               Enabled         =   0   'False
               Height          =   375
               Left            =   120
               TabIndex        =   51
               Top             =   2160
               Width           =   2535
            End
            Begin VB.CommandButton Command3 
               Caption         =   "Show Choices"
               Enabled         =   0   'False
               Height          =   375
               Left            =   120
               TabIndex        =   50
               Top             =   1680
               Width           =   2535
            End
            Begin VB.CommandButton Command2 
               Caption         =   "Show Text"
               Enabled         =   0   'False
               Height          =   375
               Left            =   120
               TabIndex        =   49
               Top             =   720
               Width           =   2535
            End
            Begin VB.CommandButton cmdAddText 
               Caption         =   "Add Chatbox Text"
               Height          =   375
               Left            =   120
               TabIndex        =   48
               Top             =   240
               Width           =   2535
            End
         End
      End
      Begin VB.PictureBox picCommands 
         BorderStyle     =   0  'None
         Height          =   5775
         Index           =   2
         Left            =   240
         ScaleHeight     =   5775
         ScaleWidth      =   5775
         TabIndex        =   46
         Top             =   720
         Visible         =   0   'False
         Width           =   5775
         Begin VB.Frame Frame16 
            Caption         =   "Music and Sound"
            Height          =   3135
            Left            =   3000
            TabIndex        =   83
            Top             =   0
            Width           =   2775
            Begin VB.CommandButton Command35 
               Caption         =   "Stop Sound"
               Enabled         =   0   'False
               Height          =   375
               Left            =   120
               TabIndex        =   89
               Top             =   2640
               Width           =   2535
            End
            Begin VB.CommandButton Command34 
               Caption         =   "Play Sound"
               Enabled         =   0   'False
               Height          =   375
               Left            =   120
               TabIndex        =   88
               Top             =   2160
               Width           =   2535
            End
            Begin VB.CommandButton Command33 
               Caption         =   "Fadeout BGS"
               Enabled         =   0   'False
               Height          =   375
               Left            =   120
               TabIndex        =   87
               Top             =   1680
               Width           =   2535
            End
            Begin VB.CommandButton Command32 
               Caption         =   "Play BGS"
               Enabled         =   0   'False
               Height          =   375
               Left            =   120
               TabIndex        =   86
               Top             =   1200
               Width           =   2535
            End
            Begin VB.CommandButton Command31 
               Caption         =   "Fadeout BGM"
               Enabled         =   0   'False
               Height          =   375
               Left            =   120
               TabIndex        =   85
               Top             =   720
               Width           =   2535
            End
            Begin VB.CommandButton Command30 
               Caption         =   "Play BGM"
               Enabled         =   0   'False
               Height          =   375
               Left            =   120
               TabIndex        =   84
               Top             =   240
               Width           =   2535
            End
         End
         Begin VB.Frame Frame15 
            Caption         =   "Screen Effects"
            Height          =   2655
            Left            =   0
            TabIndex        =   77
            Top             =   3120
            Width           =   2775
            Begin VB.CommandButton Command29 
               Caption         =   "Shake Screen"
               Enabled         =   0   'False
               Height          =   375
               Left            =   120
               TabIndex        =   82
               Top             =   2160
               Width           =   2535
            End
            Begin VB.CommandButton Command28 
               Caption         =   "Flash Screen"
               Enabled         =   0   'False
               Height          =   375
               Left            =   120
               TabIndex        =   81
               Top             =   1680
               Width           =   2535
            End
            Begin VB.CommandButton Command27 
               Caption         =   "Tint Screen"
               Enabled         =   0   'False
               Height          =   375
               Left            =   120
               TabIndex        =   80
               Top             =   1200
               Width           =   2535
            End
            Begin VB.CommandButton Command26 
               Caption         =   "Fadein"
               Enabled         =   0   'False
               Height          =   375
               Left            =   120
               TabIndex        =   79
               Top             =   720
               Width           =   2535
            End
            Begin VB.CommandButton Command23 
               Caption         =   "Fadeout"
               Enabled         =   0   'False
               Height          =   375
               Left            =   120
               TabIndex        =   78
               Top             =   240
               Width           =   2535
            End
         End
         Begin VB.Frame Frame14 
            Caption         =   "Animation"
            Height          =   1215
            Left            =   0
            TabIndex        =   74
            Top             =   1800
            Width           =   2775
            Begin VB.CommandButton Command25 
               Caption         =   "Play Animation"
               Enabled         =   0   'False
               Height          =   375
               Left            =   120
               TabIndex        =   76
               Top             =   240
               Width           =   2535
            End
            Begin VB.CommandButton Command24 
               Caption         =   "Play Emoticon"
               Enabled         =   0   'False
               Height          =   375
               Left            =   120
               TabIndex        =   75
               Top             =   720
               Width           =   2535
            End
         End
         Begin VB.Frame Frame13 
            Caption         =   "Movement"
            Height          =   1695
            Left            =   0
            TabIndex        =   70
            Top             =   0
            Width           =   2775
            Begin VB.CommandButton Command22 
               Caption         =   "Scroll Map"
               Enabled         =   0   'False
               Height          =   375
               Left            =   120
               TabIndex        =   73
               Top             =   1200
               Width           =   2535
            End
            Begin VB.CommandButton Command21 
               Caption         =   "Warp Party"
               Enabled         =   0   'False
               Height          =   375
               Left            =   120
               TabIndex        =   72
               Top             =   720
               Width           =   2535
            End
            Begin VB.CommandButton cmdWarpPlayer 
               Caption         =   "Warp Player"
               Height          =   375
               Left            =   120
               TabIndex        =   71
               Top             =   240
               Width           =   2535
            End
         End
      End
      Begin VB.CommandButton cmdCancelCommand 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   4560
         TabIndex        =   90
         Top             =   6840
         Width           =   1455
      End
      Begin MSComctlLib.TabStrip tabCommands 
         Height          =   7095
         Left            =   120
         TabIndex        =   44
         Top             =   240
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   12515
         TabMinWidth     =   1764
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   2
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "1"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "2"
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame17 
      Caption         =   "Commands"
      Height          =   735
      Left            =   6240
      TabIndex        =   103
      Top             =   8280
      Width           =   6255
      Begin VB.CommandButton cmdClearCommand 
         Caption         =   "Clear"
         Height          =   375
         Left            =   4680
         TabIndex        =   107
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdDeleteCommand 
         Caption         =   "Delete"
         Height          =   375
         Left            =   3120
         TabIndex        =   106
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdEditCommand 
         Caption         =   "Edit"
         Height          =   375
         Left            =   1560
         TabIndex        =   105
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdAddCommand 
         Caption         =   "Add"
         Height          =   375
         Left            =   120
         TabIndex        =   104
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   9720
      TabIndex        =   42
      Top             =   9360
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   11280
      TabIndex        =   41
      Top             =   9360
      Width           =   1455
   End
   Begin VB.Frame Frame8 
      Caption         =   "General"
      Height          =   855
      Left            =   120
      TabIndex        =   32
      Top             =   120
      Width           =   12615
      Begin VB.CommandButton cmdClearPage 
         Caption         =   "Clear Page"
         Height          =   375
         Left            =   10920
         TabIndex        =   39
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdDeletePage 
         Caption         =   "Delete Page"
         Enabled         =   0   'False
         Height          =   375
         Left            =   9360
         TabIndex        =   38
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdPastePage 
         Caption         =   "Paste Page"
         Enabled         =   0   'False
         Height          =   375
         Left            =   7800
         TabIndex        =   37
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdCopyPage 
         Caption         =   "Copy Page"
         Height          =   375
         Left            =   6240
         TabIndex        =   36
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdNewPage 
         Caption         =   "New Page"
         Height          =   375
         Left            =   4680
         TabIndex        =   35
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   840
         TabIndex        =   34
         Top             =   300
         Width           =   3135
      End
      Begin VB.Label Label5 
         Caption         =   "Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   330
         Width           =   735
      End
   End
   Begin VB.Frame Frame7 
      Height          =   1335
      Left            =   2760
      TabIndex        =   31
      Top             =   7680
      Width           =   3375
   End
   Begin VB.Frame Frame6 
      Caption         =   "Trigger"
      Height          =   975
      Left            =   2760
      TabIndex        =   29
      Top             =   6600
      Width           =   3375
      Begin VB.ComboBox cmbTrigger 
         Height          =   315
         ItemData        =   "frmEditor_Events.frx":005D
         Left            =   120
         List            =   "frmEditor_Events.frx":006A
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   360
         Width           =   3135
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Priority"
      Height          =   975
      Left            =   2760
      TabIndex        =   27
      Top             =   5520
      Width           =   3375
      Begin VB.ComboBox cmbPriority 
         Height          =   315
         ItemData        =   "frmEditor_Events.frx":0094
         Left            =   120
         List            =   "frmEditor_Events.frx":00A1
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   360
         Width           =   3135
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Options"
      Height          =   1935
      Left            =   360
      TabIndex        =   22
      Top             =   7080
      Width           =   2295
      Begin VB.CheckBox chkWalkThrough 
         Caption         =   "Walk Through"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   1440
         Width           =   1695
      End
      Begin VB.CheckBox chkDirFix 
         Caption         =   "Direction Fix"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   1080
         Width           =   1575
      End
      Begin VB.CheckBox chkStepAnim 
         Caption         =   "Stepping Animation"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   720
         Width           =   2055
      End
      Begin VB.CheckBox chkWalkAnim 
         Caption         =   "Walking Animation"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Movement"
      Height          =   1935
      Left            =   2760
      TabIndex        =   15
      Top             =   3480
      Width           =   3375
      Begin VB.ComboBox cmbMoveFreq 
         Height          =   315
         ItemData        =   "frmEditor_Events.frx":00DD
         Left            =   840
         List            =   "frmEditor_Events.frx":00F0
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   1320
         Width           =   2415
      End
      Begin VB.ComboBox cmbMoveSpeed 
         Height          =   315
         ItemData        =   "frmEditor_Events.frx":011C
         Left            =   840
         List            =   "frmEditor_Events.frx":0132
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   840
         Width           =   2415
      End
      Begin VB.ComboBox cmbMoveType 
         Height          =   315
         ItemData        =   "frmEditor_Events.frx":0175
         Left            =   840
         List            =   "frmEditor_Events.frx":017F
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label10 
         Caption         =   "Freq:"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1350
         Width           =   1695
      End
      Begin VB.Label Label9 
         Caption         =   "Speed:"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   870
         Width           =   1695
      End
      Begin VB.Label Label8 
         Caption         =   "Type:"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   390
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Graphic"
      Height          =   3495
      Left            =   360
      TabIndex        =   13
      Top             =   3480
      Width           =   2295
      Begin VB.PictureBox picGraphic 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2895
         Left            =   240
         ScaleHeight     =   193
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   121
         TabIndex        =   14
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Conditions"
      Height          =   1815
      Left            =   360
      TabIndex        =   2
      Top             =   1560
      Width           =   5775
      Begin VB.CheckBox chkHasItem 
         Caption         =   "Has Item"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1350
         Width           =   1695
      End
      Begin VB.ComboBox cmbHasItem 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmEditor_Events.frx":0192
         Left            =   1920
         List            =   "frmEditor_Events.frx":0194
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1320
         Width           =   1695
      End
      Begin VB.CheckBox chkSelfSwitch 
         Caption         =   "Self Switch"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   870
         Width           =   1695
      End
      Begin VB.ComboBox cmbSelfSwitch 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmEditor_Events.frx":0196
         Left            =   1920
         List            =   "frmEditor_Events.frx":01A9
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   840
         Width           =   1695
      End
      Begin VB.CheckBox chkPlayerVar 
         Caption         =   "Player Variable"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   390
         Width           =   1695
      End
      Begin VB.ComboBox cmbPlayerVar 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox txtPlayerVariable 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3960
         TabIndex        =   3
         Top             =   390
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "is ON"
         Height          =   255
         Left            =   3720
         TabIndex        =   10
         Top             =   870
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "is"
         Height          =   255
         Left            =   3720
         TabIndex        =   7
         Top             =   390
         Width           =   255
      End
      Begin VB.Label Label3 
         Caption         =   "or above"
         Height          =   255
         Left            =   4800
         TabIndex        =   6
         Top             =   390
         Width           =   855
      End
   End
   Begin VB.ListBox lstCommands 
      Height          =   6495
      Left            =   6240
      TabIndex        =   1
      Top             =   1680
      Width           =   6255
   End
   Begin MSComctlLib.TabStrip tabPages 
      Height          =   8175
      Left            =   120
      TabIndex        =   40
      Top             =   1080
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   14420
      MultiRow        =   -1  'True
      ShowTips        =   0   'False
      TabMinWidth     =   529
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "1"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "List of commands:"
      Height          =   255
      Left            =   6240
      TabIndex        =   0
      Top             =   1560
      Width           =   6255
   End
End
Attribute VB_Name = "frmEditor_Events"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private copyPage As EventPageRec
Private isEdit As Boolean

Private Sub chkDirFix_Click()
    tmpEvent.EventPage(curPageNum).DirFix = chkDirFix.value
End Sub

Private Sub chkHasItem_Click()
    tmpEvent.EventPage(curPageNum).chkHasItem = chkHasItem.value
    If chkHasItem.value = 0 Then cmbHasItem.enabled = False Else cmbHasItem.enabled = True
End Sub

Private Sub chkPlayerVar_Click()
    tmpEvent.EventPage(curPageNum).chkPlayerVar = chkPlayerVar.value
    If chkPlayerVar.value = 0 Then
        cmbPlayerVar.enabled = False
        txtPlayerVariable.enabled = False
    Else
        cmbPlayerVar.enabled = True
        txtPlayerVariable.enabled = True
    End If
End Sub

Private Sub chkSelfSwitch_Click()
    tmpEvent.EventPage(curPageNum).chkSelfSwitch = chkSelfSwitch.value
    If chkSelfSwitch.value = 0 Then cmbSelfSwitch.enabled = False Else cmbSelfSwitch.enabled = True
End Sub

Private Sub chkStepAnim_Click()
    tmpEvent.EventPage(curPageNum).StepAnim = chkStepAnim.value
End Sub

Private Sub chkWalkAnim_Click()
    tmpEvent.EventPage(curPageNum).WalkAnim = chkWalkAnim.value
End Sub

Private Sub chkWalkThrough_Click()
    tmpEvent.EventPage(curPageNum).WalkThrough = chkWalkThrough.value
End Sub

Private Sub cmbChatBubbleType_Click()
Dim i As Long
    cmbChatBubble.Clear
    With tmpEvent.EventPage(curPageNum).Commands(curCommand)
        .TargetType = cmbChatBubbleType.ListIndex
        Select Case .TargetType
            Case 0
                cmbChatBubble.AddItem "None"
                cmbChatBubble.enabled = False
            Case TARGET_TYPE_PLAYER
                cmbChatBubble.AddItem "The Player"
                cmbChatBubble.enabled = False
            Case TARGET_TYPE_NPC
                cmbChatBubble.AddItem "None"
                For i = 1 To MAX_MAP_NPCS
                    If MapNpc(i).num > 0 Then
                        cmbChatBubble.AddItem i & ": " & Trim$(Npc(MapNpc(i).num).name)
                    Else
                        cmbChatBubble.AddItem i & ": Empty"
                    End If
                Next
                cmbChatBubble.enabled = True
            Case TARGET_TYPE_EVENT
                cmbChatBubble.AddItem "None"
                For i = 1 To map.TileData.EventCount
                    cmbChatBubble.AddItem i & ": " & map.TileData.Events(i).name
                Next
                cmbChatBubble.enabled = True
        End Select
        cmbChatBubble.ListIndex = 0
    End With
End Sub

Private Sub cmbGraphic_Click()
    If cmbGraphic.ListIndex = -1 Then Exit Sub
    tmpEvent.EventPage(curPageNum).GraphicType = cmbGraphic.ListIndex
    ' set the max on the scrollbar
    Select Case cmbGraphic.ListIndex
        Case 0 ' None
            scrlGraphic.value = 1
            scrlGraphic.Max = 1
            scrlGraphic.enabled = False
        Case 1 ' character
            scrlGraphic.Max = Count_Char
            scrlGraphic.enabled = True
        Case 2 ' Tileset
            scrlGraphic.Max = Count_Tileset
            scrlGraphic.enabled = True
    End Select
End Sub

Private Sub cmbHasItem_Click()
    If cmbHasItem.ListIndex = -1 Then Exit Sub
    tmpEvent.EventPage(curPageNum).HasItemNum = cmbHasItem.ListIndex
End Sub

Private Sub cmbMoveSpeed_Click()
    If cmbMoveSpeed.ListIndex = -1 Then Exit Sub
    tmpEvent.EventPage(curPageNum).MoveSpeed = cmbMoveSpeed.ListIndex
End Sub

Private Sub cmbMoveType_Click()
    If cmbMoveType.ListIndex = -1 Then Exit Sub
    tmpEvent.EventPage(curPageNum).MoveType = cmbMoveType.ListIndex
End Sub

Private Sub cmbPlayerVar_Click()
    If cmbPlayerVar.ListIndex = -1 Then Exit Sub
    tmpEvent.EventPage(curPageNum).PlayerVarNum = cmbPlayerVar.ListIndex
End Sub

Private Sub cmbPriority_Click()
    If cmbPriority.ListIndex = -1 Then Exit Sub
    tmpEvent.EventPage(curPageNum).Priority = cmbPriority.ListIndex
End Sub

Private Sub cmbSelfSwitch_Click()
    If cmbSelfSwitch.ListIndex = -1 Then Exit Sub
    tmpEvent.EventPage(curPageNum).SelfSwitchNum = cmbSelfSwitch.ListIndex
End Sub

Private Sub cmbTrigger_Click()
    If cmbTrigger.ListIndex = -1 Then Exit Sub
    tmpEvent.EventPage(curPageNum).Trigger = cmbTrigger.ListIndex
End Sub

Private Sub cmdAddCommand_Click()
    isEdit = False
    tabCommands.SelectedItem = tabCommands.Tabs(1)
    fraCommands.visible = True
    picCommands(1).visible = True
    picCommands(2).visible = False
End Sub

Private Sub cmdAddText_Cancel_Click()
    If Not isEdit Then fraCommands.visible = True Else fraCommands.visible = False
    fraDialogue.visible = False
    fraAddText.visible = False
End Sub

Private Sub cmdAddText_Click()
    ' reset form
    txtAddText_Text.text = vbNullString
    scrlAddText_Colour.value = 0
    optAddText_Game.value = True
    ' show
    fraDialogue.visible = True
    fraAddText.visible = True
    ' hide
    fraCommands.visible = False
End Sub

Private Sub cmdAddText_Ok_Click()
    If Not isEdit Then
        AddCommand EventType.evAddText
    Else
        EditCommand
    End If
    ' hide
    fraDialogue.visible = False
    fraAddText.visible = False
    fraCommands.visible = False
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCancelCommand_Click()
    fraCommands.visible = False
End Sub

Private Sub cmdChatBubble_Click()
    ' reset form
    txtChatBubble.text = vbNullString
    scrlChatBubble.value = 0
    cmbChatBubbleType.ListIndex = 0
    cmbChatBubble.Clear
    cmbChatBubble.AddItem "The Player"
    cmbChatBubble.enabled = False
    ' show
    fraDialogue.visible = True
    fraChatBubble.visible = True
    ' hide
    fraCommands.visible = False
End Sub

Private Sub cmdChatBubbleCancel_Click()
    If Not isEdit Then fraCommands.visible = True Else fraCommands.visible = False
    fraDialogue.visible = False
    fraChatBubble.visible = False
End Sub

Private Sub cmdChatBubbleOk_Click()
    If Not isEdit Then
        AddCommand EventType.evShowChatBubble
    Else
        EditCommand
    End If
    ' hide
    fraDialogue.visible = False
    fraChatBubble.visible = False
    fraCommands.visible = False
End Sub

Private Sub cmdClearCommand_Click()
Dim i As Long
    If tmpEvent.EventPage(curPageNum).CommandCount = 0 Then Exit Sub
    For i = 1 To tmpEvent.EventPage(curPageNum).CommandCount
        ZeroMemory ByVal VarPtr(tmpEvent.EventPage(curPageNum).Commands(i)), LenB(tmpEvent.EventPage(curPageNum).Commands(i))
    Next
    EventListCommands
End Sub

Private Sub cmdClearPage_Click()
    ZeroMemory ByVal VarPtr(tmpEvent.EventPage(curPageNum)), LenB(tmpEvent.EventPage(curPageNum))
    EventEditorLoadPage curPageNum
End Sub

Private Sub cmdCopyPage_Click()
    CopyMemory ByVal VarPtr(copyPage), ByVal VarPtr(tmpEvent.EventPage(curPageNum)), LenB(tmpEvent.EventPage(curPageNum))
    cmdPastePage.enabled = True
End Sub

Private Sub cmdDeleteCommand_Click()
Dim i As Long
    If tmpEvent.EventPage(curPageNum).CommandCount = 0 Then Exit Sub
    ZeroMemory ByVal VarPtr(tmpEvent.EventPage(curPageNum).Commands(curCommand)), LenB(tmpEvent.EventPage(curPageNum).Commands(curCommand))
    ' move everything down a page
    If tmpEvent.EventPage(curPageNum).CommandCount > 1 Then
        For i = curCommand To tmpEvent.EventPage(curPageNum).CommandCount - 1
            
        Next
    Else
        tmpEvent.EventPage(curPageNum).CommandCount = 0
    End If
End Sub

Private Sub cmdDeletePage_Click()
Dim i As Long
    ZeroMemory ByVal VarPtr(tmpEvent.EventPage(curPageNum)), LenB(tmpEvent.EventPage(curPageNum))
    ' move everything else down a notch
    If curPageNum < tmpEvent.pageCount Then
        For i = curPageNum To tmpEvent.pageCount - 1
            CopyMemory ByVal VarPtr(tmpEvent.EventPage(i)), ByVal VarPtr(tmpEvent.EventPage(i + 1)), LenB(tmpEvent.EventPage(i + 1))
        Next
    End If
    tmpEvent.pageCount = tmpEvent.pageCount - 1
    ' set the tabs
    tabPages.Tabs.Clear
    For i = 1 To tmpEvent.pageCount
        tabPages.Tabs.Add , , Str(i)
    Next
    ' set the tab back
    If curPageNum <= tmpEvent.pageCount Then
        tabPages.SelectedItem = tabPages.Tabs(curPageNum)
    Else
        tabPages.SelectedItem = tabPages.Tabs(tmpEvent.pageCount)
    End If
    ' make sure we disable
    If tmpEvent.pageCount <= 1 Then
        cmdDeletePage.enabled = False
    End If
End Sub

Private Sub cmdEditCommand_Click()
Dim i As Long
    isEdit = True
    With tmpEvent.EventPage(curPageNum).Commands(curCommand)
        Select Case .Type
            Case EventType.evAddText
                ' reset form
                txtAddText_Text.text = .text
                scrlAddText_Colour.value = .Colour
                Select Case .channel
                    Case 0
                        optAddText_Game.value = True
                    Case 1
                        optAddText_Map.value = True
                    Case 2
                        optAddText_Global.value = True
                End Select
                ' show
                fraDialogue.visible = True
                fraAddText.visible = True
            Case EventType.evShowChatBubble
                txtChatBubble.text = .text
                scrlChatBubble.value = .Colour
                cmbChatBubbleType.ListIndex = .TargetType
                cmbChatBubble.Clear
                If .TargetType = 0 Then .TargetType = 1
                Select Case .TargetType
                    Case 0
                        cmbChatBubble.AddItem "None"
                    Case TARGET_TYPE_PLAYER
                        cmbChatBubble.AddItem "The Player"
                    Case TARGET_TYPE_NPC
                        cmbChatBubble.AddItem "None"
                        For i = 1 To MAX_MAP_NPCS
                            If MapNpc(i).num > 0 Then
                                cmbChatBubble.AddItem i & ": " & Trim$(Npc(MapNpc(i).num).name)
                            Else
                                cmbChatBubble.AddItem i & ": Empty"
                            End If
                        Next
                    Case TARGET_TYPE_EVENT
                        cmbChatBubble.AddItem "None"
                        For i = 1 To map.TileData.EventCount
                            cmbChatBubble.AddItem i & ": " & map.TileData.Events(i).name
                        Next
                End Select
                If .target > 0 And .target <= cmbChatBubble.ListCount Then
                    cmbChatBubble.ListIndex = .target
                Else
                    cmbChatBubble.ListIndex = 0
                End If
                ' show
                fraDialogue.visible = True
                fraChatBubble.visible = True
            Case EventType.evPlayerVar
                ' reset form
                cmbVariable.Clear
                cmbVariable.AddItem "None"
                For i = 1 To MAX_BYTE
                    cmbVariable.AddItem i
                Next
                txtVariable.text = .Colour
                cmbVariable.ListIndex = .target
                ' show
                fraDialogue.visible = True
                fraPlayerVar.visible = True
            Case EventType.evWarpPlayer
                ' reset form
                scrlWPMap.value = .target
                scrlWPX.value = .x
                scrlWPY.value = .y
                ' show
                fraDialogue.visible = True
                fraWarpPlayer.visible = True
        End Select
    End With
End Sub

Private Sub cmdGraphicCancel_Click()
    fraGraphic.visible = False
End Sub

Private Sub cmdGraphicOK_Click()
    tmpEvent.EventPage(curPageNum).GraphicType = cmbGraphic.ListIndex
    tmpEvent.EventPage(curPageNum).Graphic = scrlGraphic.value
    tmpEvent.EventPage(curPageNum).GraphicX = GraphicSelX
    tmpEvent.EventPage(curPageNum).GraphicY = GraphicSelY
    fraGraphic.visible = False
End Sub

Private Sub cmdNewPage_Click()
Dim pageCount As Long, i As Long
    pageCount = tmpEvent.pageCount + 1
    ' redim the array
    ReDim Preserve tmpEvent.EventPage(1 To pageCount)
    tmpEvent.pageCount = pageCount
    ' set the tabs
    tabPages.Tabs.Clear
    For i = 1 To tmpEvent.pageCount
        tabPages.Tabs.Add , , Str(i)
    Next
    cmdDeletePage.enabled = True
End Sub

Private Sub cmdOk_Click()
    EventEditorOK
End Sub

Private Sub cmdPastePage_Click()
    CopyMemory ByVal VarPtr(tmpEvent.EventPage(curPageNum)), ByVal VarPtr(copyPage), LenB(tmpEvent.EventPage(curPageNum))
    EventEditorLoadPage curPageNum
End Sub

Private Sub cmdPlayerVar_Click()
Dim i As Long
    ' reset form
    cmbVariable.Clear
    cmbVariable.AddItem "None"
    For i = 1 To MAX_BYTE
        cmbVariable.AddItem i
    Next
    txtVariable.text = vbNullString
    cmbVariable.ListIndex = 0
    ' show
    fraDialogue.visible = True
    fraPlayerVar.visible = True
    ' hide
    fraCommands.visible = False
End Sub

Private Sub cmdVariableCancel_Click()
    If Not isEdit Then fraCommands.visible = True Else fraCommands.visible = False
    fraDialogue.visible = False
    fraPlayerVar.visible = False
End Sub

Private Sub cmdVariableOK_Click()
    If Not isEdit Then
        AddCommand EventType.evPlayerVar
    Else
        EditCommand
    End If
    ' hide
    fraDialogue.visible = False
    fraPlayerVar.visible = False
    fraCommands.visible = False
End Sub

Private Sub cmdWarpPlayer_Click()
    ' reset form
    scrlWPMap.value = 0
    scrlWPX.value = 0
    scrlWPY.value = 0
    ' show
    fraDialogue.visible = True
    fraWarpPlayer.visible = True
    ' hide
    fraCommands.visible = False
End Sub

Private Sub cmdWPCancel_Click()
    If Not isEdit Then fraCommands.visible = True Else fraCommands.visible = False
    fraDialogue.visible = False
    fraWarpPlayer.visible = False
End Sub

Private Sub cmdWPOkay_Click()
    If Not isEdit Then
        AddCommand EventType.evWarpPlayer
    Else
        EditCommand
    End If
    ' hide
    fraDialogue.visible = False
    fraWarpPlayer.visible = False
    fraCommands.visible = False
End Sub

Private Sub lstCommands_Click()
    curCommand = lstCommands.ListIndex + 1
End Sub

Private Sub picGraphic_Click()
    fraGraphic.width = 841
    fraGraphic.height = 609
    fraGraphic.visible = True
End Sub

Private Sub picGraphicSel_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    GraphicSelX = CLng(x) \ 32
    GraphicSelY = CLng(y) \ 32
End Sub

Private Sub scrlAddText_Colour_Change()
    lblAddText_Colour.caption = "Colour: " & GetColourString(scrlAddText_Colour.value)
End Sub

Private Sub scrlChatBubble_Change()
    lblChatBubble.caption = "Colour: " & GetColourString(scrlChatBubble.value)
End Sub

Private Sub scrlGraphic_Change()
    lblGraphic.caption = "Graphic: " & scrlGraphic.value
    tmpEvent.EventPage(curPageNum).Graphic = scrlGraphic.value
End Sub

Private Sub scrlWPMap_Change()
    lblWPMap.caption = "Map: " & scrlWPMap.value
End Sub

Private Sub scrlWPX_Change()
    lblWPX.caption = "X: " & scrlWPX.value
End Sub

Private Sub scrlWPY_Change()
    lblWPY.caption = "Y: " & scrlWPY.value
End Sub

Private Sub tabCommands_Click()
Dim i As Long
    For i = 1 To 2
        picCommands(i).visible = False
    Next
    picCommands(tabCommands.SelectedItem.index).visible = True
End Sub

Private Sub tabPages_Click()
    If tabPages.SelectedItem.index <> curPageNum Then
        curPageNum = tabPages.SelectedItem.index
        EventEditorLoadPage curPageNum
    End If
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
    tmpEvent.name = Trim$(txtName.text)
End Sub

Private Sub txtPlayerVariable_Validate(Cancel As Boolean)
    tmpEvent.EventPage(curPageNum).PlayerVariable = Val(Trim$(txtPlayerVariable.text))
End Sub
