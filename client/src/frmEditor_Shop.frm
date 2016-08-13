VERSION 5.00
Begin VB.Form frmEditor_Shop 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Shop Editor"
   ClientHeight    =   9075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8775
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
   Icon            =   "frmEditor_Shop.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   605
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   5280
      TabIndex        =   20
      Top             =   8520
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Shop Properties"
      Height          =   8295
      Left            =   3360
      TabIndex        =   5
      Top             =   120
      Width           =   5295
      Begin VB.CommandButton cmdCopy 
         Caption         =   "Cut"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   7560
         Width           =   2415
      End
      Begin VB.CommandButton cmdPaste 
         Caption         =   "Paste"
         Height          =   255
         Left            =   2760
         TabIndex        =   22
         Top             =   7560
         Width           =   2415
      End
      Begin VB.CommandButton cmdDeleteTrade 
         Caption         =   "Delete"
         Height          =   255
         Left            =   2760
         TabIndex        =   21
         Top             =   7920
         Width           =   2415
      End
      Begin VB.HScrollBar scrlBuy 
         Height          =   255
         Left            =   120
         Max             =   1000
         Min             =   1
         TabIndex        =   19
         Top             =   840
         Value           =   100
         Width           =   5055
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   720
         TabIndex        =   12
         Top             =   240
         Width           =   4455
      End
      Begin VB.ListBox lstTradeItem 
         Height          =   5460
         ItemData        =   "frmEditor_Shop.frx":3332
         Left            =   120
         List            =   "frmEditor_Shop.frx":334E
         TabIndex        =   11
         Top             =   2040
         Width           =   5055
      End
      Begin VB.ComboBox cmbItem 
         Height          =   300
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1320
         Width           =   3135
      End
      Begin VB.TextBox txtItemValue 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4560
         TabIndex        =   9
         Text            =   "1"
         Top             =   1320
         Width           =   615
      End
      Begin VB.ComboBox cmbCostItem 
         Height          =   300
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1680
         Width           =   3135
      End
      Begin VB.TextBox txtCostValue 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4560
         TabIndex        =   7
         Text            =   "1"
         Top             =   1680
         Width           =   615
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   7920
         Width           =   2415
      End
      Begin VB.Label lblBuy 
         AutoSize        =   -1  'True
         Caption         =   "Buy Rate: 100%"
         Height          =   180
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   180
         Left            =   120
         TabIndex        =   17
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   495
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Item:"
         Height          =   180
         Left            =   120
         TabIndex        =   16
         Top             =   1320
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Value:"
         Height          =   180
         Left            =   3960
         TabIndex        =   15
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Price:"
         Height          =   180
         Left            =   120
         TabIndex        =   14
         Top             =   1680
         Width           =   450
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Value:"
         Height          =   180
         Left            =   3960
         TabIndex        =   13
         Top             =   1680
         UseMnemonic     =   0   'False
         Width           =   495
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Shop List"
      Height          =   8295
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3135
      Begin VB.ListBox lstIndex 
         Height          =   7980
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.CommandButton cmdArray 
      Caption         =   "Change Array Size"
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   8520
      Width           =   2895
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   8520
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6960
      TabIndex        =   1
      Top             =   8520
      Width           =   1575
   End
End
Attribute VB_Name = "frmEditor_Shop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private tmpTrade As TradeItemRec

Private Sub cmdCopy_Click()
Dim index As Long

    index = lstTradeItem.ListIndex + 1
    tmpTrade.Item = Shop(EditorIndex).TradeItem(index).Item
    tmpTrade.ItemValue = Shop(EditorIndex).TradeItem(index).ItemValue
    tmpTrade.CostItem = Shop(EditorIndex).TradeItem(index).CostItem
    tmpTrade.CostValue = Shop(EditorIndex).TradeItem(index).CostValue
    
    cmdDeleteTrade_Click
End Sub

Private Sub cmdPaste_Click()
Dim index As Long, tmpPos As Long
    tmpPos = lstTradeItem.ListIndex

    index = lstTradeItem.ListIndex + 1
    Shop(EditorIndex).TradeItem(index).Item = tmpTrade.Item
    Shop(EditorIndex).TradeItem(index).ItemValue = tmpTrade.ItemValue
    Shop(EditorIndex).TradeItem(index).CostItem = tmpTrade.CostItem
    Shop(EditorIndex).TradeItem(index).CostValue = tmpTrade.CostValue
    
    UpdateShopTrade tmpPos
End Sub

Private Sub cmdSave_Click()

    If LenB(Trim$(txtName)) = 0 Then
        Call MsgBox("Name required.")
    Else
        Call ShopEditorOk
    End If

End Sub

Private Sub cmdCancel_Click()
    Call ShopEditorCancel
End Sub

Private Sub cmdUpdate_Click()
    Dim index As Long
    Dim tmpPos As Long
    tmpPos = lstTradeItem.ListIndex
    index = lstTradeItem.ListIndex + 1

    If index = 0 Then Exit Sub

    With Shop(EditorIndex).TradeItem(index)
        .Item = cmbItem.ListIndex
        .ItemValue = Val(txtItemValue.text)
        .CostItem = cmbCostItem.ListIndex
        .CostValue = Val(txtCostValue.text)
    End With

    UpdateShopTrade tmpPos
End Sub

Private Sub cmdDeleteTrade_Click()
    Dim index As Long
    Dim tmpPos As Long
    tmpPos = lstTradeItem.ListIndex
    index = lstTradeItem.ListIndex + 1

    If index = 0 Then Exit Sub

    With Shop(EditorIndex).TradeItem(index)
        .Item = 0
        .ItemValue = 0
        .CostItem = 0
        .CostValue = 0
    End With

    UpdateShopTrade tmpPos
End Sub

Private Sub lstIndex_Click()
    ShopEditorInit
End Sub

Private Sub scrlBuy_Change()
    lblBuy.caption = "Buy Rate: " & scrlBuy.value & "%"
    Shop(EditorIndex).BuyRate = scrlBuy.value
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
    Dim tmpIndex As Long

    If EditorIndex = 0 Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    Shop(EditorIndex).name = Trim$(txtName.text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Shop(EditorIndex).name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
End Sub
