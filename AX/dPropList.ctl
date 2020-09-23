VERSION 5.00
Begin VB.UserControl dPropList 
   ClientHeight    =   2835
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3540
   ScaleHeight     =   189
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   236
   ToolboxBitmap   =   "dPropList.ctx":0000
   Begin VB.PictureBox pBase 
      BackColor       =   &H00FFFFFF&
      Height          =   2295
      Left            =   0
      ScaleHeight     =   149
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   220
      TabIndex        =   0
      Top             =   0
      Width           =   3360
      Begin VB.PictureBox pHolder 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   2145
         Left            =   0
         ScaleHeight     =   2145
         ScaleWidth      =   2220
         TabIndex        =   2
         Top             =   0
         Width           =   2220
         Begin VB.ComboBox CboA 
            Height          =   315
            Index           =   0
            Left            =   915
            TabIndex        =   5
            Top             =   330
            Visible         =   0   'False
            Width           =   1200
         End
         Begin VB.TextBox TxtA 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   915
            TabIndex        =   4
            Top             =   30
            Visible         =   0   'False
            Width           =   1185
         End
         Begin VB.Label LblA 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "BackColor"
            Height          =   195
            Index           =   0
            Left            =   30
            TabIndex        =   3
            Top             =   75
            Visible         =   0   'False
            Width           =   735
         End
      End
      Begin VB.VScrollBar vBar1 
         Height          =   285
         Left            =   2220
         TabIndex        =   1
         Top             =   0
         Width           =   255
      End
   End
End
Attribute VB_Name = "dPropList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private TmpCbo As String

Enum CrtType
    tTextbox = 1
    tComboBox = 2
End Enum

Event PropChanged(sKey As String, ItemProp As CrtType, ItemValue As String)
Event PropClick(sKey As String, ItemProp As CrtType)

Private Function IndexOfControl(sKey As String, Optional ItemProp As CrtType = tTextbox) As Integer
Dim c As Control
Dim Idx As Integer
Dim sTypeName As String
    
    'This function is used to return the index of a control
    
    Idx = -1
    
    If (ItemProp = tComboBox) Then sTypeName = "COMBOBOX" 'Return the index of a combobox control
    If (ItemProp = tTextbox) Then sTypeName = "TEXTBOX" ' Return the index of a textbox control
    
    For Each c In UserControl.Controls
        'Only serach the controls for TxtA and CboA
        If (c.Name = "TxtA") Or (c.Name = "CboA") Then
            'Only index's greator than zero to be checked
            If (c.Index > 0) Then
                'Check to see if the typename matches sTypeName
                If UCase(TypeName(c)) = sTypeName Then
                    'Compare the Tag with the propery Key
                    If (StrComp(sKey, c.Tag, vbTextCompare) = 0) Then
                        'Return the controls index
                        Idx = c.Index
                        'Exit loop
                        Exit For
                    End If
                End If
            End If
        End If
    Next c
    'Return the found index
    IndexOfControl = Idx
    
    'Clear up
    sTypeName = vbNullString
    Set c = Nothing
    Idx = 0
End Function

Public Sub Clear()
Dim c As Control
    'Used to unload all the control arrays.
    For Each c In UserControl.Controls
        If (c.Name = "LblA") Or (c.Name = "TxtA") Or (c.Name = "CboA") Then
            If (c.Index > 0) Then
                'Unload all controls except the first one.
                Unload c
            End If
        End If
    Next c
    Set c = Nothing
End Sub

Public Function GetPropItemValue(sKey As String, Optional ItemProp As CrtType = tTextbox) As String
Dim cIdx As Integer
On Error GoTo TErr:
    'Get the index of the control
    cIdx = IndexOfControl(sKey, ItemProp)
    'Check for a vaild return index
    If (cIdx = -1) Then
        Err.Raise 9
        Exit Function
    Else
        If (ItemProp = tTextbox) Then
            'Return textbox value.
            GetPropItemValue = TxtA(cIdx).Text
        End If
        If (ItemProp = tComboBox) Then
            'Return comboxbox value.
            GetPropItemValue = CboA(cIdx).Text
        End If
    End If
    
    'Clear up
    cIdx = 0
    Exit Function
TErr:
    If Err Then Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub SetPropItemValue(sKey As String, lValues, Optional ItemProp As CrtType = tTextbox)
Dim cIdx As Integer
On Error GoTo TErr:

    'Get the index of the control
    cIdx = IndexOfControl(sKey, ItemProp)
    'Check for a vaild return index
    If (cIdx = -1) Then
        Err.Raise 9
        Exit Sub
    Else
        'We are dealing with a textbox control
        'Set the textbox's data
        If (ItemProp = tTextbox) Then
            'Set up the control to assign the text to.
            TxtA(cIdx).Text = lValues
        End If
        
        'Set the items for the combobox
        If (ItemProp = tComboBox) Then
            'Clear the combobox
            CboA(cIdx).Clear
            'Add the items from the collection to the combobox
            For Each Item In lValues
                CboA(cIdx).AddItem Item
            Next Item
            'Set the first top item
            CboA(cIdx).ListIndex = 0
        End If
    End If
    
    'Clear up
    cIdx = 0
    Item = ""
    Exit Sub
    
TErr:
    If Err Then Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub AddProp(mCaption As String, sKey As String, ItemProp As CrtType)
Dim LCount As Integer
Dim ItemTop As Long
Dim ItemCount As Integer

    'Used to add the property items.
    
    LCount = LblA.Count 'Return the number of labels
    ItemCount = LCount
    
    Load LblA(LCount)
    
    If (LCount = 1) Then
        'Deafult top for the first label
        ItemTop = 75
    Else
        'All the rest.
        ItemTop = LblA(LCount - 1).Top + LblA(LCount).Height + 120
    End If
    
    LblA(LCount).Top = ItemTop
    LblA(LCount).Caption = mCaption
    LblA(LCount).Visible = True

    Select Case ItemProp
        Case tTextbox 'Text Field
            LCount = TxtA.Count
            Load TxtA(ItemCount)
            TxtA(ItemCount).Top = ItemTop - 35
            TxtA(ItemCount).Visible = True
            TxtA(ItemCount).Tag = sKey
        Case tComboBox 'ComboBox
            LCount = CboA.Count
            Load CboA(ItemCount)
            CboA(ItemCount).Top = ItemTop - 35
            CboA(ItemCount).Visible = True
            CboA(ItemCount).Tag = sKey
    End Select
    
    LCount = 0
    ItemCount = 0
    pHolder.Height = (ItemTop \ Screen.TwipsPerPixelY) + 35
    UserControl_Resize
End Sub

Private Sub CboA_Click(Index As Integer)
    RaiseEvent PropClick(CboA(Index).Tag, tComboBox)
    RaiseEvent PropChanged(CboA(Index).Tag, tComboBox, CboA(Index).Text)
End Sub

Private Sub CboA_LostFocus(Index As Integer)
    RaiseEvent PropChanged(CboA(Index).Tag, tComboBox, CboA(Index).Text)
End Sub

Private Sub TxtA_Click(Index As Integer)
    RaiseEvent PropClick(TxtA(Index).Tag, tTextbox)
End Sub

Private Sub TxtA_KeyPress(Index As Integer, KeyAscii As Integer)
    If (KeyAscii = 13) Then
        RaiseEvent PropChanged(TxtA(Index).Tag, tTextbox, TxtA(Index).Text)
        KeyAscii = 0
    End If
End Sub

Private Sub TxtA_LostFocus(Index As Integer)
    RaiseEvent PropChanged(TxtA(Index).Tag, tTextbox, TxtA(Index).Text)
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
Dim vMax As Integer
    'Resize eveything
    pBase.Width = UserControl.ScaleWidth
    pBase.Height = UserControl.ScaleHeight
    pHolder.Width = (pBase.ScaleWidth - vBar1.Width)
    '
    vBar1.Left = (pBase.ScaleWidth - vBar1.Width)
    vBar1.Height = (pBase.ScaleHeight)
    vMax = (pHolder.Height - pBase.Height)
    
    If (vMax < 0) Then vMax = 0
    vBar1.Max = vMax
End Sub

Private Sub vBar1_Change()
    pHolder.Top = -vBar1.Value
End Sub

Private Sub vBar1_Scroll()
    vBar1_Change
End Sub

