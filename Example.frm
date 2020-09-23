VERSION 5.00
Begin VB.Form frmmain 
   Caption         =   "Properties List - ActiveX Example"
   ClientHeight    =   2850
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4335
   LinkTopic       =   "Form1"
   ScaleHeight     =   2850
   ScaleWidth      =   4335
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   350
      Left            =   1665
      TabIndex        =   2
      Top             =   2370
      Width           =   915
   End
   Begin VB.CommandButton cmdGetSel 
      Caption         =   "Get Selected"
      Height          =   350
      Left            =   255
      TabIndex        =   1
      Top             =   2370
      Width           =   1215
   End
   Begin Project1.dPropList dPropList1 
      Height          =   2130
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2490
      _ExtentX        =   4392
      _ExtentY        =   3757
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private m_SelKey As String
Private m_SelType As CrtType

Private Sub SetupList()
Dim cLst1 As New Collection
Dim cLst2 As New Collection
    
    cLst1.Add "True"
    cLst1.Add "False"
    
    cLst2.Add "Soild"
    cLst2.Add "Dotted"
    cLst2.Add "Dashed"
    cLst2.Add "Dash-Dot"
    cLst2.Add "Transparent"
    'Clear properties box
    dPropList1.Clear
    'Add some property items
    dPropList1.AddProp "Name", "A", tTextbox
    dPropList1.AddProp "BackColor", "B", tTextbox
    dPropList1.AddProp "ForeColor", "C", tTextbox
    dPropList1.AddProp "Caption", "D", tTextbox
    dPropList1.AddProp "Enabled", "E", tComboBox
    dPropList1.AddProp "BorderStyle", "F", tComboBox
    dPropList1.AddProp "Height", "G", tTextbox
    dPropList1.AddProp "Width", "h", tTextbox
    'Set default property data
    
    dPropList1.SetPropItemValue "A", "ShpButton", tTextbox
    dPropList1.SetPropItemValue "B", "#000000", tTextbox
    dPropList1.SetPropItemValue "C", "#fff000", tTextbox
    dPropList1.SetPropItemValue "D", "E&xit", tTextbox
    dPropList1.SetPropItemValue "E", cLst1, tComboBox
    dPropList1.SetPropItemValue "F", cLst2, tComboBox
    dPropList1.SetPropItemValue "G", 350, tTextbox
    dPropList1.SetPropItemValue "H", 1155, tTextbox
    
    Set cLst1 = Nothing
    Set cLst2 = Nothing
End Sub

Private Sub cmdExit_Click()
    dPropList1.Clear
    Unload frmmain
End Sub

Private Sub cmdGetSel_Click()
    MsgBox dPropList1.GetPropItemValue(m_SelKey, m_SelType), vbInformation
End Sub

Private Sub dPropList1_PropClick(sKey As String, ItemProp As CrtType)
    m_SelKey = sKey
    m_SelType = ItemProp
End Sub

Private Sub Form_Load()
    Call SetupList
End Sub

Private Sub Form_Unload(Cancel As Integer)
    dPropList1.Clear
    Set frmmain = Nothing
End Sub
