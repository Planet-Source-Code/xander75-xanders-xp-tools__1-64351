VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl XandersXPListBox 
   BackStyle       =   0  'Transparent
   ClientHeight    =   1905
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2220
   ScaleHeight     =   1905
   ScaleWidth      =   2220
   ToolboxBitmap   =   "XandersXPListBox.ctx":0000
   Begin MSComctlLib.ListView lvwCombo 
      Height          =   1650
      Left            =   15
      TabIndex        =   0
      Top             =   15
      Width           =   2025
      _ExtentX        =   3572
      _ExtentY        =   2910
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      AllowReorder    =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.Shape shpBorder 
      BorderColor     =   &H00DC9670&
      Height          =   1695
      Left            =   0
      Top             =   0
      Width           =   2055
   End
End
Attribute VB_Name = "XandersXPListBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim m_BorderColorOver As OLE_COLOR
Dim SelectedBorderColor As OLE_COLOR

' Set the Margins within the Textbox
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const LVM_FIRST = &H1000
Private Const LVM_SETCOLUMNWIDTH = (LVM_FIRST + 30)
Private Const LVSCW_AUTOSIZE = -1
Private Const LVSCW_AUTOSIZE_USEHEADER = -2

' Events
Event Change()
Event Click()
Event DblClick()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Public Function AddColumn(Columns As Integer, Headers As String) As Variant
    If Trim(Headers) = "" Then
        For i = 0 To Columns - 1
            lvwCombo.ColumnHeaders.Add , , ""
        Next
    Else
        Dim sItems() As String
        sItems() = Split(Headers, vbTab)
        
        For i = 0 To Columns - 1
            lvwCombo.ColumnHeaders.Add , , sItems(i)
        Next
    End If
End Function

Public Function AddItem(Item As String) As Variant
    
    Dim sItems() As String
    sItems() = Split(Item, vbTab)
    
    If lvwCombo.ColumnHeaders.Count = 0 Then Call AddColumn(1, "")
    
    With lvwCombo.ListItems
        .Add , , sItems(0)
    End With
    ListCount = lvwCombo.ListItems.Count
    
    For i = 1 To lvwCombo.ColumnHeaders.Count - 1
        lvwCombo.ListItems.Item(ListCount).SubItems(i) = sItems(i)
    Next
    
End Function

Public Function AutoColumnWitdh() As Variant
    Dim Column As Long
    Dim Counter As Long
    Counter = 0
    For Column = Counter To lvwCombo.ColumnHeaders.Count - 2
       SendMessage lvwCombo.hwnd, LVM_SETCOLUMNWIDTH, _
       Column, LVSCW_AUTOSIZE_USEHEADER
    Next
End Function

Public Function Clear() As Variant
    lvwCombo.ListItems.Clear
End Function

Public Function ClearColumnHeaders() As Variant
    lvwCombo.ColumnHeaders.Clear
End Function

Public Sub AboutBox()
    frmAbout.Show vbModal, Me
End Sub

Private Sub lvwCombo_GotFocus()
    SelectedBorderColor = shpBorder.BorderColor
    shpBorder.BorderColor = m_BorderColorOver
End Sub

Private Sub lvwCombo_LostFocus()
    shpBorder.BorderColor = SelectedBorderColor
End Sub

Public Property Get BackColor() As OLE_COLOR
    BackColor = lvwCombo.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    lvwCombo.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

Public Property Get BorderColor() As OLE_COLOR
    BorderColor = shpBorder.BorderColor
End Property

Public Property Let BorderColor(ByVal New_BorderColor As OLE_COLOR)
    shpBorder.BorderColor() = New_BorderColor
    PropertyChanged "BorderColor"
End Property

Public Property Get BorderColorOver() As OLE_COLOR
    BorderColorOver = m_BorderColorOver
End Property

Public Property Let BorderColorOver(ByVal New_BorderColorOver As OLE_COLOR)
    m_BorderColorOver = New_BorderColorOver
    PropertyChanged "BorderColorOver"
End Property

Public Property Get Enabled() As Boolean
    Enabled = lvwCombo.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    lvwCombo.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

' Return the font.
Public Property Get Font() As Font
    Set Font = lvwCombo.Font
End Property

' Set the font.
Public Property Set Font(ByVal New_Font As Font)
    Set lvwCombo.Font = New_Font
    PropertyChanged "Font"
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = lvwCombo.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    lvwCombo.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

Public Property Get FullRowSelect() As Boolean
    FullRowSelect = lvwCombo.FullRowSelect
End Property

Public Property Let FullRowSelect(ByVal New_FullRowSelect As Boolean)
    lvwCombo.FullRowSelect() = New_FullRowSelect
    PropertyChanged "FullRowSelect"
End Property

Public Property Get Gridlines() As Boolean
    Gridlines = lvwCombo.Gridlines
End Property

Public Property Let Gridlines(ByVal New_Gridlines As Boolean)
    lvwCombo.Gridlines() = New_Gridlines
    PropertyChanged "Gridlines"
End Property

Public Property Get HideColumnHeaders() As Boolean
    HideColumnHeaders = lvwCombo.HideColumnHeaders
End Property

Public Property Let HideColumnHeaders(ByVal New_Header As Boolean)
    lvwCombo.HideColumnHeaders() = New_Header
    PropertyChanged "HideColumnHeaders"
End Property

Public Property Get MouseIcon() As Picture
    Set MouseIcon = lvwCombo.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set lvwCombo.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

Public Property Get MousePointer() As Integer
    MousePointer = lvwCombo.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As Integer)
    lvwCombo.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

'Private Sub lvwCombo_Change()
'    RaiseEvent Change
'End Sub
'
'Private Sub lvwCombo_Click()
'    RaiseEvent Click
'End Sub
'
'Private Sub lvwCombo_DblClick()
'    RaiseEvent DblClick
'End Sub
'
'Private Sub lvwCombo_KeyDown(KeyCode As Integer, Shift As Integer)
'    RaiseEvent KeyDown(KeyCode, Shift)
'End Sub
'
'Private Sub lvwCombo_KeyPress(KeyAscii As Integer)
'    RaiseEvent KeyPress(KeyAscii)
'End Sub
'
'Private Sub lvwCombo_KeyUp(KeyCode As Integer, Shift As Integer)
'    RaiseEvent KeyUp(KeyCode, Shift)
'End Sub
'
'Private Sub lvwCombo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    RaiseEvent MouseDown(Button, Shift, X, Y)
'End Sub
'
'Private Sub lvwCombo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    RaiseEvent MouseMove(Button, Shift, X, Y)
'End Sub
'
'Private Sub lvwCombo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    RaiseEvent MouseUp(Button, Shift, X, Y)
'End Sub

Private Sub UserControl_Initialize()
    m_BorderColorOver = &H96E7&
End Sub

Private Sub UserControl_Resize()
    shpBorder.Height = UserControl.Height
    shpBorder.Width = UserControl.Width
    lvwCombo.Top = shpBorder.Top + 15
    lvwCombo.Left = shpBorder.Left + 15
    lvwCombo.Height = shpBorder.Height - 30
    lvwCombo.Width = shpBorder.Width - 30
End Sub

' Load saved properties.
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    lvwCombo.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    lvwCombo.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    shpBorder.BorderColor = PropBag.ReadProperty("BorderColor", &HB99D7F)
    m_BorderColorOver = PropBag.ReadProperty("BorderColorOver", m_def_BorderColorOver)
    lvwCombo.Enabled = PropBag.ReadProperty("Enabled", 0)
    Set lvwCombo.Font = PropBag.ReadProperty("Font", Ambient.Font)
    lvwCombo.FullRowSelect = PropBag.ReadProperty("FullRowSelect", 0)
    lvwCombo.Gridlines = PropBag.ReadProperty("Gridlines", 0)
    lvwCombo.HideColumnHeaders = PropBag.ReadProperty("HideColumnHeaders", 0)
    Set lvwCombo.MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    lvwCombo.MousePointer = PropBag.ReadProperty("MousePointer", 0)
End Sub

' Save properties.
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", lvwCombo.BackColor, &H80000005)
    Call PropBag.WriteProperty("ForeColor", lvwCombo.ForeColor, &H80000008)
    Call PropBag.WriteProperty("BorderColor", shpBorder.BorderColor, &HB99D7F)
    Call PropBag.WriteProperty("BorderColorOver", m_BorderColorOver, m_def_BorderColorOver)
    Call PropBag.WriteProperty("Enabled", lvwCombo.Enabled, 0)
    Call PropBag.WriteProperty("Font", lvwCombo.Font, Ambient.Font)
    Call PropBag.WriteProperty("FullRowSelect", lvwCombo.FullRowSelect, 0)
    Call PropBag.WriteProperty("Gridlines", lvwCombo.Gridlines, 0)
    Call PropBag.WriteProperty("HideColumnHeaders", lvwCombo.HideColumnHeaders, 0)
    Call PropBag.WriteProperty("MouseIcon", lvwCombo.MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", lvwCombo.MousePointer, 0)
End Sub
