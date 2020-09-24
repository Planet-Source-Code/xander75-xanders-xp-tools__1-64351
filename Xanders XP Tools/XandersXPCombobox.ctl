VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl XandersXPCombobox 
   BackStyle       =   0  'Transparent
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "XandersXPCombobox.ctx":0000
   Begin MSComctlLib.ListView lvwCombo 
      Height          =   1650
      Left            =   15
      TabIndex        =   1
      Top             =   375
      Visible         =   0   'False
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
   Begin VB.TextBox txtXText 
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   15
      TabIndex        =   0
      Top             =   15
      Width           =   1080
   End
   Begin VB.Shape shpBorder 
      BorderColor     =   &H00DC9670&
      Height          =   1695
      Left            =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Image imgArrow 
      Height          =   150
      Left            =   1080
      Picture         =   "XandersXPCombobox.ctx":0312
      Top             =   120
      Width           =   225
   End
   Begin VB.Image picButton 
      Height          =   255
      Left            =   1080
      Picture         =   "XandersXPCombobox.ctx":0668
      Stretch         =   -1  'True
      Top             =   0
      Width           =   225
   End
   Begin VB.Shape shBorder 
      BorderColor     =   &H00B99D7F&
      Height          =   330
      Left            =   0
      Top             =   0
      Width           =   1365
   End
End
Attribute VB_Name = "XandersXPCombobox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim m_AutoSelect As Boolean
Dim UserResize As Boolean
Dim m_BorderColorOver As OLE_COLOR
Dim SelectedBorderColor As OLE_COLOR

' Set the Margins within the Textbox
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const EM_SETMARGINS = &HD3
Private Const EC_LEFTMARGIN = &H1
Private Const EC_RIGHTMARGIN = &H2

'Private Declare Function SendWidthMessage Lib "user32.dll" _
'  Alias "SendWidthMessageA" (ByVal hWnd As Long, _
'  ByVal Msg As Long, ByVal wParam As Long, _
'  ByVal lParam As Long) As Long
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
       SendMessage lvwCombo.hWnd, LVM_SETCOLUMNWIDTH, _
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

Private Sub imgArrow_Click()
    If lvwCombo.Visible = False Then
        shpBorder.Visible = True
        lvwCombo.Visible = True
        If shBorder.Width <= shpBorder.Width Then
            shpBorder.Top = shBorder.Height + 30
            UserControl.Height = shBorder.Height + shpBorder.Height + 30
            UserControl.Width = shpBorder.Width
            If lvwCombo.ListItems.Count = 0 Then
                shpBorder.Height = 330
                lvwCombo.Height = 300
            End If
        Else
            shpBorder.Top = shBorder.Height + 30
            UserControl.Height = shBorder.Height + shpBorder.Height + 30
            UserControl.Width = shBorder.Width
            If lvwCombo.ListItems.Count = 0 Then
                shpBorder.Height = 330
                lvwCombo.Height = 300
            End If
        End If
    ElseIf lvwCombo.Visible = True Then
        shpBorder.Visible = False
        lvwCombo.Visible = False
        UserResize = True
        If shBorder.Width <= shpBorder.Width Then
            UserControl.Height = shBorder.Height
            UserControl.Width = shBorder.Width
        Else
            UserControl.Height = shBorder.Height
'            UserControl.Width = shBorder.Width
        End If
    End If
    
End Sub

Private Sub lvwCombo_Click()
    txtXText.Text = lvwCombo.ListItems.Item(lvwCombo.SelectedItem.Index)
    shpBorder.Visible = False
    lvwCombo.Visible = False
    UserResize = True
End Sub

Private Sub lvwCombo_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If ColumnHeader = "" Then
        lvwCombo.SortKey = 0
    Else
        lvwCombo.SortKey = ColumnHeader.Index - 1
    End If
End Sub

Private Sub picButton_Click()
    Call imgArrow_Click
End Sub

Private Sub txtXText_GotFocus()
    SelectedBorderColor = shBorder.BorderColor
    shBorder.BorderColor = m_BorderColorOver
End Sub

Private Sub txtXText_LostFocus()
    If shpBorder.Visible = False Then
        shBorder.BorderColor = SelectedBorderColor '&HB99D7F
        shpBorder.Visible = False
        lvwCombo.Visible = False
        UserResize = True
        Call UserControl_Resize
    End If
End Sub

Private Sub UserControl_Initialize()
    m_BorderColorOver = &H96E7&

    Dim left_margin As Integer
    Dim right_margin As Integer
    Dim long_value As Long

    left_margin = CInt(4)
    right_margin = CInt(4)
    long_value = right_margin * &H10000 + left_margin

    SendMessage txtXText.hWnd, _
        EM_SETMARGINS, _
        EC_LEFTMARGIN Or EC_RIGHTMARGIN, _
        long_value
End Sub

Private Sub UserControl_InitProperties()
    'txtXText.Text = UserControl.Extender.Name
    UserResize = False
End Sub

Private Sub UserControl_Resize()
    If lvwCombo.Visible = False Then
        If UserResize = False Then
            shBorder.Height = UserControl.Height
            shBorder.Width = UserControl.Width
        End If
        
        txtXText.Height = shBorder.Height - 25
        txtXText.Left = shBorder.Left + 10
        txtXText.Top = shBorder.Top + 15
        txtXText.Width = shBorder.Width - picButton.Width - 30
        
        picButton.Top = shBorder.Top + 15
        picButton.Left = shBorder.Width - 240
        picButton.Height = shBorder.Height - 30
        
        imgArrow.Top = (picButton.Height / 2) - (imgArrow.Height / 2) + 30
        imgArrow.Left = picButton.Left
        
        shpBorder.Top = shBorder.Height + 30
        shpBorder.Width = shBorder.Width
        lvwCombo.Top = shpBorder.Top + 15
        lvwCombo.Left = shpBorder.Left + 15
        lvwCombo.Width = shpBorder.Width - 30
    Else
        UserControl.Height = shBorder.Height + shpBorder.Height + 45
    End If
End Sub

Public Property Get Alignment() As AlignmentConstants
    Alignment = txtXText.Alignment
End Property

Public Property Let Alignment(ByVal New_Alignment As AlignmentConstants)
    txtXText.Alignment() = New_Alignment
    PropertyChanged "Alignment"
End Property

Public Property Get AutoSelect() As Boolean
    AutoSelect = m_AutoSelect
End Property

Public Property Let AutoSelect(ByVal New_AutoSelect As Boolean)
    m_AutoSelect = New_AutoSelect
    PropertyChanged "AutoSelect"
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = txtXText.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    txtXText.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

Public Property Get BorderColor() As OLE_COLOR
    BorderColor = shBorder.BorderColor
End Property

Public Property Let BorderColor(ByVal New_BorderColor As OLE_COLOR)
    shBorder.BorderColor() = New_BorderColor
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
    Enabled = txtXText.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    txtXText.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

' Return the font.
Public Property Get Font() As Font
    Set Font = txtXText.Font
End Property

' Set the font.
Public Property Set Font(ByVal New_Font As Font)
    Set txtXText.Font = New_Font
    PropertyChanged "Font"
End Property

Public Property Get FontBold() As Boolean
    FontBold = txtXText.FontBold
End Property

Public Property Let FontBold(ByVal New_FontBold As Boolean)
    txtXText.FontBold() = New_FontBold
    PropertyChanged "FontBold"
End Property

Public Property Get FontItalic() As Boolean
    FontItalic = txtXText.FontItalic
End Property

Public Property Let FontItalic(ByVal New_FontItalic As Boolean)
    txtXText.FontItalic() = New_FontItalic
    PropertyChanged "FontItalic"
End Property

Public Property Get FontName() As String
    FontName = txtXText.FontName
End Property

Public Property Let FontName(ByVal New_FontName As String)
    txtXText.FontName() = New_FontName
    PropertyChanged "FontName"
End Property

Public Property Get FontSize() As Single
    FontSize = txtXText.FontSize
End Property

Public Property Let FontSize(ByVal New_FontSize As Single)
    txtXText.FontSize() = New_FontSize
    PropertyChanged "FontSize"
End Property

Public Property Get FontStrikeOut() As Boolean
    FontStrikeOut = txtXText.FontStrikethru
End Property

Public Property Let FontStrikeOut(ByVal New_FontStrikeOut As Boolean)
    txtXText.FontStrikethru() = New_FontStrikeOut
    PropertyChanged "FontStrikeOut"
End Property

Public Property Get FontUnderline() As Boolean
    FontUnderline = txtXText.FontUnderline
End Property

Public Property Let FontUnderline(ByVal New_FontUnderline As Boolean)
    txtXText.FontUnderline() = New_FontUnderline
    PropertyChanged "FontUnderline"
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = txtXText.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    txtXText.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
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

Public Property Get MaxLength() As Long
    MaxLength = txtXText.MaxLength
End Property

Public Property Let MaxLength(ByVal New_MaxLength As Long)
    txtXText.MaxLength() = New_MaxLength
    PropertyChanged "MaxLength"
End Property

Public Property Get MouseIcon() As Picture
    Set MouseIcon = txtXText.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set txtXText.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

Public Property Get MousePointer() As Integer
    MousePointer = txtXText.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As Integer)
    txtXText.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

Public Property Get PasswordChar() As String
    PasswordChar = txtXText.PasswordChar
End Property

Public Property Let PasswordChar(ByVal New_PasswordChar As String)
    txtXText.PasswordChar() = New_PasswordChar
    PropertyChanged "PasswordChar"
End Property

' Return the Text.
Public Property Get Text() As String
    Text = txtXText.Text
End Property

' Set the Text.
Public Property Let Text(ByVal New_Text As String)
    txtXText.Text() = New_Text
    PropertyChanged "Text"
End Property

'Public Property Get MultiLine() As Boolean
'    MultiLine = txtXText.MultiLine
'End Property
'
'Public Property Let MultiLine(ByVal New_MultiLine As Boolean)
'    txtXText.MultiLine() = New_MultiLine
'    PropertyChanged "MultiLine"
'End Property


Private Sub txtXText_Change()
    RaiseEvent Change
End Sub

Private Sub txtXText_Click()
    RaiseEvent Click
End Sub

Private Sub txtXText_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub txtXText_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub txtXText_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub txtXText_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub txtXText_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub txtXText_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub txtXText_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

' Load saved properties.
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    txtXText.Alignment = PropBag.ReadProperty("Alignment", txtXText.Alignment)
    m_AutoSelect = PropBag.ReadProperty("AutoSelect", m_def_AutoSelect)
    txtXText.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    txtXText.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    shBorder.BorderColor = PropBag.ReadProperty("BorderColor", &HB99D7F)
    m_BorderColorOver = PropBag.ReadProperty("BorderColorOver", m_def_BorderColorOver)
    txtXText.Enabled = PropBag.ReadProperty("Enabled", 0)
    Set txtXText.Font = PropBag.ReadProperty("Font", Ambient.Font)
    txtXText.FontBold = PropBag.ReadProperty("FontBold", 0)
    txtXText.FontItalic = PropBag.ReadProperty("FontItalic", 0)
    txtXText.FontName = PropBag.ReadProperty("FontName", "")
    txtXText.FontSize = PropBag.ReadProperty("FontSize", 0)
    txtXText.FontStrikethru = PropBag.ReadProperty("FontStrikeOut", 0)
    txtXText.FontUnderline = PropBag.ReadProperty("FontUnderline", 0)
    lvwCombo.Gridlines = PropBag.ReadProperty("Gridlines", 0)
    lvwCombo.HideColumnHeaders = PropBag.ReadProperty("HideColumnHeaders", 0)
    txtXText.MaxLength = PropBag.ReadProperty("MaxLength", 0)
    Set txtXText.MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    txtXText.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    txtXText.PasswordChar = PropBag.ReadProperty("PasswordChar", "")
    Set txtXText.Font = PropBag.ReadProperty("Font", Ambient.Font)
    txtXText.Text = PropBag.ReadProperty("Text", UserControl.Name)
'    txtXText.MultiLine = PropBag.ReadProperty("MultiLine", 0)

    If m_AutoSelect = True Then
        txtXText.SelStart = 0
        txtXText.SelLength = Len(txtXText.Text)
    End If

    If m_ComputerInfo = None Then
        txtXText.Text = UserControl.Extender.Name
    ElseIf m_ComputerInfo = Computername Then
        txtXText.Text = Environ("ComputerName")
    ElseIf m_ComputerInfo = Username Then
        txtXText.Text = Environ("UserName")
    End If
    
End Sub

' Save properties.
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Alignment", txtXText.Alignment, "Alignment")
    Call PropBag.WriteProperty("AutoSelect", m_AutoSelect, m_def_AutoSelect)
    Call PropBag.WriteProperty("BackColor", txtXText.BackColor, &H80000005)
    Call PropBag.WriteProperty("ForeColor", txtXText.ForeColor, &H80000008)
    Call PropBag.WriteProperty("BorderColor", shBorder.BorderColor, &HB99D7F)
    Call PropBag.WriteProperty("BorderColorOver", m_BorderColorOver, m_def_BorderColorOver)
    Call PropBag.WriteProperty("Enabled", txtXText.Enabled, 0)
    Call PropBag.WriteProperty("Font", txtXText.Font, Ambient.Font)
    Call PropBag.WriteProperty("FontBold", txtXText.FontBold, 0)
    Call PropBag.WriteProperty("FontItalic", txtXText.FontItalic, 0)
    Call PropBag.WriteProperty("FontName", txtXText.FontName, "")
    Call PropBag.WriteProperty("FontSize", txtXText.FontSize, 0)
    Call PropBag.WriteProperty("FontStrikeOut", txtXText.FontStrikethru, 0)
    Call PropBag.WriteProperty("FontUnderline", txtXText.FontUnderline, 0)
    Call PropBag.WriteProperty("Gridlines", lvwCombo.Gridlines, 0)
    Call PropBag.WriteProperty("HideColumnHeaders", lvwCombo.HideColumnHeaders, 0)
    Call PropBag.WriteProperty("MaxLength", txtXText.MaxLength, 0)
    Call PropBag.WriteProperty("MouseIcon", txtXText.MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", txtXText.MousePointer, 0)
    Call PropBag.WriteProperty("PasswordChar", txtXText.PasswordChar, "")
    Call PropBag.WriteProperty("Text", txtXText.Text, UserControl.Name)
'    Call PropBag.WriteProperty("MultiLine", txtXText.MultiLine, 0)
End Sub


