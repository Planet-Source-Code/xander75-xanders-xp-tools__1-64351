VERSION 5.00
Begin VB.UserControl XandersXPText 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "XandersXPText.ctx":0000
   Begin VB.TextBox MyTxt 
      Height          =   495
      Left            =   720
      TabIndex        =   1
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox txtXText 
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   15
      TabIndex        =   0
      Text            =   "XandersXPText"
      Top             =   15
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Shape shBorder 
      BorderColor     =   &H00B99D7F&
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1365
   End
End
Attribute VB_Name = "XandersXPText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Dim m_AutoSelect As Boolean
''Dim m_ComputerInfo As Boolean
'Dim m_BorderColorOver As OLE_COLOR
'Dim SelectedBorderColor As OLE_COLOR
'
'Public Enum XCompInfo
'    None = 0
'    Username = 1
'    Computername = 2
'End Enum
'
'Private m_ComputerInfo As XCompInfo
'
'' Set the Margins within the Textbox
'Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'Private Const EM_SETMARGINS = &HD3
'Private Const EC_LEFTMARGIN = &H1
'Private Const EC_RIGHTMARGIN = &H2
'
'Public Enum XTextStyle
'    Normal = 0
'    lowercase = 1
'    UPPERCASE = 2
'End Enum
'
'' Events
'Event Change()
'Event Click()
'Event DblClick()
'Event KeyDown(KeyCode As Integer, Shift As Integer)
'Event KeyPress(KeyAscii As Integer)
'Event KeyUp(KeyCode As Integer, Shift As Integer)
'Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'
'Public Sub AboutBox()
'    frmAbout.Show vbModal, Me
'End Sub
'
'Private Sub txtXText_GotFocus()
'    SelectedBorderColor = shBorder.BorderColor
'    shBorder.BorderColor = m_BorderColorOver
'End Sub
'
'Private Sub txtXText_LostFocus()
'    shBorder.BorderColor = SelectedBorderColor
'End Sub
'
'Private Sub UserControl_Initialize()
'    m_BorderColorOver = &H96E7&
'
'    Dim left_margin As Integer
'    Dim right_margin As Integer
'    Dim long_value As Long
'
'    left_margin = CInt(4)
'    right_margin = CInt(4)
'    long_value = right_margin * &H10000 + left_margin
'
'    SendMessage txtXText.hWnd, _
'        EM_SETMARGINS, _
'        EC_LEFTMARGIN Or EC_RIGHTMARGIN, _
'        long_value
'End Sub
'
'Private Sub UserControl_InitProperties()
'    txtXText.Text = UserControl.Extender.Name
'End Sub
'
'Private Sub UserControl_Resize()
'    shBorder.Height = UserControl.Height
'    shBorder.Width = UserControl.Width
'
'    txtXText.Height = shBorder.Height - 25
'    txtXText.Left = shBorder.Left + 10
'    txtXText.Top = shBorder.Top + 15
'    txtXText.Width = shBorder.Width - 40
'End Sub
'
'Public Property Get Alignment() As AlignmentConstants
'    Alignment = txtXText.Alignment
'End Property
'
'Public Property Let Alignment(ByVal New_Alignment As AlignmentConstants)
'    txtXText.Alignment() = New_Alignment
'    PropertyChanged "Alignment"
'End Property
'
'Public Property Get AutoSelect() As Boolean
'    AutoSelect = m_AutoSelect
'End Property
'
'Public Property Let AutoSelect(ByVal New_AutoSelect As Boolean)
'    m_AutoSelect = New_AutoSelect
'    PropertyChanged "AutoSelect"
'End Property
'
'Public Property Get BackColor() As OLE_COLOR
'    BackColor = txtXText.BackColor
'End Property
'
'Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
'    txtXText.BackColor() = New_BackColor
'    PropertyChanged "BackColor"
'End Property
'
'Public Property Get BorderColor() As OLE_COLOR
'    BorderColor = shBorder.BorderColor
'End Property
'
'Public Property Let BorderColor(ByVal New_BorderColor As OLE_COLOR)
'    shBorder.BorderColor() = New_BorderColor
'    PropertyChanged "BorderColor"
'End Property
'
'Public Property Get BorderColorOver() As OLE_COLOR
'    BorderColorOver = m_BorderColorOver
'End Property
'
'Public Property Let BorderColorOver(ByVal New_BorderColorOver As OLE_COLOR)
'    m_BorderColorOver = New_BorderColorOver
'    PropertyChanged "BorderColorOver"
'End Property
'
'Public Property Get ComputerInfo() As XCompInfo
'    ComputerInfo = m_ComputerInfo
'End Property
'
'Public Property Let ComputerInfo(val As XCompInfo)
'    m_ComputerInfo = val
'
'    If m_ComputerInfo = None Then
'        txtXText.Text = UserControl.Extender.Name
'    ElseIf m_ComputerInfo = Computername Then
'        txtXText.Text = Environ("ComputerName")
'    ElseIf m_ComputerInfo = Username Then
'        txtXText.Text = Environ("UserName")
'    End If
'End Property
'
'Public Property Get Enabled() As Boolean
'    Enabled = txtXText.Enabled
'End Property
'
'Public Property Let Enabled(ByVal New_Enabled As Boolean)
'    txtXText.Enabled() = New_Enabled
'    PropertyChanged "Enabled"
'End Property
'
'' Return the font.
'Public Property Get Font() As Font
'    Set Font = txtXText.Font
'End Property
'
'' Set the font.
'Public Property Set Font(ByVal New_Font As Font)
'    Set txtXText.Font = New_Font
'    PropertyChanged "Font"
'End Property
'
'Public Property Get FontBold() As Boolean
'    FontBold = txtXText.FontBold
'End Property
'
'Public Property Let FontBold(ByVal New_FontBold As Boolean)
'    txtXText.FontBold() = New_FontBold
'    PropertyChanged "FontBold"
'End Property
'
'Public Property Get FontItalic() As Boolean
'    FontItalic = txtXText.FontItalic
'End Property
'
'Public Property Let FontItalic(ByVal New_FontItalic As Boolean)
'    txtXText.FontItalic() = New_FontItalic
'    PropertyChanged "FontItalic"
'End Property
'
'Public Property Get FontName() As String
'    FontName = txtXText.FontName
'End Property
'
'Public Property Let FontName(ByVal New_FontName As String)
'    txtXText.FontName() = New_FontName
'    PropertyChanged "FontName"
'End Property
'
'Public Property Get FontSize() As Single
'    FontSize = txtXText.FontSize
'End Property
'
'Public Property Let FontSize(ByVal New_FontSize As Single)
'    txtXText.FontSize() = New_FontSize
'    PropertyChanged "FontSize"
'End Property
'
'Public Property Get FontStrikeOut() As Boolean
'    FontStrikeOut = txtXText.FontStrikethru
'End Property
'
'Public Property Let FontStrikeOut(ByVal New_FontStrikeOut As Boolean)
'    txtXText.FontStrikethru() = New_FontStrikeOut
'    PropertyChanged "FontStrikeOut"
'End Property
'
'Public Property Get FontUnderline() As Boolean
'    FontUnderline = txtXText.FontUnderline
'End Property
'
'Public Property Let FontUnderline(ByVal New_FontUnderline As Boolean)
'    txtXText.FontUnderline() = New_FontUnderline
'    PropertyChanged "FontUnderline"
'End Property
'
'Public Property Get ForeColor() As OLE_COLOR
'    ForeColor = txtXText.ForeColor
'End Property
'
'Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
'    txtXText.ForeColor() = New_ForeColor
'    PropertyChanged "ForeColor"
'End Property
'
'Public Property Get MaxLength() As Long
'    MaxLength = txtXText.MaxLength
'End Property
'
'Public Property Let MaxLength(ByVal New_MaxLength As Long)
'    txtXText.MaxLength() = New_MaxLength
'    PropertyChanged "MaxLength"
'End Property
'
'Public Property Get MouseIcon() As Picture
'    Set MouseIcon = txtXText.MouseIcon
'End Property
'
'Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
'    Set txtXText.MouseIcon = New_MouseIcon
'    PropertyChanged "MouseIcon"
'End Property
'
'Public Property Get MousePointer() As Integer
'    MousePointer = txtXText.MousePointer
'End Property
'
'Public Property Let MousePointer(ByVal New_MousePointer As Integer)
'    txtXText.MousePointer() = New_MousePointer
'    PropertyChanged "MousePointer"
'End Property
'
'Public Property Get PasswordChar() As String
'    PasswordChar = txtXText.PasswordChar
'End Property
'
'Public Property Let PasswordChar(ByVal New_PasswordChar As String)
'    txtXText.PasswordChar() = New_PasswordChar
'    PropertyChanged "PasswordChar"
'End Property
'
'' Return the Text.
'Public Property Get Text() As String
'    Text = txtXText.Text
'End Property
'
'' Set the Text.
'Public Property Let Text(ByVal New_Text As String)
'    txtXText.Text() = New_Text
'    PropertyChanged "Text"
'End Property
'
'Public Property Get TextStyle() As XTextStyle
'    TextStyle = m_TextStyle
'End Property
'
'Public Property Let TextStyle(val As XTextStyle)
'    m_TextStyle = val
'End Property
'
''Public Property Get MultiLine() As Boolean
''    MultiLine = txtXText.MultiLine
''End Property
''
''Public Property Let MultiLine(ByVal New_MultiLine As Boolean)
''    txtXText.MultiLine() = New_MultiLine
''    PropertyChanged "MultiLine"
''End Property
'
'Private Sub txtXText_Change()
'    RaiseEvent Change
'End Sub
'
'Private Sub txtXText_Click()
'    RaiseEvent Click
'End Sub
'
'Private Sub txtXText_DblClick()
'    RaiseEvent DblClick
'End Sub
'
'Private Sub txtXText_KeyDown(KeyCode As Integer, Shift As Integer)
'    RaiseEvent KeyDown(KeyCode, Shift)
'End Sub
'
'Private Sub txtXText_KeyPress(KeyAscii As Integer)
'
'    If m_TextStyle = Default Then
'    ElseIf m_TextStyle = lowercase Then
'        If KeyAscii >= 65 And KeyAscii <= 90 Then
'        KeyAscii = KeyAscii + 32
'    End If
'    ElseIf m_TextStyle = UPPERCASE Then
'        If KeyAscii >= 97 And KeyAscii <= 122 Then
'            KeyAscii = KeyAscii - 32
'        End If
'    End If
'
'    RaiseEvent KeyPress(KeyAscii)
'End Sub
'
'Private Sub txtXText_KeyUp(KeyCode As Integer, Shift As Integer)
'    RaiseEvent KeyUp(KeyCode, Shift)
'End Sub
'
'Private Sub txtXText_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    RaiseEvent MouseDown(Button, Shift, X, Y)
'End Sub
'
'Private Sub txtXText_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    RaiseEvent MouseMove(Button, Shift, X, Y)
'End Sub
'
'Private Sub txtXText_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    RaiseEvent MouseUp(Button, Shift, X, Y)
'End Sub
'
'' Load saved properties.
'Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
'    txtXText.Alignment = PropBag.ReadProperty("Alignment", txtXText.Alignment)
'    m_AutoSelect = PropBag.ReadProperty("AutoSelect", m_def_AutoSelect)
'    txtXText.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
'    txtXText.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
'    shBorder.BorderColor = PropBag.ReadProperty("BorderColor", &HB99D7F)
'    m_BorderColorOver = PropBag.ReadProperty("BorderColorOver", m_def_BorderColorOver)
'    txtXText.Enabled = PropBag.ReadProperty("Enabled", 0)
'    Set txtXText.Font = PropBag.ReadProperty("Font", Ambient.Font)
'    txtXText.FontBold = PropBag.ReadProperty("FontBold", 0)
'    txtXText.FontItalic = PropBag.ReadProperty("FontItalic", 0)
'    txtXText.FontName = PropBag.ReadProperty("FontName", "")
'    txtXText.FontSize = PropBag.ReadProperty("FontSize", 0)
'    txtXText.FontStrikethru = PropBag.ReadProperty("FontStrikeOut", 0)
'    txtXText.FontUnderline = PropBag.ReadProperty("FontUnderline", 0)
'    txtXText.MaxLength = PropBag.ReadProperty("MaxLength", 0)
'    Set txtXText.MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
'    txtXText.MousePointer = PropBag.ReadProperty("MousePointer", 0)
'    txtXText.PasswordChar = PropBag.ReadProperty("PasswordChar", "")
'    txtXText.Text = PropBag.ReadProperty("Text", 0)
'    m_TextStyle = PropBag.ReadProperty("TextStyle", m_def_TextStyle)
'    m_ComputerInfo = PropBag.ReadProperty("ComputerInfo", m_def_ComputerInfo)
''    txtXText.MultiLine = PropBag.ReadProperty("MultiLine", 0)
'
'    If m_AutoSelect = True Then
'        txtXText.SelStart = 0
'        txtXText.SelLength = Len(txtXText.Text)
'    End If
'
'    If m_ComputerInfo = None Then
'        txtXText.Text = txtXText.Text
'    ElseIf m_ComputerInfo = Computername Then
'        txtXText.Text = Environ("ComputerName")
'    ElseIf m_ComputerInfo = Username Then
'        txtXText.Text = Environ("UserName")
'    End If
'
'End Sub
'
'' Save properties.
'Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
'    Call PropBag.WriteProperty("Alignment", txtXText.Alignment, "Alignment")
'    Call PropBag.WriteProperty("AutoSelect", m_AutoSelect, m_def_AutoSelect)
'    Call PropBag.WriteProperty("BackColor", txtXText.BackColor, &H80000005)
'    Call PropBag.WriteProperty("ForeColor", txtXText.ForeColor, &H80000008)
'    Call PropBag.WriteProperty("BorderColor", shBorder.BorderColor, &HB99D7F)
'    Call PropBag.WriteProperty("BorderColorOver", m_BorderColorOver, m_def_BorderColorOver)
'    Call PropBag.WriteProperty("Enabled", txtXText.Enabled, 0)
'    Call PropBag.WriteProperty("Font", txtXText.Font, Ambient.Font)
'    Call PropBag.WriteProperty("FontBold", txtXText.FontBold, 0)
'    Call PropBag.WriteProperty("FontItalic", txtXText.FontItalic, 0)
'    Call PropBag.WriteProperty("FontName", txtXText.FontName, "")
'    Call PropBag.WriteProperty("FontSize", txtXText.FontSize, 0)
'    Call PropBag.WriteProperty("FontStrikeOut", txtXText.FontStrikethru, 0)
'    Call PropBag.WriteProperty("FontUnderline", txtXText.FontUnderline, 0)
'    Call PropBag.WriteProperty("MaxLength", txtXText.MaxLength, 0)
'    Call PropBag.WriteProperty("MouseIcon", txtXText.MouseIcon, Nothing)
'    Call PropBag.WriteProperty("MousePointer", txtXText.MousePointer, 0)
'    Call PropBag.WriteProperty("PasswordChar", txtXText.PasswordChar, "")
'    Call PropBag.WriteProperty("Text", txtXText.Text, 0)
'    Call PropBag.WriteProperty("TextStyle", m_TextStyle, m_def_TextStyle)
'    Call PropBag.WriteProperty("ComputerInfo", m_ComputerInfo, m_def_ComputerInfo)
''    Call PropBag.WriteProperty("MultiLine", txtXText.MultiLine, 0)
'End Sub
'
'
'
Public Enum states
    Normal = 0
    Disable = 1
    ReadOnly = 2
End Enum
Const m_def_BorderColor = &HB99D7F
Const m_def_BorderColorOver = &H96E7&
Const m_def_DataFields = ""
Dim m_AutoSelect As Boolean
Dim m_BorderColor As OLE_COLOR
Dim m_BorderColorOver As OLE_COLOR
Dim m_DataFields As String
Event Change()
Event Click()
Event DblClick()
Event KeyPress(KeyAscii As Integer)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=MyTxt,MyTxt,-1,MouseMove
Sub RePos()
On Error Resume Next
    With UserControl
        MyTxt.Width = .Width - 120
        MyTxt.Height = .Height - 120
        MyTxt.Left = 60
        MyTxt.Top = 60
    End With
End Sub
Private Sub MyTxt_GotFocus()
    SetMyFocus m_BorderColorOver
    If m_AutoSelect = True Then
        MyTxt.SelStart = 0
        MyTxt.SelLength = Len(MyTxt.Text)
    End If
End Sub
Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
    MyTxt.SetFocus
End Sub

Private Sub UserControl_ExitFocus()
    SetMyFocus m_BorderColor
End Sub
Private Sub UserControl_Resize()
    RePos
    Call MyXPtxt(MyTxt, vbWhite, Normal)
End Sub

Private Function MyXPtxt(Txt As TextBox, BackColor As ColorConstants, State As states)
    UserControl.Cls
    UserControl.BackColor = BackColor
    UserControl.ScaleMode = 1
    Txt.Appearance = 0
    Txt.BorderStyle = 0
    UserControl.AutoRedraw = True
    UserControl.DrawWidth = 1
    UserControl.Line (0, 0)-(UserControl.Width, 0), m_BorderColor
    UserControl.Line (0, 0)-(0, UserControl.Height), m_BorderColor
    UserControl.Line (UserControl.Width - 15, 0)-(UserControl.Width - 15, UserControl.Height), m_BorderColor
    UserControl.Line (0, UserControl.Height - 15)-(UserControl.Width, UserControl.Height - 15), m_BorderColor
    
    If State = Normal Then
        Txt.BackColor = vbWhite
        Txt.Enabled = True
        Txt.Locked = False
    ElseIf State = Disable Then
        Txt.Enabled = False
        Txt.BackColor = RGB(235, 235, 228)
        Txt.ForeColor = RGB(161, 161, 146)
    ElseIf State = ReadOnly Then
        Txt.Enabled = True
        Txt.Locked = True
    End If
    
End Function
Public Property Get Alignment() As Integer
    Alignment = MyTxt.Alignment
End Property
Public Property Let Alignment(ByVal New_Alignment As Integer)
    If New_Alignment > 2 Then New_Alignment = 0
    MyTxt.Alignment() = New_Alignment
    PropertyChanged "Alignment"
End Property

Public Property Get AutoSelect() As Boolean
    AutoSelect = m_AutoSelect
End Property

Public Property Let AutoSelect(ByVal New_AutoSelect As Boolean)
    m_AutoSelect = New_AutoSelect
    PropertyChanged "AutoSelect"
End Property

Private Sub MyTxt_Change()
    RaiseEvent Change
End Sub
Private Sub MyTxt_Click()
    RaiseEvent Click
End Sub
Private Sub MyTxt_DblClick()
    RaiseEvent DblClick
End Sub
Public Property Get Enabled() As Boolean
    Enabled = MyTxt.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    MyTxt.Enabled() = New_Enabled
    PropertyChanged "Enabled"
    If New_Enabled Then
        SetMyFocus RGB(127, 157, 185)
    Else
        SetMyFocus RGB(191, 167, 128)
    End If
End Property
Public Property Get Font() As Font
    Set Font = MyTxt.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set MyTxt.Font = New_Font
    PropertyChanged "Font"
End Property
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = MyTxt.ForeColor
End Property
Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    MyTxt.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property
Private Sub MyTxt_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub
Public Property Get Locked() As Boolean
    Locked = MyTxt.Locked
End Property
Public Property Let Locked(ByVal New_Locked As Boolean)
    MyTxt.Locked() = New_Locked
    PropertyChanged "Locked"
End Property
Public Property Get MaxLength() As Long
    MaxLength = MyTxt.MaxLength
End Property
Public Property Let MaxLength(ByVal New_MaxLength As Long)
    MyTxt.MaxLength() = New_MaxLength
    PropertyChanged "MaxLength"
End Property
Private Sub MyTxt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub
Public Property Get PasswordChar() As String
    PasswordChar = MyTxt.PasswordChar
End Property
Public Property Let PasswordChar(ByVal New_PasswordChar As String)
    MyTxt.PasswordChar() = New_PasswordChar
    PropertyChanged "PasswordChar"
End Property
Public Property Get SelStart() As Long
    SelStart = MyTxt.SelStart
End Property
Public Property Let SelStart(ByVal New_SelStart As Long)
    MyTxt.SelStart() = New_SelStart
    PropertyChanged "SelStart"
End Property
Public Property Get SelText() As String
    SelText = MyTxt.SelText
End Property
Public Property Let SelText(ByVal New_SelText As String)
    MyTxt.SelText() = New_SelText
    PropertyChanged "SelText"
End Property
Public Property Get SelLength() As Long
    SelLength = MyTxt.SelLength
End Property
Public Property Let SelLength(ByVal New_SelLength As Long)
    MyTxt.SelLength() = New_SelLength
    PropertyChanged "SelLength"
End Property
Public Property Get Text() As String
    Text = MyTxt.Text
End Property

Public Property Let Text(ByVal New_Text As String)
    MyTxt.Text() = New_Text
    PropertyChanged "Text"
End Property
Public Property Get ToolTipText() As String
    ToolTipText = MyTxt.ToolTipText
End Property

Public Property Let ToolTipText(ByVal New_ToolTipText As String)
    MyTxt.ToolTipText() = New_ToolTipText
    PropertyChanged "ToolTipText"
End Property
Private Sub UserControl_InitProperties()
    m_DataFields = m_def_DataFields
    MyTxt.Text = UserControl.Extender.Name
    UserControl.Height = 330
    MyTxt.FontName = "Verdana"
    UserControl_Resize
    m_BorderColor = m_def_BorderColor
    m_BorderColorOver = m_def_BorderColorOver
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    MyTxt.Alignment = PropBag.ReadProperty("Alignment", 0)
    m_AutoSelect = PropBag.ReadProperty("AutoSelect", m_def_AutoSelect)
    MyTxt.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    MyTxt.Enabled = PropBag.ReadProperty("Enabled", True)
    Set MyTxt.Font = PropBag.ReadProperty("Font", Ambient.Font)
    MyTxt.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    MyTxt.Locked = PropBag.ReadProperty("Locked", False)
    MyTxt.MaxLength = PropBag.ReadProperty("MaxLength", 0)
    MyTxt.PasswordChar = PropBag.ReadProperty("PasswordChar", "")
    MyTxt.SelStart = PropBag.ReadProperty("SelStart", 0)
    MyTxt.SelText = PropBag.ReadProperty("SelText", "")
    MyTxt.SelLength = PropBag.ReadProperty("SelLength", 0)
    MyTxt.Text = PropBag.ReadProperty("Text", "Text1")
    MyTxt.ToolTipText = PropBag.ReadProperty("ToolTipText", "")
    m_BorderColor = PropBag.ReadProperty("BorderColor", m_def_BorderColor)
    m_BorderColorOver = PropBag.ReadProperty("BorderColorOver", m_def_BorderColorOver)
    
    If m_AutoSelect = True Then
        MyTxt.SelStart = 0
        MyTxt.SelLength = Len(MyTxt.Text)
    End If
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Alignment", MyTxt.Alignment, 0)
    Call PropBag.WriteProperty("AutoSelect", m_AutoSelect, m_def_AutoSelect)
    Call PropBag.WriteProperty("BackColor", MyTxt.BackColor, &H80000005)
    Call PropBag.WriteProperty("Enabled", MyTxt.Enabled, True)
    Call PropBag.WriteProperty("Font", MyTxt.Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", MyTxt.ForeColor, &H80000008)
    Call PropBag.WriteProperty("Locked", MyTxt.Locked, False)
    Call PropBag.WriteProperty("MaxLength", MyTxt.MaxLength, 0)
    Call PropBag.WriteProperty("PasswordChar", MyTxt.PasswordChar, "")
    Call PropBag.WriteProperty("SelStart", MyTxt.SelStart, 0)
    Call PropBag.WriteProperty("SelText", MyTxt.SelText, "")
    Call PropBag.WriteProperty("SelLength", MyTxt.SelLength, 0)
    Call PropBag.WriteProperty("Text", MyTxt.Text, "Text1")
    Call PropBag.WriteProperty("ToolTipText", MyTxt.ToolTipText, "")
    Call PropBag.WriteProperty("Value", val(MyTxt.Text), 0)
    Call PropBag.WriteProperty("BorderColor", m_BorderColor, m_def_BorderColor)
    Call PropBag.WriteProperty("BorderColorOver", m_BorderColorOver, m_def_BorderColorOver)
End Sub
Private Sub SetMyFocus(LineColor As ColorConstants)
    UserControl.AutoRedraw = True
    UserControl.DrawWidth = 1
    UserControl.Line (0, 0)-(UserControl.Width, 0), LineColor
    UserControl.Line (0, 0)-(0, UserControl.Height), LineColor
    UserControl.Line (UserControl.Width - 15, 0)-(UserControl.Width - 15, UserControl.Height), LineColor
    UserControl.Line (0, UserControl.Height - 15)-(UserControl.Width, UserControl.Height - 15), LineColor
End Sub
Public Property Get Value() As Double
    Value = val(MyTxt.Text)
End Property
Public Property Let Value(ByVal New_Value As Double)
    MyTxt.Text() = New_Value
    PropertyChanged "Value"
End Property
Public Property Get BorderColor() As OLE_COLOR
    BorderColor = m_BorderColor
End Property
Public Property Let BorderColor(ByVal New_BorderColor As OLE_COLOR)
    m_BorderColor = New_BorderColor
    MyXPtxt MyTxt, vbWhite, Normal
    PropertyChanged "BorderColor"
End Property
Public Property Get BorderColorFocus() As OLE_COLOR
    BorderColorFocus = m_BorderColorOver
End Property
Public Property Let BorderColorFocus(ByVal New_BorderColorOver As OLE_COLOR)
    m_BorderColorOver = New_BorderColorOver
    PropertyChanged "BorderColorOver"
End Property


