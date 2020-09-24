VERSION 5.00
Begin VB.UserControl XandersXPSpin 
   ClientHeight    =   1110
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2595
   ScaleHeight     =   1110
   ScaleWidth      =   2595
   ToolboxBitmap   =   "XandersXPSpin.ctx":0000
   Begin VB.Timer tmrDown 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2040
      Top             =   480
   End
   Begin VB.Timer tmrUp 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2040
      Top             =   0
   End
   Begin VB.TextBox txtXText 
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   15
      TabIndex        =   0
      Top             =   15
      Width           =   1080
   End
   Begin VB.Image imgDownArrow 
      Height          =   120
      Left            =   1440
      Picture         =   "XandersXPSpin.ctx":0312
      Top             =   600
      Width           =   255
   End
   Begin VB.Image imgUpArrow 
      Height          =   120
      Left            =   1440
      Picture         =   "XandersXPSpin.ctx":0661
      Top             =   480
      Width           =   255
   End
   Begin VB.Image imgDown 
      Height          =   150
      Left            =   1080
      Picture         =   "XandersXPSpin.ctx":09AE
      Stretch         =   -1  'True
      Top             =   600
      Width           =   255
   End
   Begin VB.Image imgUp 
      Height          =   135
      Left            =   1080
      Picture         =   "XandersXPSpin.ctx":0BF8
      Stretch         =   -1  'True
      Top             =   480
      Width           =   255
   End
   Begin VB.Shape shBorder 
      BorderColor     =   &H00B99D7F&
      Height          =   330
      Left            =   0
      Top             =   0
      Width           =   1365
   End
End
Attribute VB_Name = "XandersXPSpin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim m_AutoSelect As Boolean
Dim m_BorderColorOver As OLE_COLOR
Dim SelectedBorderColor As OLE_COLOR

' Set the Margins within the Textbox
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const EM_SETMARGINS = &HD3
Private Const EC_LEFTMARGIN = &H1
Private Const EC_RIGHTMARGIN = &H2

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

Private Sub txtXText_GotFocus()
    SelectedBorderColor = shBorder.BorderColor
    shBorder.BorderColor = m_BorderColorOver
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

Private Sub txtXText_LostFocus()
    shBorder.BorderColor = SelectedBorderColor
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

Private Sub imgDown_Click()
    txtXText.Text = val(txtXText.Text) - 1
End Sub

Private Sub imgDown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    tmrDown.Enabled = True
End Sub

Private Sub imgDownArrow_Click()
    txtXText.Text = val(txtXText.Text) - 1
End Sub

Private Sub imgDownArrow_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    tmrDown.Enabled = True
End Sub

Private Sub imgDownArrow_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    tmrDown.Enabled = False
End Sub

Private Sub imgUp_Click()
    txtXText.Text = val(txtXText.Text) + 1
End Sub

Private Sub imgUp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    tmrUp.Enabled = True
End Sub

Private Sub imgUpArrow_Click()
    txtXText.Text = val(txtXText.Text) + 1
End Sub

Private Sub imgUpArrow_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    tmrUp.Enabled = True
End Sub

Private Sub imgUpArrow_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    tmrUp.Enabled = False
End Sub

Private Sub tmrDown_Timer()
    txtXText.Text = val(txtXText.Text) - 1
End Sub

Private Sub tmrUp_Timer()
    txtXText.Text = val(txtXText.Text) + 1
End Sub

Private Sub UserControl_Initialize()
    m_BorderColorOver = &H96E7&
    
    Dim left_margin As Integer
    Dim right_margin As Integer
    Dim long_value As Long

    left_margin = CInt(2)
    right_margin = CInt(2)
    long_value = right_margin * &H10000 + left_margin

    SendMessage txtXText.hwnd, _
        EM_SETMARGINS, _
        EC_LEFTMARGIN Or EC_RIGHTMARGIN, _
        long_value
    
    If txtXText.Text = "" Then txtXText.Text = 0
End Sub

Private Sub UserControl_InitProperties()
    If txtXText.Text = "" Then txtXText.Text = 0
End Sub

Private Sub UserControl_Resize()

    shBorder.Height = UserControl.Height
    shBorder.Width = UserControl.Width
    
    txtXText.Height = shBorder.Height - 25
    txtXText.Left = shBorder.Left + 10
    txtXText.Top = shBorder.Top + 15
    txtXText.Width = shBorder.Width - imgUp.Width - 30
    
    imgUp.Top = shBorder.Top + 15
    imgUp.Left = txtXText.Width + 15
    imgUp.Height = (shBorder.Height / 2) - 10
    
    imgUpArrow.Left = imgUp.Left
    imgUpArrow.Top = (shBorder.Height / 2) / 2 - 30
    
    imgDown.Top = imgUp.Height
    imgDown.Left = imgUp.Left
    imgDown.Height = (shBorder.Height / 2) - 10

    imgDownArrow.Left = imgUp.Left
    imgDownArrow.Top = (shBorder.Height / 2) + (shBorder.Height / 2) / 2 - 65

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
    txtXText.MaxLength = PropBag.ReadProperty("MaxLength", 0)
    Set txtXText.MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    txtXText.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    txtXText.PasswordChar = PropBag.ReadProperty("PasswordChar", "")
    Set txtXText.Font = PropBag.ReadProperty("Font", Ambient.Font)
    txtXText.Text = PropBag.ReadProperty("Text", UserControl.Extender.Name)
'    txtXText.MultiLine = PropBag.ReadProperty("MultiLine", 0)

    If m_AutoSelect = True Then
        txtXText.SelStart = 0
        txtXText.SelLength = Len(txtXText.Text)
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
    Call PropBag.WriteProperty("MaxLength", txtXText.MaxLength, 0)
    Call PropBag.WriteProperty("MouseIcon", txtXText.MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", txtXText.MousePointer, 0)
    Call PropBag.WriteProperty("PasswordChar", txtXText.PasswordChar, "")
    Call PropBag.WriteProperty("Text", txtXText.Text, UserControl.Extender.Name)
'    Call PropBag.WriteProperty("MultiLine", txtXText.MultiLine, 0)
End Sub
