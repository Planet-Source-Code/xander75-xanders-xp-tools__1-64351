VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl XandersXPCheckbox 
   BackStyle       =   0  'Transparent
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "XandersXPCheckbox.ctx":0000
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2280
      Top             =   0
   End
   Begin MSComctlLib.ImageList CheckboxImages 
      Left            =   2880
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "XandersXPCheckbox.ctx":0312
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "XandersXPCheckbox.ctx":056C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "XandersXPCheckbox.ctx":07C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "XandersXPCheckbox.ctx":0D24
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image imgCheckbox 
      Height          =   195
      Left            =   120
      Picture         =   "XandersXPCheckbox.ctx":1282
      Top             =   120
      Width           =   195
   End
   Begin VB.Label lblCaption 
      Caption         =   "XandersXPCheckbox"
      ForeColor       =   &H00C65D21&
      Height          =   195
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   1515
   End
   Begin VB.Shape shpCheckbox 
      BackColor       =   &H0022A21F&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0022A21F&
      FillColor       =   &H0022A21F&
      Height          =   180
      Left            =   120
      Top             =   120
      Width           =   195
   End
End
Attribute VB_Name = "XandersXPCheckbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim m_AutoSize As Boolean
Dim highlighted As Boolean
Dim m_CheckColor As OLE_COLOR
Dim m_ForeColorDown As OLE_COLOR
Dim m_ForeColorOver As OLE_COLOR
Dim m_Value As Boolean
Dim Checked As Boolean

Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Declare Function WindowFromPointXY Lib "user32" Alias "WindowFromPoint" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Event Click()
Event DblClick()
Event Initialize()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Public Sub AboutBox()
Attribute AboutBox.VB_UserMemId = -552
    frmAbout.Show vbModal, Me
End Sub

Private Sub imgCheckbox_Click()
    RaiseEvent Click
    Call UserControl_Click
End Sub

Private Sub imgCheckbox_DblClick()
    RaiseEvent DblClick
End Sub
    
Private Sub imgCheckbox_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub imgCheckbox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub imgCheckbox_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub Timer1_Timer()
    Dim pt As POINTAPI

    ' See where the cursor is.
    GetCursorPos pt
    
    ' Translate into window coordinates.
    If WindowFromPointXY(pt.X, pt.Y) <> UserControl.hWnd _
        Then
        highlighted = False
        If imgCheckbox.Picture = CheckboxImages.ListImages(2).Picture Then
            imgCheckbox.Picture = CheckboxImages.ListImages(1).Picture
        ElseIf imgCheckbox.Picture = CheckboxImages.ListImages(4).Picture Then
            imgCheckbox.Picture = CheckboxImages.ListImages(3).Picture
        End If
        Timer1.Enabled = False
    End If
End Sub

Private Sub UserControl_Click()
    If imgCheckbox.Picture = CheckboxImages.ListImages(1).Picture Or imgCheckbox.Picture = CheckboxImages.ListImages(2).Picture Then
        If highlighted = True Then
            imgCheckbox.Picture = CheckboxImages.ListImages(4).Picture
            m_Value = True
            Exit Sub
        Else
            imgCheckbox.Picture = CheckboxImages.ListImages(3).Picture
            m_Value = True
            Exit Sub
        End If
    ElseIf imgCheckbox.Picture = CheckboxImages.ListImages(3).Picture Or imgCheckbox.Picture = CheckboxImages.ListImages(4).Picture Then
        If highlighted = True Then
            imgCheckbox.Picture = CheckboxImages.ListImages(2).Picture
            m_Value = False
            Exit Sub
        Else
            imgCheckbox.Picture = CheckboxImages.ListImages(1).Picture
            m_Value = False
            Exit Sub
        End If
    End If
End Sub

Private Sub UserControl_GotFocus()
    If imgCheckbox.Picture = CheckboxImages.ListImages(1).Picture Then
        imgCheckbox.Picture = CheckboxImages.ListImages(2).Picture
    ElseIf imgCheckbox.Picture = CheckboxImages.ListImages(3).Picture Then
        imgCheckbox.Picture = CheckboxImages.ListImages(4).Picture
    End If
End Sub

Private Sub UserControl_Initialize()
    m_CheckColor = &H22A21F
End Sub

Private Sub UserControl_InitProperties()
    m_Value = m_def_Value
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_LostFocus()
    If imgCheckbox.Picture = CheckboxImages.ListImages(2).Picture Then
        imgCheckbox.Picture = CheckboxImages.ListImages(1).Picture
    ElseIf imgCheckbox.Picture = CheckboxImages.ListImages(4).Picture Then
        imgCheckbox.Picture = CheckboxImages.ListImages(3).Picture
    End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If highlighted Then Exit Sub
    highlighted = True
    If imgCheckbox.Picture = CheckboxImages.ListImages(1).Picture Then
        imgCheckbox.Picture = CheckboxImages.ListImages(2).Picture
    ElseIf imgCheckbox.Picture = CheckboxImages.ListImages(3).Picture Then
        imgCheckbox.Picture = CheckboxImages.ListImages(4).Picture
    End If
    Timer1.Enabled = True
End Sub

Private Sub UserControl_Resize()

    If lblCaption.Caption = "" Then Exit Sub
    
    ' Do nothing unless AutoSize is True.
    If Not m_AutoSize Then

        lblCaption.Width = TextWidth(lblCaption.Caption)
        lblCaption.Height = TextHeight(Left(UCase(lblCaption.Caption), 1))
        
        UserControl.Height = lblCaption.Height + 200
        UserControl.Width = imgCheckbox.Width + 300 + lblCaption.Width
        Exit Sub
    End If
    
    lblCaption.Width = TextWidth(lblCaption.Caption)
    lblCaption.Height = TextHeight(Left(UCase(lblCaption.Caption), 1))
    
    UserControl.Width = imgCheckbox.Width + 300 + lblCaption.Width
    UserControl.Height = lblCaption.Height + 200
    
End Sub

Public Property Get AutoSize() As Boolean
    AutoSize = m_AutoSize
End Property

Public Property Let AutoSize(ByVal New_AutoSize As Boolean)
    m_AutoSize = New_AutoSize
    PropertyChanged "AutoSize"
    Call UserControl_Resize ' Resize if necessary.
End Property

' Return the caption.
Public Property Get Caption() As String
Attribute Caption.VB_UserMemId = -518
    Caption = lblCaption.Caption
End Property

' Set the caption.
Public Property Let Caption(ByVal New_TheCaption As String)
    lblCaption.Caption() = New_TheCaption
    PropertyChanged "Caption"
    Call UserControl_Resize
End Property

Public Property Get Enabled() As Boolean
    Enabled = lblCaption.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    lblCaption.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

' Return the font.
Public Property Get Font() As Font
    Set Font = lblCaption.Font
End Property

' Set the font.
Public Property Set Font(ByVal New_Font As Font)
    Set lblCaption.Font = New_Font
    PropertyChanged "Font"
    Call UserControl_Resize
End Property

Public Property Get FontBold() As Boolean
    FontBold = lblCaption.FontBold
End Property

Public Property Let FontBold(ByVal New_FontBold As Boolean)
    lblCaption.FontBold() = New_FontBold
    PropertyChanged "FontBold"
End Property

Public Property Get FontItalic() As Boolean
    FontItalic = lblCaption.FontItalic
End Property

Public Property Let FontItalic(ByVal New_FontItalic As Boolean)
    lblCaption.FontItalic() = New_FontItalic
    PropertyChanged "FontItalic"
End Property

Public Property Get FontName() As String
    FontName = lblCaption.FontName
End Property

Public Property Let FontName(ByVal New_FontName As String)
    lblCaption.FontName() = New_FontName
    PropertyChanged "FontName"
End Property

Public Property Get FontSize() As Single
    FontSize = lblCaption.FontSize
End Property

Public Property Let FontSize(ByVal New_FontSize As Single)
    lblCaption.FontSize() = New_FontSize
    PropertyChanged "FontSize"
    Call UserControl_Resize
End Property

Public Property Get FontStrikeOut() As Boolean
    FontStrikeOut = lblCaption.FontStrikethru
End Property

Public Property Let FontStrikeOut(ByVal New_FontStrikeOut As Boolean)
    lblCaption.FontStrikethru() = New_FontStrikeOut
    PropertyChanged "FontStrikeOut"
End Property

Public Property Get FontUnderline() As Boolean
    FontUnderline = lblCaption.FontUnderline
End Property

Public Property Let FontUnderline(ByVal New_FontUnderline As Boolean)
    lblCaption.FontUnderline() = New_FontUnderline
    PropertyChanged "FontUnderline"
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = lblCaption.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    lblCaption.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

Public Property Get CheckColor() As OLE_COLOR
    CheckColor = m_CheckColor
End Property

Public Property Let CheckColor(ByVal New_CheckColor As OLE_COLOR)
    m_CheckColor = New_CheckColor
    PropertyChanged "CheckColor"
    shpCheckbox.BackColor = m_CheckColor
    shpCheckbox.BorderColor = m_CheckColor
    shpCheckbox.FillColor = m_CheckColor
End Property

Public Property Get Value() As Boolean
    Value = m_Value
End Property

Public Property Let Value(ByVal New_Value As Boolean)
    m_Value = New_Value
    PropertyChanged "Value"
    If m_Value = False Then
        imgCheckbox.Picture = CheckboxImages.ListImages(1).Picture
    ElseIf m_Value = True Then
        imgCheckbox.Picture = CheckboxImages.ListImages(3).Picture
    End If
End Property

' Load saved properties.
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_AutoSize = PropBag.ReadProperty("AutoSize", m_def_AutoSize)
    m_CheckColor = PropBag.ReadProperty("CheckColor", m_def_CheckColor)
    lblCaption.Caption = PropBag.ReadProperty("Caption", "XandersXPCheckbox")
    lblCaption.Enabled = PropBag.ReadProperty("Enabled", 0)
    Set lblCaption.Font = PropBag.ReadProperty("Font", Ambient.Font)
    lblCaption.FontBold = PropBag.ReadProperty("FontBold", 0)
    lblCaption.FontItalic = PropBag.ReadProperty("FontItalic", 0)
    lblCaption.FontName = PropBag.ReadProperty("FontName", "")
    lblCaption.FontSize = PropBag.ReadProperty("FontSize", 0)
    lblCaption.FontStrikethru = PropBag.ReadProperty("FontStrikeOut", 0)
    lblCaption.FontUnderline = PropBag.ReadProperty("FontUnderline", 0)
    lblCaption.ForeColor = PropBag.ReadProperty("ForeColor", &HC65D21)
    m_Value = PropBag.ReadProperty("Value", m_def_Value)

    shpCheckbox.BackColor = m_CheckColor
    shpCheckbox.BorderColor = m_CheckColor
    shpCheckbox.FillColor = m_CheckColor

    If m_Value = False Then
        imgCheckbox.Picture = CheckboxImages.ListImages(1).Picture
    ElseIf m_Value = True Then
        imgCheckbox.Picture = CheckboxImages.ListImages(3).Picture
    End If
End Sub

' Save properties.
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("AutoSize", m_AutoSize, m_def_AutoSize)
    Call PropBag.WriteProperty("CheckColor", m_CheckColor, m_def_CheckColor)
    Call PropBag.WriteProperty("Caption", lblCaption.Caption, "XandersXPCheckbox")
    Call PropBag.WriteProperty("Enabled", lblCaption.Enabled, 0)
    Call PropBag.WriteProperty("Font", lblCaption.Font, Ambient.Font)
    Call PropBag.WriteProperty("FontBold", lblCaption.FontBold, 0)
    Call PropBag.WriteProperty("FontItalic", lblCaption.FontItalic, 0)
    Call PropBag.WriteProperty("FontName", lblCaption.FontName, "")
    Call PropBag.WriteProperty("FontSize", lblCaption.FontSize, 0)
    Call PropBag.WriteProperty("FontStrikeOut", lblCaption.FontStrikethru, 0)
    Call PropBag.WriteProperty("FontUnderline", lblCaption.FontUnderline, 0)
    Call PropBag.WriteProperty("ForeColor", lblCaption.ForeColor, &HC65D21)
    Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
End Sub






