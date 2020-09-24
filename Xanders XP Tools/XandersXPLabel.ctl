VERSION 5.00
Begin VB.UserControl XandersXPLabel 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "XandersXPLabel.ctx":0000
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1560
      Top             =   120
   End
   Begin VB.Label lblCaption 
      Caption         =   "XandersXPLabel"
      ForeColor       =   &H00C65D21&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1335
   End
End
Attribute VB_Name = "XandersXPLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim m_AutoSize As Boolean
Dim m_ForeColorDown As OLE_COLOR
Dim m_ForeColorOver As OLE_COLOR
Dim DownForeColor As OLE_COLOR
Dim SelectedForeColor As OLE_COLOR

Event Click()
Event DblClick()
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim highlighted As Boolean

Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Declare Function WindowFromPointXY Lib "User32" Alias "WindowFromPoint" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "User32" (lpPoint As POINTAPI) As Long

Public Sub AboutBox()
Attribute AboutBox.VB_UserMemId = -552
    frmAbout.Show vbModal, Me
End Sub

Private Sub lblCaption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    DownForeColor = lblCaption.ForeColor
    lblCaption.ForeColor = m_ForeColorDown
End Sub

Private Sub lblCaption_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If highlighted Then Exit Sub
    highlighted = True
    SelectedForeColor = lblCaption.ForeColor
    lblCaption.ForeColor = m_ForeColorOver
    Timer1.Enabled = True
End Sub

Private Sub lblCaption_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If highlighted = True Then
        lblCaption.ForeColor = m_ForeColorOver
    ElseIf highlighted = False Then
        lblCaption.ForeColor = SelectedForeColor
    End If
End Sub

Private Sub Timer1_Timer()
    Dim pt As POINTAPI

    ' See where the cursor is.
    GetCursorPos pt
    
    ' Translate into window coordinates.
    If WindowFromPointXY(pt.X, pt.Y) <> UserControl.hWnd _
        Then
        highlighted = False
        lblCaption.ForeColor = SelectedForeColor
        Timer1.Enabled = False
    End If
End Sub

Private Sub UserControl_Initialize()
    m_ForeColorDown = &H8000&
    m_ForeColorOver = &H96E7&
End Sub

Private Sub UserControl_InitProperties()
    lblCaption.Caption = UserControl.Extender.Name
End Sub

Private Sub UserControl_Resize()

'    If lblCaption.Caption = "" Then Exit Sub
    
    ' Do nothing unless AutoSize is True.
    If Not m_AutoSize Then
        lblCaption.Width = UserControl.Width
        lblCaption.Height = UserControl.Height
        Exit Sub
    End If
    
    lblCaption.Width = TextWidth(lblCaption.Caption)
    lblCaption.Height = TextHeight(Left(UCase(lblCaption.Caption), 1))
    
    UserControl.Width = lblCaption.Width
    UserControl.Height = lblCaption.Height
        
End Sub

Public Property Get AutoSize() As Boolean
    AutoSize = m_AutoSize
End Property

Public Property Let AutoSize(ByVal New_AutoSize As Boolean)
    m_AutoSize = New_AutoSize
    PropertyChanged "AutoSize"
    Call UserControl_Resize ' Resize if necessary.
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
Attribute Font.VB_UserMemId = -512
    Set Font = lblCaption.Font
End Property

' Set the font.
Public Property Set Font(ByVal New_Font As Font)
    Set lblCaption.Font = New_Font
    PropertyChanged "Font"
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

Public Property Get ForeColorDown() As OLE_COLOR
    ForeColorDown = m_ForeColorDown
End Property

Public Property Let ForeColorDown(ByVal New_ForeColorDown As OLE_COLOR)
    m_ForeColorDown = New_ForeColorDown
    PropertyChanged "ForeColorDown"
End Property

Public Property Get ForeColorOver() As OLE_COLOR
    ForeColorOver = m_ForeColorOver
End Property

Public Property Let ForeColorOver(ByVal New_ForeColorOver As OLE_COLOR)
    m_ForeColorOver = New_ForeColorOver
    PropertyChanged "ForeColorOver"
End Property

Public Property Get MouseIcon() As Picture
    Set MouseIcon = lblCaption.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set lblCaption.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

Public Property Get MousePointer() As Integer
    MousePointer = lblCaption.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As Integer)
    lblCaption.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
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

' Load saved properties.
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_AutoSize = PropBag.ReadProperty("AutoSize", m_def_AutoSize)
    lblCaption.Enabled = PropBag.ReadProperty("Enabled", 0)
    Set lblCaption.Font = PropBag.ReadProperty("Font", Ambient.Font)
    lblCaption.FontBold = PropBag.ReadProperty("FontBold", 0)
    lblCaption.FontItalic = PropBag.ReadProperty("FontItalic", 0)
    lblCaption.FontName = PropBag.ReadProperty("FontName", "")
    lblCaption.FontSize = PropBag.ReadProperty("FontSize", 0)
    lblCaption.FontStrikethru = PropBag.ReadProperty("FontStrikeOut", 0)
    lblCaption.FontUnderline = PropBag.ReadProperty("FontUnderline", 0)
    lblCaption.Caption = PropBag.ReadProperty("Caption", UserControl.Extender.Name)
    lblCaption.ForeColor = PropBag.ReadProperty("ForeColor", &HC65D21)
    m_ForeColorDown = PropBag.ReadProperty("ForeColorDown", m_def_ForeColorDown)
    m_ForeColorOver = PropBag.ReadProperty("ForeColorOver", m_def_ForeColorOver)
    Set lblCaption.MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    lblCaption.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    
    lblCaption.Caption = UserControl.Extender.Name
End Sub

' Save properties.
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("AutoSize", m_AutoSize, m_def_AutoSize)
    Call PropBag.WriteProperty("Caption", lblCaption.Caption, UserControl.Extender.Name)
    Call PropBag.WriteProperty("Enabled", lblCaption.Enabled, 0)
    Call PropBag.WriteProperty("Font", lblCaption.Font, Ambient.Font)
    Call PropBag.WriteProperty("FontBold", lblCaption.FontBold, 0)
    Call PropBag.WriteProperty("FontItalic", lblCaption.FontItalic, 0)
    Call PropBag.WriteProperty("FontName", lblCaption.FontName, "")
    Call PropBag.WriteProperty("FontSize", lblCaption.FontSize, 0)
    Call PropBag.WriteProperty("FontStrikeOut", lblCaption.FontStrikethru, 0)
    Call PropBag.WriteProperty("FontUnderline", lblCaption.FontUnderline, 0)
    Call PropBag.WriteProperty("ForeColor", lblCaption.ForeColor, &HC65D21)
    Call PropBag.WriteProperty("ForeColorDown", m_ForeColorDown, m_def_ForeColorDown)
    Call PropBag.WriteProperty("ForeColorOver", m_ForeColorOver, m_def_ForeColorOver)
    Call PropBag.WriteProperty("MouseIcon", lblCaption.MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", lblCaption.MousePointer, 0)
End Sub


