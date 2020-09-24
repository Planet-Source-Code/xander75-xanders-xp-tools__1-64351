VERSION 5.00
Begin VB.UserControl XandersXPPanel 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3270
   ControlContainer=   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   3270
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2760
      Top             =   1440
   End
   Begin VB.PictureBox picPanel 
      BorderStyle     =   0  'None
      Height          =   3615
      Left            =   0
      ScaleHeight     =   3615
      ScaleWidth      =   2400
      TabIndex        =   0
      Top             =   0
      Width           =   2400
   End
End
Attribute VB_Name = "XandersXPPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim EnabledColour As OLE_COLOR

Private m_ColorSchemes As XStyles

Public Enum XGradient
    xpHorizontal = 0
    xpVertical = 1
End Enum

Private m_Gradient As XGradient

Event Click()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

' Fade colors in a vertical area of a PictureBox.
Private Sub FadeVertical(ByVal pic As PictureBox, ByVal start_r As Single, ByVal start_g As Single, ByVal start_b As Single, ByVal end_r As Single, ByVal end_g As Single, ByVal end_b As Single, ByVal start_y, ByVal end_y)
    Dim hgt As Single
    Dim wid As Single
    Dim r As Single
    Dim g As Single
    Dim b As Single
    Dim dr As Single
    Dim dg As Single
    Dim db As Single
    Dim Y As Single

    wid = pic.ScaleWidth
    hgt = end_y - start_y
    dr = (end_r - start_r) / hgt
    dg = (end_g - start_g) / hgt
    db = (end_b - start_b) / hgt
    r = start_r
    g = start_g
    b = start_b
    For Y = start_y To end_y
        pic.Line (0, Y)-(wid, Y), RGB(r, g, b)
        r = r + dr
        g = g + dg
        b = b + db
    Next Y
End Sub

Private Sub FadeHorizontal(ByVal pic As PictureBox, ByVal start_r As Single, ByVal start_g As Single, ByVal start_b As Single, ByVal end_r As Single, ByVal end_g As Single, ByVal end_b As Single, ByVal start_y, ByVal end_y)
    Dim hgt As Single
    Dim hei As Single
    Dim r As Single
    Dim g As Single
    Dim b As Single
    Dim dr As Single
    Dim dg As Single
    Dim db As Single
    Dim Y As Single

    hgt = pic.ScaleHeight
    wid = end_y - start_y
    dr = (end_r - start_r) / wid
    dg = (end_g - start_g) / wid
    db = (end_b - start_b) / wid
    r = start_r
    g = start_g
    b = start_b
    For Y = start_y To end_y
        pic.Line (Y, 0)-(Y, hgt), RGB(r, g, b)
        r = r + dr
        g = g + dg
        b = b + db
    Next Y
End Sub

Public Property Get ColorScheme() As XStyles
    ColorScheme = m_ColorSchemes
End Property

Public Property Let ColorScheme(val As XStyles)
    m_ColorSchemes = val
    Call picPanel_Paint
End Property

Public Property Get Enabled() As Boolean
    Enabled = picPanel.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    picPanel.Enabled() = New_Enabled
    PropertyChanged "Enabled"
    Call picPanel_Paint
End Property

Public Property Get GradientStyle() As XGradient
    GradientStyle = m_Gradient
End Property

Public Property Let GradientStyle(val As XGradient)
    m_Gradient = val
    picPanel_Paint
End Property

Public Property Get MouseIcon() As Picture
    Set MouseIcon = picPanel.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set picPanel.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

Public Property Get MousePointer() As Integer
    MousePointer = picPanel.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As Integer)
    picPanel.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

Private Sub picPanel_Paint()
    If picPanel.Enabled = False Then
        EnabledColour = &HE0E7E7
        picPanel.BackColor = EnabledColour
    Else
        If m_Gradient = xpHorizontal Then
            If m_ColorSchemes = xpDefault Then
                FadeHorizontal picPanel, _
                254, 254, 254, _
                231, 231, 224, _
                0, picPanel.ScaleWidth
            ElseIf m_ColorSchemes = xpBlue Then
                FadeHorizontal picPanel, _
                139, 169, 229, _
                117, 134, 220, _
                0, picPanel.ScaleWidth
            ElseIf m_ColorSchemes = xpOliveGreen Then
                FadeHorizontal picPanel, _
                254, 254, 254, _
                172, 186, 135, _
                0, picPanel.ScaleWidth
            ElseIf m_ColorSchemes = xpSilver Then
                FadeHorizontal picPanel, _
                254, 254, 254, _
                205, 204, 223, _
                0, picPanel.ScaleWidth
            End If
        ElseIf m_Gradient = xpVertical Then
            If m_ColorSchemes = xpDefault Then
                FadeVertical picPanel, _
                254, 254, 254, _
                231, 231, 224, _
                0, picPanel.ScaleHeight
            ElseIf m_ColorSchemes = xpBlue Then
                FadeVertical picPanel, _
                139, 169, 229, _
                117, 134, 220, _
                0, picPanel.ScaleHeight
            ElseIf m_ColorSchemes = xpOliveGreen Then
                FadeVertical picPanel, _
                254, 254, 254, _
                172, 186, 135, _
                0, picPanel.ScaleHeight
            ElseIf m_ColorSchemes = xpSilver Then
                FadeVertical picPanel, _
                254, 254, 254, _
                205, 204, 223, _
                0, picPanel.ScaleHeight
            End If
        End If
    End If
    
End Sub

Private Sub UserControl_InitProperties()
    UserControl.Extender.Left = picPanel.Left
    UserControl.Extender.Top = picPanel.Top
    picPanel.Height = Parent.Height
    picPanel.Width = UserControl.Width
    m_Gradient = xpVertical
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_ColorSchemes = PropBag.ReadProperty("ColorScheme", m_def_ColorSchemes)
    picPanel.Enabled = PropBag.ReadProperty("Enabled", 0)
    m_Gradient = PropBag.ReadProperty("GradientStyle", m_def_Gradient)
    picPanel.MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    picPanel.MousePointer = PropBag.ReadProperty("MousePointer", 0)
End Sub

Private Sub UserControl_Resize()
    UserControl.Extender.Left = picPanel.Left
    UserControl.Extender.Top = picPanel.Top
    picPanel.Height = Parent.Height
    picPanel.Width = UserControl.Width
End Sub

' Save properties.
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("ColorScheme", m_ColorSchemes, m_def_ColorSchemes)
    Call PropBag.WriteProperty("Enabled", picPanel.Enabled, 0)
    Call PropBag.WriteProperty("GradientStyle", m_Gradient, m_def_Gradient)
    Call PropBag.WriteProperty("MouseIcon", picPanel.MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", picPanel.MousePointer, 0)
End Sub
