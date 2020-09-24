VERSION 5.00
Begin VB.UserControl XandersXPButton 
   ClientHeight    =   1485
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2895
   ScaleHeight     =   1485
   ScaleWidth      =   2895
   ToolboxBitmap   =   "XandersXPButton.ctx":0000
   Begin VB.Timer Hovertimer 
      Enabled         =   0   'False
      Left            =   1920
      Top             =   720
   End
   Begin VB.PictureBox picButton 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   1575
      TabIndex        =   0
      Top             =   0
      Width           =   1575
      Begin VB.Label lblCaption 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "XandersXPButton"
         Height          =   255
         Left            =   0
         TabIndex        =   1
         Top             =   120
         Visible         =   0   'False
         Width           =   1455
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1920
      Top             =   120
   End
End
Attribute VB_Name = "XandersXPButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim highlighted As Boolean
Dim EnabledColour As OLE_COLOR

Public Enum XStyles
    xpDefault = 0
    xpBlue = 1
    xpOliveGreen = 2
    xpSilver = 3
End Enum

Private m_ColorSchemes As XStyles

Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Declare Function WindowFromPointXY Lib "user32" Alias "WindowFromPoint" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

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
    Dim R As Single
    Dim G As Single
    Dim B As Single
    Dim dr As Single
    Dim dg As Single
    Dim db As Single
    Dim Y As Single

    wid = pic.ScaleWidth
    hgt = end_y - start_y
    dr = (end_r - start_r) / hgt
    dg = (end_g - start_g) / hgt
    db = (end_b - start_b) / hgt
    R = start_r
    G = start_g
    B = start_b
    For Y = start_y To end_y
        pic.Line (0, Y)-(wid, Y), RGB(R, G, B)
        R = R + dr
        G = G + dg
        B = B + db
    Next Y
End Sub

Public Sub PrintCaption()
    picButton.CurrentX = (picButton.Width / 2) - ((TextWidth(lblCaption.Caption) / 2) + 20)
    picButton.CurrentY = (picButton.Height / 2) - 80

    If picButton.Enabled = False Then
        EnabledColour = &H90ABAB
        picButton.ForeColor = EnabledColour
        picButton.Print lblCaption.Caption
    ElseIf picButton.Enabled = True Then
        picButton.ForeColor = &H0&
        picButton.Print lblCaption.Caption
    End If
End Sub

Public Sub AboutBox()
    frmAbout.Show vbModal, Me
End Sub

Public Property Get ColorScheme() As XStyles
    ColorScheme = m_ColorSchemes
End Property

Public Property Let ColorScheme(val As XStyles)
    m_ColorSchemes = val

    Call picButton_Paint
    Call PrintCaption
End Property

' Return the caption.
Public Property Get Caption() As String
    Caption = lblCaption.Caption
End Property

' Set the caption.
Public Property Let Caption(ByVal New_TheCaption As String)
    lblCaption.Caption() = New_TheCaption
    PropertyChanged "Caption"
End Property

Public Property Get Enabled() As Boolean
    Enabled = picButton.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    picButton.Enabled() = New_Enabled
    PropertyChanged "Enabled"
    Call UserControl_Paint
    Call picButton_Paint
    Call PrintCaption
End Property

' Return the font.
Public Property Get Font() As Font
    Set Font = lblCaption.Font
End Property

' Set the font.
Public Property Set Font(ByVal New_Font As Font)
    Set lblCaption.Font = New_Font
    PropertyChanged "Font"
End Property

Public Property Get MouseIcon() As Picture
    Set MouseIcon = picButton.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set picButton.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

Public Property Get MousePointer() As Integer
    MousePointer = picButton.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As Integer)
    picButton.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

Private Sub lblCaption_Change()
    Call picButton_Paint
    Call PrintCaption
End Sub

Private Sub lblCaption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call picButton_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub lblCaption_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call picButton_MouseMove(Button, Shift, X, Y)
End Sub

Private Sub lblCaption_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call picButton_MouseUp(Button, Shift, X, Y)
End Sub

Private Sub picButton_Click()
    RaiseEvent Click
End Sub

Private Sub picButton_GotFocus()
    If highlighted Then Exit Sub

    ' Top Line Highlighted
    picButton.Line (0, 15)-(picButton.Width, 15), &HF7D7BD               ' Inside Border
    picButton.Line (15, 0)-(picButton.Width - 15, 0), &HFFE7CE           ' Outside Border

    ' Left Line Highlighted
    picButton.Line (0, 15)-(0, picButton.Height - 30), &HE7AE8C            ' Inside Border
    picButton.Line (15, 15)-(15, picButton.Height - 30), &HE7AE8C          ' Outside Border

    ' Right Line Highlighted
    picButton.Line (picButton.Width - 15, 15)-(picButton.Width - 15, picButton.Height - 30), &HE7AE8C      ' Inside Border
    picButton.Line (picButton.Width - 30, 15)-(picButton.Width - 30, picButton.Height - 30), &HE7AE8C      ' Outside Border

    ' Bottom Line Highlighted
    picButton.Line (0, picButton.Height - 30)-(picButton.Width, picButton.Height - 30), &HEF826B            ' Inside Border
    picButton.Line (15, picButton.Height - 15)-(picButton.Width - 15, picButton.Height - 15), &HEF826B      ' Outside Border

End Sub

Private Sub picButton_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode <> 13 Then Exit Sub
    If m_ColorSchemes = xpDefault Then
        FadeVertical picButton, _
        231, 231, 224, _
        254, 254, 254, _
        0, picButton.ScaleHeight
    ElseIf m_ColorSchemes = xpBlue Then
        FadeVertical picButton, _
        132, 171, 226, _
        254, 254, 254, _
        0, picButton.ScaleHeight
    ElseIf m_ColorSchemes = xpOliveGreen Then
        FadeVertical picButton, _
        172, 186, 135, _
        254, 254, 254, _
        0, picButton.ScaleHeight
    ElseIf m_ColorSchemes = xpSilver Then
        FadeVertical picButton, _
        205, 204, 223, _
        254, 254, 254, _
        0, picButton.ScaleHeight
    End If

    ' Add a Pixel to each corner of the Picture Box
    picButton.Line (0, 0)-(15, 15), &H733C00
    picButton.Line (picButton.Width - 15, 0)-(picButton.Width - 15, 15), &H733C00
    picButton.Line (0, picButton.Height - 15)-(15, picButton.Height - 15), &H733C00
    picButton.Line (picButton.Width - 15, picButton.Height - 15)-(picButton.Width, picButton.Height - 15), &H733C00

    Call PrintCaption
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub picButton_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub picButton_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode <> 13 Then Exit Sub
    Call picButton_Paint

    If highlighted = False Then Exit Sub
    highlighted = True
    Call PrintCaption
    ' Top Line Highlighted
    picButton.Line (0, 15)-(picButton.Width, 15), &H8CDBFF              ' Inside Border
    picButton.Line (30, 30)-(picButton.Width - 30, 30), &HCEF3FF        ' Outside Border

    ' Left Line Highlighted
    picButton.Line (0, 15)-(0, picButton.Height - 30), &H6BCBFF         ' Inside Border
    picButton.Line (15, 15)-(15, picButton.Height - 30), &H6BCBFF       ' Outside Border

    ' Right Line Highlighted
    picButton.Line (picButton.Width - 15, 15)-(picButton.Width - 15, picButton.Height - 30), &H6BCBFF   ' Inside Border
    picButton.Line (picButton.Width - 30, 15)-(picButton.Width - 30, picButton.Height - 30), &H6BCBFF   ' Outside Border

    ' Bottom Line Highlighted
    picButton.Line (0, picButton.Height - 30)-(picButton.Width, picButton.Height - 30), &H96E7&         ' Inside Border
    picButton.Line (15, picButton.Height - 15)-(picButton.Width - 15, picButton.Height - 15), &H96E7&   ' Outside Border

    Timer1.Enabled = True
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub picButton_LostFocus()
    Call picButton_Paint
    highlighted = False
    Call PrintCaption
End Sub

Private Sub picButton_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If m_ColorSchemes = xpDefault Then
        FadeVertical picButton, _
        231, 231, 224, _
        254, 254, 254, _
        0, picButton.ScaleHeight
    ElseIf m_ColorSchemes = xpBlue Then
        FadeVertical picButton, _
        132, 171, 226, _
        254, 254, 254, _
        0, picButton.ScaleHeight
    ElseIf m_ColorSchemes = xpOliveGreen Then
        FadeVertical picButton, _
        172, 186, 135, _
        254, 254, 254, _
        0, picButton.ScaleHeight
    ElseIf m_ColorSchemes = xpSilver Then
        FadeVertical picButton, _
        205, 204, 223, _
        254, 254, 254, _
        0, picButton.ScaleHeight
    End If

    ' Add a Pixel to each corner of the Picture Box
    picButton.Line (0, 0)-(15, 15), &H733C00
    picButton.Line (picButton.Width - 15, 0)-(picButton.Width - 15, 15), &H733C00
    picButton.Line (0, picButton.Height - 15)-(15, picButton.Height - 15), &H733C00
    picButton.Line (picButton.Width - 15, picButton.Height - 15)-(picButton.Width, picButton.Height - 15), &H733C00

    Call PrintCaption

    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub picButton_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If highlighted Then Exit Sub
    highlighted = True
    Call PrintCaption
    ' Top Line Highlighted
    picButton.Line (0, 15)-(picButton.Width, 15), &H8CDBFF              ' Inside Border
    picButton.Line (30, 30)-(picButton.Width - 30, 30), &HCEF3FF        ' Outside Border

    ' Left Line Highlighted
    picButton.Line (0, 15)-(0, picButton.Height - 30), &H6BCBFF         ' Inside Border
    picButton.Line (15, 15)-(15, picButton.Height - 30), &H6BCBFF       ' Outside Border

    ' Right Line Highlighted
    picButton.Line (picButton.Width - 15, 15)-(picButton.Width - 15, picButton.Height - 30), &H6BCBFF   ' Inside Border
    picButton.Line (picButton.Width - 30, 15)-(picButton.Width - 30, picButton.Height - 30), &H6BCBFF   ' Outside Border

    ' Bottom Line Highlighted
    picButton.Line (0, picButton.Height - 30)-(picButton.Width, picButton.Height - 30), &H96E7&         ' Inside Border
    picButton.Line (15, picButton.Height - 15)-(picButton.Width - 15, picButton.Height - 15), &H96E7&   ' Outside Border
    Timer1.Enabled = True

    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub picButton_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call picButton_Paint

    If highlighted = False Then Exit Sub
    highlighted = True
    Call PrintCaption
    ' Top Line Highlighted
    picButton.Line (0, 15)-(picButton.Width, 15), &H8CDBFF              ' Inside Border
    picButton.Line (30, 30)-(picButton.Width - 30, 30), &HCEF3FF        ' Outside Border

    ' Left Line Highlighted
    picButton.Line (0, 15)-(0, picButton.Height - 30), &H6BCBFF         ' Inside Border
    picButton.Line (15, 15)-(15, picButton.Height - 30), &H6BCBFF       ' Outside Border

    ' Right Line Highlighted
    picButton.Line (picButton.Width - 15, 15)-(picButton.Width - 15, picButton.Height - 30), &H6BCBFF   ' Inside Border
    picButton.Line (picButton.Width - 30, 15)-(picButton.Width - 30, picButton.Height - 30), &H6BCBFF   ' Outside Border

    ' Bottom Line Highlighted
    picButton.Line (0, picButton.Height - 30)-(picButton.Width, picButton.Height - 30), &H96E7&         ' Inside Border
    picButton.Line (15, picButton.Height - 15)-(picButton.Width - 15, picButton.Height - 15), &H96E7&   ' Outside Border

    Timer1.Enabled = True
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub picButton_Paint()

    If picButton.Enabled = False Then
        EnabledColour = &HE0E7E7
        picButton.BackColor = EnabledColour
    Else
        If m_ColorSchemes = xpDefault Then
            FadeVertical picButton, _
            254, 254, 254, _
            231, 231, 224, _
            0, picButton.ScaleHeight
        ElseIf m_ColorSchemes = xpBlue Then
            FadeVertical picButton, _
            254, 254, 254, _
            132, 171, 226, _
            0, picButton.ScaleHeight
        ElseIf m_ColorSchemes = xpOliveGreen Then
            FadeVertical picButton, _
            254, 254, 254, _
            172, 186, 135, _
            0, picButton.ScaleHeight
        ElseIf m_ColorSchemes = xpSilver Then
            FadeVertical picButton, _
            254, 254, 254, _
            205, 204, 223, _
            0, picButton.ScaleHeight
        End If
    End If

    If picButton.Enabled = False Then
        EnabledColour = &HE0E7E7
        ' Add a Pixel to each corner of the Picture Box
        picButton.Line (0, 0)-(15, 15), EnabledColour
        picButton.Line (picButton.Width - 15, 0)-(picButton.Width - 15, 15), EnabledColour
        picButton.Line (0, picButton.Height - 15)-(15, picButton.Height - 15), EnabledColour
        picButton.Line (picButton.Width - 15, picButton.Height - 15)-(picButton.Width, picButton.Height - 15), EnabledColour

    Else
        ' Add a Pixel to each corner of the Picture Box
        picButton.Line (0, 0)-(15, 15), &H733C00
        picButton.Line (picButton.Width - 15, 0)-(picButton.Width - 15, 15), &H733C00
        picButton.Line (0, picButton.Height - 15)-(15, picButton.Height - 15), &H733C00
        picButton.Line (picButton.Width - 15, picButton.Height - 15)-(picButton.Width, picButton.Height - 15), &H733C00
    End If

    Call PrintCaption
End Sub

Private Sub Timer1_Timer()
    Dim pt As POINTAPI

    ' See where the cursor is.
    GetCursorPos pt

    ' Translate into window coordinates.
    If WindowFromPointXY(pt.X, pt.Y) <> picButton.hWnd _
        Then
        highlighted = False

        Call picButton_Paint

        Timer1.Enabled = False
    End If
End Sub

Private Sub UserControl_InitProperties()
    highlighted = False
    lblCaption.Caption = UserControl.Extender.Name
End Sub

Private Sub UserControl_Paint()
    ' Draws an XP Style Button
    ' co-ordinates = (x1, y1) are start coordinates and (x2, y2) are end coordinates
    ' (Horizontal, Vertical)
    picButton.Cls
    If picButton.Enabled = False Then
        EnabledColour = &H90ABAB
        UserControl.Line (30, 0)-(UserControl.Width - 30, 0), EnabledColour          ' Top Line
        UserControl.Line (30, UserControl.Height - 15)-(UserControl.Width - 30, UserControl.Height - 15), EnabledColour  ' Bottom Line

        ' Side Lines
        UserControl.Line (0, 30)-(0, UserControl.Height - 30), EnabledColour         ' Left Line
        UserControl.Line (UserControl.Width - 15, 30)-(UserControl.Width - 15, UserControl.Height - 30), EnabledColour   ' Right Line

        ' Corner Pixels
        UserControl.Line (0, 15)-(30, 30), EnabledColour
        UserControl.Line (15, 0)-(30, 30), EnabledColour

        UserControl.Line (UserControl.Width - 30, 0)-(UserControl.Width - 15, 30), EnabledColour
        UserControl.Line (UserControl.Width - 15, 15)-(UserControl.Width - 15, 30), EnabledColour

        UserControl.Line (0, UserControl.Height - 30)-(30, UserControl.Height - 15), EnabledColour
        UserControl.Line (15, UserControl.Height - 15)-(30, UserControl.Height - 15), EnabledColour

        UserControl.Line (UserControl.Width - 30, UserControl.Height - 15)-(UserControl.Width - 15, UserControl.Height - 15), EnabledColour
        UserControl.Line (UserControl.Width - 15, UserControl.Height - 30)-(UserControl.Width - 15, UserControl.Height - 15), EnabledColour
    Else
        ' Top & Bottom Lines
        UserControl.Line (30, 0)-(UserControl.Width - 30, 0), &H733C00          ' Top Line
        UserControl.Line (30, UserControl.Height - 15)-(UserControl.Width - 30, UserControl.Height - 15), &H733C00  ' Bottom Line

        ' Side Lines
        UserControl.Line (0, 30)-(0, UserControl.Height - 30), &H733C00         ' Left Line
        UserControl.Line (UserControl.Width - 15, 30)-(UserControl.Width - 15, UserControl.Height - 30), &H733C00   ' Right Line

        ' Corner Pixels
        UserControl.Line (0, 15)-(30, 30), &HA8957A
        UserControl.Line (15, 0)-(30, 30), &HA8957A

        UserControl.Line (UserControl.Width - 30, 0)-(UserControl.Width - 15, 30), &HA8957A
        UserControl.Line (UserControl.Width - 15, 15)-(UserControl.Width - 15, 30), &HA8957A

        UserControl.Line (0, UserControl.Height - 30)-(30, UserControl.Height - 15), &HA8957A
        UserControl.Line (15, UserControl.Height - 15)-(30, UserControl.Height - 15), &HA8957A

        UserControl.Line (UserControl.Width - 30, UserControl.Height - 15)-(UserControl.Width - 15, UserControl.Height - 15), &HA8957A
        UserControl.Line (UserControl.Width - 15, UserControl.Height - 30)-(UserControl.Width - 15, UserControl.Height - 15), &HA8957A
    End If
    picButton.Refresh
End Sub

Private Sub UserControl_Resize()
    picButton.Left = 15
    picButton.Top = 10
    picButton.Height = UserControl.Height - 30
    picButton.Width = UserControl.Width - 30

    lblCaption.Width = picButton.Width - 30
    lblCaption.Left = picButton.Left + 15
    lblCaption.Top = (picButton.Height / 2) - 80

    Call UserControl_Paint
    Call PrintCaption
End Sub

' Load saved properties.
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_ColorSchemes = PropBag.ReadProperty("ColorScheme", m_def_ColorSchemes)
    lblCaption.Caption = PropBag.ReadProperty("Caption", m_def_Caption)
    picButton.Enabled = PropBag.ReadProperty("Enabled", 0)
    Set lblCaption.Font = PropBag.ReadProperty("Font", Ambient.Font)
    picButton.MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    picButton.MousePointer = PropBag.ReadProperty("MousePointer", 0)

    Call picButton_Paint
    Call PrintCaption
End Sub

' Save properties.
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("ColorScheme", m_ColorSchemes, m_def_ColorSchemes)
    Call PropBag.WriteProperty("Caption", lblCaption.Caption, m_def_Caption)
    Call PropBag.WriteProperty("Enabled", picButton.Enabled, 0)
    Call PropBag.WriteProperty("Font", lblCaption.Font, Ambient.Font)
    Call PropBag.WriteProperty("MouseIcon", picButton.MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", picButton.MousePointer, 0)
End Sub
