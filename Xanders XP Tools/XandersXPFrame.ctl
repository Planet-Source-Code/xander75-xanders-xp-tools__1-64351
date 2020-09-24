VERSION 5.00
Begin VB.UserControl XandersXPFrame 
   BackStyle       =   0  'Transparent
   ClientHeight    =   1965
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3300
   ControlContainer=   -1  'True
   MaskColor       =   &H00000000&
   MaskPicture     =   "XandersXPFrame.ctx":0000
   ScaleHeight     =   1965
   ScaleWidth      =   3300
   ToolboxBitmap   =   "XandersXPFrame.ctx":2D8A12
   Begin VB.PictureBox picFrame 
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   0
      ScaleHeight     =   1455
      ScaleWidth      =   3015
      TabIndex        =   0
      Top             =   0
      Width           =   3015
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         Caption         =   "XandersXPFrame"
         ForeColor       =   &H00C65D21&
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   0
         Width           =   1230
      End
   End
End
Attribute VB_Name = "XandersXPFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Event Click()
Event DblClick()
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

Private Sub picFrame_Click()
    RaiseEvent Click
End Sub

Private Sub picFrame_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub picFrame_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub picFrame_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub picFrame_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub picFrame_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub picFrame_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub picFrame_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub picFrame_Paint()

    picFrame.Cls
    ' Top & Bottom Lines
    picFrame.Line (30, 90)-(picFrame.Width - 30, 90), &HBFD0D0  ' Top Line
    picFrame.Line (30, picFrame.Height - 15)-(picFrame.Width - 30, picFrame.Height - 15), &HBFD0D0     ' Bottom Line

    ' Side Lines
    picFrame.Line (0, 120)-(0, picFrame.Height - 30), &HBFD0D0 ' Left Line
    picFrame.Line (picFrame.Width - 15, 120)-(picFrame.Width - 15, picFrame.Height - 30), &HBFD0D0      ' Right Line

    ' Add the curved Pixels to each corner of the Picture Box
    picFrame.Line (15, 105)-(30, 90), &HBFD0D0    ' Top Left Corner
    picFrame.Line (30, 105)-(30, 90), &HBFD0D0    ' Top Left Corner
    picFrame.Line (15, 120)-(30, 90), &HBFD0D0    ' Top Left Corner
    
    picFrame.Line (picFrame.Width - 30, 120)-(picFrame.Width - 45, 105), &HBFD0D0 ' Top Right Corner
    picFrame.Line (picFrame.Width - 30, 105)-(picFrame.Width - 30, 90), &HBFD0D0 ' Top Right Corner
    picFrame.Line (picFrame.Width - 45, 105)-(picFrame.Width - 30, 90), &HBFD0D0 ' Top Right Corner
    
    picFrame.Line (15, picFrame.Height - 30)-(30, picFrame.Height - 30), &HBFD0D0 ' Bottom Left Corner
    picFrame.Line (30, picFrame.Height - 30)-(45, picFrame.Height - 45), &HBFD0D0 ' Bottom Left Corner
    picFrame.Line (15, picFrame.Height - 45)-(30, picFrame.Height - 30), &HBFD0D0 ' Bottom Left Corner
    
    picFrame.Line (picFrame.Width - 30, picFrame.Height - 30)-(picFrame.Width - 45, picFrame.Height - 45), &HBFD0D0 ' Bottom Right Corner
    picFrame.Line (picFrame.Width - 45, picFrame.Height - 30)-(picFrame.Width - 45, picFrame.Height - 45), &HBFD0D0 ' Bottom Right Corner
    picFrame.Line (picFrame.Width - 30, picFrame.Height - 45)-(picFrame.Width - 45, picFrame.Height - 45), &HBFD0D0 ' Bottom Right Corner
End Sub

Public Property Get BackColor() As OLE_COLOR
    BackColor = picFrame.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    picFrame.BackColor() = New_BackColor
    lblCaption.BackColor = New_BackColor
    PropertyChanged "BackColor"
    Call picFrame_Paint
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

' Set the ForeColor
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = lblCaption.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    lblCaption.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property
'
'' Set the BorderColor
'Public Property Get BorderColor() As OLE_COLOR
'    BorderColor = shFrame.BorderColor
'End Property
'
'Public Property Let BorderColor(ByVal New_BorderColor As OLE_COLOR)
'    shFrame.BorderColor() = New_BorderColor
'    PropertyChanged "BorderColor"
'End Property

Private Sub UserControl_InitProperties()
    lblCaption.Caption = UserControl.Extender.Name
End Sub

' Load saved properties.
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    picFrame.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    Set lblCaption.Font = PropBag.ReadProperty("Font", Ambient.Font)
    lblCaption.Caption = PropBag.ReadProperty("Caption", m_def_Caption)
'    shFrame.BorderColor = PropBag.ReadProperty("BorderColor", &HBFD0D0)
    lblCaption.ForeColor = PropBag.ReadProperty("ForeColor", &HC65D21)
    
    lblCaption.BackColor = picFrame.BackColor
End Sub

Private Sub UserControl_Resize()
    picFrame.Height = UserControl.Height
    picFrame.Width = UserControl.Width
End Sub

' Save properties.
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", picFrame.BackColor, &H80000005)
    Call PropBag.WriteProperty("Font", lblCaption.Font, Ambient.Font)
    Call PropBag.WriteProperty("Caption", lblCaption.Caption, m_def_Caption)
'    Call PropBag.WriteProperty("BorderColor", shFrame.BorderColor, &HBFD0D0)
    Call PropBag.WriteProperty("ForeColor", lblCaption.ForeColor, &HC65D21)
End Sub

