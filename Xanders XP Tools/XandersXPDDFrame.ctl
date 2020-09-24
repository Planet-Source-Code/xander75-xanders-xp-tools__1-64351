VERSION 5.00
Begin VB.UserControl XandersXPDDFrame 
   BackStyle       =   0  'Transparent
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ControlContainer=   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "XandersXPDDFrame.ctx":0000
   Begin VB.PictureBox picDropDown 
      BackColor       =   &H00F1F4F4&
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   0
      ScaleHeight     =   1815
      ScaleWidth      =   2655
      TabIndex        =   0
      Top             =   360
      Width           =   2655
   End
   Begin VB.Image imgDefaultRight 
      Height          =   375
      Left            =   4440
      Picture         =   "XandersXPDDFrame.ctx":0312
      Top             =   360
      Visible         =   0   'False
      Width           =   30
   End
   Begin VB.Image imgDefaultMiddle 
      Height          =   375
      Left            =   3000
      Picture         =   "XandersXPDDFrame.ctx":0652
      Stretch         =   -1  'True
      Top             =   360
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   840
      TabIndex        =   1
      Top             =   120
      Width           =   585
   End
   Begin VB.Image imgGreenRight 
      Height          =   375
      Left            =   4440
      Picture         =   "XandersXPDDFrame.ctx":3A54
      Top             =   1320
      Visible         =   0   'False
      Width           =   30
   End
   Begin VB.Image imgGreenMiddle 
      Height          =   375
      Left            =   3000
      Picture         =   "XandersXPDDFrame.ctx":3D94
      Stretch         =   -1  'True
      Top             =   1320
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Image imgLeftTitle 
      Height          =   375
      Left            =   120
      Picture         =   "XandersXPDDFrame.ctx":6EB6
      Top             =   0
      Width           =   30
   End
   Begin VB.Image imgArrow 
      Height          =   255
      Left            =   2280
      Picture         =   "XandersXPDDFrame.ctx":71F6
      Top             =   0
      Width           =   240
   End
   Begin VB.Image imgMiddleTitle 
      Height          =   375
      Left            =   600
      Picture         =   "XandersXPDDFrame.ctx":7794
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1320
   End
   Begin VB.Image imgRightTitle 
      Height          =   375
      Left            =   2640
      Picture         =   "XandersXPDDFrame.ctx":AB96
      Top             =   0
      Width           =   30
   End
   Begin VB.Image imgSilverRight 
      Height          =   375
      Left            =   4440
      Picture         =   "XandersXPDDFrame.ctx":AED6
      Top             =   1800
      Visible         =   0   'False
      Width           =   30
   End
   Begin VB.Image imgBlueRight 
      Height          =   375
      Left            =   4440
      Picture         =   "XandersXPDDFrame.ctx":B216
      Top             =   840
      Visible         =   0   'False
      Width           =   30
   End
   Begin VB.Image imgSilverMiddle 
      Height          =   345
      Left            =   3000
      Picture         =   "XandersXPDDFrame.ctx":B556
      Stretch         =   -1  'True
      Top             =   1800
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Image imgBlueMiddle 
      Height          =   375
      Left            =   3000
      Picture         =   "XandersXPDDFrame.ctx":E678
      Stretch         =   -1  'True
      Top             =   840
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Shape shpBorder 
      BorderColor     =   &H00FEFDFC&
      Height          =   1095
      Left            =   0
      Top             =   1320
      Width           =   2655
   End
End
Attribute VB_Name = "XandersXPDDFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim PicDDHeight As Long
Dim ControlHeight As Long
Dim m_UserControlHeight As Long
Dim m_PanelColor As OLE_COLOR
Dim m_Expanded As Boolean

Private m_ColorSchemes As XStyles

Private Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function CreateMenu Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Function FindWindow& Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String)
Private Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Function AnyPopup Lib "user32" () As Long

' Events
Event Click()
Event DblClick()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Private Sub Pop1(picDropDown As PictureBox, speed As Integer)
    
    If picDropDown.Visible = True Then Exit Sub

    picDropDown.Height = 0
    picDropDown.Visible = True
    
    'the next line has to be set so the menu(form) knows how far to open too
    Do Until picDropDown.Width >= picDropDown.Width And picDropDown.Height >= PicDDHeight 'width and height must be smaller than
                                                                 'H and W stated in each form
    'the next line tell how fast you want it to open (speed= speed * number, the higher the faster
    'the form opens and what part of menu(form) opens first.

    DoEvents
    picDropDown.Height = picDropDown.Height + speed * 250 ' The higher the number the faster the form will drop
    picDropDown.Top = picDropDown.Top + 0
    picDropDown.Width = picDropDown.Width
    
    Loop
    shpBorder.Visible = True
    UserControl.Height = ControlHeight
    m_UserControlHeight = ControlHeight
    picDropDown.Height = PicDDHeight
End Sub

Private Sub Pop2(picDropDown As PictureBox, speed As Integer)
    
    PicDDHeight = picDropDown.Height
    ControlHeight = UserControl.Height
    m_UserControlHeight = UserControl.Height
    
    'the next line has to be set so the menu(form) knows how far to open too
    Do Until picDropDown.Width <= picDropDown.Width And picDropDown.Height <= 550 'width and height must be smaller than
                                                                 'H and W stated in each form
    'the next line tell how fast you want it to open (speed= speed * number, the higher the faster
    'the form opens and what part of menu(form) opens first.
    
    DoEvents
    picDropDown.Height = picDropDown.Height - speed * 250 ' The higher the number the faster the form will drop
    picDropDown.Top = picDropDown.Top + 0
    picDropDown.Width = picDropDown.Width
    
    Loop
    picDropDown.Visible = False
    shpBorder.Visible = False
    UserControl.Height = imgMiddleTitle.Height
End Sub

Public Property Get UserControlHeight() As Long
    UserControlHeight = m_UserControlHeight
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
'    Call UserControl_Resize
End Property

Public Property Get ColorScheme() As XStyles
    ColorScheme = m_ColorSchemes
End Property

Public Property Let ColorScheme(val As XStyles)
    m_ColorSchemes = val

    If m_ColorSchemes = xpDefault Then
        If m_Expanded = False Then
            imgArrow.Picture = frmImageList.ImageList1.ListImages(10).Picture '   imgBlueDownArrow.Picture
            picDropDown.Visible = False
        ElseIf m_Expanded = True Then
            imgArrow.Picture = frmImageList.ImageList1.ListImages(9).Picture ' imgBlueUpArrow.Picture
            picDropDown.Visible = True
        End If
        
        imgMiddleTitle.Picture = imgDefaultMiddle
        imgRightTitle.Picture = imgDefaultRight
        picDropDown.BackColor = &HF1F4F4
        lblCaption.ForeColor = &H0&
    ElseIf m_ColorSchemes = xpBlue Then
        If m_Expanded = False Then
            imgArrow.Picture = frmImageList.ImageList1.ListImages(12).Picture '   imgBlueDownArrow.Picture
            picDropDown.Visible = False
        ElseIf m_Expanded = True Then
            imgArrow.Picture = frmImageList.ImageList1.ListImages(11).Picture ' imgBlueUpArrow.Picture
            picDropDown.Visible = True
        End If
        
        imgMiddleTitle.Picture = imgBlueMiddle
        imgRightTitle.Picture = imgBlueRight
        picDropDown.BackColor = &HF7DFD6
        lblCaption.ForeColor = &HC65D21
    ElseIf m_ColorSchemes = xpOliveGreen Then
        If m_Expanded = False Then
            imgArrow.Picture = frmImageList.ImageList1.ListImages(14).Picture '   imgBlueDownArrow.Picture
            picDropDown.Visible = False
        ElseIf m_Expanded = True Then
            imgArrow.Picture = frmImageList.ImageList1.ListImages(13).Picture ' imgBlueUpArrow.Picture
            picDropDown.Visible = True
        End If
        
        imgMiddleTitle.Picture = imgGreenMiddle
        imgRightTitle.Picture = imgGreenRight
        picDropDown.BackColor = &HECF6F6
        lblCaption.ForeColor = &H2D6656
    ElseIf m_ColorSchemes = xpSilver Then
        If m_Expanded = False Then
            imgArrow.Picture = frmImageList.ImageList1.ListImages(16).Picture '   imgBlueDownArrow.Picture
            picDropDown.Visible = False
        ElseIf m_Expanded = True Then
            imgArrow.Picture = frmImageList.ImageList1.ListImages(15).Picture ' imgBlueUpArrow.Picture
            picDropDown.Visible = True
        End If
        
        imgMiddleTitle.Picture = imgSilverMiddle
        imgRightTitle.Picture = imgSilverRight
        picDropDown.BackColor = &HF5F1F0
        lblCaption.ForeColor = &H3D3D3F
    End If
End Property

Public Property Get Expanded() As Boolean
    Expanded = m_Expanded
End Property

Public Property Let Expanded(ByVal New_Expanded As Boolean)
    m_Expanded = New_Expanded
    PropertyChanged "Expanded"
        
    Call imgMiddleTitle_MouseDown(1, 0, 100, 100)
End Property

Sub UserChoice()

    If m_ColorSchemes = xpDefault Then
        If m_Expanded = False Then
            imgArrow.Picture = frmImageList.ImageList1.ListImages(10).Picture '   imgBlueDownArrow.Picture
            picDropDown.Visible = False
        ElseIf m_Expanded = True Then
            imgArrow.Picture = frmImageList.ImageList1.ListImages(9).Picture ' imgBlueUpArrow.Picture
            picDropDown.Visible = True
        End If
        
        imgMiddleTitle.Picture = imgDefaultMiddle
        imgRightTitle.Picture = imgDefaultRight
        picDropDown.BackColor = &HF1F4F4
        lblCaption.ForeColor = &H0&
    ElseIf m_ColorSchemes = xpBlue Then
        If m_Expanded = False Then
            imgArrow.Picture = frmImageList.ImageList1.ListImages(12).Picture '   imgBlueDownArrow.Picture
            picDropDown.Visible = False
        ElseIf m_Expanded = True Then
            imgArrow.Picture = frmImageList.ImageList1.ListImages(11).Picture ' imgBlueUpArrow.Picture
            picDropDown.Visible = True
        End If
        
        imgMiddleTitle.Picture = imgBlueMiddle
        imgRightTitle.Picture = imgBlueRight
        picDropDown.BackColor = &HF7DFD6
        lblCaption.ForeColor = &HC65D21
    ElseIf m_ColorSchemes = xpOliveGreen Then
        If m_Expanded = False Then
            imgArrow.Picture = frmImageList.ImageList1.ListImages(14).Picture '   imgBlueDownArrow.Picture
            picDropDown.Visible = False
        ElseIf m_Expanded = True Then
            imgArrow.Picture = frmImageList.ImageList1.ListImages(13).Picture ' imgBlueUpArrow.Picture
            picDropDown.Visible = True
        End If
        
        imgMiddleTitle.Picture = imgGreenMiddle
        imgRightTitle.Picture = imgGreenRight
        picDropDown.BackColor = &HECF6F6
        lblCaption.ForeColor = &H2D6656
    ElseIf m_ColorSchemes = xpSilver Then
        If m_Expanded = False Then
            imgArrow.Picture = frmImageList.ImageList1.ListImages(16).Picture '   imgBlueDownArrow.Picture
            picDropDown.Visible = False
        ElseIf m_Expanded = True Then
            imgArrow.Picture = frmImageList.ImageList1.ListImages(15).Picture ' imgBlueUpArrow.Picture
            picDropDown.Visible = True
        End If
        
        imgMiddleTitle.Picture = imgSilverMiddle
        imgRightTitle.Picture = imgSilverRight
        picDropDown.BackColor = &HF5F1F0
        lblCaption.ForeColor = &H3D3D3F
    End If

End Sub

Private Sub imgArrow_Click()
    Call UserControl_Click
End Sub

Private Sub imgArrow_DblClick()
    Call UserControl_DblClick
End Sub

Private Sub imgArrow_KeyDown(KeyCode As Integer, Shift As Integer)
    Call UserControl_KeyDown(KeyCode, Shift)
End Sub

Private Sub imgArrow_KeyPress(KeyAscii As Integer)
    Call UserControl_KeyPress(KeyAscii)
End Sub

Private Sub imgArrow_KeyUp(KeyCode As Integer, Shift As Integer)
    Call UserControl_KeyUp(KeyCode, Shift)
End Sub

Private Sub imgArrow_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (picDropDown.Visible = False And m_Expanded = True) Or (picDropDown.Visible = False And m_Expanded = False) Then
        If m_ColorSchemes = xpDefault Then
            imgArrow.Picture = frmImageList.ImageList1.ListImages(9).Picture
        ElseIf m_ColorSchemes = xpBlue Then
            imgArrow.Picture = frmImageList.ImageList1.ListImages(11).Picture 'imgBlueUpArrow.Picture
        ElseIf m_ColorSchemes = xpOliveGreen Then
            imgArrow.Picture = frmImageList.ImageList1.ListImages(13).Picture
        ElseIf m_ColorSchemes = xpSilver Then
            imgArrow.Picture = frmImageList.ImageList1.ListImages(15).Picture
        End If
        Call Pop1(picDropDown, 1)
    ElseIf (picDropDown.Visible = True And m_Expanded = False) Or (picDropDown.Visible = True And m_Expanded = True) Then
        If m_ColorSchemes = xpDefault Then
            imgArrow.Picture = frmImageList.ImageList1.ListImages(10).Picture
        ElseIf m_ColorSchemes = xpBlue Then
            imgArrow.Picture = frmImageList.ImageList1.ListImages(12).Picture 'imgBlueDownArrow.Picture
        ElseIf m_ColorSchemes = xpOliveGreen Then
            imgArrow.Picture = frmImageList.ImageList1.ListImages(14).Picture
        ElseIf m_ColorSchemes = xpSilver Then
            imgArrow.Picture = frmImageList.ImageList1.ListImages(16).Picture
        End If
        Call Pop2(picDropDown, 1)
    End If
    Call UserControl_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub imgArrow_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseMove(Button, Shift, X, Y)
End Sub

Private Sub imgArrow_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseUp(Button, Shift, X, Y)
End Sub

Private Sub imgMiddleTitle_Click()
    Call UserControl_Click
End Sub

Private Sub imgMiddleTitle_DblClick()
    Call UserControl_DblClick
End Sub

Private Sub imgMiddleTitle_KeyDown(KeyCode As Integer, Shift As Integer)
    Call UserControl_KeyDown(KeyCode, Shift)
End Sub

Private Sub imgMiddleTitle_KeyPress(KeyAscii As Integer)
    Call UserControl_KeyPress(KeyAscii)
End Sub

Private Sub imgMiddleTitle_KeyUp(KeyCode As Integer, Shift As Integer)
    Call UserControl_KeyUp(KeyCode, Shift)
End Sub

Private Sub imgMiddleTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (picDropDown.Visible = False And m_Expanded = True) Or (picDropDown.Visible = False And m_Expanded = False) Then
        If m_ColorSchemes = xpDefault Then
            imgArrow.Picture = frmImageList.ImageList1.ListImages(9).Picture
        ElseIf m_ColorSchemes = xpBlue Then
            imgArrow.Picture = frmImageList.ImageList1.ListImages(11).Picture 'imgBlueUpArrow.Picture
        ElseIf m_ColorSchemes = xpOliveGreen Then
            imgArrow.Picture = frmImageList.ImageList1.ListImages(13).Picture
        ElseIf m_ColorSchemes = xpSilver Then
            imgArrow.Picture = frmImageList.ImageList1.ListImages(15).Picture
        End If
        Call Pop1(picDropDown, 1)
    ElseIf (picDropDown.Visible = True And m_Expanded = False) Or (picDropDown.Visible = True And m_Expanded = True) Then
        If m_ColorSchemes = xpDefault Then
            imgArrow.Picture = frmImageList.ImageList1.ListImages(10).Picture
        ElseIf m_ColorSchemes = xpBlue Then
            imgArrow.Picture = frmImageList.ImageList1.ListImages(12).Picture 'imgBlueDownArrow.Picture
        ElseIf m_ColorSchemes = xpOliveGreen Then
            imgArrow.Picture = frmImageList.ImageList1.ListImages(14).Picture
        ElseIf m_ColorSchemes = xpSilver Then
            imgArrow.Picture = frmImageList.ImageList1.ListImages(16).Picture
        End If
        Call Pop2(picDropDown, 1)
    End If
    Call UserControl_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub imgMiddleTitle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseMove(Button, Shift, X, Y)
End Sub

Private Sub imgMiddleTitle_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseUp(Button, Shift, X, Y)
End Sub

Private Sub lblCaption_Click()
    Call UserControl_Click
End Sub

Private Sub lblCaption_DblClick()
    Call UserControl_DblClick
End Sub

Private Sub lblCaption_KeyDown(KeyCode As Integer, Shift As Integer)
    Call UserControl_KeyDown(KeyCode, Shift)
End Sub

Private Sub lblCaption_KeyPress(KeyAscii As Integer)
    Call UserControl_KeyPress(KeyAscii)
End Sub

Private Sub lblCaption_KeyUp(KeyCode As Integer, Shift As Integer)
    Call UserControl_KeyUp(KeyCode, Shift)
End Sub

Private Sub lblCaption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (picDropDown.Visible = False And m_Expanded = True) Or (picDropDown.Visible = False And m_Expanded = False) Then
        If m_ColorSchemes = xpDefault Then
            imgArrow.Picture = frmImageList.ImageList1.ListImages(9).Picture
        ElseIf m_ColorSchemes = xpBlue Then
            imgArrow.Picture = frmImageList.ImageList1.ListImages(11).Picture 'imgBlueUpArrow.Picture
        ElseIf m_ColorSchemes = xpOliveGreen Then
            imgArrow.Picture = frmImageList.ImageList1.ListImages(13).Picture
        ElseIf m_ColorSchemes = xpSilver Then
            imgArrow.Picture = frmImageList.ImageList1.ListImages(15).Picture
        End If
        Call Pop1(picDropDown, 1)
    ElseIf (picDropDown.Visible = True And m_Expanded = False) Or (picDropDown.Visible = True And m_Expanded = True) Then
        If m_ColorSchemes = xpDefault Then
            imgArrow.Picture = frmImageList.ImageList1.ListImages(10).Picture
        ElseIf m_ColorSchemes = xpBlue Then
            imgArrow.Picture = frmImageList.ImageList1.ListImages(12).Picture 'imgBlueDownArrow.Picture
        ElseIf m_ColorSchemes = xpOliveGreen Then
            imgArrow.Picture = frmImageList.ImageList1.ListImages(14).Picture
        ElseIf m_ColorSchemes = xpSilver Then
            imgArrow.Picture = frmImageList.ImageList1.ListImages(16).Picture
        End If
        Call Pop2(picDropDown, 1)
    End If
    Call UserControl_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub lblCaption_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseMove(Button, Shift, X, Y)
End Sub

Private Sub lblCaption_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseUp(Button, Shift, X, Y)
End Sub

Private Sub picDropDown_Click()
    Call UserControl_Click
End Sub

Private Sub picDropDown_DblClick()
    Call UserControl_DblClick
End Sub

Private Sub picDropDown_KeyDown(KeyCode As Integer, Shift As Integer)
    Call UserControl_KeyDown(KeyCode, Shift)
End Sub

Private Sub picDropDown_KeyPress(KeyAscii As Integer)
    Call UserControl_KeyPress(KeyAscii)
End Sub

Private Sub picDropDown_KeyUp(KeyCode As Integer, Shift As Integer)
    Call UserControl_KeyUp(KeyCode, Shift)
End Sub

Private Sub picDropDown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub picDropDown_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseMove(Button, Shift, X, Y)
End Sub

Private Sub picDropDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
    UserControl.BackColor = UserControl.Ambient.BackColor
End Sub

Private Sub UserControl_InitProperties()
    m_Expanded = True
    lblCaption.Caption = UserControl.Extender.Name
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
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

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_Resize()
    imgMiddleTitle.Width = UserControl.Width - 60

    imgArrow.Top = 45
    imgArrow.Left = UserControl.Width - 330
    
    imgLeftTitle.Left = 0
    imgLeftTitle.Top = 0
    imgLeftTitle.Height = 375

    imgMiddleTitle.Left = imgLeftTitle.Width
    imgMiddleTitle.Top = 0
    imgMiddleTitle.Height = 375

    imgRightTitle.Left = UserControl.Width - 30
    imgRightTitle.Top = 0
    imgRightTitle.Height = 375
    
    shpBorder.Left = 0
    shpBorder.Width = UserControl.Width
    shpBorder.Top = imgMiddleTitle.Height - 30
    If UserControl.Height <= 375 Then
        shpBorder.Height = 0
    Else
        shpBorder.Height = (UserControl.Height - imgMiddleTitle.Height) + 15
    End If
    
    picDropDown.Left = 15
    picDropDown.Width = UserControl.Width - 30
    picDropDown.Top = imgMiddleTitle.Height
    If UserControl.Height <= 375 Then
        picDropDown.Height = 0
    Else
        picDropDown.Height = (UserControl.Height - imgMiddleTitle.Height) - 30
    End If
    
    lblCaption.Left = 210
    lblCaption.Top = 105
    
    PicDDHeight = picDropDown.Height

End Sub

' Load saved properties.
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
'    picDropDown.BackColor = PropBag.ReadProperty("BackColor", &HC65D21)
    m_UserControlHeight = PropBag.ReadProperty("UserControlHeight", m_def_UserControlHeight)
    lblCaption.Caption = PropBag.ReadProperty("Caption", UserControl.Extender.Name)
    m_ColorSchemes = PropBag.ReadProperty("ColorScheme", m_def_ColorSchemes)
    m_Expanded = PropBag.ReadProperty("Expanded", m_def_Expanded)
    ' picButton.Enabled = PropBag.ReadProperty("Enabled", 0)
    
    ControlHeight = m_UserControlHeight

    If m_ColorSchemes = xpDefault Then
        If m_Expanded = False Then
            imgArrow.Picture = frmImageList.ImageList1.ListImages(10).Picture '   imgBlueDownArrow.Picture
            picDropDown.Visible = False
        ElseIf m_Expanded = True Then
            imgArrow.Picture = frmImageList.ImageList1.ListImages(9).Picture ' imgBlueUpArrow.Picture
            picDropDown.Visible = True
        End If

        imgMiddleTitle.Picture = imgDefaultMiddle
        imgRightTitle.Picture = imgDefaultRight
        picDropDown.BackColor = &HF1F4F4
        lblCaption.ForeColor = &H0&
    ElseIf m_ColorSchemes = xpBlue Then
        If m_Expanded = False Then
            imgArrow.Picture = frmImageList.ImageList1.ListImages(12).Picture '   imgBlueDownArrow.Picture
            picDropDown.Visible = False
        ElseIf m_Expanded = True Then
            imgArrow.Picture = frmImageList.ImageList1.ListImages(11).Picture ' imgBlueUpArrow.Picture
            picDropDown.Visible = True
        End If

        imgMiddleTitle.Picture = imgBlueMiddle
        imgRightTitle.Picture = imgBlueRight
        picDropDown.BackColor = &HF7DFD6
        lblCaption.ForeColor = &HC65D21
    ElseIf m_ColorSchemes = xpOliveGreen Then
        If m_Expanded = False Then
            imgArrow.Picture = frmImageList.ImageList1.ListImages(14).Picture '   imgBlueDownArrow.Picture
            picDropDown.Visible = False
        ElseIf m_Expanded = True Then
            imgArrow.Picture = frmImageList.ImageList1.ListImages(13).Picture ' imgBlueUpArrow.Picture
            picDropDown.Visible = True
        End If

        imgMiddleTitle.Picture = imgGreenMiddle
        imgRightTitle.Picture = imgGreenRight
        picDropDown.BackColor = &HECF6F6
        lblCaption.ForeColor = &H2D6656
    ElseIf m_ColorSchemes = xpSilver Then
        If m_Expanded = False Then
            imgArrow.Picture = frmImageList.ImageList1.ListImages(16).Picture '   imgBlueDownArrow.Picture
            picDropDown.Visible = False
        ElseIf m_Expanded = True Then
            imgArrow.Picture = frmImageList.ImageList1.ListImages(15).Picture ' imgBlueUpArrow.Picture
            picDropDown.Visible = True
        End If

        imgMiddleTitle.Picture = imgSilverMiddle
        imgRightTitle.Picture = imgSilverRight
        picDropDown.BackColor = &HF5F1F0
        lblCaption.ForeColor = &H3D3D3F
    End If
    
End Sub

Private Sub UserControl_Show()
    UserControl.BackColor = UserControl.Ambient.BackColor
End Sub

' Save properties.
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
'    Call PropBag.WriteProperty("BackColor", picDropDown.BackColor, &HC65D21)
    Call PropBag.WriteProperty("UserControlHeight", m_UserControlHeight, m_def_UserControlHeight)
    Call PropBag.WriteProperty("Caption", lblCaption.Caption, UserControl.Extender.Name)
    Call PropBag.WriteProperty("ColorScheme", m_ColorSchemes, m_def_ColorSchemes)
    Call PropBag.WriteProperty("Expanded", m_Expanded, m_def_Expanded)
    'Call PropBag.WriteProperty("Enabled", picButton.Enabled, 0)
End Sub


