VERSION 5.00
Begin VB.UserControl XandersXPTransparency 
   ClientHeight    =   2385
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4005
   ScaleHeight     =   2385
   ScaleWidth      =   4005
   ToolboxBitmap   =   "XandersXPTransparency.ctx":0000
   Begin VB.Timer tmrDown 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1800
      Top             =   480
   End
   Begin VB.Timer tmrUp 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1320
      Top             =   480
   End
   Begin VB.Image Image1 
      Height          =   15
      Left            =   45
      Picture         =   "XandersXPTransparency.ctx":0312
      Top             =   450
      Width           =   960
   End
   Begin VB.Image imgBlank 
      Height          =   465
      Left            =   3720
      Picture         =   "XandersXPTransparency.ctx":0414
      Top             =   0
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image imgUp 
      Height          =   150
      Left            =   770
      Picture         =   "XandersXPTransparency.ctx":09AA
      Top             =   90
      Width           =   195
   End
   Begin VB.Image imgDown 
      Height          =   150
      Left            =   770
      Picture         =   "XandersXPTransparency.ctx":0EBC
      Top             =   270
      Width           =   195
   End
   Begin VB.Image imgNum3 
      Height          =   465
      Left            =   525
      Picture         =   "XandersXPTransparency.ctx":13CE
      Top             =   0
      Width           =   210
   End
   Begin VB.Image imgNum2 
      Height          =   465
      Left            =   315
      Picture         =   "XandersXPTransparency.ctx":1964
      Top             =   0
      Width           =   210
   End
   Begin VB.Image imgNum1 
      Height          =   465
      Left            =   105
      Picture         =   "XandersXPTransparency.ctx":1EFA
      Top             =   0
      Width           =   210
   End
   Begin VB.Image imgRightBackground 
      Height          =   465
      Left            =   735
      Picture         =   "XandersXPTransparency.ctx":2490
      Top             =   0
      Width           =   300
   End
   Begin VB.Image img9 
      Height          =   465
      Left            =   3480
      Picture         =   "XandersXPTransparency.ctx":28AB
      Top             =   0
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image img8 
      Height          =   465
      Left            =   3240
      Picture         =   "XandersXPTransparency.ctx":2E41
      Top             =   0
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image img7 
      Height          =   465
      Left            =   3000
      Picture         =   "XandersXPTransparency.ctx":33D7
      Top             =   0
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image img6 
      Height          =   465
      Left            =   2760
      Picture         =   "XandersXPTransparency.ctx":396D
      Top             =   0
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image img5 
      Height          =   465
      Left            =   2520
      Picture         =   "XandersXPTransparency.ctx":3F03
      Top             =   0
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image img4 
      Height          =   465
      Left            =   2280
      Picture         =   "XandersXPTransparency.ctx":4499
      Top             =   0
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image img3 
      Height          =   465
      Left            =   2040
      Picture         =   "XandersXPTransparency.ctx":4A2F
      Top             =   0
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image img2 
      Height          =   465
      Left            =   1800
      Picture         =   "XandersXPTransparency.ctx":4FC5
      Top             =   0
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image img1 
      Height          =   465
      Left            =   1560
      Picture         =   "XandersXPTransparency.ctx":555B
      Top             =   0
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image img0 
      Height          =   465
      Left            =   1320
      Picture         =   "XandersXPTransparency.ctx":5AF1
      Top             =   0
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image imgLeftBackground 
      Height          =   465
      Left            =   0
      Picture         =   "XandersXPTransparency.ctx":6087
      Top             =   0
      Width           =   105
   End
End
Attribute VB_Name = "XandersXPTransparency"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Private m_TransparencyLevel As Long
Private moForm As Form

' Make a Semi Transparent Form
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function UpdateLayeredWindow Lib "user32" (ByVal hWnd As Long, ByVal hdcDst As Long, pptDst As Any, psize As Any, ByVal hdcSrc As Long, pptSrc As Any, crKey As Long, ByVal pblend As Long, ByVal dwFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const G = (-20)
Private Const LWA_COLORKEY = &H1
Private Const LWA_ALPHA = &H2
Private Const ULW_COLORKEY = &H1
Private Const ULW_ALPHA = &H2
Private Const ULW_OPAQUE = &H4
Private Const WS_EX_LAYERED = &H80000

Private Sub GetNumbers()
    Select Case Len(Trim(m_TransparencyLevel))
        Case 1
            imgNum1.Picture = imgBlank.Picture
            imgNum2.Picture = imgBlank.Picture
            Call FillNum3(val(m_TransparencyLevel))
        Case 2
            imgNum1.Picture = imgBlank.Picture
            Call FillNum2(Trim(Left(m_TransparencyLevel, 1)))
            Call FillNum3(Trim(Mid(m_TransparencyLevel, 2, 1)))
        Case 3
            Call FillNum1(Trim(Left(m_TransparencyLevel, 1)))
            Call FillNum2(Trim(Mid(m_TransparencyLevel, 2, 1)))
            Call FillNum3(Trim(Mid(m_TransparencyLevel, 3, 1)))
    End Select
End Sub

Private Sub FillNum1(FillNumber1 As Integer)

    Select Case FillNumber1
        Case 0
            imgNum1.Picture = img0.Picture
        Case 1
            imgNum1.Picture = img1.Picture
        Case 2
            imgNum1.Picture = img2.Picture
        Case 3
            imgNum1.Picture = img3.Picture
        Case 4
            imgNum1.Picture = img4.Picture
        Case 5
            imgNum1.Picture = img5.Picture
        Case 6
            imgNum1.Picture = img6.Picture
        Case 7
            imgNum1.Picture = img7.Picture
        Case 8
            imgNum1.Picture = img8.Picture
        Case 9
            imgNum1.Picture = img9.Picture
    End Select

End Sub

Private Sub FillNum2(FillNumber2 As Integer)

    Select Case FillNumber2
        Case 0
            imgNum2.Picture = img0.Picture
        Case 1
            imgNum2.Picture = img1.Picture
        Case 2
            imgNum2.Picture = img2.Picture
        Case 3
            imgNum2.Picture = img3.Picture
        Case 4
            imgNum2.Picture = img4.Picture
        Case 5
            imgNum2.Picture = img5.Picture
        Case 6
            imgNum2.Picture = img6.Picture
        Case 7
            imgNum2.Picture = img7.Picture
        Case 8
            imgNum2.Picture = img8.Picture
        Case 9
            imgNum2.Picture = img9.Picture
    End Select

End Sub

Private Sub FillNum3(FillNumber3 As Integer)

    Select Case FillNumber3
        Case 0
            imgNum3.Picture = img0.Picture
        Case 1
            imgNum3.Picture = img1.Picture
        Case 2
            imgNum3.Picture = img2.Picture
        Case 3
            imgNum3.Picture = img3.Picture
        Case 4
            imgNum3.Picture = img4.Picture
        Case 5
            imgNum3.Picture = img5.Picture
        Case 6
            imgNum3.Picture = img6.Picture
        Case 7
            imgNum3.Picture = img7.Picture
        Case 8
            imgNum3.Picture = img8.Picture
        Case 9
            imgNum3.Picture = img9.Picture
    End Select

End Sub

Public Function MakeSemiTransparent(ByVal hWnd As Long, ByVal Perc As Integer) As Long
    Dim Msg As Long
    On Error Resume Next
     
    Perc = ((100 - Perc) / 100) * 255
    If Perc < 0 Or Perc > 255 Then
        MakeSemiTransparent = 1
    Else
        Msg = GetWindowLong(hWnd, G)
        Msg = Msg Or WS_EX_LAYERED
        SetWindowLong hWnd, G, Msg
        SetLayeredWindowAttributes hWnd, 0, Perc, LWA_ALPHA
        MakeSemiTransparent = 0
    End If
    If Err Then
        MakeSemiTransparent = 2
    End If
End Function

Private Sub imgDown_Click()
    If val(m_TransparencyLevel) > 0 Then
        m_TransparencyLevel = m_TransparencyLevel - 1
        Call GetNumbers
        MakeSemiTransparent UserControl.Parent.hWnd, m_TransparencyLevel
    End If
End Sub

Private Sub imgDown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    tmrDown.Enabled = True
End Sub

Private Sub imgDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    tmrDown.Enabled = False
End Sub

Private Sub imgUp_Click()
    If val(m_TransparencyLevel) < 100 Then
        m_TransparencyLevel = m_TransparencyLevel + 1
        Call GetNumbers
        MakeSemiTransparent UserControl.Parent.hWnd, m_TransparencyLevel
    End If
End Sub

Public Property Get TransparencyLevel() As Long
    TransparencyLevel = m_TransparencyLevel
End Property

Public Property Let TransparencyLevel(ByVal New_TransparencyLevel As Long)
    m_TransparencyLevel = New_TransparencyLevel
    PropertyChanged "TransparencyLevel"
    Call GetNumbers
End Property

Private Sub imgUp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    tmrUp.Enabled = True
End Sub

Private Sub imgUp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    tmrUp.Enabled = False
End Sub

Private Sub tmrDown_Timer()
    If val(m_TransparencyLevel) > 0 Then
        m_TransparencyLevel = m_TransparencyLevel - 1
        Call GetNumbers
        MakeSemiTransparent UserControl.Parent.hWnd, m_TransparencyLevel
    End If
End Sub

Private Sub tmrUp_Timer()
    If val(m_TransparencyLevel) < 100 Then
        m_TransparencyLevel = m_TransparencyLevel + 1
        Call GetNumbers
        MakeSemiTransparent UserControl.Parent.hWnd, m_TransparencyLevel
    End If
End Sub

Private Sub UserControl_Initialize()
    m_TransparencyLevel = 0
End Sub

Private Sub UserControl_InitProperties()
    UserControl.Height = 465
    UserControl.Width = imgRightBackground.Left + imgRightBackground.Width
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_TransparencyLevel = PropBag.ReadProperty("TransparencyLevel", m_def_TransparencyLevel)
    Call GetNumbers
    MakeSemiTransparent UserControl.Parent.hWnd, m_TransparencyLevel
End Sub

Private Sub UserControl_Resize()
    UserControl.Height = 465
    UserControl.Width = imgRightBackground.Left + imgRightBackground.Width
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("TransparencyLevel", m_TransparencyLevel, m_def_TransparencyLevel)
End Sub
