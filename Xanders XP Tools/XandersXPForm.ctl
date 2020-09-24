VERSION 5.00
Begin VB.UserControl XandersXPForm 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
End
Attribute VB_Name = "XandersXPForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvPara As Any, ByVal fuWinIni As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long

Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2
Dim bTransparent As Boolean
Private WithEvents MyForm As Form
Attribute MyForm.VB_VarHelpID = -1
Public Event Closed()
 Const MF_BYPOSITION = &H400&
 Const MF_BYCOMMAND = 0
 Const SC_RESTORE = &HF120
 Const SC_MOVE = &HF010
 Const SC_SIZE = &HF000
 Const SC_MINIMIZE = &HF020
 Const SC_MAXIMIZE = &HF030
 Const SC_CLOSE = &HF060
 Const WM_GETSYSMENU = &H313


Const GWL_STYLE = (-16)
Const WS_SYSMENU = &H80000

Private Sub RePos()
    'This repositions the different controls on the form when it is resized
    Dim X As Single
    Dim Y As Single
    
    If UserControl.Height < 615 Then UserControl.Height = 615   'Checks that form
    If UserControl.Width < 1695 Then UserControl.Width = 1695   'is not too small
    
    X = UserControl.Width / Screen.TwipsPerPixelX   'Registers the size of the
    Y = UserControl.Height / Screen.TwipsPerPixelY  'form in pixels
    
    'Titlebar
    With TitleLeft
        .Left = 0
        .Top = 0
    End With
    
    With Title
        .Left = TitleLeft.Width
        .Top = 0
        .Width = X - TitleLeft.Width - TitleRight.Width
    End With
    
    With TitleRight
        .Left = Title.Left + Title.Width
        .Top = 0
    End With
    
    'Borders
    With BottomLeft
        .Left = 0
        .Top = Y - .Height
    End With
    
    With BottomRight
        .Left = X - .Width
        .Top = Y - .Height
    End With
    
    With Left
        .Left = 0
        .Top = TitleLeft.Top + TitleLeft.Height
        .Height = BottomLeft.Top - .Top
    End With
    
    With Right
        .Left = X - .Width
        .Top = TitleRight.Top + TitleRight.Height
        .Height = BottomRight.Top - .Top
    End With
    
    With Bottom
        .Left = BottomLeft.Width
        .Top = Y - Bottom.Height
        .Width = X - BottomLeft.Width - BottomRight.Width
    End With
    
    'Buttons
    With CloseButton
        .Left = Right.Left - .Width - 2
        .Top = (Title.Height - .Height) / 2
    End With
    
    With MaximizeButton
        .Left = CloseButton.Left - .Width - 2
        .Top = (Title.Height - .Height) / 2
    End With
    
    With MinimizeButton
        .Left = MaximizeButton.Left - .Width - 2
        .Top = (Title.Height - .Height) / 2
    End With
    
    'Icon
    With TitleIcon
        .Left = Left.Left + Left.Width + 2
        .Top = (Title.Height - .Height) / 2.5
    End With
    
    'Titlebar Caption
    With Caption1
        If TitleIcon.Visible = True Then
        .Left = TitleIcon.Left + TitleIcon.Width + 3
        Else
        .Left = Left.Left + Left.Width + 2.5
        End If
        .Top = ((Title.Height - 13) / 2) - 1
        .Width = MinimizeButton.Left - TitleIcon.Left - TitleIcon.Width - 10
        If MinimizeButton.Visible = False Then
            .Width = MaximizeButton.Left - TitleIcon.Left - TitleIcon.Width - 10
        End If
        If MinimizeButton.Visible = False And TitleIcon.Visible = False Then
            .Width = MaximizeButton.Left - Left.Left - Left.Width - 10
        End If
        If MinimizeButton.Visible = False And MaximizeButton.Visible = False Then
            .Width = CloseButton.Left - TitleIcon.Left - TitleIcon.Width - 10
        End If
        If MinimizeButton.Visible = False And MaximizeButton.Visible = False And TitleIcon.Visible = False Then
            .Width = CloseButton.Left - Left.Left - Left.Width - 10
        End If
        
        '.Height = 13
    End With
    
    With Caption2
        If TitleIcon.Visible = True Then
        .Left = TitleIcon.Left + TitleIcon.Width + 2
        Else
        .Left = Left.Left + Left.Width + 1.5
        End If
        .Top = ((Title.Height - 13) / 2) '+ 1
        .Width = Caption1.Width
       ' .Height = 13
    End With
    'Checks if it should have transparent corners
    If bTransparent = True Then
        ReTrans
    End If
End Sub

Private Sub TransparentEdges()
    'This is used as a safe guard set when the application starts,
    'otherwise the control would set the corners transparent at design time
    bTransparent = True
    RePos
End Sub

Private Sub ReTrans()
    Dim Add As Long
    Dim Sum As Long
    
    Dim X As Single
    Dim Y As Single
    
    If UserControl.Height < 615 Then UserControl.Height = 615   'Checks that form
    If UserControl.Width < 1695 Then UserControl.Width = 1695   'is not too small
    
    X = UserControl.Width / Screen.TwipsPerPixelX   'Registers the size of the
    Y = UserControl.Height / Screen.TwipsPerPixelY  'form in pixels
    
    Sum = CreateRectRgn(5, 0, X - 5, 1)
    CombineRgn Sum, Sum, CreateRectRgn(3, 1, X - 3, 2), 2
    CombineRgn Sum, Sum, CreateRectRgn(2, 2, X - 2, 3), 2
    CombineRgn Sum, Sum, CreateRectRgn(1, 3, X - 1, 4), 2
    CombineRgn Sum, Sum, CreateRectRgn(1, 4, X - 1, 5), 2
    CombineRgn Sum, Sum, CreateRectRgn(0, 5, X, Y), 2
    SetWindowRgn UserControl.ContainerHwnd, Sum, True   'Sets corners transparent
End Sub

Private Sub Caption1_Change()
    Caption2.Caption = Caption1.Caption
End Sub

Private Sub CloseButton_Click()
On Error GoTo EF
    Unload UserControl.Parent
EF:
End Sub

Private Sub CloseButton_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        CloseButton.Picture = Cb_CLose(3).Picture
    End If
End Sub

Private Sub CloseButton_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        CloseButton.Picture = Cb_CLose(3).Picture
    Else
        CloseButton.Picture = Cb_CLose(2).Picture
    End If
End Sub
Private Sub MYFORM_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CloseButton.Picture = Cb_CLose(1).Picture
    
End Sub
Private Sub Title_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 On Error GoTo H
    CloseButton.Picture = Cb_CLose(1).Picture
H:
End Sub

Private Sub TitleRight_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CloseButton.Picture = Cb_CLose(1).Picture
End Sub

Private Sub UserControl_Initialize()
    bTransparent = False  'So we do not set the corners transparent while still in design mode
    RePos   'Reposition
End Sub

Private Sub UserControl_Resize()
    RePos   'Reposition
End Sub

Private Sub Title_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Lets user move parent form
    Call ReleaseCapture: Call SendMessage(UserControl.ContainerHwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
End Sub

Private Sub TitleLeft_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Lets user move parent form
    Call ReleaseCapture: Call SendMessage(UserControl.ContainerHwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
End Sub

Private Sub TitleRight_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Lets user move parent form
    Call ReleaseCapture: Call SendMessage(UserControl.ContainerHwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
End Sub

Private Sub Caption1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Lets user move parent form
    Call ReleaseCapture: Call SendMessage(UserControl.ContainerHwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
End Sub

Private Sub Caption2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Lets user move parent form
    Call ReleaseCapture: Call SendMessage(UserControl.ContainerHwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
End Sub

Private Function DefaultBackgroundColor() As String
    DefaultBackgroundColor = &HEAF1F1   'Returns a common off-white Windows XP color
End Function
Private Sub LoadXP()
Dim IpForm As Form
    TitleIcon.Visible = False
    Set IpForm = UserControl.Parent
    Set MyForm = IpForm
    Caption1.Caption = IpForm.Caption
    IpForm.Width = UserControl.Width
    IpForm.Height = UserControl.Height
    If IpForm.BorderStyle <> 0 Then
        IpForm.Height = UserControl.Height + 375
    End If
    SetStyle IpForm
    UserControl.Width = IpForm.Width    ':   XP_Name.Top = 0
    UserControl.Height = IpForm.Height  ': XP_Name.Left = 0
    ReTransObj IpForm
    DoEvents
    IpForm.Hide
End Sub
Private Sub ReTransObj(IpObject As Object)
    Dim Add As Long
    Dim Sum As Long
    Dim X As Single
    Dim Y As Single
    If IpObject.Height < 615 Then IpObject.Height = 615   'Checks that form
    If IpObject.Width < 1695 Then IpObject.Width = 1695   'is not too small
    X = IpObject.Width / Screen.TwipsPerPixelX   'Registers the size of the
    Y = IpObject.Height / Screen.TwipsPerPixelY  'form in pixels
    Sum = CreateRectRgn(5, 0, X - 5, 1)
    CombineRgn Sum, Sum, CreateRectRgn(3, 1, X - 3, 2), 2
    CombineRgn Sum, Sum, CreateRectRgn(2, 2, X - 2, 3), 2
    CombineRgn Sum, Sum, CreateRectRgn(1, 3, X - 1, 4), 2
    CombineRgn Sum, Sum, CreateRectRgn(1, 4, X - 1, 5), 2
    CombineRgn Sum, Sum, CreateRectRgn(0, 5, X, Y), 2
    SetWindowRgn IpObject.hWnd, Sum, True   'Sets corners transparent
End Sub
Private Sub SetStyle(IpForm As Object)
    Dim lCurrentSettings As Long
    Const WS_MINIMIZEBOX = &H20000
    Const WS_MAXIMIZEBOX = &H10000
    Const WS_THICKFRAME = &H40000
    Const WS_DLGFRAME = &H400000
    Const WS_CAPTION = &HC00000
    lCurrentSettings = GetWindowLong(IpForm.hWnd, GWL_STYLE)
    lCurrentSettings = lCurrentSettings And Not WS_THICKFRAME
    lCurrentSettings = lCurrentSettings And Not WS_DLGFRAME
    lCurrentSettings = lCurrentSettings And Not WS_CAPTION
    lCurrentSettings = lCurrentSettings And Not WS_MINIMIZEBOX
    lCurrentSettings = lCurrentSettings And Not WS_MAXIMIZEBOX
    lCurrentSettings = lCurrentSettings Or WS_SYSMENU
    SetWindowLong IpForm.hWnd, GWL_STYLE, lCurrentSettings
    SetWindowPos IpForm.hWnd, 0, IpForm.Left / 15, IpForm.Top / 15, (IpForm.Width / 15), (IpForm.Height / 15), &H40
    If IpForm.BorderStyle <> 0 Then
    IpForm.Height = IpForm.Height - 365
    End If
    IpForm.Left = (Screen.Width - IpForm.Width) / 2
    IpForm.Top = (Screen.Height - IpForm.Height) / 2
End Sub

Private Sub UserControl_Terminate()
    Set MyForm = Nothing
End Sub
Private Function CapWidth() As Long
    CapWidth = Caption1.Width
End Function
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Caption1,Caption1,-1,Caption
Private Property Get Caption() As String
    Caption = Caption1.Caption
End Property

Private Property Let Caption(ByVal New_Caption As String)
    Caption1.Caption() = New_Caption
    PropertyChanged "Caption"
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Caption1.Caption = PropBag.ReadProperty("Caption", "Osen Kusnadi")
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Caption", Caption1.Caption, "Osen Kusnadi")
End Sub

Private Sub SetMyCurentForm(MyObj As Object)
    Set MyForm = MyObj
End Sub

