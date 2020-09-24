VERSION 5.00
Begin VB.UserControl XandersOSXForm 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   BackStyle       =   0  'Transparent
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12450
   ScaleHeight     =   3600
   ScaleWidth      =   12450
   ToolboxBitmap   =   "XandersOSXForm.ctx":0000
   Begin VB.PictureBox picForm 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   660
      Left            =   120
      Picture         =   "XandersOSXForm.ctx":0312
      ScaleHeight     =   660
      ScaleWidth      =   3045
      TabIndex        =   3
      Top             =   1080
      Visible         =   0   'False
      Width           =   3045
   End
   Begin VB.PictureBox picRight 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   345
      Left            =   3000
      Picture         =   "XandersOSXForm.ctx":6C84
      ScaleHeight     =   345
      ScaleWidth      =   60
      TabIndex        =   2
      Top             =   480
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.PictureBox picLeft 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   345
      Left            =   240
      Picture         =   "XandersOSXForm.ctx":6DDA
      ScaleHeight     =   345
      ScaleWidth      =   60
      TabIndex        =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.PictureBox picMiddle 
      BorderStyle     =   0  'None
      Height          =   405
      Left            =   480
      Picture         =   "XandersOSXForm.ctx":6F30
      ScaleHeight     =   405
      ScaleWidth      =   2385
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   2385
   End
   Begin VB.Image imgRightSilver 
      Height          =   345
      Left            =   12120
      Picture         =   "XandersOSXForm.ctx":997E
      Top             =   2160
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.Image imgMiddleSilver 
      Height          =   345
      Left            =   360
      Picture         =   "XandersOSXForm.ctx":9AD4
      Top             =   2160
      Visible         =   0   'False
      Width           =   11715
   End
   Begin VB.Image imgLeftSilver 
      Height          =   345
      Left            =   240
      Picture         =   "XandersOSXForm.ctx":16DAE
      Top             =   2160
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.Image imgFormSilver 
      Height          =   795
      Left            =   240
      Picture         =   "XandersOSXForm.ctx":16F04
      Top             =   2640
      Visible         =   0   'False
      Width           =   11745
   End
   Begin VB.Image imgRightDefault 
      Height          =   345
      Left            =   7080
      Picture         =   "XandersOSXForm.ctx":35636
      Top             =   480
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.Image imgMiddleDefault 
      Height          =   345
      Left            =   4560
      Picture         =   "XandersOSXForm.ctx":3578C
      Top             =   480
      Visible         =   0   'False
      Width           =   2325
   End
   Begin VB.Image imgLeftDefault 
      Height          =   345
      Left            =   4320
      Picture         =   "XandersOSXForm.ctx":381DA
      Top             =   480
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.Image imgFormDefault 
      Height          =   660
      Left            =   4200
      Picture         =   "XandersOSXForm.ctx":38330
      Top             =   1080
      Visible         =   0   'False
      Width           =   3045
   End
   Begin VB.Image imgRight 
      Height          =   180
      Left            =   4200
      Picture         =   "XandersOSXForm.ctx":3ECA2
      Top             =   120
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgGreen 
      Height          =   240
      Left            =   780
      Picture         =   "XandersOSXForm.ctx":3F23C
      Top             =   60
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgYellow 
      Height          =   240
      Left            =   420
      Picture         =   "XandersOSXForm.ctx":3F7C6
      Top             =   60
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgRed 
      Height          =   240
      Left            =   60
      Picture         =   "XandersOSXForm.ctx":3FD50
      Top             =   60
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "XandersOSXForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Private moForm As Form
Private WithEvents moveForm As Form
Attribute moveForm.VB_VarHelpID = -1

Public Enum OSXStyles
    osxDefault = 0
    osxSilver = 1
End Enum

Private m_ColorSchemes As OSXStyles

'Our region combine consts
Private Const RGN_AND = 1 'Combines an intersection
Private Const RGN_OR = 2 'Creates a union of two regions
Private Const RGN_XOR = 3 'Creations a union of two objects with the exception of overlapping
Private Const RGN_DIFF = 4 'Combines two regions
Private Const RGN_COPY = 5 'Copy a region

'Our API declarations
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function OffsetRgn Lib "gdi32" (ByVal hRgn As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

'Our declarations for retrieving colors
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long

' Move a Titleless Window
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const HTCAPTION = 2
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const WM_SYSCOMMAND = &H112

' Show a Form in the Taskbar
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function ShowWindow Lib "user32" _
    (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
    
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_APPWINDOW = &H40000

Const SW_HIDE = 0
Const SW_NORMAL = 1

Private Sub ShowInTheTaskbar(hWnd As Long, bShow As Boolean)
    Dim lStyle As Long
    
    ShowWindow hWnd, SW_HIDE
    
    lStyle = GetWindowLong(hWnd, GWL_EXSTYLE)
    
    If bShow = False Then
        If lStyle And WS_EX_APPWINDOW Then
            lStyle = lStyle - WS_EX_APPWINDOW
        End If
    Else
        lStyle = lStyle Or WS_EX_APPWINDOW
    End If
    
    SetWindowLong hWnd, GWL_EXSTYLE, lStyle
    
    App.TaskVisible = bShow
    
    ShowWindow hWnd, SW_NORMAL
End Sub

Private Function MakeTransparent(ByRef Frm As Form, ByVal TrnsColor As Long)
 
    Frm.BorderStyle = 0
     
    Dim ScaleSize As Long
    Dim Width, Height As Long 'Width and height of the image on our form
    Dim rgnMain As Long 'The main region which will be skinned then will be applied to our form
    Dim X, Y As Long 'Variables containing current X, Y in loop below
    Dim rgnPixel As Long 'A single pixel to be cut out of our image
    Dim rgbColor As Long 'A variable to store a color in the loop below
    Dim dcMain As Long 'The temporary DC of where all the skinning takes place
    Dim bmpMain As Long '1x1 bitmap created when dcMain is created
    
    ScaleSize = Frm.ScaleMode
    Frm.ScaleMode = 3 'Set the scale mode to pixels
    
    'This will get the height and width of the image on our form
    Width = Frm.ScaleX(Frm.Width, vbTwips, vbPixels)
    Height = Frm.ScaleY(Frm.Height, vbTwips, vbPixels)
    'vbHimetric
'    Frm.Width = Width * Screen.TwipsPerPixelX
'    Frm.Height = Height * Screen.TwipsPerPixelY
    
    'This will create our basic region to fit the dimensions of our
    'forms image
    rgnMain = CreateRectRgn(0, 0, Width, Height)
    
    'This will create a DC where all the skinning takes place
    dcMain = CreateCompatibleDC(Frm.hdc)
    bmpMain = SelectObject(dcMain, Frm.Picture.Handle)
    
    For Y = 0 To Height
    For X = 0 To Width
    
    rgbColor = GetPixel(dcMain, X, Y) 'Gets the color of a pixel on dcMain
    
    If rgbColor = TrnsColor Then 'If we found a mask color then cut it out of dcMain
    rgnPixel = CreateRectRgn(X, Y, X + 1, Y + 1) 'Create a region of a single pixel
    CombineRgn rgnMain, rgnMain, rgnPixel, RGN_XOR 'Cut it out
    DeleteObject rgnPixel 'Delete it from the memory
    End If
    
    Next X
    Next Y
     
    'Clear up our memory
    SelectObject dcMain, bmpMain
    DeleteDC dcMain
    DeleteObject bmpMain
    
    If rgnMain <> 0 Then
        SetWindowRgn Frm.hWnd, rgnMain, True 'Apply rgnMain to our form
    End If
     
    'Frm.ScaleMode = ScaleSize
 
End Function

Private Function RemoveTransparent(ByRef Frm As Form)

    Dim Width, Height As Long
    Dim rgnMain As Long
    
    'Get size of form
    Width = Frm.ScaleWidth
    Height = Frm.ScaleHeight
    
    rgnMain = CreateRectRgn(0, 0, Width, Height) 'Create a plain old region
    SetWindowRgn Frm.hWnd, rgnMain, True 'Apply to our window
 
End Function

Private Sub moveForm_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngReturnValue As Long

    If Button = 1 Then
        If Y <= 22 Then
            Call ReleaseCapture
            lngReturnValue = SendMessage(moveForm.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
        End If
    End If
End Sub

Public Property Get ColorScheme() As OSXStyles
    ColorScheme = m_ColorSchemes
End Property

Public Property Let ColorScheme(val As OSXStyles)
    m_ColorSchemes = val

    Call UserControl_Paint
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next
    Set moForm = UserControl.Parent
    Set moveForm = UserControl.Parent
    m_ColorSchemes = PropBag.ReadProperty("ColorScheme", m_def_ColorSchemes)
    Call UserControl_Paint
    MakeTransparent moForm, &HFF00FF
    ShowInTheTaskbar moForm.hWnd, True
End Sub

Private Sub UserControl_Resize()
    Set moForm = UserControl.Parent
    UserControl.Extender.Align = vbAlignTop
    UserControl.Height = moForm.Height
    UserControl.Width = moForm.Width
End Sub

Private Sub UserControl_InitProperties()
    Set moForm = UserControl.Parent
    moForm.AutoRedraw = True
    moForm.BorderStyle = 0
    UserControl.Extender.Left = moForm.Left
    UserControl.Extender.Top = moForm.Top
    UserControl.Height = moForm.Height
    UserControl.Width = moForm.Width
    UserControl.Extender.Align = vbAlignTop
    Call UserControl_Paint
    MakeTransparent moForm, &HFF00FF
End Sub

Private Sub UserControl_Paint()
    Set moForm = UserControl.Parent
    Dim wid As Long
    Dim hgt As Long
    
    moForm.AutoRedraw = True
    
    If m_ColorSchemes = osxDefault Then
        picLeft.Picture = imgLeftDefault.Picture
        picMiddle.Picture = imgMiddleDefault.Picture
        picRight.Picture = imgRightDefault.Picture
        picForm.Picture = imgFormDefault.Picture
    ElseIf m_ColorSchemes = osxSilver Then
        picLeft.Picture = imgLeftSilver.Picture
        picMiddle.Picture = imgMiddleSilver.Picture
        picRight.Picture = imgRightSilver.Picture
        picForm.Picture = imgFormSilver.Picture
    End If
    
    hgt = 0
    Do Until hgt >= moForm.Height
        wid = 0
        Do Until wid >= moForm.Width
            UserControl.PaintPicture picForm.Picture, wid, hgt
            wid = wid + picForm.Width
        Loop
        hgt = hgt + picForm.Height
    Loop
    UserControl.PaintPicture picLeft.Picture, 0, 0
    UserControl.PaintPicture picMiddle.Picture, 60, 0, moForm.Width - 60
    UserControl.PaintPicture picRight.Picture, moForm.Width - 60, 0
    
    UserControl.PaintPicture imgRed.Picture, 120, 60
    UserControl.PaintPicture imgYellow.Picture, 480, 60
    UserControl.PaintPicture imgGreen.Picture, 840, 60
    UserControl.PaintPicture imgRight.Picture, moForm.Width - 500, 90
    moForm.Picture = UserControl.Image
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("ColorScheme", m_ColorSchemes, m_def_ColorSchemes)
End Sub
