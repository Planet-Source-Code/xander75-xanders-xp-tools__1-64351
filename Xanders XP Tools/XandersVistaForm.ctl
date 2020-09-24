VERSION 5.00
Begin VB.UserControl XandersVistaForm 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   BackStyle       =   0  'Transparent
   ClientHeight    =   360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3135
   ScaleHeight     =   360
   ScaleWidth      =   3135
   ToolboxBitmap   =   "XandersVistaForm.ctx":0000
   Begin VB.Image imgClose 
      Height          =   210
      Left            =   2400
      Picture         =   "XandersVistaForm.ctx":0312
      Top             =   0
      Width           =   570
   End
   Begin VB.Image imgMax 
      Height          =   210
      Left            =   2040
      Picture         =   "XandersVistaForm.ctx":09AC
      Top             =   0
      Width           =   360
   End
   Begin VB.Image imgMin 
      Height          =   210
      Left            =   1630
      Picture         =   "XandersVistaForm.ctx":0DDE
      Top             =   0
      Width           =   405
   End
   Begin VB.Image imgRightMask 
      Height          =   45
      Left            =   480
      Picture         =   "XandersVistaForm.ctx":12B8
      Top             =   120
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Image imgLeftMask 
      Height          =   45
      Left            =   240
      Picture         =   "XandersVistaForm.ctx":15F2
      Top             =   120
      Visible         =   0   'False
      Width           =   45
   End
End
Attribute VB_Name = "XandersVistaForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private moForm As Form
Private WithEvents moveForm As Form
Attribute moveForm.VB_VarHelpID = -1

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

' Move a Titleless Window
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const HTCAPTION = 2
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const WM_SYSCOMMAND = &H112

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

' Show a Form in the Taskbar
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

Private Sub moveForm_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngReturnValue As Long

    If Button = 1 Then
        If Y <= 22 Then
            Call ReleaseCapture
            lngReturnValue = SendMessage(moveForm.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
        End If
    End If
End Sub

Private Sub UserControl_InitProperties()
    Set moForm = UserControl.Parent
    moForm.AutoRedraw = True
    moForm.BorderStyle = 0
    UserControl.Extender.Left = moForm.Left
    UserControl.Extender.Top = moForm.Top
    UserControl.Height = moForm.Height
    UserControl.Width = moForm.Width
    imgMin.Left = UserControl.Width - 1450
    imgMin.Top = 0
    imgMax.Left = UserControl.Width - 1055
    imgMax.Top = 0
    imgClose.Left = UserControl.Width - 695
    UserControl.Extender.Align = vbAlignTop
    Call UserControl_Paint
    MakeTransparent moForm, &HFF00FF
    MakeSemiTransparent UserControl.Parent.hWnd, 50
End Sub

Private Sub UserControl_Paint()
    Set moForm = UserControl.Parent
    Dim wid As Long
    Dim hgt As Long
    UserControl.PaintPicture imgLeftMask.Picture, 0, 0
    UserControl.PaintPicture imgRightMask.Picture, moForm.Width - 45, 0
    moForm.Picture = UserControl.Image
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next
    Set moForm = UserControl.Parent
    Set moveForm = UserControl.Parent
    Call UserControl_Paint
    MakeTransparent moForm, &HFF00FF
    MakeSemiTransparent UserControl.Parent.hWnd, 50
    ShowInTheTaskbar moForm.hWnd, True
End Sub

Private Sub UserControl_Resize()
    Set moForm = UserControl.Parent
    UserControl.Height = moForm.Height
    UserControl.Width = moForm.Width
    imgMin.Left = UserControl.Width - 1450
    imgMin.Top = 0
    imgMax.Left = UserControl.Width - 1055
    imgMax.Top = 0
    imgClose.Left = UserControl.Width - 695
    imgClose.Top = 0
End Sub
