VERSION 5.00
Begin VB.UserControl XandersXPTitleBar 
   AutoRedraw      =   -1  'True
   BackStyle       =   0  'Transparent
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.PictureBox picForm 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   660
      Left            =   360
      Picture         =   "XandersXPTitleBar.ctx":0000
      ScaleHeight     =   660
      ScaleWidth      =   3045
      TabIndex        =   3
      Top             =   1320
      Visible         =   0   'False
      Width           =   3045
   End
   Begin VB.PictureBox picRight 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   345
      Left            =   3000
      Picture         =   "XandersXPTitleBar.ctx":6972
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
      Picture         =   "XandersXPTitleBar.ctx":6AC8
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
      Picture         =   "XandersXPTitleBar.ctx":6C1E
      ScaleHeight     =   405
      ScaleWidth      =   2385
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   2385
   End
   Begin VB.Image imgRight 
      Height          =   180
      Left            =   4200
      Picture         =   "XandersXPTitleBar.ctx":966C
      Top             =   120
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgGreen 
      Height          =   240
      Left            =   780
      Picture         =   "XandersXPTitleBar.ctx":9C06
      Top             =   60
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgYellow 
      Height          =   240
      Left            =   420
      Picture         =   "XandersXPTitleBar.ctx":A190
      Top             =   60
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgRed 
      Height          =   240
      Left            =   60
      Picture         =   "XandersXPTitleBar.ctx":A71A
      Top             =   60
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "XandersXPTitleBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Private moForm As Form
Private WithEvents moveForm As Form
Attribute moveForm.VB_VarHelpID = -1

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
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

'Our declarations for retrieving colors
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long

' Move a Titleless Window
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const HTCAPTION = 2
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const WM_SYSCOMMAND = &H112

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
        SetWindowRgn Frm.hwnd, rgnMain, True 'Apply rgnMain to our form
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
    SetWindowRgn Frm.hwnd, rgnMain, True 'Apply to our window
 
End Function

Private Sub moveForm_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngReturnValue As Long

    If Button = 1 Then
        If Y <= 22 Then
            Call ReleaseCapture
            lngReturnValue = SendMessage(moveForm.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
        End If
    End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next
    Set moForm = UserControl.Parent
    Set moveForm = UserControl.Parent
    MakeTransparent moForm, &HFF00FF
End Sub

Private Sub UserControl_Resize()
    Set moForm = UserControl.Parent
    UserControl.Extender.Left = moForm.Left
    UserControl.Extender.Top = moForm.Top
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
    Call UserControl_Paint
    MakeTransparent moForm, &HFF00FF
End Sub

Private Sub UserControl_Paint()
    Set moForm = UserControl.Parent
    Dim wid As Long
    Dim hgt As Long
    
    moForm.AutoRedraw = True
    
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



