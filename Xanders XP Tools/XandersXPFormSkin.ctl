VERSION 5.00
Begin VB.UserControl XandersXPFormSkin 
   BackStyle       =   0  'Transparent
   ClientHeight    =   810
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2445
   InvisibleAtRuntime=   -1  'True
   MaskColor       =   &H00FF00FF&
   MaskPicture     =   "XandersXPFormSkin.ctx":0000
   Picture         =   "XandersXPFormSkin.ctx":0876
   PropertyPages   =   "XandersXPFormSkin.ctx":10EC
   ScaleHeight     =   810
   ScaleWidth      =   2445
   ToolboxBitmap   =   "XandersXPFormSkin.ctx":111F
   Begin VB.PictureBox picMask 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   1080
      ScaleHeight     =   495
      ScaleWidth      =   1095
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "XandersXPFormSkin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"

Private moForm As Form
Private WithEvents moveForm As Form
Attribute moveForm.VB_VarHelpID = -1

Dim m_FormMove As Boolean
Dim m_MaskColor As OLE_COLOR
Private m_Picture As Picture

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
    Width = Frm.ScaleX(Frm.Picture.Width, vbHimetric, vbPixels)
    Height = Frm.ScaleY(Frm.Picture.Height, vbHimetric, vbPixels)
    
    Frm.Width = Width * Screen.TwipsPerPixelX
    Frm.Height = Height * Screen.TwipsPerPixelY
    
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
     
    Frm.ScaleMode = ScaleSize
 
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
    MsgBox m_FormMove
    If Button = 1 Then
        If m_FormMove = True Then
            Call ReleaseCapture
            lngReturnValue = SendMessage(moveForm.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
        End If
    End If
End Sub

Public Property Get FormMove() As Boolean
Attribute FormMove.VB_ProcData.VB_Invoke_Property = "Properties"
    FormMove = m_FormMove
End Property

Public Property Let FormMove(ByVal New_FormMove As Boolean)
    m_FormMove = New_FormMove
    PropertyChanged "FormMove"
End Property

Public Property Get MaskColor() As OLE_COLOR
    MaskColor = m_MaskColor
End Property

Public Property Let MaskColor(ByVal New_MaskColor As OLE_COLOR)
    m_MaskColor = New_MaskColor
    PropertyChanged "MaskColor"
End Property

Public Property Get Picture() As Picture
    Set Picture = picMask.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set moForm = UserControl.Parent
    Set picMask.Picture = New_Picture
    PropertyChanged "Picture"

    moForm.BorderStyle = 0
    moForm.Height = picMask.Height
    moForm.Width = picMask.Width
    moForm.Cls
    moForm.Picture = picMask.Picture
    moForm.Refresh

End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Set moForm = UserControl.Parent
    m_FormMove = PropBag.ReadProperty("FormMove", m_def_FormMove)
    m_MaskColor = PropBag.ReadProperty("MaskColor", m_def_MaskColor)
    Set picMask.Picture = PropBag.ReadProperty("Picture", m_def_Picture)
    
    MakeTransparent moForm, m_MaskColor
End Sub

Private Sub UserControl_Resize()
    UserControl.Height = 375
    UserControl.Width = 420
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("FormMove", m_FormMove, m_def_FormMove)
    Call PropBag.WriteProperty("MaskColor", m_MaskColor, m_def_MaskColor)
    Call PropBag.WriteProperty("Picture", picMask.Picture, m_def_Picture)
End Sub



