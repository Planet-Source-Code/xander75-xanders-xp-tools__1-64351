VERSION 5.00
Begin VB.UserControl XandersXPTaskBar 
   BackStyle       =   0  'Transparent
   ClientHeight    =   2025
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1485
   InvisibleAtRuntime=   -1  'True
   MaskColor       =   &H00FF00FF&
   MaskPicture     =   "XandersXPTaskBar.ctx":0000
   Picture         =   "XandersXPTaskBar.ctx":0A02
   ScaleHeight     =   2025
   ScaleWidth      =   1485
   ToolboxBitmap   =   "XandersXPTaskBar.ctx":1404
   Begin VB.Timer tmrExpand 
      Enabled         =   0   'False
      Interval        =   25
      Left            =   960
      Top             =   480
   End
   Begin VB.Timer tmrAppFocus 
      Enabled         =   0   'False
      Left            =   960
      Top             =   960
   End
   Begin VB.Timer tmrCheckMouseOver 
      Left            =   960
      Top             =   1440
   End
End
Attribute VB_Name = "XandersXPTaskBar"
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
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As RECT, ByVal fuWinIni As Long) As Long
Private Const SPI_GETWORKAREA = 48

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Type POINTAPI
        x As Long
        y As Long
End Type

' Alignment
Public Enum XTaskBarAlignment
'    vbBottomLeft = 0
'    vbBottomCenter = 1
'    vbBottomRight = 2
    vbLeftCenter = 0
    vbRightCenter = 1
    vbTopLeft = 2
    vbTopCenter = 3
    vbTopRight = 4
End Enum

Private m_TaskBarAlignment As XTaskBarAlignment

Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, rectangle As RECT) As Long

Private Const SWP_NOMOVE = 2
Private Const SWP_NOSIZE = 1
Private Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Declare Function SetWindowPos Lib "user32" _
            (ByVal hwnd As Long, _
            ByVal hWndInsertAfter As Long, _
            ByVal x As Long, _
            ByVal y As Long, _
            ByVal cx As Long, _
            ByVal cy As Long, _
            ByVal wFlags As Long) As Long

' ########## Member Vars #######
Private moForm As Form
Private mbHaveFocus As Boolean
Private mbMouseOver As Boolean

Dim wa_info As RECT
Dim wa_wid As Single
Dim wa_hgt As Single
Dim wa_left As Single
Dim wa_top As Single


' Center the form taking the task bar
' into account.
Private Sub PlaceForm(ByVal Frm As Form)
    
    wa_wid = 0
    wa_hgt = 0
    wa_left = 0
    wa_top = 0

    If SystemParametersInfo(SPI_GETWORKAREA, _
        0, wa_info, 0) <> 0 _
    Then
        ' We got the work area bounds.
        ' Center the form in the work area.
        wa_wid = ScaleX(wa_info.Right, vbPixels, vbTwips)
        wa_hgt = ScaleY(wa_info.Bottom, vbPixels, vbTwips)
        wa_left = ScaleX(wa_info.Left, vbPixels, vbTwips)
        wa_top = ScaleY(wa_info.Top, vbPixels, vbTwips)
    Else
        ' We did not get the work area bounds.
        ' Center the form on the whole screen.
        wa_wid = Screen.Width
        wa_hgt = Screen.Height
    End If
    
'    If m_TaskBarAlignment = vbBottomLeft Then
'        Frm.Move (wa_left + wa_left), (wa_hgt - 50)
'    ElseIf m_TaskBarAlignment = vbBottomCenter Then
'        Frm.Move (wa_wid - moForm.Width + wa_left) / 2, (wa_hgt - 50)
'    ElseIf m_TaskBarAlignment = vbBottomRight Then
'        Frm.Move (wa_wid - moForm.Width + wa_left), (wa_hgt - 50)
'    Else
    If m_TaskBarAlignment = vbLeftCenter Then
        Frm.Move (((wa_left + wa_left) - moForm.Width) + 50), (wa_hgt - moForm.Height + wa_top) / 2
    ElseIf m_TaskBarAlignment = vbRightCenter Then
        Frm.Move (wa_wid - 50), (wa_hgt - moForm.Height + wa_top) / 2
    ElseIf m_TaskBarAlignment = vbTopLeft Then
        Frm.Move (wa_left + wa_left), (((wa_top + wa_top) - moForm.Height) + 50)
    ElseIf m_TaskBarAlignment = vbTopCenter Then
        Frm.Move (wa_wid - moForm.Width + wa_left) / 2, (((wa_top + wa_top) - moForm.Height) + 50)
    ElseIf m_TaskBarAlignment = vbTopRight Then
        Frm.Move (wa_wid - moForm.Width), (((wa_top + wa_top) - moForm.Height) + 50)
    End If

    tmrCheckMouseOver.Enabled = True
    tmrCheckMouseOver.Interval = 200
    
'    tmrAppFocus.Enabled = True
'    tmrAppFocus.Interval = 200
    
End Sub

Private Sub lblCaption_Click()

End Sub

Private Sub tmrCheckMouseOver_Timer()
    Dim bThisMouseOver As Boolean
    
    Dim p As POINTAPI
    Call GetCursorPos(p)

    ' Check the screen coordinates of our window.  If it's not in ours, close ourselves up.
    Dim r As RECT
    Call GetWindowRect(moForm.hwnd, r)
    bThisMouseOver = (p.x >= r.Left And p.x <= r.Right And p.y >= r.Top And p.y <= r.Bottom)
    If (bThisMouseOver <> mbMouseOver) Then
        mbMouseOver = bThisMouseOver
        
        If mbMouseOver = True Then             ' Just got the mouse over
            tmrExpand.Enabled = True
            If (Not mbHaveFocus) Then
                
            End If

        ElseIf mbMouseOver = False Then        ' Just lost mouse over
            tmrExpand.Enabled = True
            
            If (Not mbHaveFocus) Then
                
            End If

        End If
    End If
End Sub

Private Sub tmrExpand_Timer()

    Dim new_left As Single
    Dim new_top As Single
    Call SetTopMost(moForm.hwnd)
    If mbMouseOver = True Then
'        If (m_TaskBarAlignment = vbBottomLeft Or m_TaskBarAlignment = vbBottomCenter Or m_TaskBarAlignment = vbBottomRight) Then
'            new_top = moForm.Top - 240
'            If (new_top + moForm.Height) < wa_hgt Then
'                new_top = wa_hgt - moForm.Height
'                tmrExpand.Enabled = False
'            End If
'            moForm.Top = new_top
'        Else
        If m_TaskBarAlignment = vbLeftCenter Then
            new_left = moForm.Left + 240
            If new_left > 0 Then
                new_left = 0
                tmrExpand.Enabled = False
            End If
            moForm.Left = new_left
        ElseIf m_TaskBarAlignment = vbRightCenter Then
            new_left = moForm.Left - 240
            If (new_left + moForm.Width) < wa_wid Then
                new_left = wa_wid - moForm.Width
                tmrExpand.Enabled = False
            End If
            moForm.Left = new_left
        ElseIf (m_TaskBarAlignment = vbTopLeft Or m_TaskBarAlignment = vbTopCenter Or m_TaskBarAlignment = vbTopRight) Then
            new_top = moForm.Top + 240
            If new_top > 0 Then
                new_top = 0
                tmrExpand.Enabled = False
            End If
            moForm.Top = new_top
        End If
    Else
'        If (m_TaskBarAlignment = vbBottomLeft Or m_TaskBarAlignment = vbBottomCenter Or m_TaskBarAlignment = vbBottomRight) Then
'            new_top = moForm.Top + 240
'            If new_top > (wa_hgt - 50) Then
'                new_top = wa_hgt - 50
'                tmrExpand.Enabled = False
'            End If
'            moForm.Top = new_top
'        Else
        If m_TaskBarAlignment = vbLeftCenter Then
            new_left = moForm.Left - 240
            If new_left < ((0 - moForm.Width) + 50) Then
                new_left = ((0 - moForm.Width) + 50)
                tmrExpand.Enabled = False
            End If
            moForm.Left = new_left
        ElseIf m_TaskBarAlignment = vbRightCenter Then
            new_left = moForm.Left + 240
            If new_left > (wa_wid - 50) Then
                new_left = wa_wid - 50
                tmrExpand.Enabled = False
            End If
            moForm.Left = new_left
        ElseIf (m_TaskBarAlignment = vbTopLeft Or m_TaskBarAlignment = vbTopCenter Or m_TaskBarAlignment = vbTopRight) Then
            new_top = moForm.Top - 240
            If new_top < ((0 - moForm.Height) + 50) Then
                new_top = ((0 - moForm.Height) + 50)
                tmrExpand.Enabled = False
            End If
            moForm.Top = new_top
        End If
    End If

End Sub

'Private Sub tmrAppFocus_Timer()
'    Dim bThisHaveFocus As Boolean
'
'    bThisHaveFocus = (GetForegroundWindow() = moForm.hwnd)
'
'    ' We've just changed states
'    If (bThisHaveFocus <> mbHaveFocus) Then
'        mbHaveFocus = bThisHaveFocus
'
'        If (mbHaveFocus) Then        ' Got focus
'
'        Else                        ' Lost focus
'
'        End If
'    End If
'End Sub

Private Sub SetTopMost(hwnd As Long)
    Call SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
End Sub

Private Sub UserControl_Initialize()
    m_TaskBarAlignment = vbTopCenter
End Sub

Public Property Get Alignment() As XTaskBarAlignment
    Alignment = m_TaskBarAlignment
End Property

Public Property Let Alignment(val As XTaskBarAlignment)
    m_TaskBarAlignment = val
    Call PlaceForm(moForm)
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    Set moForm = UserControl.Parent
    moForm.BorderStyle = 0
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Set moForm = UserControl.Parent

    m_TaskBarAlignment = PropBag.ReadProperty("Alignment", m_def_TaskBarAlignment)
'    m_NumSteps = PropBag.ReadProperty("NumSteps", m_def_NumSteps)
'    m_HangDown = PropBag.ReadProperty("HangDown", m_def_HangDown)
       
    Call PlaceForm(moForm)
    
    Call SetTopMost(moForm.hwnd)
End Sub

Private Sub UserControl_Resize()
    UserControl.Height = 480
    UserControl.Width = 480
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Alignment", m_TaskBarAlignment, m_def_TaskBarAlignment)
'    Call PropBag.WriteProperty("NumSteps", m_NumSteps, m_def_NumSteps)
'    Call PropBag.WriteProperty("HangDown", m_HangDown, m_def_HangDown)
End Sub





