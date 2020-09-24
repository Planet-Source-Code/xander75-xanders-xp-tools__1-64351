VERSION 5.00
Begin VB.UserControl XandersXPCalendar 
   BackColor       =   &H00FFFFFF&
   BackStyle       =   0  'Transparent
   ClientHeight    =   3270
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2880
   ScaleHeight     =   3270
   ScaleWidth      =   2880
   ToolboxBitmap   =   "XandersXPCalendar.ctx":0000
   Begin VB.PictureBox picCalendar 
      BackColor       =   &H00EDF1F1&
      BorderStyle     =   0  'None
      Height          =   2575
      Left            =   15
      ScaleHeight     =   2580
      ScaleWidth      =   2700
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   480
      Visible         =   0   'False
      Width           =   2700
      Begin XandersXPTools.XandersXPCombobox cboMonth 
         Height          =   300
         Left            =   90
         TabIndex        =   1
         Top             =   0
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         Alignment       =   0
         BorderColorOver =   38631
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         HideColumnHeaders=   -1  'True
         Text            =   "cboMonth"
      End
      Begin XandersXPTools.XandersXPSpin spnYear 
         Height          =   300
         Left            =   1680
         TabIndex        =   2
         Top             =   0
         Visible         =   0   'False
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   529
         Alignment       =   0
         BorderColorOver =   38631
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         MaxLength       =   4
         Text            =   "0"
      End
      Begin VB.Label lblToday 
         BackStyle       =   0  'Transparent
         Caption         =   "Today:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   46
         Top             =   2280
         Width           =   2055
      End
      Begin VB.Label lblNumbers 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   13
         Left            =   2280
         TabIndex        =   45
         Top             =   960
         Width           =   255
      End
      Begin VB.Label lblNumbers 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   41
         Left            =   2280
         TabIndex        =   44
         Top             =   1920
         Width           =   255
      End
      Begin VB.Label lblNumbers 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   43
         Top             =   720
         Width           =   255
      End
      Begin VB.Label lblNumbers 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   42
         Top             =   720
         Width           =   255
      End
      Begin VB.Label lblNumbers 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   2
         Left            =   840
         TabIndex        =   41
         Top             =   720
         Width           =   255
      End
      Begin VB.Label lblNumbers 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   3
         Left            =   1200
         TabIndex        =   40
         Top             =   720
         Width           =   255
      End
      Begin VB.Label lblNumbers 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   4
         Left            =   1560
         TabIndex        =   39
         Top             =   720
         Width           =   255
      End
      Begin VB.Label lblNumbers 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   5
         Left            =   1920
         TabIndex        =   38
         Top             =   720
         Width           =   255
      End
      Begin VB.Label lblNumbers 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   6
         Left            =   2280
         TabIndex        =   37
         Top             =   720
         Width           =   255
      End
      Begin VB.Label lblNumbers 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   36
         Top             =   960
         Width           =   255
      End
      Begin VB.Label lblNumbers 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   8
         Left            =   480
         TabIndex        =   35
         Top             =   960
         Width           =   255
      End
      Begin VB.Label lblNumbers 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   9
         Left            =   840
         TabIndex        =   34
         Top             =   960
         Width           =   255
      End
      Begin VB.Label lblNumbers 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   10
         Left            =   1200
         TabIndex        =   33
         Top             =   960
         Width           =   255
      End
      Begin VB.Label lblNumbers 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   11
         Left            =   1560
         TabIndex        =   32
         Top             =   960
         Width           =   255
      End
      Begin VB.Label lblNumbers 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   12
         Left            =   1920
         TabIndex        =   31
         Top             =   960
         Width           =   255
      End
      Begin VB.Label lblNumbers 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   14
         Left            =   120
         TabIndex        =   30
         Top             =   1200
         Width           =   255
      End
      Begin VB.Label lblNumbers 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   15
         Left            =   480
         TabIndex        =   29
         Top             =   1200
         Width           =   255
      End
      Begin VB.Label lblNumbers 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   16
         Left            =   840
         TabIndex        =   28
         Top             =   1200
         Width           =   255
      End
      Begin VB.Label lblNumbers 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   17
         Left            =   1200
         TabIndex        =   27
         Top             =   1200
         Width           =   255
      End
      Begin VB.Label lblNumbers 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   18
         Left            =   1560
         TabIndex        =   26
         Top             =   1200
         Width           =   255
      End
      Begin VB.Label lblNumbers 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   19
         Left            =   1920
         TabIndex        =   25
         Top             =   1200
         Width           =   255
      End
      Begin VB.Label lblNumbers 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   20
         Left            =   2280
         TabIndex        =   24
         Top             =   1200
         Width           =   255
      End
      Begin VB.Label lblNumbers 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   21
         Left            =   120
         TabIndex        =   23
         Top             =   1440
         Width           =   255
      End
      Begin VB.Label lblNumbers 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   22
         Left            =   480
         TabIndex        =   22
         Top             =   1440
         Width           =   255
      End
      Begin VB.Label lblNumbers 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   23
         Left            =   840
         TabIndex        =   21
         Top             =   1440
         Width           =   255
      End
      Begin VB.Label lblNumbers 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   24
         Left            =   1200
         TabIndex        =   20
         Top             =   1440
         Width           =   255
      End
      Begin VB.Label lblNumbers 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   25
         Left            =   1560
         TabIndex        =   19
         Top             =   1440
         Width           =   255
      End
      Begin VB.Label lblNumbers 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   26
         Left            =   1920
         TabIndex        =   18
         Top             =   1440
         Width           =   255
      End
      Begin VB.Label lblNumbers 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   27
         Left            =   2280
         TabIndex        =   17
         Top             =   1440
         Width           =   255
      End
      Begin VB.Label lblNumbers 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   28
         Left            =   120
         TabIndex        =   16
         Top             =   1680
         Width           =   255
      End
      Begin VB.Label lblNumbers 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   29
         Left            =   480
         TabIndex        =   15
         Top             =   1680
         Width           =   255
      End
      Begin VB.Label lblNumbers 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   30
         Left            =   840
         TabIndex        =   14
         Top             =   1680
         Width           =   255
      End
      Begin VB.Label lblNumbers 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   31
         Left            =   1200
         TabIndex        =   13
         Top             =   1680
         Width           =   255
      End
      Begin VB.Label lblNumbers 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   32
         Left            =   1560
         TabIndex        =   12
         Top             =   1680
         Width           =   255
      End
      Begin VB.Label lblNumbers 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   33
         Left            =   1920
         TabIndex        =   11
         Top             =   1680
         Width           =   255
      End
      Begin VB.Label lblNumbers 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   34
         Left            =   2280
         TabIndex        =   10
         Top             =   1680
         Width           =   255
      End
      Begin VB.Label lblNumbers 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   35
         Left            =   120
         TabIndex        =   9
         Top             =   1920
         Width           =   255
      End
      Begin VB.Label lblNumbers 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   36
         Left            =   480
         TabIndex        =   8
         Top             =   1920
         Width           =   255
      End
      Begin VB.Label lblNumbers 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   37
         Left            =   840
         TabIndex        =   7
         Top             =   1920
         Width           =   255
      End
      Begin VB.Label lblNumbers 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   38
         Left            =   1200
         TabIndex        =   6
         Top             =   1920
         Width           =   255
      End
      Begin VB.Label lblNumbers 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   39
         Left            =   1560
         TabIndex        =   5
         Top             =   1920
         Width           =   255
      End
      Begin VB.Label lblNumbers 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   40
         Left            =   1920
         TabIndex        =   4
         Top             =   1920
         Width           =   255
      End
      Begin VB.Image imgDays 
         Height          =   1965
         Left            =   60
         Picture         =   "XandersXPCalendar.ctx":0312
         Top             =   315
         Width           =   2610
      End
   End
   Begin VB.TextBox txtXText 
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   15
      TabIndex        =   0
      Top             =   15
      Width           =   1080
   End
   Begin VB.Image imgArrow 
      Height          =   150
      Left            =   1080
      Picture         =   "XandersXPCalendar.ctx":10F78
      Top             =   120
      Width           =   225
   End
   Begin VB.Shape shBorder 
      BorderColor     =   &H00B99D7F&
      Height          =   330
      Left            =   0
      Top             =   0
      Width           =   1365
   End
   Begin VB.Image picButton 
      Height          =   255
      Left            =   1080
      Picture         =   "XandersXPCalendar.ctx":112CE
      Stretch         =   -1  'True
      Top             =   0
      Width           =   225
   End
   Begin VB.Shape shpCalendar 
      BackColor       =   &H00EDF1F1&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00DC9670&
      Height          =   2775
      Left            =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   2730
   End
End
Attribute VB_Name = "XandersXPCalendar"
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

Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, rectangle As RECT) As Long

' ########## Member Vars #######
Private moForm As Form


Dim m_AutoSelect As Boolean
Dim UserResize As Boolean
Dim m_BorderColorOver As OLE_COLOR
Dim SelectedBorderColor As OLE_COLOR

' Set the Margins within the Textbox
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const EM_SETMARGINS = &HD3
Private Const EC_LEFTMARGIN = &H1
Private Const EC_RIGHTMARGIN = &H2

' Events
Event Click()
Event DblClick()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Public Sub AboutBox()
    frmAbout.Show vbModal, Me
End Sub

Sub CalcCalendar()

    Select Case cboMonth.Text
        Case "January"
            GetMonth = 1
        Case "February"
            GetMonth = 2
        Case "March"
            GetMonth = 3
        Case "April"
            GetMonth = 4
        Case "May"
            GetMonth = 5
        Case "June"
            GetMonth = 6
        Case "July"
            GetMonth = 7
        Case "August"
            GetMonth = 8
        Case "September"
            GetMonth = 9
        Case "October"
            GetMonth = 10
        Case "November"
            GetMonth = 11
        Case "December"
            GetMonth = 12
    End Select
    
    Dim NewDate As Date
    
    For i = 1 To 41
        If Day(Date) = i Then
            GetDay = i
        End If
    Next
    NewDate = GetDay & "/" & GetMonth & "/" & spnYear.Text
    MonthStart = DateSerial(Year(NewDate), Month(NewDate), 1)
    'Number of days in the month
    NumDays = DateDiff("d", MonthStart, DateAdd("m", 1, MonthStart))
      
    FirstDayWD = Weekday(MonthStart, vbSunday)
    
    For i = 0 To 41
        lblNumbers(i).Caption = ""
        lblNumbers(i).Font.Bold = False
    Next
    
    ' Sunday
    If FirstDayWD = 1 Then
        j = 6
        For i = 1 To NumDays
            lblNumbers(j).Caption = i
            If Day(Date) = i Then
                lblNumbers(j).Font.Bold = True
            End If
            j = j + 1
        Next
    ' Monday
    ElseIf FirstDayWD = 2 Then
        j = 0
        For i = 1 To NumDays
            lblNumbers(j).Caption = i
            If Day(Date) = i Then
                lblNumbers(j).Font.Bold = True
            End If
            j = j + 1
        Next
    ' Tuesday
    ElseIf FirstDayWD = 3 Then
        j = 1
        For i = 1 To NumDays
            lblNumbers(j).Caption = i
            If Day(Date) = i Then
                lblNumbers(j).Font.Bold = True
            End If
            j = j + 1
        Next
    ' Wednesday
    ElseIf FirstDayWD = 4 Then
        j = 2
        For i = 1 To NumDays
            lblNumbers(j).Caption = i
            If Day(Date) = i Then
                lblNumbers(j).Font.Bold = True
            End If
            j = j + 1
        Next
    ' Thursday
    ElseIf FirstDayWD = 5 Then
        j = 3
        For i = 1 To NumDays
            lblNumbers(j).Caption = i
            If Day(Date) = i Then
                lblNumbers(j).Font.Bold = True
            End If
            j = j + 1
        Next
    ' Friday
    ElseIf FirstDayWD = 6 Then
        j = 4
        For i = 1 To NumDays
            lblNumbers(j).Caption = i
            If Day(Date) = i Then
                lblNumbers(j).Font.Bold = True
            End If
            j = j + 1
        Next
    ' Saturday
    ElseIf FirstDayWD = 7 Then
        j = 5
        For i = 1 To NumDays
            lblNumbers(j).Caption = i
            If Day(Date) = i Then
                lblNumbers(j).Font.Bold = True
            End If
            j = j + 1
        Next
    End If

End Sub

Private Sub cboMonth_Change()
    Call CalcCalendar
End Sub

Private Sub cboMonth_KeyPress(KeyAscii As Integer)
    ' Disallow the user from typing any key in the Database selection combobox
    If KeyAscii >= 0 And KeyAscii <= 127 Then KeyAscii = 0
End Sub

Private Sub imgArrow_Click()
    If shpCalendar.Visible = False Then
        shpCalendar.Visible = True
        picCalendar.Visible = True
        cboMonth.Visible = True
        spnYear.Visible = True
        imgDays.Visible = True
        If shBorder.Width <= shpCalendar.Width Then
            shpCalendar.Top = shBorder.Height + 15
            picCalendar.Top = shpCalendar.Top + 105
            UserControl.Height = shBorder.Height + shpCalendar.Height + 30
            UserControl.Width = shpCalendar.Width
        Else
            shpCalendar.Top = shBorder.Height + 15
            picCalendar.Top = shpCalendar.Top + 105
            UserControl.Height = shBorder.Height + shpCalendar.Height + 30
            UserControl.Width = shBorder.Width
        End If
    ElseIf shpCalendar.Visible = True Then
        shpCalendar.Visible = False
        picCalendar.Visible = False
        cboMonth.Visible = False
        spnYear.Visible = False
        imgDays.Visible = False
        UserResize = True
        If shBorder.Width <= shpCalendar.Width Then
            UserControl.Height = shBorder.Height
            UserControl.Width = shBorder.Width
        Else
            UserControl.Height = shBorder.Height
        End If
    End If
    
End Sub

Private Sub lblNumbers_Click(Index As Integer)
    
    If lblNumbers(Index).Caption = "" Then Exit Sub

    Select Case cboMonth.Text
        Case "January"
            GetMonth = 1
        Case "February"
            GetMonth = 2
        Case "March"
            GetMonth = 3
        Case "April"
            GetMonth = 4
        Case "May"
            GetMonth = 5
        Case "June"
            GetMonth = 6
        Case "July"
            GetMonth = 7
        Case "August"
            GetMonth = 8
        Case "September"
            GetMonth = 9
        Case "October"
            GetMonth = 10
        Case "November"
            GetMonth = 11
        Case "December"
            GetMonth = 12
    End Select
    
    Dim NewDate As Date
    NewDate = lblNumbers(Index).Caption & "/" & GetMonth & "/" & spnYear.Text
    txtXText.Text = NewDate

End Sub

Private Sub picButton_Click()
    Call imgArrow_Click
End Sub

Private Sub spnYear_Change()
    Call CalcCalendar
End Sub

Private Sub txtXText_GotFocus()
    SelectedBorderColor = shBorder.BorderColor
    shBorder.BorderColor = m_BorderColorOver
End Sub

Private Sub txtXText_LostFocus()
    If shpCalendar.Visible = False Then
        shBorder.BorderColor = SelectedBorderColor '&HB99D7F
        shpCalendar.Visible = False
        picCalendar.Visible = False
        imgDays.Visible = False
        UserResize = True
        Call UserControl_Resize
    End If
End Sub

Private Sub UserControl_Initialize()
    m_BorderColorOver = &H96E7&

    Dim left_margin As Integer
    Dim right_margin As Integer
    Dim long_value As Long

    left_margin = CInt(2)
    right_margin = CInt(2)
    long_value = right_margin * &H10000 + left_margin

    SendMessage txtXText.hWnd, _
        EM_SETMARGINS, _
        EC_LEFTMARGIN Or EC_RIGHTMARGIN, _
        long_value
    
    For i = 0 To 41
        lblNumbers(i).BackStyle = transparent
    Next

    cboMonth.AddItem "January"
    cboMonth.AddItem "February"
    cboMonth.AddItem "March"
    cboMonth.AddItem "April"
    cboMonth.AddItem "May"
    cboMonth.AddItem "June"
    cboMonth.AddItem "July"
    cboMonth.AddItem "August"
    cboMonth.AddItem "September"
    cboMonth.AddItem "October"
    cboMonth.AddItem "November"
    cboMonth.AddItem "December"
    
    GetMonth = DatePart("m", Date)
    Select Case GetMonth
        Case 1
            cboMonth.Text = "January"
        Case 2
            cboMonth.Text = "February"
        Case 3
            cboMonth.Text = "March"
        Case 4
            cboMonth.Text = "April"
        Case 5
            cboMonth.Text = "May"
        Case 6
            cboMonth.Text = "June"
        Case 7
            cboMonth.Text = "July"
        Case 8
            cboMonth.Text = "August"
        Case 9
            cboMonth.Text = "September"
        Case 10
            cboMonth.Text = "October"
        Case 11
            cboMonth.Text = "November"
        Case 12
            cboMonth.Text = "December"
    End Select
    
    spnYear.Text = DatePart("yyyy", Date)

    Call CalcCalendar
    
    lblToday.Caption = "Today: " & Date
End Sub

Private Sub UserControl_InitProperties()
    txtXText.Text = UserControl.Extender.Name
    UserResize = False
    shpCalendar.ZOrder 0
    picCalendar.ZOrder 0
End Sub

Private Sub UserControl_Resize()

    If shpCalendar.Visible = False Then
        If UserResize = False Then
            shBorder.Height = UserControl.Height
            shBorder.Width = UserControl.Width
        End If
         
        txtXText.Height = shBorder.Height - 25
        txtXText.Left = shBorder.Left + 10
        txtXText.Top = shBorder.Top + 15
        txtXText.Width = shBorder.Width - picButton.Width - 30
        
        picButton.Top = shBorder.Top + 15
        picButton.Left = shBorder.Width - 240
        picButton.Height = shBorder.Height - 30
        
        imgArrow.Top = (picButton.Height / 2) - (imgArrow.Height / 2) + 30
        imgArrow.Left = picButton.Left
        
    Else
        UserControl.Height = shBorder.Height + shpCalendar.Height + 30
    End If
    
End Sub

Public Property Get Alignment() As AlignmentConstants
    Alignment = txtXText.Alignment
End Property

Public Property Let Alignment(ByVal New_Alignment As AlignmentConstants)
    txtXText.Alignment() = New_Alignment
    PropertyChanged "Alignment"
End Property

Public Property Get AutoSelect() As Boolean
    AutoSelect = m_AutoSelect
End Property

Public Property Let AutoSelect(ByVal New_AutoSelect As Boolean)
    m_AutoSelect = New_AutoSelect
    PropertyChanged "AutoSelect"
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = txtXText.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    txtXText.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

Public Property Get BorderColor() As OLE_COLOR
    BorderColor = shBorder.BorderColor
End Property

Public Property Let BorderColor(ByVal New_BorderColor As OLE_COLOR)
    shBorder.BorderColor() = New_BorderColor
    PropertyChanged "BorderColor"
End Property

Public Property Get BorderColorOver() As OLE_COLOR
    BorderColorOver = m_BorderColorOver
End Property

Public Property Let BorderColorOver(ByVal New_BorderColorOver As OLE_COLOR)
    m_BorderColorOver = New_BorderColorOver
    PropertyChanged "BorderColorOver"
End Property

Public Property Get Enabled() As Boolean
    Enabled = txtXText.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    txtXText.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

' Return the font.
Public Property Get Font() As Font
    Set Font = txtXText.Font
End Property

' Set the font.
Public Property Set Font(ByVal New_Font As Font)
    Set txtXText.Font = New_Font
    PropertyChanged "Font"
End Property

Public Property Get FontBold() As Boolean
    FontBold = txtXText.FontBold
End Property

Public Property Let FontBold(ByVal New_FontBold As Boolean)
    txtXText.FontBold() = New_FontBold
    PropertyChanged "FontBold"
End Property

Public Property Get FontItalic() As Boolean
    FontItalic = txtXText.FontItalic
End Property

Public Property Let FontItalic(ByVal New_FontItalic As Boolean)
    txtXText.FontItalic() = New_FontItalic
    PropertyChanged "FontItalic"
End Property

Public Property Get FontName() As String
    FontName = txtXText.FontName
End Property

Public Property Let FontName(ByVal New_FontName As String)
    txtXText.FontName() = New_FontName
    PropertyChanged "FontName"
End Property

Public Property Get FontSize() As Single
    FontSize = txtXText.FontSize
End Property

Public Property Let FontSize(ByVal New_FontSize As Single)
    txtXText.FontSize() = New_FontSize
    PropertyChanged "FontSize"
End Property

Public Property Get FontStrikeOut() As Boolean
    FontStrikeOut = txtXText.FontStrikethru
End Property

Public Property Let FontStrikeOut(ByVal New_FontStrikeOut As Boolean)
    txtXText.FontStrikethru() = New_FontStrikeOut
    PropertyChanged "FontStrikeOut"
End Property

Public Property Get FontUnderline() As Boolean
    FontUnderline = txtXText.FontUnderline
End Property

Public Property Let FontUnderline(ByVal New_FontUnderline As Boolean)
    txtXText.FontUnderline() = New_FontUnderline
    PropertyChanged "FontUnderline"
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = txtXText.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    txtXText.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

Public Property Get MaxLength() As Long
    MaxLength = txtXText.MaxLength
End Property

Public Property Let MaxLength(ByVal New_MaxLength As Long)
    txtXText.MaxLength() = New_MaxLength
    PropertyChanged "MaxLength"
End Property

Public Property Get MouseIcon() As Picture
    Set MouseIcon = txtXText.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set txtXText.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

Public Property Get MousePointer() As Integer
    MousePointer = txtXText.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As Integer)
    txtXText.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

Public Property Get PasswordChar() As String
    PasswordChar = txtXText.PasswordChar
End Property

Public Property Let PasswordChar(ByVal New_PasswordChar As String)
    txtXText.PasswordChar() = New_PasswordChar
    PropertyChanged "PasswordChar"
End Property

' Return the Text.
Public Property Get Text() As String
    Text = txtXText.Text
End Property

' Set the Text.
Public Property Let Text(ByVal New_Text As String)
    txtXText.Text() = New_Text
    PropertyChanged "Text"
End Property

'Public Property Get MultiLine() As Boolean
'    MultiLine = txtXText.MultiLine
'End Property
'
'Public Property Let MultiLine(ByVal New_MultiLine As Boolean)
'    txtXText.MultiLine() = New_MultiLine
'    PropertyChanged "MultiLine"
'End Property

Private Sub txtXText_Click()
    RaiseEvent Click
End Sub

Private Sub txtXText_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub txtXText_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub txtXText_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub txtXText_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub txtXText_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub txtXText_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub txtXText_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

' Load saved properties.
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    txtXText.Alignment = PropBag.ReadProperty("Alignment", txtXText.Alignment)
    m_AutoSelect = PropBag.ReadProperty("AutoSelect", m_def_AutoSelect)
    txtXText.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    txtXText.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    shBorder.BorderColor = PropBag.ReadProperty("BorderColor", &HB99D7F)
    m_BorderColorOver = PropBag.ReadProperty("BorderColorOver", m_def_BorderColorOver)
    txtXText.Enabled = PropBag.ReadProperty("Enabled", 0)
    Set txtXText.Font = PropBag.ReadProperty("Font", Ambient.Font)
    txtXText.FontBold = PropBag.ReadProperty("FontBold", 0)
    txtXText.FontItalic = PropBag.ReadProperty("FontItalic", 0)
    txtXText.FontName = PropBag.ReadProperty("FontName", "")
    txtXText.FontSize = PropBag.ReadProperty("FontSize", 0)
    txtXText.FontStrikethru = PropBag.ReadProperty("FontStrikeOut", 0)
    txtXText.FontUnderline = PropBag.ReadProperty("FontUnderline", 0)
    txtXText.MaxLength = PropBag.ReadProperty("MaxLength", 0)
    Set txtXText.MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    txtXText.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    txtXText.PasswordChar = PropBag.ReadProperty("PasswordChar", "")
    Set txtXText.Font = PropBag.ReadProperty("Font", Ambient.Font)
    txtXText.Text = PropBag.ReadProperty("Text", txtXText.Text)
'    txtXText.MultiLine = PropBag.ReadProperty("MultiLine", 0)

    If m_AutoSelect = True Then
        txtXText.SelStart = 0
        txtXText.SelLength = Len(txtXText.Text)
    End If
    
    If m_ComputerInfo = None Then
        txtXText.Text = txtXText.Text
    ElseIf m_ComputerInfo = Computername Then
        txtXText.Text = Environ("ComputerName")
    ElseIf m_ComputerInfo = Username Then
        txtXText.Text = Environ("UserName")
    End If
    
End Sub

' Save properties.
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Alignment", txtXText.Alignment, "Alignment")
    Call PropBag.WriteProperty("AutoSelect", m_AutoSelect, m_def_AutoSelect)
    Call PropBag.WriteProperty("BackColor", txtXText.BackColor, &H80000005)
    Call PropBag.WriteProperty("ForeColor", txtXText.ForeColor, &H80000008)
    Call PropBag.WriteProperty("BorderColor", shBorder.BorderColor, &HB99D7F)
    Call PropBag.WriteProperty("BorderColorOver", m_BorderColorOver, m_def_BorderColorOver)
    Call PropBag.WriteProperty("Enabled", txtXText.Enabled, 0)
    Call PropBag.WriteProperty("Font", txtXText.Font, Ambient.Font)
    Call PropBag.WriteProperty("FontBold", txtXText.FontBold, 0)
    Call PropBag.WriteProperty("FontItalic", txtXText.FontItalic, 0)
    Call PropBag.WriteProperty("FontName", txtXText.FontName, "")
    Call PropBag.WriteProperty("FontSize", txtXText.FontSize, 0)
    Call PropBag.WriteProperty("FontStrikeOut", txtXText.FontStrikethru, 0)
    Call PropBag.WriteProperty("FontUnderline", txtXText.FontUnderline, 0)
    Call PropBag.WriteProperty("MaxLength", txtXText.MaxLength, 0)
    Call PropBag.WriteProperty("MouseIcon", txtXText.MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", txtXText.MousePointer, 0)
    Call PropBag.WriteProperty("PasswordChar", txtXText.PasswordChar, "")
    Call PropBag.WriteProperty("Text", txtXText.Text, txtXText.Text)
'    Call PropBag.WriteProperty("MultiLine", txtXText.MultiLine, 0)
End Sub







Private Sub XandersXPCombobox1_Click()

End Sub
