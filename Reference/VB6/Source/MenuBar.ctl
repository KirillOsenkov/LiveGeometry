VERSION 5.00
Begin VB.UserControl ctlMenuBar 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   ClientHeight    =   480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5400
   ClipControls    =   0   'False
   ControlContainer=   -1  'True
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   32
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   360
   ToolboxBitmap   =   "MenuBar.ctx":0000
   Begin VB.Timer TrackCapture 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   4200
      Top             =   0
   End
   Begin VB.PictureBox MenuBox 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   210
      Index           =   0
      Left            =   360
      ScaleHeight     =   14
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   104
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.Line linEast 
      BorderColor     =   &H80000015&
      Index           =   1
      Visible         =   0   'False
      X1              =   224
      X2              =   225
      Y1              =   -5
      Y2              =   35
   End
   Begin VB.Line linSouth 
      BorderColor     =   &H80000015&
      Index           =   1
      Visible         =   0   'False
      X1              =   174
      X2              =   210
      Y1              =   17
      Y2              =   15
   End
   Begin VB.Line linWest 
      BorderColor     =   &H80000014&
      Index           =   1
      Visible         =   0   'False
      X1              =   156
      X2              =   152
      Y1              =   2
      Y2              =   49
   End
   Begin VB.Line linNorth 
      BorderColor     =   &H80000014&
      Index           =   1
      Visible         =   0   'False
      X1              =   159
      X2              =   218
      Y1              =   2
      Y2              =   6
   End
   Begin VB.Menu mnuWhatsThisPopup 
      Caption         =   "popup"
      Visible         =   0   'False
      Begin VB.Menu mnuWhatsThis 
         Caption         =   "What's this?"
      End
   End
End
Attribute VB_Name = "ctlMenuBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'API
Const STRETCH_ANDSCANS = 1
Const STRETCH_DELETESCANS = 3
Const STRETCH_HALFTONE = 4
Const STRETCH_ORSCANS = 2

Const colIconTransparent = vbMagenta

Private Const LOGPIXELSY As Long = 90
Private Const LF_FACESIZE As Long = 32
Private Const SPI_GETNONCLIENTMETRICS As Long = 41
Private Const WS_CHILD As Long = &H40000000
Private Const WS_POPUP As Long = &H80000000
Private Const WS_EX_APPWINDOW = &H40000
Private Const GWL_STYLE = (-16)
Private Const GWL_EXSTYLE = (-20)
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOOWNERZORDER = &H200
Private Const SWP_NOREDRAW = &H8
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOZORDER = &H4
Private Const SWP_DRAWFRAME = SWP_FRAMECHANGED
Private Const BF_FLAT = &H4000
Private Const BF_MONO = &H8000
Private Const BF_SOFT = &H1000
Private Const BF_LEFT = &H1
Private Const BF_TOP = &H2
Private Const BF_RIGHT = &H4
Private Const BF_BOTTOM = &H8
Private Const BDR_RAISEDINNER = &H4
Private Const BDR_RAISEDOUTER = &H1
Private Const BDR_SUNKENINNER = &H8
Private Const BDR_SUNKENOUTER = &H2
Private Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)
Private Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
Private Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Private Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)
Private Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)

Private Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(0 To LF_FACESIZE - 1) As Byte
End Type
Private Type NONCLIENTMETRICS
    cbSize As Long
    iBorderWidth As Long
    iScrollWidth As Long
    iScrollHeight As Long
    iCaptionWidth As Long
    iCaptionHeight As Long
    lfCaptionFont As LOGFONT
    iSMCaptionWidth As Long
    iSMCaptionHeight As Long
    lfSMCaptionFont As LOGFONT
    iMenuWidth As Long
    iMenuHeight As Long
    lfMenuFont As LOGFONT
    lfStatusFont As LOGFONT
    lfMessageFont As LOGFONT
End Type

Private Declare Function GetStretchBltMode Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC As Long, ByVal nStretchMode As Long) As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As NONCLIENTMETRICS, ByVal fuWinIni As Long) As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal ptx As Long, ByVal pty As Long) As Long
Private Declare Function sndPlaySound Lib "WinMM.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Declare Function GetCapture Lib "user32" () As Long
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hDC As Long, lpRect As RECT) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
'_________________________________________________
'_________________________________________________
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

Private Const MenuLeftArrowSymbol = "3"
Private Const MenuRightArrowSymbol = "4"
Private Const MenuUpArrowSymbol = "5"
Private Const MenuDownArrowSymbol = "6"
Private Const SelMargin = 1
Private Const IconMargin = 4
Private Const GroupMargin = 3
Private Const MenuSmallOffSet = 2
Private Const MenuSoundEnabled As Boolean = False
Private Const SysColorTranslationBase = 2147483648#
Private Const MaxSubMenuLevels = 4
Private Const AutoShowSubMenus As Boolean = True
Private Const ColDecrease = 24

Private Const MenuHorizSpace = 10
Private Const MenuVerticSpace = 6
Private Const MenuHorizIndent = 6
Private Const MenuVerticIndent = 6
Private Const MenuPopupHorizIndent = 20

Private Const DragAreaIndent = 2
Private Const DragAreaSize = 4
'_________________________________________________

Enum MenuBorderStyle
    mbsNone = 0
    mbs2D = 1
    mbs3DThin = 2
    mbs3DThick = 3
End Enum
Enum PictureStyle
    psNormal = 0
    psStretched = 1
    psTiled = 2
End Enum
Enum OrientationStyle
    Horizontal = 0
    Vertical = 1
End Enum
Enum DragBarConstants
    BeginDrag
    Dragged
    EndDrag
    AboutToPlace
    Placed
End Enum

'Default Property Values:
Private Const m_def_Orientation = 0
Private Const m_def_PictureState = 2
Private Const m_def_GradientInverse = False
Private Const m_def_GradientVertical = True
Private Const m_def_Gradiented = True
Private Const m_def_GradientColor = vbButtonFace
Private Const m_def_BorderStyle = 2

Public m As Long 'global array identifier
Private m_WhatsThisHelp As Boolean

'Property Variables:
Dim m_Orientation As OrientationStyle
Dim m_PictureState As PictureStyle
Dim m_GradientInverse As Boolean
Dim m_GradientVertical As Boolean
Dim m_Gradiented As Boolean
Dim m_GradientColor As OLE_COLOR
Dim m_BorderStyle As MenuBorderStyle
Dim m_ToolbarMode As MenuBarState

'Event Declarations:
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseHover(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Event Resize() 'MappingInfo=UserControl,UserControl,-1,Resize
Attribute Resize.VB_Description = "Occurs when a form is first displayed or the size of an object changes."
Event Command(ByVal MenuNumber As Integer, ByVal MenuText As String)
Event CommandHover(ByVal MenuNumber As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal Shift As Integer)
Event DragBar(ByVal EventType As DragBarConstants, ByVal AuxInfo As Integer)
Event WhatsThisRequest(ByVal ItemIndex As Integer)
Event MustHideTooltip()

Dim DragArea As RECT
Public DragAreaVisible As Boolean
Dim ScrollAreaLeft As RECT
Dim ScrollAreaRight As RECT
Dim CanScrollLeft As Boolean
Dim CanScrollRight As Boolean
Dim ValidWidth As Long
Dim ItemOffset As Long
Dim MaxOffset As Long

Dim MenuCount() As Long
Dim TopItem() As Long
Dim SelItem() As Long
Dim DontCloseRect() As RECT

Dim CaptureIsSet As Boolean
Dim CaptureWindow As Long
Dim TotalItems As Integer
Dim UpperMenu As Integer
Dim DraggingMe As Boolean
Dim OX As Long, OY As Long, Rt As RECT
Dim OldAlign As Integer
Dim WasMouseDown As Integer
Dim SentCommand As Integer
Dim ControlLoaded As Boolean
Dim ShowFromWhere As AlignConstants
Dim CommandHoverSent As Integer
Dim OldGradient As Long, OldForeColor As Long, OldBackColor As Long
Dim IsAPopup As Integer
Dim ResizeBlock As Boolean
Dim WhatsThisCandidate As Integer
Dim DontRaiseResize As Boolean

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
UserControl.BackColor() = NotTooDark(New_BackColor)
PropertyChanged "BackColor"
Refresh
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
UserControl.ForeColor() = New_ForeColor
PropertyChanged "ForeColor"
If Not UserControl.Enabled Then
    OldForeColor = EnsureRGB(New_ForeColor)
    UserControl.ForeColor = RGB(Abs(Red(OldForeColor) - 2 * ColDecrease), Abs(Green(OldForeColor) - 2 * ColDecrease), Abs(Blue(OldForeColor) - 2 * ColDecrease))
End If
Refresh
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
If New_Enabled = UserControl.Enabled Then Exit Property
UserControl.Enabled = New_Enabled
PropertyChanged "Enabled"
If Not New_Enabled Then
    'OldGradient = m_GradientColor
    'OldBackColor = UserControl.BackColor
    'If OldGradient < 0 Then OldGradient = GetSysColor(OldGradient + SysColorTranslationBase)
    'm_GradientColor = RGB(Abs(Red(m_GradientColor) - ColDecrease), Abs(Green(m_GradientColor) - ColDecrease), Abs(Blue(m_GradientColor) - ColDecrease))
    'OldForeColor = UserControl.ForeColor
    'If OldForeColor < 0 Then OldForeColor = GetSysColor(OldForeColor + SysColorTranslationBase)
    'UserControl.BackColor = RGB(Abs(Red(OldBackColor) - ColDecrease), Abs(Green(OldBackColor) - ColDecrease), Abs(Blue(OldBackColor) - ColDecrease))
    'Dim hBitmap As Long
    'hBitmap = CreateCompatibleBitmap(UserControl.hDC, UserControl.ScaleWidth, UserControl.ScaleHeight)
    'DrawState UserControl.hDC, 0, 0, hBitmap, 0, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, DSS_DISABLED Or DST_BITMAP
    'DeleteObject hBitmap
Else
    'm_GradientColor = OldGradient
    'UserControl.ForeColor = OldForeColor
    'UserControl.BackColor = OldBackColor
End If
'Refresh
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
    Refresh
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get BorderStyle() As MenuBorderStyle
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
Attribute BorderStyle.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As MenuBorderStyle)
    m_BorderStyle = New_BorderStyle
    PropertyChanged "BorderStyle"
    Refresh
End Property

Private Sub MenuBox_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Over As Long, OldForeCol As Long, i As Long

Over = GetItemFromCursor(Index + 1, X, Y)
If Over = 0 Then
    Do
        TerminateMenu UpperMenu
    Loop Until UpperMenu = 1
    Exit Sub
End If

If Button = 1 And Shift = 0 Then
    ReleaseMouseCapture
    If MenuSoundEnabled Then sndPlaySound "MenuCommand", SND_ASYNC Or SND_NODEFAULT
    SelectItem Menus(m).Items(Over).Bounds, EDGE_SUNKEN, Index + 1
    OldForeCol = EnsureRGB(MenuBox(Index).ForeColor)
    MenuBox(i).FontBold = Menus(m).Items(Over).Checked
    MenuBox(i).FontUnderline = MenuBox(i).FontBold
    MenuBox(Index).ForeColor = RGB(Abs(Red(OldForeCol) - 2 * ColDecrease), Abs(Green(OldForeCol) - 2 * ColDecrease), Abs(Blue(OldForeCol) - 2 * ColDecrease))
    If Menus(m).Items(Over).Style <> GraphicsOnly Then
        MenuBox(Index).CurrentX = Menus(m).Items(Over).Bounds.Left
        If Menus(m).Items(Over).Style = TextAndGraphics Then
            MenuBox(Index).CurrentX = Menus(m).Items(Over).Bounds.Left + IconMargin + GetPicWidth(Menus(m).Items(Over).Icon)
        End If
        MenuBox(Index).CurrentY = Menus(m).Items(Over).Bounds.Top + ((Menus(m).Items(Over).Bounds.Bottom - Menus(m).Items(Over).Bounds.Top) - MenuBox(Index).TextHeight(Menus(m).Items(Over).Caption)) / 2
        MenuBox(Index).Print Menus(m).Items(Over).Caption
    End If
    MenuBox(Index).ForeColor = OldForeCol
    SelItem(Index + 1) = Over
    SentCommand = Over
    WasMouseDown = Index + 1
End If

End Sub

Private Sub MenuBox_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Over As Long, lPA As POINTAPI, WR1 As RECT, PX As Long, PY As Long

If WasMouseDown Then Exit Sub

If GetCapture = 0 Then SetMouseCapture MenuBox(Index).hWnd

GetWindowRect MenuBox(Index).hWnd, WR1
WR1.Bottom = WR1.Bottom + 1
WR1.Right = WR1.Right + 1
lPA.X = X
lPA.Y = Y
ClientToScreen MenuBox(Index).hWnd, lPA

If PtInRect(WR1, lPA.X, lPA.Y) = 0 And PtInRect(DontCloseRect(Index), lPA.X, lPA.Y) = 0 Then
    TerminateMenu Index + 1
    Exit Sub
End If

Over = GetItemFromCursor(Index + 1, X, Y)

If Over <> 0 And Button = 0 Then
    If CommandHoverSent <> Over Then CommandHoverSent = Over: RaiseEvent CommandHover(Over, X, Y, Shift)
    If SelItem(Index + 1) <> Over Then
        If SelItem(Index + 1) <> 0 Then SelectItem Menus(m).Items(TopItem(Index + 1)).Bounds, 0, Index + 1: SelItem(Index + 1) = 0
        If IsSubMenu(Over) Then
            'w
            
            If AutoShowSubMenus Then
                DontCloseRect(Index + 1) = Menus(m).Items(Over).Bounds
                DontCloseRect(Index + 1).Left = 0
                DontCloseRect(Index + 1).Right = MenuBox(Index).ScaleWidth + MenuSmallOffSet + 1
                lPA.X = MenuBox(Index).Left / Screen.TwipsPerPixelX
                lPA.Y = MenuBox(Index).Top / Screen.TwipsPerPixelY
                OffsetRect DontCloseRect(Index + 1), lPA.X, lPA.Y
                PX = DontCloseRect(Index + 1).Right - MenuSmallOffSet - 1
                PY = DontCloseRect(Index + 1).Top - SelMargin
                SelectItem Menus(m).Items(Over).Bounds, EDGE_SUNKEN, Index + 1
                linEast(Index + 1).Refresh
                linNorth(Index + 1).Refresh
                linSouth(Index + 1).Refresh
                linWest(Index + 1).Refresh
                SelItem(Index + 1) = Over
                PopUpChild PX, PY, Index + 2, Over + 1
            Else
                'don't show it automatically
            End If
        Else
            SelectItem Menus(m).Items(Over).Bounds, EDGE_RAISED, Index + 1
            SelItem(Index + 1) = Over
        End If
    End If
Else
    If SelItem(Index + 1) <> 0 Then
        SelectItem Menus(m).Items(SelItem(Index + 1)).Bounds, 0, Index + 1
        SelItem(Index + 1) = 0
    End If
End If

End Sub

Private Sub MenuBox_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Over As Long

If WasMouseDown = Index + 1 Then
    SelectItem Menus(m).Items(SentCommand).Bounds, 0, Index + 1 'EDGE_RAISED
    SelItem(Index + 1) = 0 'SentCommand
    'MenuBox(I).FontBold = Menus(M).Items(SentCommand).Checked
    'MenuBox(I).FontUnderline = MenuBox(I).FontBold
    If Menus(m).Items(SentCommand).Style <> GraphicsOnly Then
        MenuBox(Index).CurrentX = Menus(m).Items(SentCommand).Bounds.Left
        If Menus(m).Items(SentCommand).Style = TextAndGraphics Then
            MenuBox(Index).CurrentX = Menus(m).Items(SentCommand).Bounds.Left + IconMargin + GetPicWidth(Menus(m).Items(SentCommand).Icon)
        End If
        MenuBox(Index).CurrentY = Menus(m).Items(SentCommand).Bounds.Top + ((Menus(m).Items(SentCommand).Bounds.Bottom - Menus(m).Items(SentCommand).Bounds.Top) - MenuBox(Index).TextHeight(Menus(m).Items(SentCommand).Caption)) / 2
        MenuBox(Index).Print Menus(m).Items(SentCommand).Caption
    End If
    Over = SentCommand
    WasMouseDown = 0
    SentCommand = 0
    Do
        TerminateMenu UpperMenu
    Loop Until UpperMenu = 1
    DoEvents
    If IsSubMenu(Over) = False Then RaiseEvent Command(Over, Menus(m).Items(Over).Caption)
    'DoEvents
End If

End Sub

Private Sub mnuWhatsThis_Click()
RaiseEvent WhatsThisRequest(WhatsThisCandidate)
End Sub

Private Sub TrackCapture_Timer()
If GetCapture <> CaptureWindow Then
    ReleaseMouseCapture
End If
End Sub

Private Sub UserControl_Initialize()
On Local Error Resume Next

Me.MenuID = AddNewMenu

'M = 1 'UserControl.Extender.Index
'ReDim Menus(M).Items(1 To 1)
ReDim MenuCount(1 To 1)
ReDim TopItem(1 To MaxSubMenuLevels)
ReDim SelItem(1 To MaxSubMenuLevels)
ReDim DontCloseRect(0 To MaxSubMenuLevels)

DragAreaVisible = True
UserControl.BackColor = NotTooDark(colMenuBackGround)
UserControl.ForeColor = colMenuForeGround
SelectSystemFont
Exit Sub

End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hDC
Public Property Get hDC() As Long
Attribute hDC.VB_Description = "Returns a handle (from Microsoft Windows) to the object's device context."
    hDC = UserControl.hDC
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Picture
Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set Picture = UserControl.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set UserControl.Picture = New_Picture
    PropertyChanged "Picture"
    Refresh
End Property

Private Sub UserControl_Resize()
If Not ControlLoaded Or ResizeBlock Then Exit Sub
ResizeBlock = True
If UserControl.ScaleWidth < 4 Then UserControl.Width = 4 * Screen.TwipsPerPixelX
If UserControl.ScaleHeight < 4 Then UserControl.Height = 4 * Screen.TwipsPerPixelY
If Not DontRaiseResize Then RaiseEvent Resize
ResizeBlock = False
Refresh
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ScaleHeight
Public Property Get ScaleHeight() As Single
Attribute ScaleHeight.VB_Description = "Returns/sets the number of units for the vertical measurement of an object's interior."
    ScaleHeight = UserControl.ScaleHeight
End Property

'Public Property Let ScaleHeight(ByVal New_ScaleHeight As Single)
'    If New_ScaleHeight < 2 Then Exit Property
'    UserControl.ScaleHeight() = New_ScaleHeight
'    PropertyChanged "ScaleHeight"
'End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ScaleWidth
Public Property Get ScaleWidth() As Single
Attribute ScaleWidth.VB_Description = "Returns/sets the number of units for the horizontal measurement of an object's interior."
    ScaleWidth = UserControl.ScaleWidth
End Property

'Public Property Let ScaleWidth(ByVal New_ScaleWidth As Single)
'    If New_ScaleWidth < 2 Then Exit Property
'    UserControl.ScaleWidth() = New_ScaleWidth
'    PropertyChanged "ScaleWidth"
'End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
m_BorderStyle = m_def_BorderStyle
m_Gradiented = m_def_Gradiented
m_GradientColor = m_def_GradientColor
m_GradientInverse = m_def_GradientInverse
m_GradientVertical = m_def_GradientVertical
m_PictureState = m_def_PictureState
m_Orientation = m_def_Orientation
m_WhatsThisHelp = False
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Local Error Resume Next
ControlLoaded = True

'M = UserControl.Extender.Index
'ReDim Menus(M).Items(1 To 1)
'UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
'UserControl.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
Set Picture = PropBag.ReadProperty("Picture", Nothing)
'UserControl.ScaleHeight = PropBag.ReadProperty("ScaleHeight", 26)
'UserControl.ScaleWidth = PropBag.ReadProperty("ScaleWidth", 353)
m_Gradiented = PropBag.ReadProperty("Gradiented", m_def_Gradiented)
m_GradientColor = PropBag.ReadProperty("GradientColor", m_def_GradientColor)
m_GradientInverse = PropBag.ReadProperty("GradientInverse", m_def_GradientInverse)
m_GradientVertical = PropBag.ReadProperty("GradientVertical", m_def_GradientVertical)
m_PictureState = PropBag.ReadProperty("PictureState", m_def_PictureState)
m_Orientation = PropBag.ReadProperty("Orientation", m_def_Orientation)
m_WhatsThisHelp = PropBag.ReadProperty("WhatsThisHelp", False)
Exit Sub

End Sub

Private Sub UserControl_Terminate()
RemoveMenu m
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, &H80000012)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
    Call PropBag.WriteProperty("ScaleHeight", UserControl.ScaleHeight, 26)
    Call PropBag.WriteProperty("ScaleWidth", UserControl.ScaleWidth, 353)
    Call PropBag.WriteProperty("Gradiented", m_Gradiented, m_def_Gradiented)
    Call PropBag.WriteProperty("GradientColor", m_GradientColor, m_def_GradientColor)
    Call PropBag.WriteProperty("GradientInverse", m_GradientInverse, m_def_GradientInverse)
    Call PropBag.WriteProperty("GradientVertical", m_GradientVertical, m_def_GradientVertical)
    Call PropBag.WriteProperty("PictureState", m_PictureState, m_def_PictureState)
    Call PropBag.WriteProperty("Orientation", m_Orientation, m_def_Orientation)
    Call PropBag.WriteProperty("WhatsThisHelp", m_WhatsThisHelp, False)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,true
Public Property Get Gradiented() As Boolean
Attribute Gradiented.VB_Description = "Sets/returns whether the menu bar is painted in gradient color."
    Gradiented = m_Gradiented
End Property

Public Property Let Gradiented(ByVal New_Gradiented As Boolean)
    m_Gradiented = New_Gradiented
    PropertyChanged "Gradiented"
    Refresh
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get GradientColor() As OLE_COLOR
Attribute GradientColor.VB_Description = "Sets/returns the color of the gradient fill."
    GradientColor = m_GradientColor
End Property

Public Property Let GradientColor(ByVal New_GradientColor As OLE_COLOR)
m_GradientColor = New_GradientColor
PropertyChanged "GradientColor"
If Not UserControl.Enabled() Then
    OldGradient = EnsureRGB(m_GradientColor)
    m_GradientColor = RGB(Abs(Red(m_GradientColor) - ColDecrease), Abs(Green(m_GradientColor) - ColDecrease), Abs(Blue(m_GradientColor) - ColDecrease))
End If
Refresh
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,false
Public Property Get GradientInverse() As Boolean
Attribute GradientInverse.VB_Description = "Determines if source and destination fill colors are swapped."
    GradientInverse = m_GradientInverse
End Property

Public Property Let GradientInverse(ByVal New_GradientInverse As Boolean)
    m_GradientInverse = New_GradientInverse
    PropertyChanged "GradientInverse"
    Refresh
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,false
Public Property Get GradientVertical() As Boolean
Attribute GradientVertical.VB_Description = "Defines the direction of the gradient fill."
    GradientVertical = m_GradientVertical
End Property

Public Property Let GradientVertical(ByVal New_GradientVertical As Boolean)
    m_GradientVertical = New_GradientVertical
    PropertyChanged "GradientVertical"
    Refresh
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get PictureState() As PictureStyle
Attribute PictureState.VB_Description = "Defines in what way the picture will be displayed in the menubar."
    PictureState = m_PictureState
End Property

Public Property Let PictureState(ByVal New_PictureState As PictureStyle)
    m_PictureState = New_PictureState
    PropertyChanged "PictureState"
    Refresh
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,1,1,0
Public Property Get Orientation() As OrientationStyle
Attribute Orientation.VB_Description = "Indicates whether the menu bar is horizontal or vertical."
    Orientation = m_Orientation
End Property

Private Sub FillMenu(Optional ByVal ShouldRedraw As Boolean = True)
Dim Z As Long
TotalItems = UBound(Menus(m).Items)
If Menus(m).Items(1).Caption = "#" Then TotalItems = 0
ReDim TopItem(1 To MaxSubMenuLevels)
ReDim MenuCount(1 To MaxSubMenuLevels)
ReDim SelItem(1 To MaxSubMenuLevels)
ReDim DontCloseRect(0 To MaxSubMenuLevels)

UpperMenu = 1
TopItem(1) = 1
If TotalItems <> 0 Then
    For Z = 1 To TotalItems
        If Menus(m).Items(Z).SubLevel = 1 Then MenuCount(1) = MenuCount(1) + 1
    Next
End If

If ShouldRedraw Then
    
    Refresh
End If
End Sub

Private Sub DrawEdgeThin(qrc As RECT, ByVal lStyle As Long, Optional ByVal Level As Integer = 1)
Dim Col As Long

Select Case lStyle
    Case EDGE_RAISED
        If Level = 1 Then
            UserControl.Line (qrc.Left, qrc.Top)-(qrc.Right, qrc.Top), vb3DHighlight, BF
            UserControl.Line (qrc.Left, qrc.Top)-(qrc.Left, qrc.Bottom), vb3DHighlight, BF
            UserControl.Line (qrc.Right, qrc.Top)-(qrc.Right, qrc.Bottom), vb3DDKShadow, BF
            UserControl.Line (qrc.Left, qrc.Bottom)-(qrc.Right, qrc.Bottom), vb3DDKShadow, BF
        Else
            MenuBox(Level - 1).Line (qrc.Left, qrc.Top)-(qrc.Right, qrc.Top), vb3DHighlight, BF
            MenuBox(Level - 1).Line (qrc.Left, qrc.Top)-(qrc.Left, qrc.Bottom), vb3DHighlight, BF
            MenuBox(Level - 1).Line (qrc.Right, qrc.Top)-(qrc.Right, qrc.Bottom), vb3DDKShadow, BF
            MenuBox(Level - 1).Line (qrc.Left, qrc.Bottom)-(qrc.Right, qrc.Bottom), vb3DDKShadow, BF
        End If
    Case EDGE_SUNKEN
        If Level = 1 Then
            UserControl.Line (qrc.Left, qrc.Top)-(qrc.Right, qrc.Top), vb3DDKShadow, B
            UserControl.Line (qrc.Left, qrc.Top)-(qrc.Left, qrc.Bottom), vb3DDKShadow, B
            UserControl.Line (qrc.Right, qrc.Top)-(qrc.Right, qrc.Bottom), vb3DHighlight, B
            UserControl.Line (qrc.Left, qrc.Bottom)-(qrc.Right, qrc.Bottom), vb3DHighlight, B
        Else
            MenuBox(Level - 1).Line (qrc.Left, qrc.Top)-(qrc.Right, qrc.Top), vb3DDKShadow, B
            MenuBox(Level - 1).Line (qrc.Left, qrc.Top)-(qrc.Left, qrc.Bottom), vb3DDKShadow, B
            MenuBox(Level - 1).Line (qrc.Right, qrc.Top)-(qrc.Right, qrc.Bottom), vb3DHighlight, B
            MenuBox(Level - 1).Line (qrc.Left, qrc.Bottom)-(qrc.Right, qrc.Bottom), vb3DHighlight, B
        End If
    Case EDGE_ETCHED, EDGE_BUMP, 0
        If lStyle = EDGE_ETCHED Then Col = vb3DHighlight
        If lStyle = EDGE_BUMP Then Col = vb3DDKShadow
        If lStyle = 0 Then Col = UserControl.BackColor
        If Level = 1 Then
            UserControl.Line (qrc.Left, qrc.Top)-(qrc.Right, qrc.Bottom), Col, B
        Else
            MenuBox(Level - 1).Line (qrc.Left, qrc.Top)-(qrc.Right, qrc.Bottom), Col, B
        End If
End Select
End Sub

Private Sub SelectItem(qrc As RECT, ByVal lStyle As Long, Optional ByVal Level As Integer = 1)
Dim tVis As Boolean
tVis = True
Select Case lStyle
    Case EDGE_RAISED
        linEast(Level).BorderColor = vb3DDKShadow
        linNorth(Level).BorderColor = vb3DHighlight
        linSouth(Level).BorderColor = vb3DDKShadow
        linWest(Level).BorderColor = vb3DHighlight
    Case EDGE_SUNKEN
        linEast(Level).BorderColor = vb3DHighlight
        linNorth(Level).BorderColor = vb3DDKShadow
        linSouth(Level).BorderColor = vb3DHighlight
        linWest(Level).BorderColor = vb3DDKShadow
    Case Else
        tVis = False
End Select
linEast(Level).X1 = qrc.Right + SelMargin
linEast(Level).X2 = qrc.Right + SelMargin
linEast(Level).Y1 = qrc.Top - SelMargin
linEast(Level).Y2 = qrc.Bottom + SelMargin
linNorth(Level).X1 = qrc.Left - SelMargin
linNorth(Level).X2 = qrc.Right + SelMargin
linNorth(Level).Y1 = qrc.Top - SelMargin
linNorth(Level).Y2 = qrc.Top - SelMargin
linSouth(Level).X1 = qrc.Left - SelMargin
linSouth(Level).X2 = qrc.Right + 1 + SelMargin
linSouth(Level).Y1 = qrc.Bottom + SelMargin
linSouth(Level).Y2 = qrc.Bottom + SelMargin
linWest(Level).X1 = qrc.Left - SelMargin
linWest(Level).X2 = qrc.Left - SelMargin
linWest(Level).Y1 = qrc.Top - SelMargin
linWest(Level).Y2 = qrc.Bottom + SelMargin
linEast(Level).Visible = tVis
linNorth(Level).Visible = tVis
linSouth(Level).Visible = tVis
linWest(Level).Visible = tVis

If Not tVis Then SelItem(Level) = 0
End Sub

Private Function GetProperSize(Optional ByVal Level As Long = 1, Optional ByVal NewAlign As Integer = -1) As Integer
Dim Z As Long
Dim TVar As Long, TempH As Long, MaxH As Long, Max As Double, MaxW As Double, TempW As Double
If NewAlign = -1 Then NewAlign = UserControl.Extender.Align
If Level > 1 Then NewAlign = vbAlignLeft
UserControl.Font.Charset = DefaultFontCharset

Select Case NewAlign
    Case vbAlignTop, vbAlignBottom, vbAlignNone
        MaxH = 0
        TVar = (2 * MenuVerticIndent)
        If MenuCount(Level) > 0 Then
            For Z = TopItem(Level) To TotalItems
                If Menus(m).Items(Z).PopUp And IsAPopup = 0 Then Exit For
                If Menus(m).Items(Z).SubLevel = Level And Menus(m).Items(Z).Visible Then
                    TempH = 0
                    If Menus(m).Items(Z).Style <> TextOnly Then
                        TempH = GetPicHeight(Menus(m).Items(Z).Icon)
                    End If
                    If Menus(m).Items(Z).Style <> GraphicsOnly Then
                        TempH = Maximum(TempH, GetItemHeight(Z))
                    End If
                    If TempH > MaxH Then MaxH = TempH
                End If
                If Menus(m).Items(Z).SubLevel < Level Then Exit For
            Next
        End If
        TVar = MaxH + TVar
        
    Case vbAlignLeft, vbAlignRight
        If MenuCount(Level) > 0 Then
            Max = 0
            MaxW = 0
            For Z = TopItem(Level) To TotalItems
                If Menus(m).Items(Z).PopUp And IsAPopup = 0 Then Exit For
                If Menus(m).Items(Z).SubLevel = Level And Menus(m).Items(Z).Visible Then
                    TempW = 0
                    If Menus(m).Items(Z).Style <> TextOnly Then
                        TempW = TempW + GetPicWidth(Menus(m).Items(Z).Icon)
                    End If
                    If Menus(m).Items(Z).Style <> GraphicsOnly Then
                        TempW = TempW + GetItemWidth(Z)
                    End If
                    If Menus(m).Items(Z).Style = TextAndGraphics Then
                        TempW = TempW + IconMargin
                    End If
                    If Level = 1 And IsSubMenu(Z) Then TempW = TempW + 16
                    If TempW > MaxW Then MaxW = TempW: Max = Z
                End If
                If Menus(m).Items(Z).SubLevel < Level Then Exit For
            Next
        End If
        If Max <> 0 Then
            TVar = (2 * IIf(Level > 1, MenuPopupHorizIndent, MenuHorizIndent) + MaxW)
        Else
            If Level = 1 Then
                TVar = (DragAreaSize + 2 * DragAreaIndent)
            Else
                TVar = 2 * IIf(Level > 1, MenuPopupHorizIndent, MenuHorizIndent)
            End If
        End If
End Select

GetProperSize = TVar '+ 4

End Function

Public Function GetItemFromCursor(ByVal Level As Integer, ByVal X As Long, ByVal Y As Long) As Long
Dim OldBounds As RECT, tItem As Long, PX As Long, Z As Long
If MenuCount(Level) > 0 Then
    Z = TopItem(Level)
    Do While Z <= TotalItems And tItem = 0
        If Menus(m).Items(Z).SubLevel < Level Then Exit Do
        If Menus(m).Items(Z).SubLevel = Level And RectInVisibleArea(Menus(m).Items(Z).Bounds, Level) Then
            If PtInRect(Menus(m).Items(Z).Bounds, X, Y) And Menus(m).Items(Z).Visible Then tItem = Z
            If Level > 1 Or (Level = 1 And m_Orientation = Vertical) Then
                OldBounds = Menus(m).Items(Z).Bounds
                Menus(m).Items(Z).Bounds.Left = 0
                If Level = 1 Then Menus(m).Items(Z).Bounds.Right = UserControl.ScaleWidth Else Menus(m).Items(Z).Bounds.Right = MenuBox(Level - 1).ScaleWidth
                If PtInRect(Menus(m).Items(Z).Bounds, X, Y) And Menus(m).Items(Z).Visible Then tItem = Z
                Menus(m).Items(Z).Bounds = OldBounds
            End If
        End If
        Z = Z + 1
    Loop
End If
GetItemFromCursor = tItem
End Function

Private Function PopUpChild(ByVal X As Long, ByVal Y As Long, ByVal Level As Integer, ByVal MenuNumber As Integer, Optional ByVal ShowFromWh As AlignConstants) As Boolean
Dim tXS As Long, tYS As Long, Offset As Long, i As Long, R As RECT, R2 As RECT, tempP As POINTAPI
Dim Z As Long, TempWidth As Long, TempHeight As Long, TempPW As Long, TempPH As Long
Dim T As String

UpperMenu = Level
ReleaseMouseCapture
TopItem(Level) = -1
i = Level - 1
If MenuBox.UBound < (i) Then
    For Z = MenuBox.UBound + 1 To i
        Load MenuBox(Z)
        'SetWindowLong MenuBox(Z).hWnd, GWL_STYLE, WS_POPUP
        SetWindowLong MenuBox(Z).hWnd, GWL_EXSTYLE, WS_EX_TOPMOST Or &H80
        SetWindowPos MenuBox(Z).hWnd, -1, 0, 0, 0, 0, 3
        SetParent MenuBox(Z).hWnd, 0
        'SetWindowLong MenuBox(Z).hWnd, GWL_STYLE, 0 'GetWindowLong(MenuBox(Z).hWnd, GWL_STYLE) And Not WS_CHILD
        'SetWindowLong MenuBox(Z).hWnd, GWL_EXSTYLE, GetWindowLong(MenuBox(Z).hWnd, GWL_EXSTYLE) Or WS_EX_APPWINDOW
    Next
End If

Set MenuBox(i).Font = UserControl.Font
Set MenuBox(i).Picture = UserControl.Picture
MenuBox(i).ForeColor = UserControl.ForeColor
MenuBox(i).BackColor = UserControl.BackColor
SetMouseCapture MenuBox(i).hWnd

If linEast.UBound < Level Then
    For Z = linEast.UBound + 1 To Level
        Load linEast(Z)
        Set linEast(Z).Container = MenuBox(i)
        Load linWest(Z)
        Set linWest(Z).Container = MenuBox(i)
        Load linSouth(Z)
        Set linSouth(Z).Container = MenuBox(i)
        Load linNorth(Z)
        Set linNorth(Z).Container = MenuBox(i)
    Next
End If

Do While MenuNumber <= TotalItems
    If Menus(m).Items(MenuNumber).SubLevel < Level Or (Menus(m).Items(MenuNumber).PopUp = True And IsAPopup = 0) Then Exit Do
    If Menus(m).Items(MenuNumber).SubLevel = Level And Menus(m).Items(MenuNumber).Visible Then
        If TopItem(Level) = -1 Then TopItem(Level) = MenuNumber
        Offset = Offset + MenuVerticSpace
        
        TempWidth = 0
        TempHeight = 0
        'MenuBox(I).FontBold = Menus(M).Items(MenuNumber).Checked
        'MenuBox(I).FontUnderline = MenuBox(I).FontBold
        Menus(m).Items(MenuNumber).Bounds.Top = Offset
        If Menus(m).Items(MenuNumber).Style <> TextOnly Then
            TempWidth = GetPicWidth(Menus(m).Items(MenuNumber).Icon)
            TempHeight = GetPicHeight(Menus(m).Items(MenuNumber).Icon)
        End If
        If Menus(m).Items(MenuNumber).Style <> GraphicsOnly Then
            If TempWidth <> 0 Then TempWidth = TempWidth + IconMargin
            TempWidth = TempWidth + MenuBox(i).TextWidth(Menus(m).Items(MenuNumber).Caption)
            TempHeight = Maximum(MenuBox(i).TextHeight(Menus(m).Items(MenuNumber).Caption), TempHeight)
        End If
        Offset = Offset + TempHeight
        Menus(m).Items(MenuNumber).Bounds.Bottom = Offset
        Menus(m).Items(MenuNumber).Bounds.Left = MenuPopupHorizIndent
        Menus(m).Items(MenuNumber).Bounds.Right = MenuPopupHorizIndent + TempWidth
        MenuCount(Level) = MenuCount(Level) + 1
    End If
    MenuNumber = MenuNumber + 1
Loop

tYS = Offset + 4
tXS = GetProperSize(Level, vbAlignLeft)
Const SC = 2

ShowFromWhere = vbAlignTop
If IsAPopup = 0 Then
    ShowFromWhere = vbAlignLeft
    If Level = 2 Then
        ShowFromWhere = UserControl.Extender.Align
        If UserControl.Extender.Align = vbAlignRight Then
            X = DontCloseRect(Level - 1).Left - tXS
        End If
        If UserControl.Extender.Align = vbAlignBottom Then
            Y = DontCloseRect(Level - 1).Top - tYS
        End If
    End If
Else
    ShowFromWhere = ShowFromWh
    Select Case ShowFromWhere
        Case vbAlignTop, vbAlignNone
            X = X - SC
            Y = Y - SC
        Case vbAlignBottom
            X = X - SC
            Y = Y + SC
        Case vbAlignLeft
            X = X - SC
            Y = Y - SC
        Case vbAlignRight
            X = X + SC
            Y = Y - SC
    End Select
End If

If (X + tXS) > GetSystemMetrics(SM_CXSCREEN) Then
    If IsAPopup Then ShowFromWhere = vbAlignRight
    If (X - tXS) < 0 Then
        X = 0
    Else
        X = X - tXS + 2 * SC
    End If
End If
If (Y + tYS) > GetSystemMetrics(SM_CYSCREEN) Then
    If IsAPopup Then ShowFromWhere = vbAlignBottom
    If (Y - tYS) * Screen.TwipsPerPixelY < 0 Then
        Y = 0
    Else
        Y = Y - tYS + 2 * SC
    End If
End If

MenuBox(i).Move X * Screen.TwipsPerPixelX, Y * Screen.TwipsPerPixelY, tXS * Screen.TwipsPerPixelX, tYS * Screen.TwipsPerPixelY
PaintMenuArea Level
MenuBox(i).Font.Charset = DefaultFontCharset

If IsAPopup <> 0 Then
    GetWindowRect MenuBox(i).hWnd, R
    R.Bottom = R.Bottom + 1
    R.Right = R.Right + 1
    DontCloseRect(IsAPopup - 1) = R
End If

MenuNumber = TopItem(Level)
Do While MenuNumber <= TotalItems
    If Menus(m).Items(MenuNumber).SubLevel < Level Or (Menus(m).Items(MenuNumber).PopUp = True And IsAPopup = 0) Then Exit Do
    If Menus(m).Items(MenuNumber).SubLevel = Level And Menus(m).Items(MenuNumber).Visible Then
        tempP.X = Menus(m).Items(MenuNumber).Bounds.Left
        tempP.Y = Menus(m).Items(MenuNumber).Bounds.Top
        
        If Menus(m).Items(MenuNumber).BeginAGroup Then
            R2.Right = MenuBox(i).ScaleWidth
            R2.Top = tempP.Y - GroupMargin
            R2.Bottom = R2.Top
            DrawEdge MenuBox(i).hDC, R2, EDGE_ETCHED, BF_TOP
        End If
        
        With Menus(m).Items(MenuNumber)
            If .Checked Then
                MenuBox(i).DrawMode = MenuCheckDrawMode
                MenuBox(i).Line (.Bounds.Left, .Bounds.Top)-(.Bounds.Right, .Bounds.Bottom), colMenuCheck, BF
                MenuBox(i).DrawMode = 13
            End If
        End With
        
        If Menus(m).Items(MenuNumber).Style <> TextOnly Then
            TempPW = GetPicWidth(Menus(m).Items(MenuNumber).Icon)
            TempPH = GetPicHeight(Menus(m).Items(MenuNumber).Icon)
            
            'MenuBox(i).PaintPicture Menus(M).Items(MenuNumber).Icon, tempP.X, tempP.Y + ((Menus(M).Items(MenuNumber).Bounds.Bottom - Menus(M).Items(MenuNumber).Bounds.Top) - TempPH) / 2, TempPW, TempPH
            TransparentBlt MenuBox(i).hDC, Menus(m).Items(MenuNumber).Icon, tempP.X, tempP.Y + ((Menus(m).Items(MenuNumber).Bounds.Bottom - Menus(m).Items(MenuNumber).Bounds.Top) - TempPH) / 2, colIconTransparent
            
            tempP.X = tempP.X + IconMargin + TempPW
        End If
        'MenuBox(I).FontBold = Menus(M).Items(MenuNumber).Checked
        'MenuBox(I).FontUnderline = MenuBox(I).FontBold

        If Menus(m).Items(MenuNumber).Style <> GraphicsOnly Then
            MenuBox(i).CurrentX = tempP.X
            MenuBox(i).CurrentY = tempP.Y + ((Menus(m).Items(MenuNumber).Bounds.Bottom - Menus(m).Items(MenuNumber).Bounds.Top) - MenuBox(i).TextHeight(Menus(m).Items(MenuNumber).Caption)) / 2
            MenuBox(i).Print Menus(m).Items(MenuNumber).Caption
        End If
        If IsSubMenu(MenuNumber) Then
            T = MenuBox(i).Font.Name
            MenuBox(i).Font.Name = "Marlett"
            MenuBox(i).CurrentX = MenuBox(i).ScaleWidth - 2 * MenuBox(i).TextWidth(MenuRightArrowSymbol)
            MenuBox(i).CurrentY = Menus(m).Items(MenuNumber).Bounds.Top + (Menus(m).Items(MenuNumber).Bounds.Bottom - Menus(m).Items(MenuNumber).Bounds.Top) \ 2 - MenuBox(i).TextHeight(MenuRightArrowSymbol) \ 2
            MenuBox(i).Print MenuRightArrowSymbol
            MenuBox(i).Font.Name = T$
        End If
        MenuCount(Level) = MenuCount(Level) + 1
    End If
    MenuNumber = MenuNumber + 1
Loop

SelectItem Menus(m).Items(TopItem(Level)).Bounds, 0, Level

ShowSubMenu Level
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Refresh
Public Sub Refresh()
On Local Error GoTo EH:

ItemOffset = 0
MaxOffset = 0
CanScrollLeft = False
CanScrollRight = False
AdjustSize
PaintMenuArea
DrawDragArea
AdjustItems
DrawItems

If MenuCount(1) > 0 And SelItem(1) <> 0 Then
    SelectItem Menus(m).Items(SelItem(1)).Bounds, EDGE_RAISED
End If

UserControl.Refresh
Exit Sub

EH:
Resume Next
End Sub

Public Sub UpdateMenu()
FillMenu
End Sub

Private Sub SelectSystemFont()
Dim NCM As NONCLIENTMETRICS, Z As Long, S As String

NCM.cbSize = Len(NCM)
SystemParametersInfo SPI_GETNONCLIENTMETRICS, Len(NCM), NCM, 0
For Z = 0 To UBound(NCM.lfMenuFont.lfFaceName)
    S = S & Chr(NCM.lfMenuFont.lfFaceName(Z))
    If Right(S, 1) = vbNullChar Then
        S = Left(S, Len(S) - 1)
        Exit For
    End If
Next
UserControl.Font.Name = S 'Left(NCM.lfMenuFont.lfFaceName, InStr(NCM.lfMenuFont.lfFaceName, Chr$(0)) - 1)
UserControl.Font.Size = Abs(NCM.lfMenuFont.lfHeight) * 72 / GetDeviceCaps(hDC, LOGPIXELSY)
UserControl.Font.Weight = NCM.lfMenuFont.lfWeight
UserControl.Font.Italic = NCM.lfMenuFont.lfItalic > 0
UserControl.Font.Strikethrough = NCM.lfMenuFont.lfStrikeOut > 0
UserControl.Font.Underline = NCM.lfMenuFont.lfUnderline > 0

End Sub

Public Function IsSubMenu(ByVal MenuNumber As Integer) As Boolean
If MenuNumber < 1 Or MenuNumber >= TotalItems Then Exit Function
If Menus(m).Items(MenuNumber + 1).SubLevel > Menus(m).Items(MenuNumber).SubLevel Then
    If Menus(m).Items(MenuNumber).PopUp = Menus(m).Items(MenuNumber + 1).PopUp Then IsSubMenu = True
End If
End Function

Private Sub PaintMenuArea(Optional ByVal Level As Integer = 1)
Dim SCol As Long, DCol As Long, DC As Long
Dim X As Long, Y As Long
Dim tPX As Long, tPY As Long
Dim qrc As RECT, tSW As Long, tSH As Long

If UserControl.AutoRedraw = False Then UserControl.AutoRedraw = True

If Level = 1 Then
    tSW = UserControl.ScaleWidth - 1
    tSH = UserControl.ScaleHeight - 1
Else
    tSW = MenuBox(Level - 1).ScaleWidth - 1
    tSH = MenuBox(Level - 1).ScaleHeight - 1
End If

If UserControl.Picture = LoadPicture Then
    If m_Gradiented Then
        If Not m_GradientInverse Then
            SCol = m_GradientColor
            DCol = UserControl.BackColor
        Else
            SCol = UserControl.BackColor
            DCol = m_GradientColor
        End If
        If Level = 1 Then DC = UserControl.hDC Else DC = MenuBox(Level - 1).hDC
        Gradient DC, SCol, DCol, 0, 0, tSW + 1, tSH + 1, m_GradientVertical
    Else
        If Level = 1 Then UserControl.Cls Else MenuBox(Level - 1).Cls
    End If
Else
    Select Case m_PictureState
        Case psNormal
            If Level = 1 Then UserControl.Cls Else MenuBox(Level - 1).Cls
        Case psStretched
            If Level = 1 Then UserControl.PaintPicture UserControl.Picture, 0, 0, tSW + 1, tSH + 1 Else MenuBox(Level - 1).PaintPicture UserControl.Picture, 0, 0, tSW + 1, tSH + 1
        Case psTiled
            tPX = ScaleX(UserControl.Picture.Width, vbHimetric, vbPixels)
            tPY = ScaleY(UserControl.Picture.Height, vbHimetric, vbPixels)
            For X = 0 To Int((tSW + 1) / tPX) + 1
                For Y = 0 To Int((tSH + 1) / tPY) + 1
                    If Level = 1 Then
                        UserControl.PaintPicture UserControl.Picture, X * tPX, Y * tPY
                    Else
                        MenuBox(Level - 1).PaintPicture UserControl.Picture, X * tPX, Y * tPY
                    End If
                Next
            Next
    End Select
End If

Select Case m_BorderStyle
    Case mbs2D
        If Level = 1 Then
            UserControl.Line (0, 0)-(tSW, tSH), 0, B
        Else
            MenuBox(Level - 1).Line (0, 0)-(tSW, tSH), 0, B
        End If
    Case mbs3DThin
        qrc.Right = tSW
        qrc.Bottom = tSH
        DrawEdgeThin qrc, EDGE_RAISED, Level
    Case mbs3DThick
        qrc.Right = tSW + 1
        qrc.Bottom = tSH + 1
        If Level = 1 Then DC = UserControl.hDC Else DC = MenuBox(Level - 1).hDC
        DrawEdge DC, qrc, EDGE_RAISED, BF_RECT
End Select

End Sub

Private Sub ShowSubMenu(ByVal Level As Long, Optional ByVal bShow As Boolean = True)
Const MenuAnim = 1
Const StepsConst = 100
Const TimeToSlide = 150
Dim DestR As RECT
Dim Z As Long
Dim P As POINTAPI, P2 As POINTAPI
Dim SourceR As RECT, T As Double, Steps As Single

If bShow Then
    If MenuSoundEnabled Then sndPlaySound "MenuPopup", SND_ASYNC Or SND_NODEFAULT
Select Case MenuAnim
Case 1
    SourceR.Right = MenuBox(Level - 1).ScaleWidth
    SourceR.Bottom = MenuBox(Level - 1).ScaleHeight
    ClientToScreen MenuBox(Level - 1).hWnd, P
    SourceR.Left = P.X
    SourceR.Top = P.Y
    DestR = SourceR
    Steps = StepsConst
    'If ShowFromWhere < 2.5 Then Steps = DestR.Bottom Else Steps = DestR.Right
    For Z = 1 To Steps
        Select Case ShowFromWhere
            Case vbAlignBottom
                DestR.Top = SourceR.Top + SourceR.Bottom * (1 - Z / Steps)
                DestR.Bottom = SourceR.Bottom - (DestR.Top - SourceR.Top)
                P2.X = 0: P2.Y = 0
            Case vbAlignLeft
                DestR.Right = SourceR.Right * Z / Steps
                P2.X = SourceR.Right - DestR.Right
                P2.Y = SourceR.Bottom - DestR.Bottom
            Case vbAlignRight
                DestR.Left = SourceR.Left + SourceR.Right * (1 - Z / Steps)
                DestR.Right = SourceR.Right - (DestR.Left - SourceR.Left)
                P2.X = 0: P2.Y = 0
            Case Else
                DestR.Bottom = SourceR.Bottom * Z / Steps
                P2.X = SourceR.Right - DestR.Right
                P2.Y = SourceR.Bottom - DestR.Bottom
        End Select
        BitBlt GetWindowDC(GetDesktopWindow), DestR.Left, DestR.Top, DestR.Right, DestR.Bottom, MenuBox(Level - 1).hDC, P2.X, P2.Y, vbSrcCopy
        T = timeGetTime: Do: Loop Until timeGetTime >= T + TimeToSlide / StepsConst
    Next
End Select
End If

MenuBox(Level - 1).Visible = bShow
End Sub

Private Sub TerminateMenu(ByVal Level As Long)
Dim lPA As POINTAPI, hW As Long
If Level > UpperMenu Or Level < 2 Then Exit Sub
ReleaseMouseCapture
SelItem(Level) = 0
ShowSubMenu Level, False
'SetParent MenuBox(Level - 1).hWnd, UserControl.hWnd
'SetWindowLong MenuBox(Level - 1).hWnd, GWL_STYLE, GetWindowLong(MenuBox(Level - 1).hWnd, GWL_STYLE) Or WS_CHILD And Not WS_POPUP
'If Level >= 2 Then Unload MenuBox(Level - 1)
'If Level >= 2 Then SetupTerminationTimer Level - 1
UpperMenu = Level - 1

If UpperMenu = 1 Then hW = UserControl.hWnd Else hW = MenuBox(UpperMenu - 1).hWnd
GetCursorPos lPA

If WindowFromPoint(lPA.X, lPA.Y) = hW Then
    SelectItem Menus(m).Items(TopItem(UpperMenu + 1) - 1).Bounds, 0, UpperMenu
    SelItem(UpperMenu) = 0
    SetMouseCapture hW
Else
    If PtInRect(DontCloseRect(UpperMenu), lPA.X, lPA.Y) Then
        SetMouseCapture hW
    Else
        If UpperMenu > 1 Then TerminateMenu UpperMenu
        SelectItem Menus(m).Items(TopItem(UpperMenu + 1) - 1).Bounds, 0, UpperMenu
        SelItem(UpperMenu) = 0
    End If
End If

If IsAPopup = Level Then IsAPopup = 0

'If PtInRect(DontCloseRect(UpperMenu), lPA.X, lPA.Y) Then
'    If UpperMenu = 1 Then hW = UserControl.hWnd Else hW = MenuBox(UpperMenu - 1).hWnd
'    SetMouseCapture hW
'End If
'SelectItem Menus(M).Items(TopItem(UpperMenu + 1) - 1).Bounds, 0, UpperMenu
'SelItem(UpperMenu) = 0
'
'If WindowFromPoint(lPA.X, lPA.Y) <> hW And UpperMenu > 1 Then
'    TerminateMenu UpperMenu
'End If
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseDown(Button, Shift, X, Y)
Dim Over As Long, lPA As POINTAPI
Dim OldForeCol As Long

'If PtInRect(ScrollAreaLeft, X, Y) And Button = 1 And CanScrollLeft Then
'    ScrollLeft
'    Exit Sub
'End If
'
'If PtInRect(ScrollAreaRight, X, Y) And Button = 1 And CanScrollRight Then
'    ScrollRight
'    Exit Sub
'End If

If PtInRect(DragArea, X, Y) And Button = 1 And DragAreaVisible Then
    DraggingMe = True
    TrackCapture.Enabled = False
    OX = X
    OY = Y
    UserControl.AutoRedraw = False
    Rt.Left = UserControl.Extender.Left
    Rt.Top = UserControl.Extender.Top
    Rt.Right = UserControl.ScaleWidth + Rt.Left
    Rt.Bottom = UserControl.ScaleHeight + Rt.Top
    ClientToScreen UserControl.ContainerHwnd, lPA
    OffsetRect Rt, lPA.X, lPA.Y
    DrawFocusRect GetWindowDC(GetDesktopWindow), Rt
    OldAlign = UserControl.Extender.Align
    RaiseEvent DragBar(BeginDrag, OldAlign)
    Exit Sub
End If

If GetCapture <> CaptureWindow And CaptureWindow <> 0 Then SetMouseCapture CaptureWindow

Over = GetItemFromCursor(1, X, Y)
If Over = 0 Then Exit Sub
If IsSubMenu(Over) Then Exit Sub

If Button = 2 And Shift = 0 Then
    RaiseEvent MustHideTooltip
    mnuWhatsThis.Caption = GetString(ResWhatsThis)
    If SelItem(1) <> 0 Then SelectItem Menus(m).Items(SelItem(1)).Bounds, 0
    CommandHoverSent = 0
    SelItem(1) = 0
    ReleaseMouseCapture
    WhatsThisCandidate = Over
    If m_ToolbarMode = mbsToolBar And m_WhatsThisHelp Then PopupMenu mnuWhatsThisPopup
    Exit Sub
End If

If Button = 1 And Shift = 0 Then
    
    If MenuSoundEnabled Then sndPlaySound "MenuCommand", SND_ASYNC Or SND_NODEFAULT
    SelectItem Menus(m).Items(Over).Bounds, EDGE_SUNKEN, 1
    
    If Menus(m).Items(Over).Style <> GraphicsOnly Then
        OldForeCol = EnsureRGB(UserControl.ForeColor)
        UserControl.ForeColor = RGB(Abs(Red(OldForeCol) - 2 * ColDecrease), Abs(Green(OldForeCol) - 2 * ColDecrease), Abs(Blue(OldForeCol) - 2 * ColDecrease))
        'UserControl.FontBold = Menus(M).Items(Over).Checked
        'UserControl.FontUnderline = Menus(M).Items(Over).Checked
        UserControl.CurrentX = Menus(m).Items(Over).Bounds.Left
        If Menus(m).Items(Over).Style = TextAndGraphics Then
            UserControl.CurrentX = Menus(m).Items(Over).Bounds.Left + IconMargin + GetPicWidth(Menus(m).Items(Over).Icon)
        End If
        UserControl.CurrentY = Menus(m).Items(Over).Bounds.Top + ((Menus(m).Items(Over).Bounds.Bottom - Menus(m).Items(Over).Bounds.Top) - GetItemHeight(Over)) / 2
        UserControl.Print Menus(m).Items(Over).Caption
    
        UserControl.ForeColor = OldForeCol
    End If
    
    SentCommand = Over
    WasMouseDown = 1
End If

End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseMove(Button, Shift, X, Y)

If UpperMenu > 1 Then Exit Sub

Dim Over As Integer, qrc As RECT
Dim NewAlign As Integer, PX As Long, PY As Long
Dim PSW As Long, PSH As Long
Dim lPA As POINTAPI

Over = GetItemFromCursor(1, X, Y)

If Over <> 0 And Button = 0 And Not DraggingMe Then
    If SelItem(1) <> Over Then
        If SelItem(1) <> 0 Then SelectItem Menus(m).Items(SelItem(1)).Bounds, 0: SelItem(1) = 0
        If CommandHoverSent <> Over Then
'
'            '=====================
'            ' Wait to see if the user really
'            ' wants something from this
'            ' point on the screen...
'            Const TooltipDelay = 400 ' wait 400 milliseconds
'            ' to check user's confidence in his choice
'            Dim P As POINTAPI, NewP As POINTAPI
'            Dim T As Long
'
'            GetCursorPos P
'            GetCursorPos NewP
'            T = timeGetTime
'            Do
'                DoEvents
'                GetCursorPos NewP
'            Loop Until timeGetTime - T > TooltipDelay Or NewP.X <> P.X Or NewP.Y <> P.Y
'
'            If NewP.X = P.X And NewP.Y = P.Y Then
                CommandHoverSent = Over
                RaiseEvent CommandHover(Over, X, Y, Shift)
'            End If
        End If
        If IsSubMenu(Over) Then
            If AutoShowSubMenus Then
                DontCloseRect(1) = Menus(m).Items(Over).Bounds
                If m_Orientation = Horizontal Then
                    DontCloseRect(1).Top = 0
                    DontCloseRect(1).Bottom = UserControl.ScaleHeight + MenuSmallOffSet + 1
                Else
                    DontCloseRect(1).Left = 0
                    DontCloseRect(1).Right = UserControl.ScaleWidth + MenuSmallOffSet + 1
                End If
                ClientToScreen UserControl.hWnd, lPA
                OffsetRect DontCloseRect(1), lPA.X, lPA.Y
                Select Case UserControl.Extender.Align
                    Case vbAlignLeft
                        PX = DontCloseRect(1).Right - MenuSmallOffSet - 1
                        PY = DontCloseRect(1).Top - SelMargin
                    Case vbAlignTop, 0
                        PX = DontCloseRect(1).Left - SelMargin
                        PY = DontCloseRect(1).Bottom - MenuSmallOffSet - 1
                    Case vbAlignRight
                        PX = DontCloseRect(1).Left + MenuSmallOffSet + 1
                        PY = DontCloseRect(1).Top - SelMargin
                    Case vbAlignBottom
                        PX = DontCloseRect(1).Left - SelMargin
                        PY = DontCloseRect(1).Top + MenuSmallOffSet + 1
                End Select
                SelectItem Menus(m).Items(Over).Bounds, EDGE_SUNKEN
                linEast(1).Refresh
                linNorth(1).Refresh
                linSouth(1).Refresh
                linWest(1).Refresh
                SelItem(1) = Over
                PopUpChild PX, PY, 2, Over + 1
            Else
                'don't show it automatically
            End If
        Else
            SetMouseCapture UserControl.hWnd
            SelectItem Menus(m).Items(Over).Bounds, EDGE_RAISED
            SelItem(1) = Over
        End If
    End If
Else
    If WasMouseDown = 0 Then
        If SelItem(1) <> 0 Then
            SelectItem Menus(m).Items(SelItem(1)).Bounds, 0
            RaiseEvent CommandHover(-1, 0, 0, 0)
            CommandHoverSent = 0
            SelItem(1) = 0
            ReleaseMouseCapture
        End If
    End If
End If

If (X < 0 Or Y < 0 Or X >= UserControl.ScaleWidth Or Y >= UserControl.ScaleHeight) And Button = 0 And WasMouseDown = 0 And Not DraggingMe Then
    If CaptureWindow <> 0 Then
        If SelItem(1) <> 0 Then
            SelectItem Menus(m).Items(SelItem(1)).Bounds, 0
            SelItem(1) = 0
        End If
        ReleaseMouseCapture
    End If
End If

If Button = 1 And DraggingMe Then
    DrawFocusRect GetWindowDC(GetDesktopWindow), Rt
    
    PSW = UserControl.Extender.Parent.ScaleWidth
    PSH = UserControl.Extender.Parent.ScaleHeight
    PX = ((X + UserControl.Extender.Left) - PSW \ 2) * PSH \ PSW
    PY = (Y + UserControl.Extender.Top) - PSH \ 2

    NewAlign = OldAlign
    If Abs(PY) > Abs(PX) Then
        If PY > PX Then
            NewAlign = vbAlignBottom
        Else
            NewAlign = vbAlignTop
        End If
    Else
        If PX > PY Then
            NewAlign = vbAlignRight
        Else
            NewAlign = vbAlignLeft
        End If
    End If

    If (NewAlign > 2.5) <> (OldAlign > 2.5) Then
        If NewAlign > 2.5 Then
            Rt.Right = Rt.Left + GetProperSize(1, NewAlign)
            Rt.Bottom = Rt.Top + PSH
        Else
            Rt.Right = Rt.Left + PSW
            Rt.Bottom = Rt.Top + GetProperSize(1, NewAlign)
        End If
    End If

    Rt.Left = Rt.Left + (X - OX)
    Rt.Top = Rt.Top + (Y - OY)
    Rt.Right = Rt.Right + (X - OX)
    Rt.Bottom = Rt.Bottom + (Y - OY)
    OX = X
    OY = Y
    OldAlign = NewAlign

    DrawFocusRect GetWindowDC(GetDesktopWindow), Rt
    RaiseEvent DragBar(Dragged, OldAlign)
End If

End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseUp(Button, Shift, X, Y)

Dim NewAlign As Integer, PX As Long, PY As Long
Dim PSW As Long, PSH As Long, Over As Long

If DraggingMe Then
    PSW = UserControl.Extender.Parent.ScaleWidth
    PSH = UserControl.Extender.Parent.ScaleHeight
    PX = ((X + UserControl.Extender.Left) - PSW \ 2) * PSH \ PSW
    PY = (Y + UserControl.Extender.Top) - PSH \ 2

    NewAlign = UserControl.Extender.Align
    If Abs(PY) > Abs(PX) Then
        If PY > PX Then
            NewAlign = vbAlignBottom
        Else
            NewAlign = vbAlignTop
        End If
    Else
        If PX > PY Then
            NewAlign = vbAlignRight
        Else
            NewAlign = vbAlignLeft
        End If
    End If

    DrawFocusRect GetWindowDC(GetDesktopWindow), Rt
    UserControl.AutoRedraw = True

    If NewAlign = UserControl.Extender.Align Then Exit Sub
    
    RaiseEvent DragBar(AboutToPlace, NewAlign)
    
    UserControl.Extender.Align = vbAlignNone
    If NewAlign > 2.5 Then
        m_Orientation = Vertical
        UserControl.Width = ScaleX(GetProperSize(1, NewAlign), vbPixels, vbTwips)
    Else
        m_Orientation = Horizontal
        UserControl.Height = ScaleY(GetProperSize(1, NewAlign), vbPixels, vbTwips)
    End If
    UserControl.Extender.Align = NewAlign
    
    DraggingMe = False
    
    RaiseEvent DragBar(EndDrag, NewAlign)
    Exit Sub
End If
'_____________END DRAGGING ME

If WasMouseDown = 1 Then
    If SentCommand <> 0 Then Over = SentCommand Else Over = GetItemFromCursor(1, X, Y)
    If Over = SentCommand Then
        SelectItem Menus(m).Items(SentCommand).Bounds, 0 'EDGE_RAISED
        SelItem(1) = 0 'SentCommand

        If Menus(m).Items(Over).Style <> GraphicsOnly Then
            'UserControl.FontBold = Menus(M).Items(Over).Checked
            'UserControl.FontUnderline = Menus(M).Items(Over).Checked
            UserControl.CurrentX = Menus(m).Items(Over).Bounds.Left
            If Menus(m).Items(Over).Style = TextAndGraphics Then
                UserControl.CurrentX = Menus(m).Items(Over).Bounds.Left + IconMargin + GetPicWidth(Menus(m).Items(Over).Icon)
            End If
            UserControl.CurrentY = Menus(m).Items(Over).Bounds.Top + ((Menus(m).Items(Over).Bounds.Bottom - Menus(m).Items(Over).Bounds.Top) - GetItemHeight(Over)) / 2
            UserControl.Print Menus(m).Items(Over).Caption
        End If
        SetMouseCapture UserControl.hWnd
        WasMouseDown = 0
        SentCommand = 0
        DoEvents
        RaiseEvent Command(Over, Menus(m).Items(Over).Caption)
        DoEvents
    Else
        SelectItem Menus(m).Items(SentCommand).Bounds, 0
        If Menus(m).Items(SentCommand).Style <> GraphicsOnly Then
            'UserControl.FontBold = Menus(M).Items(SentCommand).Checked
            'UserControl.FontUnderline = Menus(M).Items(SentCommand).Checked
            UserControl.CurrentX = Menus(m).Items(SentCommand).Bounds.Left
            If Menus(m).Items(SentCommand).Style = TextAndGraphics Then
                UserControl.CurrentX = Menus(m).Items(SentCommand).Bounds.Left + IconMargin + GetPicWidth(Menus(m).Items(SentCommand).Icon)
            End If
            UserControl.CurrentY = Menus(m).Items(SentCommand).Bounds.Top + ((Menus(m).Items(SentCommand).Bounds.Bottom - Menus(m).Items(SentCommand).Bounds.Top) - GetItemHeight(SentCommand)) / 2
            UserControl.Print Menus(m).Items(SentCommand).Caption
        End If
        SelItem(1) = 0
        WasMouseDown = 0
        SentCommand = 0
    End If
End If

End Sub

'Private Sub SetupTerminationTimer(ByVal Level As Integer)
'SelfDestruct.Tag = Level
'SelfDestruct.Enabled = True
'End Sub

Private Function GetItemWidth(ByVal ItemNum As Integer, Optional ByVal iOrientation As OrientationStyle = Horizontal)
Dim W As Long

If Menus(m).Items(ItemNum).SubLevel = 1 Then
    W = UserControl.TextWidth(Menus(m).Items(ItemNum).Caption)
Else
    W = MenuBox(Menus(m).Items(ItemNum).SubLevel - 1).TextWidth(Menus(m).Items(ItemNum).Caption)
End If
'If iOrientation = Horizontal Then
'    W = W + MenuHorizSpace
'Else
    'dummy
'End If
GetItemWidth = W
End Function

Private Function GetItemHeight(ByVal ItemNum As Integer, Optional ByVal iOrientation As OrientationStyle = Horizontal)
Dim W As Long, Level As Long

W = UserControl.TextHeight(Menus(m).Items(ItemNum).Caption)

'If iOrientation = Horizontal Then
    'dummy
'Else
'    W = W + MenuVerticSpace
'End If
GetItemHeight = W
End Function

Private Sub SetMouseCapture(ByVal hWnd As Long)
If GetCapture <> CaptureWindow Or CaptureWindow <> 0 Then
    ReleaseCapture
End If
SetCapture hWnd
CaptureWindow = hWnd
TrackCapture.Enabled = True
End Sub

Private Sub ReleaseMouseCapture()
If CaptureWindow = 0 Then Exit Sub
TrackCapture.Enabled = False
ReleaseCapture
CaptureWindow = 0
If UpperMenu = 1 And SelItem(1) <> 0 Then SelectItem Menus(m).Items(SelItem(1)).Bounds, 0, 1
End Sub

Private Sub Gradient(ByVal DC As Long, ByVal SCol As Long, ByVal DCol As Long, ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long, Optional ByVal Vertical As Boolean = True)
On Error GoTo EH:
If Height < 1 Or Width < 1 Then Exit Sub
Dim hBrush As Long
Dim hOldBrush As Long
Dim hPen As Long
Dim hOldPen As Long
Dim DrawColorR(1 To 2) As Integer
Dim DrawColorG(1 To 2) As Integer
Dim DrawColorB(1 To 2) As Integer
Dim St As Integer, Coeff As Double, i As Long

If SCol < 0 Then SCol = GetSysColor(SCol + SysColorTranslationBase)
If DCol < 0 Then DCol = GetSysColor(DCol + SysColorTranslationBase)

DrawColorR(1) = Red(SCol)
DrawColorG(1) = Green(SCol)
DrawColorB(1) = Blue(SCol)
DrawColorR(2) = Red(DCol) - DrawColorR(1)
DrawColorG(2) = Green(DCol) - DrawColorG(1)
DrawColorB(2) = Blue(DCol) - DrawColorB(1)

hPen = CreatePen(5, 0, 0)
hOldPen = SelectObject(DC, hPen)

If Vertical Then
    St = Width \ 128
    If St < 1 Then St = 1
    St = St + 1
    Height = Top + Height + 1
    For i = Left To Left + Width Step St - 1
        Coeff = (i - Left) / Width
    
        hBrush = CreateSolidBrush(RGB(DrawColorR(1) + DrawColorR(2) * Coeff, DrawColorG(1) + DrawColorG(2) * Coeff, DrawColorB(1) + DrawColorB(2) * Coeff))
        hOldBrush = SelectObject(DC, hBrush)
        Rectangle DC, i, Top, i + St, Height
        SelectObject DC, hOldBrush
        DeleteObject hBrush
    
    Next i
Else
    St = Height \ 128
    If St < 1 Then St = 1
    St = St + 1
    Width = Left + Width + 1
    For i = Top To Top + Height Step St - 1
        Coeff = (i - Top) / Height
    
        hBrush = CreateSolidBrush(RGB(DrawColorR(1) + DrawColorR(2) * Coeff, DrawColorG(1) + DrawColorG(2) * Coeff, DrawColorB(1) + DrawColorB(2) * Coeff))
        hOldBrush = SelectObject(DC, hBrush)
        Rectangle DC, Left, i, Width, i + St
        SelectObject DC, hOldBrush
        DeleteObject hBrush
    
    Next i
End If

SelectObject DC, hOldPen
DeleteObject hPen
EH:
End Sub

Private Function Red(ByVal RGB As Long) As Long
Red = RGB And 255
End Function

Private Function Green(ByVal RGB As Long) As Long
Green = (RGB And 65535) \ 256
End Function

Private Function Blue(ByVal RGB As Long) As Long
Blue = RGB \ 65536
End Function

Public Sub AddItem(ByVal Caption As String, ByVal SubLevel As Integer, Optional ByVal ToolTipText As String, Optional ByVal Index As Integer = -1, Optional ByVal PopUp As Boolean = False, Optional Icon As StdPicture, Optional ByVal Style As MenuItemStyle = TextOnly, Optional ByVal BeginAGroup As Boolean = False, Optional ByVal Visible As Boolean = True, Optional AuxIndex, Optional Enabled As Boolean = True, Optional ByVal ShortcutKey As Long = 0)
Dim Z As Long
On Local Error Resume Next

Dim UI As Long
UI = UBound(Menus(m).Items) + 1
If MenuCount(1) = 0 Then UI = 1
ReDim Preserve Menus(m).Items(1 To UI)
If Index > 0 And Index < UI - 1 Then
    For Z = UI To Index + 1 Step -1
        Menus(m).Items(Z) = Menus(m).Items(Z - 1)
    Next
Else
    Index = UI
End If
Menus(m).Items(Index).Caption = Caption
Menus(m).Items(Index).SubLevel = SubLevel
Menus(m).Items(Index).ToolTipText = ToolTipText
Menus(m).Items(Index).Enabled = Enabled
Menus(m).Items(Index).Visible = Visible
Menus(m).Items(Index).PopUp = PopUp
Menus(m).Items(Index).BeginAGroup = BeginAGroup
Menus(m).Items(Index).ShortcutKey = ShortcutKey
If Not IsMissing(AuxIndex) Then Menus(m).Items(Index).AuxIndex = AuxIndex
Set Menus(m).Items(Index).Icon = Icon
Menus(m).Items(Index).Style = Style
FillMenu False
End Sub

Public Sub RemoveItem(ByVal Index As Integer, Optional ByVal PopUp As Boolean = False)
Dim Z As Long
Dim UI As Integer

UI = UBound(Menus(m).Items)
If Index < 1 Or Index > UI Then Exit Sub
If Index < UI Then
    For Z = Index To UI - 1
        Menus(m).Items(Z) = Menus(m).Items(Z + 1)
    Next
End If
If UI > 1 Then
    ReDim Preserve Menus(m).Items(1 To UI - 1)
Else
    Menus(m).Items(1).Caption = "#"
End If
UpdateMenu
End Sub

Public Sub Clear()
If SelItem(1) <> 0 Then
    SelectItem Menus(m).Items(SelItem(1)).Bounds, 0
    SelItem(1) = 0
End If

ReDim Menus(m).Items(1 To 1)
Menus(m).Items(1).Caption = "#"

UpdateMenu
End Sub

Public Sub ShowPopup(ByVal StartNum As Integer, ByVal X As Integer, ByVal Y As Integer, Optional ByVal ShowFromWhere As AlignConstants = vbAlignLeft)
IsAPopup = Menus(m).Items(StartNum).SubLevel
PopUpChild X, Y, IsAPopup, StartNum, ShowFromWhere
End Sub

Private Function GetPicWidth(Pic As IPictureDisp) As Long
GetPicWidth = setIconSize 'ScaleX(Pic.Width, vbHimetric, vbPixels)
End Function

Private Function GetPicHeight(Pic As IPictureDisp) As Long
GetPicHeight = setIconSize 'ScaleY(Pic.Height, vbHimetric, vbPixels)
End Function

Public Function Item(ByVal ItemName As String) As Long
Dim Z As Long
For Z = 1 To UBound(Menus(m).Items)
    If Menus(m).Items(Z).Caption = ItemName Then Item = Z: Exit Function
Next
End Function

Public Function IsAboveDragArea(ByVal X As Long, ByVal Y As Long) As Boolean
If PtInRect(DragArea, X, Y) Then IsAboveDragArea = True Else IsAboveDragArea = False
End Function

Private Function RectInVisibleArea(ptRect As RECT, Optional ByVal Level As Integer = 1) As Boolean
If Level = 1 Then
    If m_Orientation = Vertical Then
        If ptRect.Top > DragArea.Bottom + DragAreaIndent And ptRect.Bottom < UserControl.ScaleHeight - 10 Then RectInVisibleArea = True Else RectInVisibleArea = False
    Else
        If ptRect.Left > DragArea.Right + DragAreaIndent And ptRect.Right < UserControl.ScaleWidth - 10 Then RectInVisibleArea = True Else RectInVisibleArea = False
    End If
Else
    RectInVisibleArea = True
End If
End Function

Private Function AdjustSize()
Dim ProperSize As Long

ProperSize = GetProperSize
If UserControl.Extender.Align > 2.5 Then
    m_Orientation = OrientationStyle.Vertical
    If MenuCount(1) <> 0 Then
        ResizeBlock = True
        If UserControl.ScaleWidth <> ProperSize Then
            UserControl.Width = ProperSize * Screen.TwipsPerPixelX
            RaiseEvent Resize
        End If
        ResizeBlock = False
    End If
Else
    m_Orientation = OrientationStyle.Horizontal
    If MenuCount(1) <> 0 Then
        If UserControl.ScaleHeight <> ProperSize Then
            ResizeBlock = True
            UserControl.Height = ProperSize * Screen.TwipsPerPixelY
            RaiseEvent Resize
            ResizeBlock = False
        End If
    End If
End If
End Function

Private Sub DrawDragArea()
If Not DragAreaVisible Then Exit Sub
If m_Orientation = Horizontal Then
    DragArea.Left = DragAreaIndent
    DragArea.Top = DragAreaIndent
    DragArea.Right = DragArea.Left + DragAreaSize
    DragArea.Bottom = UserControl.ScaleHeight - DragAreaIndent
    ItemOffset = DragArea.Right
Else
    DragArea.Left = DragAreaIndent
    DragArea.Top = DragAreaIndent
    DragArea.Right = UserControl.ScaleWidth - DragAreaIndent
    DragArea.Bottom = DragArea.Top + DragAreaSize
    ItemOffset = DragArea.Bottom
End If
If DragArea.Right < DragArea.Left Or DragArea.Bottom < DragArea.Top Then
    DragArea.Left = 0
    DragArea.Right = UserControl.ScaleWidth
    DragArea.Top = 0
    DragArea.Bottom = UserControl.ScaleHeight - 2
End If
'DrawEdge UserControl.hDC, DragArea, EDGE_RAISED, BF_RECT
DrawEdgeThin DragArea, EDGE_RAISED
End Sub

Private Sub AdjustItems()
Dim Offset As Long, OldOffset As Long
Dim tSW As Long, tSH As Long, Z As Long
Dim Opt As Long, tempP As POINTAPI, TempHeight As Long, TempWidth As Long
If MenuCount(1) = 0 Then Exit Sub
tSW = UserControl.ScaleWidth - 1
tSH = UserControl.ScaleHeight - 1

If ItemOffset > IIf(m_Orientation = Horizontal, DragArea.Right, DragArea.Bottom) Then EnableScrollRight
MaxOffset = 0
Offset = ItemOffset
UserControl.Font.Charset = DefaultFontCharset
For Z = 1 To UBound(Menus(m).Items)
    If Menus(m).Items(Z).SubLevel = 1 And Menus(m).Items(Z).Visible And Not Menus(m).Items(Z).PopUp Then
        'UserControl.FontBold = Menus(M).Items(Z).Checked
        'UserControl.FontUnderline = Menus(M).Items(Z).Checked
        OldOffset = Offset

        If m_Orientation = Horizontal Then
            Offset = Offset + MenuHorizSpace
            Menus(m).Items(Z).Bounds.Left = Offset
            TempHeight = 0
            If Menus(m).Items(Z).Style <> TextOnly Then
                Offset = Offset + GetPicWidth(Menus(m).Items(Z).Icon)
                TempHeight = GetPicHeight(Menus(m).Items(Z).Icon)
            End If
            If Menus(m).Items(Z).Style <> GraphicsOnly Then
                Offset = Offset + GetItemWidth(Z)
                TempHeight = Maximum(TempHeight, GetItemHeight(Z))
            End If
            If Menus(m).Items(Z).Style = TextAndGraphics Then Offset = Offset + IconMargin
            Menus(m).Items(Z).Bounds.Right = Offset
            Menus(m).Items(Z).Bounds.Top = (tSH + 1 - TempHeight) \ 2
            Menus(m).Items(Z).Bounds.Bottom = tSH + 1 - Menus(m).Items(Z).Bounds.Top
            If Offset > UserControl.ScaleWidth - 2 And MaxOffset = 0 Then
                MaxOffset = OldOffset
            End If
        Else
            Offset = Offset + MenuVerticSpace
            TempWidth = 0
            TempHeight = 0
            Menus(m).Items(Z).Bounds.Top = Offset
            If Menus(m).Items(Z).Style <> TextOnly Then
                TempWidth = GetPicWidth(Menus(m).Items(Z).Icon)
                TempHeight = GetPicHeight(Menus(m).Items(Z).Icon)
            End If
            If Menus(m).Items(Z).Style <> GraphicsOnly Then
                If TempWidth <> 0 Then TempWidth = TempWidth + IconMargin
                TempWidth = TempWidth + GetItemWidth(Z)
                TempHeight = Maximum(GetItemHeight(Z), TempHeight)
            End If
            Offset = Offset + TempHeight
            Menus(m).Items(Z).Bounds.Bottom = Offset
            Menus(m).Items(Z).Bounds.Left = MenuHorizIndent
            Menus(m).Items(Z).Bounds.Right = TempWidth + MenuHorizIndent
            If Offset > UserControl.ScaleHeight - 10 And MaxOffset = 0 Then
                MaxOffset = OldOffset
            End If
        End If
    End If
Next Z
If MaxOffset <> 0 And ValidWidth = 0 Then ValidWidth = MaxOffset - DragAreaSize - DragAreaIndent
If MaxOffset <> 0 Then EnableScrollLeft: ItemOffset = MaxOffset Else EnableScrollLeft False
End Sub

Private Sub DrawItems()
Dim tSW As Long, tSH As Long
Dim Opt As Long, tempP As POINTAPI, TempPW As Long, TempPH As Long
Dim lpRect As RECT, Z As Long, T As String, Char As String
If MenuCount(1) = 0 Then Exit Sub
tSW = UserControl.ScaleWidth - 1
tSH = UserControl.ScaleHeight - 1
UserControl.Font.Charset = DefaultFontCharset

For Z = 1 To UBound(Menus(m).Items)
    If Menus(m).Items(Z).SubLevel = 1 And Menus(m).Items(Z).Visible And Not Menus(m).Items(Z).PopUp And RectInVisibleArea(Menus(m).Items(Z).Bounds) Then
        If Menus(m).Items(Z).BeginAGroup Then
            If m_Orientation = Horizontal Then
                lpRect.Left = Menus(m).Items(Z).Bounds.Left - MenuHorizSpace \ 2
                lpRect.Right = lpRect.Left
                lpRect.Top = 2
                lpRect.Bottom = UserControl.ScaleHeight - 2
                DrawEdge UserControl.hDC, lpRect, EDGE_ETCHED, BF_LEFT
            Else
                lpRect.Left = 2
                lpRect.Right = UserControl.ScaleWidth - 2
                lpRect.Top = Menus(m).Items(Z).Bounds.Top - MenuVerticSpace \ 2
                lpRect.Bottom = lpRect.Top
                DrawEdge UserControl.hDC, lpRect, EDGE_ETCHED, BF_TOP
            End If
        End If
        
        With Menus(m).Items(Z)
            If .Checked Then
                UserControl.DrawMode = MenuCheckDrawMode
                UserControl.Line (.Bounds.Left, .Bounds.Top)-(.Bounds.Right, .Bounds.Bottom), colMenuCheck, BF
                UserControl.DrawMode = vbCopyPen
            End If
        End With
        
        tempP.X = Menus(m).Items(Z).Bounds.Left
        tempP.Y = Menus(m).Items(Z).Bounds.Top
        If Menus(m).Items(Z).Style <> TextOnly Then
            TempPW = GetPicWidth(Menus(m).Items(Z).Icon)
            TempPH = GetPicHeight(Menus(m).Items(Z).Icon)
            'SetStretchBltMode UserControl.hDC, STRETCH_DELETESCANS    'BLACKONWHITE
            
            'UserControl.PaintPicture Menus(M).Items(Z).Icon, tempP.X, tempP.Y + ((Menus(M).Items(Z).Bounds.Bottom - Menus(M).Items(Z).Bounds.Top) - TempPH) / 2, TempPW, TempPH
            If Menus(m).Items(Z).Enabled Then
                TransparentBlt UserControl.hDC, Menus(m).Items(Z).Icon, tempP.X, tempP.Y + ((Menus(m).Items(Z).Bounds.Bottom - Menus(m).Items(Z).Bounds.Top) - TempPH) / 2, colIconTransparent
            Else
                DrawState UserControl.hDC, 0, 0, Menus(m).Items(Z).Icon, 0, tempP.X, tempP.Y, 0, 0, DSS_DISABLED Or DST_BITMAP
            End If
            
            'use bitmaps instead of icons
            tempP.X = tempP.X + IconMargin + TempPW
        End If
        If Menus(m).Items(Z).Style <> GraphicsOnly Then
            UserControl.CurrentX = tempP.X
            UserControl.CurrentY = tempP.Y + ((Menus(m).Items(Z).Bounds.Bottom - Menus(m).Items(Z).Bounds.Top) - GetItemHeight(Z)) / 2
            UserControl.Print Menus(m).Items(Z).Caption
        End If
        'If m_Orientation = Vertical Then Ang = -90 Else Ang = 0
        'PrintText UserControl, Menus(M).Items(Z).Caption, X, Y, Ang
        
        If IsSubMenu(Z) And m_Orientation = Vertical Then
            T$ = UserControl.Font.Name
            UserControl.Font.Name = "Marlett"
            If UserControl.Extender.Align = vbAlignRight Then
                Char$ = MenuLeftArrowSymbol
                UserControl.CurrentX = UserControl.TextWidth(Char$)
            ElseIf UserControl.Extender.Align = vbAlignLeft Then
                Char$ = MenuRightArrowSymbol
                UserControl.CurrentX = UserControl.ScaleWidth - 2 * UserControl.TextWidth(Char$)
            End If
            UserControl.CurrentY = Menus(m).Items(Z).Bounds.Top + (Menus(m).Items(Z).Bounds.Bottom - Menus(m).Items(Z).Bounds.Top) \ 2 - UserControl.TextHeight(Char$) \ 2
            UserControl.Print Char$
            UserControl.Font.Name = T$
        End If
        
    End If
Next Z
End Sub

Private Sub EnableScrollLeft(Optional ByVal Enable As Boolean = True)
Dim CW As Long, CH As Long
Dim Char As String, T As String
CanScrollLeft = Enable
If Enable Then
    T = UserControl.Font.Name
    UserControl.Font.Name = "Marlett"
    
    If m_Orientation = Horizontal Then
        Char = MenuRightArrowSymbol
        CW = UserControl.TextWidth(Char)
        CH = UserControl.TextHeight(Char)
        With ScrollAreaLeft
            .Left = UserControl.ScaleWidth - 2 - CW
            .Top = (UserControl.ScaleHeight - CH) \ 2
            .Right = UserControl.ScaleWidth - 2
            .Bottom = .Top + CH + 1
        End With
    ElseIf m_Orientation = Vertical Then
        Char = MenuDownArrowSymbol
        CW = UserControl.TextWidth(Char)
        CH = UserControl.TextHeight(Char)
        With ScrollAreaLeft
            .Left = (UserControl.ScaleWidth - CW) \ 2
            .Top = UserControl.ScaleHeight - CH - 2
            .Right = .Left + CW + 1
            .Bottom = UserControl.ScaleHeight - 2
        End With
    End If
    UserControl.CurrentX = ScrollAreaLeft.Left
    UserControl.CurrentY = ScrollAreaLeft.Top
    UserControl.ForeColor = 0
    UserControl.Print Char
    UserControl.Font.Name = T
End If
End Sub

Private Sub EnableScrollRight(Optional ByVal Enable As Boolean = True)
Dim Char As String, T As String, CW As Long, CH As Long
CanScrollRight = Enable
If Enable Then
    T = UserControl.Font.Name
    UserControl.Font.Name = "Marlett"
    If m_Orientation = Horizontal Then
        Char = MenuLeftArrowSymbol
        CW = UserControl.TextWidth(Char)
        CH = UserControl.TextHeight(Char)
        With ScrollAreaRight
            .Left = 4
            .Top = (UserControl.ScaleHeight - CH) \ 2
            .Right = 6 + CW
            .Bottom = .Top + CH + 1
        End With
    Else
        Char = MenuUpArrowSymbol
        CW = UserControl.TextWidth(Char)
        CH = UserControl.TextHeight(Char)
        With ScrollAreaRight
            .Left = (UserControl.ScaleWidth - CW) \ 2
            .Top = 4
            .Right = .Left + CW + 1
            .Bottom = 6 + CH
        End With
    End If
    DragAreaVisible = False
    UserControl.CurrentX = ScrollAreaRight.Left
    UserControl.CurrentY = ScrollAreaRight.Top
    UserControl.ForeColor = 0
    UserControl.Print Char
    UserControl.Font.Name = T
End If
End Sub

Private Sub ScrollLeft()
PaintMenuArea
ItemOffset = ItemOffset - ValidWidth + 4
AdjustItems
DrawDragArea
DrawItems
EnableScrollRight
End Sub

Private Sub ScrollRight()
PaintMenuArea
ItemOffset = ItemOffset + ValidWidth - 4
AdjustItems
If ItemOffset > 0 Then EnableScrollRight False
DrawDragArea
DrawItems
EnableScrollLeft
End Sub

Private Function NotTooDark(ByVal C As Long) As Long
Const D = 82
If C < 0 Then C = GetSysColor(C + SysColorTranslationBase)
If Red(C) < D And Green(C) < D And Blue(C) < D Then NotTooDark = &H808080 Else NotTooDark = C
End Function

Public Property Get ToolbarMode() As MenuBarState
ToolbarMode = m_ToolbarMode
End Property

Public Property Let ToolbarMode(ByVal vNewValue As MenuBarState)
m_ToolbarMode = vNewValue
End Property

Public Property Get DontRaiseResizeEvent() As Boolean
DontRaiseResizeEvent = DontRaiseResize
End Property

Public Property Let DontRaiseResizeEvent(ByVal vNewValue As Boolean)
DontRaiseResize = vNewValue
End Property

'========================================================
'Interface to menu specific data
'========================================================

Public Sub CheckItem(ByVal Index As Long, ByVal Check As Boolean)
If Not IsValidIndex(Index) Then Exit Sub
Menus(m).Items(Index).Checked = Check
End Sub

Public Sub ShortcutKeyClick(ByVal KeyCode As Long)
Dim Z As Long

For Z = 1 To UBound(Menus(m).Items)
    If Menus(m).Items(Z).ShortcutKey = KeyCode Then
        RaiseEvent Command(Z, Menus(m).Items(Z).Caption)
        Exit Sub
    End If
Next
End Sub

Public Function ItemChecked(ByVal Index As Long) As Boolean
If Not IsValidIndex(Index) Then Exit Function
ItemChecked = Menus(m).Items(Index).Checked
End Function

Public Function ItemCaption(ByVal Index As Long) As String
If Not IsValidIndex(Index) Then Exit Function
ItemCaption = Menus(m).Items(Index).Caption
End Function

Public Function ItemTooltipText(ByVal Index As Long) As String
If Not IsValidIndex(Index) Then Exit Function
ItemTooltipText = Menus(m).Items(Index).ToolTipText
End Function

Public Function ItemAuxIndex(ByVal Index As Long) As Variant
If Not IsValidIndex(Index) Then Exit Function
ItemAuxIndex = Menus(m).Items(Index).AuxIndex
End Function

'========================================================
'Checks index sanity
'========================================================

Public Function IsValidIndex(ByVal Index As Long) As Boolean
If Index >= LBound(Menus(m).Items) And Index <= UBound(Menus(m).Items) Then IsValidIndex = True
End Function

'========================================================
'Holds an index of this ctl menu in global Menus() array
'========================================================

Public Property Get MenuID() As Long
MenuID = m
End Property

Public Property Let MenuID(ByVal vNewValue As Long)
m = vNewValue
End Property

Public Property Get WhatsThisHelp() As Boolean
WhatsThisHelp = m_WhatsThisHelp
End Property

Public Property Let WhatsThisHelp(ByVal vNewValue As Boolean)
m_WhatsThisHelp = vNewValue
End Property

