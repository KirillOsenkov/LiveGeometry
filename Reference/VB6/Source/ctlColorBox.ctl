VERSION 5.00
Begin VB.UserControl ctlColorBox 
   ClientHeight    =   240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   240
   ScaleHeight     =   20
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   20
   ToolboxBitmap   =   "ctlColorBox.ctx":0000
   Begin VB.PictureBox picPopup 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   360
      ScaleHeight     =   141
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   181
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   240
      Visible         =   0   'False
      Width           =   2175
   End
End
Attribute VB_Name = "ctlColorBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'====================================================
' This ActiveX control is a color choice box. It looks like a simple button
' with a single-colored background. When pushed, it pops up a nice small
' color palette with 40 colors and More and Cancel buttons. When selected
' a color, its background color becomes the selected color.
' More button pops up standard Windows ChooseColor dialog.
'
' Has only one significant Public property:
'
' ctlColorBox.Color
'
' Generates a ColorChanged event, that brings NewColor and OldColor parameters
' This event is generated AFTER the color has changed.
'
' which represents the color currently selected into the control.
' 8 lower colors are 8 most recently selected colors.
' Rows 4 and 5 are standard Windows Color choice dialog customizable
' colors.
'====================================================
Option Explicit

Const SWP_SHOWWINDOW = &H40

Private Declare Function GetCapture Lib "user32" () As Long
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetActiveWindow Lib "user32" () As Long
Private Declare Function SetActiveWindow Lib "user32" (ByVal hWnd As Long) As Long

'====================================================
' Interesting parameters...
'====================================================
Const Columns = 8
Const Rows = 6
Const ColorCount = Rows * Columns
Const Offset As Long = 2 ' pixels from the border
Const defButtonHeight = 25 ' More and Cancel button height in pixels
Const defCellWidth = 16, defCellHeight = 16

Const ActiveCellColor = vbWhite
Const InactiveCellColor = vbBlack

Private Type ColBox
    Color As Long
    R As RECT
End Type

Dim Colors() As ColBox
Dim MRUColors(1 To Columns) As Long
Dim CancelBox As RECT
Dim SelectBox As RECT
Dim Alignment As Long
Dim HasFocus As Boolean
Dim ButtonWasDown As Boolean

Dim CellWidth As Long, CellHeight As Long
Dim ButtonHeight As Long
Dim SW As Long, SH As Long

Public Cancelled As Boolean

Dim CD As New clsCommonDialog
Dim Col As Long
Dim ActiveCell As Long
Dim PopupShown As Long
Dim WasMouseDown As Boolean
Dim ParentKeyPreview As Boolean
Dim m_DefaultValue As Boolean

Public Event ColorChanged(ByVal NewColor As Long, ByVal OldColor As Long)
Public Event PopupIsShowing()
Public Event PopupWasHidden()

Const HS_BDIAGONAL = 3

Private Sub SelectBoxMouseDown()
Cancel
DoEvents

UserControl.Extender.Parent.Enabled = False
CD.Color = Color
CD.ShowColor
UserControl.Extender.Parent.Enabled = True
If CD.Cancelled Then Exit Sub
UserControl.SetFocus

SetColor CD.Color
End Sub

Private Sub picPopup_KeyDown(KeyCode As Integer, Shift As Integer)
Cancel
End Sub

Private Sub picPopup_LostFocus()
Cancel
End Sub

'====================================================
'                                   picPopup EVENTS
'====================================================
Private Sub picPopup_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim C As Long
C = FindCell(X, Y)
ButtonWasDown = False

If C > 0 Then
    SetColor Colors(C).Color
    Cancelled = False
    HidePopup
    Exit Sub
Else
    If X >= SelectBox.Left And X <= SelectBox.Right And Y >= SelectBox.Top And Y <= SelectBox.Bottom Then
        DrawEdge picPopup.hDC, SelectBox, EDGE_SUNKEN, BF_RECT
        picPopup.Refresh
        SelectBoxMouseDown
        Exit Sub
    End If
    If X >= CancelBox.Left And X <= CancelBox.Right And Y >= CancelBox.Top And Y <= CancelBox.Bottom Then
        DrawEdge picPopup.hDC, CancelBox, EDGE_SUNKEN, BF_RECT
        picPopup.Refresh
        Cancel
        Exit Sub
    End If
End If

Cancel
End Sub

Private Sub picPopup_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim C As Long
C = FindCell(X, Y)

If C > 0 Then
'    SetColor Colors(C).Color
'    Cancelled = False
'    HidePopup
    ButtonWasDown = True
    DrawSelectedCell ActiveCell, True
    Exit Sub
Else
    If X >= SelectBox.Left And X <= SelectBox.Right And Y >= SelectBox.Top And Y <= SelectBox.Bottom Then
        DrawEdge picPopup.hDC, SelectBox, EDGE_SUNKEN, BF_RECT
        picPopup.Refresh
        'SelectBoxMouseDown
        Exit Sub
    End If
    If X >= CancelBox.Left And X <= CancelBox.Right And Y >= CancelBox.Top And Y <= CancelBox.Bottom Then
        DrawEdge picPopup.hDC, CancelBox, EDGE_SUNKEN, BF_RECT
        picPopup.Refresh
        'Cancel
        Exit Sub
    End If
End If

'Cancel
End Sub

Private Sub UserControl_GotFocus()
HasFocus = True
Repaint
End Sub

'====================================================
'                           UserControl events
'====================================================
Private Sub UserControl_Initialize()
SetWindowLong picPopup.hWnd, GWL_EXSTYLE, WS_EX_TOPMOST Or &H80
SetWindowPos picPopup.hWnd, -1, 0, 0, 0, 0, 3
SetParent picPopup.hWnd, 0
ButtonWasDown = False

CellWidth = defCellWidth
CellHeight = defCellHeight
ButtonHeight = picPopup.TextHeight("W") + 2 * 6

Dim tWidth As Long
tWidth = picPopup.TextWidth(GetString(ResOtherColor)) + picPopup.TextWidth(GetString(ResCancel)) + 16
If (CellWidth + 1) * Columns + 2 * Offset - 8 < tWidth Then
    CellWidth = tWidth \ Columns + 1
    CellHeight = CellWidth
End If

SW = 2 * Offset + Columns * (CellWidth + 1) + 1
SH = 3 * Offset + Rows * (CellHeight + 1) + 1 + Maximum(ButtonHeight, picPopup.TextHeight("W") + 4)

LoadMRUColors
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
If PopupShown Then
    Cancel
Else
    If KeyCode = vbKeySpace Then ShowPopup
End If
End Sub

Private Sub UserControl_LostFocus()
HasFocus = False
Cancel
Repaint
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Shift <> 0 Or Button <> 1 Then Exit Sub
WasMouseDown = True
Repaint True
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Shift <> 0 Or Button <> 1 Then Exit Sub
If WasMouseDown And Button = 1 And ((Y > UserControl.ScaleHeight And (Alignment And 2) = 0) Or (Y < 0 And (Alignment And 2) = 2)) Then
    ButtonWasDown = True
    ShowPopup
End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Shift <> 0 Or Button <> 1 Then Exit Sub
If Not WasMouseDown Then Exit Sub
WasMouseDown = False
If Not PopupShown Then ShowPopup Else Cancel
End Sub

Private Sub UserControl_Paint()
Repaint
End Sub

Private Sub UserControl_Resize()
Repaint
End Sub

'====================================================
'                           POPUP SECTION
'====================================================
Private Sub ShowPopup()
On Local Error Resume Next
Dim P As POINTAPI, R As RECT
Dim ScreenX As Long, ScreenY As Long

ScreenX = GetSystemMetrics(SM_CXSCREEN)
ScreenY = GetSystemMetrics(SM_CYSCREEN)

picPopup.AutoRedraw = True

Cancelled = False
LoadMRUColors
FillColors

RaiseEvent PopupIsShowing

PopupShown = True
Repaint Sunken:=True

UserControl.Parent.Enabled = False
GetWindowRect UserControl.hWnd, R

Alignment = 0
If R.Left + SW > ScreenX Then
    R.Left = R.Right - SW
    Alignment = 1
End If
If R.Bottom + SH > ScreenY Then
    R.Bottom = R.Top - SH
    Alignment = Alignment + 2
End If

Dim SWidth As Long
SWidth = Maximum(SW, picPopup.TextWidth(GetString(ResOtherColor)) + picPopup.TextWidth(GetString(ResCancel)) + 20)

SetWindowPos picPopup.hWnd, HWND_TOPMOST, R.Left, R.Bottom, SWidth, SH, SWP_SHOWWINDOW
With SelectBox
    .Left = Offset
    .Top = 2 * Offset + Rows * (CellHeight + 1)
    .Right = picPopup.ScaleWidth \ 2
    .Bottom = picPopup.ScaleHeight - Offset
    Dim tWidth As Long
    tWidth = picPopup.TextWidth(GetString(ResOtherColor))
    If .Right - .Left - 10 < tWidth Then .Right = .Left + picPopup.TextWidth(GetString(ResOtherColor)) + 10
End With
With CancelBox
    .Left = SelectBox.Right
    .Top = SelectBox.Top
    .Right = picPopup.ScaleWidth - Offset
    .Bottom = picPopup.ScaleHeight - Offset
End With

PaintPopup

Set modHookWindow.ctlActiveColorbox = Me

picPopup.SetFocus

HookWindow picPopup.hWnd
SetCapture picPopup.hWnd

picPopup.SetFocus

SetForegroundWindow picPopup.hWnd
End Sub

Private Sub HidePopup()
On Local Error Resume Next
If Not PopupShown Then Exit Sub
PopupShown = False

ShowWindow picPopup.hWnd, 0
ReleaseCapture
UnhookWindow picPopup.hWnd
Set modHookWindow.ctlActiveColorbox = Nothing

picPopup.AutoRedraw = False
UserControl.Parent.Enabled = True
UserControl.Parent.SetFocus
RaiseEvent PopupWasHidden

Repaint
End Sub

Private Sub PaintPopup()
Dim i, j As Long
Dim R As RECT
Dim Offs As Long

picPopup.Cls

Offs = (picPopup.ScaleWidth - Columns * (CellWidth + 1)) \ 2
For j = 1 To Rows
    For i = 1 To Columns
        With Colors(i + (j - 1) * Columns).R
            .Left = Offs + (CellWidth + 1) * (i - 1) + 1
            .Right = .Left + CellWidth - 1
            .Top = Offset + (CellHeight + 1) * (j - 1) + 1
            .Bottom = .Top + CellHeight - 1
            picPopup.Line (.Left, .Top)-(.Right, .Bottom), Colors(i + (j - 1) * Columns).Color, BF
            picPopup.Line (.Left - 1, .Top - 1)-(.Right + 1, .Bottom + 1), InactiveCellColor, B
        End With
    Next
Next

R.Right = picPopup.ScaleWidth
R.Bottom = picPopup.ScaleHeight
DrawEdge picPopup.hDC, R, EDGE_RAISED, BF_RECT

picPopup.Line (SelectBox.Left, SelectBox.Top)-(SelectBox.Right - 1, SelectBox.Bottom - 1), vbButtonFace, BF
picPopup.Line (CancelBox.Left, CancelBox.Top)-(CancelBox.Right - 1, CancelBox.Bottom - 1), vbButtonFace, BF

DrawEdge picPopup.hDC, SelectBox, EDGE_RAISED, BF_RECT
DrawEdge picPopup.hDC, CancelBox, EDGE_RAISED, BF_RECT

Dim sColor As String
sColor = GetString(ResOtherColor)
DrawText picPopup.hDC, sColor, Len(sColor), SelectBox, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE
sColor = GetString(ResCancel)
DrawText picPopup.hDC, sColor, Len(sColor), CancelBox, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE
End Sub

Private Sub FillColors()
Dim Z As Long
ReDim Colors(1 To Columns * Rows)
Colors(1).Color = RGB(255, 0, 0)
Colors(2).Color = RGB(255, 255, 0)
Colors(3).Color = RGB(0, 255, 0)
Colors(4).Color = RGB(0, 255, 255)
Colors(5).Color = RGB(0, 0, 255)
Colors(6).Color = RGB(255, 0, 255)
Colors(7).Color = RGB(0, 0, 0)
Colors(8).Color = RGB(255, 255, 255)
Colors(9).Color = RGB(255, 192, 192)
Colors(10).Color = RGB(255, 255, 192)
Colors(11).Color = RGB(192, 255, 192)
Colors(12).Color = RGB(192, 255, 255)
Colors(13).Color = RGB(192, 192, 255)
Colors(14).Color = RGB(255, 192, 255)
Colors(15).Color = RGB(128, 128, 128)
Colors(16).Color = RGB(192, 192, 192)
For Z = 17 To 32
    Colors(Z).Color = ColorArray(Z - 17)
Next
For Z = 33 To 40
    Colors(Z).Color = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
Next
For Z = 41 To 48
    Colors(Z).Color = MRUColors(Z - 40)
Next
End Sub

'====================================================
'                           CELL OPERATION SECTION
'====================================================
Public Sub MouseMoved(ByVal X As Long, ByVal Y As Long)
Dim OldActiveCell As Long

OldActiveCell = ActiveCell
ActiveCell = FindCell(X, Y)

If OldActiveCell <> ActiveCell Then
    If OldActiveCell <> 0 Then DrawSelectedCell OldActiveCell
    If ActiveCell <> 0 Then DrawSelectedCell ActiveCell, True
End If
End Sub

Public Sub Cancel()
If Not PopupShown Then Exit Sub
Cancelled = True
HidePopup
End Sub

Public Sub LosingCapture()
Dim P As POINTAPI, R As RECT
GetCursorPos P
GetWindowRect picPopup.hWnd, R
If P.X < R.Left Or P.X > R.Right Or P.Y < R.Top Or P.Y > R.Bottom Then Cancel
End Sub

Private Sub DrawSelectedCell(ByVal Index As Long, Optional ByVal DoSelect As Boolean = False)
Dim cLight As Long, cDark As Long
If Index < 1 Or Index > ColorCount Then Exit Sub

If DoSelect Then
    If ButtonWasDown Then
        cLight = vb3DDKShadow
        cDark = vb3DHighlight
    Else
        cLight = vb3DHighlight
        cDark = vb3DDKShadow
    End If
    picPopup.Line (Colors(Index).R.Left - 1, Colors(Index).R.Top - 1)-(Colors(Index).R.Right + 1, Colors(Index).R.Top - 1), cLight
    picPopup.Line (Colors(Index).R.Left - 1, Colors(Index).R.Top - 1)-(Colors(Index).R.Left - 1, Colors(Index).R.Bottom + 1), cLight
    picPopup.Line (Colors(Index).R.Right + 1, Colors(Index).R.Top - 1)-(Colors(Index).R.Right + 1, Colors(Index).R.Bottom + 1), cDark
    picPopup.Line (Colors(Index).R.Left - 1, Colors(Index).R.Bottom + 1)-(Colors(Index).R.Right + 1, Colors(Index).R.Bottom + 1), cDark
Else
    picPopup.Line (Colors(Index).R.Left - 1, Colors(Index).R.Top - 1)-(Colors(Index).R.Right + 1, Colors(Index).R.Bottom + 1), InactiveCellColor, B
End If

End Sub

Private Function FindCell(ByVal X As Long, ByVal Y As Long) As Long
Dim Z As Long

For Z = 1 To Rows * Columns
    If X >= Colors(Z).R.Left And X <= Colors(Z).R.Right + 1 And Y >= Colors(Z).R.Top And Y <= Colors(Z).R.Bottom + 1 Then
        FindCell = Z
        Exit Function
    End If
Next

FindCell = 0
End Function

'====================================================
'                           MRU COLORS SECTION
'====================================================
Private Sub LoadMRUColors()
Dim Z As Long

For Z = 1 To Columns
    MRUColors(Z) = Val(GetSetting(AppName, "Custom colors", "MRUColor" & Format(Z), vbWhite))
Next
End Sub

Private Sub SaveMRUColors()
Dim Z As Long

For Z = 1 To Columns
    SaveSetting AppName, "Custom colors", "MRUColor" & Z, Format(MRUColors(Z))
Next
End Sub

Private Sub AddMRUColor(ByVal C As Long)
Dim Z As Long

For Z = 1 To Columns - 1
    MRUColors(Z) = MRUColors(Z + 1)
Next

MRUColors(Columns) = C
End Sub

'====================================================
'                           PROPERTIES AND EVENTS
'====================================================
Private Sub Repaint(Optional Sunken As Boolean = False)
Dim R As RECT, hBrush As Long, EdgeStyle As Long

UserControl.BackColor = EnsureRGB(IIf(Enabled And Not m_DefaultValue, Col, vbButtonFace))
UserControl.Cls
R.Right = UserControl.ScaleWidth
R.Bottom = UserControl.ScaleHeight
EdgeStyle = EDGE_RAISED - 5 * Sunken
If Not UserControl.Enabled Then EdgeStyle = EDGE_ETCHED
If m_DefaultValue And UserControl.Enabled Then
    hBrush = CreateHatchBrush(HS_BDIAGONAL, EnsureRGB(vbButtonText))
    FillRect UserControl.hDC, R, hBrush
    DeleteObject hBrush
End If
DrawEdge UserControl.hDC, R, EdgeStyle, BF_RECT
If HasFocus Then DrawFocus Sunken
'UserControl.Refresh
End Sub

Public Property Get Color() As OLE_COLOR
Color = Col
End Property

Public Property Let Color(ByVal vNewValue As OLE_COLOR)
If Col <> vNewValue Then
    Col = vNewValue
    Repaint
End If
End Property

Public Sub SetColor(ByVal NewColor As Long)
Dim OldColor As Long

m_DefaultValue = False
OldColor = Col
Color = NewColor
If NewColor <> MRUColors(Columns) Then
    AddMRUColor NewColor
    SaveMRUColors
End If
RaiseEvent ColorChanged(NewColor, OldColor)
End Sub

Public Property Get Enabled() As Boolean
Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal vNewValue As Boolean)
UserControl.Enabled = vNewValue
Repaint
End Property

Private Sub DrawFocus(Optional ByVal Sunken As Boolean = False)
Dim R As RECT
R.Left = 2 - Sunken
R.Top = 2 - Sunken
R.Right = UserControl.ScaleWidth - 3 - Sunken
R.Bottom = UserControl.ScaleHeight - 3 - Sunken
DrawFocusRect UserControl.hDC, R
End Sub

Public Property Get DefaultValue() As Boolean
DefaultValue = m_DefaultValue
End Property

Public Property Let DefaultValue(ByVal vNewValue As Boolean)
If m_DefaultValue <> vNewValue Then
    m_DefaultValue = vNewValue
    Repaint
End If
End Property
