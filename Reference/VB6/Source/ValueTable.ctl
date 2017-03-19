VERSION 5.00
Begin VB.UserControl ctlValueTable 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   ClientHeight    =   2715
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2670
   ClipBehavior    =   0  'None
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForwardFocus    =   -1  'True
   HitBehavior     =   0  'None
   MousePointer    =   1  'Arrow
   ScaleHeight     =   181
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   178
   ToolboxBitmap   =   "ValueTable.ctx":0000
   Begin VB.CommandButton cmdEdit 
      Enabled         =   0   'False
      Height          =   300
      Left            =   840
      MaskColor       =   &H00BFBFBF&
      Picture         =   "ValueTable.ctx":0312
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   360
      UseMaskColor    =   -1  'True
      Width           =   300
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "-"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   480
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   360
      Width           =   300
   End
   Begin VB.ListBox lstValues 
      Enabled         =   0   'False
      Height          =   1530
      Left            =   1200
      MousePointer    =   1  'Arrow
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   735
      Width           =   900
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   360
      Width           =   300
   End
   Begin VB.ListBox lstVars 
      Enabled         =   0   'False
      Height          =   1530
      Left            =   60
      MousePointer    =   1  'Arrow
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   735
      Width           =   975
   End
   Begin VB.Line LineN 
      BorderColor     =   &H80000014&
      Visible         =   0   'False
      X1              =   126
      X2              =   140
      Y1              =   1
      Y2              =   0
   End
   Begin VB.Line LineW 
      BorderColor     =   &H80000014&
      Visible         =   0   'False
      X1              =   129
      X2              =   130
      Y1              =   3
      Y2              =   12
   End
   Begin VB.Line LineE 
      BorderColor     =   &H80000015&
      Visible         =   0   'False
      X1              =   144
      X2              =   144
      Y1              =   2
      Y2              =   12
   End
   Begin VB.Line LineS 
      BorderColor     =   &H80000015&
      Visible         =   0   'False
      X1              =   143
      X2              =   132
      Y1              =   19
      Y2              =   18
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "r"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   6.75
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   1995
      TabIndex        =   3
      Top             =   75
      Width           =   150
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Experiment expressions"
      Height          =   240
      Left            =   -15
      TabIndex        =   0
      Top             =   15
      Width           =   2055
   End
End
Attribute VB_Name = "ctlValueTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Event Declarations:
Event NeedToHide()
Event NeedToShow()
Event NeedToClose()

Dim Hidden As Boolean

Public Sub Resize()
UserControl_Resize
End Sub

Public Sub Clear()
'If WECount = 0 Then Exit Sub
'Do While WECount > 0
'    RemoveWatchExpression WECount, False, False
'Loop
lstVars.Clear
lstValues.Clear
cmdRemove.Enabled = False
cmdEdit.Enabled = False
lstValues.Enabled = False
lstVars.Enabled = False
End Sub

Public Sub UpdateExpressions()
Dim Z As Long
On Local Error Resume Next
If WECount = 0 Then Exit Sub
For Z = 0 To lstVars.ListCount - 1
    If lstVars.List(Z) <> WatchExpressions(Z + 1).Name Then lstVars.List(Z) = WatchExpressions(Z + 1).Name
    WatchExpressions(Z + 1).Value = RecalculateTree(WatchExpressions(Z + 1).WatchTree, 1)
    If WasThereAnErrorEvaluatingLastExpression <> eetEverythingOK Then
        WasThereAnErrorEvaluatingLastExpression = eetEverythingOK
        lstValues.List(Z) = GetString(ResError)
    Else
        lstValues.List(Z) = Round(WatchExpressions(Z + 1).Value, setNumberPrecision)
    End If
Next
lstValues.Refresh
End Sub

Private Sub cmdAdd_Click()
On Local Error GoTo EH
Dim InputStr As String, EqvIns As Long
Dim tName As String, tExpression As String

SetFocus
InputStr = Trim(UCase(InputBox(GetString(ResEnterExpression), , "", , , App.HelpFile, ResHlp_Interface_Expressions)))
If InputStr = "" Then Exit Sub
EqvIns = InStr(InputStr, "=")
If EqvIns <> 0 Then
    tName = Left(InputStr, EqvIns - 1)
    tExpression = Right(InputStr, Len(InputStr) - EqvIns)
Else
    tName = InputStr
    tExpression = InputStr
End If
If Not AddWatchExpression(tName, tExpression) Then
    'do nothing
Else
    lstVars.AddItem tName
    lstValues.AddItem Round(WatchExpressions(WECount).Value, setNumberPrecision)
    lstVars.ListIndex = WECount - 1
    lstValues.ListIndex = WECount - 1
    cmdRemove.Enabled = True
    cmdEdit.Enabled = True
    lstValues.Enabled = True
    lstVars.Enabled = True
    MakeHSB
End If
Exit Sub

EH:
End Sub

Private Sub cmdEdit_Click()
lstVars_DblClick
End Sub

Private Sub cmdRemove_Click()
RemoveWatchExpression lstVars.ListIndex + 1
End Sub

Private Sub lblCaption_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
UserControl_MouseMove Button, Shift, X, Y
End Sub

Private Sub lblX_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
LineN.BorderColor = vb3DDKShadow
LineE.BorderColor = vb3DHighlight
LineS.BorderColor = vb3DHighlight
LineW.BorderColor = vb3DDKShadow
End Sub

Private Sub lblX_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Not LineN.Visible Then
    LineN.Visible = True
    LineE.Visible = True
    LineS.Visible = True
    LineW.Visible = True
End If
SelectLabel lblX
End Sub

Private Sub lblX_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
LineE.BorderColor = vb3DDKShadow
LineN.BorderColor = vb3DHighlight
LineW.BorderColor = vb3DHighlight
LineS.BorderColor = vb3DDKShadow
RaiseEvent NeedToClose
End Sub

Private Sub lstValues_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then
    If lstVars.ListIndex < lstVars.ListCount - 1 Then lstVars.ListIndex = lstValues.ListIndex + 1
ElseIf KeyCode = vbKeyUp Then
    If lstVars.ListIndex > 0 Then lstVars.ListIndex = lstValues.ListIndex - 1
End If
End Sub

Private Sub lstValues_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lstVars.ListIndex = lstValues.ListIndex
End Sub

Private Sub lstVars_DblClick()
On Local Error GoTo EH:
Dim InputStr As String, EqvIns As Long, Q As Long
Dim tName As String, tExpression As String
SetFocus
Q = lstVars.ListIndex + 1
If WatchExpressions(Q).Name = WatchExpressions(Q).Expression Then WatchExpressions(Q).Name = ""
InputStr = InputBox(GetString(ResEnterExpression), , WatchExpressions(Q).Expression)
If InputStr = "" Then Exit Sub

EqvIns = InStr(InputStr, "=")
If EqvIns <> 0 Then
    tExpression = Trim(UCase(Right(InputStr, Len(InputStr) - EqvIns)))
Else
    tExpression = Trim(UCase(InputStr))
End If

If WatchExpressions(Q).Name = "" Then tName = tExpression Else tName = WatchExpressions(Q).Name
If EditWatchExpression(Q, tName, tExpression) Then
    lstVars.List(Q - 1) = tName
    lstValues.List(Q - 1) = Round(WatchExpressions(Q).Value, setNumberPrecision)
    lstVars.ListIndex = Q - 1
    lstValues.ListIndex = Q - 1
    
    UpdateExpressions
End If
Exit Sub

EH:
End Sub

Private Sub lstVars_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then
    If lstValues.ListIndex < lstValues.ListCount - 1 Then lstValues.ListIndex = lstVars.ListIndex + 1
ElseIf KeyCode = vbKeyUp Then
    If lstValues.ListIndex > 0 Then lstValues.ListIndex = lstVars.ListIndex - 1
End If
End Sub

Private Sub lstVars_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lstValues.ListIndex = lstVars.ListIndex
End Sub

Private Sub UserControl_Initialize()
cmdAdd.BackColor = UserControl.BackColor
lblCaption.Caption = GetString(ResWatchExpressions)
End Sub

Private Sub MakeHSB()
Dim T As Long, Z As Long, T2 As Long
If lstVars.ListCount <= 0 Then Exit Sub
T = 0
For Z = 0 To lstVars.ListCount - 1
    T2 = TextWidth(lstVars.List(Z))
    If T2 > T Then T = T2
Next Z
T = T + 4
If T > lstVars.Width Then SendMessage lstVars.hWnd, &H194, T, 0
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If (X >= LineN.X1 And X <= LineN.X2 And Y >= LineE.Y1 And Y <= LineE.Y2) Then Exit Sub
If LineN.Visible Then
    LineN.Visible = False
    LineE.Visible = False
    LineS.Visible = False
    LineW.Visible = False
End If
End Sub

Private Sub UserControl_Resize()
Dim R As RECT, tHeight As Long
If UserControl.ScaleHeight < 4 * TextHeight("W") Or UserControl.ScaleWidth < 16 Then Exit Sub

lblX.Move UserControl.ScaleWidth - 4 - lblX.Width, 4
lblCaption.Move 4, 4, UserControl.ScaleWidth - 8 - lblX.Width, UserControl.TextHeight("W") + 2
tHeight = lblCaption.Top + lblCaption.Height + 2

cmdAdd.Move 4, tHeight
cmdRemove.Move cmdAdd.Left + cmdAdd.Width, cmdAdd.Top, cmdAdd.Width, cmdAdd.Height
cmdEdit.Move cmdRemove.Left + cmdRemove.Width, cmdAdd.Top, cmdAdd.Width, cmdAdd.Height

tHeight = tHeight + cmdAdd.Height + 2
lstVars.Move 4, tHeight, UserControl.ScaleWidth / 2 - 4, ScaleHeight - 4 - tHeight
lstValues.Move lstVars.Width + 4, tHeight, UserControl.ScaleWidth - 8 - lstVars.Width, lstVars.Height

If setGradientFill Then
    Gradient UserControl.hDC, setcolToolbar, UserControl.BackColor, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
Else
    UserControl.Cls
End If

R.Right = UserControl.ScaleWidth
R.Bottom = UserControl.ScaleHeight
DrawEdge UserControl.hDC, R, EDGE_RAISED, BF_RECT
LineN.X1 = lblX.Left - 1
LineN.X2 = lblX.Left + lblX.Width - 1
LineN.Y1 = lblX.Top - 1
LineN.Y2 = LineN.Y1
LineE.X1 = LineN.X2
LineE.X2 = LineN.X2
LineE.Y1 = LineN.Y1
LineE.Y2 = LineN.Y1 + lblX.Height
LineS.X1 = LineN.X1
LineS.X2 = LineN.X2
LineS.Y1 = LineE.Y2
LineS.Y2 = LineE.Y2
LineW.X1 = LineN.X1
LineW.X2 = LineN.X1
LineW.Y1 = LineN.Y1
LineW.Y2 = LineE.Y2
LineE.Y2 = LineE.Y2 + 1

End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
If lblCaption.Caption <> GetString(ResWatchExpressions) Then lblCaption.Caption = GetString(ResWatchExpressions)
UserControl.Refresh
End Sub

Private Sub SelectLabel(lblLabel As Label)
If LineN.Y1 = lblLabel.Top - 1 Then Exit Sub
LineN.X1 = lblLabel.Left - 1
LineN.X2 = lblLabel.Left + lblLabel.Width - 1
LineN.Y1 = lblLabel.Top - 1
LineN.Y2 = LineN.Y1
LineE.X1 = LineN.X2
LineE.X2 = LineN.X2
LineE.Y1 = LineN.Y1
LineE.Y2 = LineN.Y1 + lblLabel.Height
LineS.X1 = LineN.X1
LineS.X2 = LineN.X2
LineS.Y1 = LineE.Y2
LineS.Y2 = LineE.Y2
LineW.X1 = LineN.X1
LineW.X2 = LineN.X1
LineW.Y1 = LineN.Y1
LineW.Y2 = LineE.Y2
LineE.Y2 = LineE.Y2 + 1
End Sub

Public Sub AddExpression(ByVal exName As String, ByVal exExpression As String, Optional ByVal ShouldRecord As Boolean = True)
AddWatchExpression exName, exExpression, ShouldRecord
lstVars.AddItem exName
lstValues.AddItem Round(WatchExpressions(WECount).Value, setNumberPrecision)
lstVars.ListIndex = WECount - 1
lstValues.ListIndex = WECount - 1
cmdRemove.Enabled = True
cmdEdit.Enabled = True
lstValues.Enabled = True
lstVars.Enabled = True
MakeHSB
End Sub

Public Sub RecreateWE()
Dim Z As Long
lstVars.Clear
lstValues.Clear
For Z = 1 To WECount
    lstVars.AddItem WatchExpressions(Z).Name
    lstValues.AddItem Round(WatchExpressions(WECount).Value, setNumberPrecision)
Next Z
cmdRemove.Enabled = WECount > 0
cmdEdit.Enabled = WECount > 0
lstValues.Enabled = WECount > 0
lstVars.Enabled = WECount > 0
If WECount > 0 Then
    lstVars.ListIndex = WECount - 1
    lstValues.ListIndex = WECount - 1
End If
MakeHSB
End Sub

Public Sub RemoveExpression(ByVal Index As Long, Optional ByVal ShouldUpdateExpressions As Boolean = True)
Dim Q As Long
Q = Index - 1
If Q < 0 Then Exit Sub
lstVars.RemoveItem Q
lstValues.RemoveItem Q
Q = Q + 1
If Q <= WECount Then
    lstVars.ListIndex = Q - 1
    lstValues.ListIndex = Q - 1
Else
    lstVars.ListIndex = Q - 2
    lstValues.ListIndex = Q - 2
End If

If WECount = 0 Then
    cmdRemove.Enabled = False
    cmdEdit.Enabled = False
    lstValues.Enabled = False
    lstVars.Enabled = False
    lstVars.Clear
    lstValues.Clear
Else
    If ShouldUpdateExpressions Then UpdateExpressions
End If
End Sub

Public Sub WEInsert(ByVal sName As String, ByVal sValue As String)
lstVars.AddItem sName
lstValues.AddItem sValue
lstVars.ListIndex = lstVars.ListCount - 1
lstValues.ListIndex = lstVars.ListCount - 1
cmdRemove.Enabled = True
cmdEdit.Enabled = True
lstValues.Enabled = True
lstVars.Enabled = True
MakeHSB
End Sub

Public Sub WERemove(ByVal Index As Long) 'Index from 1 to XCount
Dim Q As Long
Q = Index - 1
If Q < 0 Or Q > lstVars.ListCount - 1 Then Exit Sub

lstVars.RemoveItem Q
lstValues.RemoveItem Q

If lstVars.ListCount > 0 Then
    If Q <= lstVars.ListCount - 1 Then
        lstVars.ListIndex = Q
        lstValues.ListIndex = Q
    Else
        lstVars.ListIndex = lstVars.ListCount - 1
        lstValues.ListIndex = lstVars.ListCount - 1
    End If
Else
    lstVars.Clear
    lstValues.Clear
    cmdRemove.Enabled = False
    cmdEdit.Enabled = False
    lstValues.Enabled = False
    lstVars.Enabled = False
End If

End Sub

Public Sub WEEdit(ByVal Index As Long, ByVal sName As String, ByVal sValue As String)
If sName <> "" Then
    If lstVars.List(Index - 1) <> sName Then lstVars.List(Index - 1) = sName
End If
If sValue <> "" Then
    If lstVars.List(Index - 1) <> sName Then lstVars.List(Index - 1) = sName '????? measure which
End If
End Sub

Public Sub WEClear()
lstVars.Clear
lstValues.Clear
cmdRemove.Enabled = False
cmdEdit.Enabled = False
lstValues.Enabled = False
lstVars.Enabled = False
End Sub

Public Property Get CanAutoClose() As Boolean
CanAutoClose = lblX.Visible
End Property

Public Property Let CanAutoClose(ByVal vNewValue As Boolean)
lblX.Visible = vNewValue
End Property
