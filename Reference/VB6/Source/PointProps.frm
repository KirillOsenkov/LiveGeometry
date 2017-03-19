VERSION 5.00
Begin VB.Form frmPointProps 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Point properties"
   ClientHeight    =   4260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5325
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   39
   Icon            =   "PointProps.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   5325
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkEnabled 
      Caption         =   "Disabled"
      Height          =   255
      Left            =   2640
      TabIndex        =   19
      Top             =   3120
      Width           =   2535
   End
   Begin VB.CheckBox chkHide 
      Caption         =   "Visible"
      Height          =   255
      Left            =   2640
      TabIndex        =   20
      Top             =   3360
      Width           =   2535
   End
   Begin VB.PictureBox picCoord 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   2520
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   177
      TabIndex        =   31
      Top             =   120
      Width           =   2655
      Begin VB.TextBox txtYCoord 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   2
         Text            =   "0"
         Top             =   240
         Width           =   2295
      End
      Begin VB.TextBox txtXCoord 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   1
         Text            =   "0"
         Top             =   0
         Width           =   2295
      End
      Begin VB.Label lblYCoord 
         Alignment       =   1  'Right Justify
         Caption         =   "Y = "
         Height          =   255
         Left            =   0
         TabIndex        =   33
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblXCoord 
         Alignment       =   1  'Right Justify
         Caption         =   "X = "
         Height          =   255
         Left            =   0
         TabIndex        =   32
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.TextBox txtName 
      Height          =   315
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   3360
      TabIndex        =   21
      Top             =   3720
      Width           =   735
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4200
      TabIndex        =   22
      Top             =   3720
      Width           =   975
   End
   Begin VB.Frame fraGeneral 
      Caption         =   "Appearance"
      Height          =   3255
      Left            =   120
      TabIndex        =   24
      Top             =   840
      Width           =   2295
      Begin DG.ctlColorBox csbNameColor 
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1800
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
      End
      Begin DG.ctlColorBox csbFillColor 
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
      End
      Begin DG.ctlColorBox csbForeColor 
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
      End
      Begin VB.CheckBox chkShowCoordinates 
         Caption         =   "Show coordinates"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1440
         Width           =   2055
      End
      Begin VB.CheckBox chkFill 
         Caption         =   "Fill:"
         Height          =   255
         Left            =   480
         TabIndex        =   4
         Top             =   600
         Width           =   1695
      End
      Begin VB.VScrollBar vsbSize 
         Height          =   315
         Left            =   480
         Max             =   30
         Min             =   2
         TabIndex        =   11
         Top             =   2760
         Value           =   2
         Width           =   240
      End
      Begin VB.CheckBox chkShowName 
         Caption         =   "Show name"
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Top             =   1080
         Width           =   2025
      End
      Begin VB.ComboBox cmbShape 
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   300
         ItemData        =   "PointProps.frx":014A
         Left            =   120
         List            =   "PointProps.frx":0154
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   2280
         Width           =   615
      End
      Begin VB.TextBox txtSize 
         ForeColor       =   &H80000012&
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Text            =   "12"
         Top             =   2760
         Width           =   360
      End
      Begin VB.Label lblNameColor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name color"
         Height          =   210
         Left            =   480
         TabIndex        =   28
         Top             =   1800
         Width           =   1410
      End
      Begin VB.Label lblShape 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Shape"
         Height          =   210
         Left            =   840
         TabIndex        =   27
         Top             =   2280
         Width           =   465
      End
      Begin VB.Label lblForeColor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Color"
         Height          =   210
         Left            =   480
         TabIndex        =   26
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lblSize 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Size"
         Height          =   210
         Left            =   840
         TabIndex        =   25
         Top             =   2760
         Width           =   315
      End
   End
   Begin VB.CheckBox chkLocus 
      Caption         =   "Create locus"
      Height          =   270
      Left            =   2880
      TabIndex        =   13
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Frame fraLocusProp 
      Caption         =   "Locus"
      Enabled         =   0   'False
      Height          =   2085
      Left            =   2640
      TabIndex        =   12
      Top             =   840
      Width           =   2535
      Begin DG.ctlColorBox csbLocusColor 
         Height          =   255
         Left            =   2025
         TabIndex        =   14
         Top             =   600
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
      End
      Begin VB.ComboBox cmbLocusType 
         Height          =   330
         ItemData        =   "PointProps.frx":015E
         Left            =   1665
         List            =   "PointProps.frx":0168
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   1320
         Width           =   615
      End
      Begin VB.VScrollBar vsbDrawWidth 
         Height          =   315
         Left            =   2040
         Max             =   16
         Min             =   1
         TabIndex        =   16
         Top             =   960
         Value           =   1
         Width           =   240
      End
      Begin VB.TextBox txtLocusWidth 
         ForeColor       =   &H80000012&
         Height          =   315
         Left            =   1680
         TabIndex        =   15
         Text            =   "1"
         Top             =   960
         Width           =   375
      End
      Begin VB.CommandButton cmdDeleteLocus 
         Caption         =   "Delete locus"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   1680
         Width           =   2040
      End
      Begin VB.Label lblLocusType 
         Caption         =   "Locus type"
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   1365
         Width           =   1335
      End
      Begin VB.Label lblLocusWidth 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Draw width"
         Height          =   210
         Left            =   240
         TabIndex        =   29
         Top             =   1005
         Width           =   1335
      End
      Begin VB.Label lblLocusColor 
         BackStyle       =   0  'Transparent
         Caption         =   "Locus color"
         Height          =   210
         Left            =   240
         TabIndex        =   23
         Top             =   645
         Width           =   1335
      End
   End
   Begin VB.Label lblPointType 
      Caption         =   "Type: Point"
      Height          =   255
      Left            =   120
      TabIndex        =   34
      Top             =   480
      Width           =   2295
   End
   Begin VB.Label lblPointName 
      Caption         =   "Name"
      Height          =   255
      Left            =   120
      TabIndex        =   30
      Top             =   150
      Width           =   615
   End
End
Attribute VB_Name = "frmPointProps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim unlCancel As Boolean
Dim CoordChanged As Boolean
Dim pAction As Action, pRenameAction As Action, pActionLocus As Action
Dim nLocusColor As Long

Private Sub chkFill_Click()
csbFillColor.Enabled = -chkFill.Value
End Sub

Private Sub chkHide_Click()
chkEnabled.Value = chkHide.Value
End Sub

Private Sub chkLocus_Click()
fraLocusProp.Enabled = -chkLocus.Value
txtLocusWidth.Enabled = fraLocusProp.Enabled
cmdDeleteLocus.Enabled = fraLocusProp.Enabled
lblLocusColor.Enabled = fraLocusProp.Enabled
lblLocusWidth.Enabled = fraLocusProp.Enabled
lblLocusType.Enabled = fraLocusProp.Enabled
vsbDrawWidth.Enabled = fraLocusProp.Enabled
cmbLocusType.Enabled = fraLocusProp.Enabled
txtLocusWidth.BackColor = IIf(fraLocusProp.Enabled, vbWindowBackground, vbButtonFace)
cmbLocusType.BackColor = txtLocusWidth.BackColor
csbLocusColor.Enabled = fraLocusProp.Enabled
'If fraLocusProp.Enabled = False And csbLocusColor.Color <> vbButtonFace Then nLocusColor = csbLocusColor.BackColor
'csbLocusColor.Color = IIf(fraLocusProp.Enabled, nLocusColor, vbButtonFace)
End Sub

Private Sub chkShowName_Click()
csbNameColor.Enabled = -chkShowName.Value
End Sub

Private Sub csbNameColor_ColorChanged(ByVal NewColor As Long, ByVal OldColor As Long)
txtName.ForeColor = NewColor
End Sub

Private Sub cmdDeleteLocus_Click()
' Action is recorded properly here...
If BasePoint(ActivePoint).Locus <> 0 Then EraseLocus BasePoint(ActivePoint).Locus
End Sub

'======================================================

Private Sub cmdOK_Click()
Unload Me
End Sub

Private Sub cmdCancel_Click()
unlCancel = True
Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then unlCancel = True: Unload Me
If KeyCode = vbKeyReturn Then Unload Me
End Sub

'======================================================

Private Sub Form_Load()
If Not IsPoint(ActivePoint) Then unlCancel = True: Unload Me: Exit Sub
CoordChanged = False

pAction.Type = actChangeAttrPoint
ReDim pAction.sPoint(1 To 1)
pAction.sPoint(1) = BasePoint(ActivePoint)
pAction.pPoint = ActivePoint

'pRenameAction = PrepareGenericAction(ResUndoRenamePoint)
'pRenameAction.Type = actGenericAction
'MakeStructureSnapshot pRenameAction

Caption = Replace(GetString(ResPropertiesOfAPoint), "%1", BasePoint(ActivePoint).Name) 'GetString(ResFigureBase + 2 * BasePoint(ActivePoint).Type) & " " & BasePoint(ActivePoint).Name & " " & GetString(ResPropsTitle)
lblPointType.Caption = GetString(ResFigureBase + 2 * BasePoint(ActivePoint).Type)

FillDialogStrings
FormMain.Enabled = False

txtName.BackColor = nPaperColor1
txtName.ForeColor = BasePoint(ActivePoint).NameColor

If BasePoint(ActivePoint).Type = dsPoint Then
    txtXCoord.Locked = False
    txtYCoord.Locked = False
    txtXCoord.BackColor = vbWindowBackground
    txtYCoord.BackColor = vbWindowBackground
End If
If BasePoint(ActivePoint).Type = dsAnPoint Then
    txtXCoord.Text = Figures(BasePoint(ActivePoint).ParentFigure).XS
    txtYCoord.Text = Figures(BasePoint(ActivePoint).ParentFigure).YS
Else
    txtXCoord.Text = Format(BasePoint(ActivePoint).X, setFormatDistance)
    txtYCoord.Text = Format(BasePoint(ActivePoint).Y, setFormatDistance)
End If

txtName.Text = BasePoint(ActivePoint).Name
txtSize.Text = BasePoint(ActivePoint).PhysicalWidth
vsbSize.Value = MaxPointSize + MinPointSize - BasePoint(ActivePoint).PhysicalWidth
vsbDrawWidth.Value = MaxDrawWidth + 1 - txtLocusWidth

csbFillColor.Color = BasePoint(ActivePoint).FillColor
csbForeColor.Color = BasePoint(ActivePoint).ForeColor
csbNameColor.Color = BasePoint(ActivePoint).NameColor
csbNameColor.Enabled = BasePoint(ActivePoint).ShowName

cmbShape.ListIndex = Sgn(BasePoint(ActivePoint).Shape - 1)
chkHide.Value = -(Not BasePoint(ActivePoint).Hide)
chkEnabled.Value = -(BasePoint(ActivePoint).Enabled)
chkShowName.Value = -BasePoint(ActivePoint).ShowName
chkShowCoordinates.Value = -BasePoint(ActivePoint).ShowCoordinates
chkFill.Value = 1 - BasePoint(ActivePoint).FillStyle
csbFillColor.Enabled = -chkFill.Value
cmbLocusType.ListIndex = 0

If BasePoint(ActivePoint).Locus <> 0 Then
    txtLocusWidth.Text = Locuses(BasePoint(ActivePoint).Locus).DrawWidth
    vsbDrawWidth.Value = MaxDrawWidth + 1 - Val(txtLocusWidth)
    chkLocus.Value = -(Locuses(BasePoint(ActivePoint).Locus).Enabled)
    fraLocusProp.Enabled = -chkLocus.Value
    cmbLocusType.ListIndex = Locuses(BasePoint(ActivePoint).Locus).Type
    If Locuses(BasePoint(ActivePoint).Locus).Dynamic Then
        cmdDeleteLocus.Enabled = False
        chkLocus.Enabled = False
    End If
    
    pActionLocus.Type = actChangeAttrLocus
    pActionLocus.pLocus = BasePoint(ActivePoint).Locus
    pActionLocus.pPoint = ActivePoint
    ReDim pActionLocus.sLocus(1 To 1)
    pActionLocus.sLocus(1) = Locuses(BasePoint(ActivePoint).Locus)
Else
    txtLocusWidth.Text = 1
    vsbDrawWidth.Value = MaxDrawWidth + 1 - Val(txtLocusWidth)
    cmbLocusType.ListIndex = 0
    nLocusColor = setdefcolLocus
    csbLocusColor.Enabled = False
End If

If fraLocusProp.Enabled Then csbLocusColor.Color = Locuses(BasePoint(ActivePoint).Locus).ForeColor
txtLocusWidth.Enabled = fraLocusProp.Enabled
cmdDeleteLocus.Enabled = fraLocusProp.Enabled
lblLocusColor.Enabled = fraLocusProp.Enabled
lblLocusWidth.Enabled = fraLocusProp.Enabled
lblLocusType.Enabled = fraLocusProp.Enabled
vsbDrawWidth.Enabled = fraLocusProp.Enabled
cmbLocusType.Enabled = fraLocusProp.Enabled
txtLocusWidth.BackColor = IIf(fraLocusProp.Enabled, vbWindowBackground, vbButtonFace)
cmbLocusType.BackColor = txtLocusWidth.BackColor

'ShadowControl cmdCancel
'ShadowControl cmdOK

unlCancel = False
Visible = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode <> vbFormCode Then unlCancel = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Local Error GoTo EH:
Dim lpSize As Size, TempPointNumber As Long, tempBoolean As Boolean

If unlCancel Then
    FormMain.Enabled = True
    FormMain.SetFocus
    Exit Sub
End If

If txtName.Text = "" Then
    Cancel = 1
    MsgBox GetString(ResMsgNoName), vbExclamation
    Exit Sub
End If

If txtName.Text <> BasePoint(ActivePoint).Name Then
    RenamePoint ActivePoint, txtName.Text
    'TempPointNumber = GetPointByName(txtName.Text)
    'If IsPoint(TempPointNumber) And TempPointNumber <> ActivePoint Then Cancel = True: MsgBox txtName.Text & GetString(ResMsgObjectAlreadyExists): txtName.SetFocus: Exit Sub
End If

ShowPoint Paper.hDC, ActivePoint, True

BasePoint(ActivePoint).PhysicalWidth = txtSize.Text
BasePoint(ActivePoint).Width = BasePoint(ActivePoint).PhysicalWidth
ToLogicalLength BasePoint(ActivePoint).Width

Dim NewX As Double, NewY As Double
If CoordChanged Then
    NewX = BasePoint(ActivePoint).X
    NewY = BasePoint(ActivePoint).Y
    If IsNumeric(txtXCoord.Text) Then
        NewX = CDbl(txtXCoord.Text)
    ElseIf Val(txtXCoord.Text) <> 0 Then
        NewX = Val(txtXCoord.Text)
    End If
    If IsNumeric(txtYCoord.Text) Then
        NewY = CDbl(txtYCoord.Text)
    ElseIf Val(txtYCoord.Text) <> 0 Then
        NewY = Val(txtYCoord.Text)
    End If
    MovePoint ActivePoint, NewX, NewY
    RecalcAllAuxInfo
End If

'If BasePoint(ActivePoint).Name <> txtName.Text Then
'    RenamePoint ActivePoint, txtName.Text
'    RecordAction pRenameAction
'Else

RecordAction pAction

'End If

BasePoint(ActivePoint).Hide = Not (-chkHide.Value)
BasePoint(ActivePoint).ForeColor = csbForeColor.Color
BasePoint(ActivePoint).FillStyle = 1 - chkFill.Value
BasePoint(ActivePoint).FillColor = csbFillColor.Color
BasePoint(ActivePoint).NameColor = csbNameColor.Color
BasePoint(ActivePoint).Shape = cmbShape.ListIndex * 2 + 1
BasePoint(ActivePoint).ShowName = -chkShowName.Value
BasePoint(ActivePoint).ShowCoordinates = -chkShowCoordinates.Value
BasePoint(ActivePoint).LabelLength = Len(txtName.Text)
BasePoint(PointCount).LabelOffsetX = 0
BasePoint(PointCount).LabelOffsetY = -Val(txtSize.Text) \ 2 + 1
BasePoint(ActivePoint).LabelWidth = Paper.TextWidth(txtName.Text)
BasePoint(ActivePoint).LabelHeight = Paper.TextHeight(txtName.Text)
If Not BasePoint(ActivePoint).ShowCoordinates Then BasePoint(ActivePoint).LabelLength = Len(txtName.Text)
BasePoint(ActivePoint).Enabled = (-chkEnabled.Value)

If BasePoint(ActivePoint).Locus <> 0 Then
    tempBoolean = True
    tempBoolean = tempBoolean And (CBool(chkLocus.Value = 1) = Locuses(BasePoint(ActivePoint).Locus).Enabled)
    tempBoolean = tempBoolean And (Locuses(BasePoint(ActivePoint).Locus).DrawWidth = Int(Val(txtLocusWidth.Text)))
    tempBoolean = tempBoolean And (Locuses(BasePoint(ActivePoint).Locus).ForeColor = csbLocusColor.Color)
    tempBoolean = tempBoolean And (Locuses(BasePoint(ActivePoint).Locus).Type = cmbLocusType.ListIndex)
    If Not tempBoolean Then
        RecordAction pActionLocus
        If chkLocus.Value = 1 Then
            If Locuses(BasePoint(ActivePoint).Locus).DrawWidth > Val(txtLocusWidth.Text) Then ShowLocus Paper.hDC, BasePoint(ActivePoint).Locus, False, True
            Locuses(BasePoint(ActivePoint).Locus).DrawWidth = Int(Val(txtLocusWidth.Text))
            Locuses(BasePoint(ActivePoint).Locus).ForeColor = csbLocusColor.Color
            Locuses(BasePoint(ActivePoint).Locus).Enabled = True
            Locuses(BasePoint(ActivePoint).Locus).Type = cmbLocusType.ListIndex
        Else
            Locuses(BasePoint(ActivePoint).Locus).Enabled = False
        End If
    End If
Else
    If chkLocus.Value = 1 Then
        AddLocus ActivePoint
        Locuses(BasePoint(ActivePoint).Locus).DrawWidth = Int(Val(txtLocusWidth.Text))
        Locuses(BasePoint(ActivePoint).Locus).ForeColor = csbLocusColor.Color
        Locuses(BasePoint(ActivePoint).Locus).Enabled = True
        Locuses(BasePoint(ActivePoint).Locus).Type = cmbLocusType.ListIndex
        pActionLocus.pPoint = ActivePoint
        pActionLocus.Type = actAddLocus
        pActionLocus.pLocus = LocusCount
        RecordAction pActionLocus
    End If
End If

If BasePoint(ActivePoint).ShowName And BasePoint(ActivePoint).ShowCoordinates Then
    Dim sStr As String
    sStr = BasePoint(ActivePoint).Name
    If BasePoint(ActivePoint).ShowCoordinates Then sStr = sStr & " (" & Format(BasePoint(ActivePoint).X, setFormatNumber) & "; " & Format(BasePoint(ActivePoint).Y, setFormatNumber) & ")"
    BasePoint(ActivePoint).LabelLength = Len(sStr)
    BasePoint(ActivePoint).LabelWidth = Paper.TextWidth(sStr)
    BasePoint(ActivePoint).LabelHeight = Paper.TextHeight(sStr)
End If

FormMain.Enabled = True
FormMain.SetFocus

PaperCls
ShowAll
Exit Sub

EH:
MsgBox GetString(ResError) & ": " & ERR.Description, vbExclamation
End Sub

Private Sub txtLocusWidth_Change()
If Not Visible Or Not IsNumeric(txtLocusWidth.Text) Then Exit Sub
If Val(txtLocusWidth.Text) <> Int(Val(txtLocusWidth.Text)) Then Exit Sub
If Val(txtLocusWidth.Text) < 1 Or Val(txtLocusWidth.Text) > MaxDrawWidth Then Exit Sub
If txtLocusWidth.Enabled Then
    vsbDrawWidth.Enabled = False
    vsbDrawWidth.Value = MaxDrawWidth + 1 - Val(txtLocusWidth)
    vsbDrawWidth.Enabled = True
End If
End Sub

Private Sub txtLocusWidth_Validate(Cancel As Boolean)
If Not IsNumeric(txtLocusWidth.Text) Then GoTo ERR
If Val(txtLocusWidth.Text) <> Int(Val(txtLocusWidth.Text)) Then GoTo ERR
If Val(txtLocusWidth.Text) < 1 Or Val(txtLocusWidth.Text) > MaxDrawWidth Then GoTo ERR
Exit Sub

ERR:
Cancel = 1
MsgBox GetString(ResDrawWidth) & " " & GetString(ResMsgExpressedBy) & " " & GetString(ResMsgFrom) & " 1 " & GetString(ResMsgTo) & " " & MaxDrawWidth & ".", vbInformation
End Sub

Private Sub txtSize_Change()
If Not Visible Or Not IsNumeric(txtSize.Text) Then Exit Sub
If Val(txtSize.Text) <> Int(Val(txtSize.Text)) Then Exit Sub
If txtSize < MinPointSize Or txtSize > MaxPointSize Then Exit Sub
If txtSize.Enabled Then
    vsbSize.Enabled = False
    vsbSize.Value = MaxPointSize + MinPointSize - Val(txtSize)
    vsbSize.Enabled = True
End If
End Sub

Private Sub txtSize_Validate(Cancel As Boolean)
If Not IsNumeric(txtSize) Then GoTo ERR
If Val(txtSize.Text) <> Int(Val(txtSize.Text)) Then GoTo ERR
If Val(txtSize.Text) < MinPointSize Or Val(txtSize.Text) > MaxPointSize Then GoTo ERR
Exit Sub

ERR:
Cancel = 1
MsgBox GetString(ResSize) & " " & GetString(ResMsgExpressedBy) & " " & GetString(ResMsgFrom) & " " & MinPointSize & " " & GetString(ResMsgTo) & " " & MaxPointSize & ".", vbInformation
End Sub

Private Sub txtXCoord_Change()
CoordChanged = True
End Sub

Private Sub txtYCoord_Change()
CoordChanged = True
End Sub

Private Sub vsbDrawWidth_Change()
If Not Visible Or Not vsbDrawWidth.Enabled Then Exit Sub
txtLocusWidth.Enabled = False
txtLocusWidth = MaxDrawWidth + 1 - vsbDrawWidth.Value
txtLocusWidth.Enabled = True
txtLocusWidth.SelStart = 0
txtLocusWidth.SelLength = Len(txtLocusWidth)
'txtLocusWidth.SetFocus
End Sub

Private Sub vsbSize_Change()
If Not Visible Or Not vsbSize.Enabled Then Exit Sub
If vsbSize.Value < MinPointSize Or vsbSize.Value > MaxPointSize Then Exit Sub
txtSize.Enabled = False
txtSize = MaxPointSize + MinPointSize - vsbSize.Value
txtSize.Enabled = True
txtSize.SelStart = 0
txtSize.SelLength = Len(txtSize)
'txtSize.SetFocus
End Sub

Private Sub FillDialogStrings()
chkFill.Caption = GetString(ResFill)
chkHide.Caption = GetString(ResVisible)
chkEnabled.Caption = GetString(ResDisabled)
chkShowName.Caption = GetString(ResShowName)
chkLocus.Caption = GetString(ResCreateLocus)
chkShowCoordinates.Caption = GetString(ResShowCoordinates)
cmdCancel.Caption = GetString(ResCancel)
cmdDeleteLocus.Caption = GetString(ResDeleteLocus)
fraGeneral.Caption = GetString(ResAppearance)
fraLocusProp.Caption = GetString(ResLocusProps)
lblNameColor.Caption = GetString(ResNameColor)
lblLocusColor.Caption = GetString(ResForeColor)
lblLocusType.Caption = GetString(ResLocusType)
lblShape.Caption = GetString(ResShape)
lblSize.Caption = GetString(ResSize)
lblForeColor.Caption = GetString(ResForeColor)
lblLocusWidth.Caption = GetString(ResDrawWidth)
lblPointName.Caption = GetString(ResName)
End Sub
