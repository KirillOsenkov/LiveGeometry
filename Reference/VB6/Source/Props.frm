VERSION 5.00
Begin VB.Form frmFigureProps 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Figure properties"
   ClientHeight    =   5160
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   4308
   HelpContextID   =   40
   Icon            =   "Props.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   430
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   359
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtEquation 
      Height          =   372
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   480
      Width           =   3012
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   9
      Top             =   4680
      Width           =   735
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   10
      Top             =   4680
      Width           =   975
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Top             =   120
      Width           =   3012
   End
   Begin VB.Frame fraAppearance 
      Caption         =   "Appearance"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3600
      Left            =   120
      TabIndex        =   13
      Top             =   972
      Width           =   4095
      Begin DG.ctlColorBox csbFillColor 
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   255
         _ExtentX        =   445
         _ExtentY        =   445
      End
      Begin DG.ctlColorBox csbForeColor 
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   255
         _ExtentX        =   445
         _ExtentY        =   445
      End
      Begin VB.ComboBox cmbFillStyle 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "Props.frx":030A
         Left            =   240
         List            =   "Props.frx":0326
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1920
         Width           =   3600
      End
      Begin VB.VScrollBar vsbDrawWidth 
         Height          =   330
         Left            =   2520
         Max             =   16
         Min             =   1
         TabIndex        =   5
         Top             =   1200
         Value           =   1
         Width           =   198
      End
      Begin VB.TextBox txtDrawWidth 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2160
         TabIndex        =   4
         Top             =   1200
         Width           =   372
      End
      Begin VB.CheckBox chkVisible 
         Caption         =   "Visible"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   240
         TabIndex        =   8
         Top             =   3120
         Width           =   1215
      End
      Begin VB.ComboBox cmbDrawStyle 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "Props.frx":0343
         Left            =   240
         List            =   "Props.frx":0356
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1200
         Width           =   1545
      End
      Begin VB.ComboBox cmbDrawMode 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "Props.frx":0396
         Left            =   240
         List            =   "Props.frx":03A0
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   2640
         Width           =   3600
      End
      Begin VB.Label lblFillColor 
         BackStyle       =   0  'Transparent
         Caption         =   "Fill color"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   600
         TabIndex        =   19
         Top             =   600
         Width           =   1650
      End
      Begin VB.Label lblFillStyle 
         BackStyle       =   0  'Transparent
         Caption         =   "Fill style"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         TabIndex        =   18
         Top             =   1680
         Width           =   1650
      End
      Begin VB.Label lblDrawMode 
         BackStyle       =   0  'Transparent
         Caption         =   "Drawmode"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         TabIndex        =   17
         Top             =   2400
         Width           =   1650
      End
      Begin VB.Label lblDrawStyle 
         BackStyle       =   0  'Transparent
         Caption         =   "Draw style"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         TabIndex        =   16
         Top             =   960
         Width           =   1650
      End
      Begin VB.Label lblDrawWidth 
         BackStyle       =   0  'Transparent
         Caption         =   "Draw width"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2160
         TabIndex        =   15
         Top             =   960
         Width           =   1410
      End
      Begin VB.Label lblForeColor 
         BackStyle       =   0  'Transparent
         Caption         =   "Forecolor"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   600
         TabIndex        =   14
         Top             =   360
         Width           =   1650
      End
   End
   Begin VB.CommandButton cmdEquation 
      Enabled         =   0   'False
      Height          =   372
      Left            =   1200
      TabIndex        =   20
      Top             =   480
      Width           =   3012
   End
   Begin VB.Label lblFigureEquation 
      Caption         =   "Equation"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   120
      TabIndex        =   12
      Top             =   480
      Width           =   1092
   End
   Begin VB.Label lblFigureName 
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   972
   End
End
Attribute VB_Name = "frmFigureProps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Dim unlCancel As Boolean
Dim pAction As Action
Dim DrawModeArray
Dim AntiDrawModeArray

Private Sub cmdCancel_Click()
unlCancel = True
Unload Me
End Sub

Private Sub cmdEquation_Click()
If IsLine(ActiveFigure) Then
'    R = GetLineCoordinatesAbsolute(ActiveFigure)
'    cmdEquation.Caption = GetLineEquationText(R)
    frmAnLine.EditingExisting = True
    frmAnLine.Show vbModal
    FillEquations
    Exit Sub
End If

If IsCircle(ActiveFigure) Then
'    CCent = GetCircleCenter(ActiveFigure)
'    Rad = GetCircleRadius(ActiveFigure)
'    cmdEquation = GetCircleEquationText(CCent, Rad, True)
    'If Figures(ActiveFigure).FigureType = dsCircle_ArcCenterAndRadiusAndTwoPoints Then lblEquation = lblEquation & ";  angle from " & ToDegrees * Figures(ActiveFigure).AuxInfo(2) & "° to " & ToDegrees * Figures(ActiveFigure).AuxInfo(3) & "°"
    frmAnCircle.EditingExisting = True
    frmAnCircle.Show vbModal
    FillEquations
    Exit Sub
End If

End Sub

Private Sub cmdOK_Click()
Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then unlCancel = True: Unload Me
If KeyCode = vbKeyReturn Then Unload Me
End Sub

Private Sub Form_Load()
Dim Z As Long

If Not IsFigure(ActiveFigure) Then unlCancel = True: Unload Me: Exit Sub
ReDim pAction.sFigure(1 To 1)
pAction.sFigure(1) = Figures(ActiveFigure)
pAction.pFigure = ActiveFigure

FillDialogStrings
PrepareControls

csbFillColor.Color = Figures(ActiveFigure).FillColor

DrawModeArray = Array(13, 9)
AntiDrawModeArray = Array(1, 1, 1, 1, 1, 1, 1, 1, 2, 1, 1, 1, 1, 1, 1, 1)

cmbFillStyle.ListIndex = Figures(ActiveFigure).FillStyle + 1
If Not IsCircle(ActiveFigure) Then 'Or Figures(ActiveFigure).FigureType = dsCircle_ArcCenterAndRadiusAndTwoPoints
    cmbFillStyle.Enabled = False
    csbFillColor.Enabled = False
    lblFillColor.Enabled = False
    lblFillStyle.Enabled = False
    cmbFillStyle.BackColor = vbButtonFace
End If

'lblDrawStyle.Visible = Figures(ActiveFigure).DrawWidth = 1
'cmbDrawStyle.Visible = lblDrawStyle.Visible
cmbDrawStyle.Enabled = Figures(ActiveFigure).DrawWidth = 1
cmbDrawStyle.BackColor = IIf(cmbDrawStyle.Enabled, vbWindowBackground, vbButtonFace)
lblDrawStyle.Enabled = cmbDrawStyle.Enabled

txtName.Text = Figures(ActiveFigure).Name
'cmbDrawMode.ListIndex = Format(Figures(ActiveFigure).DrawMode) - 1
cmbDrawMode.ListIndex = AntiDrawModeArray(Figures(ActiveFigure).DrawMode) - 1
'cmbDrawMode.ListIndex = Figures(ActiveFigure).DrawMode - 1
cmbDrawStyle.ListIndex = Figures(ActiveFigure).DrawStyle
txtDrawWidth.Text = Format(Figures(ActiveFigure).DrawWidth)
vsbDrawWidth.Value = MaxDrawWidth + 1 - Figures(ActiveFigure).DrawWidth
chkVisible.Caption = GetString(ResVisible)
chkVisible.Value = 1 + Figures(ActiveFigure).Hide
csbForeColor.Color = Figures(ActiveFigure).ForeColor
FormMain.Enabled = False
unlCancel = False

Show

FillEquations

'ShadowControl cmdCancel
'ShadowControl cmdOK
'ShadowControl lblEquation

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode <> vbFormCode Then unlCancel = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
FormMain.Enabled = True
FormMain.SetFocus
If unlCancel Then Exit Sub
HideFigure ActiveFigure
If Not IsNumeric(txtDrawWidth.Text) Then Cancel = 1: FormMain.Enabled = False: MsgBox GetString(ResMsgInputInteger) + ".", vbOKOnly + vbExclamation: Exit Sub
If txtName.Text = "" Then Cancel = 1: FormMain.Enabled = False: MsgBox GetString(ResMsgNoName), vbExclamation: Exit Sub

If Val(txtDrawWidth.Text) < 1 Or Val(txtDrawWidth.Text) > MaxDrawWidth Then
    Cancel = 1
    FormMain.Enabled = False
    MsgBox GetString(ResDrawWidth) & " " & GetString(ResMsgExpressedBy) & " " & GetString(ResMsgFrom) & " 1 " & GetString(ResMsgTo) & " " & MaxDrawWidth & ".", vbInformation
    txtDrawWidth.SetFocus
Else
    Figures(ActiveFigure).DrawWidth = Int(Val(txtDrawWidth.Text))
End If

Figures(ActiveFigure).DrawMode = DrawModeArray(cmbDrawMode.ListIndex + 1)

Figures(ActiveFigure).DrawStyle = cmbDrawStyle.ListIndex
Figures(ActiveFigure).ForeColor = csbForeColor.Color
If Figures(ActiveFigure).Name <> txtName.Text Then
    'FigureNames.ReplaceItem FigureNames.FindItem(Figures(ActiveFigure).Name), txtName.Text
    Figures(ActiveFigure).Name = txtName.Text
End If
Figures(ActiveFigure).Hide = Not CBool(chkVisible.Value)
Figures(ActiveFigure).FillColor = csbFillColor.Color
Figures(ActiveFigure).FillStyle = cmbFillStyle.ListIndex - 1

If Figures(ActiveFigure).Hide Then HideFigure ActiveFigure
PaperCls
ShowAll
pAction.Type = actChangeAttrFigure
RecordAction pAction
End Sub

Private Sub txtDrawWidth_Change()
If Not Visible Or Not IsNumeric(txtDrawWidth) Then Exit Sub
If Val(txtDrawWidth.Text) <> Int(Val(txtDrawWidth.Text)) Then Exit Sub
If Val(txtDrawWidth.Text) < 1 Or Val(txtDrawWidth.Text) > MaxDrawWidth Then Exit Sub
'cmbDrawStyle.Visible = Val(txtDrawWidth.Text) = 1
'lblDrawStyle.Visible = cmbDrawStyle.Visible
cmbDrawStyle.Enabled = Val(txtDrawWidth.Text) = 1
cmbDrawStyle.BackColor = IIf(cmbDrawStyle.Enabled, vbWindowBackground, vbButtonFace)
lblDrawStyle.Enabled = cmbDrawStyle.Enabled
If txtDrawWidth.Enabled Then
    vsbDrawWidth.Enabled = False
    vsbDrawWidth.Value = MaxDrawWidth + 1 - Int(Val(txtDrawWidth.Text))
    vsbDrawWidth.Enabled = True
    txtDrawWidth.SetFocus
End If
End Sub

Private Sub txtDrawWidth_Validate(Cancel As Boolean)
If Not txtDrawWidth.Enabled Then Exit Sub
If Not IsNumeric(txtDrawWidth.Text) Then GoTo ERR
If Val(txtDrawWidth.Text) <> Int(Val(txtDrawWidth.Text)) Then GoTo ERR
If Val(txtDrawWidth.Text) < 1 Or Val(txtDrawWidth.Text) > MaxDrawWidth Then GoTo ERR
'Figures(ActiveFigure).DrawWidth = Int(Val(txtDrawWidth.Text))
'vsbDrawWidth.Enabled = False
'vsbDrawWidth.Value = Int(Val(txtDrawWidth.Text))
'vsbDrawWidth.Enabled = True
Exit Sub

ERR:
Cancel = 1
FormMain.Enabled = False
MsgBox GetString(ResDrawWidth) + " " + GetString(ResMsgExpressedBy) + " " + GetString(ResMsgFrom) + " 1 " + GetString(ResMsgTo) + " 16.", vbInformation
'txtDrawWidth.SetFocus
End Sub

Private Sub vsbDrawWidth_Change()
If Not Visible Or Not vsbDrawWidth.Enabled Then Exit Sub
If Not IsNumeric(txtDrawWidth) Then Exit Sub
If Val(txtDrawWidth.Text) <> Int(Val(txtDrawWidth.Text)) Then Exit Sub
If vsbDrawWidth.Enabled Then
    txtDrawWidth.Enabled = False
    txtDrawWidth.Text = MaxDrawWidth + 1 - vsbDrawWidth.Value
    'txtDrawWidth.SelStart = 0
    'txtDrawWidth.SelLength = Len(txtDrawWidth)
    txtDrawWidth.Enabled = True
    txtDrawWidth.SetFocus
End If
End Sub

Public Sub FillEquations()
Dim Eq As LineGeneralEquation
Dim R As TwoPoints
Dim S As String
Dim CCent As OnePoint
Dim Rad As Double

cmdEquation.Enabled = False

If IsLine(ActiveFigure) Then
    R = GetLineCoordinatesAbsolute(ActiveFigure)
    cmdEquation.Caption = GetLineEquationText(R)
    cmdEquation.Enabled = IsLineAnalytic(ActiveFigure)
End If

If IsCircle(ActiveFigure) Then
    CCent = GetCircleCenter(ActiveFigure)
    Rad = GetCircleRadius(ActiveFigure)
    cmdEquation.Caption = GetCircleEquationText(CCent, Rad, True)
    'If Figures(ActiveFigure).FigureType = dsCircle_ArcCenterAndRadiusAndTwoPoints Then lblEquation = lblEquation & ";  angle from " & ToDegrees * Figures(ActiveFigure).AuxInfo(2) & "° to " & ToDegrees * Figures(ActiveFigure).AuxInfo(3) & "°"
    cmdEquation.Enabled = IsCircleAnalytic(ActiveFigure)
End If

cmdEquation.Visible = cmdEquation.Enabled
txtEquation.Visible = Not cmdEquation.Visible
txtEquation.Text = cmdEquation.Caption
End Sub

Public Sub FillDialogStrings()
Caption = Figures(ActiveFigure).Name & " " & GetString(ResPropsTitle)
lblDrawMode = GetString(ResDrawMode)
lblDrawStyle = GetString(ResDrawStyle)
lblDrawWidth = GetString(ResDrawWidth)
lblFigureName = GetString(ResName)
lblFigureEquation = GetString(ResEquation)
lblFillColor = GetString(ResFillColor) '????? GetString(ResForeColor) & " - " & GetString(ResFill)
lblFillStyle = GetString(ResFill)
fraAppearance.Caption = GetString(ResAppearance)
lblForeColor = GetString(ResForeColor)
cmdCancel.Caption = GetString(ResCancel)
End Sub

Public Sub PrepareControls()
Dim Z As Long

For Z = 0 To cmbFillStyle.ListCount - 1
    cmbFillStyle.List(Z) = GetString(ResFillStyleBase + 2 * Z)
Next

For Z = 0 To cmbDrawMode.ListCount - 1
    cmbDrawMode.List(Z) = GetString(ResDrawModeSolid + 2 * Z)
Next
End Sub
