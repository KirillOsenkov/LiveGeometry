VERSION 5.00
Begin VB.Form frmStaticProps 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Static graphic properties"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3600
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   68
   Icon            =   "frmStaticProps.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   3600
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   3120
      Width           =   735
   End
   Begin VB.Frame fraAppearance 
      Caption         =   "Appearance"
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
      Begin DG.ctlColorBox csbForeColor 
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
      End
      Begin VB.ComboBox cmbFillStyle 
         ForeColor       =   &H80000012&
         Height          =   330
         ItemData        =   "frmStaticProps.frx":0442
         Left            =   120
         List            =   "frmStaticProps.frx":045B
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1680
         Width           =   3120
      End
      Begin VB.ComboBox cmbDrawMode 
         ForeColor       =   &H80000012&
         Height          =   330
         ItemData        =   "frmStaticProps.frx":0475
         Left            =   120
         List            =   "frmStaticProps.frx":047F
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   2400
         Width           =   3120
      End
      Begin VB.ComboBox cmbDrawStyle 
         ForeColor       =   &H80000012&
         Height          =   330
         ItemData        =   "frmStaticProps.frx":0489
         Left            =   120
         List            =   "frmStaticProps.frx":049C
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   960
         Width           =   1320
      End
      Begin VB.TextBox txtDrawWidth 
         ForeColor       =   &H80000012&
         Height          =   330
         Left            =   1800
         TabIndex        =   4
         Top             =   960
         Width           =   375
      End
      Begin VB.VScrollBar vsbDrawWidth 
         Height          =   330
         Left            =   2160
         Max             =   16
         Min             =   1
         TabIndex        =   3
         Top             =   960
         Value           =   1
         Width           =   195
      End
      Begin VB.Label lblFillStyle 
         BackStyle       =   0  'Transparent
         Caption         =   "Fill style"
         Height          =   225
         Left            =   120
         TabIndex        =   12
         Top             =   1440
         Width           =   1770
      End
      Begin VB.Label lblForeColor 
         BackStyle       =   0  'Transparent
         Caption         =   "Forecolor"
         Height          =   225
         Left            =   480
         TabIndex        =   10
         Top             =   360
         Width           =   1650
      End
      Begin VB.Label lblDrawWidth 
         BackStyle       =   0  'Transparent
         Caption         =   "Draw width"
         Height          =   225
         Left            =   1800
         TabIndex        =   9
         Top             =   720
         Width           =   1410
      End
      Begin VB.Label lblDrawStyle 
         BackStyle       =   0  'Transparent
         Caption         =   "Draw style"
         Height          =   225
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   1650
      End
      Begin VB.Label lblDrawMode 
         BackStyle       =   0  'Transparent
         Caption         =   "Drawmode"
         Height          =   225
         Left            =   120
         TabIndex        =   7
         Top             =   2160
         Width           =   1650
      End
   End
End
Attribute VB_Name = "frmStaticProps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 1
'1 - Black Pen
'2 - Not Merge Pen
'3 - Mask Not Pen
'4 - Not Copy Pen
'5 - Mask Pen Not
'6 - Invert
'7 - Xor Pen
'8 - Not Mask Pen
'9 - Mask Pen
'10 - Not Xor Pen
'11 - Invisible Pen
'12 - Merge Not Pen
'13 - Copy Pen
'14 - Merge Pen Not
'15 - Merge Pen
'16 - White Pen

Option Explicit

Dim DrawModeArray
Dim AntiDrawModeArray
Dim unlCancel As Boolean
Dim pAction As Action

Private Sub cmdCancel_Click()
unlCancel = True
Unload Me
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

If ActiveStatic = 0 Then Exit Sub

ReDim pAction.sSG(1 To 1)
pAction.sSG(1) = StaticGraphics(ActiveStatic)
pAction.pSG = ActiveStatic

Select Case StaticGraphics(ActiveStatic).Type
    Case sgPolygon
        lblDrawStyle.Enabled = False
        lblDrawWidth.Enabled = False
        cmbDrawStyle.Enabled = False
        txtDrawWidth.Enabled = False
        vsbDrawWidth.Enabled = False
        txtDrawWidth.BackColor = vbButtonFace
        cmbDrawStyle.BackColor = vbButtonFace
        
    Case sgBezier
        EnableTextBox cmbFillStyle, False
        lblFillStyle.Enabled = False
    Case sgVector
        EnableTextBox cmbFillStyle, False
        lblFillStyle.Enabled = False
        EnableTextBox txtDrawWidth, False
        EnableTextBox cmbDrawMode, False
        EnableTextBox cmbDrawStyle, False
        EnableTextBox cmbFillStyle, False
        EnableTextBox cmbFillStyle, False
        EnableTextBox cmbFillStyle, False
        EnableTextBox cmbFillStyle, False
        lblDrawWidth.Enabled = False
        vsbDrawWidth.Enabled = False
        lblDrawMode.Enabled = False
        lblDrawStyle.Enabled = False
End Select

FillDialogStrings

If cmbDrawStyle.Enabled Then
    EnableTextBox cmbDrawStyle, StaticGraphics(ActiveStatic).DrawWidth = 1
    lblDrawStyle.Enabled = cmbDrawStyle.Enabled
End If

'lblDrawStyle.Visible = StaticGraphics(ActiveStatic).DrawWidth = 1
'cmbDrawStyle.Enabled = StaticGraphics(ActiveStatic).DrawWidth = 1
'cmbDrawStyle.BackColor = IIf(cmbDrawStyle.Enabled, vbWindowBackground, vbButtonFace)

DrawModeArray = Array(13, 9)
AntiDrawModeArray = Array(1, 1, 1, 1, 1, 1, 1, 1, 2, 1, 1, 1, 1, 1, 1, 1)

cmbDrawMode.ListIndex = AntiDrawModeArray(StaticGraphics(ActiveStatic).DrawMode) - 1
cmbDrawStyle.ListIndex = StaticGraphics(ActiveStatic).DrawStyle
txtDrawWidth.Text = Format(StaticGraphics(ActiveStatic).DrawWidth)
vsbDrawWidth.Value = MaxDrawWidth + 1 - StaticGraphics(ActiveStatic).DrawWidth

If StaticGraphics(ActiveStatic).Type <> sgBezier Then
    csbForeColor.Color = StaticGraphics(ActiveStatic).FillColor
Else
    csbForeColor.Color = StaticGraphics(ActiveStatic).ForeColor
End If
cmbFillStyle.ListIndex = StaticGraphics(ActiveStatic).FillStyle + 1

FormMain.Enabled = False
unlCancel = False

'ShadowControl cmdCancel
'ShadowControl cmdOK

Show
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode <> vbFormCode Then unlCancel = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
FormMain.Enabled = True
FormMain.SetFocus
If unlCancel Then Exit Sub

If Not IsNumeric(txtDrawWidth.Text) Then Cancel = 1: FormMain.Enabled = False: MsgBox GetString(ResMsgInputInteger) + ".", vbOKOnly + vbExclamation: Exit Sub

With StaticGraphics(ActiveStatic)
    If Val(txtDrawWidth.Text) < 1 Or Val(txtDrawWidth.Text) > MaxDrawWidth Then
        Cancel = 1
        FormMain.Enabled = False
        MsgBox GetString(ResDrawWidth) & " " & GetString(ResMsgExpressedBy) & " " & GetString(ResMsgFrom) & " 1 " & GetString(ResMsgTo) & " " & MaxDrawWidth & ".", vbInformation
        txtDrawWidth.SetFocus
    Else
        .DrawWidth = Int(Val(txtDrawWidth.Text))
    End If
    
    .DrawMode = DrawModeArray(cmbDrawMode.ListIndex + 1)
    .DrawStyle = cmbDrawStyle.ListIndex
    .ForeColor = csbForeColor.Color
    .FillColor = csbForeColor.Color
    .FillStyle = cmbFillStyle.ListIndex - 1
End With

PaperCls
ShowAll
pAction.Type = actChangeAttrSG
RecordAction pAction
End Sub

Private Sub txtDrawWidth_Change()
If Not Visible Or Not IsNumeric(txtDrawWidth) Then Exit Sub
If Val(txtDrawWidth.Text) <> Int(Val(txtDrawWidth.Text)) Then Exit Sub
'cmbDrawStyle.Visible = Val(txtDrawWidth.Text) = 1
'lblDrawStyle.Visible = cmbDrawStyle.Visible
cmbDrawStyle.Enabled = Val(txtDrawWidth.Text) = 1
cmbDrawStyle.BackColor = IIf(cmbDrawStyle.Enabled, vbWindowBackground, vbButtonFace)
lblDrawStyle.Enabled = cmbDrawStyle.Enabled
If txtDrawWidth.Enabled Then
    vsbDrawWidth.Enabled = False
    vsbDrawWidth.Value = MaxDrawWidth + 1 - Int(Val(txtDrawWidth.Text))
    vsbDrawWidth.Enabled = True
End If
End Sub

Private Sub txtDrawWidth_Validate(Cancel As Boolean)
If Not txtDrawWidth.Enabled Then Exit Sub
If Not IsNumeric(txtDrawWidth.Text) Then GoTo ERR
If Val(txtDrawWidth.Text) < 1 Or Val(txtDrawWidth.Text) > MaxDrawWidth Then GoTo ERR
If Val(txtDrawWidth.Text) <> Int(Val(txtDrawWidth.Text)) Then GoTo ERR
'StaticGraphics(ActiveStatic).DrawWidth = Int(Val(txtDrawWidth.Text))
'vsbDrawWidth.Value = Int(Val(txtDrawWidth.Text))
Exit Sub

ERR:
Cancel = 1
FormMain.Enabled = False
MsgBox GetString(ResDrawWidth) + " " + GetString(ResMsgExpressedBy) + " " + GetString(ResMsgFrom) + " 1 " + GetString(ResMsgTo) + " 16.", vbInformation
txtDrawWidth.SetFocus
End Sub

Private Sub vsbDrawWidth_Change()
If Not Visible Or Not vsbDrawWidth.Enabled Then Exit Sub
If Not IsNumeric(txtDrawWidth) Then Exit Sub
If Val(txtDrawWidth.Text) <> Int(Val(txtDrawWidth.Text)) Then Exit Sub
txtDrawWidth.Enabled = False
txtDrawWidth.Text = MaxDrawWidth + 1 - vsbDrawWidth.Value
txtDrawWidth.Enabled = True
txtDrawWidth.SelStart = 0
txtDrawWidth.SelLength = Len(txtDrawWidth)
txtDrawWidth.SetFocus
End Sub

Public Sub FillDialogStrings()
Dim Z As Long

lblDrawMode = GetString(ResDrawMode)
lblDrawStyle = GetString(ResDrawStyle)
lblDrawWidth = GetString(ResDrawWidth)
lblForeColor = GetString(ResForeColor)
lblFillStyle = GetString(ResFill)
fraAppearance.Caption = GetString(ResAppearance)
cmdCancel.Caption = GetString(ResCancel)
Caption = GetString(ResPropertiesOfAPolygon + StaticGraphics(ActiveStatic).Type * 2)

For Z = 0 To cmbFillStyle.ListCount - 1
    cmbFillStyle.List(Z) = GetString(ResFillStyleBase + 2 * Z)
Next

For Z = 0 To cmbDrawMode.ListCount - 1
    cmbDrawMode.List(Z) = GetString(ResDrawModeSolid + 2 * Z)
Next
End Sub
