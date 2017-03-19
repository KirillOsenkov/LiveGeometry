VERSION 5.00
Begin VB.Form frmMeasurementProps 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Measurement properties"
   ClientHeight    =   3660
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   3540
   Icon            =   "frmAngleProps.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   3540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   295
      TabIndex        =   8
      Top             =   3165
      Width           =   3540
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
         Left            =   1560
         TabIndex        =   10
         Top             =   0
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
         Left            =   2400
         TabIndex        =   9
         Top             =   0
         Width           =   975
      End
   End
   Begin VB.Frame fraAppearance 
      Caption         =   "Appearance"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2865
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3255
      Begin VB.Frame fraAngleMark 
         Caption         =   "Angle mark"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   2775
         Begin VB.TextBox txtRadius 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.4
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   330
            Left            =   120
            TabIndex        =   13
            Top             =   1080
            Width           =   372
         End
         Begin VB.VScrollBar vsbRadius 
            Height          =   330
            Left            =   480
            Max             =   40
            Min             =   10
            TabIndex        =   12
            Top             =   1080
            Value           =   10
            Width           =   198
         End
         Begin VB.CheckBox chkHideMeasurement 
            Caption         =   "Hide measurement text"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.4
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   1560
            Width           =   2535
         End
         Begin VB.VScrollBar vsbDrawWidth 
            Height          =   330
            Left            =   1680
            Max             =   4
            Min             =   1
            TabIndex        =   6
            Top             =   1080
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
            ForeColor       =   &H80000012&
            Height          =   330
            Left            =   1320
            TabIndex        =   5
            Top             =   1080
            Width           =   372
         End
         Begin VB.ComboBox cmbDrawStyle 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.6
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            ItemData        =   "frmAngleProps.frx":0442
            Left            =   120
            List            =   "frmAngleProps.frx":0452
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   360
            Width           =   975
         End
         Begin VB.Label lblRadius 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Radius"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.4
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   14
            Top             =   840
            Width           =   495
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
            Left            =   1320
            TabIndex        =   7
            Top             =   840
            Width           =   1410
         End
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
      Begin VB.Label lblForeColor 
         BackStyle       =   0  'Transparent
         Caption         =   "Forecolor"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   600
         TabIndex        =   2
         Top             =   360
         Width           =   1650
      End
   End
End
Attribute VB_Name = "frmMeasurementProps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const MaxAngleMarkWidth = 4
Const MaxRad = 40
Const MinRad = 10

Dim unlCancel As Boolean

Private Sub cmbDrawStyle_Change()
chkHideMeasurement.Enabled = cmbDrawStyle.ListIndex > 0
txtDrawWidth.Enabled = chkHideMeasurement.Enabled
vsbDrawWidth.Enabled = chkHideMeasurement.Enabled
lblDrawWidth.Enabled = chkHideMeasurement.Enabled
txtRadius.Enabled = chkHideMeasurement.Enabled
vsbRadius.Enabled = chkHideMeasurement.Enabled
lblRadius.Enabled = chkHideMeasurement.Enabled
txtRadius.BackColor = IIf(txtDrawWidth.Enabled, vbWindowBackground, vbButtonFace)
txtDrawWidth.BackColor = IIf(txtDrawWidth.Enabled, vbWindowBackground, vbButtonFace)
End Sub

Private Sub cmbDrawStyle_Click()
cmbDrawStyle_Change
End Sub

Private Sub cmdCancel_Click()
unlCancel = True
Unload Me
End Sub

Private Sub cmdOK_Click()
unlCancel = False
Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then unlCancel = True: Unload Me
If KeyCode = vbKeyReturn Then Unload Me
End Sub

Private Sub Form_Load()
Dim HeightDelta As Single

If Figures(ActiveFigure).FigureType = dsMeasureDistance Then
    fraAngleMark.Visible = False
    HeightDelta = -fraAngleMark.Height
    fraAppearance.Height = fraAppearance.Height + HeightDelta
    Height = Height + HeightDelta

Else
    
    txtDrawWidth.Text = Format(Figures(ActiveFigure).DrawWidth)
    vsbDrawWidth.Value = MaxAngleMarkWidth + 1 - Figures(ActiveFigure).DrawWidth
    If Figures(ActiveFigure).DrawStyle > 0 Then
        cmbDrawStyle.ListIndex = ((Figures(ActiveFigure).DrawStyle - 1) Mod 3) + 1
        chkHideMeasurement.Value = -(Figures(ActiveFigure).DrawStyle > 3)
    Else
        cmbDrawStyle.ListIndex = 0
    End If
    
    lblRadius.Caption = GetString(ResSize)
    If Figures(ActiveFigure).AuxInfo(2) < MinRad Or Figures(ActiveFigure).AuxInfo(2) > MaxRad Then Figures(ActiveFigure).AuxInfo(2) = defAngleMarkRadius
    txtRadius.Text = Figures(ActiveFigure).AuxInfo(2)
    vsbRadius.Value = MaxRad + MinRad - Val(txtRadius.Text)
End If

csbForeColor.Color = Figures(ActiveFigure).ForeColor

FillDialogStrings

'ShadowControl cmdCancel
'ShadowControl cmdOK
End Sub

Public Sub FillDialogStrings()
Caption = GetString(ResMnuMeasurementProperties)
cmdCancel.Caption = GetString(ResCancel)
lblForeColor.Caption = GetString(ResColor)
lblDrawWidth.Caption = GetString(ResDrawWidth)
fraAppearance.Caption = GetString(ResAppearance)
fraAngleMark.Caption = GetString(ResAngleMark)
chkHideMeasurement.Caption = GetString(ResHideMeasurementText)
cmbDrawStyle.List(0) = GetString(ResNone)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode <> vbFormCode Then unlCancel = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
If unlCancel = True Then Exit Sub

Figures(ActiveFigure).ForeColor = csbForeColor.Color
If Figures(ActiveFigure).FigureType = dsMeasureAngle Then
    Figures(ActiveFigure).DrawStyle = cmbDrawStyle.ListIndex
    If chkHideMeasurement.Enabled And chkHideMeasurement.Value = 1 And Figures(ActiveFigure).DrawStyle > 0 Then Figures(ActiveFigure).DrawStyle = Figures(ActiveFigure).DrawStyle + 3
    Figures(ActiveFigure).DrawWidth = Val(txtDrawWidth.Text)
    Figures(ActiveFigure).AuxInfo(2) = Val(txtRadius.Text)
End If

PaperCls
ShowAll
End Sub

Private Sub txtDrawWidth_Change()
If Not Visible Or Not txtDrawWidth.Visible Or Not IsNumeric(txtDrawWidth) Then Exit Sub
If Val(txtDrawWidth.Text) <> Int(Val(txtDrawWidth.Text)) Then Exit Sub
If txtDrawWidth.Enabled Then
    vsbDrawWidth.Enabled = False
    vsbDrawWidth.Value = MaxAngleMarkWidth + 1 - Int(Val(txtDrawWidth.Text))
    vsbDrawWidth.Enabled = True
    txtDrawWidth.SetFocus
End If
End Sub

Private Sub txtDrawWidth_Validate(Cancel As Boolean)
If Not txtDrawWidth.Enabled Or Not txtDrawWidth.Visible Then Exit Sub
If Not IsNumeric(txtDrawWidth.Text) Then GoTo ERR
If Val(txtDrawWidth.Text) <> Int(Val(txtDrawWidth.Text)) Then GoTo ERR
If Val(txtDrawWidth.Text) < 1 Or Val(txtDrawWidth.Text) > MaxAngleMarkWidth Then GoTo ERR
Exit Sub

ERR:
Cancel = 1
MsgBox GetString(ResDrawWidth) + " " + GetString(ResMsgExpressedBy) + " " + GetString(ResMsgFrom) + " 1 " + GetString(ResMsgTo) & " " & MaxAngleMarkWidth & ".", vbInformation
txtDrawWidth.SetFocus
End Sub

Private Sub vsbDrawWidth_Change()
If Not Visible Or Not vsbDrawWidth.Enabled Or Not txtDrawWidth.Visible Then Exit Sub
If Not IsNumeric(txtDrawWidth) Then Exit Sub
If Val(txtDrawWidth.Text) <> Int(Val(txtDrawWidth.Text)) Then Exit Sub

txtDrawWidth.Enabled = False
txtDrawWidth.Text = MaxAngleMarkWidth + 1 - vsbDrawWidth.Value
txtDrawWidth.SelStart = 0
txtDrawWidth.SelLength = Len(txtDrawWidth)
txtDrawWidth.Enabled = True
txtDrawWidth.SetFocus
End Sub

Private Sub txtRadius_Change()
If Not Visible Or Not txtRadius.Visible Or Not IsNumeric(txtRadius) Then Exit Sub
If Val(txtRadius.Text) <> Int(Val(txtRadius.Text)) Then Exit Sub
If Val(txtRadius.Text) < MinRad Or Val(txtRadius.Text) > MaxRad Then Exit Sub

If txtRadius.Enabled Then
    vsbRadius.Enabled = False
    vsbRadius.Value = MaxRad + MinRad - Int(Val(txtRadius.Text))
    vsbRadius.Enabled = True
    txtRadius.SetFocus
End If
End Sub

Private Sub txtRadius_Validate(Cancel As Boolean)
If Not txtRadius.Enabled Or Not txtRadius.Visible Then Exit Sub
If Not IsNumeric(txtRadius.Text) Then GoTo ERR
If Val(txtRadius.Text) <> Int(Val(txtRadius.Text)) Then GoTo ERR
If Val(txtRadius.Text) < MinRad Or Val(txtRadius.Text) > MaxRad Then GoTo ERR
Exit Sub

ERR:
Cancel = 1
MsgBox GetString(ResSize) + " " + GetString(ResMsgExpressedBy) + " " + GetString(ResMsgFrom) & " " & MinRad & " " & GetString(ResMsgTo) & " " & MaxRad & ".", vbInformation
txtRadius.SetFocus
End Sub

Private Sub vsbRadius_Change()
If Not Visible Or Not vsbRadius.Enabled Or Not txtRadius.Visible Then Exit Sub
If Not IsNumeric(txtRadius) Then Exit Sub
If Val(txtRadius.Text) <> Int(Val(txtRadius.Text)) Then Exit Sub

txtRadius.Enabled = False
txtRadius.Text = MaxRad + MinRad - vsbRadius.Value
txtRadius.SelStart = 0
txtRadius.SelLength = Len(txtRadius)
txtRadius.Enabled = True
txtRadius.SetFocus
End Sub

