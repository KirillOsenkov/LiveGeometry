VERSION 5.00
Begin VB.Form frmFigureList 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Figure List"
   ClientHeight    =   5160
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   6480
   HelpContextID   =   46
   Icon            =   "FigureList.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   6480
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   14
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   4680
      Width           =   1095
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
      Height          =   4455
      Left            =   3120
      TabIndex        =   15
      Top             =   120
      Visible         =   0   'False
      Width           =   3255
      Begin VB.CommandButton cmdSaveSettings 
         Caption         =   "Save these settings"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   3480
         Width           =   2760
      End
      Begin VB.CommandButton cmdRestoreSettings 
         Caption         =   "Restore settings"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   3960
         Width           =   2760
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
         ForeColor       =   &H80000012&
         Height          =   288
         ItemData        =   "FigureList.frx":0442
         Left            =   240
         List            =   "FigureList.frx":044C
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   2640
         Width           =   2760
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
         ForeColor       =   &H80000012&
         Height          =   288
         ItemData        =   "FigureList.frx":0456
         Left            =   240
         List            =   "FigureList.frx":0469
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1200
         Width           =   1545
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
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   3120
         Width           =   1215
      End
      Begin VB.TextBox txtDrawWidth 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   330
         Left            =   2040
         TabIndex        =   4
         Top             =   1200
         Width           =   372
      End
      Begin VB.VScrollBar vsbDrawWidth 
         Height          =   330
         Left            =   2400
         Max             =   16
         Min             =   1
         TabIndex        =   5
         Top             =   1200
         Value           =   1
         Width           =   198
      End
      Begin VB.ComboBox cmbFillStyle 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   288
         ItemData        =   "FigureList.frx":04A9
         Left            =   240
         List            =   "FigureList.frx":04C5
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1920
         Width           =   2760
      End
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
         TabIndex        =   21
         Top             =   360
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
         Left            =   2040
         TabIndex        =   20
         Top             =   960
         Width           =   1170
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
         TabIndex        =   19
         Top             =   960
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
         TabIndex        =   18
         Top             =   2400
         Width           =   1650
      End
      Begin VB.Label lblFillStyle 
         BackStyle       =   0  'Transparent
         Caption         =   "Fill style"
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
         Left            =   240
         TabIndex        =   17
         Top             =   1680
         Width           =   1650
      End
      Begin VB.Label lblFillColor 
         BackStyle       =   0  'Transparent
         Caption         =   "Fill color"
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
         TabIndex        =   16
         Top             =   600
         Width           =   1650
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   13
      Top             =   4680
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   12
      Top             =   4680
      Width           =   735
   End
   Begin VB.ListBox lstFigures 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4272
      Left            =   120
      MultiSelect     =   2  'Extended
      TabIndex        =   0
      Top             =   210
      Width           =   2895
   End
   Begin VB.Label lblSelectFigure 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   3120
      TabIndex        =   22
      Top             =   240
      Width           =   3255
   End
End
Attribute VB_Name = "frmFigureList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim unlCancel As Boolean
Dim ClonedFigures() As Figure
Dim DefaultSettings As Figure
Dim pAction As Action
Dim figList As ObjectList
Dim DrawModeArray, AntiDrawModeArray
Dim Expanded As Boolean
Dim FromCode As Boolean
Dim FromVsbCode As Boolean
Dim txtChangeFromCode As Boolean

'========================================================================
' Reaction to properties change
'========================================================================

Private Sub chkVisible_Click()
If Not Visible Or FromCode Or chkVisible.Value = 2 Then Exit Sub
ApplyVisible Not CBool(-chkVisible.Value)
ApplyProp
End Sub

Private Sub ApplyProp()
If Not cmdApply.Enabled Then cmdApply.Enabled = True
If SelCount > 1 Then FillMultipleFigureProperties
End Sub

Private Sub ApplyVisible(ByVal B As Boolean)
Dim Z As Long

For Z = 1 To figList.FigureCount
    If lstFigures.Selected(Z - 1) Then
        ClonedFigures(figList.Figures(Z)).Hide = B
    End If
Next
End Sub

Private Sub cmbDrawMode_Click()
If Not Visible Or FromCode Or cmbDrawMode.ListIndex = -1 Then Exit Sub
ApplyDrawMode DrawModeArray(cmbDrawMode.ListIndex)
ApplyProp
End Sub

Private Sub ApplyDrawMode(ByVal D As Long)
Dim Z As Long

For Z = 1 To figList.FigureCount
    If lstFigures.Selected(Z - 1) Then
        ClonedFigures(figList.Figures(Z)).DrawMode = D
    End If
Next
End Sub

Private Sub cmbDrawStyle_Click()
If Not Visible Or FromCode Or cmbDrawStyle.ListIndex = -1 Then Exit Sub
ApplyDrawStyle cmbDrawStyle.ListIndex
ApplyProp
End Sub

Private Sub ApplyDrawStyle(ByVal D As Long)
Dim Z As Long

For Z = 1 To figList.FigureCount
    If lstFigures.Selected(Z - 1) Then
        ClonedFigures(figList.Figures(Z)).DrawStyle = D
    End If
Next
End Sub

Private Sub cmbFillStyle_Click()
If Not Visible Or FromCode Or cmbFillStyle.ListIndex = -1 Then Exit Sub
ApplyFillStyle cmbFillStyle.ListIndex - 1
ApplyProp
End Sub

Private Sub ApplyFillStyle(ByVal F As Long)
Dim Z As Long

For Z = 1 To figList.FigureCount
    If lstFigures.Selected(Z - 1) Then
        ClonedFigures(figList.Figures(Z)).FillStyle = F
    End If
Next
End Sub

Private Sub csbFillColor_ColorChanged(ByVal NewColor As Long, ByVal OldColor As Long)
ApplyFillColor NewColor
ApplyProp
End Sub

Private Sub ApplyFillColor(ByVal C As Long)
Dim Z As Long

For Z = 1 To figList.FigureCount
    If lstFigures.Selected(Z - 1) Then
        ClonedFigures(figList.Figures(Z)).FillColor = C
    End If
Next
End Sub

Private Sub csbForeColor_ColorChanged(ByVal NewColor As Long, ByVal OldColor As Long)
ApplyForeColor NewColor
ApplyProp
End Sub

Private Sub ApplyForeColor(ByVal C As Long)
Dim Z As Long

For Z = 1 To figList.FigureCount
    If lstFigures.Selected(Z - 1) Then
        ClonedFigures(figList.Figures(Z)).ForeColor = C
    End If
Next
End Sub

'=======================================================

'==============================================

Private Sub cmdApply_Click()
Apply
End Sub

Private Sub cmdHelp_Click()
DisplayHelpTopic Me.HelpContextID
End Sub

Private Sub cmdCancel_Click()
unlCancel = True
Unload Me
End Sub

Private Sub cmdOK_Click()
unlCancel = False
Unload Me
End Sub

Private Sub cmdRestoreSettings_Click()
SetAllPropsFromFigure DefaultSettings
With DefaultSettings
    ApplyDrawMode .DrawMode
    ApplyDrawStyle .DrawStyle
    ApplyFillColor .FillColor
    ApplyFillStyle .FillStyle
    ApplyForeColor .ForeColor
    ApplyVisible .Hide
    ApplyWidth .DrawWidth
    ApplyProp
End With
End Sub

Private Sub cmdSaveSettings_Click()
WriteAllPropsToFigure DefaultSettings
End Sub

'=======================================================

Private Sub Form_DblClick()
Expanded = Not Expanded
If Expanded Then
    Me.Width = fraAppearance.Width + 3 * lstFigures.Left + lstFigures.Width + 4 * Screen.TwipsPerPixelX
    cmdOK.Left = cmdCancel.Left - cmdOK.Width - lstFigures.Left
Else
    Me.Width = 2 * lstFigures.Left + lstFigures.Width + 4 * Screen.TwipsPerPixelX
    cmdOK.Left = lstFigures.Left + lstFigures.Width - cmdOK.Width
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then cmdCancel_Click
End Sub

Private Sub Form_Load()
Dim Z As Long

If FigureCount = 0 Then unlCancel = True: Unload Me
Expanded = True

FillFigureWithDefaults DefaultSettings
ClonedFigures = Figures

ObjectListClear figList

For Z = 0 To FigureCount - 1
    If IsVisual(Z) Then ObjectListAdd figList, gotFigure, Z
Next

FillDialogStrings
PrepareControls
FillFigureList
ValidateDialogElementsVisibility
If figList.FigureCount > 0 Then FillSingleFigureProperties ClonedFigures(figList.Figures(1))

cmdApply.Enabled = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode <> vbFormCode Then unlCancel = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not unlCancel Then Apply

PaperCls
ShowAll
FormMain.Enabled = True
End Sub

'============================================

Private Sub lstFigures_Click()
ValidateDialogElementsVisibility
FillPropertyPane
DrawSelectedFigures
End Sub

Private Sub FillPropertyPane()
Select Case SelCount
Case 1
    FillSingleFigureProperties ClonedFigures(figList.Figures(CurFigure))
Case Is > 1
    FillMultipleFigureProperties
End Select
End Sub

Public Sub FillDialogStrings()
Caption = GetString(ResFigureList)
cmdCancel.Caption = GetString(ResCancel)
cmdHelp.Caption = GetString(ResHelp)
cmdApply.Caption = GetString(ResApply)
cmdSaveSettings.Caption = GetString(ResSaveAsDefaults)
cmdRestoreSettings.Caption = GetString(ResLoadDefaults)

lblDrawMode = GetString(ResDrawMode)
lblDrawStyle = GetString(ResDrawStyle)
lblDrawWidth = GetString(ResDrawWidth)
lblFillColor = GetString(ResFillColor) ' GetString(ResForeColor) & " - " & GetString(ResFill)
lblFillStyle = GetString(ResFill)
fraAppearance.Caption = GetString(ResAppearance)
lblForeColor = GetString(ResForeColor)
lblSelectFigure.Caption = GetString(ResSelectFigureFromList)
chkVisible.Caption = GetString(ResVisible)
End Sub

Public Sub PrepareControls()
Dim Z As Long

For Z = 0 To cmbFillStyle.ListCount - 1
    cmbFillStyle.List(Z) = GetString(ResFillStyleBase + 2 * Z)
Next

For Z = 0 To cmbDrawMode.ListCount - 1
    cmbDrawMode.List(Z) = GetString(ResDrawModeSolid + 2 * Z)
Next

DrawModeArray = Array(13, 9)
AntiDrawModeArray = Array(1, 1, 1, 1, 1, 1, 1, 1, 2, 1, 1, 1, 1, 1, 1, 1)
End Sub

Private Sub SetAllPropsFromFigure(Fig As Figure)
With Fig
    FromCode = True
    If csbFillColor.Color <> .FillColor Then csbFillColor.Color = .FillColor
    If csbForeColor.Color <> .ForeColor Then csbForeColor.Color = .ForeColor
    If cmbDrawStyle.ListIndex <> .DrawStyle Then cmbDrawStyle.ListIndex = .DrawStyle
    If cmbFillStyle.ListIndex <> .FillStyle + 1 Then cmbFillStyle.ListIndex = .FillStyle + 1
    If cmbDrawMode.ListIndex <> AntiDrawModeArray(.DrawMode - 1) - 1 Then cmbDrawMode.ListIndex = AntiDrawModeArray(.DrawMode - 1) - 1
    If chkVisible.Value <> -(Not .Hide) Then chkVisible.Value = -(Not .Hide)
    If Val(txtDrawWidth.Text) <> .DrawWidth Then txtDrawWidth.Text = .DrawWidth
    If vsbDrawWidth.Value <> MaxDrawWidth + 1 - .DrawWidth Then vsbDrawWidth.Value = MaxDrawWidth + 1 - .DrawWidth
    If csbForeColor.DefaultValue <> False Then csbForeColor.DefaultValue = False
    If csbFillColor.DefaultValue <> False Then csbFillColor.DefaultValue = False
    FromCode = False
End With
End Sub

Private Sub WriteAllPropsToFigure(Fig As Figure)
With Fig
    .FillColor = csbFillColor.Color
    .ForeColor = csbForeColor.Color
    .DrawStyle = cmbDrawStyle.ListIndex
    .FillStyle = cmbFillStyle.ListIndex - 1
    .DrawMode = DrawModeArray(cmbDrawMode.ListIndex)
    .Hide = CBool(chkVisible.Value - 1)
    .DrawWidth = Val(txtDrawWidth.Text)
End With
End Sub

Public Sub Apply()
If Not cmdApply.Enabled Then Exit Sub
cmdApply.Enabled = False
RecordGenericAction ResUndoFigurePropertiesChange

Figures = ClonedFigures
DrawSelectedFigures
End Sub

Public Sub FillFigureList()
Dim Z As Long

lstFigures.Clear
For Z = 1 To figList.FigureCount
    lstFigures.AddItem ClonedFigures(figList.Figures(Z)).Name
Next
AddListboxScrollbar lstFigures
End Sub

Private Property Get SelCount() As Long
SelCount = lstFigures.SelCount
End Property

Private Property Get CurFigure() As Long
Dim Z As Long

For Z = 0 To lstFigures.ListCount - 1
    If lstFigures.Selected(Z) Then
        CurFigure = Z + 1
        Exit Property
    End If
Next
End Property

Private Sub ValidateDialogElementsVisibility()
With lstFigures
    Select Case .SelCount
    Case 0
        If fraAppearance.Visible Then fraAppearance.Visible = False
        If Not lblSelectFigure.Visible Then lblSelectFigure.Visible = True
    Case 1
        If Not fraAppearance.Visible Then fraAppearance.Visible = True
        If lblSelectFigure.Visible Then lblSelectFigure.Visible = False
    Case Else
        If Not fraAppearance.Visible Then fraAppearance.Visible = True
        If lblSelectFigure.Visible Then lblSelectFigure.Visible = False
    End Select
End With
End Sub

Private Sub FillSingleFigureProperties(F As Figure)
If fraAppearance.Caption <> F.Name & " " & GetString(ResPropsTitle) Then fraAppearance.Caption = F.Name & " " & GetString(ResPropsTitle)
SetAllPropsFromFigure F
EnsureAuxControlsEnablity IsCircleType(F.FigureType)
End Sub

Private Sub FillMultipleFigureProperties()
Dim Z As Long, tEnteredSelection As Boolean

Dim tForecolor As Long
Dim tFillcolor As Long
Dim tHide As Boolean
Dim tWidth As Long
Dim tDrawStyle As Long
Dim tDrawMode As Long
Dim tFillStyle As Long

Dim newForecolor As Long
Dim newFillcolor As Long
Dim newFillcolorEnabled As Boolean
Dim newForecolorDef As Boolean
Dim newFillcolorDef As Boolean
Dim newHide As Long
Dim newWidth As Long
Dim newDrawStyle As Long
Dim newDrawMode As Long
Dim newFillStyle As Long
Dim newFillStyleEnabled As Boolean
Dim newDrawStyleEnabled As Boolean

fraAppearance.Caption = Replace(GetString(ResPropertiesOfFigures), "%1", SelCount)
'EnsureAuxControlsEnablity True

tEnteredSelection = False
FromCode = True

For Z = 1 To figList.FigureCount
    If lstFigures.Selected(Z - 1) Then
        If Not tEnteredSelection Then
            tEnteredSelection = True
            
            With ClonedFigures(figList.Figures(Z))
                tForecolor = .ForeColor
                tFillcolor = .FillColor
                If Not IsCircleType(.FigureType) Then
                    tFillcolor = -1
                End If
                tHide = .Hide
                tWidth = .DrawWidth
                tDrawMode = .DrawMode
                tDrawStyle = .DrawStyle
                tFillStyle = .FillStyle
            End With
            
            txtDrawWidth.Text = tWidth
            vsbDrawWidth.Value = 1 + MaxDrawWidth - tWidth
            vsbDrawWidth.Enabled = True
            
            newForecolorDef = False
            newFillcolorDef = False
            newForecolor = tForecolor
            newFillcolor = tFillcolor
            newFillcolorEnabled = tFillcolor <> -1
            newHide = tHide + 1
            newWidth = tWidth
            newDrawMode = AntiDrawModeArray(tDrawMode - 1) - 1
            newDrawStyle = tDrawStyle
            newFillStyle = tFillStyle + 1
            newDrawStyleEnabled = tWidth = 1
            
        Else
            With ClonedFigures(figList.Figures(Z))
                If .ForeColor <> tForecolor Then
                    newForecolorDef = True
                End If
                If tWidth <> .DrawWidth Then
                    vsbDrawWidth.Enabled = False
                    If .DrawWidth = 1 Then newDrawStyleEnabled = True
                    txtDrawWidth.Text = ""
                End If
                If IsCircleType(.FigureType) Then
                    If .FillColor >= 0 And tFillcolor = -1 Then
                        tFillcolor = .FillColor
                    End If
                    If .FillColor <> tFillcolor And newFillcolor <> -1 Then
                        newFillcolorDef = True
                    End If
                    newFillcolorEnabled = True
                    newFillcolor = .FillColor
                End If
                If tDrawMode <> .DrawMode Then
                    newDrawMode = -1
                End If
                If tHide <> .Hide Then
                    newHide = 2
                End If
                If tDrawStyle <> .DrawStyle Then
                    newDrawStyle = -1
                End If
                If tFillStyle <> .FillStyle Then
                    newFillStyle = -1
                End If
            End With
        End If
    End If
Next Z

If csbForeColor.Color <> newForecolor Then csbForeColor.Color = newForecolor
If csbForeColor.DefaultValue <> newForecolorDef Then csbForeColor.DefaultValue = newForecolorDef
If csbFillColor.Color <> newFillcolor Then csbFillColor.Color = newFillcolor
If csbFillColor.DefaultValue <> newFillcolorDef Then csbFillColor.DefaultValue = newFillcolorDef
If csbFillColor.Enabled <> newFillcolorEnabled Then csbFillColor.Enabled = newFillcolorEnabled
If cmbFillStyle.Enabled <> csbFillColor.Enabled Then EnableFillStyle csbFillColor.Enabled
If cmbDrawMode.ListIndex <> newDrawMode Then cmbDrawMode.ListIndex = newDrawMode
If chkVisible.Value <> newHide Then chkVisible.Value = newHide
If cmbDrawStyle.ListIndex <> newDrawStyle Then cmbDrawStyle.ListIndex = newDrawStyle
If cmbDrawStyle.Enabled <> newDrawStyleEnabled Then EnableDrawStyle newDrawStyleEnabled
If cmbFillStyle.ListIndex <> newFillStyle Then cmbFillStyle.ListIndex = newFillStyle

FromCode = False
End Sub

Private Sub EnsureAuxControlsEnablity(ByVal B As Boolean)
EnableFillColor B
EnableFillStyle B
If Val(txtDrawWidth.Text) = 1 Xor cmbDrawStyle.Enabled Then EnableDrawStyle Val(txtDrawWidth.Text) = 1
End Sub

Private Sub lstFigures_KeyDown(KeyCode As Integer, Shift As Integer)
Dim Z As Long, Sel As Boolean

If KeyCode = vbKeyA And Shift = 2 Then
    Sel = lstFigures.SelCount < lstFigures.ListCount
    For Z = 0 To lstFigures.ListCount - 1
        lstFigures.Selected(Z) = Sel
    Next
End If
End Sub

'============================================

Private Sub txtDrawWidth_Change()
If Not Visible Or Not IsNumeric(txtDrawWidth) Or Not txtDrawWidth.Enabled Then Exit Sub
If Val(txtDrawWidth.Text) <> Int(Val(txtDrawWidth.Text)) Then Exit Sub
If Val(txtDrawWidth.Text) < 1 Or Val(txtDrawWidth.Text) > MaxDrawWidth Then Exit Sub

If SelCount = 1 Then EnableDrawStyle Val(txtDrawWidth.Text) = 1

If Not FromVsbCode Then
    FromVsbCode = True
    vsbDrawWidth.Value = MaxDrawWidth + 1 - Int(Val(txtDrawWidth.Text))
    FromVsbCode = False
End If

ApplyWidth Val(txtDrawWidth.Text)
ApplyProp
End Sub

Private Sub vsbDrawWidth_Change()
If Not Visible Or Not vsbDrawWidth.Enabled Then Exit Sub
If Not IsNumeric(txtDrawWidth) Then Exit Sub
If Val(txtDrawWidth.Text) <> Int(Val(txtDrawWidth.Text)) Then Exit Sub

If Not FromVsbCode Then
    FromVsbCode = True
    txtDrawWidth.Text = MaxDrawWidth + 1 - vsbDrawWidth.Value
    FromVsbCode = False
End If
End Sub

Private Sub ApplyWidth(ByVal W As Long)
Dim Z As Long
For Z = 1 To figList.FigureCount
    If lstFigures.Selected(Z - 1) Then
        ClonedFigures(figList.Figures(Z)).DrawWidth = W
    End If
Next
End Sub

Private Sub txtDrawWidth_Validate(Cancel As Boolean)
If Not txtDrawWidth.Enabled Then Exit Sub
If Not IsNumeric(txtDrawWidth.Text) Then GoTo ERR
If Val(txtDrawWidth.Text) <> Int(Val(txtDrawWidth.Text)) Then GoTo ERR
If Val(txtDrawWidth.Text) < 1 Or Val(txtDrawWidth.Text) > MaxDrawWidth Then GoTo ERR
Exit Sub

ERR:
Cancel = 1
FormMain.Enabled = False
MsgBox GetString(ResDrawWidth) + " " + GetString(ResMsgExpressedBy) + " " + GetString(ResMsgFrom) + " 1 " + GetString(ResMsgTo) + " 16.", vbInformation
End Sub

'============================================

Private Sub DrawSelectedFigures()
On Local Error Resume Next

Dim Z As Long

PaperCls
ShowAll , , False

For Z = 1 To figList.FigureCount
    If lstFigures.Selected(Z - 1) Then
        If Figures(figList.Figures(Z)).Hide Then
            ShowSelectedFigure Paper.hDC, figList.Figures(Z), , True, True
        Else
            ShowSelectedFigure Paper.hDC, figList.Figures(Z)
        End If
    End If
Next

Paper.Refresh
End Sub

Private Sub EnableDrawStyle(ByVal Enable As Boolean)
cmbDrawStyle.Enabled = Enable
cmbDrawStyle.BackColor = IIf(Enable, vbWindowBackground, vbButtonFace)
lblDrawStyle.Enabled = Enable
End Sub

Private Sub EnableFillStyle(ByVal B As Boolean)
If cmbFillStyle.Enabled Xor B Then cmbFillStyle.Enabled = B
If lblFillStyle.Enabled Xor B Then lblFillStyle.Enabled = B
If cmbFillStyle.BackColor Xor IIf(B, vbWindowBackground, vbButtonFace) Then cmbFillStyle.BackColor = IIf(B, vbWindowBackground, vbButtonFace)
End Sub

Private Sub EnableFillColor(ByVal B As Boolean)
If csbFillColor.Enabled Xor B Then csbFillColor.Enabled = B
If lblFillColor.Enabled Xor B Then lblFillColor.Enabled = B
End Sub
