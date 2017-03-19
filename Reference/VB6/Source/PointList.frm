VERSION 5.00
Begin VB.Form frmPointList 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Point list"
   ClientHeight    =   4560
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   7080
   HelpContextID   =   45
   Icon            =   "PointList.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   7080
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraProps 
      Caption         =   "Properties"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3810
      Left            =   3120
      TabIndex        =   20
      Top             =   120
      Visible         =   0   'False
      Width           =   3855
      Begin VB.CommandButton cmdSetName 
         Caption         =   "ü"
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
         Height          =   285
         Left            =   1920
         TabIndex        =   2
         Top             =   330
         Width           =   1695
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
         TabIndex        =   15
         Top             =   3360
         Width           =   3375
      End
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
         TabIndex        =   14
         Top             =   2880
         Width           =   3375
      End
      Begin VB.CheckBox chkVisible 
         Caption         =   "Visible"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   2400
         Width           =   3375
      End
      Begin VB.TextBox txtSize 
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
         Height          =   315
         Left            =   2280
         TabIndex        =   8
         Text            =   "12"
         Top             =   1080
         Width           =   360
      End
      Begin VB.ComboBox cmbShape 
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.4
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   276
         ItemData        =   "PointList.frx":014A
         Left            =   2280
         List            =   "PointList.frx":0154
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   720
         Width           =   615
      End
      Begin VB.CheckBox chkShowName 
         Caption         =   "Show name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   10
         Top             =   1560
         Width           =   3315
      End
      Begin VB.VScrollBar vsbSize 
         Height          =   315
         Left            =   2640
         Max             =   30
         Min             =   2
         TabIndex        =   9
         Top             =   1080
         Value           =   2
         Width           =   240
      End
      Begin VB.CheckBox chkFill 
         Caption         =   "Fill:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   4
         Top             =   960
         Width           =   1575
      End
      Begin VB.CheckBox chkShowCoordinates 
         Caption         =   "Show coordinates"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   1800
         Width           =   3375
      End
      Begin VB.TextBox txtName 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1200
         TabIndex        =   1
         Top             =   330
         Width           =   735
      End
      Begin VB.CheckBox chkEnabled 
         Caption         =   "Enabled"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   2160
         Width           =   3375
      End
      Begin DG.ctlColorBox csbNameColor 
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1200
         Width           =   255
         _ExtentX        =   445
         _ExtentY        =   445
      End
      Begin DG.ctlColorBox csbFillColor 
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   960
         Width           =   255
         _ExtentX        =   445
         _ExtentY        =   445
      End
      Begin DG.ctlColorBox csbForeColor 
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   255
         _ExtentX        =   445
         _ExtentY        =   445
      End
      Begin VB.Label lblSize 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Size"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   192
         Left            =   3000
         TabIndex        =   25
         Top             =   1128
         Width           =   312
      End
      Begin VB.Label lblForeColor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Color"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   192
         Left            =   600
         TabIndex        =   24
         Top             =   720
         Width           =   360
      End
      Begin VB.Label lblShape 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Shape"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   192
         Left            =   3000
         TabIndex        =   23
         Top             =   768
         Width           =   444
      End
      Begin VB.Label lblNameColor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name color"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   192
         Left            =   600
         TabIndex        =   22
         Top             =   1212
         Width           =   792
      End
      Begin VB.Label lblName 
         Caption         =   "Name:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   360
         Width           =   855
      End
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
      TabIndex        =   16
      Top             =   4080
      Width           =   1095
   End
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
      Left            =   5760
      TabIndex        =   19
      Top             =   4080
      Width           =   1215
   End
   Begin VB.ListBox lstPoints 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3696
      Left            =   120
      MultiSelect     =   2  'Extended
      TabIndex        =   0
      Top             =   165
      Width           =   2895
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
      Left            =   3720
      TabIndex        =   17
      Top             =   4080
      Width           =   735
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
      Left            =   4560
      TabIndex        =   18
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label lblSelectPoint 
      Caption         =   "Select a point from the list to change its attributes. You can select multiple points by using Ctrl+click or Shift+click."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   3120
      TabIndex        =   26
      Top             =   120
      Width           =   3615
   End
   Begin VB.Menu mnuPopup1 
      Caption         =   "Popup1"
      Visible         =   0   'False
      Begin VB.Menu mnuRename 
         Caption         =   "Rename"
      End
   End
End
Attribute VB_Name = "frmPointList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim unlCancel As Boolean
Dim Expanded As Boolean
Dim ClonedPoints() As BasePointType
Dim DefaultSettings As BasePointType
Dim pAction As Action
Dim vsbSizeFromCode As Boolean
Dim txtSizeFromCode As Boolean
Dim FromCode As Boolean

'========================================================================
' Reaction to properties change
'========================================================================

Private Sub chkEnabled_Click()
If Not chkEnabled.Enabled Or chkEnabled.Value = 2 Or FromCode Then Exit Sub
ApplyEnabled -chkEnabled.Value
ApplyProp
End Sub

Private Sub chkFill_Click()
If Not chkFill.Enabled Or chkFill.Value = 2 Or FromCode Then Exit Sub
ApplyFill 1 - chkFill.Value
csbFillColor.Enabled = -chkFill.Value
ApplyProp
End Sub

Private Sub chkShowCoordinates_Click()
If Not chkShowCoordinates.Enabled Or chkShowCoordinates.Value = 2 Or FromCode Then Exit Sub
ApplyShowCoords -chkShowCoordinates.Value
ApplyProp
End Sub

Private Sub chkShowName_Click()
If Not chkShowName.Enabled Or chkShowName.Value = 2 Or FromCode Then Exit Sub
ApplyShowName -chkShowName.Value
ApplyProp
End Sub

Private Sub chkVisible_Click()
If Not chkVisible.Enabled Or chkVisible.Value = 2 Or FromCode Then Exit Sub
ApplyVisible Not CBool(-chkVisible.Value)
ApplyProp
End Sub

Private Sub cmbShape_Click()
If Not cmbShape.Enabled Or cmbShape.ListIndex = -1 Or FromCode Then Exit Sub
ApplyShape cmbShape.ListIndex * 2 + 1
ApplyProp
End Sub

Private Sub ApplyEnabled(ByVal B As Boolean)
Dim Z As Long

For Z = 1 To PointCount
    If lstPoints.Selected(Z - 1) Then
        ClonedPoints(Z).Enabled = B
    End If
Next
End Sub

Private Sub ApplyFill(ByVal FS As Long)
Dim Z As Long

For Z = 1 To PointCount
    If lstPoints.Selected(Z - 1) Then
        ClonedPoints(Z).FillStyle = FS
    End If
Next
End Sub

Private Sub ApplyShowCoords(ByVal B As Boolean)
Dim Z As Long

For Z = 1 To PointCount
    If lstPoints.Selected(Z - 1) Then
        ClonedPoints(Z).ShowCoordinates = B
    End If
Next
End Sub

Private Sub ApplyShowName(ByVal B As Boolean)
Dim Z As Long

For Z = 1 To PointCount
    If lstPoints.Selected(Z - 1) Then
        ClonedPoints(Z).ShowName = B
    End If
Next
csbNameColor.Enabled = B
End Sub

Private Sub ApplyVisible(ByVal B As Boolean)
Dim Z As Long

For Z = 1 To PointCount
    If lstPoints.Selected(Z - 1) Then
        ClonedPoints(Z).Hide = B
    End If
Next
End Sub

Private Sub ApplyShape(ByVal SH As Long)
Dim Z As Long

For Z = 1 To PointCount
    If lstPoints.Selected(Z - 1) Then
        ClonedPoints(Z).Shape = SH
    End If
Next
End Sub

Private Sub ApplyProp()
If Not cmdApply.Enabled Then cmdApply.Enabled = True
If SelCount > 1 Then FillMultiplePointProperties
End Sub

'==============================================

Private Sub cmdSetName_Click()
If SelCount <> 1 Then Exit Sub
If CurPoint < 1 Or CurPoint > PointCount Or Not Visible Then Exit Sub

If txtName.Text = "" Then
    MsgBox GetString(ResMsgNoName), vbExclamation
    Exit Sub
End If

If txtName.Text <> ClonedPoints(CurPoint).Name Then
    RenamePoint CurPoint, txtName.Text
    UpdateClonedPointNames
    FillPointListBox
End If

cmdSetName.Enabled = False

lstPoints.SetFocus
End Sub

'==============================================

Private Sub csbFillColor_ColorChanged(ByVal NewColor As Long, ByVal OldColor As Long)
ApplyFillColor NewColor
ApplyProp
End Sub

Private Sub csbForeColor_ColorChanged(ByVal NewColor As Long, ByVal OldColor As Long)
ApplyForeColor NewColor
ApplyProp
End Sub

Private Sub csbNameColor_ColorChanged(ByVal NewColor As Long, ByVal OldColor As Long)
ApplyNameColor NewColor
ApplyProp
End Sub

Private Sub ApplyFillColor(ByVal C As Long)
Dim Z As Long

For Z = 1 To PointCount
    If lstPoints.Selected(Z - 1) Then
        ClonedPoints(Z).FillColor = C
    End If
Next
End Sub

Private Sub ApplyForeColor(ByVal C As Long)
Dim Z As Long

For Z = 1 To PointCount
    If lstPoints.Selected(Z - 1) Then
        ClonedPoints(Z).ForeColor = C
    End If
Next
End Sub

Private Sub ApplyNameColor(ByVal C As Long)
Dim Z As Long

For Z = 1 To PointCount
    If lstPoints.Selected(Z - 1) Then
        ClonedPoints(Z).NameColor = C
    End If
Next
End Sub

'=============================================
'=============================================

Private Sub Form_DblClick()
Expanded = Not Expanded
If Expanded Then
    Me.Width = fraProps.Width + 3 * lstPoints.Left + lstPoints.Width + 4 * Screen.TwipsPerPixelX
    cmdOK.Left = cmdCancel.Left - cmdOK.Width - lstPoints.Left
Else
    Me.Width = 2 * lstPoints.Left + lstPoints.Width + 4 * Screen.TwipsPerPixelX
    cmdOK.Left = lstPoints.Left + lstPoints.Width - cmdOK.Width
End If
End Sub

Private Sub lstPoints_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lstPoints_Click
End Sub

Private Sub txtName_Change()
If Not Visible Or Not txtName.Enabled Then Exit Sub

If CurPoint < 1 Or CurPoint > PointCount Then
    cmdSetName.Enabled = False
    Exit Sub
End If
cmdSetName.Enabled = txtName <> ClonedPoints(CurPoint).Name And Len(txtName) > 0
End Sub

'========================================================================
'
'========================================================================

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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then cmdCancel_Click
End Sub

'========================================================================
' Properties of a point
'========================================================================


'Private Sub chkEnabled_Click()
'If Not IsPoint(lstPoints.ListIndex + 1) Then Exit Sub
'If BasePoint(lstPoints.ListIndex + 1).Enabled <> -(chkEnabled.Value) Then
'    BasePoint(lstPoints.ListIndex + 1).Enabled = (-chkEnabled.Value)
'End If
'End Sub

'Private Sub cmdRename_Click()
'Dim S As String, pAction As Action
'Do
'    S = InputBox(GetString(ResRename) & ": " & BasePoint(lstPoints.ListIndex + 1).Name, , GenerateNewPointName(False))
'    If S = "" Or S = BasePoint(lstPoints.ListIndex + 1).Name Then Exit Sub
'    If Not IsPoint(GetPointByName(S)) Then
'        pAction.Type = actRenamePoint
'        MakeStructureSnapshot pAction
'        RecordAction pAction
'
'        RenamePoint lstPoints.ListIndex + 1, S
'        lstPoints.List(lstPoints.ListIndex) = S & " - " & GetString(ResFigureBase + 2 * BasePoint(lstPoints.ListIndex + 1).Type)
'        fraProps.Caption = StrConv(BasePoint(lstPoints.ListIndex + 1).Name & " " & GetString(ResPropsTitle), vbProperCase)
'        PaperCls
'        ShowAll
'        Exit Sub
'    End If
'    MsgBox S & GetString(ResMsgObjectAlreadyExists)
'Loop Until Not IsPoint(GetPointByName(S))
'End Sub

Private Sub Form_Load()
Dim Z As Long

If PointCount = 0 Then unlCancel = True: Unload Me 'This is done to prevent Form_Unload to mess up with Basepoint(0), which is illegal ("Subscript out of range", he-he...)

ClonedPoints = BasePoint

FillPointWithDefaults DefaultSettings
FillDialogStrings
FillPointListBox
ValidateDialogElementsVisibility
If PointCount > 0 Then FillSinglePointProperties 1

cmdApply.Enabled = False

Expanded = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode <> vbFormCode Then unlCancel = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not unlCancel Then Apply

PaperCls
ShowAll
FormMain.Enabled = True
FormMain.SetFocus
End Sub

Private Sub cmdRestoreSettings_Click()
SetAllPropsFromPoint DefaultSettings

With DefaultSettings
    ApplyEnabled .Enabled
    ApplyFill .FillStyle
    ApplyShape .Shape
    ApplyShowCoords .ShowCoordinates
    ApplyShowName .ShowName
    ApplyVisible .Hide
    ApplyForeColor .ForeColor
    ApplyFillColor .FillColor
    ApplyNameColor .NameColor
    ApplySize .PhysicalWidth
End With
ApplyProp

End Sub

Private Sub cmdSaveSettings_Click()
WriteAllPropsToPoint DefaultSettings
End Sub

Public Sub FillDialogStrings()
fraProps.Caption = GetString(ResPropsTitle)
Caption = GetString(ResPointList)
cmdHelp.Caption = GetString(ResHelp)
cmdApply.Caption = GetString(ResApply)
cmdSaveSettings.Caption = GetString(ResSaveAsDefaults)
cmdRestoreSettings.Caption = GetString(ResLoadDefaults)

lblSelectPoint.Caption = GetString(ResSelectPointFromList)

chkFill.Caption = GetString(ResFill)
chkVisible.Caption = GetString(ResVisible)
chkEnabled.Caption = GetString(ResDisabled)
chkShowName.Caption = GetString(ResShowName)
chkShowCoordinates.Caption = GetString(ResShowCoordinates)

cmdCancel.Caption = GetString(ResCancel)
cmdSetName.Caption = GetString(ResRename)

lblNameColor.Caption = GetString(ResNameColor)
lblShape.Caption = GetString(ResShape)
lblSize.Caption = GetString(ResSize)
lblForeColor.Caption = GetString(ResForeColor)
lblName.Caption = GetString(ResName)

End Sub

Private Sub lstPoints_Click()
ValidateDialogElementsVisibility
FillPropertyPane
DrawSelectedPoints
End Sub

Private Sub FillPropertyPane()
Select Case SelCount
Case 1
    FillSinglePointProperties CurPoint
Case Is > 1
    FillMultiplePointProperties
End Select
End Sub

Private Sub FillSinglePointProperties(ByVal Index As Long)
If Index < 1 Or Index > PointCount Then Exit Sub

With ClonedPoints(Index)
    FromCode = True
    
    csbFillColor.Color = .FillColor
    csbForeColor.Color = .ForeColor
    csbNameColor.Color = .NameColor
    csbForeColor.DefaultValue = False
    csbFillColor.DefaultValue = False
    csbNameColor.DefaultValue = False
    If csbNameColor.Enabled <> -.ShowName Then csbNameColor.Enabled = -.ShowName
    If txtName.Text <> .Name Then txtName.Text = .Name
    
    If chkFill.Value <> 1 - .FillStyle Then chkFill.Value = 1 - .FillStyle
    If chkEnabled.Value <> -.Enabled Then chkEnabled.Value = -.Enabled
    If chkVisible.Value <> -(Not .Hide) Then chkVisible.Value = -(Not .Hide)
    If chkShowName.Value <> -.ShowName Then chkShowName.Value = -.ShowName
    If chkShowCoordinates.Value <> -.ShowCoordinates Then chkShowCoordinates.Value = -.ShowCoordinates
    
    If cmbShape.ListIndex <> Sgn(.Shape - 1) Then cmbShape.ListIndex = Sgn(.Shape - 1)
    If Val(txtSize.Text) <> .PhysicalWidth Then txtSize.Text = .PhysicalWidth
    If vsbSize.Value <> MaxPointSize + MinPointSize - .PhysicalWidth Then vsbSize.Value = MaxPointSize + MinPointSize - .PhysicalWidth
    
    csbFillColor.Enabled = chkFill.Value <> 0
    csbNameColor.Enabled = chkShowName.Value <> 0
    
    FromCode = False
    
    If fraProps.Caption <> Replace(GetString(ResPropertiesOfAPoint), "%1", .Name) Then fraProps.Caption = Replace(GetString(ResPropertiesOfAPoint), "%1", .Name)
End With
End Sub

Private Sub FillMultiplePointProperties()
Dim tEnabled As Boolean
Dim tVisible As Boolean
Dim tShowName As Boolean
Dim tShowCoords As Boolean
Dim tFill As Boolean
Dim tShape As Long
Dim tSize As Long

Dim tForecolor As Long
Dim tFillcolor As Long
Dim tNameColor As Long

Dim tEnteredSelection As Boolean
Dim Z As Long

Dim newFill As Long
Dim newEnabled As Long
Dim newVisible As Long
Dim newShowName As Long
Dim newShowCoords As Long
Dim newShape As Long

Dim newFillcolorEnabled As Boolean
Dim newNamecolorEnabled As Boolean
Dim newForecolor As Long
Dim newFillcolor As Long
Dim newNamecolor As Long
Dim newForecolorDef As Boolean
Dim newFillcolorDef As Boolean
Dim newNamecolorDef As Boolean

tEnteredSelection = False

For Z = 0 To lstPoints.ListCount - 1
    If lstPoints.Selected(Z) Then
        If Not tEnteredSelection Then
            tEnteredSelection = True
            
            tEnabled = ClonedPoints(Z + 1).Enabled
            tVisible = Not ClonedPoints(Z + 1).Hide
            tShowName = ClonedPoints(Z + 1).ShowName
            tShowCoords = ClonedPoints(Z + 1).ShowCoordinates
            tFill = ClonedPoints(Z + 1).FillStyle - 1
            tForecolor = ClonedPoints(Z + 1).ForeColor
            tFillcolor = ClonedPoints(Z + 1).FillColor
            tNameColor = ClonedPoints(Z + 1).NameColor
            tShape = Sgn(ClonedPoints(Z + 1).Shape - 1)
            tSize = ClonedPoints(Z + 1).PhysicalWidth
    
            
'            With newProps
'                .Enabled = -tEnabled
'                .Visible = -tVisible
'                .ShowName = -tShowName
'                .ShowCoordinates = -tShowCoords
'                .FillStyle = 1 + tFill
'                .ForeColor = tForeColor
'                .FillColor = tFillColor
'                .Name = tNameColor
'            End With
            
            
            EnablePropControls False
            
            'chkEnabled.Value = -tEnabled
            newEnabled = -tEnabled
            'chkVisible.Value = -tVisible
            newVisible = -tVisible
            'chkShowName.Value = -tShowName
            newShowName = -tShowName
            'chkShowCoordinates.Value = -tShowCoords
            newShowCoords = -tShowCoords
            'chkFill.Value = -tFill
            newFill = -tFill
            txtSize.Text = tSize
            vsbSize.Value = MaxPointSize + MinPointSize - tSize
            vsbSize.Enabled = True
            
            'csbNameColor.Enabled = tShowName
            newNamecolorEnabled = tShowName
            'csbFillColor.Enabled = tFill
            newFillcolorEnabled = tFill
            
            'cmbShape.ListIndex = tShape
            newShape = tShape
            csbForeColor.Color = tForecolor
            csbFillColor.Color = tFillcolor
            csbNameColor.Color = tNameColor
            'csbForeColor.DefaultValue = False
            'csbFillColor.DefaultValue = False
            'csbNameColor.DefaultValue = False
            
            EnablePropControls True
        Else
            If tEnabled <> ClonedPoints(Z + 1).Enabled Then
                'chkEnabled.Value = 2
                newEnabled = 2
            End If
            If tVisible <> Not ClonedPoints(Z + 1).Hide Then
                'chkVisible.Value = 2
                newVisible = 2
            End If
            If tShowName <> ClonedPoints(Z + 1).ShowName Then
                'chkShowName.Value = 2
                newShowName = 2
                'csbNameColor.Enabled = True
                newNamecolorEnabled = True
            End If
            If tShowCoords <> ClonedPoints(Z + 1).ShowCoordinates Then
                'chkShowCoordinates.Value = 2
                newShowCoords = 2
            End If
            If tFill <> ClonedPoints(Z + 1).FillStyle - 1 Then
                'chkFill.Value = 2
                newFill = 2
                'csbFillColor.Enabled = True
                newFillcolorEnabled = True
            End If
            If tForecolor <> ClonedPoints(Z + 1).ForeColor Then
                'csbForeColor.DefaultValue = True
                newForecolorDef = True
            End If
            If tFillcolor <> ClonedPoints(Z + 1).FillColor And CBool(ClonedPoints(Z + 1).FillStyle - 1) Then
                'csbFillColor.DefaultValue = True
                newFillcolorDef = True
            End If
            If tNameColor <> ClonedPoints(Z + 1).NameColor And CBool(ClonedPoints(Z + 1).ShowName) Then
                'csbNameColor.DefaultValue = True
                newNamecolorDef = True
            End If
            If tShape <> Sgn(ClonedPoints(Z + 1).Shape - 1) Then
                'If cmbShape.ListCount = 2 Then cmbShape.AddItem " "
                'cmbShape.ListIndex = -1
                newShape = -1
            End If
            If tSize <> ClonedPoints(Z + 1).PhysicalWidth Then
                vsbSize.Enabled = False
                txtSize.Text = ""
            End If
        End If
    End If
Next

If chkEnabled.Value <> newEnabled Then chkEnabled.Value = newEnabled
If chkVisible.Value <> newVisible Then chkVisible.Value = newVisible
If chkShowName.Value <> newShowName Then chkShowName.Value = newShowName
If chkShowCoordinates.Value <> newShowCoords Then chkShowCoordinates.Value = newShowCoords
If chkFill.Value <> newFill Then chkFill.Value = newFill
If cmbShape.ListIndex <> newShape Then cmbShape.ListIndex = newShape
If csbNameColor.Enabled <> newNamecolorEnabled Then csbNameColor.Enabled = newNamecolorEnabled
If csbFillColor.Enabled <> newFillcolorEnabled Then csbFillColor.Enabled = newFillcolorEnabled
If csbForeColor.DefaultValue <> newForecolorDef Then csbForeColor.DefaultValue = newForecolorDef
If csbFillColor.DefaultValue <> newFillcolorDef Then csbFillColor.DefaultValue = newFillcolorDef
If csbNameColor.DefaultValue <> newNamecolorDef Then csbNameColor.DefaultValue = newNamecolorDef


If fraProps.Caption <> Replace(GetString(ResPropertiesOfPoints), "%1", SelCount) Then fraProps.Caption = Replace(GetString(ResPropertiesOfPoints), "%1", SelCount)
End Sub

Private Sub FillPointListBox()
Dim Z As Long
Dim Sels() As Boolean, ShouldPreserveSelection As Boolean

If lstPoints.ListCount = PointCount Then
    ReDim Sels(1 To PointCount)
    For Z = 1 To PointCount
        Sels(Z) = lstPoints.Selected(Z - 1)
    Next
    ShouldPreserveSelection = True
End If

lstPoints.Clear
For Z = 1 To PointCount
    lstPoints.AddItem BasePoint(Z).Name & " - " & GetString(ResFigureBase + 2 * BasePoint(Z).Type)
Next

AddListboxScrollbar lstPoints

If ShouldPreserveSelection Then
    For Z = 1 To PointCount
        lstPoints.Selected(Z - 1) = Sels(Z)
    Next
End If
End Sub

Private Sub UpdateClonedPointNames()
Dim Z As Long

For Z = 1 To PointCount
    ClonedPoints(Z).Name = BasePoint(Z).Name
Next
End Sub

Private Sub ValidateDialogElementsVisibility()
With lstPoints
    Select Case .SelCount
    Case 0
        If fraProps.Visible Then fraProps.Visible = False
        If Not lblSelectPoint.Visible Then lblSelectPoint.Visible = True
    Case 1
        If Not fraProps.Visible Then fraProps.Visible = True
        If Not lblName.Visible Then lblName.Visible = True
        If Not txtName.Visible Then txtName.Visible = True
        If Not cmdSetName.Visible Then cmdSetName.Visible = True
        'If cmdSetAttr.Visible Then cmdSetAttr.Visible = False
        If lblSelectPoint.Visible Then lblSelectPoint.Visible = False
    Case Else
        If Not fraProps.Visible Then fraProps.Visible = True
        If lblName.Visible Then lblName.Visible = False
        If txtName.Visible Then txtName.Visible = False
        If cmdSetName.Visible Then cmdSetName.Visible = False
        'If Not cmdSetAttr.Visible Then cmdSetAttr.Visible = True
        If lblSelectPoint.Visible Then lblSelectPoint.Visible = False
    End Select
End With
End Sub

'============================================

Private Sub vsbSize_Change()
If Not Visible Or Not vsbSize.Enabled Or FromCode Then Exit Sub
If vsbSize.Value < MinPointSize Or vsbSize.Value > MaxPointSize Then Exit Sub

txtSizeFromCode = True
txtSize = MaxPointSize + MinPointSize - vsbSize.Value
txtSize.SelStart = 0
txtSize.SelLength = Len(txtSize)
txtSizeFromCode = False
End Sub

Private Sub txtSize_Change()
If Not Visible Or Not txtSize.Enabled Or Not IsNumeric(txtSize.Text) Or FromCode Then Exit Sub
If Val(txtSize.Text) <> Int(Val(txtSize.Text)) Then Exit Sub
If txtSize < MinPointSize Or txtSize > MaxPointSize Then Exit Sub

If Not txtSizeFromCode Then
    vsbSize.Enabled = False
    vsbSize.Value = MaxPointSize + MinPointSize - Val(txtSize)
    vsbSize.Enabled = True
End If

ApplySize Val(txtSize.Text)
ApplyProp '????? There wasn't this line before; inserted for generality?
End Sub

Private Sub ApplySize(ByVal S As Long)
Dim Z As Long

For Z = 1 To PointCount
    If lstPoints.Selected(Z - 1) Then
        ClonedPoints(Z).PhysicalWidth = S
        ClonedPoints(Z).Width = S
        ToLogicalLength ClonedPoints(Z).Width
    End If
Next
End Sub

Private Sub txtSize_Validate(Cancel As Boolean)
If Not IsNumeric(txtSize) Or Not txtSize.Enabled Then GoTo ERR
If Val(txtSize.Text) <> Int(Val(txtSize.Text)) Then GoTo ERR
If Val(txtSize.Text) < MinPointSize Or Val(txtSize.Text) > MaxPointSize Then GoTo ERR
Exit Sub

ERR:
Cancel = 1
MsgBox GetString(ResSize) & " " & GetString(ResMsgExpressedBy) & " " & GetString(ResMsgFrom) & " " & MinPointSize & " " & GetString(ResMsgTo) & " " & MaxPointSize & ".", vbExclamation
End Sub

'============================================

Private Property Get SelCount() As Long
SelCount = lstPoints.SelCount
End Property

Private Property Get CurPoint() As Long
Dim Z As Long

For Z = 0 To lstPoints.ListCount - 1
    If lstPoints.Selected(Z) Then
        CurPoint = Z + 1
        Exit Property
    End If
Next
End Property

Private Sub WriteAllPropsToPoint(P As BasePointType)
FillPointWithDefaults P
With P
    If chkEnabled.Value < 2 Then P.Enabled = -chkEnabled.Value
    P.FillColor = csbFillColor.Color
    If chkFill.Value < 2 Then P.FillStyle = 1 - chkFill.Value
    P.ForeColor = csbForeColor.Color
    If chkVisible.Value < 2 Then P.Hide = Not CBool(-chkVisible.Value)
    P.NameColor = csbNameColor.Color
    If Val(txtSize.Text) > 2 And Val(txtSize.Text) < 30 Then P.PhysicalWidth = Val(txtSize.Text)
    If cmbShape.ListIndex >= 0 Then P.Shape = cmbShape.ListIndex * 2 + 1
    If chkShowCoordinates.Value < 2 Then P.ShowCoordinates = -chkShowCoordinates.Value
    If chkShowName.Value < 2 Then P.ShowName = -chkShowName.Value
End With
End Sub

Private Sub SetAllPropsFromPoint(P As BasePointType, Optional ByVal bFromCode As Boolean = True)
With P
    FromCode = bFromCode
    
    csbForeColor.Color = P.ForeColor
    csbFillColor.Color = P.FillColor
    csbNameColor.Color = P.NameColor
    chkEnabled.Value = -P.Enabled
    chkVisible.Value = -(Not P.Hide)
    chkFill.Value = 1 - P.FillStyle
    txtSize.Text = Format(P.PhysicalWidth)
    vsbSize.Value = MinPointSize + MaxPointSize - P.PhysicalWidth
    cmbShape.ListIndex = Sgn(P.Shape - 1)
    chkShowCoordinates.Value = -P.ShowCoordinates
    chkShowName.Value = -P.ShowName
    FromCode = False
End With
End Sub

Private Sub Apply()
If Not cmdApply.Enabled Then Exit Sub
cmdApply.Enabled = False
RecordGenericAction ResUndoPointPropertiesChange

BasePoint = ClonedPoints
DrawSelectedPoints
End Sub

Private Sub DrawSelectedPoints()
Dim Z As Long

PaperCls
ShowAll , , False

For Z = 1 To PointCount
    If lstPoints.Selected(Z - 1) Then
        If BasePoint(Z).Visible And BasePoint(Z).Hide Then
            ShowSelectedPoint Paper.hDC, Z, , , True, True
        Else
            ShowSelectedPoint Paper.hDC, Z
        End If
    End If
Next

Paper.Refresh
End Sub

Private Sub EnablePropControls(Optional ByVal Enable As Boolean = True)
FromCode = Not Enable
'chkEnabled.Enabled = Enable
'chkVisible.Enabled = Enable
'chkShowName.Enabled = Enable
'chkShowCoordinates.Enabled = Enable
'chkFill.Enabled = Enable
'cmbShape.Enabled = Enable
'vsbSize.Enabled = Enable
'txtSize.Enabled = Enable
End Sub
