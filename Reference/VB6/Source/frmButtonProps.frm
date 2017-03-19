VERSION 5.00
Begin VB.Form frmButtonProps 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Button properties"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4785
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   204
   Icon            =   "frmButtonProps.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   4785
   StartUpPosition =   2  'CenterScreen
   Begin DG.ctlColorBox csbForeColor 
      Height          =   255
      Left            =   1440
      TabIndex        =   23
      Top             =   3900
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
   End
   Begin VB.CheckBox chkFix 
      Caption         =   "Fixed"
      Height          =   255
      Left            =   2520
      TabIndex        =   12
      Top             =   3900
      Width           =   2175
   End
   Begin VB.CommandButton cmdFont 
      Caption         =   "Font"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2880
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   4440
      Width           =   735
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3720
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   4440
      Width           =   975
   End
   Begin VB.Frame fraTypeProps 
      Caption         =   "Show/hide button"
      Height          =   2535
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Visible         =   0   'False
      Width           =   4575
      Begin VB.CommandButton cmdSelectObjects 
         Caption         =   "Select objects"
         Height          =   615
         Left            =   2760
         TabIndex        =   10
         Top             =   480
         Width           =   1692
      End
      Begin VB.CheckBox chkInitialVisible 
         Caption         =   "Initially visible"
         Height          =   615
         Left            =   2760
         TabIndex        =   9
         Top             =   1800
         Width           =   1695
      End
      Begin VB.ListBox lstObjectList 
         Height          =   1950
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   2532
      End
      Begin VB.Label lblObjectList 
         Caption         =   "Object list"
         Height          =   252
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1932
      End
   End
   Begin VB.ComboBox cmbType 
      Height          =   288
      ItemData        =   "frmButtonProps.frx":030A
      Left            =   2040
      List            =   "frmButtonProps.frx":031A
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   480
      Width           =   2655
   End
   Begin VB.TextBox txtName 
      Height          =   315
      Left            =   2040
      TabIndex        =   1
      Top             =   120
      Width           =   2655
   End
   Begin VB.Frame fraTypeProps 
      Caption         =   "Launch file"
      Height          =   2535
      Index           =   3
      Left            =   120
      TabIndex        =   18
      Top             =   960
      Visible         =   0   'False
      Width           =   4575
      Begin VB.CheckBox chkRemind 
         Caption         =   "Remind to save file"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   840
         Value           =   1  'Checked
         Width           =   4335
      End
      Begin VB.CommandButton cmdBrowseForFile 
         Caption         =   "..."
         Height          =   375
         Left            =   4080
         TabIndex        =   20
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtFilePath 
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   3975
      End
   End
   Begin VB.Frame fraTypeProps 
      Caption         =   "Play sound"
      Height          =   2535
      Index           =   2
      Left            =   120
      TabIndex        =   15
      Top             =   960
      Visible         =   0   'False
      Width           =   4575
      Begin VB.CommandButton cmdBrowseForSound 
         Caption         =   "..."
         Height          =   375
         Left            =   4080
         TabIndex        =   17
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtSoundPath 
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   3975
      End
   End
   Begin VB.Frame fraTypeProps 
      Caption         =   "Message button"
      Height          =   2535
      Index           =   1
      Left            =   120
      TabIndex        =   13
      Top             =   960
      Visible         =   0   'False
      Width           =   4575
      Begin VB.TextBox txtMessage 
         Height          =   2172
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   14
         Top             =   240
         Width           =   4332
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   2
      X1              =   120
      X2              =   4680
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      Index           =   2
      X1              =   120
      X2              =   4695
      Y1              =   3735
      Y2              =   3735
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   1
      X1              =   120
      X2              =   4680
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      Index           =   1
      X1              =   120
      X2              =   4695
      Y1              =   4335
      Y2              =   4335
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      Index           =   0
      X1              =   150
      X2              =   5850
      Y1              =   5415
      Y2              =   5415
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000015&
      Index           =   0
      X1              =   150
      X2              =   5850
      Y1              =   5400
      Y2              =   5400
   End
   Begin VB.Label lblForeColor 
      BackStyle       =   0  'Transparent
      Caption         =   "Color"
      Height          =   255
      Left            =   1800
      TabIndex        =   22
      Top             =   3900
      Width           =   615
   End
   Begin VB.Label lblType 
      Caption         =   "Type"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   518
      Width           =   1935
   End
   Begin VB.Label lblName 
      Caption         =   "Name"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   1935
   End
End
Attribute VB_Name = "frmButtonProps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim SelObjectList As ObjectList

Dim unlCancel As Boolean
Dim CurrentType As Long

Dim TempFontName As String
Dim TempFontSize As Long
Dim TempFontBold As Boolean
Dim TempFontItalic As Boolean
Dim TempFontUnderline As Boolean
Dim TempFontCharset As Long

Private Sub cmbType_Change()
fraTypeProps(CurrentType).Visible = False
CurrentType = cmbType.ListIndex
fraTypeProps(CurrentType).Visible = True
End Sub

Private Sub cmbType_Click()
fraTypeProps(CurrentType).Visible = False
CurrentType = cmbType.ListIndex
fraTypeProps(CurrentType).Visible = True
End Sub

Private Sub cmdBrowseForFile_Click()
CD.Filter = "*.*|*.*"
CD.Flags = &H1000 + &H4
CD.InitDir = LastFileLinkPath
CD.DialogTitle = GetString(ResOpen)
CD.ShowOpen
If CD.Cancelled = True Or Dir(CD.FileName) = "" Then Exit Sub
txtFilePath.Text = CD.FileName
If IsValidPath(RetrieveDir(CD.FileName)) Then LastFileLinkPath = AddDirSep(RetrieveDir(CD.FileName))
End Sub

Private Sub cmdBrowseForSound_Click()
CD.Filter = "*." & extWAV & "|*." & extWAV
CD.Flags = &H1000 + &H4
CD.InitDir = LastSoundPath
CD.DialogTitle = GetString(ResOpen)
CD.ShowOpen
If CD.Cancelled = True Or Dir(CD.FileName) = "" Then Exit Sub
txtSoundPath.Text = CD.FileName
If IsValidPath(RetrieveDir(CD.FileName)) Then LastSoundPath = AddDirSep(RetrieveDir(CD.FileName))
End Sub

Private Sub cmdCancel_Click()
unlCancel = True
Unload Me
End Sub

Private Sub cmdFont_Click()
CD.FontName = TempFontName
CD.FontSize = TempFontSize
CD.FontBold = TempFontBold
CD.FontItalic = TempFontItalic
CD.FontUnderline = TempFontUnderline
CD.FontCharset = TempFontCharset

CD.ShowFont

TempFontName = CD.FontName
TempFontSize = CD.FontSize
TempFontBold = CD.FontBold
TempFontItalic = CD.FontItalic
TempFontUnderline = CD.FontUnderline
TempFontCharset = CD.FontCharset
End Sub

Private Sub cmdOK_Click()
Unload Me
End Sub

Private Sub cmdSelectObjects_Click()
If ActiveButton <> 0 Then
    If Buttons(ActiveButton).CurrentState = 0 Then ButtonPushed ActiveButton
End If
Me.Hide
TempObjectSelection = SelObjectList
ObjectSelectionBegin ostShowHideObjects, False, oscButton
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then unlCancel = True: Unload Me
End Sub

Private Sub Form_Load()
FillDialogStrings

If ActiveButton <> 0 Then
    txtName.Text = Buttons(ActiveButton).Caption
    CurrentType = Buttons(ActiveButton).Type
    cmbType.ListIndex = Buttons(ActiveButton).Type
    EnableTextBox cmbType, False
    chkFix.Value = -Buttons(ActiveButton).Fixed
    csbForeColor.Color = Buttons(ActiveButton).ForeColor
    
    
    Select Case Buttons(ActiveButton).Type
    Case butShowHide
        FillListBoxWithObjectList lstObjectList, Buttons(ActiveButton).ObjectListAux, gotGeneric
        SelObjectList = Buttons(ActiveButton).ObjectListAux
        chkInitialVisible.Value = -Buttons(ActiveButton).InitiallyVisible
        fraTypeProps(0).Visible = True
        'cmdSelectObjects.Enabled = False
    Case butMsgBox
        fraTypeProps(1).Visible = True
        txtMessage.Text = Buttons(ActiveButton).Message
    Case butPlaySound
        fraTypeProps(2).Visible = True
        txtSoundPath.Text = Buttons(ActiveButton).Path
    Case butLaunchFile
        fraTypeProps(3).Visible = True
        txtFilePath.Text = Buttons(ActiveButton).Path
        chkRemind.Value = -Buttons(ActiveButton).RemindToSaveFile
    End Select
    TempFontBold = Buttons(ActiveButton).FontBold
    TempFontCharset = Buttons(ActiveButton).Charset
    TempFontItalic = Buttons(ActiveButton).FontItalic
    TempFontName = Buttons(ActiveButton).FontName
    TempFontSize = Buttons(ActiveButton).FontSize
    TempFontUnderline = Buttons(ActiveButton).FontUnderline
Else
    txtName.Text = GetString(ResButton) & (ButtonCount + 1)
    cmbType.ListIndex = 0
    csbForeColor.Color = vbButtonText
    CurrentType = 0
    fraTypeProps(0).Visible = True
    TempFontBold = setdefLabelBold
    ObjectListClear SelObjectList
    TempFontCharset = setdefLabelCharset
    TempFontItalic = setdefLabelItalic
    TempFontName = setdefLabelFont
    TempFontSize = setdefLabelFontSize
    TempFontUnderline = setdefLabelUnderline
End If

txtName.SelStart = Len(txtName.Text)

'ShadowControl cmdOK
'ShadowControl cmdCancel

FormMain.Enabled = False
unlCancel = False
Show
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode <> vbFormCode Then unlCancel = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim pAction As Action

If unlCancel Then
    ActiveButton = 0
    FormMain.Enabled = True
    FormMain.SetFocus
    Exit Sub
End If

If SelObjectList.TotalCount = 0 And CurrentType = 0 Then
    MsgBox GetString(ResObjectsNotSelected), vbExclamation, GetString(ResError)
    Cancel = 1
    Exit Sub
End If

If txtName.Text = "" Then
    MsgBox GetString(ResMsgNoName), vbExclamation
    Cancel = 1
    Exit Sub
End If

If CurrentType = 1 And txtMessage.Text = "" Then
    MsgBox GetString(ResError), vbExclamation
    Cancel = 1
    Exit Sub
End If

If CurrentType = 2 And txtSoundPath = "" Then
    MsgBox GetString(ResError), vbExclamation
    Cancel = 1
    Exit Sub
End If

If CurrentType = 2 And Dir(txtSoundPath) = "" Then
    MsgBox GetString(ResError), vbExclamation
    Cancel = 1
    Exit Sub
End If

If CurrentType = 3 And txtFilePath = "" Then
    MsgBox GetString(ResError), vbExclamation
    Cancel = 1
    Exit Sub
End If

If ActiveButton = 0 Then
    If Not AddButton(cmbType.ListIndex, txtName.Text, 0, 0) Then
        ActiveButton = 0
        FormMain.Enabled = True
        FormMain.SetFocus
        Exit Sub
    End If
    ActiveButton = ButtonCount
Else
    pAction.Type = actChangeAttrButton
    ReDim pAction.sButton(1 To 1)
    pAction.sButton(1) = Buttons(ActiveButton)
    pAction.pButton = ActiveButton
    RecordAction pAction
End If

With Buttons(ActiveButton)
    .Caption = txtName
    .Type = cmbType.ListIndex
    .Fixed = -chkFix.Value
    .FontBold = TempFontBold
    .ForeColor = csbForeColor.Color
    .Charset = TempFontCharset
    .FontItalic = TempFontItalic
    .FontName = TempFontName
    .FontSize = TempFontSize
    .FontUnderline = TempFontUnderline
    Select Case .Type
    Case butShowHide
        .CurrentState = -chkInitialVisible.Value
        .InitiallyVisible = -chkInitialVisible.Value
        .ObjectListAux = SelObjectList
        ObjectListShowHideAll Buttons(ActiveButton).ObjectListAux, CBool(-chkInitialVisible.Value), False
    Case butMsgBox
        .Message = txtMessage.Text
    Case butPlaySound
        .Path = txtSoundPath.Text
    Case butLaunchFile
        .Path = txtFilePath.Text
        .RemindToSaveFile = -chkRemind.Value
    End Select
End With

GetButtonSize ActiveButton

FormMain.Enabled = True
FormMain.SetFocus

ActiveButton = 0

PaperCls
ShowAll
End Sub

Public Sub ObjectSelectionComplete_ShowHide()
FormMain.Enabled = False
FillListBoxWithObjectList lstObjectList, TempObjectSelection, gotGeneric
SelObjectList = TempObjectSelection
ObjectListClear TempObjectSelection
Me.Show
Me.SetFocus
End Sub

Public Sub ObjectSelectionCancel_ShowHide()
If ActiveButton = 0 Then
    ObjectListClear TempObjectSelection
    unlCancel = True
    Unload Me
Else
    FormMain.Enabled = False
    ObjectListClear TempObjectSelection
    Me.Show
    Me.SetFocus
End If
End Sub

Private Sub FillDialogStrings()
Dim z As Long
Caption = GetString(ResButtonProperties)
lblName.Caption = GetString(ResCaptionName)
cmdCancel.Caption = GetString(ResCancel)
lblType.Caption = GetString(ResType)
cmbType.List(0) = GetString(ResShowHideObjects)
cmbType.List(1) = GetString(ResMessageButton)
cmbType.List(2) = GetString(ResPlaySound)
cmbType.List(3) = GetString(ResLaunchFile)
cmdSelectObjects.Caption = GetString(ResAddRemoveObjects)
cmdFont.Caption = GetString(ResFont)
chkInitialVisible.Caption = GetString(ResInitiallyVisible)
chkFix.Caption = GetString(ResFix)
lblForeColor.Caption = GetString(ResForeColor)
chkRemind.Caption = GetString(ResRemindToSaveFile)
lblObjectList.Caption = GetString(ResObjectList)
For z = 0 To 3
    fraTypeProps(z).Caption = cmbType.List(z)
Next
End Sub

Private Sub lstObjectList_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Then
    DeleteObjectFromList
End If
End Sub

Private Sub DeleteObjectFromList()
Dim objIndex As Long, ObjType As GeometryObjectType
Dim z As Long

z = lstObjectList.ListIndex
If z < 0 Then Exit Sub

ObjectListGetItemFromGenericIndex SelObjectList, z + 1, objIndex, ObjType
ObjectListDelete SelObjectList, ObjType, objIndex
FillListBoxWithObjectList lstObjectList, SelObjectList
End Sub
