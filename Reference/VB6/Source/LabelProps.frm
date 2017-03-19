VERSION 5.00
Begin VB.Form frmLabelProps 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Label properties"
   ClientHeight    =   3264
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   7980
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.4
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   41
   Icon            =   "LabelProps.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3264
   ScaleWidth      =   7980
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   665
      TabIndex        =   5
      Top             =   2775
      Width           =   7980
      Begin VB.CommandButton cmdLabelFont 
         Caption         =   "Font"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   0
         Width           =   1215
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Height          =   375
         Left            =   6000
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   6840
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   0
         Width           =   975
      End
      Begin VB.CheckBox chkFix 
         Caption         =   "Fix in place"
         Height          =   375
         Left            =   3120
         TabIndex        =   7
         Top             =   0
         Width           =   2175
      End
      Begin DG.ctlColorBox csbForeColor 
         Height          =   255
         Left            =   1560
         TabIndex        =   6
         Top             =   60
         Width           =   255
         _ExtentX        =   445
         _ExtentY        =   445
      End
      Begin VB.Label lblForeColor 
         BackStyle       =   0  'Transparent
         Caption         =   "Color"
         Height          =   255
         Left            =   1920
         TabIndex        =   11
         Top             =   60
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdInsertExpression 
      Caption         =   "Insert expression in brackets"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2280
      Width           =   7695
   End
   Begin VB.PictureBox picLayout 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   7560
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1800
      Width           =   240
      Begin VB.Label lblLayout 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         Caption         =   "En"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   240
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   240
      End
   End
   Begin DG.ctlCalculator ctlCalculator1 
      Height          =   2415
      Left            =   120
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2280
      Visible         =   0   'False
      Width           =   7695
      _ExtentX        =   13568
      _ExtentY        =   4255
      BracketsVisible =   -1  'True
   End
   Begin VB.TextBox txtLabelCaption 
      Height          =   1935
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   7695
   End
End
Attribute VB_Name = "frmLabelProps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim unlCancel As Boolean
Dim TempFontName As String
Dim TempFontSize As Long
Dim TempFontBold As Boolean
Dim TempFontItalic As Boolean
Dim TempFontUnderline As Boolean
Dim TempFontCharset As Long
Dim TempCaptionOnLoad As String

Dim pAction As Action

Private Sub cmdCancel_Click()
unlCancel = True
Unload Me
End Sub

Private Sub cmdInsertExpression_Click()
If Not Me.Visible Then Exit Sub
LockWindowUpdate Me.hWnd

cmdInsertExpression.Visible = False
ctlCalculator1.Visible = True
Me.Height = Me.Height + ctlCalculator1.Height - cmdInsertExpression.Height
CenterForm Me

LockWindowUpdate 0
Me.Refresh

ctlCalculator1.InsertBrackets
End Sub

Private Sub cmdLabelFont_Click()
CD.FontName = TempFontName
CD.FontSize = TempFontSize
CD.FontBold = TempFontBold
CD.FontItalic = TempFontItalic
CD.FontUnderline = TempFontUnderline
CD.Color = csbForeColor.Color
CD.FontCharset = TempFontCharset

CD.ShowFont

TempFontName = CD.FontName
TempFontSize = CD.FontSize
TempFontBold = CD.FontBold
TempFontItalic = CD.FontItalic
TempFontUnderline = CD.FontUnderline
TempFontCharset = CD.FontCharset
txtLabelCaption.FontBold = TempFontBold
txtLabelCaption.FontItalic = TempFontItalic
txtLabelCaption.FontUnderline = TempFontUnderline
txtLabelCaption.FontSize = TempFontSize
txtLabelCaption.FontName = TempFontName
txtLabelCaption.Font.Charset = TempFontCharset
txtLabelCaption.Refresh
If CD.Color <> csbForeColor.Color And CD.Color <> 0 Then csbForeColor.Color = CD.Color: txtLabelCaption.ForeColor = CD.Color
End Sub

Private Sub cmdOK_Click()
Unload Me
End Sub

Private Sub csbForeColor_ColorChanged(ByVal NewColor As Long, ByVal OldColor As Long)
txtLabelCaption.ForeColor = csbForeColor.Color
End Sub

Private Sub Form_Activate()
RefreshKeyboardLayout
End Sub

Private Sub Form_GotFocus()
RefreshKeyboardLayout
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then unlCancel = True: Unload Me
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
RefreshKeyboardLayout
End Sub

Private Sub Form_Load()

Set ctlCalculator1.ParentTextbox = txtLabelCaption

If ActiveLabel > 0 Then
    ReDim pAction.sLabel(1 To 1)
    pAction.sLabel(1) = TextLabels(ActiveLabel)
    pAction.pLabel = ActiveLabel
    pAction.Type = actChangeAttrLabel
    
    TempFontName = TextLabels(ActiveLabel).FontName
    TempFontSize = TextLabels(ActiveLabel).FontSize
    TempFontBold = TextLabels(ActiveLabel).FontBold
    TempFontItalic = TextLabels(ActiveLabel).FontItalic
    TempFontUnderline = TextLabels(ActiveLabel).FontUnderline
    TempFontCharset = TextLabels(ActiveLabel).Charset

    txtLabelCaption.Text = TextLabels(ActiveLabel).Caption
    csbForeColor.Color = TextLabels(ActiveLabel).ForeColor
    chkFix.Value = -TextLabels(ActiveLabel).Fixed
Else
    TempFontBold = setdefLabelBold
    TempFontItalic = setdefLabelItalic
    TempFontUnderline = setdefLabelUnderline
    TempFontSize = setdefLabelFontSize
    TempFontName = setdefLabelFont
    TempFontCharset = DefaultFontCharset
    csbForeColor.Color = setdefcolLabelColor
End If

txtLabelCaption.FontBold = TempFontBold
txtLabelCaption.FontItalic = TempFontItalic
txtLabelCaption.FontUnderline = TempFontUnderline
txtLabelCaption.FontSize = TempFontSize
txtLabelCaption.FontName = TempFontName
txtLabelCaption.BackColor = nPaperColor1
txtLabelCaption.Font.Charset = TempFontCharset

Caption = GetString(ResLabelProperties)
cmdCancel.Caption = GetString(ResCancel)
lblForeColor.Caption = GetString(ResColor)
txtLabelCaption.ForeColor = csbForeColor.Color
cmdLabelFont.Caption = GetString(ResFont)
chkFix.Caption = GetString(ResFix)
cmdInsertExpression.Caption = GetString(ResCalcBase + 2 * ResCalcBrackets)

If InStr(txtLabelCaption.Text, "[") > 0 And InStr(txtLabelCaption, "]") > 0 Then
    cmdInsertExpression.Visible = False
    ctlCalculator1.Visible = True
    Me.Height = Me.Height + ctlCalculator1.Height - cmdInsertExpression.Height
    'CenterForm Me
    'ctlCalculator1.Visible = True
    'cmdInsertExpression_Click
    'ctlCalculator1.Visible = True
End If

Dim SBSize As Long
Dim Offset As Long
SBSize = ScaleX(GetSystemMetrics(SM_CYHSCROLL), vbPixels, ScaleMode)
Offset = ScaleX(2, vbPixels, ScaleMode)

picLayout.Move txtLabelCaption.Left + txtLabelCaption.Width - SBSize - Offset, txtLabelCaption.Top + txtLabelCaption.Height - SBSize - Offset, SBSize, SBSize
lblLayout.Move 0, 0, picLayout.ScaleWidth, picLayout.ScaleHeight

'ShadowControl cmdLabelFont
'ShadowControl csbForeColor
'ShadowControl cmdOK
'ShadowControl cmdCancel

FormMain.Enabled = False
unlCancel = False
Show
txtLabelCaption.SetFocus
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode <> vbFormCode Then unlCancel = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
If unlCancel Then FormMain.Enabled = True: FormMain.SetFocus: Exit Sub

If txtLabelCaption.Text = "" Then
    Cancel = 1
    MsgBox GetString(ResMsgNoName), vbExclamation
    Exit Sub
End If

FormMain.Enabled = True
FormMain.SetFocus

If ActiveLabel = 0 Then
    AddLabelJob RTrim(txtLabelCaption.Text), csbForeColor.Color, TempFontBold, TempFontItalic, _
        TempFontUnderline, TempFontSize, TempFontName, TempFontCharset, -chkFix.Value
    Exit Sub
'    AddTextLabel txtLabelCaption.Text
'    ActiveLabel = LabelCount
'    With TextLabels(ActiveLabel)
'        .Caption = txtLabelCaption.Text
'        .DisplayName = .Caption
'        .LenDisplayName = Len(.Caption)
'        .FontBold = TempFontBold
'        .FontItalic = TempFontItalic
'        .FontUnderline = TempFontUnderline
'        .ForeColor = csbForeColor.Color
'        .FontSize = TempFontSize
'        .FontName = TempFontName
'        .Charset = TempFontCharset
'        .Fixed = -chkFix.Value
'    End With
'    pAction.Type = actAddLabel
'    pAction.pLabel = ActiveLabel
'    RecordAction pAction
Else
    With TextLabels(ActiveLabel)
        .Caption = RTrim(txtLabelCaption.Text)
        .DisplayName = .Caption
        .LenDisplayName = Len(.Caption)
        .FontBold = TempFontBold
        .FontItalic = TempFontItalic
        .FontUnderline = TempFontUnderline
        .ForeColor = csbForeColor.Color
        .FontSize = TempFontSize
        .FontName = TempFontName
        .Charset = TempFontCharset
        .Fixed = -chkFix.Value
    End With
    RecordAction pAction
End If

ParseLabel ActiveLabel
GetLabelSize ActiveLabel
PaperCls
ShowAll
End Sub

Public Sub RefreshKeyboardLayout()
Dim A As String * 40, S As String, Z As Long

Z = GetKeyboardLayout(0) And 65535
GetLocaleInfo Z, LOCALE_SABBREVLANGNAME, A, 40
S = Left(StripNull(A), 2)
S = StrConv(S, vbProperCase)
lblLayout.Caption = S
End Sub

Public Sub SetNextKeyboardLayout()
ActivateKeyboardLayout HKL_NEXT, 0
RefreshKeyboardLayout
End Sub

Private Sub lblLayout_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
SetNextKeyboardLayout
End Sub

Private Sub txtLabelCaption_GotFocus()
RefreshKeyboardLayout
End Sub

Private Sub txtLabelCaption_KeyDown(KeyCode As Integer, Shift As Integer)
On Local Error GoTo EH:
Dim P As Integer, L As Integer, S As String, C As Integer

With txtLabelCaption
    If .Locked = True Then .Locked = False
    
    If KeyCode = vbKeyBack And Shift = 2 Then
        If .SelStart > 0 Then
            If .SelLength = 0 Then
                S = .Text
                C = .SelStart
                If C > 0 Then
                    Do While Mid(S, C, 1) = " " And C > 0
                        S = Left(S, C - 1) & Right(S, Len(S) - C)
                        C = C - 1
                        If S = "" Then Exit Do
                    Loop
                End If
                If S <> "" Then
                    P = InStrRev(S, " ", C)
                    L = C - P
                    .Text = S
                    .SelStart = P
                    .SelLength = L
                Else
                    .Text = ""
                End If
            End If
            .SelText = ""
        End If
        KeyCode = 0
        .Locked = True
        Exit Sub
    End If

End With

EH:
End Sub
