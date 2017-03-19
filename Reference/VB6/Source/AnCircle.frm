VERSION 5.00
Begin VB.Form frmAnCircle 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Circle equation"
   ClientHeight    =   2424
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   4464
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.4
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   44
   Icon            =   "AnCircle.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2424
   ScaleWidth      =   4464
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtX 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   720
      TabIndex        =   10
      Text            =   "0"
      Top             =   600
      Width           =   855
   End
   Begin VB.TextBox txtR 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   3360
      TabIndex        =   12
      Text            =   "0"
      Top             =   600
      Width           =   855
   End
   Begin VB.TextBox txtY 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   2040
      TabIndex        =   11
      Text            =   "0"
      Top             =   600
      Width           =   855
   End
   Begin VB.OptionButton optEqType 
      Caption         =   "(x - xc)² + (y - yc)² = R²"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   240
      Value           =   -1  'True
      Width           =   3975
   End
   Begin VB.OptionButton optEqType 
      Caption         =   "x² + y² + ax + by + c = 0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1080
      Width           =   3975
   End
   Begin VB.TextBox txtB 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   255
      Left            =   2040
      TabIndex        =   1
      Text            =   "0"
      Top             =   1440
      Width           =   855
   End
   Begin VB.TextBox txtC 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   255
      Left            =   3360
      TabIndex        =   2
      Text            =   "0"
      Top             =   1440
      Width           =   855
   End
   Begin VB.TextBox txtA 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Text            =   "0"
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label lblX 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "xc = "
      Height          =   255
      Left            =   360
      TabIndex        =   15
      Top             =   600
      Width           =   375
   End
   Begin VB.Label lblR 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "R = "
      Height          =   255
      Left            =   3000
      TabIndex        =   14
      Top             =   600
      Width           =   375
   End
   Begin VB.Label lblY 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "yc = "
      Height          =   255
      Left            =   1680
      TabIndex        =   13
      Top             =   600
      Width           =   375
   End
   Begin VB.Label lblB 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "b = "
      Height          =   255
      Left            =   1680
      TabIndex        =   8
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label lblC 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "c = "
      Height          =   255
      Left            =   3000
      TabIndex        =   7
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label lblA 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "a = "
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   1440
      Width           =   375
   End
End
Attribute VB_Name = "frmAnCircle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim unlCancel As Boolean
Public EditingExisting As Boolean

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
Caption = GetString(ResMnuAnCircle)
cmdCancel.Caption = GetString(ResCancel)

unlCancel = False

If EditingExisting Then
    txtX.Text = GetCircleCenter(ActiveFigure).X
    txtY.Text = GetCircleCenter(ActiveFigure).Y
    txtR.Text = GetCircleRadius(ActiveFigure)
    txtA.Text = Figures(ActiveFigure).AuxInfo(1)
    txtB.Text = Figures(ActiveFigure).AuxInfo(2)
    txtC.Text = Figures(ActiveFigure).AuxInfo(3)
End If

End Sub

Private Sub Form_Paint()
Dim ptRect As RECT
Const VSP = 4
Dim top0 As Long
Dim top1 As Long
top0 = ToScale(optEqType(0).Top)
top1 = ToScale(optEqType(1).Top)

#Const DrawShadows = 1

#If DrawShadows = 0 Then
    Gradient hDC, BackColor, colStatusGradient, 0, 0, ToScale(ScaleWidth), ToScale(ScaleHeight), False
#End If

'ptRect.Left = ScaleWidth - cmdCancel.Left - cmdCancel.Width
'ptRect.Top = ptRect.Left
'ptRect.Right = ScaleWidth - ptRect.Left
'ptRect.Bottom = cmdCancel.Top - ptRect.Left
'DrawEdge hDC, ptRect, EDGE_ETCHED, BF_RECT

ptRect.Left = ToScale(ScaleWidth - cmdCancel.Left - cmdCancel.Width)
ptRect.Top = top0 - VSP
ptRect.Right = ToScale(ScaleWidth) - ptRect.Left
ptRect.Bottom = top1 - 2 * VSP
DrawEdge hDC, ptRect, EDGE_ETCHED, BF_RECT

ptRect.Top = top1 - VSP
ptRect.Bottom = ptRect.Bottom + top1 - top0
DrawEdge hDC, ptRect, EDGE_ETCHED, BF_RECT

#If DrawShadows = 1 Then
    'ShadowControl cmdOK
    'ShadowControl cmdCancel
    'ShadowControl txtA
    'ShadowControl txtB
    'ShadowControl txtC
#End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode <> vbFormCode Then unlCancel = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
If unlCancel Then
    FormMain.Enabled = True
    EditingExisting = False
    Exit Sub
End If

Dim A As Double, B As Double, C As Double
Dim X As Double, Y As Double, R As Double
Dim Z As Long

If optEqType(1).Value Then
    A = Evaluate(txtA.Text)
    B = Evaluate(txtB.Text)
    C = Evaluate(txtC.Text)
    If C = 0 And B = 0 And A = 0 Or A = EmptyVar Or B = EmptyVar Or C = EmptyVar Then MsgBox "a = 0; b = 0; c = 0: " & GetString(ResMsgABUnequalTo0), vbOKOnly + vbExclamation: Cancel = 1: Exit Sub
    If A * A + B * B - 4 * C <= 0 Then MsgBox GetString(ResImaginaryCircle), vbOKOnly + vbExclamation: Cancel = 1: Exit Sub
Else
    X = Evaluate(txtX.Text)
    Y = Evaluate(txtY.Text)
    R = Evaluate(txtR.Text)
    If X = EmptyVar Or Y = EmptyVar Or Z = EmptyVar Then MsgBox GetString(ResError) & ": R <= 0", vbExclamation: Cancel = 1: Exit Sub
    If R <= 0 Then MsgBox GetString(ResImaginaryCircle), vbExclamation: Cancel = 1: Exit Sub
End If

If optEqType(0).Value Then
    A = -2 * X
    B = -2 * Y
    C = X * X + Y * Y - R * R
End If

If Not EditingExisting Then
    AddAnCircle A, B, C
Else
    EditAnCircle A, B, C
    EditingExisting = False
End If

PaperCls
ShowAll
End Sub

Sub SetCircleType(ByVal Index As Integer)
optEqType(Index).Value = True
optEqType_Click Index
End Sub

Private Sub optEqType_Click(Index As Integer)
Select Case Index
    Case 0
        EnableText txtA, False
        EnableText txtB, False
        EnableText txtC, False
        EnableText txtX, True
        EnableText txtY, True
        EnableText txtR, True
    Case 1
        EnableText txtA, True
        EnableText txtB, True
        EnableText txtC, True
        EnableText txtX, False
        EnableText txtY, False
        EnableText txtR, False
End Select
End Sub

Private Sub txtA_GotFocus()
txtA.SelStart = 0
txtA.SelLength = Len(txtA.Text)
End Sub

Private Sub txtA_Validate(Cancel As Boolean)
Dim X As Double
On Error GoTo EH
If txtA.Text = "" Then GoTo EH
X = Evaluate(txtA.Text)
If X < -2 ^ 20 Or X > 2 ^ 20 Then GoTo EH
Exit Sub

EH:
MsgBox GetString(ResMsgCannotEvaluate), vbInformation
Cancel = True
End Sub

Private Sub txtB_GotFocus()
txtB.SelStart = 0
txtB.SelLength = Len(txtB.Text)
End Sub

Private Sub txtB_Validate(Cancel As Boolean)
Dim X As Double
On Error GoTo EH
If txtB.Text = "" Then GoTo EH
X = Evaluate(txtB.Text)
If X < -2 ^ 20 Or X > 2 ^ 20 Then GoTo EH
Exit Sub

EH:
MsgBox GetString(ResMsgCannotEvaluate), vbInformation
Cancel = True
End Sub

Private Sub txtC_GotFocus()
txtC.SelStart = 0
txtC.SelLength = Len(txtC.Text)
End Sub

Private Sub txtC_Validate(Cancel As Boolean)
Dim X As Double
On Error GoTo EH
If txtC.Text = "" Then GoTo EH
X = Evaluate(txtC.Text)
If X < -2 ^ 20 Or X > 2 ^ 20 Then GoTo EH
Exit Sub

EH:
MsgBox GetString(ResMsgCannotEvaluate), vbInformation
Cancel = True
End Sub

Private Sub txtX_GotFocus()
txtX.SelStart = 0
txtX.SelLength = Len(txtX.Text)
End Sub

Private Sub txtX_Validate(Cancel As Boolean)
Dim X As Double
On Error GoTo EH
If txtX.Text = "" Then GoTo EH
X = Evaluate(txtX.Text)
If X < -2 ^ 20 Or X > 2 ^ 20 Then GoTo EH
Exit Sub

EH:
MsgBox GetString(ResMsgCannotEvaluate), vbInformation
Cancel = True
End Sub

Private Sub txtY_GotFocus()
txtY.SelStart = 0
txtY.SelLength = Len(txtY.Text)
End Sub

Private Sub txtY_Validate(Cancel As Boolean)
Dim X As Double
On Error GoTo EH
If txtY.Text = "" Then GoTo EH
X = Evaluate(txtY.Text)
If X < -2 ^ 20 Or X > 2 ^ 20 Then GoTo EH
Exit Sub

EH:
MsgBox GetString(ResMsgCannotEvaluate), vbInformation
Cancel = True
End Sub

Private Sub txtR_GotFocus()
txtR.SelStart = 0
txtR.SelLength = Len(txtR.Text)
End Sub

Private Sub txtR_Validate(Cancel As Boolean)
Dim X As Double
On Error GoTo EH
If txtR.Text = "" Then GoTo EH
X = Evaluate(txtR.Text)
If X < -2 ^ 20 Or X > 2 ^ 20 Then GoTo EH
Exit Sub

EH:
MsgBox GetString(ResMsgCannotEvaluate), vbInformation
Cancel = True
End Sub

Private Sub EnableText(TextB As TextBox, ByVal Enable As Boolean)
TextB.Enabled = Enable
TextB.BackColor = IIf(Enable, vbWindowBackground, vbButtonFace)
End Sub

Public Function ToScale(ByVal V As Long, Optional ByVal NewScaleMode As Long = vbPixels) As Long
ToScale = ScaleX(V, ScaleMode, NewScaleMode)
End Function
