VERSION 5.00
Begin VB.Form frmAnLine 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Line"
   ClientHeight    =   6156
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   5064
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.4
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   43
   Icon            =   "AnLine.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6156
   ScaleWidth      =   5064
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtK 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   1440
      TabIndex        =   38
      Text            =   "0"
      Top             =   600
      Width           =   735
   End
   Begin VB.TextBox txtM 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   2760
      TabIndex        =   37
      Text            =   "0"
      Top             =   600
      Width           =   735
   End
   Begin VB.OptionButton optLineType 
      Caption         =   "y = kx + b"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   36
      Top             =   240
      Value           =   -1  'True
      Width           =   4575
   End
   Begin VB.OptionButton optLineType 
      Caption         =   "ax + by + c = 0"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   0
      Top             =   1320
      Width           =   4575
   End
   Begin VB.OptionButton optLineType 
      Caption         =   "Canonic"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   1
      Top             =   2400
      Width           =   4575
   End
   Begin VB.OptionButton optLineType 
      Caption         =   "Normal   (x * cos(a) + y * sin(a) + d = 0)"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   2
      Top             =   3480
      Width           =   4575
   End
   Begin VB.TextBox txtB 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   2760
      TabIndex        =   7
      Text            =   "0"
      Top             =   1680
      Width           =   735
   End
   Begin VB.TextBox txtC 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   4080
      TabIndex        =   8
      Text            =   "0"
      Top             =   1680
      Width           =   735
   End
   Begin VB.TextBox txtA 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   1440
      TabIndex        =   6
      Text            =   "0"
      Top             =   1680
      Width           =   735
   End
   Begin VB.TextBox txtX0 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   2760
      TabIndex        =   9
      Text            =   "0"
      Top             =   2760
      Width           =   735
   End
   Begin VB.TextBox txtY0 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   4080
      TabIndex        =   10
      Text            =   "0"
      Top             =   2760
      Width           =   735
   End
   Begin VB.TextBox txtA1 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   2760
      TabIndex        =   11
      Text            =   "0"
      Top             =   3000
      Width           =   735
   End
   Begin VB.TextBox txtA2 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   4080
      TabIndex        =   12
      Text            =   "0"
      Top             =   3000
      Width           =   735
   End
   Begin VB.TextBox txtAng 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   2760
      TabIndex        =   13
      Text            =   "0"
      Top             =   3840
      Width           =   735
   End
   Begin VB.TextBox txtD 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   4080
      TabIndex        =   14
      Text            =   "0"
      Top             =   3840
      Width           =   735
   End
   Begin VB.TextBox txtA2NV 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   4080
      TabIndex        =   18
      Text            =   "0"
      Top             =   5160
      Width           =   735
   End
   Begin VB.TextBox txtA1NV 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   2760
      TabIndex        =   17
      Text            =   "0"
      Top             =   5160
      Width           =   735
   End
   Begin VB.TextBox txtY0NV 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   4080
      TabIndex        =   16
      Text            =   "0"
      Top             =   4920
      Width           =   735
   End
   Begin VB.TextBox txtX0NV 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   2760
      TabIndex        =   15
      Text            =   "0"
      Top             =   4920
      Width           =   735
   End
   Begin VB.OptionButton optLineType 
      Caption         =   "Normal (normal vector)"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   3
      Top             =   4560
      Width           =   4575
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   5640
      Width           =   735
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3960
      TabIndex        =   5
      Top             =   5640
      Width           =   975
   End
   Begin VB.Label lblK 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "k = "
      Height          =   255
      Left            =   1080
      TabIndex        =   40
      Top             =   600
      Width           =   375
   End
   Begin VB.Label lblM 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "b = "
      Height          =   255
      Left            =   2400
      TabIndex        =   39
      Top             =   600
      Width           =   375
   End
   Begin VB.Label lblB 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "b = "
      Height          =   255
      Left            =   2400
      TabIndex        =   35
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label lblC 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "c = "
      Height          =   255
      Left            =   3720
      TabIndex        =   34
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label lblA 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "a = "
      Height          =   255
      Left            =   1080
      TabIndex        =   33
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label lblX0 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "x0 = "
      Height          =   255
      Left            =   2400
      TabIndex        =   32
      Top             =   2760
      Width           =   375
   End
   Begin VB.Label lblY0 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "y0 = "
      Height          =   255
      Left            =   3720
      TabIndex        =   31
      Top             =   2760
      Width           =   375
   End
   Begin VB.Label lblA1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "a1 = "
      Height          =   255
      Left            =   2400
      TabIndex        =   30
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label lblA2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "a2 = "
      Height          =   255
      Left            =   3720
      TabIndex        =   29
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label lblLinePoint 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Line point:"
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Label lblGuideVector 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Guide vector:"
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   3000
      Width           =   2175
   End
   Begin VB.Label lblAng 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "a = "
      Height          =   255
      Left            =   2400
      TabIndex        =   26
      Top             =   3840
      Width           =   375
   End
   Begin VB.Label lblD 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "d = "
      Height          =   255
      Left            =   3720
      TabIndex        =   25
      Top             =   3840
      Width           =   375
   End
   Begin VB.Label lblNormalVector 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Normal vector:"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   5160
      Width           =   1695
   End
   Begin VB.Label lblLinePointNV 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Line point:"
      Height          =   255
      Left            =   360
      TabIndex        =   23
      Top             =   4920
      Width           =   1455
   End
   Begin VB.Label lblA2NV 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "a2 = "
      Height          =   255
      Left            =   3720
      TabIndex        =   22
      Top             =   5160
      Width           =   375
   End
   Begin VB.Label lblA1NV 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "a1 = "
      Height          =   255
      Left            =   2400
      TabIndex        =   21
      Top             =   5160
      Width           =   375
   End
   Begin VB.Label lblY0NV 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "y0 = "
      Height          =   255
      Left            =   3720
      TabIndex        =   20
      Top             =   4920
      Width           =   375
   End
   Begin VB.Label lblX0NV 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "x0 = "
      Height          =   255
      Left            =   2400
      TabIndex        =   19
      Top             =   4920
      Width           =   375
   End
End
Attribute VB_Name = "frmAnLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim unlCancel As Boolean
Dim iLineType As Integer
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
If KeyCode = vbKeyReturn Then unlCancel = False: Unload Me
End Sub

Private Sub Form_Load()
unlCancel = False

Caption = GetString(ResMnuAnLine)
cmdCancel.Caption = GetString(ResCancel)
optLineType(2).Caption = GetString(ResCanonic)
optLineType(3).Caption = GetString(ResNormal) & "   x * cos(a) + y * sin(a) + d = 0"
optLineType(4).Caption = GetString(ResNormal) & " (" & GetString(ResNormalVector) & ")"
lblLinePoint.Caption = GetString(ResLinePoint) & ":"
lblLinePointNV.Caption = GetString(ResLinePoint) & ":"
lblGuideVector.Caption = GetString(ResGuideVector) & ":"
lblNormalVector.Caption = GetString(ResNormalVector) & ":"

If Not EditingExisting Then
    SetLineType 0
    Exit Sub
End If

Dim lineType As Integer
lineType = Figures(ActiveFigure).AuxInfo(6)
SetLineType lineType

If lineType = 1 Then
    Dim a As Long, b As Long, c As Long
    a = Figures(ActiveFigure).AuxInfo(1)
    b = Figures(ActiveFigure).AuxInfo(2)
    c = Figures(ActiveFigure).AuxInfo(3)
    
    If b = -1 Then
        optLineType(0).Value = True
        txtK.Text = a
        txtM.Text = c
        Exit Sub
    End If
    
    txtA.Text = a
    txtB.Text = b
    txtC.Text = c
ElseIf lineType = 2 Then
    txtX0.Text = Figures(ActiveFigure).AuxInfo(1)
    txtY0.Text = Figures(ActiveFigure).AuxInfo(2)
    txtA1.Text = Figures(ActiveFigure).AuxInfo(3)
    txtA2.Text = Figures(ActiveFigure).AuxInfo(4)
ElseIf lineType = 3 Then
    txtAng.Text = Figures(ActiveFigure).AuxInfo(1)
    txtD.Text = Figures(ActiveFigure).AuxInfo(2)
ElseIf lineType = 4 Then
    txtX0NV.Text = Figures(ActiveFigure).AuxInfo(1)
    txtY0NV.Text = Figures(ActiveFigure).AuxInfo(2)
    txtA1NV.Text = Figures(ActiveFigure).AuxInfo(3)
    txtA2NV.Text = Figures(ActiveFigure).AuxInfo(4)
End If

End Sub

Private Sub Form_Paint()
Dim ptRect As RECT
Const VSP = 4
Dim Top(0 To 4) As Long
Dim Z As Long

For Z = 0 To 4
    Top(Z) = ToScale(optLineType(Z).Top)
Next

ptRect.Left = ToScale(ScaleWidth - cmdCancel.Left - cmdCancel.Width)
ptRect.Top = Top(0) - VSP
ptRect.Right = ToScale(ScaleWidth) - ptRect.Left
ptRect.Bottom = Top(1) - 2 * VSP
DrawEdge hDC, ptRect, EDGE_ETCHED, BF_RECT

ptRect.Top = Top(1) - VSP
ptRect.Bottom = Top(2) - 2 * VSP
DrawEdge hDC, ptRect, EDGE_ETCHED, BF_RECT

ptRect.Top = Top(2) - VSP
ptRect.Bottom = Top(3) - 2 * VSP
DrawEdge hDC, ptRect, EDGE_ETCHED, BF_RECT

ptRect.Top = Top(3) - VSP
ptRect.Bottom = Top(4) - 2 * VSP
DrawEdge hDC, ptRect, EDGE_ETCHED, BF_RECT

ptRect.Top = Top(4) - VSP
ptRect.Bottom = ptRect.Bottom + Top(2) - Top(1)
DrawEdge hDC, ptRect, EDGE_ETCHED, BF_RECT

'ShadowControl cmdOK
'ShadowControl cmdCancel
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode <> vbFormCode Then unlCancel = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim kA As Double, kB As Double, kC As Double, kA1 As Double, kA2 As Double, kD As Double
FormMain.Enabled = True

If unlCancel Then
    EditingExisting = False
    Exit Sub
End If

If EditingExisting Then
    EditingExisting = False
    
    Select Case iLineType
        Case 0
            kA = Evaluate(txtK.Text)
            kB = -1
            kC = Evaluate(txtM.Text)
            If kA = EmptyVar Or kC = EmptyVar Then Exit Sub
            EditAnLineGeneral kA, kB, kC
        Case 1
            kA = Evaluate(txtA.Text)
            kB = Evaluate(txtB.Text)
            kC = Evaluate(txtC.Text)
            If kA = EmptyVar Or kB = EmptyVar Or kC = EmptyVar Then Exit Sub
            If kA = 0 And kB = 0 Then MsgBox "a = 0; b = 0: " & GetString(ResMsgABUnequalTo0), vbInformation: Cancel = 1: Exit Sub
            EditAnLineGeneral kA, kB, kC
        Case 2
            kA1 = Evaluate(txtA1.Text)
            kA2 = Evaluate(txtA2.Text)
            If kA1 = 0 And kA2 = 0 Then MsgBox "a1 = 0; a2 = 0: " & GetString(ResMsgABUnequalTo0), vbInformation: Cancel = 1: Exit Sub
            EditAnLineCanonic Evaluate(txtX0.Text), Evaluate(txtY0.Text), kA1, kA2
        Case 3
            EditAnLineNormal Evaluate(txtAng.Text), Evaluate(txtD.Text)
        Case 4
            kA1 = Evaluate(txtA1NV.Text)
            kA2 = Evaluate(txtA2NV.Text)
            If kA1 = 0 And kA2 = 0 Then MsgBox "a1 = 0; a2 = 0: " & GetString(ResMsgABUnequalTo0), vbInformation: Cancel = 1: Exit Sub
            EditAnLineNormalPoint Evaluate(txtX0NV.Text), Evaluate(txtY0NV.Text), kA1, kA2
    End Select
    
Else
    Select Case iLineType
        Case 0
            kA = Evaluate(txtK.Text)
            kB = -1
            kC = Evaluate(txtM.Text)
            If kA = EmptyVar Or kC = EmptyVar Then Exit Sub
            AddAnLineGeneral kA, kB, kC
        Case 1
            kA = Evaluate(txtA.Text)
            kB = Evaluate(txtB.Text)
            kC = Evaluate(txtC.Text)
            If kA = EmptyVar Or kB = EmptyVar Or kC = EmptyVar Then Exit Sub
            If kA = 0 And kB = 0 Then MsgBox "a = 0; b = 0: " & GetString(ResMsgABUnequalTo0), vbInformation: Cancel = 1: Exit Sub
            AddAnLineGeneral kA, kB, kC
        Case 2
            kA1 = Evaluate(txtA1.Text)
            kA2 = Evaluate(txtA2.Text)
            If kA1 = 0 And kA2 = 0 Then MsgBox "a1 = 0; a2 = 0: " & GetString(ResMsgABUnequalTo0), vbInformation: Cancel = 1: Exit Sub
            AddAnLineCanonic Evaluate(txtX0.Text), Evaluate(txtY0.Text), kA1, kA2
        Case 3
            AddAnLineNormal Evaluate(txtAng.Text), Evaluate(txtD.Text)
        Case 4
            kA1 = Evaluate(txtA1NV.Text)
            kA2 = Evaluate(txtA2NV.Text)
            If kA1 = 0 And kA2 = 0 Then MsgBox "a1 = 0; a2 = 0: " & GetString(ResMsgABUnequalTo0), vbInformation: Cancel = 1: Exit Sub
            AddAnLineNormalPoint Evaluate(txtX0NV.Text), Evaluate(txtY0NV.Text), kA1, kA2
    End Select
End If

PaperCls
ShowAll
End Sub

Sub SetLineType(ByVal index As Integer)
optLineType(index).Value = True
optLineType_Click index
End Sub

Private Sub optLineType_Click(index As Integer)
iLineType = index
Select Case index
    Case 0
        EnableText txtK, True
        EnableText txtM, True
        EnableText txtA, False
        EnableText txtB, False
        EnableText txtC, False
        EnableText txtX0, False
        EnableText txtY0, False
        EnableText txtA1, False
        EnableText txtA2, False
        EnableText txtAng, False
        EnableText txtD, False
        EnableText txtX0NV, False
        EnableText txtY0NV, False
        EnableText txtA1NV, False
        EnableText txtA2NV, False
    Case 1
        EnableText txtA, True
        EnableText txtB, True
        EnableText txtC, True
        EnableText txtK, False
        EnableText txtM, False
        EnableText txtX0, False
        EnableText txtY0, False
        EnableText txtA1, False
        EnableText txtA2, False
        EnableText txtAng, False
        EnableText txtD, False
        EnableText txtX0NV, False
        EnableText txtY0NV, False
        EnableText txtA1NV, False
        EnableText txtA2NV, False
    Case 2
        EnableText txtK, False
        EnableText txtM, False
        EnableText txtA, False
        EnableText txtB, False
        EnableText txtC, False
        EnableText txtX0, True
        EnableText txtY0, True
        EnableText txtA1, True
        EnableText txtA2, True
        EnableText txtAng, False
        EnableText txtD, False
        EnableText txtX0NV, False
        EnableText txtY0NV, False
        EnableText txtA1NV, False
        EnableText txtA2NV, False
    Case 3
        EnableText txtK, False
        EnableText txtM, False
        EnableText txtA, False
        EnableText txtB, False
        EnableText txtC, False
        EnableText txtX0, False
        EnableText txtY0, False
        EnableText txtA1, False
        EnableText txtA2, False
        EnableText txtAng, True
        EnableText txtD, True
        EnableText txtX0NV, False
        EnableText txtY0NV, False
        EnableText txtA1NV, False
        EnableText txtA2NV, False
    Case 4
        EnableText txtK, False
        EnableText txtM, False
        EnableText txtX0NV, True
        EnableText txtY0NV, True
        EnableText txtA1NV, True
        EnableText txtA2NV, True
        EnableText txtA, False
        EnableText txtB, False
        EnableText txtC, False
        EnableText txtX0, False
        EnableText txtY0, False
        EnableText txtA1, False
        EnableText txtA2, False
        EnableText txtAng, False
        EnableText txtD, False
End Select
End Sub

Private Sub txtA_GotFocus()
txtA.SelStart = 0
txtA.SelLength = Len(txtA.Text)
End Sub

Private Sub txtA_Validate(Cancel As Boolean)
Dim X As Double
'If Not IsNumeric(txtA.Text) Then Cancel = True: MsgBox GetString(ResMsgInputInteger)
On Error GoTo EH
If txtA.Text = "" Then GoTo EH
X = Evaluate(txtA.Text)
If X < -2 ^ 20 Or X > 2 ^ 20 Then GoTo EH
Exit Sub

EH:
MsgBox GetString(ResMsgCannotEvaluate), vbInformation
Cancel = True
End Sub

Private Sub txtA1_GotFocus()
txtA1.SelStart = 0
txtA1.SelLength = Len(txtA1.Text)
End Sub

Private Sub txtA1_Validate(Cancel As Boolean)
Dim X As Double
'If Not IsNumeric(txtA.Text) Then Cancel = True: MsgBox GetString(ResMsgInputInteger)
On Error GoTo EH
If txtA1.Text = "" Then GoTo EH
X = Evaluate(txtA1.Text)
If X < -2 ^ 20 Or X > 2 ^ 20 Then GoTo EH
Exit Sub

EH:
MsgBox GetString(ResMsgCannotEvaluate), vbInformation
Cancel = True
End Sub

Private Sub txtA2_GotFocus()
txtA2.SelStart = 0
txtA2.SelLength = Len(txtA2.Text)
End Sub

Private Sub txtA2_Validate(Cancel As Boolean)
Dim X As Double
On Error GoTo EH
If txtA2.Text = "" Then GoTo EH
X = Evaluate(txtA2.Text)
If X < -2 ^ 20 Or X > 2 ^ 20 Then GoTo EH
Exit Sub

EH:
MsgBox GetString(ResMsgCannotEvaluate), vbInformation
Cancel = True
End Sub

Private Sub txtA1NV_GotFocus()
txtA1NV.SelStart = 0
txtA1NV.SelLength = Len(txtA1NV.Text)
End Sub

Private Sub txtA1NV_Validate(Cancel As Boolean)
Dim X As Double
On Error GoTo EH
If txtA1NV.Text = "" Then GoTo EH
X = Evaluate(txtA1NV.Text)
If X < -2 ^ 20 Or X > 2 ^ 20 Then GoTo EH
Exit Sub

EH:
MsgBox GetString(ResMsgCannotEvaluate), vbInformation
Cancel = True
End Sub

Private Sub txtA2NV_GotFocus()
txtA2NV.SelStart = 0
txtA2NV.SelLength = Len(txtA2NV.Text)
End Sub

Private Sub txtA2NV_Validate(Cancel As Boolean)
Dim X As Double
On Error GoTo EH
If txtA2NV.Text = "" Then GoTo EH
X = Evaluate(txtA2NV.Text)
If X < -2 ^ 20 Or X > 2 ^ 20 Then GoTo EH
Exit Sub

EH:
MsgBox GetString(ResMsgCannotEvaluate), vbInformation
Cancel = True
End Sub

Private Sub txtAng_GotFocus()
txtAng.SelStart = 0
txtAng.SelLength = Len(txtAng.Text)
End Sub

Private Sub txtAng_Validate(Cancel As Boolean)
Dim X As Double
On Error GoTo EH
If txtAng.Text = "" Then GoTo EH
X = Evaluate(txtAng.Text)
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
'If Not IsNumeric(txtB.Text) Then Cancel = True: MsgBox GetString(ResMsgInputInteger)
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
'If Not IsNumeric(txtC.Text) Then Cancel = True: MsgBox GetString(ResMsgInputInteger)
On Error GoTo EH
If txtC.Text = "" Then GoTo EH
X = Evaluate(txtC.Text)
If X < -2 ^ 20 Or X > 2 ^ 20 Then GoTo EH
Exit Sub

EH:
MsgBox GetString(ResMsgCannotEvaluate), vbInformation
Cancel = True
End Sub

Private Sub txtD_GotFocus()
txtD.SelStart = 0
txtD.SelLength = Len(txtD.Text)
End Sub

Private Sub txtD_Validate(Cancel As Boolean)
Dim X As Double
On Error GoTo EH
If txtD.Text = "" Then GoTo EH
X = Evaluate(txtD.Text)
If X < -2 ^ 20 Or X > 2 ^ 20 Then GoTo EH
Exit Sub

EH:
MsgBox GetString(ResMsgCannotEvaluate), vbInformation
Cancel = True
End Sub

Private Sub txtX0_GotFocus()
txtX0.SelStart = 0
txtX0.SelLength = Len(txtX0.Text)
End Sub

Private Sub txtX0_Validate(Cancel As Boolean)
Dim X As Double
On Error GoTo EH
If txtX0.Text = "" Then GoTo EH
X = Evaluate(txtX0.Text)
If X < -2 ^ 20 Or X > 2 ^ 20 Then GoTo EH
Exit Sub

EH:
MsgBox GetString(ResMsgCannotEvaluate), vbInformation
Cancel = True
End Sub

Private Sub txtY0_GotFocus()
txtY0.SelStart = 0
txtY0.SelLength = Len(txtY0.Text)
End Sub

Private Sub txtY0_Validate(Cancel As Boolean)
Dim X As Double
On Error GoTo EH
If txtY0.Text = "" Then GoTo EH
X = Evaluate(txtY0.Text)
If X < -2 ^ 20 Or X > 2 ^ 20 Then GoTo EH
Exit Sub

EH:
MsgBox GetString(ResMsgCannotEvaluate), vbInformation
Cancel = True
End Sub

Private Sub txtX0NV_GotFocus()
txtX0NV.SelStart = 0
txtX0NV.SelLength = Len(txtX0NV.Text)
End Sub

Private Sub txtX0NV_Validate(Cancel As Boolean)
Dim X As Double
On Error GoTo EH
If txtX0NV.Text = "" Then GoTo EH
X = Evaluate(txtX0NV.Text)
If X < -2 ^ 20 Or X > 2 ^ 20 Then GoTo EH
Exit Sub

EH:
MsgBox GetString(ResMsgCannotEvaluate), vbInformation
Cancel = True
End Sub

Private Sub txtY0NV_GotFocus()
txtY0NV.SelStart = 0
txtY0NV.SelLength = Len(txtY0NV.Text)
End Sub

Private Sub txtY0NV_Validate(Cancel As Boolean)
Dim X As Double
On Error GoTo EH
If txtY0NV.Text = "" Then GoTo EH
X = Evaluate(txtY0NV.Text)
If X < -2 ^ 20 Or X > 2 ^ 20 Then GoTo EH
Exit Sub

EH:
MsgBox GetString(ResMsgCannotEvaluate), vbInformation
Cancel = True
End Sub

Private Sub txtK_GotFocus()
txtK.SelStart = 0
txtK.SelLength = Len(txtK.Text)
End Sub

Private Sub txtK_Validate(Cancel As Boolean)
Dim X As Double
On Error GoTo EH
If txtK.Text = "" Then GoTo EH
X = Evaluate(txtK.Text)
If X < -2 ^ 20 Or X > 2 ^ 20 Then GoTo EH
Exit Sub

EH:
MsgBox GetString(ResMsgCannotEvaluate), vbInformation
Cancel = True
End Sub

Private Sub txtM_GotFocus()
txtM.SelStart = 0
txtM.SelLength = Len(txtM.Text)
End Sub

Private Sub txtM_Validate(Cancel As Boolean)
Dim X As Double
On Error GoTo EH
If txtM.Text = "" Then GoTo EH
X = Evaluate(txtM.Text)
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

