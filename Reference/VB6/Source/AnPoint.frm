VERSION 5.00
Begin VB.Form frmAnPoint 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Point"
   ClientHeight    =   3744
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   7908
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.4
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   42
   Icon            =   "AnPoint.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   312
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   659
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6840
      TabIndex        =   3
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   6000
      TabIndex        =   2
      Top             =   3240
      Width           =   735
   End
   Begin DG.ctlCalculator ctlCalculator1 
      Height          =   2415
      Left            =   120
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   720
      Width           =   7695
      _ExtentX        =   13568
      _ExtentY        =   4255
   End
   Begin VB.TextBox txtYCoord 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   4440
      TabIndex        =   1
      Top             =   240
      Width           =   3375
   End
   Begin VB.TextBox txtXCoord 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   3375
   End
   Begin VB.Label lblYCoord 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "Y = "
      Height          =   255
      Left            =   4080
      TabIndex        =   5
      Top             =   240
      Width           =   375
   End
   Begin VB.Label lblXCoord 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "X = "
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   375
   End
End
Attribute VB_Name = "frmAnPoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim unlCancel As Boolean

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
Dim ptRect As RECT
unlCancel = False
Caption = GetString(ResMnuAnPoint)
cmdCancel.Caption = GetString(ResCancel)
FormMain.Enabled = False
ctlCalculator1.BracketsVisible = False

ptRect.Left = ScaleWidth - cmdCancel.Left - cmdCancel.Width
ptRect.Top = ptRect.Left
ptRect.Right = ScaleWidth - ptRect.Left
ptRect.Bottom = cmdCancel.Top - ptRect.Left
'DrawEdge hDC, ptRect, EDGE_ETCHED, BF_RECT

'ShadowControl cmdOK
'ShadowControl cmdCancel

Show
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode <> vbFormCode Then unlCancel = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim X As Double, Y As Double
FormMain.Enabled = True
If unlCancel Then Exit Sub

If txtXCoord.Text = "" Then txtXCoord = "0"
If txtYCoord.Text = "" Then txtYCoord = "0"

X = Evaluate(txtXCoord.Text)
If (txtXCoord.Text = "") Or (X < -2 ^ 20 Or X > 2 ^ 20) Then
    MsgBox GetString(ResMsgCannotEvaluate) & " (X)" & vbCrLf & GetString(ResError) & ": " & GetString(ResEvalErrorBase + 2 * WasThereAnErrorEvaluatingLastExpression - 2), vbExclamation
    Cancel = True
    FormMain.Enabled = False
    Exit Sub
End If
Y = Evaluate(txtYCoord.Text)
If (txtYCoord.Text = "") Or (Y < -2 ^ 20 Or Y > 2 ^ 20) Then
    MsgBox GetString(ResMsgCannotEvaluate) & " (Y)" & vbCrLf & GetString(ResError) & ": " & GetString(ResEvalErrorBase + 2 * WasThereAnErrorEvaluatingLastExpression - 2), vbExclamation
    Cancel = True
    FormMain.Enabled = False
    Exit Sub
End If

AddAnPoint txtXCoord.Text, txtYCoord.Text
PaperCls
ShowAll
End Sub

Private Sub txtXCoord_GotFocus()
'txtXCoord.SelStart = 0
'txtXCoord.SelLength = Len(txtXCoord.Text)
Set ctlCalculator1.ParentTextbox = txtXCoord
End Sub

Private Sub txtYCoord_GotFocus()
'txtYCoord.SelStart = 0
'txtYCoord.SelLength = Len(txtYCoord.Text)
Set ctlCalculator1.ParentTextbox = txtYCoord
End Sub
