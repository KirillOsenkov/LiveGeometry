VERSION 5.00
Begin VB.Form frmCalculator 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calculator"
   ClientHeight    =   3672
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   7884
   HelpContextID   =   267
   Icon            =   "frmCalculator.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   306
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   657
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCreateLabel 
      Caption         =   "Create a label to measure this expression"
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
      Left            =   120
      TabIndex        =   4
      Top             =   3240
      Width           =   4575
   End
   Begin VB.TextBox txtResult 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   375
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   120
      Width           =   2415
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
      Left            =   7080
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3240
      Width           =   735
   End
   Begin DG.ctlCalculator ctlCalculator1 
      Height          =   2415
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   7695
      _ExtentX        =   13568
      _ExtentY        =   4255
   End
   Begin VB.TextBox txtExpression 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4935
   End
   Begin VB.Label lblEqualSign 
      Caption         =   "=  "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "frmCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim unlCancel As Boolean

Private Sub cmdCreateLabel_Click()
On Local Error Resume Next
Dim pAction As Action
Dim T As Tree, V As Single

If txtExpression.Text = "" Then
    txtResult.Text = ""
    lblEqualSign.Visible = False
    Exit Sub
End If

If IsNumeric(txtExpression.Text) Then
    txtResult.Text = txtExpression.Text 'Format(CDbl(txtExpression.Text), setFormatNumber): Exit Sub
End If

T = BuildTree(txtExpression.Text)
If T.Erroneous Or WasThereAnErrorEvaluatingLastExpression <> eetEverythingOK Then txtResult.Text = "": Exit Sub

AddTextLabel "[" & txtExpression.Text & "]"
pAction.Type = actAddLabel
pAction.pLabel = LabelCount
RecordAction pAction

PaperCls
ShowAll

cmdCreateLabel.Enabled = False
End Sub

Private Sub cmdOK_Click()
Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
Set ctlCalculator1.ParentTextbox = txtExpression
'ShadowControl cmdOK

Caption = GetString(ResCalculator)
cmdCreateLabel.Caption = GetString(ResCreateLabelWithMeasurement)

FormMain.Enabled = False
Set Me.Font = txtResult.Font
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode <> vbFormCode Then unlCancel = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
FormMain.Enabled = True
End Sub

Private Sub txtExpression_Change()
On Local Error Resume Next
Dim T As Tree, V, D As Double

cmdCreateLabel.Enabled = False

If txtExpression.Text = "" Then
    txtResult.Text = ""
    lblEqualSign.Visible = False
    Exit Sub
End If

If IsNumeric(txtExpression.Text) Then txtResult.Text = txtExpression.Text 'Format(CDbl(txtExpression.Text), setFormatNumber): Exit Sub

T = BuildTree(txtExpression.Text)
If T.Erroneous Or WasThereAnErrorEvaluatingLastExpression <> eetEverythingOK Then
    txtResult.Text = ""
    lblEqualSign.Visible = False
    Exit Sub
End If

D = RecalculateTree(T, 1)
If Abs(D) > 7.92281625142643E+28 Then
    txtResult.Text = ""
    lblEqualSign.Visible = False
    Exit Sub
End If
V = CDec(D)
If ERR <> 0 Then
    txtResult.Text = ""
    lblEqualSign.Visible = False
    Exit Sub
End If
If T.Erroneous Or WasThereAnErrorEvaluatingLastExpression <> eetEverythingOK Then
    txtResult.Text = ""
    lblEqualSign.Visible = False
    Exit Sub
End If

txtResult.Text = V 'Format(V, "0.####")
lblEqualSign.Visible = True
If IsTreeDynamic(T) Then cmdCreateLabel.Enabled = True

Do While TextWidth(txtResult.Text) > txtResult.Width And Len(txtResult.Text) > 3
    If Right(txtResult.Text, 3) = "..." Then txtResult.Text = Left(txtResult.Text, Len(txtResult.Text) - 3)
    txtResult.Text = Left(txtResult.Text, Len(txtResult.Text) - 1) & "..."
Loop
End Sub

Private Sub txtExpression_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtExpression_Change
End Sub
