VERSION 5.00
Begin VB.Form frmMacroSave 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Завершение создания макроса"
   ClientHeight    =   5400
   ClientLeft      =   36
   ClientTop       =   336
   ClientWidth     =   4824
   HelpContextID   =   210
   Icon            =   "frmMacroSave.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   4824
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtMacroDescription 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   1680
      TabIndex        =   10
      Top             =   4200
      Width           =   3012
   End
   Begin VB.TextBox txtMacroName 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   1680
      TabIndex        =   0
      Top             =   3840
      Width           =   3012
   End
   Begin VB.CommandButton cmdReselect 
      Caption         =   "Вернуться и выбрать заново"
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
      TabIndex        =   6
      Top             =   2640
      Width           =   4575
   End
   Begin VB.ListBox lstResults 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1992
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   5
      Top             =   480
      Width           =   4575
   End
   Begin VB.PictureBox picContainer 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   600
      Left            =   0
      ScaleHeight     =   600
      ScaleWidth      =   4824
      TabIndex        =   3
      Top             =   4800
      Width           =   4824
      Begin VB.CommandButton cmdHelp 
         Caption         =   "Помощь"
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
         TabIndex        =   12
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK - сохранить макрос"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   2280
         TabIndex        =   11
         Top             =   120
         Width           =   2415
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "Отмена"
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
         Left            =   1200
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      Index           =   0
      X1              =   120
      X2              =   4680
      Y1              =   4695
      Y2              =   4695
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000015&
      Index           =   0
      X1              =   120
      X2              =   4680
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      Index           =   1
      X1              =   120
      X2              =   4680
      Y1              =   3255
      Y2              =   3255
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000015&
      Index           =   1
      X1              =   120
      X2              =   4680
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Label lblMacroDescription 
      Caption         =   "Описание"
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
      Left            =   120
      TabIndex        =   9
      Top             =   4200
      Width           =   1455
   End
   Begin VB.Label lblMacroName 
      Caption         =   "Имя"
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
      Left            =   120
      TabIndex        =   8
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Label lblEnterName 
      Caption         =   "Теперь введите имя и описание макроса:"
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
      TabIndex        =   7
      Top             =   3480
      Width           =   4095
   End
   Begin VB.Label lbl3 
      AutoSize        =   -1  'True
      Caption         =   "3."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.4
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   480
      Left            =   120
      TabIndex        =   2
      Top             =   3360
      Width           =   345
   End
   Begin VB.Label lblResultsSelected 
      Caption         =   "Вы выбрали такие построения-результаты:"
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
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "frmMacroSave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim unlCancel As Boolean

Private Sub cmdCancel_Click()
CancelDialog
End Sub

Private Sub cmdHelp_Click()
DisplayHelpTopic Me.HelpContextID
End Sub

Private Sub cmdReselect_Click()
tempMacro.Name = txtMacroName.Text
tempMacro.Description = txtMacroDescription.Text
unlCancel = True
Unload Me

MacroResetResults

PaperCls
ShowAllWithResults
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then CancelDialog
End Sub

Private Sub cmdOK_Click()
Unload Me
End Sub

Public Sub CancelDialog()
unlCancel = True
Unload Me
i_CancelMacro
End Sub

Private Sub Form_Load()
txtMacroName.Text = tempMacro.Name
txtMacroDescription.Text = tempMacro.Description
If txtMacroName.Text = "" Then cmdOK.Enabled = False
FillDialogStrings

MacroFillResults

FillResultsList
unlCancel = False

'Me.Visible = True
'If txtMacroName.Enabled And txtMacroName.Visible Then txtMacroName.SetFocus
End Sub

Public Sub FillResultsList()
Dim Z As Long
lstResults.Clear

For Z = 1 To tempMacro.ResultCount
    lstResults.AddItem Z & ". " & tempMacro.Results(Z).Description
    lstResults.Selected(Z - 1) = Not tempMacro.Results(Z).Hide
Next

For Z = 1 To tempMacro.SGCount
    lstResults.AddItem Z + tempMacro.ResultCount & ". " & tempMacro.SG(Z).Description
    lstResults.Selected(Z + tempMacro.ResultCount - 1) = True
Next

AddListboxScrollbar lstResults

If tempMacro.ResultCount + tempMacro.SGCount > 0 Then
    lstResults.ListIndex = 0
Else
    lstResults.Enabled = False
    lstResults.BackColor = vbButtonFace
    cmdOK.Enabled = False
    txtMacroName.Enabled = False
    txtMacroDescription.Enabled = False
    txtMacroName.BackColor = vbButtonFace
    txtMacroDescription.BackColor = vbButtonFace
    lbl3.Enabled = False
    lblMacroName.Enabled = False
    lblMacroDescription.Enabled = False
    lblEnterName.Enabled = False
    lblResultsSelected.Caption = GetString(ResMacroErrBase + 2 * meResultsNotSelected)
End If
End Sub

Public Sub FillDialogStrings()
cmdCancel.Caption = GetString(ResCancel)
cmdHelp.Caption = GetString(ResHelp)
cmdReselect.Caption = GetString(ResReturnAndReselect)
Caption = GetString(ResMacroCompletingTask)
lblResultsSelected.Caption = GetString(ResMacroChosenSuchResults)
lblMacroName.Caption = GetString(ResName)
lblMacroDescription.Caption = GetString(ResDescription)
lblEnterName.Caption = GetString(ResMacroEnterNameAndDescription)
cmdOK.Caption = GetString(ResMacroOKSaveAs)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode <> vbFormCode Then unlCancel = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim S As String, Z As Long

InitZeroTags

If unlCancel = True Then Exit Sub

For Z = 1 To MacroCount
    If Macros(Z).Name = tempMacro.Name Then
        S = tempMacro.Name & GetString(ResMsgObjectAlreadyExists)
        tempMacro.Name = ""
        MsgBox S, vbInformation
        Cancel = 1
        Exit Sub
    End If
Next

If tempMacro.Name = "" Then
    MsgBox GetString(ResEnterMacroName), vbExclamation
    Cancel = 1
    Exit Sub
End If

Visible = False

MacroCreateSave
End Sub

Private Sub lstResults_ItemCheck(Item As Integer)
Static AlreadyBusy As Boolean
Dim Z As Long

If Not Visible Or AlreadyBusy Then Exit Sub
If Item < 0 Then Exit Sub
AlreadyBusy = True

If Item + 1 > tempMacro.ResultCount Then
    lstResults.Selected(Item) = True
Else
    tempMacro.Results(Item + 1).Hide = Not lstResults.Selected(Item)
    
    For Z = 0 To tempMacro.Results(Item + 1).NumberOfPoints - 1
        If IsChildPointPos(tempMacro.Results(Item + 1), Z) Then
            With tempMacro.Results(Item + 1)
                If .FigureType = dsIntersect Then
                    'If Not tempMacro.FigurePoints(tempMacro.Results(Item + 1).Points(Z)).Tag Then tempMacro.FigurePoints(tempMacro.Results(Item + 1).Points(Z)).Hide = tempMacro.Results(Item + 1).Hide
                    'If Not tempMacro.FigurePoints(tempMacro.Results(Item + 1).Points(Z)).Tag Then tempMacro.FigurePoints(tempMacro.Results(Item + 1).Points(Z)).Enabled = Not tempMacro.Results(Item + 1).Hide
                    tempMacro.FigurePoints(.Points(Z)).Hide = tempMacro.FigurePoints(.Points(Z)).Tag Or .Hide
                    tempMacro.FigurePoints(.Points(Z)).Enabled = Not tempMacro.FigurePoints(.Points(Z)).Hide
                Else
                    tempMacro.FigurePoints(.Points(Z)).Hide = tempMacro.Results(Item + 1).Hide
                    tempMacro.FigurePoints(.Points(Z)).Enabled = Not tempMacro.Results(Item + 1).Hide
                End If
            End With
            
        End If
    Next Z
End If

AlreadyBusy = False
End Sub

Private Sub txtMacroDescription_Change()
tempMacro.Description = txtMacroDescription.Text
End Sub

Private Sub txtMacroName_Change()
tempMacro.Name = txtMacroName.Text
cmdOK.Enabled = txtMacroName.Text <> ""
End Sub
