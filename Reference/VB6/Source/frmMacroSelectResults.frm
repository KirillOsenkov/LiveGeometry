VERSION 5.00
Begin VB.Form frmMacroSelectResults 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Выбор результатов"
   ClientHeight    =   5364
   ClientLeft      =   36
   ClientTop       =   336
   ClientWidth     =   4824
   HelpContextID   =   211
   Icon            =   "frmMacroSelectResults.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5364
   ScaleWidth      =   4824
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtGivenHint 
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
      Left            =   120
      TabIndex        =   10
      Top             =   2760
      Width           =   3015
   End
   Begin VB.CommandButton cmdGivenHintDefault 
      Caption         =   "По умолчанию"
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
      Left            =   3120
      TabIndex        =   9
      Top             =   2760
      Width           =   1575
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
      Height          =   735
      Left            =   2880
      TabIndex        =   7
      Top             =   480
      Width           =   1812
   End
   Begin VB.ListBox lstGivens 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1968
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   2652
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   960
      Left            =   0
      ScaleHeight     =   960
      ScaleWidth      =   4824
      TabIndex        =   3
      Top             =   4404
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
         HelpContextID   =   211
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Width           =   975
      End
      Begin VB.CheckBox chkDoNotShow 
         Caption         =   "Пропускать это окно в дальнейшем."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   120
         TabIndex        =   5
         Top             =   0
         Width           =   3372
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
         Left            =   3720
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdSelectResults 
      Caption         =   "Выбрать построения-результаты"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   120
      TabIndex        =   1
      Top             =   3840
      Width           =   4572
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      Index           =   0
      X1              =   120
      X2              =   4680
      Y1              =   4335
      Y2              =   4335
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000015&
      Index           =   0
      X1              =   120
      X2              =   4680
      Y1              =   4320
      Y2              =   4320
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
   Begin VB.Label lblGivenHint 
      Caption         =   "Подсказка при выборе этого элемента исходных данных"
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
      TabIndex        =   11
      Top             =   2520
      Width           =   4575
   End
   Begin VB.Label lblSelectResults 
      Caption         =   "Теперь выберите результаты:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   8
      Top             =   3360
      Width           =   4095
   End
   Begin VB.Label lbl2 
      AutoSize        =   -1  'True
      Caption         =   "2."
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
      Top             =   3240
      Width           =   345
   End
   Begin VB.Label lblGivens 
      Caption         =   "Вы выбрали такие исходные данные:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4692
   End
End
Attribute VB_Name = "frmMacroSelectResults"
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
'MacroResetGivens
PaperCls
ShowAllWithGivens
unlCancel = True
Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then CancelDialog
End Sub

Public Sub CancelDialog()
unlCancel = True
i_CancelMacro
Unload Me
End Sub

Private Sub cmdGivenHintDefault_Click()
If lstGivens.ListCount < 1 Or lstGivens.ListIndex < 0 Or Not txtGivenHint.Enabled Then Exit Sub
txtGivenHint.Text = GetString(ResFigureBase + 2 * tempMacro.Givens(lstGivens.ListIndex + 1).Type)
End Sub

Private Sub cmdSelectResults_Click()
Unload Me
End Sub

Private Sub Form_Load()
FillDialogStrings
FillGivensList

unlCancel = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode <> vbFormCode Then unlCancel = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
If unlCancel Then Exit Sub

If chkDoNotShow.Value = 1 Then
    setShowMacroResultsDialog = False
    SaveSetting AppName, "General", "ShowMacroResultsDialog", "0"
End If

Visible = False
MacroCreateResults
End Sub

Public Sub FillGivensList()
Dim Z As Long
Dim Q As Long
Dim S As String

lstGivens.Clear
For Z = 1 To tempMacro.GivenCount
    lstGivens.AddItem Z & ". " & GetString(ResFigureBase + tempMacro.Givens(Z).Type * 2)
Next
AddListboxScrollbar lstGivens

If tempMacro.GivenCount > 0 Then
    lstGivens.ListIndex = 0
Else
    lstGivens.Enabled = False
    txtGivenHint.Visible = False
    lstGivens.BackColor = vbButtonFace
    cmdGivenHintDefault.Visible = False
    lblGivenHint.Visible = False
    lblGivens.Caption = GetString(ResMacroChosenNoGivens)
    If IndependentFigureCount > 0 Then
        lblSelectResults.Caption = GetString(ResMacroNeverthelessSelectResults)
    Else
        lblSelectResults.Caption = GetString(ResMacroNoResultsWithoutGivens)
        cmdSelectResults.Enabled = False
        lbl2.Enabled = False
    End If
    
End If
End Sub

Public Sub FillDialogStrings()
cmdCancel.Caption = GetString(ResCancel)
cmdHelp.Caption = GetString(ResHelp)
chkDoNotShow.Caption = GetString(ResDoNotShowThisDialog)
Caption = GetString(ResMacroResultSelection)
lblGivenHint.Caption = GetString(ResMacroThisGivenPrompt)
lblGivens.Caption = GetString(ResMacroChosenSuchGivens)
lblSelectResults.Caption = GetString(ResNowSelectResults)
cmdGivenHintDefault.Caption = GetString(ResDefault)
cmdReselect.Caption = GetString(ResReturnAndReselect)
cmdSelectResults.Caption = GetString(ResMacroSelectResults)
End Sub

Private Sub lstGivens_Click()
txtGivenHint.Enabled = False
txtGivenHint.Text = tempMacro.Givens(lstGivens.ListIndex + 1).Description
txtGivenHint.Enabled = True
End Sub

Private Sub txtGivenHint_Change()
If lstGivens.ListCount < 1 Or lstGivens.ListIndex < 0 Or Not txtGivenHint.Enabled Then Exit Sub
tempMacro.Givens(lstGivens.ListIndex + 1).Description = txtGivenHint.Text
'lstGivens.List(lstGivens.ListIndex) = (lstGivens.ListIndex + 1) & ". " & txtGivenHint.Text
End Sub
