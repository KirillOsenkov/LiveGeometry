VERSION 5.00
Begin VB.Form frmPointRename 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Point rename"
   ClientHeight    =   2820
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   4488
   Icon            =   "frmPointRename.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   4488
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   2520
      TabIndex        =   6
      Top             =   2280
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
      Left            =   3360
      TabIndex        =   7
      Top             =   2280
      Width           =   975
   End
   Begin VB.TextBox txtScrollName 
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
      Left            =   3720
      TabIndex        =   0
      Text            =   "D"
      Top             =   1125
      Width           =   615
   End
   Begin VB.TextBox txtRename 
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
      Left            =   3720
      TabIndex        =   1
      Text            =   "C"
      Top             =   645
      Width           =   615
   End
   Begin VB.OptionButton optRename 
      Caption         =   "Выбрать другое имя для точки %1:"
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
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   3255
   End
   Begin VB.OptionButton optRename 
      Caption         =   "Присвоить точке %1 новое имя:"
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
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Value           =   -1  'True
      Width           =   3255
   End
   Begin VB.OptionButton optRename 
      Caption         =   "Поменять имена точек %1 и %2 местами"
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
      Index           =   2
      Left            =   240
      TabIndex        =   4
      Top             =   1560
      Width           =   3975
   End
   Begin VB.Label lblExists 
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
      Left            =   120
      TabIndex        =   8
      Top             =   2280
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      Index           =   1
      X1              =   120
      X2              =   4320
      Y1              =   2055
      Y2              =   2055
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   1
      X1              =   120
      X2              =   4320
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Label lblPrompt 
      Caption         =   "Точка с именем %1 уже существует. Что делать?"
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
      TabIndex        =   5
      Top             =   240
      Width           =   4215
   End
End
Attribute VB_Name = "frmPointRename"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum RenameErrorMessageType
    remNameOK
    remNameAlreadyExists
    remNameNotSpecified
End Enum

Dim unlCancel As Boolean

Public OldName As String
Public NewName As String
Public Index As Long
Public NewIndex As Long

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
If KeyCode = vbKeyReturn Then cmdOK_Click
End Sub

Private Sub Form_Load()
unlCancel = False
NewIndex = GetPointByName(NewName)
FillDialogStrings
txtScrollName.SelStart = 0
txtScrollName.SelLength = Len(txtScrollName)
ValidateDialog
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode <> vbFormCode Then unlCancel = True
End Sub

Public Sub FillDialogStrings()
cmdCancel.Caption = GetString(ResCancel)
Caption = Replace(Replace(GetString(ResRenamePoint), "%1", OldName), "%2", NewName)
lblPrompt.Caption = Replace(GetString(ResRenameWhatToDo), "%1", NewName)
lblExists.Caption = GetString(ResName) & GetString(ResMsgObjectAlreadyExists)
optRename(0).Caption = Replace(GetString(ResRenameChooseAnotherName), "%1", OldName)
optRename(1).Caption = Replace(GetString(ResRenameAssignAnother), "%1", NewName)
optRename(2).Caption = Replace(Replace(GetString(ResRenameSwapNames), "%1", OldName), "%2", NewName)
txtRename = NewName
txtScrollName = GenerateNewPointName(ShouldAllocate:=False)
End Sub

Private Sub Form_Unload(Cancel As Integer)
If unlCancel Then Exit Sub

RecordGenericAction ResUndoRenamePoint

If optRename(0).Value Then
    RenamePointActions Index, txtRename
ElseIf optRename(1).Value Then
    RenamePointActions NewIndex, txtScrollName
    RenamePointActions Index, NewName
Else
    BasePoint(NewIndex).Name = BasePoint(Index).Name
    BasePoint(Index).Name = NewName
    ReRestoreExpressionsFromTrees
End If

PaperCls
ShowAll
End Sub

Private Sub optRename_Click(Index As Integer)
ValidateDialog
If Index = 0 Then
    txtRename.SelStart = 0
    txtRename.SelLength = Len(txtRename)
    txtRename.SetFocus
End If
If Index = 1 Then
    txtScrollName.SelStart = 0
    txtScrollName.SelLength = Len(txtScrollName)
    txtScrollName.SetFocus
End If
End Sub

Private Sub txtRename_Change()
ValidateDialog
End Sub

Private Sub txtScrollName_Change()
ValidateDialog
End Sub

Public Sub EnableOK(ByVal En As Boolean)
cmdOK.Enabled = En
End Sub

Public Sub ValidateDialog()
Dim B As Boolean, Index As Long
For Index = 0 To 2
    If optRename(Index).Value = True Then Exit For
Next

If Index = 0 Then
    B = Not GetPointByName(txtRename) > 0 And txtRename <> ""
    EnableOK B
    If B Then
        ShowMessage remNameOK
    Else
        If txtRename <> "" Then ShowMessage remNameAlreadyExists Else ShowMessage remNameNotSpecified
    End If
    EnableTextBox txtRename, True
    EnableTextBox txtScrollName, False
End If
If Index = 1 Then
    B = Not GetPointByName(txtScrollName) > 0 And txtScrollName <> ""
    EnableOK B
    If B Then
        ShowMessage remNameOK
    Else
        If txtScrollName <> "" Then ShowMessage remNameAlreadyExists Else ShowMessage remNameNotSpecified
    End If
    EnableTextBox txtRename, False
    EnableTextBox txtScrollName, True
End If
If Index = 2 Then
    EnableOK True
    ShowMessage remNameOK
    EnableTextBox txtRename, False
    EnableTextBox txtScrollName, False
End If
End Sub

Private Sub ShowMessage(R As RenameErrorMessageType)
Select Case R
Case remNameAlreadyExists
    lblExists = GetString(ResName) & GetString(ResMsgObjectAlreadyExists)
    lblExists.Visible = True
Case remNameNotSpecified
    lblExists = GetString(ResMsgNoName)
    lblExists.Visible = True
Case remNameOK
    lblExists.Visible = False
End Select
End Sub
