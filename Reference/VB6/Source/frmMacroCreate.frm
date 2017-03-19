VERSION 5.00
Begin VB.Form frmMacroCreate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Начало создания макроса"
   ClientHeight    =   2796
   ClientLeft      =   36
   ClientTop       =   336
   ClientWidth     =   4824
   HelpContextID   =   208
   Icon            =   "frmMacroCreate.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2796
   ScaleWidth      =   4824
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picCancelContainer 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   840
      Left            =   0
      ScaleHeight     =   840
      ScaleWidth      =   4824
      TabIndex        =   3
      Top             =   1956
      Width           =   4824
      Begin VB.CommandButton cmdHelp 
         Caption         =   "Справка"
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
         HelpContextID   =   208
         Left            =   120
         TabIndex        =   6
         Top             =   360
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
         Height          =   255
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
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdSelectGivens 
      Caption         =   "Указать исходные данные"
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
      Left            =   600
      TabIndex        =   1
      Top             =   1200
      Width           =   4092
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      Index           =   1
      X1              =   120
      X2              =   4680
      Y1              =   1815
      Y2              =   1815
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000015&
      Index           =   1
      X1              =   120
      X2              =   4680
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000015&
      Index           =   0
      X1              =   120
      X2              =   4680
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      Index           =   0
      X1              =   120
      X2              =   4680
      Y1              =   975
      Y2              =   975
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      Caption         =   "1."
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
      Top             =   1200
      Width           =   345
   End
   Begin VB.Label lblInstruction 
      BackStyle       =   0  'Transparent
      Caption         =   "Сейчас Вы сможете создать макрос. Для дополнительной информации см. Справка."
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "frmMacroCreate"
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

Private Sub cmdHelp_Click()
DisplayHelpTopic Me.HelpContextID
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    unlCancel = True
    Unload Me
End If
End Sub

'==================================================
Private Sub cmdSelectGivens_Click()
Unload Me
MacroCreateBegin
End Sub

Private Sub Form_Load()
FillDialogStrings
unlCancel = False
End Sub

Public Sub FillDialogStrings()
cmdCancel.Caption = GetString(ResCancel)
cmdHelp.Caption = GetString(ResHelp)
cmdSelectGivens.Caption = GetString(ResMnuMacroSelectGivens)
lblInstruction.Caption = GetString(ResMacroGivenPrompt)
chkDoNotShow.Caption = GetString(ResDoNotShowThisDialog)
Caption = GetString(ResMacroCreation)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode <> vbFormCode Then unlCancel = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
If chkDoNotShow.Value = 1 Then
    setShowMacroCreateDialog = False
    SaveSetting AppName, "General", "ShowMacroCreateDialog", "0"
End If
End Sub

